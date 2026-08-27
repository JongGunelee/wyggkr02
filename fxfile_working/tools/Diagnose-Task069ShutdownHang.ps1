[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$StageRoot,

    [string]$EvidenceRoot = (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) '__BUILD_TEMP_BACKUP__\task069_shutdown_diagnostic'),

    [int]$ReadyTimeoutSeconds = 180,

    [int]$ExitTimeoutSeconds = 15,

    [int]$HoldSeconds = 0
)

$ErrorActionPreference = 'Stop'

if (-not ('FxTask069.NativeMethods' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace FxTask069
{
    public static class NativeMethods
    {
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr GetProp(IntPtr hWnd, string name);

        public static IntPtr GetLayoutWindow(int processId, out int readyViewCount)
        {
            IntPtr layoutWindow = IntPtr.Zero;
            int maximum = 0;
            EnumWindows(delegate(IntPtr hWnd, IntPtr unused)
            {
                uint ownerProcessId;
                GetWindowThreadProcessId(hWnd, out ownerProcessId);
                if (ownerProcessId != (uint)processId)
                    return true;

                long value = GetProp(hWnd, "FxFile.StartupLayoutReadyViewCount").ToInt64();
                if (value > maximum)
                {
                    maximum = (int)value;
                    layoutWindow = hWnd;
                }
                return true;
            }, IntPtr.Zero);
            readyViewCount = maximum;
            return layoutWindow;
        }
    }
}
'@
}

$exe = Join-Path $StageRoot 'fxfile.exe'
if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) {
    throw "Staged fxfile.exe was not found: $exe"
}

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$runRoot = Join-Path $EvidenceRoot $stamp
New-Item -ItemType Directory -Path $runRoot -Force | Out-Null
$dumpPath = Join-Path $runRoot 'fxfile_shutdown_hang.dmp'
$stackPath = Join-Path $runRoot 'fxfile_shutdown_hang_stacks.txt'
$reportPath = Join-Path $runRoot 'shutdown_diagnostic_report.json'
$process = $null
$forced = $false
$readyCount = 0
$layoutWindow = [IntPtr]::Zero
$watch = [Diagnostics.Stopwatch]::StartNew()
$report = [ordered]@{
    StageRoot = $StageRoot
    ExecutableSha256 = (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash
    StartedAt = (Get-Date).ToString('o')
    ReadySeconds = $null
    ReadyViewCount = 0
    HoldSeconds = $HoldSeconds
    ResponsiveSamples = 0
    UnresponsiveSamples = 0
    WorkingSetStartBytes = $null
    WorkingSetPeakBytes = $null
    PrivateMemoryStartBytes = $null
    PrivateMemoryPeakBytes = $null
    HandleCountStart = $null
    HandleCountPeak = $null
    ThreadCountStart = $null
    ThreadCountPeak = $null
    ExitCommandPosted = $false
    ExitedNormally = $false
    ExitSeconds = $null
    ExitCode = $null
    DumpPath = $null
    StackPath = $null
    ForcedTermination = $false
}

try {
    $oldCompat = $env:__COMPAT_LAYER
    try {
        $env:__COMPAT_LAYER = 'RunAsInvoker'
        $process = Start-Process -FilePath $exe -WorkingDirectory $StageRoot -PassThru
    }
    finally {
        $env:__COMPAT_LAYER = $oldCompat
    }

    $readyDeadline = [DateTime]::UtcNow.AddSeconds($ReadyTimeoutSeconds)
    do {
        if ($process.HasExited) {
            throw "The diagnostic process exited before becoming ready. ExitCode=$($process.ExitCode)"
        }
        $process.Refresh()
        $layoutWindow = [FxTask069.NativeMethods]::GetLayoutWindow($process.Id, [ref]$readyCount)
        if ($layoutWindow -ne [IntPtr]::Zero -and $readyCount -ge 4 -and $process.Responding) {
            break
        }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $readyDeadline)

    if ($layoutWindow -eq [IntPtr]::Zero -or $readyCount -lt 4) {
        throw "The diagnostic process did not publish a ready 2x2 layout. ReadyViewCount=$readyCount"
    }

    $report.ReadySeconds = [math]::Round($watch.Elapsed.TotalSeconds, 3)
    $report.ReadyViewCount = $readyCount

    if ($HoldSeconds -gt 0) {
        $holdDeadline = [DateTime]::UtcNow.AddSeconds($HoldSeconds)
        do {
            if ($process.HasExited) {
                throw "The diagnostic process exited during the hold interval. ExitCode=$($process.ExitCode)"
            }
            $process.Refresh()
            if ($process.Responding) {
                ++$report.ResponsiveSamples
            }
            else {
                ++$report.UnresponsiveSamples
            }

            if ($null -eq $report.WorkingSetStartBytes) {
                $report.WorkingSetStartBytes = $process.WorkingSet64
                $report.PrivateMemoryStartBytes = $process.PrivateMemorySize64
                $report.HandleCountStart = $process.HandleCount
                $report.ThreadCountStart = $process.Threads.Count
            }
            $report.WorkingSetPeakBytes = [math]::Max([int64]$report.WorkingSetPeakBytes, $process.WorkingSet64)
            $report.PrivateMemoryPeakBytes = [math]::Max([int64]$report.PrivateMemoryPeakBytes, $process.PrivateMemorySize64)
            $report.HandleCountPeak = [math]::Max([int]$report.HandleCountPeak, $process.HandleCount)
            $report.ThreadCountPeak = [math]::Max([int]$report.ThreadCountPeak, $process.Threads.Count)
            Start-Sleep -Seconds 1
        } while ([DateTime]::UtcNow -lt $holdDeadline)
    }

    $report.ExitCommandPosted = [FxTask069.NativeMethods]::PostMessage($layoutWindow, 0x0111, [IntPtr]30120, [IntPtr]::Zero)
    if (-not $report.ExitCommandPosted) {
        throw 'The normal exit command could not be posted.'
    }

    if ($process.WaitForExit($ExitTimeoutSeconds * 1000)) {
        $report.ExitedNormally = $true
        $report.ExitSeconds = [math]::Round($watch.Elapsed.TotalSeconds, 3)
        $report.ExitCode = $process.ExitCode
    }
    else {
        $rundll32 = Join-Path $env:SystemRoot 'System32\rundll32.exe'
        $comsvcs = Join-Path $env:SystemRoot 'System32\comsvcs.dll'
        $dumpArgs = ('"{0}", MiniDump {1} "{2}" full' -f $comsvcs, $process.Id, $dumpPath)
        $dumpProcess = Start-Process -FilePath $rundll32 -ArgumentList $dumpArgs -PassThru -Wait
        $cdb = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\Debuggers\x64\cdb.exe'
        if ($dumpProcess.ExitCode -eq 0 -and (Test-Path -LiteralPath $dumpPath -PathType Leaf)) {
            $report.DumpPath = $dumpPath
        }

        if (Test-Path -LiteralPath $cdb -PathType Leaf) {
            $pdbRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'bin\x64\Release'
            if ($null -ne $report.DumpPath) {
                $commands = ".sympath `"$pdbRoot`"; .reload /f fxfile.exe; ~* kb; !locks; q"
                & $cdb -z $dumpPath -c $commands 2>&1 | Out-File -LiteralPath $stackPath -Encoding utf8
            }
            else {
                $commands = ".sympath `"$pdbRoot`"; .reload /f fxfile.exe; ~* kb; !locks; .detach; q"
                & $cdb -p $process.Id -c $commands 2>&1 | Out-File -LiteralPath $stackPath -Encoding utf8
            }
            $report.StackPath = $stackPath
        }
        elseif ($null -eq $report.DumpPath) {
            throw "MiniDump failed and cdb.exe is unavailable. MiniDumpExitCode=$($dumpProcess.ExitCode)"
        }
    }
}
finally {
    if ($null -ne $process -and -not $process.HasExited) {
        Stop-Process -Id $process.Id -Force
        $process.WaitForExit()
        $forced = $true
    }
    $report.ForcedTermination = $forced
    $report | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $reportPath -Encoding utf8
}

$report | ConvertTo-Json -Depth 5
Write-Host "REPORT=$reportPath"
