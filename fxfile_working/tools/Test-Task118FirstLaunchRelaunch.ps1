[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$EvidenceRoot,
    [Parameter(Mandatory = $true)] [string]$PackageRoot,
    [switch]$UsePackageInPlace,
    [ValidateRange(2, 5)] [int]$RunCount = 3,
    [ValidateRange(1, 6)] [int]$ExpectedPaneCount = 6,
    [ValidateRange(10, 180)] [int]$TimeoutSeconds = 90,
    [ValidateRange(20, 1000)] [int]$PollIntervalMilliseconds = 50
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$workspace = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$allowed = (Join-Path $workspace '__BUILD_TEMP_BACKUP__').TrimEnd('\') + '\'
$evidence = [IO.Path]::GetFullPath($EvidenceRoot)
$package = [IO.Path]::GetFullPath($PackageRoot)
if (-not $evidence.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {
    throw "EvidenceRoot is outside the approved task boundary: $evidence"
}
if (-not (Test-Path -LiteralPath (Join-Path $package 'fxfile.exe') -PathType Leaf)) {
    throw "FxFile package is incomplete: $package"
}
if (Test-Path -LiteralPath $evidence) {
    throw "Refusing to overwrite existing evidence: $evidence"
}

$existing = @(Get-Process -Name 'fxfile' -ErrorAction SilentlyContinue)
if ($existing.Count -ne 0) {
    throw "Close all FxFile processes before the test. Running PID(s): $($existing.Id -join ', ')"
}

New-Item -ItemType Directory -Path $evidence | Out-Null
$stage = $package
if (-not $UsePackageInPlace) {
    $stage = Join-Path $evidence 'package'
    New-Item -ItemType Directory -Path $stage | Out-Null
    Get-ChildItem -LiteralPath $package -Force | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $stage -Recurse -Force
    }
}

$oldTemp = $env:TEMP
$oldTmp = $env:TMP
$testTemp = Join-Path $evidence 'temp'
New-Item -ItemType Directory -Path $testTemp | Out-Null
$env:TEMP = $testTemp
$env:TMP = $testTemp

try {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class FxTask118TimingNative {
    public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr data);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc callback, IntPtr data);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern IntPtr GetProp(IntPtr hwnd, string name);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);

    public const uint WM_CLOSE = 0x0010;

    public static IntPtr FindFrame(int processId) {
        IntPtr result = IntPtr.Zero;
        int largestArea = 0;
        EnumWindows(delegate(IntPtr hwnd, IntPtr data) {
            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);
            if (pid != (uint)processId || !IsWindowVisible(hwnd)) return true;
            RECT rect;
            if (!GetWindowRect(hwnd, out rect)) return true;
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;
            int area = width * height;
            if (width >= 600 && height >= 400 && area > largestArea) {
                largestArea = area;
                result = hwnd;
            }
            return true;
        }, IntPtr.Zero);
        return result;
    }

    public static long Property(IntPtr hwnd, string name) {
        return hwnd == IntPtr.Zero ? 0 : GetProp(hwnd, name).ToInt64();
    }
}
'@

    $runs = [Collections.Generic.List[object]]::new()
    $exe = Join-Path $stage 'fxfile.exe'
    for ($runIndex = 1; $runIndex -le $RunCount; $runIndex++) {
        $process = $null
        try {
            $watch = [Diagnostics.Stopwatch]::StartNew()
            $process = Start-Process -FilePath $exe -WorkingDirectory $stage -PassThru
            $frame = [IntPtr]::Zero
            $frameMs = $null
            $skeletonMs = $null
            $readyMs = $null
            $readyCount = 0

            while ($watch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
                $process.Refresh()
                if ($process.HasExited) {
                    throw "FxFile exited during launch $runIndex with code $($process.ExitCode)."
                }
                if ($frame -eq [IntPtr]::Zero) {
                    $frame = [FxTask118TimingNative]::FindFrame($process.Id)
                    if ($frame -ne [IntPtr]::Zero) { $frameMs = $watch.ElapsedMilliseconds }
                }
                if ($frame -ne [IntPtr]::Zero) {
                    if ($null -eq $skeletonMs -and
                        [FxTask118TimingNative]::Property($frame, 'FxFile.StartupLayoutSkeletonPainted') -eq 1) {
                        $skeletonMs = $watch.ElapsedMilliseconds
                    }
                    $readyCount = [int][FxTask118TimingNative]::Property(
                        $frame, 'FxFile.StartupLayoutReadyViewCount')
                    if ($readyCount -ge $ExpectedPaneCount) {
                        $readyMs = $watch.ElapsedMilliseconds
                        break
                    }
                }
                Start-Sleep -Milliseconds $PollIntervalMilliseconds
            }
            if ($null -eq $readyMs) {
                throw "Launch $runIndex timed out at ready count $readyCount/$ExpectedPaneCount."
            }

            $process.Refresh()
            $runs.Add([pscustomobject]@{
                Run = $runIndex
                Kind = if ($runIndex -eq 1) { 'FirstProcessLaunch' } else { 'CloseAndRelaunch' }
                ProcessId = $process.Id
                FrameVisibleMilliseconds = $frameMs
                SkeletonPaintedMilliseconds = $skeletonMs
                AllPanesReadyMilliseconds = $readyMs
                ReadyPaneCount = $readyCount
                RespondingAtReady = $process.Responding
                WorkingSetMiB = [math]::Round($process.WorkingSet64 / 1MB, 1)
                CpuSeconds = [math]::Round($process.TotalProcessorTime.TotalSeconds, 3)
            })
        }
        finally {
            if ($null -ne $process) {
                try {
                    $process.Refresh()
                    if (-not $process.HasExited) {
                        $main = if ($process.MainWindowHandle -ne 0) {
                            [IntPtr]$process.MainWindowHandle
                        } else {
                            [FxTask118TimingNative]::FindFrame($process.Id)
                        }
                        if ($main -ne [IntPtr]::Zero) {
                            [FxTask118TimingNative]::PostMessage(
                                $main, [FxTask118TimingNative]::WM_CLOSE,
                                [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
                        }
                        if (-not $process.WaitForExit(10000)) {
                            Stop-Process -Id $process.Id -Force
                            $process.WaitForExit()
                        }
                    }
                } catch {
                    try { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue } catch {}
                }
            }
        }
        if ($runIndex -lt $RunCount) { Start-Sleep -Milliseconds 750 }
    }

    $relaunch = @($runs | Where-Object Kind -eq 'CloseAndRelaunch' |
        Select-Object -ExpandProperty AllPanesReadyMilliseconds | Sort-Object)
    $medianRelaunch = if (($relaunch.Count % 2) -eq 1) {
        [double]$relaunch[[int][math]::Floor($relaunch.Count / 2)]
    } else {
        ([double]$relaunch[$relaunch.Count / 2 - 1] +
         [double]$relaunch[$relaunch.Count / 2]) / 2.0
    }
    $firstReady = [double]$runs[0].AllPanesReadyMilliseconds
    $report = [ordered]@{
        Result = 'PASS'
        CapturedAt = (Get-Date).ToString('o')
        Scope = 'FxFile process first launch versus close-and-relaunch; Windows reboot excluded'
        Executable = $exe
        ExecutableSha256 = (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash
        ExpectedPaneCount = $ExpectedPaneCount
        RunCount = $RunCount
        Runs = @($runs)
        FirstReadyMilliseconds = $firstReady
        MedianRelaunchReadyMilliseconds = $medianRelaunch
        FirstMinusMedianRelaunchMilliseconds = $firstReady - $medianRelaunch
        StagingNote = if ($UsePackageInPlace) {
            'All launches use the requested package in place; no package staging copy was used.'
        } else {
            'The package was copied once before all launches; all runs use the identical staged executable and configuration.'
        }
    }
    $report | ConvertTo-Json -Depth 6 | Set-Content `
        -LiteralPath (Join-Path $evidence 'first_launch_relaunch_report.json') -Encoding utf8
    $report | ConvertTo-Json -Depth 6
}
finally {
    $env:TEMP = $oldTemp
    $env:TMP = $oldTmp
}
