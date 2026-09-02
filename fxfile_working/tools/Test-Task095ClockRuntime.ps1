[CmdletBinding()]
param(
    [string]$PackageRoot = 'D:\03 금일작업\00 임시\0000 FxFile\fxfile_run_x64',
    [string]$EvidenceRoot = 'D:\03 금일작업\00 임시\0000 FxFile\__BUILD_TEMP_BACKUP__',
    [string]$ReportPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$packagePath = [IO.Path]::GetFullPath($PackageRoot)
$evidencePath = [IO.Path]::GetFullPath($EvidenceRoot)
if (-not (Test-Path -LiteralPath (Join-Path $packagePath 'fxfile.exe') -PathType Leaf)) {
    throw "FxFile package is missing: $packagePath"
}
if (-not (Test-Path -LiteralPath $evidencePath -PathType Container)) {
    throw "Evidence root is missing: $evidencePath"
}

$probeName = 'task095_clock_runtime_probe_{0}_{1}' -f (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), $PID
$probeRoot = Join-Path $evidencePath $probeName
New-Item -ItemType Directory -Path $probeRoot | Out-Null
Get-ChildItem -LiteralPath $packagePath -Force | Copy-Item -Destination $probeRoot -Recurse -Force

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class Task095Win32 {
  public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr lp);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr lp);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
  [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
  [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr insertAfter, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint message, UIntPtr wParam, IntPtr lParam);
}
'@

function Get-MainWindow([int]$ProcessId, [int]$TimeoutSeconds) {
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        $script:mainWindow = [IntPtr]::Zero
        [Task095Win32]::EnumWindows({
            param($window, $parameter)
            $windowProcessId = 0
            [Task095Win32]::GetWindowThreadProcessId($window, [ref]$windowProcessId) | Out-Null
            if ($windowProcessId -eq $ProcessId -and [Task095Win32]::IsWindowVisible($window)) {
                $script:mainWindow = $window
                return $false
            }
            return $true
        }, [IntPtr]::Zero) | Out-Null
        if ($script:mainWindow -ne [IntPtr]::Zero) { return $script:mainWindow }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $deadline)
    return [IntPtr]::Zero
}

function Get-ClockState([IntPtr]$MainWindow, [string]$Stage) {
    $script:clockWindow = [IntPtr]::Zero
    [Task095Win32]::EnumChildWindows($MainWindow, {
        param($window, $parameter)
        if ([Task095Win32]::GetDlgCtrlID($window) -eq 1055) {
            $script:clockWindow = $window
            return $false
        }
        return $true
    }, [IntPtr]::Zero) | Out-Null

    if ($script:clockWindow -eq [IntPtr]::Zero) {
        return [pscustomobject]@{ Stage=$Stage; Found=$false; Visible=$false; Width=0; Height=0; ParentWidth=0; ParentHeight=0; InsideParent=$false }
    }

    $clockRect = New-Object Task095Win32+RECT
    $parentRect = New-Object Task095Win32+RECT
    [Task095Win32]::GetWindowRect($script:clockWindow, [ref]$clockRect) | Out-Null
    $parentWindow = [Task095Win32]::GetParent($script:clockWindow)
    [Task095Win32]::GetWindowRect($parentWindow, [ref]$parentRect) | Out-Null

    return [pscustomobject]@{
        Stage = $Stage
        Found = $true
        Visible = [Task095Win32]::IsWindowVisible($script:clockWindow)
        Width = $clockRect.Right - $clockRect.Left
        Height = $clockRect.Bottom - $clockRect.Top
        ParentWidth = $parentRect.Right - $parentRect.Left
        ParentHeight = $parentRect.Bottom - $parentRect.Top
        InsideParent = ($clockRect.Left -ge $parentRect.Left -and $clockRect.Right -le $parentRect.Right)
    }
}

$process = [Diagnostics.Process]::Start((Join-Path $probeRoot 'fxfile.exe'))
$forced = $false
try {
    $mainWindow = Get-MainWindow $process.Id 25
    if ($mainWindow -eq [IntPtr]::Zero) { throw 'FxFile main window was not found.' }

    Start-Sleep -Seconds 8
    $states = @()
    $states += Get-ClockState $mainWindow 'SavedLayout'

    [Task095Win32]::SetWindowPos($mainWindow, [IntPtr]::Zero, 40, 40, 1200, 800, 0x0014) | Out-Null
    Start-Sleep -Seconds 2
    $states += Get-ClockState $mainWindow 'Width1200'

    [Task095Win32]::SetWindowPos($mainWindow, [IntPtr]::Zero, 40, 40, 900, 700, 0x0014) | Out-Null
    Start-Sleep -Seconds 2
    $states += Get-ClockState $mainWindow 'Width900'

    [Task095Win32]::SetWindowPos($mainWindow, [IntPtr]::Zero, 40, 40, 600, 700, 0x0014) | Out-Null
    Start-Sleep -Seconds 2
    $states += Get-ClockState $mainWindow 'Width600'

    [Task095Win32]::SetWindowPos($mainWindow, [IntPtr]::Zero, 40, 40, 420, 700, 0x0014) | Out-Null
    Start-Sleep -Seconds 2
    $states += Get-ClockState $mainWindow 'Width420'

    $states | Format-Table -AutoSize
    $failed = @($states | Where-Object { -not $_.Found -or -not $_.Visible -or $_.Width -le 0 -or $_.Height -le 0 -or -not $_.InsideParent })

    [Task095Win32]::PostMessage($mainWindow, 0x0111, [UIntPtr]30120, [IntPtr]::Zero) | Out-Null
    if (-not $process.WaitForExit(15000)) { throw 'FxFile did not exit through command 30120.' }

    $runtimePassed = ($process.ExitCode -eq 0 -and $failed.Count -eq 0)
    if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
        $resolvedReportPath = [IO.Path]::GetFullPath($ReportPath)
        $reportParent = Split-Path -Parent $resolvedReportPath
        if (-not (Test-Path -LiteralPath $reportParent -PathType Container)) {
            throw "Runtime report parent is missing: $reportParent"
        }

        $report = [ordered]@{
            Task = 'Task095'
            CreatedAt = (Get-Date).ToString('o')
            PackageRoot = $packagePath
            FxFileSHA256 = (Get-FileHash -LiteralPath (Join-Path $packagePath 'fxfile.exe') -Algorithm SHA256).Hash
            ExitCode = $process.ExitCode
            RuntimePassed = $runtimePassed
            ForcedTermination = $false
            States = @($states)
        }
        [IO.File]::WriteAllText($resolvedReportPath, ($report | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
        "REPORT_PATH=$resolvedReportPath"
    }

    "EXIT_CODE=$($process.ExitCode)"
    "RUNTIME_PASS=$runtimePassed"
    "PROBE_ROOT=$probeRoot"
    if (-not $runtimePassed) { exit 1 }
}
finally {
    if (-not $process.HasExited) {
        $forced = $true
        $process.Kill()
        $process.WaitForExit()
    }
    "FORCED_TERMINATION=$forced"
}
