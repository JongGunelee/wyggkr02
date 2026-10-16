[CmdletBinding()]
param(
    [string]$PackageRoot = 'C:\Users\PC\Downloads\0000 FxFile\fxfile_run_x64',
    [string]$TestFolder = 'D:\03 금일작업\00 임시\00000 스크립트\01 Scripts\automated_scripts',
    [Parameter(Mandatory = $true)]
    [string]$EvidenceRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class FxTask060Win32 {
    public delegate bool EnumProc(IntPtr hwnd, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr hwnd, EnumProc cb, IntPtr data);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, StringBuilder text, int max);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);
}
'@

function Get-MainWindow([int]$ProcessId) {
    $script:mainWindow = [IntPtr]::Zero
    $callback = [FxTask060Win32+EnumProc]{
        param([IntPtr]$hwnd, [IntPtr]$data)
        [uint32]$ownerProcessId = 0
        [void][FxTask060Win32]::GetWindowThreadProcessId($hwnd, [ref]$ownerProcessId)
        if ($ownerProcessId -eq $ProcessId -and [FxTask060Win32]::IsWindowVisible($hwnd)) {
            $script:mainWindow = $hwnd
            return $false
        }
        return $true
    }
    [void][FxTask060Win32]::EnumWindows($callback, [IntPtr]::Zero)
    return $script:mainWindow
}

function Get-VisibleListViewCount([IntPtr]$MainWindow) {
    $script:listViewCount = 0
    $callback = [FxTask060Win32+EnumProc]{
        param([IntPtr]$hwnd, [IntPtr]$data)
        $className = [Text.StringBuilder]::new(128)
        [void][FxTask060Win32]::GetClassName($hwnd, $className, $className.Capacity)
        if ($className.ToString() -eq 'SysListView32' -and [FxTask060Win32]::IsWindowVisible($hwnd)) {
            $script:listViewCount++
        }
        return $true
    }
    [void][FxTask060Win32]::EnumChildWindows($MainWindow, $callback, [IntPtr]::Zero)
    return $script:listViewCount
}

$package = [IO.Path]::GetFullPath($PackageRoot)
$testPath = [IO.Path]::GetFullPath($TestFolder)
$evidence = [IO.Path]::GetFullPath($EvidenceRoot)
$workspace = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$allowedPrefix = (Join-Path $workspace '__BUILD_TEMP_BACKUP__').TrimEnd('\') + '\'
if (-not $evidence.StartsWith($allowedPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Evidence must stay under the workspace backup boundary: $evidence"
}
if (-not (Test-Path -LiteralPath $testPath -PathType Container)) {
    throw "Test folder does not exist: $testPath"
}

New-Item -ItemType Directory -Path $evidence -Force | Out-Null
$smoke = Join-Path $evidence 'isolated_package'
if (Test-Path -LiteralPath $smoke) { throw "Smoke path already exists: $smoke" }
Copy-Item -LiteralPath $package -Destination $smoke -Recurse

$exe = Join-Path $smoke 'fxfile.exe'
if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { throw "Missing smoke executable: $exe" }
if ((Test-Path -LiteralPath (Join-Path $smoke 'fxfile.ini')) -or
    (Test-Path -LiteralPath (Join-Path $smoke '.fxfile'))) {
    throw 'The isolated package contains a forbidden root pointer.'
}

$process = $null
$main = [IntPtr]::Zero
$forced = $false
$samples = @()
$failure = $null
try {
    $quoted = '"' + $testPath.Replace('"', '\"') + '"'
    $arguments = "-w 2x2 --dir1 $quoted --dir2 $quoted --dir3 $quoted --dir4 $quoted"
    $process = Start-Process -FilePath $exe -WorkingDirectory $smoke -ArgumentList $arguments -PassThru
    $deadline = (Get-Date).AddSeconds(60)
    $viewCount = 0
    do {
        Start-Sleep -Milliseconds 250
        $process.Refresh()
        $main = Get-MainWindow $process.Id
        if ($main -ne [IntPtr]::Zero) { $viewCount = Get-VisibleListViewCount $main }
    } while (($main -eq [IntPtr]::Zero -or $viewCount -lt 4 -or -not $process.Responding) -and
             (Get-Date) -lt $deadline -and -not $process.HasExited)
    if ($process.HasExited) { throw "FxFile exited early: $($process.ExitCode)" }
    if ($main -eq [IntPtr]::Zero -or $viewCount -lt 4) { throw "Four panes did not become visible; count=$viewCount" }

    # Four controls can become visible before initial Shell metadata and the
    # responsive-column pass have settled.  Keep pumping a responsiveness
    # observation window first; the historical defect never settled and made
    # the main window stop responding, whereas legitimate startup work does.
    $consecutiveNonResponding = 0
    $maxConsecutiveNonResponding = 0
    for ($i = 0; $i -lt 40; $i++) {
        Start-Sleep -Milliseconds 500
        $process.Refresh()
        if ($process.Responding) {
            $consecutiveNonResponding = 0
        }
        else {
            $consecutiveNonResponding++
            $maxConsecutiveNonResponding = [math]::Max($maxConsecutiveNonResponding,
                                                       $consecutiveNonResponding)
        }
        $samples += [pscustomobject]@{
            At = (Get-Date).ToString('o')
            Phase = 'Warmup'
            Responding = $process.Responding
            CpuSeconds = $process.TotalProcessorTime.TotalSeconds
            ViewCount = Get-VisibleListViewCount $main
        }
        # A fully loaded 2x2 skeleton can briefly cross a Shell/AV scheduling
        # boundary on a saturated PC.  Treat a continuous 10-second stall as
        # the hang defect, then require the final five seconds to be stable.
        if ($consecutiveNonResponding -ge 20) {
            throw 'FxFile remained non-responsive for 10 seconds while Shell metadata was settling.'
        }
    }
    if (@($samples | Select-Object -Last 10 | Where-Object { -not $_.Responding }).Count -ne 0) {
        throw 'FxFile did not reach a stable responding state after the warmup window.'
    }

    $process.Refresh()
    $cpuBefore = $process.TotalProcessorTime.TotalSeconds
    for ($i = 0; $i -lt 20; $i++) {
        Start-Sleep -Milliseconds 250
        $process.Refresh()
        $samples += [pscustomobject]@{
            At = (Get-Date).ToString('o')
            Phase = 'SteadyState'
            Responding = $process.Responding
            CpuSeconds = $process.TotalProcessorTime.TotalSeconds
            ViewCount = Get-VisibleListViewCount $main
        }
    }
    $process.Refresh()
    $cpuDelta = $process.TotalProcessorTime.TotalSeconds - $cpuBefore
    if (@($samples | Where-Object { -not $_.Responding }).Count -ne 0) { throw 'FxFile became non-responsive during the post-navigation observation.' }
    if (@($samples | Where-Object { $_.ViewCount -lt 4 }).Count -ne 0) { throw 'A 2x2 pane disappeared during the observation.' }
    if ($cpuDelta -ge 2.5) { throw "Icon/overlay feedback loop suspected after warmup: CPU delta=$cpuDelta seconds over 5 seconds." }

    [void][FxTask060Win32]::PostMessage($main, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
    if (-not $process.WaitForExit(20000)) { throw 'FxFile did not close normally within 20 seconds.' }
    if ($process.ExitCode -ne 0) { throw "FxFile exit code was $($process.ExitCode)." }

    $report = [ordered]@{
        Result = 'PASS'
        CapturedAt = (Get-Date).ToString('o')
        PackageSHA256 = (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash
        TestFolder = $testPath
        ViewCount = $viewCount
        WarmupSeconds = 20
        ObservationSeconds = 5
        CpuDeltaSeconds = [math]::Round($cpuDelta, 3)
        NonRespondingSamples = @($samples | Where-Object { -not $_.Responding }).Count
        MaxConsecutiveNonRespondingSamples = $maxConsecutiveNonResponding
        ExitCode = $process.ExitCode
        ForcedTermination = $false
        RootPointerFilesCreated = ((Test-Path -LiteralPath (Join-Path $smoke 'fxfile.ini')) -or (Test-Path -LiteralPath (Join-Path $smoke '.fxfile')))
        Samples = $samples
    }
    $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $evidence 'runtime_report.json') -Encoding UTF8
    $report | ConvertTo-Json -Depth 4
}
catch {
    $failure = $_
    [ordered]@{
        Result = 'FAIL'
        CapturedAt = (Get-Date).ToString('o')
        Message = $_.Exception.Message
        ScriptStackTrace = $_.ScriptStackTrace
        ProcessId = $(if ($process) { $process.Id } else { $null })
        MainWindow = $main.ToInt64()
        ForcedTermination = $forced
        Samples = $samples
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $evidence 'runtime_report.json') -Encoding UTF8
}
finally {
    if ($process -and -not $process.HasExited) {
        if ($main -ne [IntPtr]::Zero) { [void][FxTask060Win32]::PostMessage($main, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) }
        if (-not $process.WaitForExit(5000)) {
            Stop-Process -Id $process.Id -Force
            $forced = $true
        }
    }
    if (Test-Path -LiteralPath $smoke -PathType Container) {
        $item = Get-Item -LiteralPath $smoke -Force
        $hasReparse = (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) -or
            @(Get-ChildItem -LiteralPath $smoke -Recurse -Force | Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 }).Count -gt 0
        if ($hasReparse) { throw "Refusing to clean a smoke package containing a reparse point: $smoke" }
        Get-ChildItem -LiteralPath $smoke -Recurse -Force -File | ForEach-Object { $_.IsReadOnly = $false }
        $removed = $false
        for ($attempt = 0; $attempt -lt 10 -and -not $removed; $attempt++) {
            try {
                Remove-Item -LiteralPath $smoke -Recurse -Force -ErrorAction Stop
                $removed = -not (Test-Path -LiteralPath $smoke)
            }
            catch {
                if ($attempt -eq 9) { throw }
                Start-Sleep -Milliseconds 500
            }
        }
    }
}
if ($failure) { throw $failure }
