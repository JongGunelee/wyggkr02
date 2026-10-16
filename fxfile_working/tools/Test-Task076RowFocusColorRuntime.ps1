[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$PackageRoots,

    [Parameter(Mandatory = $true)]
    [string]$EvidenceRoot,

    [ValidateRange(10, 120)]
    [int]$HoldSeconds = 20,

    [ValidateRange(0, 120)]
    [int]$WarmupSeconds = 30,

    [ValidateRange(15, 180)]
    [int]$ReadyTimeoutSeconds = 90
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$workspaceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$allowedEvidenceBase = (Join-Path $workspaceRoot '__BUILD_TEMP_BACKUP__').TrimEnd('\')
$allowedEvidenceRoot = $allowedEvidenceBase + '\'
$resolvedEvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if (-not ($resolvedEvidenceRoot.Equals($allowedEvidenceBase, [StringComparison]::OrdinalIgnoreCase) -or
          $resolvedEvidenceRoot.StartsWith($allowedEvidenceRoot, [StringComparison]::OrdinalIgnoreCase))) {
    throw "EvidenceRoot must stay below the workspace backup boundary: $allowedEvidenceRoot"
}

if (@(Get-Process fxfile, 'fxfile-launcher', 'fxfile-upchecker', 'fxfile-updater' -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'Close every FxFile-related process before the isolated runtime comparison.'
}

if (-not ('FxTask076.NativeMethods' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace FxTask076
{
    public static class NativeMethods
    {
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool EnumChildWindows(IntPtr hWnd, EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder className, int maxCount);

        [DllImport("user32.dll")]
        public static extern bool PostMessage(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern uint GetGuiResources(IntPtr process, uint flags);
    }
}
'@
}

function Get-MainWindow([int]$processId) {
    $script:mainWindow = [IntPtr]::Zero
    $callback = [FxTask076.NativeMethods+EnumWindowsProc]{
        param([IntPtr]$window, [IntPtr]$unused)
        [uint32]$ownerProcessId = 0
        [void][FxTask076.NativeMethods]::GetWindowThreadProcessId($window, [ref]$ownerProcessId)
        if ($ownerProcessId -eq $processId -and [FxTask076.NativeMethods]::IsWindowVisible($window)) {
            $script:mainWindow = $window
            return $false
        }
        return $true
    }
    [void][FxTask076.NativeMethods]::EnumWindows($callback, [IntPtr]::Zero)
    return $script:mainWindow
}

function Get-VisibleListViewCount([IntPtr]$mainWindow) {
    $script:listViewCount = 0
    $callback = [FxTask076.NativeMethods+EnumWindowsProc]{
        param([IntPtr]$window, [IntPtr]$unused)
        $className = [Text.StringBuilder]::new(128)
        [void][FxTask076.NativeMethods]::GetClassName($window, $className, $className.Capacity)
        if ($className.ToString() -eq 'SysListView32' -and [FxTask076.NativeMethods]::IsWindowVisible($window)) {
            ++$script:listViewCount
        }
        return $true
    }
    [void][FxTask076.NativeMethods]::EnumChildWindows($mainWindow, $callback, [IntPtr]::Zero)
    return $script:listViewCount
}

function Get-TextEncoding([string]$path) {
    $bytes = [IO.File]::ReadAllBytes($path)
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xff -and $bytes[1] -eq 0xfe) { return [Text.Encoding]::Unicode }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xfe -and $bytes[1] -eq 0xff) { return [Text.Encoding]::BigEndianUnicode }
    return [Text.UTF8Encoding]::new($false)
}

function Set-ConfigValue([string]$configPath, [string]$key, [string]$value) {
    $encoding = Get-TextEncoding $configPath
    $text = [IO.File]::ReadAllText($configPath, $encoding)
    $pattern = '(?m)^' + [regex]::Escape($key) + '\s*=.*$'
    $replacement = "$key = $value"
    if ([regex]::IsMatch($text, $pattern)) {
        $text = [regex]::Replace($text, $pattern, $replacement)
    }
    else {
        if (-not $text.EndsWith("`r`n")) { $text += "`r`n" }
        $text += "$replacement`r`n"
    }
    [IO.File]::WriteAllText($configPath, $text, $encoding)
}

function Set-ScenarioConfig([string]$configPath, [string]$mainConfigPath, [bool]$fullRowFocus) {
    Set-ConfigValue $configPath 'config.file_list.full_row_select' $(if ($fullRowFocus) { '1' } else { '0' })
    for ($view = 1; $view -le 6; ++$view) {
        # Dark blue produces the cached white-text contrast path while the
        # OFF scenario verifies that the legacy list selection path remains idle.
        Set-ConfigValue $configPath "config.view$view.file_list.row_focus_color" '48,96,224'
    }

    # The runtime comparison must isolate row-focus rendering. A copied
    # profile may contain path locks for a removable/offline drive from a
    # different PC; let the package load its saved C:/Shell tabs instead.
    Set-ConfigValue $mainConfigPath 'main.view.path_locked' '0'
    # Keep the measurement independent from saved network/removable-folder
    # enumeration. This changes only the disposable scenario copy and gives
    # the renderer a bounded local data set in every pane.
    for ($view = 1; $view -le 6; ++$view) {
        Set-ConfigValue $configPath "config.view$view.file_list.init_folder" '1'
        Set-ConfigValue $configPath "config.view$view.file_list.init_folder_path" 'C:\Windows\Temp'
    }
}

function Get-ProcessSample([Diagnostics.Process]$process, [IntPtr]$mainWindow) {
    $process.Refresh()
    [pscustomobject]@{
        AtUtc = [DateTime]::UtcNow.ToString('o')
        Responding = $process.Responding
        VisibleFileLists = Get-VisibleListViewCount $mainWindow
        CpuSeconds = [math]::Round($process.TotalProcessorTime.TotalSeconds, 3)
        WorkingSetBytes = [int64]$process.WorkingSet64
        PrivateMemoryBytes = [int64]$process.PrivateMemorySize64
        HandleCount = [int]$process.HandleCount
        ThreadCount = [int]$process.Threads.Count
        GdiHandles = [int][FxTask076.NativeMethods]::GetGuiResources($process.Handle, 0)
        UserHandles = [int][FxTask076.NativeMethods]::GetGuiResources($process.Handle, 1)
    }
}

function Remove-IsolatedRoot([string]$path) {
    $resolved = [IO.Path]::GetFullPath($path)
    if (-not $resolved.StartsWith($allowedEvidenceRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove a path outside the test evidence boundary: $resolved"
    }
    if (Test-Path -LiteralPath $resolved -PathType Container) {
        $rootItem = Get-Item -LiteralPath $resolved -Force
        $hasReparsePoint = (($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) -or
            @(Get-ChildItem -LiteralPath $resolved -Recurse -Force | Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 }).Count -ne 0
        if ($hasReparsePoint) { throw "Refusing to remove an isolated test root containing a reparse point: $resolved" }
        Get-ChildItem -LiteralPath $resolved -Recurse -Force -File | ForEach-Object { $_.IsReadOnly = $false }
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

function Test-ProcessAlive([Diagnostics.Process]$process) {
    if ($null -eq $process) { return $false }
    try {
        $process.Refresh()
        return -not $process.HasExited
    }
    catch {
        # A normally exited process can disappear between Refresh and
        # HasExited.  Treat that race as already stopped so the isolated copy
        # is still removed and the runtime result remains authoritative.
        return $false
    }
}

function Request-NormalExit([IntPtr]$mainWindow) {
    if ($mainWindow -eq [IntPtr]::Zero) { return $false }
    # Use FxFile's own tray-exit command rather than WM_CLOSE so the normal
    # shutdown path used by the production smoke tests is exercised.
    [FxTask076.NativeMethods]::PostMessage($mainWindow, 0x0111, [IntPtr]30120, [IntPtr]::Zero)
}

function Invoke-Scenario([string]$packageRoot, [string]$architecture, [string]$scenarioName, [bool]$fullRowFocus, [string]$runRoot) {
    $scenarioRoot = Join-Path $runRoot ("{0}_{1}" -f $architecture, $scenarioName)
    $process = $null
    $mainWindow = [IntPtr]::Zero
    $forcedTermination = $false
    $completed = $false
    $samples = [Collections.Generic.List[object]]::new()
    try {
        Copy-Item -LiteralPath $packageRoot -Destination $scenarioRoot -Recurse
        $exePath = Join-Path $scenarioRoot 'fxfile.exe'
        $configPath = Join-Path $scenarioRoot 'fxfile\fxfile.conf'
        $mainConfigPath = Join-Path $scenarioRoot 'fxfile\fxfile-main.conf'
        if (-not (Test-Path -LiteralPath $exePath -PathType Leaf)) { throw "Missing staged executable: $exePath" }
        if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) { throw "Missing staged configuration: $configPath" }
        if (-not (Test-Path -LiteralPath $mainConfigPath -PathType Leaf)) { throw "Missing staged main configuration: $mainConfigPath" }
        if ((Test-Path -LiteralPath (Join-Path $scenarioRoot 'fxfile.ini')) -or (Test-Path -LiteralPath (Join-Path $scenarioRoot '.fxfile'))) {
            throw 'The isolated package contains a forbidden root pointer.'
        }
        Set-ScenarioConfig $configPath $mainConfigPath $fullRowFocus

        $watch = [Diagnostics.Stopwatch]::StartNew()
        $process = Start-Process -FilePath $exePath -WorkingDirectory $scenarioRoot -ArgumentList '-w 2x2' -PassThru
        $readyDeadline = [DateTime]::UtcNow.AddSeconds($ReadyTimeoutSeconds)
        do {
            Start-Sleep -Milliseconds 250
            if (-not (Test-ProcessAlive $process)) { throw "FxFile exited before ready." }
            $process.Refresh()
            $mainWindow = Get-MainWindow $process.Id
            if ($mainWindow -ne [IntPtr]::Zero -and $process.Responding -and (Get-VisibleListViewCount $mainWindow) -ge 4) { break }
        } while ([DateTime]::UtcNow -lt $readyDeadline)
        if ($mainWindow -eq [IntPtr]::Zero -or -not $process.Responding -or (Get-VisibleListViewCount $mainWindow) -lt 4) {
            throw 'FxFile did not reach a responsive 2x2 layout before the ready timeout.'
        }
        [void][FxTask076.NativeMethods]::SetForegroundWindow($mainWindow)
        $readySeconds = [math]::Round($watch.Elapsed.TotalSeconds, 3)
        # Four visible list views are the boot/readiness gate, but Shell
        # enumeration can continue briefly after that point.  Separate this
        # bounded warm-up from steady-state so startup work is recorded rather
        # than misclassified as a leak or repaint loop.
        $warmupStart = Get-ProcessSample $process $mainWindow
        for ($sample = 1; $sample -le ($WarmupSeconds * 2); ++$sample) {
            Start-Sleep -Milliseconds 500
            if (-not (Test-ProcessAlive $process)) { throw 'FxFile exited during the bounded warm-up.' }
            $warmupCurrent = Get-ProcessSample $process $mainWindow
            if (-not $warmupCurrent.Responding) { throw 'FxFile reported Not Responding during the bounded warm-up.' }
            if ($warmupCurrent.VisibleFileLists -lt 4) { throw 'A 2x2 Explorer pane disappeared during the bounded warm-up.' }
        }
        $warmupEnd = Get-ProcessSample $process $mainWindow
        $start = Get-ProcessSample $process $mainWindow
        $samples.Add($start)

        for ($sample = 1; $sample -le ($HoldSeconds * 2); ++$sample) {
            Start-Sleep -Milliseconds 500
            if ($process.HasExited) { throw "FxFile exited during the hold. ExitCode=$($process.ExitCode)" }
            $current = Get-ProcessSample $process $mainWindow
            $samples.Add($current)
            if (-not $current.Responding) { throw 'FxFile reported Not Responding during the steady-state hold.' }
            if ($current.VisibleFileLists -lt 4) { throw 'A 2x2 Explorer pane disappeared during the steady-state hold.' }
        }
        $end = $samples[$samples.Count - 1]
        $cpuDelta = $end.CpuSeconds - $start.CpuSeconds
        $privateDelta = $end.PrivateMemoryBytes - $start.PrivateMemoryBytes
        $workingSetDelta = $end.WorkingSetBytes - $start.WorkingSetBytes
        $handleDelta = $end.HandleCount - $start.HandleCount
        $gdiDelta = $end.GdiHandles - $start.GdiHandles
        $userDelta = $end.UserHandles - $start.UserHandles
        if ($cpuDelta -gt ($HoldSeconds * 0.25)) { throw "[$architecture/$scenarioName] Unexpected steady-state CPU accumulation: $cpuDelta CPU seconds over $HoldSeconds seconds." }
        if ($privateDelta -gt 96MB -or $workingSetDelta -gt 128MB) { throw "[$architecture/$scenarioName] Unexpected memory accumulation: private=$privateDelta bytes, working-set=$workingSetDelta bytes." }
        if ($handleDelta -gt 64 -or $gdiDelta -gt 32 -or $userDelta -gt 32) { throw "[$architecture/$scenarioName] Unexpected handle accumulation: handle=$handleDelta, GDI=$gdiDelta, USER=$userDelta." }

        if (-not (Request-NormalExit $mainWindow)) { throw 'The normal FxFile exit command could not be posted.' }
        if (-not $process.WaitForExit(20000)) { throw 'FxFile did not close normally within 20 seconds.' }
        if ($process.ExitCode -ne 0) { throw "FxFile returned exit code $($process.ExitCode)." }
        $result = [pscustomobject]@{
            Result = 'PASS'; Architecture = $architecture; Scenario = $scenarioName; FullRowFocus = $fullRowFocus
            ExecutableSha256 = (Get-FileHash -LiteralPath $exePath -Algorithm SHA256).Hash; ReadySeconds = $readySeconds
            HoldSeconds = $HoldSeconds; WarmupSeconds = $WarmupSeconds
            WarmupCpuDeltaSeconds = [math]::Round(($warmupEnd.CpuSeconds - $warmupStart.CpuSeconds), 3)
            WarmupPrivateMemoryDeltaBytes = ($warmupEnd.PrivateMemoryBytes - $warmupStart.PrivateMemoryBytes)
            WarmupWorkingSetDeltaBytes = ($warmupEnd.WorkingSetBytes - $warmupStart.WorkingSetBytes)
            CpuDeltaSeconds = [math]::Round($cpuDelta, 3)
            PrivateMemoryDeltaBytes = $privateDelta; WorkingSetDeltaBytes = $workingSetDelta
            HandleDelta = $handleDelta; GdiHandleDelta = $gdiDelta; UserHandleDelta = $userDelta
            PeakPrivateMemoryBytes = ($samples | Measure-Object -Property PrivateMemoryBytes -Maximum).Maximum
            PeakWorkingSetBytes = ($samples | Measure-Object -Property WorkingSetBytes -Maximum).Maximum
            Samples = $samples; ExitCode = $process.ExitCode; ForcedTermination = $false
        }
        $completed = $true
        return $result
    }
    finally {
        if (Test-ProcessAlive $process) {
            try { [void](Request-NormalExit $mainWindow) } catch { }
            try {
                if (-not $process.WaitForExit(5000) -and (Test-ProcessAlive $process)) {
                    Stop-Process -Id $process.Id -Force
                    $forcedTermination = $true
                }
            }
            catch {
                if (Test-ProcessAlive $process) {
                    Stop-Process -Id $process.Id -Force
                    $forcedTermination = $true
                }
            }
        }
        Remove-IsolatedRoot $scenarioRoot
        if ($forcedTermination -and $completed) { throw 'FxFile required forced termination during the isolated runtime test.' }
    }
}

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$runRoot = Join-Path $resolvedEvidenceRoot ("task076_row_focus_color_$stamp")
New-Item -ItemType Directory -Path $runRoot -Force | Out-Null
$results = [Collections.Generic.List[object]]::new()
try {
    foreach ($packageRootInput in $PackageRoots) {
        $packageRoot = [IO.Path]::GetFullPath($packageRootInput)
        $exePath = Join-Path $packageRoot 'fxfile.exe'
        if (-not (Test-Path -LiteralPath $exePath -PathType Leaf)) { throw "Package root is missing fxfile.exe: $packageRoot" }
        $architecture = if ((Get-Item -LiteralPath $exePath).Length -gt 0) { Split-Path -Leaf $packageRoot } else { 'unknown' }
        $results.Add((Invoke-Scenario $packageRoot $architecture 'LegacyFocusOff' $false $runRoot))
        $results.Add((Invoke-Scenario $packageRoot $architecture 'CustomRowFocusOn' $true $runRoot))
    }
    $onByArchitecture = @($results | Where-Object { $_.Scenario -eq 'CustomRowFocusOn' } | Group-Object Architecture)
    foreach ($group in $onByArchitecture) {
        $off = @($results | Where-Object { $_.Architecture -eq $group.Name -and $_.Scenario -eq 'LegacyFocusOff' })[0]
        $on = $group.Group[0]
        if (($on.ReadySeconds - $off.ReadySeconds) -gt 3.0) {
            throw "Custom row-focus startup regression exceeded 3 seconds for $($group.Name): off=$($off.ReadySeconds), on=$($on.ReadySeconds)."
        }
    }
    $report = [ordered]@{ Result = 'PASS'; CapturedAt = (Get-Date).ToString('o'); HoldSeconds = $HoldSeconds; WarmupSeconds = $WarmupSeconds; Results = $results }
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $runRoot 'runtime_report.json') -Encoding UTF8
    $report | ConvertTo-Json -Depth 5
}
catch {
    [ordered]@{ Result = 'FAIL'; CapturedAt = (Get-Date).ToString('o'); Message = $_.Exception.Message; Results = $results } |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $runRoot 'runtime_report.json') -Encoding UTF8
    throw
}
finally {
    $remaining = @(Get-ChildItem -LiteralPath $runRoot -Directory -ErrorAction SilentlyContinue)
    if ($remaining.Count -ne 0) { throw "Isolated runtime package cleanup was incomplete: $($remaining.FullName -join ', ')" }
}
