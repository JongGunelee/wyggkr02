[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PackageRoot,

    [Parameter(Mandatory = $true)]
    [string]$EvidenceRoot,

    [ValidateRange(30, 900)]
    [int]$DurationSeconds = 120,

    [ValidateRange(30, 180)]
    [int]$ReadyTimeoutSeconds = 120,

    [ValidateRange(0, 240)]
    [int]$RefreshEveryTransitions = 24,

    [switch]$HideParentFolder,

    [switch]$HideStatusBar
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$workspaceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$allowedEvidenceBase = (Join-Path $workspaceRoot '__BUILD_TEMP_BACKUP__').TrimEnd('\')
$allowedEvidencePrefix = $allowedEvidenceBase + '\'
$resolvedEvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if (-not ($resolvedEvidenceRoot.Equals($allowedEvidenceBase, [StringComparison]::OrdinalIgnoreCase) -or
          $resolvedEvidenceRoot.StartsWith($allowedEvidencePrefix, [StringComparison]::OrdinalIgnoreCase))) {
    throw "EvidenceRoot must stay below $allowedEvidenceBase"
}

if (@(Get-Process fxfile, 'fxfile-launcher', 'fxfile-upchecker', 'fxfile-updater' -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'Close every FxFile-related process before the isolated six-pane test.'
}

if (-not ('FxTask111.NativeMethods' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace FxTask111
{
    public static class NativeMethods
    {
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }

        [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr value);
        [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr value);
        [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr window, StringBuilder name, int maximum);
        [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr window, out RECT bounds);
        [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr window, out RECT bounds);
        [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr window);
        [DllImport("user32.dll")] public static extern uint GetGuiResources(IntPtr process, uint flags);
    }
}
'@
}

function Get-MainWindow([int]$processId) {
    $script:foundMain = [IntPtr]::Zero
    $callback = [FxTask111.NativeMethods+EnumWindowsProc]{
        param([IntPtr]$window, [IntPtr]$unused)
        [uint32]$owner = 0
        [void][FxTask111.NativeMethods]::GetWindowThreadProcessId($window, [ref]$owner)
        if ($owner -eq $processId -and [FxTask111.NativeMethods]::IsWindowVisible($window)) {
            $script:foundMain = $window
            return $false
        }
        $true
    }
    [void][FxTask111.NativeMethods]::EnumWindows($callback, [IntPtr]::Zero)
    $script:foundMain
}

function Get-VisibleListViews([IntPtr]$mainWindow) {
    $script:foundLists = [Collections.Generic.List[object]]::new()
    $callback = [FxTask111.NativeMethods+EnumWindowsProc]{
        param([IntPtr]$window, [IntPtr]$unused)
        $name = [Text.StringBuilder]::new(128)
        [void][FxTask111.NativeMethods]::GetClassName($window, $name, $name.Capacity)
        if ($name.ToString() -eq 'SysListView32' -and [FxTask111.NativeMethods]::IsWindowVisible($window)) {
            $bounds = [FxTask111.NativeMethods+RECT]::new()
            if ([FxTask111.NativeMethods]::GetWindowRect($window, [ref]$bounds)) {
                $script:foundLists.Add([pscustomobject]@{ Handle = $window; Left = $bounds.Left; Top = $bounds.Top })
            }
        }
        $true
    }
    [void][FxTask111.NativeMethods]::EnumChildWindows($mainWindow, $callback, [IntPtr]::Zero)
    @($script:foundLists | Sort-Object Top, Left)
}

function Get-StableVisibleListViews([IntPtr]$mainWindow, [int]$expectedCount = 6) {
    [object[]]$lists = Get-VisibleListViews $mainWindow
    # A refresh intentionally hides the ListView while replacing its items.
    # Retry that bounded transition instead of reporting a permanent pane loss.
    for ($retry = 0; @($lists).Count -ne $expectedCount -and $retry -lt 20; ++$retry) {
        Start-Sleep -Milliseconds 100
        [object[]]$lists = Get-VisibleListViews $mainWindow
    }
    @($lists)
}

function Get-Encoding([string]$path) {
    $bytes = [IO.File]::ReadAllBytes($path)
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xff -and $bytes[1] -eq 0xfe) { return [Text.Encoding]::Unicode }
    [Text.UTF8Encoding]::new($false)
}

function Set-ConfigValue([string]$path, [string]$key, [string]$value) {
    $encoding = Get-Encoding $path
    $text = [IO.File]::ReadAllText($path, $encoding)
    $pattern = '(?m)^' + [regex]::Escape($key) + '\s*=.*$'
    $replacement = "$key = $value"
    if ([regex]::IsMatch($text, $pattern)) { $text = [regex]::Replace($text, $pattern, $replacement) }
    else {
        if (-not $text.EndsWith("`r`n")) { $text += "`r`n" }
        $text += "$replacement`r`n"
    }
    [IO.File]::WriteAllText($path, $text, $encoding)
}

function Get-Sample([Diagnostics.Process]$process, [IntPtr]$mainWindow, [int]$transition) {
    $process.Refresh()
    [pscustomobject]@{
        AtUtc = [DateTime]::UtcNow.ToString('o')
        Transition = $transition
        Responding = $process.Responding
        VisibleFileLists = @(Get-VisibleListViews $mainWindow).Count
        CpuSeconds = [math]::Round($process.TotalProcessorTime.TotalSeconds, 3)
        WorkingSetBytes = [int64]$process.WorkingSet64
        PrivateMemoryBytes = [int64]$process.PrivateMemorySize64
        HandleCount = [int]$process.HandleCount
        ThreadCount = [int]$process.Threads.Count
        GdiHandles = [int][FxTask111.NativeMethods]::GetGuiResources($process.Handle, 0)
        UserHandles = [int][FxTask111.NativeMethods]::GetGuiResources($process.Handle, 1)
    }
}

function Get-ListViewItemCount([IntPtr]$listView) {
    [int][FxTask111.NativeMethods]::SendMessage($listView, 0x1004, [IntPtr]::Zero, [IntPtr]::Zero).ToInt64()
}

function Get-FirstSelectedIndex([IntPtr]$listView) {
    # LVM_GETNEXTITEM does not pass a caller buffer and is safe to use across
    # the process boundary.  LVNI_SELECTED proves that the intended mixed
    # fixture row, rather than just an arbitrary blank coordinate, was hit.
    [int][FxTask111.NativeMethods]::SendMessage($listView, 0x100C, [IntPtr](-1), [IntPtr]2).ToInt64()
}

function Post-Click([IntPtr]$listView, [int]$itemIndex) {
    $client = [FxTask111.NativeMethods+RECT]::new()
    if (-not [FxTask111.NativeMethods]::GetClientRect($listView, [ref]$client)) { return $false }
    $itemCount = Get-ListViewItemCount $listView
    if ($itemIndex -lt 0 -or $itemIndex -ge $itemCount) { return $false }

    # Derive a row centre from the report view itself instead of assuming one
    # fixed DPI.  LVM_GETCOUNTPERPAGE excludes the header from its row count.
    $headerHeight = 0
    $header = [FxTask111.NativeMethods]::SendMessage($listView, 0x101F, [IntPtr]::Zero, [IntPtr]::Zero)
    if ($header -ne [IntPtr]::Zero) {
        $headerBounds = [FxTask111.NativeMethods+RECT]::new()
        if ([FxTask111.NativeMethods]::GetWindowRect($header, [ref]$headerBounds)) {
            $headerHeight = [Math]::Max(0, $headerBounds.Bottom - $headerBounds.Top)
        }
    }
    $rowsPerPage = [int][FxTask111.NativeMethods]::SendMessage($listView, 0x1028, [IntPtr]::Zero, [IntPtr]::Zero).ToInt64()
    $dataHeight = [Math]::Max(1, $client.Bottom - $headerHeight)
    $rowHeight = if ($rowsPerPage -gt 0) { [Math]::Max(16, [Math]::Floor($dataHeight / $rowsPerPage)) } else { 20 }
    $x = [Math]::Max(12, [Math]::Min(80, $client.Right - 8))
    $y = [Math]::Max(8, [Math]::Min($headerHeight + ($itemIndex * $rowHeight) + [Math]::Floor($rowHeight / 2), $client.Bottom - 8))
    $packed = [IntPtr](($y -shl 16) -bor ($x -band 0xffff))
    $down = [FxTask111.NativeMethods]::PostMessage($listView, 0x0201, [IntPtr]1, $packed)
    $up = [FxTask111.NativeMethods]::PostMessage($listView, 0x0202, [IntPtr]::Zero, $packed)
    if (-not ($down -and $up)) { return $false }
    for ($retry = 0; $retry -lt 20; ++$retry) {
        Start-Sleep -Milliseconds 25
        if ((Get-FirstSelectedIndex $listView) -eq $itemIndex) { return $true }
    }
    $false
}

function Select-ListViewItem([IntPtr]$listView, [int]$itemIndex) {
    if (Post-Click $listView $itemIndex) { return $true }

    # A narrow pane can make a lower row only partially visible.  Exercise the
    # same native ListView selection path with HOME/DOWN as a DPI-independent
    # fallback, then verify the exact selected index through LVM_GETNEXTITEM.
    foreach ($message in @(0x0100, 0x0101)) {
        [void][FxTask111.NativeMethods]::PostMessage($listView, $message, [IntPtr]0x24, [IntPtr]::Zero)
    }
    for ($step = 0; $step -lt $itemIndex; ++$step) {
        foreach ($message in @(0x0100, 0x0101)) {
            [void][FxTask111.NativeMethods]::PostMessage($listView, $message, [IntPtr]0x28, [IntPtr]::Zero)
        }
    }
    for ($retry = 0; $retry -lt 40; ++$retry) {
        Start-Sleep -Milliseconds 25
        if ((Get-FirstSelectedIndex $listView) -eq $itemIndex) { return $true }
    }
    $false
}

function Remove-IsolatedRoot([string]$path) {
    $resolved = [IO.Path]::GetFullPath($path)
    if (-not $resolved.StartsWith($allowedEvidencePrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove outside evidence boundary: $resolved"
    }
    if (Test-Path -LiteralPath $resolved -PathType Container) {
        $items = @(Get-ChildItem -LiteralPath $resolved -Recurse -Force)
        if (@($items | Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 }).Count -ne 0) {
            throw "Refusing to remove an isolated root containing a reparse point: $resolved"
        }
        $items | Where-Object { -not $_.PSIsContainer } | ForEach-Object { $_.IsReadOnly = $false }
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

$package = [IO.Path]::GetFullPath($PackageRoot)
$exe = Join-Path $package 'fxfile.exe'
if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { throw "Missing fxfile.exe: $exe" }

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$runRoot = Join-Path $resolvedEvidenceRoot "task111_six_pane_long_run_$stamp"
$stageRoot = Join-Path $runRoot 'isolated_package'
$fixtureRoot = Join-Path $runRoot 'mixed_file_and_folder_fixture'
$reportPath = Join-Path $runRoot 'runtime_report.json'
New-Item -ItemType Directory -Path $runRoot -Force | Out-Null
$process = $null
$mainWindow = [IntPtr]::Zero
$samples = [Collections.Generic.List[object]]::new()
$result = 'FAIL'
$failure = $null
$fixtureFolderNames = @('00_Folder_A', '01_Folder_B', '02_Folder_C', '03_Folder_D')
$fixtureFileNames = @(
    '10_Text_A.txt',
    '11_Data_B.csv',
    '12_Page_C.html',
    '13_Config_D.ini',
    '14_NoExtension',
    '15_Unicode_한글.txt',
    '16_ZeroLength.bin',
    '17_Long_File_Name_For_Report_Row_Repaint_E.txt'
)
$folderSelectionsByPane = [int[]]::new(6)
$fileSelectionsByPane = [int[]]::new(6)

try {
    New-Item -ItemType Directory -Path $fixtureRoot -Force | Out-Null
    foreach ($name in $fixtureFolderNames) {
        New-Item -ItemType Directory -Path (Join-Path $fixtureRoot $name) -Force | Out-Null
    }
    foreach ($name in $fixtureFileNames) {
        $filePath = Join-Path $fixtureRoot $name
        if ($name -eq '16_ZeroLength.bin') { [IO.File]::WriteAllBytes($filePath, [byte[]]::new(0)) }
        else { [IO.File]::WriteAllText($filePath, "Task111 mixed-row repaint fixture: $name`r`n", [Text.UTF8Encoding]::new($false)) }
    }

    Copy-Item -LiteralPath $package -Destination $stageRoot -Recurse
    $stageExe = Join-Path $stageRoot 'fxfile.exe'
    $config = Join-Path $stageRoot 'fxfile\fxfile.conf'
    $mainConfig = Join-Path $stageRoot 'fxfile\fxfile-main.conf'
    # The saved tab path has priority over the per-view initial-folder option.
    # Pin both sources so an existing user's history cannot leak into this
    # isolated verification and accidentally turn it into a folder-only run.
    Set-ConfigValue $mainConfig 'main.view.path_locked' '1'
    Set-ConfigValue $config 'config.file_list.full_row_select' '1'
    Set-ConfigValue $config 'config.file_list.show_parent_folder' $(if ($HideParentFolder) { '0' } else { '1' })
    Set-ConfigValue $config 'config.status_bar.show' $(if ($HideStatusBar) { '0' } else { '1' })
    for ($view = 1; $view -le 6; ++$view) {
        Set-ConfigValue $mainConfig "main.view$view.tab1.path" $fixtureRoot
        Set-ConfigValue $mainConfig "main.view$view.locked_path" $fixtureRoot
        Set-ConfigValue $config "config.view$view.file_list.init_folder" '1'
        Set-ConfigValue $config "config.view$view.file_list.init_folder_path" $fixtureRoot
        Set-ConfigValue $config "config.view$view.file_list.row_focus_color" '255,255,255'
    }

    $oldCompat = $env:__COMPAT_LAYER
    try {
        $env:__COMPAT_LAYER = 'RunAsInvoker'
        $stageConfigDir = Join-Path $stageRoot 'fxfile'
        $arguments = "-w 2x3 --conf_dir=`"$stageConfigDir`""
        for ($view = 1; $view -le 6; ++$view) {
            $arguments += " --dir$view=`"$fixtureRoot`""
        }
        $process = Start-Process -FilePath $stageExe -WorkingDirectory $stageRoot -ArgumentList $arguments -PassThru
    }
    finally { $env:__COMPAT_LAYER = $oldCompat }

    $deadline = [DateTime]::UtcNow.AddSeconds($ReadyTimeoutSeconds)
    do {
        Start-Sleep -Milliseconds 250
        if ($process.HasExited) { throw "FxFile exited before ready. ExitCode=$($process.ExitCode)" }
        $mainWindow = Get-MainWindow $process.Id
        [object[]]$lists = if ($mainWindow -ne [IntPtr]::Zero) { Get-VisibleListViews $mainWindow } else { @() }
        $process.Refresh()
        if ($mainWindow -ne [IntPtr]::Zero -and $process.Responding -and @($lists).Count -eq 6) { break }
    } while ([DateTime]::UtcNow -lt $deadline)
    if (@($lists).Count -ne 6 -or -not $process.Responding) { throw 'A responsive 2x3 layout did not become ready.' }
    [void][FxTask111.NativeMethods]::SetForegroundWindow($mainWindow)
    Start-Sleep -Seconds 8

    $expectedItemCount = 1 + $fixtureFolderNames.Count + $fixtureFileNames.Count
    foreach ($list in $lists) {
        $actualItemCount = Get-ListViewItemCount $list.Handle
        if ($actualItemCount -ne $expectedItemCount) {
            throw "Mixed fixture row count mismatch: expected=$expectedItemCount actual=$actualItemCount"
        }
    }

    $transition = 0
    $samples.Add((Get-Sample $process $mainWindow $transition))
    $workDeadline = [DateTime]::UtcNow.AddSeconds($DurationSeconds)
    while ([DateTime]::UtcNow -lt $workDeadline) {
        [object[]]$lists = Get-StableVisibleListViews $mainWindow
        if (@($lists).Count -ne 6) { throw "Visible pane count changed to $(@($lists).Count)." }
        $paneIndex = $transition % 6
        $target = $lists[$paneIndex].Handle
        $paneCycle = [int][Math]::Floor($transition / 6)
        if (($paneCycle % 2) -eq 0) {
            # Parent is index 0; FxFile's name sort keeps folders before files.
            $targetItem = 1 + (($paneIndex + [int][Math]::Floor($paneCycle / 2)) % $fixtureFolderNames.Count)
            $targetKind = 'folder'
        }
        else {
            $targetItem = 1 + $fixtureFolderNames.Count + (($paneIndex + [int][Math]::Floor($paneCycle / 2)) % $fixtureFileNames.Count)
            $targetKind = 'file'
        }
        if (-not (Select-ListViewItem $target $targetItem)) {
            $selectedItem = Get-FirstSelectedIndex $target
            throw "A $targetKind row selection was not confirmed in pane $($paneIndex + 1), expected item $targetItem, selected item $selectedItem."
        }
        if ($targetKind -eq 'folder') { ++$folderSelectionsByPane[$paneIndex] }
        else { ++$fileSelectionsByPane[$paneIndex] }

        # Refresh each pane in turn.  This repeatedly invalidates queued shell
        # icons and exercises the generation/cancellation path that previously
        # produced late partial-row repaints after prolonged use.
        if ($RefreshEveryTransitions -gt 0 -and
            ($transition % $RefreshEveryTransitions) -eq ($RefreshEveryTransitions - 1)) {
            [void][FxTask111.NativeMethods]::PostMessage($target, 0x0100, [IntPtr]0x74, [IntPtr]::Zero)
            [void][FxTask111.NativeMethods]::PostMessage($target, 0x0101, [IntPtr]0x74, [IntPtr]::Zero)
        }

        ++$transition
        Start-Sleep -Milliseconds 250
        if (($transition % 10) -eq 0) {
            $sample = Get-Sample $process $mainWindow $transition
            $samples.Add($sample)
            if (-not $sample.Responding) { throw "FxFile became Not Responding at transition $transition." }
            if ($sample.VisibleFileLists -ne 6) { throw "A pane disappeared at transition $transition." }
        }
    }

    Start-Sleep -Seconds 12
    $samples.Add((Get-Sample $process $mainWindow $transition))
    $start = $samples[0]
    $end = $samples[$samples.Count - 1]
    $privateDelta = $end.PrivateMemoryBytes - $start.PrivateMemoryBytes
    $workingDelta = $end.WorkingSetBytes - $start.WorkingSetBytes
    $handleDelta = $end.HandleCount - $start.HandleCount
    $gdiDelta = $end.GdiHandles - $start.GdiHandles
    $userDelta = $end.UserHandles - $start.UserHandles
    if (@($folderSelectionsByPane | Where-Object { $_ -eq 0 }).Count -ne 0 -or
        @($fileSelectionsByPane | Where-Object { $_ -eq 0 }).Count -ne 0) {
        throw 'Every pane must confirm at least one real folder-row and file-row selection.'
    }
    if ($privateDelta -gt 128MB -or $workingDelta -gt 160MB) {
        throw "Memory accumulation exceeded bound: private=$privateDelta working=$workingDelta"
    }
    # Same-folder refresh must now be resource-bounded.  The former PathBar
    # defect grew exactly +3 GDI/+1 USER per refresh; these limits fail that
    # staircase long before Windows' per-process quotas can affect painting.
    if ($handleDelta -gt 48 -or $gdiDelta -gt 12 -or $userDelta -gt 6) {
        throw "Handle accumulation exceeded bound: handles=$handleDelta GDI=$gdiDelta USER=$userDelta"
    }

    if (-not [FxTask111.NativeMethods]::PostMessage($mainWindow, 0x0111, [IntPtr]30120, [IntPtr]::Zero)) {
        throw 'Normal exit command could not be posted.'
    }
    if (-not $process.WaitForExit(20000)) { throw 'FxFile did not close normally within 20 seconds.' }
    if ($process.ExitCode -ne 0) { throw "FxFile returned exit code $($process.ExitCode)." }
    $result = 'PASS'
}
catch {
    $failure = $_.Exception.Message
    throw
}
finally {
    if ($null -ne $process -and -not $process.HasExited) {
        try { [void][FxTask111.NativeMethods]::PostMessage($mainWindow, 0x0111, [IntPtr]30120, [IntPtr]::Zero) } catch { }
        try { if (-not $process.WaitForExit(5000)) { Stop-Process -Id $process.Id -Force } } catch { }
    }
    $report = [ordered]@{
        Result = $result
        CapturedAt = (Get-Date).ToString('o')
        Failure = $failure
        DurationSeconds = $DurationSeconds
        RefreshEveryTransitions = $RefreshEveryTransitions
        ParentFolderVisible = -not $HideParentFolder.IsPresent
        StatusBarVisible = -not $HideStatusBar.IsPresent
        FixtureFolderCount = $fixtureFolderNames.Count
        FixtureFileCount = $fixtureFileNames.Count
        FolderRowSelectionsByPane = $folderSelectionsByPane
        FileRowSelectionsByPane = $fileSelectionsByPane
        TransitionCount = if ($null -ne (Get-Variable transition -ErrorAction SilentlyContinue)) { $transition } else { 0 }
        ExecutableSha256 = (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash
        Samples = $samples
    }
    $report | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    Remove-IsolatedRoot $stageRoot
    Remove-IsolatedRoot $fixtureRoot
}

Get-Content -LiteralPath $reportPath -Raw
