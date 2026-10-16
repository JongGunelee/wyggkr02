[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$sourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\src\fxfile'))
$explorer = Get-Content -LiteralPath (Join-Path $sourceRoot 'explorer_ctrl.cpp') -Raw
$explorerHeader = Get-Content -LiteralPath (Join-Path $sourceRoot 'explorer_ctrl.h') -Raw
$mainFrame = Get-Content -LiteralPath (Join-Path $sourceRoot 'main_frame.cpp') -Raw
$keyboardProbe = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Test-Task117KeyboardStartup.ps1') -Raw
$navigationCommit = [regex]::Match($explorer, '(?s)xpr_bool_t ExplorerCtrl::commitNavigationSelection.*?\n\}').Value
$parentLanding = [regex]::Match($explorer, '(?s)xpr_bool_t ExplorerCtrl::focusParentFolderRow.*?\n\}').Value

$checks = [ordered]@{
    'ExplorerCtrl exposes an explicit parent-row keyboard landing operation' =
        $explorerHeader.Contains('focusParentFolderRow(void)') -and
        $explorer.Contains('ExplorerCtrl::focusParentFolderRow(void)') -and
        $explorer.Contains('XPR_IS_FALSE(mOption.mParentFolder)')
    'parent-row landing clears stale selection and commits row zero' =
        $parentLanding.Contains('commitNavigationSelection(0)') -and
        $navigationCommit.Contains('SetItemState(-1, 0, LVIS_SELECTED | LVIS_FOCUSED)') -and
        $navigationCommit.Contains('SetItemState(aItemIndex, LVIS_SELECTED | LVIS_FOCUSED') -and
        $navigationCommit.Contains('SetSelectionMark(aItemIndex)') -and
        $navigationCommit.Contains('EnsureVisible(aItemIndex, XPR_FALSE)')
    'ExplorerView Tab requests direct pane movement instead of child control traversal' =
        $mainFrame.Contains('moveFocus(aCurWnd, GetAsyncKeyState(VK_SHIFT) < 0, XPR_TRUE)') -and
        -not $mainFrame.Contains('aCtrlKey = XPR_FALSE;')
    'startup and pane transfer both land on the visible parent row' =
        ([regex]::Matches($mainFrame, 'sExplorerCtrl->focusParentFolderRow\(\);').Count -ge 2)
    'runtime probe captures readiness before navigation test time' =
        $keyboardProbe.Contains('$startupReadyMilliseconds = $started.ElapsedMilliseconds') -and
        $keyboardProbe.Contains('StartupReadyMilliseconds = $startupReadyMilliseconds') -and
        $keyboardProbe.Contains('TotalTestMilliseconds = $started.ElapsedMilliseconds')
    'runtime probe requires exactly one Tab for each next file pane' =
        $keyboardProbe.Contains('for ($i = 1; $i -lt $ExpectedPaneCount; $i++)') -and
        $keyboardProbe.Contains('entered ''$focusClass'' instead of the next file pane') -and
        $keyboardProbe.Contains('FocusedRow = $focusRow') -and
        -not $keyboardProbe.Contains('$maxKeys = [Math]::Max')
    'runtime probe requires one Shift+Tab to reach previous parent row' =
        $keyboardProbe.Contains('[FxKeyboardNative]::ShiftTab()') -and
        $keyboardProbe.Contains('$reverseRow -eq 0') -and
        $keyboardProbe.Contains('directly reach the previous file pane parent row')
}

$failed = @($checks.GetEnumerator() | Where-Object { -not $_.Value })
foreach ($entry in $checks.GetEnumerator()) {
    '{0}: {1}' -f ($(if ($entry.Value) { 'PASS' } else { 'FAIL' })), $entry.Key
}
if ($failed.Count -gt 0) {
    throw ('Task118 contracts failed: ' + (($failed | ForEach-Object Key) -join '; '))
}
'Task118 contracts: PASS ({0}/{0})' -f $checks.Count
