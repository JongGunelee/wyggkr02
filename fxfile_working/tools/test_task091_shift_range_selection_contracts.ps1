[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot

function Read-Text([string]$relativePath) {
    [IO.File]::ReadAllText((Join-Path $root $relativePath))
}

function Function-Body([string]$text, [string]$signature, [string]$nextSignature) {
    $start = $text.IndexOf($signature, [StringComparison]::Ordinal)
    if ($start -lt 0) { return '' }
    $end = $text.IndexOf($nextSignature, $start + $signature.Length, [StringComparison]::Ordinal)
    if ($end -lt 0) { $end = $text.Length }
    return $text.Substring($start, $end - $start)
}

$checks = [Collections.Generic.List[object]]::new()
function Check([string]$name, [bool]$passed) {
    $checks.Add([pscustomobject]@{ Name = $name; Passed = $passed })
}

$explorer = Read-Text 'src\fxfile\explorer_ctrl.cpp'
$pane = Read-Text 'src\fxfile\explorer_pane.cpp'
$option = Read-Text 'src\fxfile\option.cpp'
$probe = Read-Text 'tools\fxfile_listview_state_probe.cpp'

$leftDown = Function-Body $explorer 'void ExplorerCtrl::OnLButtonDown' 'void ExplorerCtrl::OnMouseMove'
$leftUp = Function-Body $explorer 'void ExplorerCtrl::OnLButtonUp' 'void ExplorerCtrl::OnRButtonDown'
$click = Function-Body $explorer 'void ExplorerCtrl::OnClick' 'void ExplorerCtrl::OnLButtonDblClk'
$preCreate = Function-Body $explorer 'xpr_bool_t ExplorerCtrl::PreCreateWindow' 'void ExplorerCtrl::OnDestroy'
$snapshot = Function-Body $explorer 'void ExplorerCtrl::snapshotRowFocusItem' 'xpr_bool_t ExplorerCtrl::isFocusedSelectedItem'

Check 'Mouse down delegates the selection transition and Shift anchor to the native ListView' (
    $leftDown.Contains('super::OnLButtonDown(aFlags, aPoint);') -and
    -not $leftDown.Contains('SetSelectionMark(') -and
    -not $leftDown.Contains('SetItemState('))
Check 'Mouse up does not overwrite the native Shift range anchor after selection completes' (
    $leftUp.Contains('super::OnLButtonUp(aFlags, aPoint);') -and
    -not $leftUp.Contains('SetSelectionMark(') -and
    -not $leftUp.Contains('SetItemState('))
Check 'Visual row identity remains cached without mutating native Ctrl or Shift selection state' (
    -not $leftDown.Contains('mFocusedItemIndex =') -and
    $click.Contains('mFocusedItemIndex = sNmItemActivate->iItem;') -and
    $leftUp.Contains('mFocusedItemIndex = sItemIndex;'))
Check 'Explorer lists remain native multi-select controls' (
    $preCreate.Contains('LVS_REPORT') -and
    -not $preCreate.Contains('LVS_SINGLESEL'))
Check 'The row-focus renderer remains a read-only consumer of the native selection identity' (
    $snapshot.Contains('GetSelectionMark()') -and
    $snapshot.Contains('LVNI_FOCUSED | LVNI_SELECTED') -and
    -not $snapshot.Contains('SetSelectionMark(') -and
    -not $snapshot.Contains('SetItemState('))
Check 'Every layout pane uses the same ExplorerCtrl selection implementation' (
    $pane.Contains('ExplorerCtrl') -and
    $pane.Contains('mExplorerCtrl'))
Check 'All six pane indices retain independent row-focus color settings' (
    $pane.Contains('aOption.mConfig.mFileListRowFocusColor[mViewIndex]') -and
    ([regex]::Matches($option, 'config\.view[1-6]\.file_list\.row_focus_color').Count -eq 6))
Check 'Runtime probe records the complete selected range for each visible pane' (
    $probe.Contains('selectedItems') -and
    $probe.Contains('LVM_GETNEXTITEM') -and
    $probe.Contains('LVNI_SELECTED') -and
    $probe.Contains('selected_items='))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
