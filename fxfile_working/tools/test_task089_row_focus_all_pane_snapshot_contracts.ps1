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
$header = Read-Text 'src\fxfile\explorer_ctrl.h'
$pane = Read-Text 'src\fxfile\explorer_pane.cpp'
$option = Read-Text 'src\fxfile\option.cpp'
$probe = Read-Text 'tools\fxfile_listview_state_probe.cpp'

$snapshot = Function-Body $explorer 'void ExplorerCtrl::snapshotRowFocusItem' 'xpr_bool_t ExplorerCtrl::isFocusedSelectedItem'
$selectedFocus = Function-Body $explorer 'xpr_bool_t ExplorerCtrl::isFocusedSelectedItem' 'void ExplorerCtrl::resetCustomDrawColors'
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$prePaint = Function-Body $customDraw 'if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_PREPAINT)' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREERASE)'
$preEnumeration = Function-Body $explorer 'void ExplorerCtrl::preEnumeration' 'xpr_bool_t ExplorerCtrl::insertPidlItem'

Check 'Every native paint snapshots this ListView selection before item callbacks' (
    $header.Contains('snapshotRowFocusItem') -and
    $header.Contains('mRowFocusPaintItemIndex') -and
    $prePaint.Contains('snapshotRowFocusItem();') -and
    $prePaint.IndexOf('snapshotRowFocusItem();', [StringComparison]::Ordinal) -lt
        $prePaint.IndexOf('CDRF_NOTIFYITEMDRAW', [StringComparison]::Ordinal))
Check 'The snapshot resolves focused-selected before selection-mark and first-selected fallbacks' (
    $snapshot.Contains('GetSelectionMark()') -and
    $snapshot.Contains('GetItemState(') -and
    $snapshot.Contains('LVIS_SELECTED') -and
    $snapshot.Contains('LVNI_FOCUSED | LVNI_SELECTED') -and
    $snapshot.Contains('GetNextItem(-1, LVNI_SELECTED)') -and
    $snapshot.IndexOf('LVNI_FOCUSED | LVNI_SELECTED', [StringComparison]::Ordinal) -lt
        $snapshot.IndexOf('GetSelectionMark()', [StringComparison]::Ordinal))
Check 'Item and subitem draw use the immutable paint snapshot instead of active-window focus' (
    $selectedFocus.Contains('mRowFocusPaintItemIndex == aItem') -and
    -not $selectedFocus.Contains('mFocusedItemIndex == aItem') -and
    -not $selectedFocus.Contains('::GetFocus()'))
Check 'Folder re-enumeration cannot retain a row identity from the previous folder' (
    $preEnumeration -match 'mFocusedItemIndex\s*=\s*-1;' -and
    $preEnumeration -match 'mRowFocusPaintItemIndex\s*=\s*-1;')
Check 'The hot paint snapshot is constant-time, allocation-free, and read-only' (
    -not $snapshot.Contains('new ') -and
    -not $snapshot.Contains('SetItemState(') -and
    -not $snapshot.Contains('SetSelectionMark(') -and
    -not $snapshot.Contains('Invalidate(') -and
    -not $snapshot.Contains('RedrawItems(') -and
    -not $snapshot.Contains('PostMessage(') -and
    -not $snapshot.Contains('SetTimer('))
Check 'Custom draw still preserves the native selection model and schedules no repaint loop' (
    -not $customDraw.Contains('SetItemState(') -and
    -not $customDraw.Contains('SetSelectionMark(') -and
    -not $customDraw.Contains('Invalidate(') -and
    -not $customDraw.Contains('RedrawItems(') -and
    -not $customDraw.Contains('PostMessage(') -and
    -not $customDraw.Contains('SetTimer('))
Check 'All six view indices retain independent saved row-focus colors' (
    $pane.Contains('aOption.mConfig.mFileListRowFocusColor[mViewIndex]') -and
    ([regex]::Matches($option, 'config\.view[1-6]\.file_list\.row_focus_color').Count -eq 6))
Check 'Runtime characterization records every visible ListView native selection identity' (
    $probe.Contains('EnumChildWindows') -and
    $probe.Contains('WC_LISTVIEW') -and
    $probe.Contains('LVM_GETSELECTEDCOUNT') -and
    $probe.Contains('LVM_GETSELECTIONMARK') -and
    $probe.Contains('LVNI_FOCUSED') -and
    $probe.Contains('LVNI_SELECTED'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
