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
    $text.Substring($start, $end - $start)
}

$checks = [Collections.Generic.List[object]]::new()
function Check([string]$name, [bool]$passed) {
    $checks.Add([pscustomobject]@{ Name = $name; Passed = $passed })
}

$explorer = Read-Text 'src\fxfile\explorer_ctrl.cpp'
$pane = Read-Text 'src\fxfile\explorer_pane.cpp'
$option = Read-Text 'src\fxfile\option.cpp'
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$reportState = Function-Body $explorer 'void ExplorerCtrl::applyReportSelectionDrawState' 'void ExplorerCtrl::drawFinalReportSelection'
$finalPaint = Function-Body $explorer 'void ExplorerCtrl::drawFinalReportSelection' 'void ExplorerCtrl::drawParentFolderIcon'
$leftDown = Function-Body $explorer 'void ExplorerCtrl::OnLButtonDown' 'void ExplorerCtrl::OnMouseMove'
$leftUp = Function-Body $explorer 'void ExplorerCtrl::OnLButtonUp' 'void ExplorerCtrl::OnRButtonDown'

Check 'All six panes still provide their independent configured selection-row colour' (
    $pane.Contains('aOption.mConfig.mFileListRowFocusColor[mViewIndex]') -and
    ([regex]::Matches($option, 'config\.view[1-6]\.file_list\.row_focus_color').Count -eq 6))

Check 'Every live-selected report row receives the configured background and contrast text' (
    $reportState.Contains('GetItemState(sItemIndex, LVIS_SELECTED') -and
    $reportState.Contains('applyRowFocusDrawState(aNmLvCustomDraw);') -and
    -not $reportState.Contains('isFocusedSelectedItem(sItemIndex)') -and
    -not $reportState.Contains('COLOR_HIGHLIGHT') -and
    -not $reportState.Contains('COLOR_BTNFACE'))

Check 'Final paint applies the configured colour to range start middle and end alike' (
    $finalPaint.Contains('if (!XPR_TEST_BITS(sNativeState, LVIS_SELECTED))') -and
    $finalPaint.Contains('sBackgroundColor = mOption.mRowFocusColor;') -and
    $finalPaint.Contains('sTextColor = mRowFocusTextColor;') -and
    -not $finalPaint.Contains('isFocusedSelectedItem(sItemIndex)') -and
    -not $finalPaint.Contains('COLOR_BTNFACE'))

Check 'Drop highlighting remains an explicit transient override of selection colour' (
    $finalPaint.Contains('LVIS_DROPHILITED') -and
    $finalPaint.Contains('COLOR_HIGHLIGHT') -and
    $finalPaint.Contains('COLOR_HIGHLIGHTTEXT'))

Check 'Only the actual keyboard-focus row receives the focus rectangle' (
    $finalPaint.Contains('LVIS_FOCUSED') -and
    $finalPaint.Contains('::GetFocus() == m_hWnd') -and
    $finalPaint.Contains('::DrawFocusRect('))

Check 'Report item postpaint is requested for every live-selected row' (
    $customDraw.Contains('GetItemState(sItemIndex, LVIS_SELECTED) == LVIS_SELECTED') -and
    $customDraw.Contains('CDRF_NOTIFYPOSTPAINT') -and
    $customDraw.Contains('drawFinalReportSelection(sNmLvCustomDraw);'))

Check 'Native single Ctrl and Shift selection ownership remains untouched' (
    $leftDown.Contains('super::OnLButtonDown(aFlags, aPoint);') -and
    $leftUp.Contains('super::OnLButtonUp(aFlags, aPoint);') -and
    -not $leftDown.Contains('SetItemState(') -and
    -not $leftUp.Contains('SetItemState(') -and
    -not $leftDown.Contains('SetSelectionMark(') -and
    -not $leftUp.Contains('SetSelectionMark(') -and
    -not $finalPaint.Contains('SetItemState(') -and
    -not $finalPaint.Contains('SetSelectionMark('))

Check 'Selected-row paint stays allocation-free and cannot schedule a repaint loop' (
    $finalPaint.Contains('::SaveDC(') -and
    $finalPaint.Contains('::RestoreDC(') -and
    -not $finalPaint.Contains('new ') -and
    -not $finalPaint.Contains('Invalidate(') -and
    -not $finalPaint.Contains('RedrawItems(') -and
    -not $finalPaint.Contains('SetTimer(') -and
    -not $finalPaint.Contains('PostMessage('))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object {
    '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name
}
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
