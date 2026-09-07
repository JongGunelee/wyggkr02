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
$selectedFocus = Function-Body $explorer 'xpr_bool_t ExplorerCtrl::isFocusedSelectedItem' 'void ExplorerCtrl::resetCustomDrawColors'
$focusState = Function-Body $explorer 'void ExplorerCtrl::applyRowFocusDrawState' 'void ExplorerCtrl::applyReportSelectionDrawState'
$reportSelection = Function-Body $explorer 'void ExplorerCtrl::applyReportSelectionDrawState' 'void ExplorerCtrl::drawFinalReportSelection'
$finalPaint = Function-Body $explorer 'void ExplorerCtrl::drawFinalReportSelection' 'void ExplorerCtrl::drawParentFolderIcon'
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$subItemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM))' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)'
$itemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT)'
$itemChanged = Function-Body $explorer 'void ExplorerCtrl::OnItemchanged' 'xpr_size_t ExplorerCtrl::getRealSelCount'

Check 'The report item notification applies configured colors to every live-selected row' (
    $itemDraw.Contains('XPR_IS_TRUE(isReportView())') -and
    $itemDraw.Contains('applyReportSelectionDrawState(sNmLvCustomDraw);') -and
    $reportSelection.Contains('if (!XPR_TEST_BITS(sNativeState, LVIS_SELECTED))') -and
    -not $reportSelection.Contains('isFocusedSelectedItem(sItemIndex)') -and
    $itemDraw.Contains('CDRF_NEWFONT | CDRF_NOTIFYSUBITEMDRAW'))
Check 'The final item paint owns every selected row scope after Windows selection compositing' (
    $subItemDraw.Contains('XPR_IS_TRUE(mOption.mFullRowSelect) || sNmLvCustomDraw->iSubItem == 0') -and
    $subItemDraw.Contains('applyReportSelectionDrawState(sNmLvCustomDraw);') -and
    -not $reportSelection.Contains('::FillRect(') -and
    $finalPaint.Contains('LVIR_BOUNDS : LVIR_SELECTBOUNDS') -and
    $finalPaint.Contains('::FillRect(') -and
    $focusState.Contains('clrTextBk = mOption.mRowFocusColor;') -and
    $focusState.Contains('clrText   = mRowFocusTextColor;'))
Check 'Paint-state suppression never mutates the list selection model or schedules work' (
    -not $customDraw.Contains('SetItemState(') -and
    -not $customDraw.Contains('SetItem(') -and
    -not $customDraw.Contains('SetTimer(') -and
    -not $customDraw.Contains('PostMessage(') -and
    -not $finalPaint.Contains('SetItemState(') -and
    -not $finalPaint.Contains('Invalidate(') -and
    $focusState.Contains('uItemState &= ~CDIS_SELECTED') -and
    -not [regex]::IsMatch($focusState, '(?m)^\s*aNmLvCustomDraw->iStateId\s*=') -and
    -not [regex]::IsMatch($focusState, '(?m)^\s*aNmLvCustomDraw->clrFace\s*=') -and
    $selectedFocus.Contains('mRowFocusPaintItemIndex == aItem'))
Check 'Full-row focus does not schedule a RedrawItems loop on every selection transition' (
    -not $itemChanged.Contains('mOption.mFullRowSelect'))
Check 'Repeated owner-data selection notifications do not re-arm a one-shot row paint gate' (
    $itemChanged.Contains('if (mFocusedItemIndex != sNmListView->iItem)') -and
    -not $explorer.Contains('mRowFocusPaintPending'))
Check 'Every pane still receives its own configured colour and full-row versus first-cell scope remains explicit' (
    $pane.Contains('aOption.mConfig.mFileListRowFocusColor[mViewIndex]') -and
    $customDraw.Contains('applyReportSelectionDrawState(sNmLvCustomDraw);'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
