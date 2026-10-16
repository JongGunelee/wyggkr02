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
$focusState = Function-Body $explorer 'void ExplorerCtrl::applyRowFocusDrawState' 'void ExplorerCtrl::applyReportSelectionDrawState'
$reportSelection = Function-Body $explorer 'void ExplorerCtrl::applyReportSelectionDrawState' 'void ExplorerCtrl::drawFinalReportSelection'
$finalPaint = Function-Body $explorer 'void ExplorerCtrl::drawFinalReportSelection' 'void ExplorerCtrl::drawParentFolderIcon'
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$subItemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM))' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)'
$itemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT)'
$reportItemDraw = Function-Body $itemDraw 'if (XPR_IS_TRUE(isReportView()) &&' 'else if (getViewStyle() != VIEW_STYLE_THUMBNAIL)'

Check 'The saved full-row option still owns native report-row geometry' (
    $explorer.Contains('XPR_SET_OR_CLR_BITS(sExStyle, LVS_EX_FULLROWSELECT, aNewOption.mFullRowSelect);'))
Check 'The focused item requests subitem and final rendering before filling its exact scope' (
    $itemDraw.Contains('applyReportSelectionDrawState(sNmLvCustomDraw);') -and
    -not $reportSelection.Contains('::FillRect(') -and
    $itemDraw.Contains('CDRF_NEWFONT | CDRF_NOTIFYSUBITEMDRAW') -and
    $itemDraw.Contains('CDRF_NOTIFYPOSTPAINT') -and
    $finalPaint.Contains('LVIR_BOUNDS : LVIR_SELECTBOUNDS') -and
    $finalPaint.Contains('::FillRect('))
Check 'Every focused report subitem suppresses themed selection before applying both colours' (
    $subItemDraw.Contains('applyReportSelectionDrawState(sNmLvCustomDraw);') -and
    $reportSelection.Contains('applyRowFocusDrawState(aNmLvCustomDraw);') -and
    $focusState.IndexOf('uItemState &= ~CDIS_SELECTED', [StringComparison]::Ordinal) -lt
        $focusState.IndexOf('clrTextBk = mOption.mRowFocusColor', [StringComparison]::Ordinal) -and
    -not [regex]::IsMatch($focusState, '(?m)^\s*aNmLvCustomDraw->iStateId\s*=') -and
    -not [regex]::IsMatch($focusState, '(?m)^\s*aNmLvCustomDraw->clrFace\s*=') -and
    $focusState.Contains('clrText   = mRowFocusTextColor;') -and
    $subItemDraw.Contains('*aResult = CDRF_NEWFONT;'))
Check 'Theme suppression never changes the ListView selection model or schedules another paint' (
    -not $customDraw.Contains('SetItemState(') -and
    -not $customDraw.Contains('SetItem(') -and
    -not $customDraw.Contains('Invalidate(') -and
    -not $customDraw.Contains('RedrawItems(') -and
    -not $customDraw.Contains('PostMessage(') -and
    -not $reportSelection.Contains('SetItemState(') -and
    -not $focusState.Contains('SetItemState(') -and
    -not $finalPaint.Contains('SetItemState(') -and
    -not $finalPaint.Contains('Invalidate('))
Check 'Final icon text and alignment rendering preserves the ListView report contract' (
    -not $subItemDraw.Contains('CDRF_SKIPDEFAULT') -and
    -not $reportItemDraw.Contains('CDRF_SKIPDEFAULT') -and
    $finalPaint.Contains('mSmallImgList->Draw(') -and
    $finalPaint.Contains('LVCF_FMT') -and
    $finalPaint.Contains('DT_END_ELLIPSIS'))
Check 'All six panes retain independent saved row-focus colours' (
    $pane.Contains('aOption.mConfig.mFileListRowFocusColor[mViewIndex]') -and
    ([regex]::Matches((Read-Text 'src\fxfile\option.cpp'), 'config\.view[1-6]\.file_list\.row_focus_color').Count -eq 6))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
