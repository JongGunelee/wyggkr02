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

$definition = Read-Text 'src\fxfile\fxfile_def.h'
$option = Read-Text 'src\fxfile\option.cpp'
$pane = Read-Text 'src\fxfile\explorer_pane.cpp'
$explorer = Read-Text 'src\fxfile\explorer_ctrl.cpp'
$header = Read-Text 'src\fxfile\explorer_ctrl.h'
$view = Read-Text 'src\fxfile\explorer_view.cpp'
$longRun = Read-Text 'tools\Test-Task111SixPaneLongRun.ps1'

$focusState = Function-Body $explorer 'void ExplorerCtrl::applyRowFocusDrawState' 'void ExplorerCtrl::applyReportSelectionDrawState'
$reportState = Function-Body $explorer 'void ExplorerCtrl::applyReportSelectionDrawState' 'void ExplorerCtrl::drawFinalReportSelection'
$fill = Function-Body $explorer 'void ExplorerCtrl::drawFinalReportSelection' 'void ExplorerCtrl::drawParentFolderIcon'
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$subItemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM))' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)'
$itemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT)'

Check 'All six panes still own independent persisted row-focus colours' (
    $definition.Contains('#define MAX_VIEW_SPLIT_ROW        (2)') -and
    $definition.Contains('#define MAX_VIEW_SPLIT_COLUMN     (3)') -and
    ([regex]::Matches($option, 'config\.view[1-6]\.file_list\.row_focus_color').Count -eq 6) -and
    $pane.Contains('aOption.mConfig.mFileListRowFocusColor[mViewIndex]') -and
    $view.Contains('new ExplorerCtrl'))

Check 'Focused selection resolves the configured background and contrast text without mutating the model' (
    $focusState.Contains('clrTextBk = mOption.mRowFocusColor;') -and
    $focusState.Contains('clrText   = mRowFocusTextColor;') -and
    $focusState.Contains('uItemState &= ~CDIS_SELECTED;') -and
    -not $focusState.Contains('SetItemState(') -and
    -not $focusState.Contains('SetTextColor(') -and
    -not $focusState.Contains('SetBkColor('))

Check 'Every Ctrl and Shift selection resolves the configured readable colours' (
    $reportState.Contains('GetItemState(sItemIndex, LVIS_SELECTED | LVIS_FOCUSED)') -and
    $reportState.Contains('applyRowFocusDrawState(aNmLvCustomDraw);') -and
    -not $reportState.Contains('COLOR_HIGHLIGHT') -and
    -not $reportState.Contains('COLOR_BTNFACE') -and
    $fill.Contains('LVIS_DROPHILITED') -and
    $fill.Contains('COLOR_HIGHLIGHT') -and
    -not $fill.Contains('COLOR_BTNFACE') -and
    -not $reportState.Contains('SetItemState('))

Check 'A selected report item receives one explicit final theme-proof background fill' (
    $header.Contains('drawFinalReportSelection') -and
    $fill.Contains('GetItemState(sItemIndex') -and
    $fill.Contains('LVIR_BOUNDS : LVIR_SELECTBOUNDS') -and
    $fill.Contains('GetClientRect(&sClientRect)') -and
    $fill.Contains('IntersectRect') -and
    $fill.Contains('sBackgroundColor = mOption.mRowFocusColor') -and
    $fill.Contains('::GetStockObject(DC_BRUSH)') -and
    $fill.Contains('::FillRect('))

Check 'The explicit final fill is allocation-free and restores the complete DC state' (
    $fill.Contains('::SetDCBrushColor(') -and
    $fill.Contains('::SaveDC(') -and
    $fill.Contains('::RestoreDC(') -and
    -not $fill.Contains('CreateSolidBrush') -and
    -not $fill.Contains('DeleteObject') -and
    -not $fill.Contains('new ') -and
    -not $fill.Contains('SetItemState(') -and
    -not $fill.Contains('SetSelectionMark('))

Check 'The fill occurs only in supported report ITEMPOSTPAINT after native theme rendering' (
    $itemDraw.Contains('applyReportSelectionDrawState(sNmLvCustomDraw);') -and
    $itemDraw.Contains('CDRF_NOTIFYPOSTPAINT') -and
    $customDraw.Contains('drawFinalReportSelection(sNmLvCustomDraw);') -and
    -not $itemDraw.Contains('::FillRect(') -and
    -not $subItemDraw.Contains('drawFinalReportSelection') -and
    -not $customDraw.Contains('CDDS_ITEMPREERASE'))

Check 'Native prepaint remains and final report rendering restores icons and text' (
    $itemDraw.Contains('CDRF_NEWFONT | CDRF_NOTIFYSUBITEMDRAW') -and
    $subItemDraw.Contains('applyReportSelectionDrawState(sNmLvCustomDraw);') -and
    $subItemDraw.Contains('CDRF_NEWFONT') -and
    -not $subItemDraw.Contains('CDRF_SKIPDEFAULT') -and
    $fill.Contains('mSmallImgList->Draw(') -and
    $fill.Contains('::DrawText(') -and
    ([regex]::Matches($customDraw, 'CDRF_SKIPDEFAULT').Count -eq 1))

Check 'The six-pane runtime harness covers real folders and files, not folders alone' (
    $longRun.Contains('mixed_file_and_folder_fixture') -and
    $longRun.Contains('FolderRowSelectionsByPane') -and
    $longRun.Contains('FileRowSelectionsByPane') -and
    $longRun.Contains('Every pane must confirm at least one real folder-row and file-row selection.'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object {
    '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name
}
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
