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

$definition = Read-Text 'src\fxfile\fxfile_def.h'
$option = Read-Text 'src\fxfile\option.cpp'
$main = Read-Text 'src\fxfile\main_frame.cpp'
$view = Read-Text 'src\fxfile\explorer_view.cpp'
$pane = Read-Text 'src\fxfile\explorer_pane.cpp'
$ctrl = Read-Text 'src\fxfile\explorer_ctrl.cpp'
$dialog = Read-Text 'src\fxfile\cfg\cfg_appearance_color_dlg.cpp'

$focused = Function-Body $ctrl 'xpr_bool_t ExplorerCtrl::isFocusedSelectedItem' 'void ExplorerCtrl::resetCustomDrawColors'
$drawState = Function-Body $ctrl 'void ExplorerCtrl::applyRowFocusDrawState' 'void ExplorerCtrl::applyReportSelectionDrawState'
$reportSelection = Function-Body $ctrl 'void ExplorerCtrl::applyReportSelectionDrawState' 'void ExplorerCtrl::drawParentFolderIcon'
$customDraw = Function-Body $ctrl 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$subItemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM))' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)'
$itemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT)'
$saveAll = Function-Body $main 'void MainFrame::saveAllOptions' 'void MainFrame::destroy'
$onClose = Function-Body $main 'void MainFrame::OnClose' 'void MainFrame::exitTrayApp'
$onEndSession = Function-Body $main 'void MainFrame::OnEndSession' 'xpr_bool_t MainFrame::confirmToClose'
$setPathLock = Function-Body $main 'void MainFrame::setViewPathLocked' 'xpr_bool_t MainFrame::isViewSplitLocked'
$setSplitLock = Function-Body $main 'void MainFrame::setViewSplitLocked' 'xpr_bool_t MainFrame::isClockLocked'
$saveTab = Function-Body $view 'XPR_INLINE void saveTabOption' 'void ExplorerView::saveOption'
$viewDestroy = Function-Body $view 'void ExplorerView::OnDestroy' 'void ExplorerView::setChangedOption'
$onSelEndOk = Function-Body $dialog 'LRESULT CfgAppearanceColorDlg::OnSelEndOK' 'void CfgAppearanceColorDlg::OnItemSize'

Check 'All six row-focus defaults share the canonical white constant' (
    $definition.Contains('#define DEF_FILE_LIST_ROW_FOCUS_COLOR (RGB(255,255,255))') -and
    ([regex]::Matches($option, 'mFileListRowFocusColor\[\d\].*DEF_FILE_LIST_ROW_FOCUS_COLOR').Count -eq 6))

Check 'Every pane receives only its own indexed row-focus colour' (
    ([regex]::Matches($option, 'config\.view[1-6]\.file_list\.row_focus_color').Count -eq 6) -and
    $pane.Contains('aOption.mConfig.mFileListRowFocusColor[mViewIndex]') -and
    -not $pane.Contains('mFileListRowFocusColor[0];'))

Check 'The colour dialog saves only the active view unless Apply All is explicitly clicked' (
    $onSelEndOk.Contains('saveViewColor();') -and
    $onSelEndOk.Contains('setModified();') -and
    -not $onSelEndOk.Contains('OnApplyAll') -and
    $dialog.Contains('IDC_CFG_COLOR_APPLY_ALL') -and
    $dialog.Contains('OnApplyAll'))

Check 'Exactly one live-selected row retains keyboard-focus identity' (
    $focused.Contains('GetItemState(aItem, LVIS_SELECTED | LVIS_FOCUSED)') -and
    $focused.Contains('if (!XPR_TEST_BITS(sNativeState, LVIS_SELECTED))') -and
    $focused.Contains('if (XPR_TEST_BITS(sNativeState, LVIS_FOCUSED))') -and
    $focused.Contains('mRowFocusPaintItemIndex == aItem') -and
    -not $itemDraw.Contains('CDIS_HOT') -and
    -not $itemDraw.Contains('uItemState &'))

Check 'Subitems apply configured colour to every live-selected row without promoting hover to selection' (
    $subItemDraw.Contains('applyReportSelectionDrawState(sNmLvCustomDraw);') -and
    $reportSelection.Contains('applyRowFocusDrawState(aNmLvCustomDraw);') -and
    $reportSelection.Contains('if (!XPR_TEST_BITS(sNativeState, LVIS_SELECTED))') -and
    -not $reportSelection.Contains('isFocusedSelectedItem(sItemIndex)') -and
    -not $subItemDraw.Contains('CDIS_HOT') -and
    -not $subItemDraw.Contains('uItemState &'))

Check 'Every item recomputes live selection without a cross-item drawing token' (
    $customDraw.Contains('snapshotRowFocusItem();') -and
    $itemDraw.Contains('applyReportSelectionDrawState(sNmLvCustomDraw);') -and
    -not $ctrl.Contains('mRowFocusDrawingItemIndex'))

Check 'Row-focus draw state never leaves text or background colours on the shared HDC' (
    $drawState.Contains('uItemState &= ~CDIS_SELECTED;') -and
    $drawState.Contains('clrTextBk = mOption.mRowFocusColor;') -and
    $drawState.Contains('clrText   = mRowFocusTextColor;') -and
    -not $drawState.Contains('SetTextColor(') -and
    -not $drawState.Contains('SetBkColor('))

Check 'Every report item and subitem resets its own normal/filter colours first' (
    $itemDraw.IndexOf('resetCustomDrawColors(sNmLvCustomDraw);', [StringComparison]::Ordinal) -ge 0 -and
    $subItemDraw.IndexOf('resetCustomDrawColors(sNmLvCustomDraw);', [StringComparison]::Ordinal) -ge 0 -and
    $subItemDraw.Contains('CDRF_NEWFONT'))

Check 'Normal close and Windows session close preserve locked snapshots' (
    $onClose.Contains('saveOption();') -and
    -not $onClose.Contains('saveAllOptions();') -and
    $onEndSession.Contains('saveOption();') -and
    -not $onEndSession.Contains('saveAllOptions();'))

Check 'Explicit Save All refreshes only currently enabled lock baselines' (
    $saveAll.Contains('XPR_IS_TRUE(gOpt->mMain.mWindowPlacementLocked)') -and
    $saveAll.Contains('XPR_IS_TRUE(gOpt->mMain.mViewSplitLocked)') -and
    $saveAll.Contains('XPR_IS_TRUE(gOpt->mMain.mViewPathLocked)') -and
    $saveAll.Contains('saveWindowPlacement();') -and
    $saveAll.Contains('mSplitter.getPaneCount(') -and
    $saveAll.Contains('sExplorerCtrl->getCurPath(sPath);'))

Check 'Enabling path and split locks captures the current state and persists it' (
    $setPathLock.Contains('if (XPR_IS_TRUE(aLocked))') -and
    $setPathLock.Contains('getCurPath(sPath)') -and
    $setPathLock.Contains('saveMainOption();') -and
    $setSplitLock.Contains('if (XPR_IS_TRUE(aLocked))') -and
    $setSplitLock.Contains('mSplitter.getPaneCount(') -and
    $setSplitLock.Contains('saveMainOption();'))

Check 'Remember-none and path-lock tab saves never replace the baseline with a work path' (
    $saveTab.Contains('mViewPathLocked') -and
    $saveTab.Contains('SAVE_FOLDER_LAYOUT_NONE') -and
    $saveTab.Contains('mLockedViewPath[sViewIndex]') -and
    $saveTab.Contains('mFileListInitFolder[sViewIndex]'))

$drain = $viewDestroy.IndexOf('mExplorerPane->destroySubPane();', [StringComparison]::Ordinal)
$tabDelete = $viewDestroy.IndexOf('DESTROY_DELETE(mTabCtrl);', [StringComparison]::Ordinal)
Check 'Explorer controls are withdrawn before tab callbacks during shutdown' (
    $drain -ge 0 -and $tabDelete -gt $drain)

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object {
    '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name
}
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
