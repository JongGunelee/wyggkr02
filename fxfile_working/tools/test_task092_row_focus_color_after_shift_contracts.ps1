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
$mainFrame = Read-Text 'src\fxfile\main_frame.cpp'
$option = Read-Text 'src\fxfile\option.cpp'
$colorDialog = Read-Text 'src\fxfile\cfg\cfg_appearance_color_dlg.cpp'
$visualProfile = Read-Text 'tools\Prepare-Task092RowFocusVisualProfile.ps1'

$snapshot = Function-Body $explorer 'void ExplorerCtrl::snapshotRowFocusItem' 'xpr_bool_t ExplorerCtrl::isFocusedSelectedItem'
$click = Function-Body $explorer 'void ExplorerCtrl::OnClick' 'void ExplorerCtrl::OnLButtonDblClk'
$leftDown = Function-Body $explorer 'void ExplorerCtrl::OnLButtonDown' 'void ExplorerCtrl::OnMouseMove'
$leftUp = Function-Body $explorer 'void ExplorerCtrl::OnLButtonUp' 'void ExplorerCtrl::OnRButtonDown'
$keyUp = Function-Body $explorer 'void ExplorerCtrl::OnKeyUp' 'void ExplorerCtrl::OnMarqueebegin'
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$drawState = Function-Body $explorer 'void ExplorerCtrl::applyRowFocusDrawState' 'void ExplorerCtrl::OnCustomdraw('
$reportSelection = Function-Body $explorer 'void ExplorerCtrl::applyReportSelectionDrawState' 'void ExplorerCtrl::drawFinalReportSelection'
$finalPaint = Function-Body $explorer 'void ExplorerCtrl::drawFinalReportSelection' 'void ExplorerCtrl::drawParentFolderIcon'
$setOption = Function-Body $explorer 'void ExplorerCtrl::setOption' 'void ExplorerCtrl::setImageList'
$paneSetViewIndex = Function-Body $pane 'void ExplorerPane::setViewIndex' 'void ExplorerPane::setExplorerObserver'
$splitView = Function-Body $mainFrame 'void MainFrame::splitView' 'xpr_sint_t MainFrame::getViewCount'

$focusedPos = $snapshot.IndexOf('GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED)', [StringComparison]::Ordinal)
$markPos = $snapshot.IndexOf('GetSelectionMark()', [StringComparison]::Ordinal)
$keySuperPos = $keyUp.IndexOf('super::OnKeyUp(aChar, aRepCnt, aFlags);', [StringComparison]::Ordinal)
$keyFocusedPos = $keyUp.IndexOf('GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED)', [StringComparison]::Ordinal)

Check 'The settings dialog saves the selected row-focus color for every view slot' (
    $colorDialog.Contains('aViewColor.mFileListRowFocusColor    = mFileListRowFocusColorCtrl.GetColor();') -and
    $colorDialog.Contains('aConfig.mFileListRowFocusColor[i]    = sViewColor->mFileListRowFocusColor;') -and
    ([regex]::Matches($option, 'config\.view[1-6]\.file_list\.row_focus_color').Count -eq 6))
Check 'Every Explorer pane receives its own saved row-focus color' (
    $pane.Contains('aOption.mConfig.mFileListRowFocusColor[mViewIndex]'))
Check 'Apply and OK update row-focus color and mode without waiting for folder navigation' (
    $setOption.Contains('cacheRowFocusOption(aOption);') -and
    $setOption.Contains('LVS_EX_FULLROWSELECT') -and
    $setOption.Contains('Invalidate(XPR_FALSE);'))
Check 'Reused panes rebind existing controls to their canonical per-view options without refresh' (
    $paneSetViewIndex.Contains('sExplorerCtrl->setViewIndex(aViewIndex);') -and
    $paneSetViewIndex.Contains('setExplorerOption(sExplorerCtrl, *gOpt);') -and
    -not $paneSetViewIndex.Contains('refresh('))
Check 'Every split layout canonicalizes pane indices immediately after Splitter reuse' (
    $splitView.Contains('mSplitter.split(aRowCount, aColumnCount);') -and
    $splitView.Contains('getViewIndexFromViewSplit(aRowCount,') -and
    $splitView.Contains('sExplorerView->setViewIndex(sViewIndex);'))
Check 'Visual profile replacements preserve CRLF and force six deterministic selected rows' (
    $visualProfile.Contains("'\s*=[^\r\n]*'") -and
    $visualProfile.Contains('Generated visual-test profile contains an LF-only line') -and
    $visualProfile.Contains('$launchArguments.Add("--dir$view")'))
Check 'Paint targets the native focused selected row before the Shift anchor fallback' (
    $focusedPos -ge 0 -and $markPos -ge 0 -and $focusedPos -lt $markPos -and
    $snapshot.Contains('GetItemState(') -and $snapshot.Contains('LVIS_SELECTED'))
Check 'Mouse repaint is requested only after the native click notification identifies the final row' (
    $click.Contains('mFocusedItemIndex = sNmItemActivate->iItem;') -and
    $click.Contains('redrawFocusItemChange(sOldFocused, sNmItemActivate->iItem);') -and
    -not $click.Contains('Invalidate(XPR_FALSE);') -and
    -not $leftDown.Contains('mFocusedItemIndex =') -and
    -not $leftDown.Contains('Invalidate('))
Check 'Keyboard repaint reads the final focused selected row after native navigation' (
    $keySuperPos -ge 0 -and $keyFocusedPos -gt $keySuperPos -and
    $keyUp.Contains('redrawFocusItemChange(sOldFocused, sItemIndex);') -and
    -not $keyUp.Contains('Invalidate(XPR_FALSE);'))
Check 'Report paint fills the focused row with the exact configured color' (
    $drawState.Contains('clrTextBk = mOption.mRowFocusColor;') -and
    $drawState.Contains('clrText   = mRowFocusTextColor;') -and
    $reportSelection.Contains('applyRowFocusDrawState(aNmLvCustomDraw);') -and
    -not $reportSelection.Contains('::FillRect(') -and
    $finalPaint.Contains('sBackgroundColor = mOption.mRowFocusColor;') -and
    $finalPaint.Contains('::FillRect('))
Check 'Content and tile layouts mapped to native report view use the same row-focus paint path' (
    ([regex]::Matches($customDraw, 'XPR_IS_TRUE\(isReportView\(\)\)').Count -ge 2) -and
    -not $customDraw.Contains('getViewStyle() == VIEW_STYLE_DETAILS'))
Check 'Themed ListView paint is not forced back to its normal white background' (
    $drawState.Contains('aNmLvCustomDraw->nmcd.uItemState &= ~CDIS_SELECTED;') -and
    -not [regex]::IsMatch($drawState, '(?m)^\s*aNmLvCustomDraw->iStateId\s*=') -and
    -not [regex]::IsMatch($drawState, '(?m)^\s*aNmLvCustomDraw->clrFace\s*='))
Check 'Row-focus paint remains allocation-free and never mutates native selection' (
    -not $snapshot.Contains('new ') -and
    -not $customDraw.Contains('SetItemState(') -and
    -not $customDraw.Contains('SetSelectionMark(') -and
    -not $customDraw.Contains('PostMessage(') -and
    -not $customDraw.Contains('SetTimer(') -and
    -not $finalPaint.Contains('SetItemState(') -and
    -not $finalPaint.Contains('SetSelectionMark('))
Check 'Shift range anchors remain owned by the native ListView mouse handlers' (
    $leftDown.Contains('super::OnLButtonDown(aFlags, aPoint);') -and
    $leftUp.Contains('super::OnLButtonUp(aFlags, aPoint);') -and
    -not $leftDown.Contains('SetSelectionMark(') -and
    -not $leftUp.Contains('SetSelectionMark('))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
