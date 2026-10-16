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

$optionHeader = Read-Text 'src\fxfile\option.h'
$option = Read-Text 'src\fxfile\option.cpp'
$menuBar = Read-Text 'src\fxfile\gui\rebar\MenuBar.cpp'
$popupMenu = Read-Text 'src\fxfile\gui\BCMenu.cpp'
$toolBarHeader = Read-Text 'src\fxfile\gui\rebar\ToolBarEx.h'
$toolBar = Read-Text 'src\fxfile\gui\rebar\ToolBarEx.cpp'
$address = Read-Text 'src\fxfile\address_bar.cpp'
$folder = Read-Text 'src\fxfile\folder_ctrl.cpp'
$statusBar = Read-Text 'src\fxfile\gui\StatusBar.cpp'
$statusBarHeader = Read-Text 'src\fxfile\gui\StatusBar.h'
$tabCtrl = Read-Text 'src\fxfile\gui\TabCtrl.cpp'
$tabCtrlHeader = Read-Text 'src\fxfile\gui\TabCtrl.h'
$explorerPane = Read-Text 'src\fxfile\explorer_pane.cpp'
$explorerView = Read-Text 'src\fxfile\explorer_view.cpp'
$explorerHeader = Read-Text 'src\fxfile\explorer_ctrl.h'
$explorer = Read-Text 'src\fxfile\explorer_ctrl.cpp'

$scaleFactor = Function-Body $option 'double Option::getScaleFactor' 'double Option::getToolbarScaleFactor'
$toolbarScale = Function-Body $option 'double Option::getToolbarScaleFactor' 'void Option::scaleLogFont'
$fontScale = Function-Body $option 'void Option::scaleLogFont' 'void Option::getScaledFont'
$reportSelection = Function-Body $explorer 'void ExplorerCtrl::applyReportSelectionDrawState' 'void ExplorerCtrl::drawFinalReportSelection'
$selectionFill = Function-Body $explorer 'void ExplorerCtrl::drawFinalReportSelection' 'void ExplorerCtrl::drawParentFolderIcon'
$parentPaint = Function-Body $explorer 'void ExplorerCtrl::drawParentFolderIcon' 'void ExplorerCtrl::redrawFocusItemChange'
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$postPaint = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT)' '}'
$toolbarUpdate = Function-Body $toolBar 'void CToolBarEx::UpdateToolbarSize' 'BOOL CToolBarEx::GetButtonInfo'

Check '25/50/75 percent use one readable compact-density mapping while 100+ remain exact' (
    $scaleFactor.Contains('if (sRequestedScale < 1.0)') -and
    $scaleFactor.Contains('return 0.5 + (sRequestedScale * 0.5);') -and
    $scaleFactor.Contains('return sRequestedScale;'))
Check 'Toolbar compatibility factor cannot diverge from the common factor' (
    $toolbarScale.Contains('return getScaleFactor();') -and
    -not $toolbarScale.Contains('* 1.3'))
Check 'All font scaling is centralized and has a readable Korean glyph floor' (
    $optionHeader.Contains('static void scaleLogFont(LOGFONT &aLogFont);') -and
    $fontScale.Contains('kMinimumReadableFontHeight = 10') -and
    $fontScale.Contains('sScaledHeight') -and
    $option.Contains('Option::scaleLogFont(aOutLogFont);'))
Check 'Menu bar and every owner-drawn popup menu use the shared font scaler' (
    $menuBar.Contains('Option::getScaleFactor();') -and
    $menuBar.Contains('Option::scaleLogFont(sLf);') -and
    -not $menuBar.Contains('Option::getToolbarScaleFactor();') -and
    ([regex]::Matches($popupMenu, 'Option::scaleLogFont\(').Count -eq 5) -and
    -not $popupMenu.Contains('lfMenuFont.lfHeight = (LONG)(ncm.lfMenuFont.lfHeight * sScale)'))
Check 'Address bar and folder tree cannot bypass the readable font floor' (
    $address.Contains('Option::scaleLogFont(sLogFont);') -and
    $folder.Contains('Option::scaleLogFont(sLogFont);') -and
    -not $address.Contains('sLogFont.lfHeight = (LONG)(sLogFont.lfHeight * sScale)') -and
    -not $folder.Contains('sLogFont.lfHeight = (LONG)(sLogFont.lfHeight * sScale)'))
Check 'Default file lists, tabs and status bars cannot bypass the common scale' (
    $explorer.Contains('#include "option.h"') -and
    -not $explorer.Contains('SetFont(XPR_NULL)') -and
    ([regex]::Matches($explorer, '::fxfile::Option::getScaledFont\(sLogFont\);').Count -eq 2) -and
    ([regex]::Matches($explorer, '::fxfile::Option::scaleLogFont\(sLogFont\);').Count -eq 2) -and
    $statusBarHeader.Contains('void updateUIScale(void);') -and
    $statusBar.Contains('fxfile::Option::scaleLogFont(sLogFont);') -and
    $tabCtrlHeader.Contains('void updateUIScale(void);') -and
    $tabCtrl.Contains('fxfile::Option::scaleLogFont(sLogFont);'))
Check 'Runtime option changes reapply scale to every pane status bar and view tab' (
    $explorerPane.Contains('mStatusBar->updateUIScale();') -and
    $explorerView.Contains('mTabCtrl->updateUIScale();'))
Check 'All CToolBarEx descendants refresh their persistent scaled font before autosize' (
    $toolBarHeader.Contains('CFont           m_fontUIScale;') -and
    $toolbarUpdate.Contains('Option::getScaledFont(sLogFont);') -and
    $toolbarUpdate.Contains('SetFont(&m_fontUIScale, FALSE);') -and
    $toolbarUpdate.IndexOf('SetFont(&m_fontUIScale, FALSE);', [StringComparison]::Ordinal) -lt
        $toolbarUpdate.IndexOf('tbCtrl.AutoSize();', [StringComparison]::Ordinal))
Check 'Parent report row uses native buffered selection paint and overlays only its icon' (
    $explorerHeader.Contains('drawParentFolderIcon') -and
    $parentPaint.Contains('::DrawIconEx(') -and
    -not $parentPaint.Contains('::FillRect(') -and
    -not $parentPaint.Contains('::DrawText(') -and
    $customDraw.Contains('drawParentFolderIcon(sNmLvCustomDraw);') -and
    $customDraw.Contains('*aResult = CDRF_DODEFAULT;'))
Check 'Only live native selection controls parent and ordinary-row selection colours' (
    $reportSelection.Contains('GetItemState(sItemIndex, LVIS_SELECTED | LVIS_FOCUSED)') -and
    $reportSelection.Contains('if (!XPR_TEST_BITS(sNativeState, LVIS_SELECTED))') -and
    -not $reportSelection.Contains('XPR_TEST_BITS(aNmLvCustomDraw->nmcd.uItemState, CDIS_SELECTED) ||'))
Check 'Parent-row correction is read-only and cannot introduce repaint or selection loops' (
    -not $parentPaint.Contains('SetItemState(') -and
    -not $parentPaint.Contains('SetSelectionMark(') -and
    -not $parentPaint.Contains('Invalidate(') -and
    -not $parentPaint.Contains('RedrawItems(') -and
    -not $parentPaint.Contains('PostMessage(') -and
    -not $parentPaint.Contains('SetTimer(') -and
    -not $customDraw.Contains('SetItemState('))
Check 'Selected report rows complete icon and text after the final theme-proof background fill' (
    $customDraw.Contains('*aResult = CDRF_NEWFONT | CDRF_NOTIFYSUBITEMDRAW;') -and
    $customDraw.Contains('applyReportSelectionDrawState(sNmLvCustomDraw);') -and
    $customDraw.Contains('drawFinalReportSelection(sNmLvCustomDraw);') -and
    $selectionFill.Contains('::FillRect(') -and
    $selectionFill.Contains('mSmallImgList->Draw(') -and
    $selectionFill.Contains('::DrawText('))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | Format-Table -AutoSize
if ($failed.Count -gt 0) {
    throw ('Task 108 contract failure(s): ' + (($failed | ForEach-Object Name) -join '; '))
}

Write-Host 'PASS: Task 108 parent-row flash and unified readable UI scale contracts hold.'
