[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$explorer = [IO.File]::ReadAllText((Join-Path $root 'src\fxfile\explorer_ctrl.cpp'))
$definition = [IO.File]::ReadAllText((Join-Path $root 'src\fxfile\fxfile_def.h'))
$explorerView = [IO.File]::ReadAllText((Join-Path $root 'src\fxfile\explorer_view.cpp'))

function Function-Body([string]$text, [string]$signature, [string]$nextSignature) {
    $start = $text.IndexOf($signature, [StringComparison]::Ordinal)
    if ($start -lt 0) { return '' }
    $end = $text.IndexOf($nextSignature, $start + $signature.Length, [StringComparison]::Ordinal)
    if ($end -lt 0) { $end = $text.Length }
    return $text.Substring($start, $end - $start)
}

$onCreate = Function-Body $explorer 'xpr_sint_t ExplorerCtrl::OnCreate' 'void ExplorerCtrl::cacheRowFocusOption'
$focusTruth = Function-Body $explorer 'xpr_bool_t ExplorerCtrl::isFocusedSelectedItem' 'void ExplorerCtrl::resetCustomDrawColors'
$reportSelection = Function-Body $explorer 'void ExplorerCtrl::applyReportSelectionDrawState' 'void ExplorerCtrl::drawFinalReportSelection'
$selectionFill = Function-Body $explorer 'void ExplorerCtrl::drawFinalReportSelection' 'void ExplorerCtrl::drawParentFolderIcon'
$parentPaint = Function-Body $explorer 'void ExplorerCtrl::drawParentFolderIcon' 'void ExplorerCtrl::redrawFocusItemChange'
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'

$checks = [Collections.Generic.List[object]]::new()
function Check([string]$name, [bool]$passed) {
    $checks.Add([pscustomobject]@{ Name = $name; Passed = $passed })
}

Check 'ListView native double buffering is enabled after themed control setup' (
    $onCreate.Contains('enableVistaEnhanced(XPR_TRUE);') -and
    $onCreate.Contains('GetExtendedStyle();') -and
    $onCreate.Contains('SetExtendedStyle(sExtendedStyle | LVS_EX_DOUBLEBUFFER);') -and
    $onCreate.IndexOf('enableVistaEnhanced(XPR_TRUE);', [StringComparison]::Ordinal) -lt
        $onCreate.IndexOf('LVS_EX_DOUBLEBUFFER', [StringComparison]::Ordinal))

Check 'Live selection is the mandatory gate for every visual focus row' (
    $focusTruth.Contains('GetItemState(aItem, LVIS_SELECTED | LVIS_FOCUSED)') -and
    $focusTruth.Contains('if (!XPR_TEST_BITS(sNativeState, LVIS_SELECTED))') -and
    $focusTruth.IndexOf('if (!XPR_TEST_BITS(sNativeState, LVIS_SELECTED))', [StringComparison]::Ordinal) -lt
        $focusTruth.IndexOf('mRowFocusPaintItemIndex == aItem', [StringComparison]::Ordinal))

Check 'A newly focused selected row closes the PREPAINT snapshot timing gap' (
    $focusTruth.Contains('if (XPR_TEST_BITS(sNativeState, LVIS_FOCUSED))') -and
    $focusTruth.Contains('return XPR_TRUE;'))

Check 'Ctrl and Shift selections keep one native keyboard-focus identity' (
    -not $focusTruth.Contains('GetItemState(aItem, LVIS_SELECTED) & LVIS_SELECTED') -and
    -not $focusTruth.Contains('if (XPR_TEST_BITS(sNativeState, LVIS_SELECTED))') -and
    $focusTruth.Contains('mRowFocusPaintItemIndex == aItem'))

Check 'Selected parent row uses one explicit background and a shared-resource icon overlay' (
    $parentPaint.Contains('::DrawIconEx(') -and
    -not $parentPaint.Contains('::FillRect(') -and
    -not $parentPaint.Contains('::DrawText(') -and
    -not $parentPaint.Contains('DESTROY_ICON') -and
    $reportSelection.Contains('applyRowFocusDrawState(aNmLvCustomDraw);') -and
    $selectionFill.Contains('::FillRect(') -and
    $selectionFill.Contains('AfxGetApp()->LoadIcon(IDI_GO_UP)') -and
    -not $selectionFill.Contains('DESTROY_ICON') -and
    $customDraw.Contains('drawFinalReportSelection(sNmLvCustomDraw);'))

Check 'Parent icon overlay cannot open a second background and text paint transaction' (
    -not $parentPaint.Contains('::SaveDC(') -and
    -not $parentPaint.Contains('::RestoreDC(') -and
    -not $parentPaint.Contains('::FillRect('))

Check 'Paint refactor cannot mutate selection or schedule repaint loops' (
    -not $focusTruth.Contains('SetItemState(') -and
    -not $parentPaint.Contains('SetItemState(') -and
    -not $parentPaint.Contains('SetSelectionMark(') -and
    -not $parentPaint.Contains('Invalidate(') -and
    -not $parentPaint.Contains('RedrawItems(') -and
    -not $parentPaint.Contains('PostMessage(') -and
    -not $parentPaint.Contains('SetTimer('))

Check 'All six panes continue to share the same ExplorerCtrl paint implementation' (
    $definition.Contains('#define MAX_VIEW_SPLIT_ROW        (2)') -and
    $definition.Contains('#define MAX_VIEW_SPLIT_COLUMN     (3)') -and
    $explorerView.Contains('new ExplorerCtrl') -and
    $customDraw.Contains('drawParentFolderIcon(sNmLvCustomDraw);'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | Format-Table -AutoSize
if ($failed.Count -gt 0) {
    throw ('Task 110 contract failure(s): ' + (($failed | ForEach-Object Name) -join '; '))
}

Write-Host 'PASS: Task 110 selected-parent visibility and buffered paint contracts hold.'
