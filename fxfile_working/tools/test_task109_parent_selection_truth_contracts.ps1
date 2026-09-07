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

$explorer = Read-Text 'src\fxfile\explorer_ctrl.cpp'
$explorerHeader = Read-Text 'src\fxfile\explorer_ctrl.h'
$explorerView = Read-Text 'src\fxfile\explorer_view.cpp'
$definition = Read-Text 'src\fxfile\fxfile_def.h'

$reportSelection = Function-Body $explorer 'void ExplorerCtrl::applyReportSelectionDrawState' 'void ExplorerCtrl::fillReportSelectionBackground'
$parentPaint = Function-Body $explorer 'void ExplorerCtrl::drawParentFolderIcon' 'void ExplorerCtrl::redrawFocusItemChange'
$focusRedraw = Function-Body $explorer 'void ExplorerCtrl::redrawFocusItemChange' 'void ExplorerCtrl::OnCustomdraw('
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$postPaintStart = $customDraw.IndexOf('else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT)', [StringComparison]::Ordinal)
$postPaint = if ($postPaintStart -ge 0) { $customDraw.Substring($postPaintStart) } else { '' }

$checks = [Collections.Generic.List[object]]::new()
function Check([string]$name, [bool]$passed) {
    $checks.Add([pscustomobject]@{ Name = $name; Passed = $passed })
}

Check 'Shared report-selection helper rechecks live ListView selection before painting' (
    $explorerHeader.Contains('applyReportSelectionDrawState') -and
    $reportSelection.Contains('GetItemState(sItemIndex, LVIS_SELECTED | LVIS_FOCUSED)') -and
    $reportSelection.Contains('if (!XPR_TEST_BITS(sNativeState, LVIS_SELECTED))'))

Check 'ITEMPOSTPAINT overlays only the parent icon and returns the supported default result' (
    $postPaint.Contains('drawParentFolderIcon(sNmLvCustomDraw);') -and
    $postPaint.Contains('*aResult = CDRF_DODEFAULT;') -and
    -not $postPaint.Contains('CDRF_SKIPDEFAULT'))

Check 'Delayed callback and visual-focus cache cannot promote an unselected parent row' (
    -not $postPaint.Contains('XPR_TEST_BITS(sNmLvCustomDraw->nmcd.uItemState, CDIS_SELECTED) ||') -and
    -not $reportSelection.Contains('XPR_TEST_BITS(aNmLvCustomDraw->nmcd.uItemState, CDIS_SELECTED) ||') -and
    -not $postPaint.Contains('XPR_TEST_BITS(sNativeState, LVIS_SELECTED) ||'))

Check 'Parent correction remains paint-only and cannot mutate selection or schedule loops' (
    -not $parentPaint.Contains('SetItemState(') -and
    -not $parentPaint.Contains('SetSelectionMark(') -and
    -not $parentPaint.Contains('Invalidate(') -and
    -not $parentPaint.Contains('RedrawItems(') -and
    -not $parentPaint.Contains('PostMessage(') -and
    -not $parentPaint.Contains('SetTimer('))

Check 'Normal selection changes invalidate both old and new focus rows' (
    $focusRedraw.Contains('GetItemRect(aOldItem, &sOldRect, LVIR_BOUNDS)') -and
    $focusRedraw.Contains('GetItemRect(aNewItem, &sNewRect, LVIR_BOUNDS)') -and
    $focusRedraw.Contains('InvalidateRect(&sOldRect, XPR_FALSE);') -and
    $focusRedraw.Contains('InvalidateRect(&sNewRect, XPR_FALSE);'))

Check 'All six panes use the shared ExplorerCtrl implementation' (
    $definition.Contains('#define MAX_VIEW_SPLIT_ROW        (2)') -and
    $definition.Contains('#define MAX_VIEW_SPLIT_COLUMN     (3)') -and
    $explorerView.Contains('new ExplorerCtrl'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | Format-Table -AutoSize
if ($failed.Count -gt 0) {
    throw ('Task 109 contract failure(s): ' + (($failed | ForEach-Object Name) -join '; '))
}

Write-Host 'PASS: Task 109 parent-row selection truth is enforced for the shared six-pane control.'
