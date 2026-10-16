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
$header = Read-Text 'src\fxfile\explorer_ctrl.h'
$pane = Read-Text 'src\fxfile\explorer_pane.cpp'
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$itemPrepaint = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT)'
$itemPostpaint = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT)' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$finalPaint = Function-Body $explorer 'void ExplorerCtrl::drawFinalReportSelection' 'void ExplorerCtrl::drawParentFolderIcon'

Check 'All six panes still feed independent focus colours into the shared ExplorerCtrl' (
    $pane.Contains('aOption.mConfig.mFileListRowFocusColor[mViewIndex]') -and
    $header.Contains('drawFinalReportSelection'))

Check 'A live selected report item requests a final callback after native theme painting' (
    $itemPrepaint.Contains('LVIS_SELECTED') -and
    $itemPrepaint.Contains('CDRF_NOTIFYPOSTPAINT') -and
    $itemPrepaint.IndexOf('CDRF_NOTIFYPOSTPAINT', [StringComparison]::Ordinal) -gt
        $itemPrepaint.IndexOf('applyReportSelectionDrawState', [StringComparison]::Ordinal))

Check 'The final callback owns every selected report row rather than only the parent item' (
    $itemPostpaint.Contains('isReportView()') -and
    $itemPostpaint.Contains('drawFinalReportSelection(sNmLvCustomDraw);') -and
    $itemPostpaint.IndexOf('drawFinalReportSelection', [StringComparison]::Ordinal) -lt
        $itemPostpaint.IndexOf('drawParentFolderIcon', [StringComparison]::Ordinal))

Check 'Final painting rechecks live selection and runs after theme inside a saved DC transaction' (
    $finalPaint.Contains('GetItemState(sItemIndex') -and
    $finalPaint.Contains('LVIS_SELECTED') -and
    $finalPaint.Contains('::SaveDC(') -and
    $finalPaint.Contains('::FillRect(') -and
    $finalPaint.Contains('::RestoreDC(') -and
    $finalPaint.IndexOf('::FillRect(', [StringComparison]::Ordinal) -lt
        $finalPaint.IndexOf('::RestoreDC(', [StringComparison]::Ordinal))

Check 'Final painting restores native report content including all columns and ellipsis' (
    $finalPaint.Contains('mHeaderCtrl->GetItemCount()') -and
    $finalPaint.Contains('GetItemText(sItemIndex, sColumn)') -and
    $finalPaint.Contains('LVCF_FMT') -and
    $finalPaint.Contains('DT_END_ELLIPSIS') -and
    $finalPaint.Contains('DT_VCENTER'))

Check 'File folder and parent icons retain overlay cut and shared-resource ownership semantics' (
    $finalPaint.Contains('mSmallImgList->Draw(') -and
    $finalPaint.Contains('LVIS_OVERLAYMASK') -and
    $finalPaint.Contains('LVIS_CUT') -and
    $finalPaint.Contains('IDI_GO_UP') -and
    -not $finalPaint.Contains('DESTROY_ICON'))

Check 'Full-row legacy-cell focus and grid lines are reconstructed without changing options' (
    $finalPaint.Contains('LVIR_BOUNDS : LVIR_SELECTBOUNDS') -and
    $finalPaint.Contains('LVS_EX_GRIDLINES') -and
    $finalPaint.Contains('DrawFocusRect'))

Check 'Hover is paint input only and cannot become selection truth or start a repaint loop' (
    -not $finalPaint.Contains('CDIS_HOT') -and
    -not $finalPaint.Contains('SetItemState(') -and
    -not $finalPaint.Contains('SetSelectionMark(') -and
    -not $finalPaint.Contains('Invalidate(') -and
    -not $finalPaint.Contains('RedrawItems(') -and
    -not $finalPaint.Contains('SetTimer(') -and
    -not $finalPaint.Contains('PostMessage('))

Check 'The obsolete pre-theme background helper is no longer part of the paint path' (
    -not $header.Contains('fillReportSelectionBackground') -and
    -not $itemPrepaint.Contains('fillReportSelectionBackground'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
