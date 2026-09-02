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
$fillFocus = Function-Body $explorer 'void ExplorerCtrl::fillRowFocusBackground' 'void ExplorerCtrl::applyRowFocusDrawState'
$focusState = Function-Body $explorer 'void ExplorerCtrl::applyRowFocusDrawState' 'void ExplorerCtrl::OnCustomdraw('
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$subItemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM))' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)'
$itemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT)'
$itemChanged = Function-Body $explorer 'void ExplorerCtrl::OnItemchanged' 'xpr_size_t ExplorerCtrl::getRealSelCount'

Check 'The report item notification applies native colors only to the cached focused row' (
    $itemDraw.Contains('XPR_IS_TRUE(isReportView())') -and
    $itemDraw.Contains('XPR_IS_TRUE(sFocusedSelected)') -and
    $itemDraw.Contains('fillRowFocusBackground(sNmLvCustomDraw);') -and
    $itemDraw.Contains('CDRF_NEWFONT | CDRF_NOTIFYSUBITEMDRAW'))
Check 'The native item paint owns only the configured row scope after Windows selection compositing' (
    $fillFocus.Contains('LVIR_BOUNDS : LVIR_SELECTBOUNDS') -and
    $fillFocus.Contains('::FillRect(') -and
    $subItemDraw.Contains('XPR_IS_TRUE(mOption.mFullRowSelect) || sNmLvCustomDraw->iSubItem == 0') -and
    $subItemDraw.Contains('applyRowFocusDrawState(sNmLvCustomDraw);') -and
    $focusState.Contains('clrTextBk = mOption.mRowFocusColor;') -and
    $focusState.Contains('clrText   = mRowFocusTextColor;'))
Check 'Paint-state suppression never mutates the list selection model or schedules work' (
    -not $customDraw.Contains('SetItemState(') -and
    -not $customDraw.Contains('SetItem(') -and
    -not $customDraw.Contains('SetTimer(') -and
    -not $customDraw.Contains('PostMessage(') -and
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
    $customDraw.Contains('XPR_IS_TRUE(sFocusedSelected)'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
