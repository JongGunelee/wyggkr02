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
$fillFocus = Function-Body $explorer 'void ExplorerCtrl::fillRowFocusBackground' 'void ExplorerCtrl::applyRowFocusDrawState'
$focusState = Function-Body $explorer 'void ExplorerCtrl::applyRowFocusDrawState' 'void ExplorerCtrl::OnCustomdraw('
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$subItemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM))' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)'
$itemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT)'
$reportItemDraw = Function-Body $itemDraw 'if (XPR_IS_TRUE(isReportView()) &&' 'else if (getViewStyle() != VIEW_STYLE_THUMBNAIL)'

Check 'Focused report items receive an explicit background before native icon and text paint' (
    $itemDraw.Contains('XPR_IS_TRUE(isReportView())') -and
    $itemDraw.Contains('XPR_IS_TRUE(sFocusedSelected)') -and
    $itemDraw.Contains('fillRowFocusBackground(sNmLvCustomDraw);') -and
    $itemDraw.Contains('CDRF_NEWFONT | CDRF_NOTIFYSUBITEMDRAW'))
Check 'Final paint colors every selected column in full-row mode and only column zero in legacy mode' (
    $fillFocus.Contains('LVIR_BOUNDS : LVIR_SELECTBOUNDS') -and
    $subItemDraw.Contains('XPR_IS_TRUE(mOption.mFullRowSelect) || sNmLvCustomDraw->iSubItem == 0') -and
    $subItemDraw.Contains('applyRowFocusDrawState(sNmLvCustomDraw);') -and
    $focusState.Contains('clrTextBk = mOption.mRowFocusColor;') -and
    $focusState.Contains('clrText   = mRowFocusTextColor;'))
Check 'Selected report cells suppress theme compositing while retaining native icon and text rendering' (
    $fillFocus.Contains('::FillRect(') -and
    $focusState.Contains('uItemState &= ~CDIS_SELECTED') -and
    -not [regex]::IsMatch($focusState, '(?m)^\s*aNmLvCustomDraw->iStateId\s*=') -and
    -not [regex]::IsMatch($focusState, '(?m)^\s*aNmLvCustomDraw->clrFace\s*=') -and
    -not $subItemDraw.Contains('CDRF_SKIPDEFAULT') -and
    -not $reportItemDraw.Contains('CDRF_SKIPDEFAULT') -and
    -not $customDraw.Contains('mRowFocusPaintPending'))
Check 'Selection remains read-only paint state throughout the corrected path' (
    -not $customDraw.Contains('SetItemState(') -and
    -not $customDraw.Contains('SetItem(') -and
    -not $fillFocus.Contains('SetItemState(') -and
    $explorer.Contains('mFocusedItemIndex') -and
    $explorer.Contains('CDDS_ITEMPREPAINT') -and
    $explorer.Contains('CDDS_SUBITEM'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
