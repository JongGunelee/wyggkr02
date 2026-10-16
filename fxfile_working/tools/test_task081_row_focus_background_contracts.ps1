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
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$focusState = Function-Body $explorer 'void ExplorerCtrl::applyRowFocusDrawState' 'void ExplorerCtrl::applyReportSelectionDrawState'
$reportSelection = Function-Body $explorer 'void ExplorerCtrl::applyReportSelectionDrawState' 'void ExplorerCtrl::drawParentFolderIcon'
$subitemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM))' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT)'

Check 'Selected report cells receive the configured background and contrast text in native subitem painting' (
    $focusState.Contains('clrTextBk = mOption.mRowFocusColor;') -and
    $focusState.Contains('clrText   = mRowFocusTextColor;') -and
    $reportSelection.Contains('applyRowFocusDrawState(aNmLvCustomDraw);') -and
    $customDraw.Contains('CDRF_NEWFONT'))
Check 'Every live-selected row receives application-owned final paint; the list selection model is not changed' (
    $customDraw.Contains('CDDS_ITEMPREPAINT') -and
    $reportSelection.Contains('if (!XPR_TEST_BITS(sNativeState, LVIS_SELECTED))') -and
    -not $reportSelection.Contains('isFocusedSelectedItem(sItemIndex)') -and
    -not $customDraw.Contains('SetItemState(') -and
    -not $customDraw.Contains('SetItem('))
Check 'Configured contrast text leaves icon, overlay and alignment rendering to the native list control' (
    $focusState.Contains('clrText   = mRowFocusTextColor;') -and
    $customDraw.Contains('CDRF_NEWFONT') -and
    -not $customDraw.Contains('mSmallImgList->Draw') -and
    -not $customDraw.Contains('DrawText('))
Check 'Full-row and legacy first-cell scopes remain mutually consistent' (
    $customDraw.Contains('XPR_IS_TRUE(mOption.mFullRowSelect)') -and
    $customDraw.Contains('applyReportSelectionDrawState(sNmLvCustomDraw);') -and
    $customDraw.Contains('CDRF_NEWFONT'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
