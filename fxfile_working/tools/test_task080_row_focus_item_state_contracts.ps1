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
$selectedFocus = Function-Body $explorer 'xpr_bool_t ExplorerCtrl::isFocusedSelectedItem' 'void ExplorerCtrl::resetCustomDrawColors'
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$thumbnailDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdrawThumbnail' 'LRESULT ExplorerCtrl::OnThumbnailProc'

Check 'Focused row target comes from each ListView prepaint snapshot, not the transient window focus owner' (
    $explorer.Contains('snapshotRowFocusItem') -and
    $selectedFocus.Contains('mRowFocusPaintItemIndex == aItem') -and
    -not $selectedFocus.Contains('::GetFocus()'))
Check 'Report view keeps the focused selected target through native row painting after a dialog or path-bar focus transition' (
    $customDraw.Contains('isFocusedSelectedItem(sItemIndex)') -and
    $customDraw.Contains('CDDS_ITEMPREPAINT') -and
    $customDraw.Contains('sNmLvCustomDraw->clrTextBk = mOption.mRowFocusColor;') -and
    $customDraw.Contains('sNmLvCustomDraw->clrText   = mRowFocusTextColor;') -and
    $customDraw.Contains('CDRF_NEWFONT'))
Check 'Thumbnail view uses the same item target instead of falling back to system colors when its list loses immediate focus' (
    $thumbnailDraw.Contains('sSelectedFocus = isFocusedSelectedItem(sItemIndex);') -and
    $thumbnailDraw.Contains('XPR_IS_TRUE(mOption.mFullRowSelect) && XPR_IS_TRUE(sSelectedFocus)') -and
    $thumbnailDraw.Contains('else if (XPR_IS_TRUE(sListHasFocus))'))
Check 'Painting remains read-only and does not change selection, navigation, enumeration, or scheduling state' (
    -not $customDraw.Contains('SetItemState(') -and
    -not $customDraw.Contains('SetItem(') -and
    -not $thumbnailDraw.Contains('SetItemState(') -and
    -not $thumbnailDraw.Contains('SetItem('))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
