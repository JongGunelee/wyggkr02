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
$explorerHeader = Read-Text 'src\fxfile\explorer_ctrl.h'
$pane = Read-Text 'src\fxfile\explorer_pane.cpp'
$driveToolbar = Read-Text 'src\fxfile\drive_toolbar.cpp'
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$filtering = Function-Body $explorer 'void ExplorerCtrl::applyCustomDrawFiltering' 'void ExplorerCtrl::OnCustomdraw'
$destroyDrivePathBar = Function-Body $pane 'void ExplorerPane::destroyDrivePathBar' 'DrivePathBar *ExplorerPane::getDrivePathBar'
$destroyDriveBar = Function-Body $driveToolbar 'void DriveToolBar::destroyDriveBar' 'void DriveToolBar::refresh'
$iconUpdate = Function-Body $driveToolbar 'LRESULT DriveToolBar::OnDriveIconUpdate' 'void DriveToolBar::createDriveBar'

Check 'Custom draw obtains selection from the list model, then handles full-row and first-cell paint separately at final item paint' (
    $explorerHeader.Contains('isFocusedSelectedItem') -and
    $explorer.Contains('mFocusedItemIndex') -and
    $customDraw.Contains('XPR_IS_TRUE(mOption.mFullRowSelect)') -and
    $customDraw.Contains('sNmLvCustomDraw->clrTextBk = mOption.mRowFocusColor;') -and
    $customDraw.Contains('sNmLvCustomDraw->clrText   = mRowFocusTextColor;') -and
    $customDraw.Contains('XPR_IS_TRUE(sFocusedSelected)') -and
    $customDraw.Contains('CDRF_NEWFONT'))
Check 'Item notifications restore filtering before the focused selection receives final application-owned paint' (
    $filtering.Contains('mOption.mTextColorType') -and
    ([regex]::Matches($customDraw, 'resetCustomDrawColors\(sNmLvCustomDraw\);\s*applyCustomDrawFiltering\(sNmLvCustomDraw\);').Count -eq 2) -and
    $customDraw.Contains('applyRowFocusDrawState(sNmLvCustomDraw);') -and
    $customDraw.Contains('CDDS_ITEMPREPAINT') -and
    $customDraw.Contains('CDRF_NEWFONT'))
Check 'Drive path bar is detached before native destruction and does not duplicate DriveToolBar cleanup' (
    $destroyDrivePathBar.Contains('DrivePathBar *sDrivePathBar = mDrivePathBar;') -and
    $destroyDrivePathBar.IndexOf('mDrivePathBar = XPR_NULL;', [StringComparison]::Ordinal) -lt $destroyDrivePathBar.IndexOf('DestroyWindow();', [StringComparison]::Ordinal) -and
    $destroyDrivePathBar.Contains('::IsWindow(sDrivePathBar->GetSafeHwnd())') -and
    -not $destroyDrivePathBar.Contains('sDrivePathBar->destroyDriveBar();'))
Check 'Drive toolbar refuses stale-window button updates while retaining deterministic worker cleanup' (
    $destroyDriveBar.Contains('::IsWindow(GetSafeHwnd()) == FALSE') -and
    $destroyDriveBar.Contains('mBarCreated = XPR_FALSE;') -and
    $iconUpdate.Contains('::IsWindow(GetSafeHwnd()) == FALSE'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
