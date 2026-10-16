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

$toolbar = Read-Text 'src\fxfile\main_toolbar.cpp'
$toolbarHeader = Read-Text 'src\fxfile\main_toolbar.h'
$coolbar = Read-Text 'src\fxfile\main_coolbar.cpp'
$mainFrame = Read-Text 'src\fxfile\main_frame.cpp'
$option = Read-Text 'src\fxfile\option.cpp'
$command = Read-Text 'src\fxfile\cmd\cmd_cfg.cpp'
$clockCtrl = Read-Text 'src\fxfile\clock_ctrl.cpp'

$showClock = Function-Body $toolbar 'void MainToolBar::showClock(xpr_bool_t aShow)' 'void MainToolBar::updateClockLayout(void)'
$updateClock = Function-Body $toolbar 'void MainToolBar::updateClock(void)' 'void MainToolBar::showClock(xpr_bool_t aShow)'
$updateLayout = Function-Body $toolbar 'void MainToolBar::updateClockLayout(void)' 'xpr_bool_t MainToolBar::HasButtonText'
$updatedSize = Function-Body $coolbar 'void MainCoolBar::onUpdatedToolbarSize(CToolBarEx &theToolBar)' 'void MainCoolBar::OnSysColorChange'
$frameShow = Function-Body $mainFrame 'void MainFrame::setShowClock(xpr_bool_t aShow)' 'void MainFrame::applyUIScale(void)'

$showPos = $showClock.IndexOf('mClockCtrl.ShowWindow(aShow ? SW_SHOW : SW_HIDE);', [StringComparison]::Ordinal)
$sizePos = $showClock.IndexOf('UpdateToolbarSize();', [StringComparison]::Ordinal)
$layoutPos = $showClock.IndexOf('updateClockLayout();', [StringComparison]::Ordinal)

Check 'Existing saved option and checked menu command remain the single show-clock state' (
    $option.Contains('"main.clock.show"') -and
    $option.Contains('&Option::mMain.mShowClock') -and
    $command.Contains('sMainFrame->setShowClock(!sMainFrame->isShowClock());') -and
    $frameShow.Contains('OptionManager::instance().saveMainOption();'))
Check 'Main toolbar publishes visible clock ideal and minimum width reservations' (
    $toolbarHeader.Contains('getClockIdealReservedWidth') -and
    $toolbarHeader.Contains('getClockMinimumReservedWidth') -and
    $toolbar.Contains('MainToolBar::getClockIdealReservedWidth') -and
    $toolbar.Contains('MainToolBar::getClockMinimumReservedWidth') -and
    $toolbar.Contains('mClockCtrl.GetStyle(), WS_VISIBLE'))
Check 'Rebar sizing includes clock width only for the main toolbar band' (
    $updatedSize.Contains('if (&theToolBar == &mMainToolBar)') -and
    $updatedSize.Contains('getClockIdealReservedWidth()') -and
    $updatedSize.Contains('getClockMinimumReservedWidth()'))
Check 'Requested-visible clock guarantees a nonzero main-toolbar band row' (
    $updatedSize.Contains('cyChild = max(cyChild, sClockRowHeight);') -and
    $updatedSize.Contains('A saved rebar state can restore a zero-height main-toolbar'))
Check 'Insufficient inline width moves the clock to a dedicated responsive toolbar row' (
    $toolbarHeader.Contains('mClockSeparateRow') -and
    $toolbarHeader.Contains('getClockRowHeight') -and
    $updateLayout.Contains('sAvailableWidth < sMinimumWidth') -and
    $updateLayout.Contains('sMaxBottom +') -and
    $updatedSize.Contains('cyChild += sClockRowHeight;'))
Check 'Show toggle recalculates the band before final clock placement' (
    $showPos -ge 0 -and $sizePos -gt $showPos -and $layoutPos -gt $sizePos)
Check 'Clock width is clamped to the toolbar client width instead of extending off-screen' (
    $updateLayout.Contains('sAvailableWidth') -and
    $updateLayout.Contains('min(sPreferredWidth, sAvailableWidth)') -and
    $updateLayout.Contains('sBarRect.Width() - minLeft - sRightMargin'))
Check 'Clock text adapts to full, medium, compact, and minimal layouts' (
    $updateClock.Contains('kClockFullTextWidth') -and
    $updateClock.Contains('kClockMediumTextWidth') -and
    $updateClock.Contains('kClockCompactTextWidth') -and
    $updateClock.Contains('_T("%02d:%02d")'))
Check 'Toolbar resize continues to drive responsive clock layout' (
    $toolbar.Contains('void MainToolBar::OnSize(UINT nType, int cx, int cy)') -and
    $toolbar.Contains('super::OnSize(nType, cx, cy);') -and
    $toolbar.Contains('updateClockLayout();'))
Check 'Startup timer repairs a requested-visible zero-sized clock after ancestor visibility settles' (
    $toolbar.Contains('During startup the toolbar can be initialized while an ancestor is') -and
    $toolbar.Contains('if (sClockRect.Width() <= 0 || sClockRect.Height() <= 0)') -and
    $toolbar.Contains('XPR_TEST_BITS(mClockCtrl.GetStyle(), WS_VISIBLE)'))
Check 'Clock position lock and unlocked drag persistence remain intact' (
    $toolbar.Contains('gOpt->mMain.mClockPosX') -and
    $clockCtrl.Contains('gOpt->mMain.mClockLocked') -and
    $clockCtrl.Contains('gOpt->mMain.mClockPosX = newLeft;') -and
    $clockCtrl.Contains('OptionManager::instance().saveMainOption();'))
Check 'Visibility repair does not create duplicate clock windows or timers' (
    ([regex]::Matches($toolbar, 'mClockCtrl\.Create\(').Count -eq 1) -and
    ([regex]::Matches($toolbar, 'SetTimer\(1055').Count -eq 1))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
