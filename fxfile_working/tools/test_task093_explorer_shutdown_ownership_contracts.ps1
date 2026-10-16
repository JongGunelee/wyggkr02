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

$ctrl = Read-Text 'src\fxfile\explorer_ctrl.cpp'
$ctrlHeader = Read-Text 'src\fxfile\explorer_ctrl.h'
$pane = Read-Text 'src\fxfile\explorer_pane.cpp'
$view = Read-Text 'src\fxfile\explorer_view.cpp'
$shutdownProbe = Read-Text 'tools\Test-Task093RepeatedShutdown.ps1'

$ctrlDataStart = $pane.IndexOf('class ExplorerPane::ExplorerCtrlData', [StringComparison]::Ordinal)
$ctrlDataEnd = $pane.IndexOf('ExplorerPane::ExplorerPane', $ctrlDataStart, [StringComparison]::Ordinal)
$ctrlData = if ($ctrlDataStart -ge 0 -and $ctrlDataEnd -gt $ctrlDataStart) {
    $pane.Substring($ctrlDataStart, $ctrlDataEnd - $ctrlDataStart)
} else { '' }
$destroyOne = Function-Body $pane 'void ExplorerPane::destroySubPane(xpr_uint_t aId)' 'void ExplorerPane::destroySubPane(void)'
$destroyAll = Function-Body $pane 'void ExplorerPane::destroySubPane(void)' 'xpr_size_t ExplorerPane::getSubPaneCount'
$paneOnDestroy = Function-Body $pane 'void ExplorerPane::OnDestroy(void)' 'void ExplorerPane::setChangedOption'
$viewOnDestroy = Function-Body $view 'void ExplorerView::OnDestroy(void)' 'void ExplorerView::OnSize'
$ctrlOnDestroy = Function-Body $ctrl 'void ExplorerCtrl::OnDestroy(void)' 'void ExplorerCtrl::setObserver'
$deleteAllItems = Function-Body $ctrl 'void ExplorerCtrl::OnDeleteallitems' 'xpr_bool_t ExplorerCtrl::OnNotify'

$detachPos = $ctrlData.IndexOf('mExplorerCtrl = XPR_NULL;', [StringComparison]::Ordinal)
$destroyWindowPos = $ctrlData.IndexOf('DestroyWindow();', [StringComparison]::Ordinal)
$erasePos = $destroyOne.IndexOf('mExplorerCtrlMap.erase(sIterator);', [StringComparison]::Ordinal)
$deletePos = $destroyOne.IndexOf('XPR_SAFE_DELETE(sExplorerCtrlData);', [StringComparison]::Ordinal)
$viewDrainPos = $viewOnDestroy.IndexOf('mExplorerPane->destroySubPane();', [StringComparison]::Ordinal)
$tabDestroyPos = $viewOnDestroy.IndexOf('DESTROY_DELETE(mTabCtrl);', [StringComparison]::Ordinal)
$paneDrainPos = $paneOnDestroy.IndexOf('destroySubPane();', [StringComparison]::Ordinal)
$paneSuperPos = $paneOnDestroy.IndexOf('super::OnDestroy();', [StringComparison]::Ordinal)
$guardPos = $ctrlOnDestroy.IndexOf('if (XPR_IS_TRUE(mDestroying))', [StringComparison]::Ordinal)
$cancelPos = $ctrlOnDestroy.IndexOf('cancelAsyncImage(', [StringComparison]::Ordinal)

Check 'ExplorerCtrlData withdraws its raw member before native window destruction' (
    $ctrlData.Contains('ExplorerCtrl *sExplorerCtrl = mExplorerCtrl;') -and
    $detachPos -ge 0 -and $destroyWindowPos -gt $detachPos -and
    $ctrlData.Contains('::IsWindow(sHwnd)') -and
    $ctrlData.Contains('CWnd::FromHandlePermanent(sHwnd) == sExplorerCtrl'))
Check 'Single subpane removal erases the published map entry before any destructor side effect' (
    $erasePos -ge 0 -and $deletePos -gt $erasePos)
Check 'Bulk subpane removal swaps the live ownership map empty before deleting controls' (
    $destroyAll.Contains('sExplorerCtrlMap.swap(mExplorerCtrlMap);') -and
    $destroyAll.IndexOf('sExplorerCtrlMap.swap(mExplorerCtrlMap);', [StringComparison]::Ordinal) -lt
        $destroyAll.IndexOf('XPR_SAFE_DELETE(sExplorerCtrlData);', [StringComparison]::Ordinal))
Check 'ExplorerView drains shared explorer controls before TabCtrl destruction callbacks' (
    $viewDrainPos -ge 0 -and $tabDestroyPos -gt $viewDrainPos)
Check 'ExplorerPane retains an idempotent native-destroy fallback before base teardown' (
    $paneDrainPos -ge 0 -and $paneSuperPos -gt $paneDrainPos)
Check 'ExplorerCtrl shutdown is guarded before asynchronous owner cancellation' (
    $ctrlHeader.Contains('xpr_bool_t                       mDestroying;') -and
    $ctrl.Contains('mDestroying         = XPR_FALSE;') -and
    $guardPos -ge 0 -and $cancelPos -gt $guardPos -and
    $ctrlOnDestroy.Contains('mDestroying = XPR_TRUE;') -and
    $ctrlOnDestroy.Contains('HWND sHwnd = GetSafeHwnd();'))
Check 'Delete-all notification does not repeat thumbnail cancellation during native teardown' (
    $deleteAllItems.Contains('if (XPR_IS_FALSE(mDestroying))') -and
    $deleteAllItems.Contains('Thumbnail::instance().cancelAsyncImage'))
Check 'Repeated shutdown probe monitors both directory and ZIP BugTrap reports' (
    $shutdownProbe.Contains('$_.PSIsContainer -or $_.Extension -ieq ''.zip''') -and
    $shutdownProbe.Contains('$_.Name -like ''fxfile_error_report_*'''))
Check 'Repeated shutdown probe supplies a directory to every startup pane' (
    $shutdownProbe.Contains('Get-LaunchArguments $layout.Name $layout.Views $package.Root') -and
    -not $shutdownProbe.Contains('Get-LaunchArguments $layout.Name $layout.Views $exe'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
