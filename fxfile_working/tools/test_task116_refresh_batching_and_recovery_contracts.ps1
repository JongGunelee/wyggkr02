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
    $end = $text.IndexOf($nextSignature, $start + $signature.Length,
        [StringComparison]::Ordinal)
    if ($end -lt 0) { $end = $text.Length }
    $text.Substring($start, $end - $start)
}

$checks = [Collections.Generic.List[object]]::new()
function Check([string]$name, [bool]$passed) {
    $checks.Add([pscustomobject]@{ Name = $name; Passed = $passed })
}

$explorer = Read-Text 'src\fxfile\explorer_ctrl.cpp'
$header = Read-Text 'src\fxfile\explorer_ctrl.h'
$schedule = Function-Body $explorer 'void ExplorerCtrl::scheduleDirectoryRefresh' 'void ExplorerCtrl::reconcileDirectoryRefresh'
$reconcile = Function-Body $explorer 'void ExplorerCtrl::reconcileDirectoryRefresh' 'void ExplorerCtrl::scheduleRefreshSort'
$sortSchedule = Function-Body $explorer 'void ExplorerCtrl::scheduleRefreshSort' 'void ExplorerCtrl::addParentItem'
$timer = Function-Body $explorer 'void ExplorerCtrl::OnTimer' 'xpr_bool_t ExplorerCtrl::doPasteSelect'
$notify = Function-Body $explorer 'LRESULT ExplorerCtrl::OnAdvFileChangeNotify' 'void ExplorerCtrl::enumerateShcn'
$finish = Function-Body $explorer 'void ExplorerCtrl::endShcn' 'LRESULT ExplorerCtrl::OnFileChangeNotify'
$adaptive = Read-Text 'src\fxfile\adaptive_file_operation.cpp'
$buildPlan = Function-Body $adaptive 'bool buildPlan' 'bool createDirectories'
$winApp = Read-Text 'src\fxfile\win_app.cpp'

Check 'Continuous failed exact events cannot restart and starve the same-path reconcile timer' (
    $schedule.Contains('mDeferredDirectoryRefresh') -and
    $schedule.Contains('mDeferredDirectoryRefreshPath') -and
    $schedule.Contains('return;') -and
    -not $schedule.Contains('KillTimer(TM_ID_NOTIFY_RECONCILE);`r`n    SetTimer'))

Check 'Directory recovery has bounded retry state and increasing delay' (
    $header.Contains('mDeferredDirectoryRefreshRetryCount') -and
    $explorer.Contains('kDirectoryRefreshRetryLimit = 4') -and
    $explorer.Contains('{250, 500, 1000, 2000}') -and
    $reconcile.Contains('++mDeferredDirectoryRefreshRetryCount') -and
    $reconcile.Contains('>= kDirectoryRefreshRetryLimit'))

Check 'Successful and stale recovery clear path and retry ownership' (
    $reconcile.Contains('mDeferredDirectoryRefresh = XPR_FALSE') -and
    $reconcile.Contains('mDeferredDirectoryRefreshPath.clear()') -and
    $reconcile.Contains('mDeferredDirectoryRefreshRetryCount = 0') -and
    $reconcile.Contains('_tcsicmp(mDeferredDirectoryRefreshPath.c_str(), getCurPath())'))

Check 'Timer allocation failure is explicit and cannot silently discard initial recovery' (
    $schedule.Contains('SetTimer(TM_ID_NOTIFY_RECONCILE') -and
    $schedule.Contains('reconcileDirectoryRefresh();'))

Check 'A rename with a temporarily unavailable new PIDL always schedules full reconciliation' (
    $notify.Contains('if (XPR_IS_NULL(sFullPidl))') -and
    $notify.Contains('OnShcnDeleteItem(XPR_NULL, sOldPath.c_str())') -and
    $notify.Contains('scheduleDirectoryRefresh();'))

Check 'Notification sorting is fixed-window batched instead of repeated per exact event' (
    $finish.Contains('scheduleRefreshSort();') -and
    $header.Contains('mDeferredRefreshSort') -and
    $sortSchedule.Contains('mDeferredRefreshSort') -and
    $sortSchedule.Contains('TM_ID_NOTIFY_SORT') -and
    -not $finish.Contains('resortItems();'))

Check 'Sort timer preserves active inline rename and current refresh-sort setting' (
    $timer.Contains('TM_ID_NOTIFY_SORT') -and
    $timer.Contains('GetEditControl()') -and
    $timer.Contains('mRenameResorting') -and
    $timer.Contains('mOption.mRefreshSort == XPR_TRUE') -and
    $timer.Contains('resortItems();'))

Check 'Navigation and destruction cancel both pending recovery and sort work' (
    ([regex]::Matches($explorer, 'KillTimer\(TM_ID_NOTIFY_RECONCILE\)').Count -ge 1) -and
    ([regex]::Matches($explorer, 'KillTimer\(TM_ID_NOTIFY_SORT\)').Count -ge 1) -and
    ([regex]::Matches($explorer, 'mDeferredDirectoryRefreshRetryCount = 0').Count -ge 3) -and
    ([regex]::Matches($explorer, 'mDeferredRefreshSort = XPR_FALSE').Count -ge 3))

Check 'Recovery remains event-driven without unconditional directory polling' (
    -not $explorer.Contains('TM_ID_PERIODIC_REFRESH') -and
    -not $explorer.Contains('TM_ID_REFRESH_POLL'))

$movePreflight = $buildPlan.IndexOf('if (aOperation->wFunc == FO_MOVE)', [StringComparison]::Ordinal)
$treeWalk = $buildPlan.IndexOf('enumerateDirectory(aPlan', [StringComparison]::Ordinal)
Check 'Same-volume move eligibility is decided before recursive tree enumeration' (
    $movePreflight -ge 0 -and
    $treeWalk -gt $movePreflight -and
    $buildPlan.Contains('sProbeSource') -and
    $buildPlan.Contains('_wcsicmp(sProbeVolume.c_str(), sTargetVolume.c_str()) == 0') -and
    $buildPlan.Contains('return false;'))

Check 'Language-pack diagnostics format the size_t count without narrowing or vararg mismatch' (
    $winApp.Contains('Found Count: %Iu') -and
    $winApp.Contains('mLanguageTable->getLanguageCount()') -and
    -not $winApp.Contains('Found Count: %d'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object {
    '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name
}
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
