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
$explorerHeader = Read-Text 'src\fxfile\explorer_ctrl.h'
$watcher = Read-Text 'src\fxfile\adv_file_change_watcher.cpp'
$watcherHeader = Read-Text 'src\fxfile\adv_file_change_watcher.h'
$pane = Read-Text 'src\fxfile\explorer_pane.cpp'
$mainFrame = Read-Text 'src\fxfile\main_frame.cpp'
$deployment = Read-Text 'tools\Build-Deploy-Verify.ps1'

$finalPaint = Function-Body $explorer 'void ExplorerCtrl::drawFinalReportSelection' 'void ExplorerCtrl::drawParentFolderIcon'
$watchFileChange = Function-Body $explorer 'void ExplorerCtrl::watchFileChange' 'void ExplorerCtrl::addParentItem'
$advancedNotify = Function-Body $explorer 'LRESULT ExplorerCtrl::OnAdvFileChangeNotify' 'void ExplorerCtrl::enumerateShcn'
$timer = Function-Body $explorer 'void ExplorerCtrl::OnTimer' 'xpr_bool_t ExplorerCtrl::doPasteSelect'
$completion = Function-Body $watcher '    void OnCompletionRoutine' '    void OnFileChanged'
$registerTask = Function-Body $watcher 'void AdvFileChangeWatcher::registerTask' 'void AdvFileChangeWatcher::modifyTask'

Check 'Selected report text maps header client coordinates into the scrolled list client' (
    $finalPaint.Contains('ClientToScreen(&sHeaderOrigin)') -and
    $finalPaint.Contains('ScreenToClient(&sHeaderOrigin)') -and
    $finalPaint.Contains('sHeaderRect.OffsetRect(sHeaderOrigin.x, 0)'))

Check 'The same scroll-corrected column geometry owns selected text and grid separators' (
    ([regex]::Matches($finalPaint, 'sHeaderRect\.OffsetRect\(sHeaderOrigin\.x, 0\)').Count -ge 2) -and
    $finalPaint.Contains('LVS_EX_GRIDLINES'))

Check 'All pane layouts still use the one shared ExplorerCtrl final paint implementation' (
    $pane.Contains('setExplorerOption') -and
    $explorerHeader.Contains('drawFinalReportSelection'))

Check 'Startup pane windows remain hidden until atomic batch publication' (
    $mainFrame.Contains('DWORD sStyle = WS_CHILD | WS_BORDER;') -and
    $mainFrame.Contains('if (XPR_IS_FALSE(mDeferStartupViews))') -and
    $mainFrame.Contains('sStyle |= WS_VISIBLE;'))

Check 'Smoke expected pane count honors the active locked split snapshot' (
    $deployment.Contains("'^main\.view\.split_locked\s*=\s*(\d+)'") -and
    $deployment.Contains("'^main\.view\.locked_row_count\s*=\s*(\d+)'") -and
    $deployment.Contains("'^main\.view\.locked_column_count\s*=\s*(\d+)'") -and
    $deployment.Contains('if ($splitLocked -and $lockedRowCount -ge 1 -and $lockedColumnCount -ge 1)'))

Check 'Advanced watcher exposes an explicit asynchronous registration or rearm failure event' (
    $watcherHeader.Contains('EventWatchFailed') -and
    $watcher.Contains('appendWatchFailedNotifies'))

Check 'Initial ReadDirectoryChangesW failure is reported instead of returning a false live watch id' (
    $registerTask.Contains('readDirectoryChanges()') -and
    $registerTask.Contains('queueWatchFailedNotify'))

Check 'A completed watch that cannot rearm reports both full reconciliation and fallback' (
    $completion.Contains('appendUpdateDirNotifies') -and
    $completion.Contains('appendWatchFailedNotifies'))

Check 'ExplorerCtrl switches a failed advanced watch to the legacy directory watcher' (
    $advancedNotify.Contains('EventWatchFailed') -and
    $advancedNotify.Contains('watchFileChangeLegacy()') -and
    $watchFileChange.Contains('watchFileChangeLegacy()'))

Check 'Transient exact-event reconciliation failure schedules one path-bound delayed refresh' (
    $advancedNotify.Contains('scheduleDirectoryRefresh()') -and
    $explorer.Contains('mDeferredDirectoryRefreshPath') -and
    $explorer.Contains('TM_ID_NOTIFY_RECONCILE'))

Check 'Delayed refresh rejects stale navigation and still honors NoRefresh' (
    $timer.Contains('TM_ID_NOTIFY_RECONCILE') -and
    $explorer.Contains('mOption.mNoRefresh') -and
    $explorer.Contains('_tcsicmp(mDeferredDirectoryRefreshPath.c_str(), getCurPath())'))

Check 'Refresh recovery is event driven and does not add an unconditional polling timer' (
    -not $explorer.Contains('TM_ID_PERIODIC_REFRESH') -and
    -not $explorer.Contains('TM_ID_REFRESH_POLL'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object {
    '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name
}
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
