param()

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

function Read-Text([string]$relativePath) {
    return [System.IO.File]::ReadAllText((Join-Path $root $relativePath))
}

$checks = New-Object System.Collections.Generic.List[object]
function Check([string]$name, [bool]$passed) {
    $checks.Add([pscustomobject]@{ Name = $name; Passed = $passed })
}

$thread = Read-Text 'src\fxfile\file_op_thread.cpp'
$threadHeader = Read-Text 'src\fxfile\file_op_thread.h'
$explorer = Read-Text 'src\fxfile\explorer_ctrl.cpp'
$explorerHeader = Read-Text 'src\fxfile\explorer_ctrl.h'
$adaptive = Read-Text 'src\fxfile\adaptive_file_operation.cpp'
$modern = Read-Text 'src\fxfile\modern_shell_file_operation.cpp'

Check 'Broken drive-root UPDATEDIR code removed' (-not $thread.Contains('_tcschr(sUpdatePath'))
Check 'Operation inputs captured before mutation' ($thread.Contains('captureOperationSources();'))
Check 'UI reconciliation runs on post-end UI thread' ($thread.Contains('reconcileOperationResult();'))
Check 'All visible panes receive one batch' ($thread.Contains('reconcileFileOperationItems('))
Check 'Source disappearance is captured by worker filesystem snapshot' (
    $thread.Contains('sSnapshot.mSourceGone =') -and
    $thread.Contains('::GetFileAttributes(sIt->mPath.c_str())'))
Check 'Target existence is checked from actual file system' ($thread.Contains('sItem.mTargetExists ='))
Check 'Small operations use exact rename event' ($thread.Contains('SHCNE_RENAMEFOLDER : SHCNE_RENAMEITEM'))
Check 'Small operations use exact delete event' ($thread.Contains('SHCNE_RMDIR : SHCNE_DELETE'))
Check 'Small operations use exact create event' ($thread.Contains('SHCNE_MKDIR : SHCNE_CREATE'))
Check 'Shell notification never blocks listeners' ($thread.Contains('SHCNF_PATH | SHCNF_FLUSHNOWAIT'))
Check 'Exact notification fan-out is conservatively bounded' (
    $thread.Contains('kExactNotificationLimit = 64'))
Check 'Large operation refreshes each directory once' ($thread.Contains('std::set<xpr::string> sChangedDirectories'))
Check 'Explorer reconciliation contract declared' ($explorerHeader.Contains('FileOperationReconcileItems'))
Check 'List deletion is descending and redraw-batched' ($explorer.Contains('sDeleteIndexes.rbegin()') -and $explorer.Contains('SetRedraw(XPR_FALSE)'))
Check 'Large target PIDL burst is bounded' ($explorer.Contains('aItems.size() <= 512'))
Check 'Status is updated after reconciliation' ($explorer.Contains('updateStatus();'))
Check 'Source snapshots retain directory type' ($threadHeader.Contains('mDirectory'))
Check 'Storage type query is volume-deduplicated first' ($adaptive.IndexOf('sQueriedVolumes.insert(sSourceVolume)') -lt $adaptive.IndexOf('queryStorageKind(aPlan.files[i].source, NULL)'))
Check 'Access-denied direct copy can fall back safely' ($adaptive.Contains('case ERROR_ACCESS_DENIED:'))
Check 'Modern shell tracks copy results and the actual destination item' (
    $modern.Contains('STDMETHODIMP PostCopyItem') -and
    $modern.Contains('record(aItem, aResult, aNewItem);'))
Check 'Modern shell tracks move results and the actual destination item' (
    $modern.Contains('STDMETHODIMP PostMoveItem') -and
    $modern.Contains('record(aItem, aResult, aNewItem);'))
Check 'Per-item shell failure prevents false success' ($modern.Contains('!sSink->failed().empty()') -and $modern.Contains('sResult = E_FAIL;'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object {
    '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name
}
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
