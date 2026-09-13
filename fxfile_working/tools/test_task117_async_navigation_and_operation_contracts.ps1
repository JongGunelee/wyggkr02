[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$sourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\src\fxfile'))
$explorer = Get-Content -LiteralPath (Join-Path $sourceRoot 'explorer_ctrl.cpp') -Raw
$explorerHeader = Get-Content -LiteralPath (Join-Path $sourceRoot 'explorer_ctrl.h') -Raw
$worker = Get-Content -LiteralPath (Join-Path $sourceRoot 'directory_enumeration_worker.cpp') -Raw
$workerHeader = Get-Content -LiteralPath (Join-Path $sourceRoot 'directory_enumeration_worker.h') -Raw
$adaptive = Get-Content -LiteralPath (Join-Path $sourceRoot 'adaptive_file_operation.cpp') -Raw
$modern = Get-Content -LiteralPath (Join-Path $sourceRoot 'modern_shell_file_operation.cpp') -Raw
$modernHeader = Get-Content -LiteralPath (Join-Path $sourceRoot 'modern_shell_file_operation.h') -Raw
$fileOp = Get-Content -LiteralPath (Join-Path $sourceRoot 'file_op_thread.cpp') -Raw
$rename = Get-Content -LiteralPath (Join-Path $sourceRoot 'rename_helper.cpp') -Raw
$mainFrame = Get-Content -LiteralPath (Join-Path $sourceRoot 'main_frame.cpp') -Raw
$probeBuild = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Build-Task117OperationProbes.ps1') -Raw
$buildStorage = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Assert-BuildStorage.ps1') -Raw
$buildDeploy = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Build-Deploy-Verify.ps1') -Raw
$buildEnvironment = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Test-BuildEnvironment.ps1') -Raw
$keyboardProbe = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Test-Task117KeyboardStartup.ps1') -Raw

$checks = [ordered]@{
    'filesystem folder enumeration runs on an STA worker' =
        $worker.Contains('CoInitializeEx(XPR_NULL, COINIT_APARTMENTTHREADED)') -and
        $explorer.Contains('startDirectoryEnumeration(')
    'Shell interfaces never cross the worker boundary' =
        $worker.Contains('mFolderFullPidl') -and
        -not $workerHeader.Contains('LPSHELLFOLDER')
    'bounded UI batches have backpressure and cancellation' =
        $worker.Contains('kEnumerationBatchSize = 128') -and
        $worker.Contains('CreateSemaphore(XPR_NULL, 4, 4') -and
        $worker.Contains('WaitForMultipleObjects')
    'posted enumeration batches have shutdown-safe ownership' =
        $worker.Contains('registerBatch(aBatch)') -and
        $worker.Contains('claimBatch') -and
        $worker.Contains('cleanupOwnerBatches') -and
        $explorer.Contains('DirectoryEnumerationWorker::claimBatch') -and
        $explorer.Contains('DirectoryEnumerationWorker::cleanupOwnerBatches')
    'worker resolves filesystem metadata and names before UI publication' =
        $worker.Contains('SHGetDataFromIDListW') -and
        $worker.Contains('sFindData.cFileName') -and
        $worker.Contains('mHasKnownMetadata') -and
        $explorer.Contains('sIt->mName.c_str()') -and
        $explorer.Contains('sIt->mHasKnownMetadata')
    'ListView insertion failure rolls back row ownership and name index' =
        $explorer.Contains('const xpr_sint_t sInsertedIndex = InsertItem') -and
        $explorer.Contains('eraseNameHash(sLvItemData)') -and
        $explorer.Contains('COM_RELEASE(sLvItemData->mShellFolder)') -and
        $explorer.Contains('sLvItemData->mPidl = XPR_NULL')
    'stale batches are rejected by generation and owner token' =
        $explorer.Contains('sBatch->mGeneration != mDirectoryEnumerationGeneration') -and
        $explorer.Contains('sBatch->mOwnerToken != mDirectoryEnumerationOwnerToken')
    'async path is restricted to ordinary local fixed directories' =
        $explorer.Contains('isAsyncLocalDirectoryEligible') -and
        $explorer.Contains('FILE_ATTRIBUTE_REPARSE_POINT') -and
        $explorer.Contains('GetDriveType(sRoot) == DRIVE_FIXED')
    'changes during enumeration are reconciled after publish' =
        $explorer.Contains('mDirectoryEnumerationDirty = XPR_TRUE') -and
        $explorer.Contains('if (XPR_IS_TRUE(mDirectoryEnumerationDirty))')
    'local full reconciliation stays asynchronous and restores view state' =
        $explorer.Contains('sResult = explore(sFullPidl, XPR_FALSE)') -and
        $explorer.Contains('captureRefreshViewState()') -and
        $explorer.Contains('restoreRefreshViewState()') -and
        $explorer.Contains('mRefreshHorizontalScroll') -and
        $explorer.Contains('mRefreshSkipSort')
    'navigation timing separates enumeration first batch and final sort' =
        $explorer.Contains('first_batch_ms=') -and
        $explorer.Contains('finalize_sort_ms=') -and
        $explorer.Contains('icon=deferred')
    'refresh recovery is generation-bound and has bounded retry' =
        $explorer.Contains('mDeferredDirectoryRefreshGeneration') -and
        $explorer.Contains('kDirectoryRefreshRetryDelay[] = {250, 500, 1000, 2000}') -and
        $explorer.Contains('reportDirectoryRefreshFailure')
    'planning appears before writes and remains cancellable' =
        $adaptive.Contains('class PlanProgress') -and
        $adaptive.Contains('PROGDLG_MARQUEEPROGRESS') -and
        $adaptive.Contains('HasUserCancelled') -and
        $adaptive.Contains('if (!sProgress.pulse())')
    'planning inventory has a safe memory bound and Shell fallback' =
        $adaptive.Contains('kMaximumInMemoryPlanItems = 100000') -and
        $adaptive.Contains('ERROR_NOT_ENOUGH_MEMORY') -and
        $fileOp.Contains('ModernShellFileOperation::tryExecute')
    'adaptive concurrency examines all distinct source volumes and target' =
        $adaptive.Contains('sQueriedVolumes') -and
        $adaptive.Contains('queryStorageKind(aPlan.files[i].source') -and
        $adaptive.Contains('sAggregateKind == StorageUnknown')
    'result contract records per-item outcome engine reason and phase timing' =
        $fileOp.Contains('OutcomeCancelled') -and
        $fileOp.Contains('OutcomeUnprocessed') -and
        $fileOp.Contains('sSnapshot.mError == ERROR_CANCELLED') -and
        $fileOp.Contains('FxFile.FileOperationItem id=') -and
        $fileOp.Contains('FxFile.FileOperation id=') -and
        $fileOp.Contains('verification=existence-and-source-state')
    'failed copy cannot be reported successful from a leftover target alone' =
        $fileOp.Contains('XPR_IS_TRUE(mOperationSucceeded)') -and
        $fileOp.Contains('prefers a conservative unprocessed/fail') -and
        $fileOp.Contains('mExistedBefore')
    'IFileOperation reports actual per-item result and destination' =
        $modernHeader.Contains('struct ItemResult') -and
        $modern.Contains('PostCopyItem') -and
        $modern.Contains('PostMoveItem') -and
        $modern.Contains('QueryInterface(IID_IUnknown') -and
        $modern.Contains('std::map<IUnknown *, std::wstring>') -and
        $modern.Contains('sResult.mTargetPath') -and
        $modern.Contains('aExecutionInfo->mItems = sSink->results()') -and
        $fileOp.Contains('sModernItem->mResult')
    'failed or cancelled items cannot emit an exact successful create notification' =
        $fileOp.Contains('sIt->mOutcome == ResultSnapshot::OutcomeSucceeded') -and
        $fileOp.Contains('It must not') -and
        $fileOp.Contains('SHCNE_UPDATEDIR')
    'large operation diagnostics are bounded off the UI hot path' =
        $fileOp.Contains('kDetailedTraceLimit = 64') -and
        $fileOp.Contains('detailed_records_omitted=')
    'single rename handles empty and unchanged labels explicitly' =
        $rename.Contains('return ResultEmptiedName;') -and
        $rename.Contains('An unchanged label is a successful no-op')
    'completed rename reconciles source and target in all panes' =
        $explorer.Contains('aOperation == FO_RENAME') -and
        $fileOp.Contains('mShFileOpStruct->wFunc == FO_RENAME')
    'startup commits keyboard focus to the final active list without mouse' =
        $mainFrame.Contains('WM_DEFERRED_STARTUP_KEYBOARD_FOCUS') -and
        $mainFrame.Contains('mStartupKeyboardFocusReady = XPR_TRUE') -and
        $mainFrame.Contains('requestStartupKeyboardFocus()') -and
        $mainFrame.Contains('sExplorerCtrl->SetFocus()') -and
        $mainFrame.Contains('GetForegroundWindow()') -and
        $explorer.Contains('gFrame->requestStartupKeyboardFocus()') -and
        -not $mainFrame.Contains("mSplitter.showPane(XPR_TRUE);`r`n    recalcLayout();`r`n    mSplitter.setFocus();") -and
        $mainFrame.Contains('--sViewIndex;') -and
        $mainFrame.Contains('getExplorerCtrl(sViewIndex)')
    'standalone x64/x32 probes use product-compatible UTF-8 and Win32 linkage' =
        ([regex]::Matches($probeBuild, '/utf-8').Count -ge 4) -and
        ([regex]::Matches($probeBuild, 'User32\.lib').Count -ge 2)
    'low-system-drive exception enforces the documented one-GiB emergency floor' =
        $buildStorage.Contains('$systemDrive.FreeBytes -ge 1GB') -and
        $buildDeploy.Contains('$script:SystemEmergencyFloorBytes = [int64](1GB)') -and
        $buildEnvironment.Contains('$systemEmergencyFloorBytes = [int64](1GB)') -and
        -not $buildStorage.Contains('$systemDrive.FreeBytes -ge 100MB') -and
        -not $buildDeploy.Contains('$script:SystemEmergencyFloorBytes = [int64]100MB') -and
        -not $buildEnvironment.Contains('$systemEmergencyFloorBytes = [int64]100MB')
    'keyboard runtime probe has no optional System.Drawing assembly dependency' =
        $keyboardProbe.Contains('public struct RECT') -and
        -not $keyboardProbe.Contains('System.Drawing.Rectangle') -and
        -not $keyboardProbe.Contains('-ReferencedAssemblies System.Drawing')
}

$failed = @($checks.GetEnumerator() | Where-Object { -not $_.Value })
foreach ($entry in $checks.GetEnumerator()) {
    '{0}: {1}' -f ($(if ($entry.Value) { 'PASS' } else { 'FAIL' })), $entry.Key
}
if ($failed.Count -gt 0) {
    throw ('Task117 contracts failed: ' + (($failed | ForEach-Object Key) -join '; '))
}
'Task117 contracts: PASS ({0}/{0})' -f $checks.Count
