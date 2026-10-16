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
    $end = if ([string]::IsNullOrEmpty($nextSignature)) {
        $text.Length
    } else {
        $text.IndexOf($nextSignature, $start + $signature.Length,
                      [StringComparison]::Ordinal)
    }
    if ($end -lt 0) { $end = $text.Length }
    $text.Substring($start, $end - $start)
}

$checks = [Collections.Generic.List[object]]::new()
function Check([string]$name, [bool]$passed) {
    $checks.Add([pscustomobject]@{ Name = $name; Passed = $passed })
}

$fileOp = Read-Text 'src\fxfile\file_op_thread.cpp'
$fileOpResult = Function-Body $fileOp 'void FileOpThread::reconcileOperationResult' 'LRESULT FileOpThread::OnPostEnd'
$shellIconH = Read-Text 'src\fxfile\shell_icon.h'
$shellIcon = Read-Text 'src\fxfile\shell_icon.cpp'
$shellColumn = Read-Text 'src\fxfile\shell_column_manager.cpp'
$shellNotify = Read-Text 'src\fxfile\shell_change_notify.cpp'
$thumbnail = Read-Text 'src\fxfile\thumbnail.cpp'
$thumbList = Read-Text 'src\fxfile\thumb_img_list.cpp'
$explorer = Read-Text 'src\fxfile\explorer_ctrl.cpp'
$systemInfo = Read-Text 'src\fxfile\SystemInfo.cpp'
$drive = Read-Text 'src\fxfile\drive_toolbar.cpp'
$advWatcher = Read-Text 'src\fxfile\adv_file_change_watcher.cpp'
$shellNotifyHeader = Read-Text 'src\fxfile\shell_change_notify.h'
$clipboard = Read-Text 'src\fxfile\clipboard.cpp'
$systemInfoHeader = Read-Text 'src\fxfile\SystemInfo.h'
$folderCtrl = Read-Text 'src\fxfile\folder_ctrl.cpp'
$searchCtrl = Read-Text 'src\fxfile\search_result_ctrl.cpp'
$addressBar = Read-Text 'src\fxfile\address_bar.cpp'
$bookmark = Read-Text 'src\fxfile\bookmark.cpp'
$folderSize = Read-Text 'src\fxfile\folder_size.cpp'
$syncDirs = Read-Text 'src\fxfile\sync_dirs.cpp'

Check 'Worker captures operation result before completion signal' (
    $fileOp.IndexOf('captureOperationResult();') -ge 0 -and
    $fileOp.IndexOf('captureOperationResult();') -lt $fileOp.LastIndexOf('::SetEvent(mStopEvent);'))
Check 'UI reconciliation contains no filesystem probes' (
    -not $fileOpResult.Contains('GetFileAttributes'))
Check 'Exact UI and shell event fan-out is capped at 64' (
    $fileOp.Contains('kExactUiLimit = 64') -and
    $fileOp.Contains('kExactNotificationLimit = 64'))
Check 'Cross-apartment raw shell folder removed from async icon payload' (
    -not $shellIconH.Contains('LPSHELLFOLDER mShellFolder'))
Check 'Async icon worker rebinds absolute PIDL in its COM apartment' (
    $shellIcon.Contains('SHBindToParent') -and
    $shellIcon.Contains('CoEnableCallCancellation'))
Check 'Shell column worker is woken before join and supports COM cancellation' (
    $shellColumn.IndexOf('mThread.stop();') -lt $shellColumn.IndexOf('::SetEvent(mEvent)') -and
    $shellColumn.IndexOf('::SetEvent(mEvent)') -lt $shellColumn.IndexOf('mThread.join();') -and
    $shellColumn.Contains('CoEnableCallCancellation'))
Check 'Shell column requests are bounded and duplicate-suppressed' (
    $shellColumn.Contains('kMaxAsyncColumnQueue = 512') -and
    $shellColumn.Contains('mInFlightAsyncInfo'))
Check 'Shell change queue is bounded and TerminateThread removed' (
    $shellNotify.Contains('kMaxQueueSize = 512') -and
    -not $shellNotify.Contains('::TerminateThread('))
Check 'Thumbnail decoder queue is deduplicated and bounded' (
    $thumbnail.Contains('kMaxPendingThumbnails = 256') -and
    $thumbnail.Contains('mInFlightThumbItem'))
Check 'Thumbnail result ownership is cancelled and drained on pane destroy' (
    $thumbnail.Contains('cancelAsyncImage') -and
    $explorer.Contains('cancelAsyncImage(m_hWnd, WM_THUMBNAIL_PROC)'))
Check 'Runtime thumbnail image list has a hard record cap and Add failure check' (
    $thumbList.Contains('mThumbDeque.size() >= kMaxThumbnailRecords') -and
    $thumbList.Contains('if (sImageIndex < 0)'))
Check 'File-name query no longer force-terminates a CRT worker' (
    -not $systemInfo.Contains('TerminateThread( hThread') -and
    $systemInfo.Contains('CancelSynchronousIo( hThread') -and
    $systemInfo.Contains('ReleaseFileNameThreadParam'))
Check 'Disconnected drives are not probed by shell icon worker' (
    $drive.Contains('SHGFI_USEFILEATTRIBUTES') -and
    $drive.Contains('sDriveType != DRIVE_REMOTE'))
Check 'Automatic column content measurement is strictly sampled' (
    $explorer.Contains('kMaximumSamples = 64'))
Check 'Advanced watcher cancellation and record bounds are enforced' (
    $advWatcher.Contains('::CancelIoEx') -and
    $advWatcher.Contains('kCancelCompletionTimeout = 2000') -and
    $advWatcher.Contains('releaseRetiredOverlapped') -and
    -not $advWatcher.Contains('::GetOverlappedResult') -and
    $advWatcher.Contains('kMaxRawNotifyQueue = 512'))
Check 'Advanced watcher filename byte count is not multiplied twice' (
    -not $advWatcher.Contains('FileNameLength * sizeof(WCHAR)'))
Check 'Shell notification identifies and cancels owner watches' (
    $shellNotifyHeader.Contains('mWatchId') -and
    $shellNotify.Contains('mCancelledWatchSet') -and
    $shellNotify.Contains('mInFlightShcn'))
Check 'Shell notification startup failures are closed and registration failure returns zero' (
    $shellNotify.Contains('if (XPR_IS_NULL(mStopEvent))') -and
    $shellNotify.Contains('if (XPR_IS_NULL(sNotifyItem->mNotify))'))
Check 'Shell notification failure uses event-aware payload release' (
    $shellNotify.Contains('sShcn->Free();') -and
    -not $shellNotify.Contains('COM_FREE(sShcn->mPidl1);'))
Check 'Thumbnail in-flight pointer clears before request deletion' (
    $thumbnail.Contains('if (mInFlightThumbItem == sThumbItem)') -and
    $thumbnail.IndexOf('mInFlightThumbItem = XPR_NULL;',
                       $thumbnail.IndexOf('if (mInFlightThumbItem == sThumbItem)')) -lt
        $thumbnail.LastIndexOf('XPR_SAFE_DELETE(sThumbItem);'))
Check 'Async worker startup uses native handle instead of racy state' (
    $shellIcon.Contains('mThread.getThreadHandle().mHandle') -and
    $shellColumn.Contains('mThread.getThreadHandle().mHandle'))
Check 'Native filename query uses pointer-sized status and byte-counted UTF16' (
    $systemInfo.Contains('ULONG_PTR Information') -and
    $systemInfo.Contains('FileNameLength') -and
    $systemInfo.Contains('maximumNameBytes') -and
    $systemInfo.Contains('DUPLICATE_SAME_ACCESS'))
Check 'Native filename result is read only on WAIT_OBJECT_0' (
    $systemInfo.Contains('waitResult != WAIT_OBJECT_0'))
Check 'Clipboard storage mediums are released by ownership contract' (
    $clipboard.Contains('::ReleaseStgMedium'))
Check 'Destroyed controls drain owned icon and shell payloads' (
    $folderCtrl.Contains('WM_SHELL_ASYNC_ICON, PM_REMOVE') -and
    $folderCtrl.Contains('WM_SHELL_CHANGE_NOTIFY, PM_REMOVE') -and
    $searchCtrl.Contains('WM_SHELL_ASYNC_ICON, PM_REMOVE') -and
    $searchCtrl.Contains('WM_SHELL_CHANGE_NOTIFY, PM_REMOVE'))
Check 'Address bar unregisters both shell watches before destruction' (
    $addressBar.Contains('unregisterWatch(mShcnDesktopId)') -and
    $addressBar.Contains('unregisterWatch(mShcnComputerId)'))
Check 'Bookmark manager drains posted icon ownership after stop' (
    $bookmark.Contains('WM_SHELL_ASYNC_ICON, PM_REMOVE'))
Check 'No worker is force-terminated while owning heap or locks' (
    -not $folderSize.Contains('TerminateThread(') -and
    -not $syncDirs.Contains('TerminateThread(') -and
    $folderSize.Contains('CancelSynchronousIo') -and
    $syncDirs.Contains('CancelSynchronousIo'))
Check 'Folder size recursion is cooperative and skips reparse cycles' (
    $folderSize.Contains('FILE_ATTRIBUTE_REPARSE_POINT') -and
    $folderSize.Contains('aStopEvent') -and
    $folderSize.Contains('nFileSizeHigh << 32'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object {
    '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name
}
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
