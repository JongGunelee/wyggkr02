[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$root=Join-Path $PSScriptRoot '../src/fxfile'
function Read-Code($name){[IO.File]::ReadAllText((Join-Path $root $name))}
$confirm=Read-Code 'file_operation_confirmation.cpp'
$thread=Read-Code 'file_op_thread.cpp'
$clipboard=Read-Code 'clipboard.cpp'
$context=Read-Code 'context_menu.cpp'
$drop=Read-Code 'gui/DropTarget.cpp'
$source=Read-Code 'gui/DropSource.cpp'
$start=$thread.Substring($thread.IndexOf('xpr_bool_t FileOpThread::start(SHFILEOPSTRUCT'),$thread.IndexOf('unsigned __stdcall FileOpThread::FileOpProc')-$thread.IndexOf('xpr_bool_t FileOpThread::start(SHFILEOPSTRUCT'))
$checks=@(
 @('Worker consent precedes operation acceptance', ($start.IndexOf('FileOperationConfirmation::confirm') -lt $start.IndexOf('mAcceptedTick ='))),
 @('Worker consent precedes event/thread creation', ($start.IndexOf('FileOperationConfirmation::confirm') -lt $start.IndexOf('::CreateEvent'))),
 @('Cancellation closes owned operation without starting worker', ($start.Contains('PostMessage(WM_CLOSE);') -and $start.Contains('return XPR_FALSE;'))),
 @('Consent has no saved option bypass', -not($confirm -match 'gOpt|mConfig|FOF_NOCONFIRMATION')),
 @('Cancel is default and owns initial focus', ($confirm.Contains('SetDefID(IDCANCEL)') -and $confirm.Contains('GetDlgItem(IDCANCEL)->SetFocus()'))),
 @('Only affirmative dialog result accepts', $confirm.Contains('return dialog.DoModal() == IDOK;')),
 @('Enter explicitly cancels without default-button dispatch', ($confirm.Contains('message->wParam == VK_RETURN') -and $confirm.Contains('OnCancel();'))),
 @('Cut visual state updated only after publication', (Read-Code 'cmd/cmd_clipboard.cpp').Contains('if (SetClipboard(sDataObject, DROPEFFECT_MOVE, sClipFormat))')),
 @('Complete source list retained in scrollable details', ($confirm.Contains('i < sources.size()') -and $confirm.Contains('edit->LimitText(0)'))),
 @('Permanent delete distinguished from recycle', ($confirm.Contains('FOF_ALLOWUNDO') -and $confirm.Contains('영구 삭제'))),
 @('Duplicate names disclosed', $confirm.Contains('FOF_RENAMEONCOLLISION')),
 @('Cut confirmed before clipboard publication', ($clipboard.IndexOf('confirmCut(') -lt $clipboard.IndexOf('sDataSource->SetClipboard()'))),
 @('Paste no longer clears clipboard before execution', -not $clipboard.Contains('::OleSetClipboard(XPR_NULL)')),
 @('Paste observes acceptance result', $clipboard.Contains('sResult = sFileOpThread->start(')),
 @('Moved clipboard cleared only after per-item full success', ($thread.Contains('sSucceededCount == mSourceSnapshots.size()') -and $thread.Contains('!sFailedCount && !sCancelledCount && !sUnprocessedCount'))),
 @('Newer clipboard protected at completion', $thread.Contains('::GetClipboardSequenceNumber() == mClipboardMoveSequence')),
 @('Repeated pending cut sequence blocked and released', ($thread.Contains('mActiveClipboardMoves.count(') -and $thread.Contains('mActiveClipboardMoves.erase('))),
 @('Shell numeric and verb invocation both guarded', ([regex]::Matches($context,'FileOperationConfirmation::confirmShellVerb').Count -eq 2)),
 @('Context menu owns COM and PIDL snapshots', ($context.Contains('mShellFolder->AddRef()') -and $context.Contains('fxfile::base::Pidl::clone(aPidls[i])') -and $context.Contains('delete[] mPidls'))),
 @('Shell paste refuses changed clipboard', $confirm.Contains('::GetClipboardSequenceNumber() != sequence')),
 @('Right drag resolves concrete operation first', ($drop.Contains('chooseDropEffect(') -and $drop.Contains('mRightDrag = XPR_FALSE;'))),
 @('External drop can cancel before OLE commits', ($source.Contains('FileOperationConfirmation::confirm(') -and $source.Contains('DRAGDROP_S_CANCEL'))),
 @('External consent binds original drop window and point', ($source.Contains('::WindowFromPoint(point) == target') -and $source.Contains('::SetCursorPos(point.x, point.y)'))),
 @('Refusal notice uses readable owned dialog', ($confirm.Contains('mNotice ? L"닫기"') -and -not $confirm.Contains('::MessageBoxW('))),
 @('Confirmation uses readable twelve point Korean font', (Read-Code 'fxfile.rc').Contains('FONT 12, "Malgun Gothic", 400, 0, 0x81')),
 @('Readable window clamped to monitor and font metrics', ($confirm.Contains('monitor.rcWork') -and $confirm.Contains('dc.GetTextMetrics(&metrics)'))),
 @('CIDA offsets and PIDL boundaries validated', ($confirm.Contains('header <= bytes') -and $confirm.Contains('validPidl(raw, bytes, ida->aoffset[i])'))),
 @('File collection bypass route guarded', (Read-Code 'file_scrap.cpp').Contains('FileOperationConfirmation::confirm(&sShFileOpStruct)')),
 @('Archive paste captured and confirmed before extraction', (Read-Code 'cmd/cmd_clipboard.cpp').Contains('extractItems(sConfirmedArchive, sConfirmedItems')),
 @('Real files never take unchecked content/text import paths', ([regex]::Matches($clipboard,'if \(HasFileSelectionData\(aOleDataObject\)\) return XPR_FALSE;').Count -eq 6)),
 @('Clipboard sequence frozen before attaching provider', $clipboard.Contains('const DWORD sSequenceBeforeAttach = ::GetClipboardSequenceNumber();')),
 @('Virtual content cancellation consumes request without fallback', $clipboard.Contains('This representation was handled, including cancellation.')),
 @('Virtual file descriptor names and sizes bounded', ($clipboard.Contains('sDeclaredCount <= (sBytes - sHeader) / sUnit') -and $clipboard.Contains('sName.find_first_of')))
)
foreach($file in @('explorer_ctrl.cpp','folder_ctrl.cpp','bookmark_toolbar.cpp')){
 $code=Read-Code $file
 $checks+=,@("Shell drop guarded in $file",($code.Contains('FileOperationConfirmation::confirmDrop(') -and $code.Contains('sDropTarget->DragLeave();')))
 $checks+=,@("Shell drop owns separate data reference in $file",($code.Contains('->GetIDataObject(XPR_TRUE)') -and -not $code.Contains('->Detach()')))
}
$fail=0
foreach($check in $checks){if($check[1]){"PASS $($check[0])"}else{"FAIL $($check[0])";$fail++}}
"SUMMARY PASS=$($checks.Count-$fail) FAIL=$fail"
if($fail){exit 1}
