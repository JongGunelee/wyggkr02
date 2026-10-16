[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
function Read([string]$p) { [IO.File]::ReadAllText((Join-Path $root $p)) }
$checks = [Collections.Generic.List[object]]::new()
function Check([string]$name, [bool]$ok) { $checks.Add([pscustomobject]@{Name=$name; Passed=$ok}) }
$policy = Read 'src\fxfile\drag_file_operation.h'
$explorer = Read 'src\fxfile\explorer_ctrl.cpp'
$rename = Read 'src\fxfile\rename_helper.cpp'
$dialog = Read 'src\fxfile\cmd\single_rename_dlg.cpp'
foreach($file in @('explorer_ctrl.cpp','folder_ctrl.cpp','drive_toolbar.cpp','bookmark_toolbar.cpp')) {
    $body = Read "src\fxfile\$file"
    Check "Shared drag policy: $file" ($body.Contains('getDragFileOperation(aKeyState,') -and $body.Contains('#include "drag_file_operation.h"'))
}
Check 'Default copy; explicit move/link and modifier precedence preserved' ($policy.Contains('return DROPEFFECT_COPY;') -and $policy.Contains('MK_SHIFT') -and $policy.Contains('DRAG_FILE_OP_MOVE') -and $policy.Contains('DRAG_FILE_OP_LINK'))
Check 'Remove synchronous source CIDA/name lookup from DragOver default' (-not $explorer.Contains('sDropEffect = sCopy ? DROPEFFECT_COPY : DROPEFFECT_MOVE;'))
Check 'Rename owns COM and relative PIDL across modal refresh' ($rename.Contains('aShellFolder->AddRef()') -and $rename.Contains('Pidl::clone(aPidl)') -and $rename.Contains('COM_RELEASE(mShellFolder)') -and $rename.Contains('COM_FREE(mPidl)'))
Check 'F2 opening uses cached source path and cached icon only' ($rename.Contains('mSourcePath = aCachedPath') -and $rename.Contains('return mCachedIconIndex;') -and -not $rename.Contains('GetItemIconIndex(mShellFolder, mPidl)'))
Check 'Unknown icon skips image extraction; no new shell provider query' ($dialog.Contains('(sIconIndex >= 0) ?') -and -not $dialog.Contains('SHGetFileInfo'))
Check 'Rename notification does not wait for third-party shell listeners' ($rename.Contains('sFlags |= SHCNF_FLUSHNOWAIT;') -and -not $rename.Contains('sFlags |= SHCNF_FLUSH;'))
Check 'Shell notification uses desktop-relative PIDLs and skips invalid identifiers' ($rename.Contains('const void *sItem1 = mFullPidl;') -and $rename.Contains('const void *sItem2 = mNewFullPidl;') -and $rename.Contains('XPR_IS_NOT_NULL(sItem1) && XPR_IS_NOT_NULL(sItem2)'))
Check 'Missing item attributes do not masquerade as protected attributes' ($rename.Contains('sFileAttributes != INVALID_FILE_ATTRIBUTES'))
Check 'Scrollbar tracking recomposes without erase or column reset' ($explorer.Contains('ON_WM_HSCROLL()') -and $explorer.Contains('void ExplorerCtrl::OnHScroll') -and $explorer.Contains('super::OnHScroll(aSbCode, aPos, aScrollBar);'))
Check 'Cell clipping restores DC and preserves live mapped header origin' ($explorer.Contains('sVisibleCellRect.right, sVisibleCellRect.bottom);') -and $explorer.Contains('RestoreDC(aNmLvCustomDraw->nmcd.hdc, sCellDc)') -and $explorer.Contains('sHeaderRect.OffsetRect(sHeaderOrigin.x, 0)'))
$failed = @($checks |Where-Object {-not $_.Passed})
$checks|ForEach-Object { '{0} {1}' -f $(if($_.Passed){'PASS'}else{'FAIL'}),$_.Name }
"SUMMARY PASS=$($checks.Count-$failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if($failed.Count){exit 1}
