$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$dialog = Get-Content -LiteralPath (Join-Path $root 'src\fxfile\cmd\sel_name_dlg.cpp') -Raw -Encoding UTF8
$explorer = Get-Content -LiteralPath (Join-Path $root 'src\fxfile\explorer_ctrl.cpp') -Raw -Encoding UTF8
function Check([bool]$condition, [string]$message) {
    if (-not $condition) { throw $message }
    Write-Host "PASS $message"
}
Check ($dialog -match 'sOnlySel\s*=\s*mSelect;') 'Select by name defaults to clear other selections; deselect dialog does not.'
Check ($dialog -match 'SetCheck\(sOnlySel\)') 'Default is visible in the checkbox when the dialog opens.'
Check ($dialog -match 'mOnlySel\s*=.*GetCheck\(') 'The user can still uncheck the option for one operation.'
Check ($explorer -match '(?s)SHGetPathFromIDList\(aFullPidl, sAccessPath\).*?CreateFile\(sAccessPath, FILE_LIST_DIRECTORY.*?ERROR_ACCESS_DENIED.*?return XPR_FALSE;.*?getSimplePidl') 'Denied local folder is rejected before replacing pane state.'
Check ($explorer -match '(?s)if \(XPR_IS_TRUE\(sSucceeded\)\)\s*\{\s*if \(XPR_IS_TRUE\(mDirectoryEnumerationDirty\)\)\s*scheduleDirectoryRefresh\(\);') 'Failed enumeration cannot schedule a dirty watcher retry.'
Check ($explorer -match '(?s)if \(XPR_IS_FALSE\(sSucceeded\) && XPR_IS_FALSE\(mDestroying\)\).*?unregisterWatch\(mAdvWatchId\).*?unregisterWatch\(mWatchId\).*?unregisterWatch\(mShcnId\).*?KillTimer\(TM_ID_NOTIFY_RECONCILE\).*?MB_RETRYCANCEL') 'All failed-folder watches and recovery timer are retired before Retry/Cancel.'
