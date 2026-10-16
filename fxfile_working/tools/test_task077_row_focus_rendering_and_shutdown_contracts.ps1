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

$explorer = Read-Text 'src\fxfile\explorer_ctrl.cpp'
$explorerView = Read-Text 'src\fxfile\explorer_view.cpp'
$folderHeader = Read-Text 'src\fxfile\folder_ctrl.h'
$folder = Read-Text 'src\fxfile\folder_ctrl.cpp'
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$resetDraw = Function-Body $explorer 'void ExplorerCtrl::resetCustomDrawColors' 'void ExplorerCtrl::applyCustomDrawFiltering'
$reportSelectionDraw = Function-Body $explorer 'void ExplorerCtrl::applyReportSelectionDrawState' 'void ExplorerCtrl::drawParentFolderIcon'
$onDestroy = Function-Body $folder 'void FolderCtrl::OnDestroy(void)' 'void FolderCtrl::setOption'
$onFileChange = Function-Body $folder 'LRESULT FolderCtrl::OnFileChangeNotify' 'LRESULT FolderCtrl::OnShellChangeNotify'
$onShellChange = Function-Body $folder 'LRESULT FolderCtrl::OnShellChangeNotify' 'xpr_bool_t FolderCtrl::canProcessShellChange'
$updateItem = Function-Body $folder 'xpr_bool_t FolderCtrl::updateShcnTvItemData' 'xpr_bool_t FolderCtrl::OnShcnDeleteItem'

$textBase = $resetDraw.IndexOf('aNmLvCustomDraw->clrText   = GetTextColor();', [StringComparison]::Ordinal)
$backgroundBase = $resetDraw.IndexOf('aNmLvCustomDraw->clrTextBk = GetTextBkColor();', [StringComparison]::Ordinal)
$filtering = $customDraw.IndexOf('applyCustomDrawFiltering(sNmLvCustomDraw);', [StringComparison]::Ordinal)
Check 'Each item custom draw resets text and background before filtering or row-focus overrides' (
    $textBase -ge 0 -and $backgroundBase -ge 0 -and
    $customDraw.IndexOf('resetCustomDrawColors(sNmLvCustomDraw);', [StringComparison]::Ordinal) -lt $filtering -and
    $reportSelectionDraw.Contains('applyRowFocusDrawState(aNmLvCustomDraw);'))
Check 'A list background image alone is the only path that requests transparent item backgrounds' (
    $resetDraw.Contains('GetBkImage(&sLvBkImage) == XPR_TRUE && sImage[0] != XPR_STRING_LITERAL(''\0'')') -and
    $resetDraw.Contains('aNmLvCustomDraw->clrTextBk = CLR_NONE;'))
Check 'The configured color queries selected plus item-focused state and covers full-row plus first-cell modes after themed drawing' (
    $customDraw.Contains('XPR_IS_TRUE(mOption.mFullRowSelect)') -and
    $explorer.Contains('mFocusedItemIndex') -and
    $reportSelectionDraw.Contains('if (!XPR_TEST_BITS(sNativeState, LVIS_SELECTED))') -and
    -not $reportSelectionDraw.Contains('isFocusedSelectedItem(sItemIndex)') -and
    $reportSelectionDraw.Contains('applyRowFocusDrawState(aNmLvCustomDraw);') -and
    $customDraw.Contains('CDRF_NEWFONT') -and
    $customDraw.Contains('CDRF_NOTIFYSUBITEMDRAW'))
Check 'An unavailable saved drive lock falls back without rewriting the user preference' (
    $explorerView.Contains('isAvailableStartupPath') -and
    $explorerView.Contains('::GetDriveType(sRoot) == DRIVE_NO_ROOT_DIR') -and
    $explorerView.Contains('XPR_IS_TRUE(isAvailableStartupPath(gOpt->mMain.mLockedViewPath[mViewIndex]))') -and
    -not $explorerView.Contains('mViewPathLocked = XPR_FALSE'))
Check 'Folder controls mark shutdown before unregistering asynchronous shell notifications' (
    $folderHeader.Contains('mDestroying') -and
    $onDestroy.IndexOf('mDestroying = XPR_TRUE;', [StringComparison]::Ordinal) -lt $onDestroy.IndexOf('unregisterWatch', [StringComparison]::Ordinal))
Check 'Queued file and shell notifications are rejected safely after shutdown begins' (
    $folderHeader.Contains('canProcessShellChange(void) const') -and
    $onFileChange.Contains('XPR_IS_FALSE(canProcessShellChange())') -and
    $onShellChange.Contains('sShcn->Free();') -and
    $onShellChange.Contains('XPR_SAFE_DELETE(sShcn);'))
Check 'The crash-report shell-update path validates window lifetime and old tree item data before use' (
    $updateItem.Contains('if (XPR_IS_FALSE(canProcessShellChange()))') -and
    $updateItem.Contains('XPR_IS_NULL(sTvItemData2)') -and
    $updateItem.Contains('Ownership of aTvItemData now belongs to the tree control.'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
