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

$option = Read-Text 'src\fxfile\option.cpp'
$dialog = Read-Text 'src\fxfile\cfg\cfg_func_refresh_dlg.cpp'
$dialogHeader = Read-Text 'src\fxfile\cfg\cfg_func_refresh_dlg.h'
$language = Read-Text 'src\fxfile\Languages\Korean.xml'
$resource = Read-Text 'src\fxfile\fxfile.rc'
$pane = Read-Text 'src\fxfile\explorer_pane.cpp'
$view = Read-Text 'src\fxfile\explorer_view.cpp'
$frame = Read-Text 'src\fxfile\main_frame.cpp'
$explorer = Read-Text 'src\fxfile\explorer_ctrl.cpp'

$setOption = Function-Body $explorer 'void ExplorerCtrl::setOption' 'void ExplorerCtrl::setImageList'
$legacyNotify = Function-Body $explorer 'LRESULT ExplorerCtrl::OnFileChangeNotify' 'LRESULT ExplorerCtrl::OnAdvFileChangeNotify'
$advancedNotify = Function-Body $explorer 'LRESULT ExplorerCtrl::OnAdvFileChangeNotify' 'void ExplorerCtrl::enumerateShcn'
$shellNotify = Function-Body $explorer 'LRESULT ExplorerCtrl::OnShellChangeNotify' 'xpr_bool_t ExplorerCtrl::beginShcn'
$finishNotify = Function-Body $explorer 'void ExplorerCtrl::endShcn' 'LRESULT ExplorerCtrl::OnFileChangeNotify'
$reconcile = Function-Body $explorer 'void ExplorerCtrl::reconcileFileOperationItems' 'LRESULT ExplorerCtrl::OnShellChangeNotify'
$runtimeTest = Read-Text 'tools\Test-Task070AutoRefreshSortRuntime.ps1'

# Windows PowerShell 5.1 treats a UTF-8 script without a BOM as the active ANSI
# code page.  Keep the Korean assertions encoding-independent so this contract
# means the same thing in the stock Windows host and in PowerShell 7.
$labelImmediateRefresh = -join [char[]]@(0xD30C,0xC77C,0x0020,0xBCC0,0xACBD,0x0020,0xC989,0xC2DC,0x0020,0xD654,0xBA74,0x0020,0xAC31,0xC2E0)
$labelRefreshThenSort = -join [char[]]@(0xD654,0xBA74,0x0020,0xAC31,0xC2E0,0x0020,0xD6C4,0x0020,0xC790,0xB3D9,0x0020,0xC815,0xB82C)
$obsoleteDoubleNegative = -join [char[]]@(0xC790,0xB3D9,0x0020,0xAC31,0xC2E0,0x0020,0xC0AC,0xC6A9,0x0020,0xC548,0xD568)

Check 'Refresh-sort option has one persisted configuration key' (
    $option.Contains('config.refresh.sort') -and
    ([regex]::Matches($option, 'config\.refresh\.sort').Count -eq 1))
Check 'No-refresh option has one persisted configuration key' (
    $option.Contains('config.refresh.no') -and
    ([regex]::Matches($option, 'config\.refresh\.no').Count -eq 1))
Check 'Settings dialog loads and applies both refresh values' (
    $dialog.Contains('SetCheck(!aConfig.mNoRefresh)') -and
    $dialog.Contains('SetCheck(aConfig.mRefreshSort)') -and
    $dialog.Contains('aConfig.mNoRefresh   = !') -and
    $dialog.Contains('aConfig.mRefreshSort =') -and
    $dialog.Contains('GetCheck()'))
Check 'Refresh UI uses independent positive labels instead of double negative wording' (
    $language.Contains('popup.cfg.body.function.bookmark.check.auto_refresh') -and
    $language.Contains($labelImmediateRefresh) -and
    $language.Contains($labelRefreshThenSort) -and
    -not $language.Contains($obsoleteDoubleNegative) -and
    $resource.Contains('Refresh file changes immediately'))
Check 'Auto-sort control is visibly dependent on immediate refresh' (
    $dialogHeader.Contains('OnAutoRefresh') -and
    $dialog.Contains('ON_BN_CLICKED(IDC_CFG_REFRESH_NO_REFRESH, OnAutoRefresh)') -and
    $dialog.Contains('IDC_CFG_REFRESH_AUTO_SORT)->EnableWindow(sAutoRefresh)'))
Check 'Main frame propagates changed settings to every explorer view' (
    $frame.Contains('for (i = 0; i < sViewCount; ++i)') -and
    $frame.Contains('sExplorerView->setChangedOption(aOption);'))
Check 'Explorer view propagates settings to all unique tab panes' (
    $view.Contains('for (i = 0; i < sTabCount; ++i)') -and
    $view.Contains('sTabPane->setChangedOption(aOption);'))
Check 'Explorer pane propagates refresh-sort to every list control' (
    $pane.Contains('mRefreshSort') -and
    $pane.Contains('mExplorerCtrlMap.begin()') -and
    $pane.Contains('setExplorerOption(sExplorerCtrlData->mExplorerCtrl, aOption);'))
Check 'Refresh policy becomes active without waiting for navigation' (
    $setOption.Contains('mOption.mNoRefresh   = aOption.mNoRefresh;') -and
    $setOption.Contains('mOption.mRefreshSort = aOption.mRefreshSort;'))
Check 'No-refresh policy suppresses shell and both directory watcher paths' (
    $frame.Contains('ShellChangeNotify::setNoRefresh(aOption.mConfig.mNoRefresh);') -and
    $legacyNotify.Contains('mOption.mNoRefresh == XPR_TRUE') -and
    $advancedNotify.Contains('mOption.mNoRefresh == XPR_TRUE'))
Check 'Windows shell notification path completes through common sort policy' (
    $shellNotify.Contains('endShcn(sEventId, sResult);'))
Check 'Legacy directory watcher completes through common sort policy' (
    $legacyNotify.Contains('const xpr_bool_t sResult = OnShcnUpdateDir') -and
    $legacyNotify.Contains('endShcn(SHCNE_UPDATEDIR, sResult);'))
Check 'Advanced watcher records changed state for create delete rename modify and directory refresh' (
    $advancedNotify.Contains('sResult = OnShcnCreateItem') -and
    $advancedNotify.Contains('sResult = OnShcnDeleteItem') -and
    $advancedNotify.Contains('sResult = OnShcnRenameItem') -and
    $advancedNotify.Contains('sResult = OnShcnUpdateDir'))
Check 'Advanced watcher completes through common sort policy' (
    $advancedNotify.Contains('endShcn(sEventId, sResult);'))
Check 'Common sort policy preserves in-place rename editing' (
    $finishNotify.Contains('mOption.mRefreshSort == XPR_TRUE') -and
    $finishNotify.Contains('GetEditControl()') -and
    $finishNotify.Contains('mRenameResorting') -and
    $finishNotify.Contains('resortItems();'))
Check 'File-operation UI reconciliation also honors refresh-sort' (
    $reconcile.Contains('mOption.mRefreshSort == XPR_TRUE') -and
    $reconcile.Contains('resortItems();'))
Check 'Runtime regression distinguishes refresh sorted refresh-only and no-refresh modes' (
    $runtimeTest.Contains("ValidateSet('Sorted', 'RefreshOnly', 'NoRefresh')") -and
    $runtimeTest.Contains("if (`$Mode -eq 'NoRefresh')") -and
    $runtimeTest.Contains("if (`$Mode -eq 'Sorted')"))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object {
    '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name
}
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
