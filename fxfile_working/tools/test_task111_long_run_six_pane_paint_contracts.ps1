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
    $text.Substring($start, $end - $start)
}

$explorer = Read-Text 'src\fxfile\explorer_ctrl.cpp'
$explorerHeader = Read-Text 'src\fxfile\explorer_ctrl.h'
$shellIcon = Read-Text 'src\fxfile\shell_icon.cpp'
$shellIconHeader = Read-Text 'src\fxfile\shell_icon.h'
$definition = Read-Text 'src\fxfile\fxfile_def.h'
$explorerView = Read-Text 'src\fxfile\explorer_view.cpp'
$pathBar = Read-Text 'src\fxfile\path_bar.cpp'
$longRunHarness = Read-Text 'tools\Test-Task111SixPaneLongRun.ps1'

$drawState = Function-Body $explorer 'void ExplorerCtrl::applyReportSelectionDrawState' 'void ExplorerCtrl::drawFinalReportSelection'
$selectionFill = Function-Body $explorer 'void ExplorerCtrl::drawFinalReportSelection' 'void ExplorerCtrl::drawParentFolderIcon'
$parentIcon = Function-Body $explorer 'void ExplorerCtrl::drawParentFolderIcon' 'void ExplorerCtrl::redrawFocusItemChange'
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$asyncHandler = Function-Body $explorer 'LRESULT ExplorerCtrl::OnShellAsyncIcon' 'LRESULT ExplorerCtrl::OnShellColumnProc'
$clearIcon = Function-Body $shellIcon 'void ShellIcon::clear' 'void ShellIcon::stopThread'
$queueIcon = Function-Body $shellIcon 'xpr_bool_t ShellIcon::getAsyncIcon' 'xpr_sint_t ShellIcon::runThread'
$workerIcon = Function-Body $shellIcon 'xpr_sint_t ShellIcon::runThread' 'HICON ShellIcon::getIcon'
$filtering = Function-Body $explorer 'void ExplorerCtrl::applyCustomDrawFiltering' 'void ExplorerCtrl::applyRowFocusDrawState'
$setPath = Function-Body $pathBar 'void PathBar::setPath(' 'void PathBar::OnPaint'

$checks = [Collections.Generic.List[object]]::new()
function Check([string]$name, [bool]$passed) {
    $checks.Add([pscustomobject]@{ Name = $name; Passed = $passed })
}

Check 'ListView custom-draw return values are emitted only at supported stages' (
    $customDraw.Contains('CDDS_PREPAINT') -and
    $customDraw.Contains('*aResult = CDRF_NOTIFYITEMDRAW;') -and
    -not $customDraw.Contains('CDDS_ITEMPREERASE') -and
    ([regex]::Matches($customDraw, 'CDRF_SKIPDEFAULT').Count -eq 1))

Check 'Report selection uses one allocation-free final background after native theme paint' (
    $explorerHeader.Contains('applyReportSelectionDrawState') -and
    $customDraw.Contains('applyReportSelectionDrawState(sNmLvCustomDraw);') -and
    $explorerHeader.Contains('drawFinalReportSelection') -and
    $customDraw.Contains('drawFinalReportSelection(sNmLvCustomDraw);') -and
    $selectionFill.Contains('::GetStockObject(DC_BRUSH)') -and
    $selectionFill.Contains('::FillRect(') -and
    $selectionFill.Contains('::SaveDC(') -and
    $selectionFill.Contains('::RestoreDC(') -and
    -not $selectionFill.Contains('CreateSolidBrush') -and
    $customDraw.Contains('CDRF_NEWFONT | CDRF_NOTIFYSUBITEMDRAW') -and
    $customDraw.Contains('CDRF_NOTIFYPOSTPAINT'))

Check 'Live ListView selection overrides stale callback state for ordinary files and folders' (
    $drawState.Contains('GetItemState(sItemIndex, LVIS_SELECTED | LVIS_FOCUSED)') -and
    $drawState.Contains('if (!XPR_TEST_BITS(sNativeState, LVIS_SELECTED))') -and
    $drawState.Contains('uItemState &= ~(CDIS_SELECTED | CDIS_FOCUS | CDIS_DEFAULT)') -and
    $drawState.Contains('applyRowFocusDrawState(aNmLvCustomDraw);'))

Check 'Every Ctrl and Shift selection receives the pane configured colour while drop highlight stays distinct' (
    $drawState.Contains('applyRowFocusDrawState(aNmLvCustomDraw);') -and
    -not $drawState.Contains('COLOR_HIGHLIGHT') -and
    -not $drawState.Contains('COLOR_BTNFACE') -and
    $selectionFill.Contains('sBackgroundColor = mOption.mRowFocusColor;') -and
    $selectionFill.Contains('LVIS_DROPHILITED') -and
    $selectionFill.Contains('COLOR_HIGHLIGHT') -and
    -not $selectionFill.Contains('COLOR_BTNFACE') -and
    -not $drawState.Contains('SetItemState('))

Check 'Six panes do not share a mutable filtering path buffer' (
    $filtering.Contains('xpr_tchar_t sParsing[XPR_MAX_PATH + 1] = {0};') -and
    -not $filtering.Contains('static xpr_tchar_t sParsing'))

Check 'Parent icon repaint uses a shared resource handle without destroying it' (
    $parentIcon.Contains('AfxGetApp()->LoadIcon(IDI_GO_UP)') -and
    $parentIcon.Contains('::DrawIconEx(') -and
    -not $parentIcon.Contains('DESTROY_ICON') -and
    -not $parentIcon.Contains('::FillRect(') -and
    -not $parentIcon.Contains('::DrawText('))

Check 'Each pane generations its own asynchronous shell-icon requests' (
    $shellIconHeader.Contains('xpr_uint_t    mGeneration;') -and
    $shellIconHeader.Contains('xpr_uint_t mGeneration;') -and
    $queueIcon.Contains('aAsyncIcon->mGeneration = mGeneration;') -and
    $clearIcon.Contains('++mGeneration;'))

Check 'Same-folder refreshes retain one path icon and real path changes release the owned icon first' (
    $setPath.Contains('fxfile::base::Pidl::compare(mFullPidl, aFullPidl) == 0') -and
    $setPath.Contains('XPR_IS_FALSE(sSamePath) || XPR_IS_NULL(mIcon)') -and
    $setPath.Contains('DESTROY_ICON(mIcon);') -and
    $setPath.IndexOf('DESTROY_ICON(mIcon);', [StringComparison]::Ordinal) -lt
        $setPath.IndexOf('mIcon = GetItemIcon(mFullPidl);', [StringComparison]::Ordinal))

Check 'Navigation cancels stale shell calls outside the queue mutex and worker drops stale completions' (
    $clearIcon.Contains('::CoCancelCall(sWorkerThreadId, 0);') -and
    $clearIcon.IndexOf('}', $clearIcon.IndexOf('mIconDeque.clear();', [StringComparison]::Ordinal), [StringComparison]::Ordinal) -lt
        $clearIcon.IndexOf('::CoCancelCall(sWorkerThreadId, 0);', [StringComparison]::Ordinal) -and
    $workerIcon.Contains('sAsyncIcon->mGeneration == mGeneration') -and
    $workerIcon.Contains('XPR_SAFE_DELETE(sAsyncIcon);') -and
    $workerIcon.IndexOf('sCurrentGeneration', [StringComparison]::Ordinal) -lt
        $workerIcon.IndexOf('::PostMessage(', [StringComparison]::Ordinal))

Check 'Late no-op icon and overlay completions cannot invalidate a row' (
    $asyncHandler.Contains('sLvItemData->mCachedIconIndex != sOldIconIndex') -and
    $asyncHandler.Contains('sState != sOldOverlayState') -and
    $asyncHandler.IndexOf('sLvItemData->mCachedIconIndex != sOldIconIndex', [StringComparison]::Ordinal) -lt
        $asyncHandler.IndexOf('SetItem(&sLvItem);', [StringComparison]::Ordinal))

Check 'All six layout slots share this one ExplorerCtrl implementation' (
    $definition.Contains('#define MAX_VIEW_SPLIT_ROW        (2)') -and
    $definition.Contains('#define MAX_VIEW_SPLIT_COLUMN     (3)') -and
    $explorerView.Contains('new ExplorerCtrl'))

Check 'Long-run GUI verification selects real folder and file rows in every pane' (
    $longRunHarness.Contains('mixed_file_and_folder_fixture') -and
    $longRunHarness.Contains('FixtureFolderCount') -and
    $longRunHarness.Contains('FixtureFileCount') -and
    $longRunHarness.Contains('FolderRowSelectionsByPane') -and
    $longRunHarness.Contains('FileRowSelectionsByPane') -and
    $longRunHarness.Contains('Every pane must confirm at least one real folder-row and file-row selection.'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | Format-Table -AutoSize
if ($failed.Count -gt 0) {
    throw ('Task 111 contract failure(s): ' + (($failed | ForEach-Object Name) -join '; '))
}

Write-Host 'PASS: Task 111 six-pane long-run paint and asynchronous icon isolation contracts hold.'
