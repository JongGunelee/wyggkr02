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

$checks = [Collections.Generic.List[object]]::new()
function Check([string]$name, [bool]$passed) {
    $checks.Add([pscustomobject]@{ Name = $name; Passed = $passed })
}

$optionHeader = Read-Text 'src\fxfile\option.h'
$option = Read-Text 'src\fxfile\option.cpp'
$definition = Read-Text 'src\fxfile\fxfile_def.h'
$dialogHeader = Read-Text 'src\fxfile\cfg\cfg_appearance_color_dlg.h'
$dialog = Read-Text 'src\fxfile\cfg\cfg_appearance_color_dlg.cpp'
$pane = Read-Text 'src\fxfile\explorer_pane.cpp'
$explorerHeader = Read-Text 'src\fxfile\explorer_ctrl.h'
$explorer = Read-Text 'src\fxfile\explorer_ctrl.cpp'
$language = Read-Text 'src\fxfile\Languages\Korean.xml'
$resource = Read-Text 'src\fxfile\fxfile.rc'
$applyOption = Function-Body $explorer 'void ExplorerCtrl::applyOption' 'void ExplorerCtrl::applyTextColor'
$cacheRowFocusOption = Function-Body $explorer 'void ExplorerCtrl::cacheRowFocusOption' 'void ExplorerCtrl::setOption'
$reportSelectionDraw = Function-Body $explorer 'void ExplorerCtrl::applyReportSelectionDrawState' 'void ExplorerCtrl::drawParentFolderIcon'
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$thumbnailDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdrawThumbnail' 'LRESULT ExplorerCtrl::OnThumbnailProc'

$rowFocusKeys = 1..6 | ForEach-Object { "config.view$_.file_list.row_focus_color" }
Check 'Automatic row-focus color is pure white' (
    $definition.Contains('#define DEF_FILE_LIST_ROW_FOCUS_COLOR (RGB(255,255,255))'))
Check 'Six per-pane row-focus color keys default to white' (
    $optionHeader.Contains('mFileListRowFocusColor[MAX_VIEW_SPLIT]') -and
    (@($rowFocusKeys | Where-Object { -not $option.Contains($_) }).Count -eq 0) -and
    ([regex]::Matches($option, 'mFileListRowFocusColor\[\d\].*DEF_FILE_LIST_ROW_FOCUS_COLOR').Count -eq 6))
Check 'Color dialog loads, saves, and applies the per-pane row-focus color' (
    $dialogHeader.Contains('mFileListRowFocusColor') -and
    $dialogHeader.Contains('mFileListRowFocusColorCtrl') -and
    $dialog.Contains('mFileListRowFocusColorCtrl.SetDefaultColor(DEF_FILE_LIST_ROW_FOCUS_COLOR)') -and
    $dialog.Contains('pViewColor->mFileListRowFocusColor    = aConfig.mFileListRowFocusColor[i];') -and
    $dialog.Contains('aConfig.mFileListRowFocusColor[i]    = sViewColor->mFileListRowFocusColor;') -and
    $dialog.Contains('aViewColor.mFileListRowFocusColor    = mFileListRowFocusColorCtrl.GetColor();'))
Check 'Every Explorer pane receives its own configured row-focus color' (
    $pane.Contains('sOption.mRowFocusColor                    = aOption.mConfig.mFileListRowFocusColor[mViewIndex];'))
Check 'Selected active report rows use the configured color in the final item paint without changing selection state' (
    $explorerHeader.Contains('mRowFocusColor') -and
    $customDraw.Contains('XPR_IS_TRUE(mOption.mFullRowSelect)') -and
    $reportSelectionDraw.Contains('if (!XPR_TEST_BITS(sNativeState, LVIS_SELECTED))') -and
    -not $reportSelectionDraw.Contains('isFocusedSelectedItem(sItemIndex)') -and
    $customDraw.Contains('CDRF_NOTIFYPOSTPAINT') -and
    $reportSelectionDraw.Contains('applyRowFocusDrawState(aNmLvCustomDraw);') -and
    $customDraw.Contains('CDRF_NEWFONT') -and
    $customDraw.Contains('applyReportSelectionDrawState(sNmLvCustomDraw);') -and
    -not $customDraw.Contains('SetItemState(') -and
    -not $customDraw.Contains('SetItem('))
Check 'Custom colors receive cached contrast text and do not add hot-path allocation or scheduling' (
    $applyOption.Contains('cacheRowFocusOption(aNewOption);') -and
    $cacheRowFocusOption.Contains('mRowFocusTextColor = sLuminance < 128') -and
    $cacheRowFocusOption.Contains('GetSysColor(COLOR_HIGHLIGHTTEXT)') -and
    -not $customDraw.Contains('new ') -and
    -not $customDraw.Contains('SetTimer(') -and
    -not $customDraw.Contains('PostMessage('))
Check 'Thumbnail mode uses the same cached color only for the active full-row presentation' (
    $thumbnailDraw.Contains('sActiveFocusColor') -and
    $thumbnailDraw.Contains('? mOption.mRowFocusColor') -and
    $thumbnailDraw.Contains('? mRowFocusTextColor'))
Check 'Environment setting has an accessible Korean label and matching resource controls' (
    $language.Contains('popup.cfg.body.appearance.color.view.label.file_list_row_focus_color') -and
    $resource.Contains('Selected &row focus color:') -and
    $resource.Contains('IDC_CFG_COLOR_FILE_LIST_ROW_FOCUS_COLOR'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object {
    '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name
}
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
