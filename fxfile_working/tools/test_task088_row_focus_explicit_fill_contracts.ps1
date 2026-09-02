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

$definition = Read-Text 'src\fxfile\fxfile_def.h'
$option = Read-Text 'src\fxfile\option.cpp'
$dialog = Read-Text 'src\fxfile\cfg\cfg_appearance_color_dlg.cpp'
$pane = Read-Text 'src\fxfile\explorer_pane.cpp'
$explorer = Read-Text 'src\fxfile\explorer_ctrl.cpp'
$header = Read-Text 'src\fxfile\explorer_ctrl.h'
$probe = Read-Text 'tools\row_focus_visual_probe.cpp'

$fillFocus = Function-Body $explorer 'void ExplorerCtrl::fillRowFocusBackground' 'void ExplorerCtrl::applyRowFocusDrawState'
$focusState = Function-Body $explorer 'void ExplorerCtrl::applyRowFocusDrawState' 'void ExplorerCtrl::OnCustomdraw('
$customDraw = Function-Body $explorer 'void ExplorerCtrl::OnCustomdraw(' 'void ExplorerCtrl::OnCustomdrawThumbnail'
$subItemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM))' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)'
$itemDraw = Function-Body $customDraw 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)' 'else if (sNmLvCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT)'
$reportItemDraw = Function-Body $itemDraw 'if (XPR_IS_TRUE(isReportView()) &&' 'else if (getViewStyle() != VIEW_STYLE_THUMBNAIL)'

Check 'Automatic and all six new-profile row-focus defaults remain pure white' (
    $definition.Contains('#define DEF_FILE_LIST_ROW_FOCUS_COLOR (RGB(255,255,255))') -and
    ([regex]::Matches($option, 'mFileListRowFocusColor\[\d\].*DEF_FILE_LIST_ROW_FOCUS_COLOR').Count -eq 6) -and
    $dialog.Contains('mFileListRowFocusColorCtrl.SetDefaultColor(DEF_FILE_LIST_ROW_FOCUS_COLOR)'))
Check 'Each of the six panes still receives its independently saved row-focus color' (
    ([regex]::Matches($option, 'config\.view[1-6]\.file_list\.row_focus_color').Count -eq 6) -and
    $pane.Contains('aOption.mConfig.mFileListRowFocusColor[mViewIndex]'))
Check 'The focused report item is painted on every native repaint with no one-shot gate' (
    $itemDraw.Contains('XPR_IS_TRUE(isReportView())') -and
    $itemDraw.Contains('XPR_IS_TRUE(sFocusedSelected)') -and
    $itemDraw.Contains('fillRowFocusBackground(sNmLvCustomDraw);') -and
    $itemDraw.Contains('CDRF_NEWFONT | CDRF_NOTIFYSUBITEMDRAW') -and
    -not $customDraw.Contains('mRowFocusPaintPending'))
Check 'Full-row and legacy modes use mutually exclusive Win32 selection geometry' (
    $header.Contains('fillRowFocusBackground') -and
    $fillFocus.Contains('LVIR_BOUNDS : LVIR_SELECTBOUNDS') -and
    $fillFocus.Contains('GetItemRect(') -and
    $fillFocus.Contains('IntersectRect('))
Check 'The background fill is allocation-free and restores the shared DC brush state' (
    $fillFocus.Contains('GetStockObject(DC_BRUSH)') -and
    $fillFocus.Contains('SetDCBrushColor(') -and
    $fillFocus.Contains('::FillRect(') -and
    ([regex]::Matches($fillFocus, 'SetDCBrushColor\(').Count -eq 2) -and
    -not $fillFocus.Contains('CreateSolidBrush') -and
    -not $fillFocus.Contains('DeleteObject') -and
    -not $fillFocus.Contains('new '))
Check 'Theme suppression changes only transient draw state and supplies readable text colors' (
    $focusState.Contains('uItemState &= ~CDIS_SELECTED') -and
    -not [regex]::IsMatch($focusState, '(?m)^\s*aNmLvCustomDraw->iStateId\s*=') -and
    -not [regex]::IsMatch($focusState, '(?m)^\s*aNmLvCustomDraw->clrFace\s*=') -and
    $focusState.Contains('clrTextBk = mOption.mRowFocusColor;') -and
    $focusState.Contains('clrText   = mRowFocusTextColor;'))
Check 'Every focused subitem is reset before full-row or column-zero scope is applied' (
    $subItemDraw.IndexOf('resetCustomDrawColors(sNmLvCustomDraw);', [StringComparison]::Ordinal) -lt
        $subItemDraw.IndexOf('applyRowFocusDrawState(sNmLvCustomDraw);', [StringComparison]::Ordinal) -and
    $subItemDraw.Contains('applyCustomDrawFiltering(sNmLvCustomDraw);') -and
    $subItemDraw.Contains('XPR_IS_TRUE(mOption.mFullRowSelect) || sNmLvCustomDraw->iSubItem == 0'))
Check 'The report path retains native icon and text rendering and never schedules repaint work' (
    -not $subItemDraw.Contains('CDRF_SKIPDEFAULT') -and
    -not $reportItemDraw.Contains('CDRF_SKIPDEFAULT') -and
    -not $customDraw.Contains('SetItemState(') -and
    -not $customDraw.Contains('Invalidate(') -and
    -not $customDraw.Contains('RedrawItems(') -and
    -not $customDraw.Contains('PostMessage(') -and
    -not $customDraw.Contains('SetTimer('))
Check 'The visual characterization covers failed theme-only paths and both passing scopes' (
    $probe.Contains('ProbeCurrent') -and
    $probe.Contains('ProbeThemeState') -and
    $probe.Contains('ProbeThemeStateAndFill') -and
    $probe.Contains('ProbeLegacyCellFill') -and
    $probe.Contains('LVIR_BOUNDS : LVIR_SELECTBOUNDS') -and
    $probe.Contains('index != ProbeLegacyCellFill'))

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
