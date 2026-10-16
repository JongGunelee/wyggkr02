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

$option = Read-Text 'src\fxfile\option.cpp'
$dialog = Read-Text 'src\fxfile\cfg\cfg_appearance_file_list_dlg.cpp'
$language = Read-Text 'src\fxfile\Languages\Korean.xml'
$resource = Read-Text 'src\fxfile\fxfile.rc'
$pane = Read-Text 'src\fxfile\explorer_pane.cpp'
$explorer = Read-Text 'src\fxfile\explorer_ctrl.cpp'
$applyOption = Function-Body $explorer 'void ExplorerCtrl::applyOption' 'void ExplorerCtrl::applyTextColor'

Check 'Legacy preference defaults to full-row selection for new profiles' (
    $option.Contains('config.file_list.full_row_select') -and
    $option.Contains('(void *)XPR_TRUE'))
Check 'Settings dialog loads and saves either focus mode' (
    -not $dialog.Contains('IDC_CFG_FILE_LIST_FULL_ROW_SELECTION)->EnableWindow(XPR_FALSE)') -and
    $dialog.Contains('IDC_CFG_FILE_LIST_FULL_ROW_SELECTION  ))->SetCheck(aConfig.mFileListFullRowSelect)') -and
    $dialog.Contains('aConfig.mFileListFullRowSelect     = ((CButton *)GetDlgItem(IDC_CFG_FILE_LIST_FULL_ROW_SELECTION ))->GetCheck();'))
Check 'All explorer panes receive the same saved focus preference' (
    $pane.Contains('sOption.mFullRowSelect                    = aOption.mConfig.mFileListFullRowSelect'))
Check 'Explorer control applies the selected focus mode without changing selection state' (
    $applyOption.Contains('XPR_SET_OR_CLR_BITS(sExStyle, LVS_EX_FULLROWSELECT, aNewOption.mFullRowSelect);') -and
    -not $applyOption.Contains('sExStyle |= LVS_EX_FULLROWSELECT;'))
$hasKoreanFullRowFocusLabel = $language.Contains('popup.cfg.body.appearance.file_list.check.full_row_selection')
$hasResourceFullRowFocusLabel = $resource.Contains('Use &full row focus')
Check 'Settings wording identifies the selectable full-row focus mode' ($hasKoreanFullRowFocusLabel -and $hasResourceFullRowFocusLabel)

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object {
    '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name
}
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
