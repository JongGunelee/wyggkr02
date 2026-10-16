param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$header = [IO.File]::ReadAllText((Join-Path $root 'src\fxfile\explorer_ctrl.h'))
$source = [IO.File]::ReadAllText((Join-Path $root 'src\fxfile\explorer_ctrl.cpp'))

$passed = 0
$failed = [Collections.Generic.List[string]]::new()

function Assert-Contains([string]$Text, [string]$Needle, [string]$Message) {
    if ($Text.Contains($Needle)) { $script:passed++ }
    else { $script:failed.Add($Message) }
}

function Assert-Matches([string]$Text, [string]$Pattern, [string]$Message) {
    if ([Text.RegularExpressions.Regex]::IsMatch($Text, $Pattern, [Text.RegularExpressions.RegexOptions]::Singleline)) {
        $script:passed++
    } else {
        $script:failed.Add($Message)
    }
}

# Both Details and Content are Win32 report views; the logical enum must not
# suppress responsive columns in the user's normal Content layout.
Assert-Contains $source 'return ((GetStyle() & LVS_TYPEMASK) == LVS_REPORT)' `
    'Responsive columns must use the actual Win32 report style.'

# Every pane owns its own cache/guard. A function-static guard couples four
# panes and can discard real user header notifications.
Assert-Contains $header 'mAutomaticPreferredColumnWidths' `
    'Preferred-width cache must be an ExplorerCtrl instance member.'
Assert-Contains $header 'mInAutomaticColumnLayout' `
    'Programmatic width guard must be an ExplorerCtrl instance member.'
Assert-Contains $header 'mLastAutomaticLayoutClientWidth' `
    'Responsive layout must remember the last real client width.'
Assert-Matches $source 'void ExplorerCtrl::OnHdnItemChanged.*?mInAutomaticColumnLayout' `
    'Header notifications must reject programmatic responsive changes.'
Assert-Matches $source 'void ExplorerCtrl::OnSize.*?mInAutomaticColumnLayout.*?mLastAutomaticLayoutClientWidth' `
    'WM_SIZE must reject programmatic re-entry and duplicate client widths.'

# WM_SIZE only schedules a bounded allocator. Content scans are deliberately
# separated and sample at most 64 rows regardless of the folder item count.
Assert-Contains $source 'ON_WM_SIZE()' 'ExplorerCtrl must receive pane-size changes.'
Assert-Matches $source 'void ExplorerCtrl::OnSize.*?scheduleAutomaticColumnReflow\(\)' `
    'WM_SIZE must debounce responsive reflow.'
Assert-Contains $source 'SetTimer(TM_ID_AUTO_COLUMN_REFLOW, 100, XPR_NULL)' `
    'Responsive resize must use the 100 ms per-pane debounce timer.'
Assert-Matches $source 'void ExplorerCtrl::adjustAutomaticColumnWidths.*?measureAutomaticColumnWidths\(\).*?reflowAutomaticColumnWidths\(\)' `
    'Content measurement and viewport reflow must be separate phases.'
Assert-Contains $source 'measureAutomaticColumnTextWidth(i, *sColumnId)' `
    'AutoFull content width must use the text-only measurement path.'
Assert-Contains $source 'const xpr_sint_t kMaximumSamples = 64' `
    'Synchronous content measurement must remain capped at 64 sampled rows.'
$reflowStart = $source.IndexOf('void ExplorerCtrl::reflowAutomaticColumnWidths')
$reflowEnd = $source.IndexOf('void ExplorerCtrl::scheduleAutomaticColumnReflow', $reflowStart)
if ($reflowStart -ge 0 -and $reflowEnd -gt $reflowStart) {
    $reflowBody = $source.Substring($reflowStart, $reflowEnd - $reflowStart)
    if (-not $reflowBody.Contains('LVSCW_AUTOSIZE')) { $passed++ }
    else { $failed.Add('WM_SIZE reflow must never rescan item text through LVSCW_AUTOSIZE.') }
    if ($reflowBody.Contains('if (XPR_IS_FALSE(sChanged))') -and
        $reflowBody.IndexOf('if (XPR_IS_FALSE(sChanged))') -lt $reflowBody.IndexOf('SetRedraw(XPR_FALSE)')) {
        $passed++
    } else {
        $failed.Add('A no-op responsive pass must return before invalidating/redrawing the list view.')
    }
} else {
    $failed.Add('Viewport reflow implementation boundaries were not found.')
}

$measureStart = $source.IndexOf('void ExplorerCtrl::measureAutomaticColumnWidths')
$measureEnd = $source.IndexOf('void ExplorerCtrl::reflowAutomaticColumnWidths', $measureStart)
if ($measureStart -ge 0 -and $measureEnd -gt $measureStart) {
    $measureBody = $source.Substring($measureStart, $measureEnd - $measureStart)
    if (-not $measureBody.Contains('SetColumnWidth(i, LVSCW_AUTOSIZE)')) { $passed++ }
    else { $failed.Add('Content measurement must not invoke native LVSCW_AUTOSIZE icon/thumbnail redraws.') }
} else {
    $failed.Add('Content measurement implementation boundaries were not found.')
}

# Auto-managed display widths are transient. Manual header widths must update
# the retained FolderLayout baseline by ColumnId before a save/restart.
Assert-Contains $source 'isAutomaticStandardColumn(mOption, *sColumnId)' `
    'All standard responsive columns must preserve retained manual widths.'
Assert-Matches $source 'void ExplorerCtrl::OnHdnItemChanged.*?rememberManualColumnWidth\(sColumnIndex, sWidth\)' `
    'A real header drag must update its retained manual width.'
Assert-Contains $source 'if (mFolderLayout->mColumnItem[i] == *sColumnId)' `
    'Manual width persistence must match stable ColumnId rather than display index.'
$headerChangedStart = $source.IndexOf('void ExplorerCtrl::OnHdnItemChanged')
$headerChangedEnd = $source.IndexOf('xpr_bool_t ExplorerCtrl::OnEraseBkgnd', $headerChangedStart)
if ($headerChangedStart -ge 0 -and $headerChangedEnd -gt $headerChangedStart) {
    $headerChangedBody = $source.Substring($headerChangedStart, $headerChangedEnd - $headerChangedStart)
    if (-not $headerChangedBody.Contains('scheduleAutomaticColumnReflow()')) { $passed++ }
    else { $failed.Add('A manual header drag must not immediately schedule viewport reflow and snap back.') }
} else {
    $failed.Add('Manual header-width handler boundaries were not found.')
}

# Folder size must only change the sampling step; it must never skip the
# responsive reflow pipeline or iterate every item in a large folder.
Assert-Contains $source 'const xpr_sint_t sStep = (sItemCount > kMaximumSamples)' `
    'Large folders must use bounded, distributed content sampling.'

if ($failed.Count -gt 0) {
    foreach ($failure in $failed) { Write-Error "FAIL: $failure" -ErrorAction Continue }
    throw "$($failed.Count) responsive-column contract checks failed; $passed passed."
}

Write-Host "PASS: $passed responsive-column contracts; static/read-only test."
