[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$mainFrame = [IO.File]::ReadAllText((Join-Path $root 'src\fxfile\main_frame.cpp'))
$definition = [IO.File]::ReadAllText((Join-Path $root 'src\fxfile\fxfile_def.h'))

function Function-Body([string]$text, [string]$signature, [string]$nextSignature) {
    $start = $text.IndexOf($signature, [StringComparison]::Ordinal)
    if ($start -lt 0) { return '' }
    $end = $text.IndexOf($nextSignature, $start + $signature.Length, [StringComparison]::Ordinal)
    if ($end -lt 0) { $end = $text.Length }
    return $text.Substring($start, $end - $start)
}

$onCreateClient = Function-Body $mainFrame 'xpr_bool_t MainFrame::OnCreateClient' 'void MainFrame::OnClose'
$checks = @(
    [pscustomobject]@{ Name = 'FxFile supports a six-pane 2x3 layout'; Passed = $definition.Contains('#define MAX_VIEW_SPLIT_ROW        (2)') -and $definition.Contains('#define MAX_VIEW_SPLIT_COLUMN     (3)') },
    [pscustomobject]@{ Name = 'The -w command applies its row argument only to rows'; Passed = $onCreateClient.Contains('sRowCount = sCmdRowCount;') },
    [pscustomobject]@{ Name = 'The -w command applies its column argument to columns'; Passed = $onCreateClient.Contains('sColumnCount = sCmdColumnCount;') -and -not $onCreateClient.Contains('sRowCount = sCmdColumnCount;') },
    [pscustomobject]@{ Name = 'Both command dimensions are range-checked before creating panes'; Passed = $onCreateClient.Contains('XPR_IS_RANGE(1, sCmdRowCount, MAX_VIEW_SPLIT_ROW)') -and $onCreateClient.Contains('XPR_IS_RANGE(1, sCmdColumnCount, MAX_VIEW_SPLIT_COLUMN)') }
)

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { '{0} {1}' -f ($(if ($_.Passed) { 'PASS' } else { 'FAIL' })), $_.Name }
"SUMMARY PASS=$($checks.Count - $failed.Count) FAIL=$($failed.Count) TOTAL=$($checks.Count)"
if ($failed.Count -ne 0) { exit 1 }
