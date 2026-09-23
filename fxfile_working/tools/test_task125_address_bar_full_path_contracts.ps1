$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$addressHeader = Get-Content -LiteralPath (Join-Path $root 'src\fxfile\address_bar.h') -Raw -Encoding UTF8
$address = Get-Content -LiteralPath (Join-Path $root 'src\fxfile\address_bar.cpp') -Raw -Encoding UTF8
$pane = Get-Content -LiteralPath (Join-Path $root 'src\fxfile\explorer_pane.cpp') -Raw -Encoding UTF8
$showPath = [regex]::Match($address, '(?s)xpr_bool_t AddressBar::showCurrentPath\(LPITEMIDLIST.*?(?=\r?\nxpr_bool_t AddressBar::showCurrentPathText)').Value

function Check([bool]$condition, [string]$message) {
    if (-not $condition) { throw $message }
    Write-Host "PASS $message"
}

Check ($addressHeader -match 'showCurrentPathText\(const xpr_tchar_t \*aFullPath\)') 'AddressBar exposes one canonical text commit path for filesystem and archive navigation.'
Check ($address -match '(?s)AddressBar::showCurrentPath\(LPITEMIDLIST aFullPidl\).*?GetDispFullPath\(aFullPidl, sFullPath\).*?showCurrentPathText\(sFullPath\)') 'Filesystem address text is always derived from the complete PIDL display path.'
Check ($address -match '(?s)AddressBar::showCurrentPathText.*?SetCurSel\(-1\);\s*mCurSel = -1;.*?SetWindowText\(aFullPath\)') 'Native and mirrored ComboBoxEx selections are cleared before the canonical full path is committed.'
Check ($address -match '(?s)if \(sResult == XPR_TRUE\)\s*\{\s*SetCurSel\(sCurSel\);.*?showCurrentPath\(sCbItemData->mFullPidl\);\s*\}') 'A successful drop-down navigation restores the full path after native SetCurSel writes its short label.'
Check ($pane -match '(?s)isInArchiveMode\(\).*?showCurrentPathText\(aCurPath\).*?showCurrentPath\(aFullPidl\)') 'Archive and normal folder navigation share the selection-safe address display contract.'
Check ($showPath -notmatch 'GetName\([^;]*SHGDN_INFOLDER') 'The current address display never falls back to a leaf-only shell name.'
