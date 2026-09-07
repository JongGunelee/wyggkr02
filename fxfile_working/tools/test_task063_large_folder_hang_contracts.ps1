
# test_task063_large_folder_hang_contracts.ps1
# Task 063 -- Large folder icon/overlay blocking prevention contracts
param([string]$WorkDir = '')

if ([string]::IsNullOrWhiteSpace($WorkDir)) {
    $WorkDir = Split-Path -Parent $PSScriptRoot
}

$pass = 0; $fail = 0
function Check([string]$desc, [bool]$cond) {
    if ($cond) { Write-Host "  PASS: $desc" -ForegroundColor Green; $script:pass++ }
    else        { Write-Host "  FAIL: $desc" -ForegroundColor Red;   $script:fail++ }
}

Write-Host "=== Task 063 Contract Tests ===" -ForegroundColor Cyan

$siCpp = [IO.File]::ReadAllText((Join-Path $WorkDir 'src\fxfile\shell_icon.cpp'))
$ecCpp = [IO.File]::ReadAllText((Join-Path $WorkDir 'src\fxfile\explorer_ctrl.cpp'))

# shell_icon.cpp contracts
Check "ShellIcon: kMaxQueueSize constant defined"     ($siCpp -match "kMaxQueueSize")
Check "ShellIcon: queue limit = 300"                  ($siCpp -match "kMaxQueueSize\s*=\s*300")
Check "ShellIcon: returns false when queue full"      ($siCpp -match "mIconDeque\.size\(\)")
Check "ShellIcon: push_back after size check"         ($siCpp.IndexOf("kMaxQueueSize") -lt $siCpp.IndexOf("push_back"))

# getFileIconIndex folder icon contracts
Check "getFileIconIndex: folder fallback is realized without touching the real folder" (
    $ecCpp -match "sCachedFolderIconIndex" -and
    $ecCpp -match "FILE_ATTRIBUTE_DIRECTORY" -and
    $ecCpp -match "SHGFI_SYSICONINDEX\s*\|\s*SHGFI_USEFILEATTRIBUTES")
Check "getFileIconIndex: unrealized sparse extension fast path remains disabled" (
    $ecCpp -match "Do NOT call GetFileExtIconIndex here" -and
    $ecCpp -match "sCachedDefaultFileIconIndex")

# Overlay filter contracts
Check "Overlay: kOverlayRelevantAttr constant defined"     ($ecCpp -match "kOverlayRelevantAttr")
Check "Overlay: SFGAO_LINK filter present"                 ($ecCpp -match "SFGAO_LINK")
Check "Overlay: SFGAO_SHARE filter present"                ($ecCpp -match "SFGAO_SHARE")
Check "Overlay: XPR_TEST_BITS filter applied"              ($ecCpp -match "XPR_TEST_BITS.*kOverlayRelevantAttr")

# Task 062 regression guards
Check "Regression: mIconResolved guard present"            ($ecCpp -match "mIconResolved")
Check "Regression: mIconRequestIssued guard present"       ($ecCpp -match "mIconRequestIssued")
Check "Regression: mOverlayResolved guard present"         ($ecCpp -match "mOverlayResolved")
Check "Regression: mOverlayRequestIssued guard present"    ($ecCpp -match "mOverlayRequestIssued")

Write-Host ""
if ($fail -eq 0) {
    Write-Host "Result: $pass PASS / $fail FAIL" -ForegroundColor Cyan
} else {
    Write-Host "Result: $pass PASS / $fail FAIL" -ForegroundColor Red
}
exit $fail
