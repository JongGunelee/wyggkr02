
# test_task063_large_folder_hang_contracts.ps1
# Task 063 -- Large folder icon/overlay blocking prevention contracts
param([string]$WorkDir = "d:\03 금일작업\00 임시\0000 FxFile\fxfile_working")

$pass = 0; $fail = 0
function Check([string]$desc, [bool]$cond) {
    if ($cond) { Write-Host "  PASS: $desc" -ForegroundColor Green; $script:pass++ }
    else        { Write-Host "  FAIL: $desc" -ForegroundColor Red;   $script:fail++ }
}

Write-Host "=== Task 063 Contract Tests ===" -ForegroundColor Cyan

$siCpp = Get-Content "$WorkDir\src\fxfile\shell_icon.cpp" -Raw -Encoding UTF8
$ecCpp = Get-Content "$WorkDir\src\fxfile\explorer_ctrl.cpp" -Raw -Encoding UTF8

# shell_icon.cpp contracts
Check "ShellIcon: kMaxQueueSize constant defined"     ($siCpp -match "kMaxQueueSize")
Check "ShellIcon: queue limit = 300"                  ($siCpp -match "kMaxQueueSize\s*=\s*300")
Check "ShellIcon: returns false when queue full"      ($siCpp -match "mIconDeque\.size\(\)")
Check "ShellIcon: push_back after size check"         ($siCpp.IndexOf("kMaxQueueSize") -lt $siCpp.IndexOf("push_back"))

# getFileIconIndex folder icon contracts
Check "getFileIconIndex: SFGAO_FOLDER returns sIconIndex=6"       ($ecCpp -match "sIconIndex\s*=\s*6.*generic closed folder")
Check "getFileIconIndex: GetFileExtIconIndex fast path preserved"  ($ecCpp -match "GetFileExtIconIndex\(sExt\)")

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
