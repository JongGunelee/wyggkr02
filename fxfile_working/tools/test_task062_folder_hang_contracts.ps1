# Test-Task062FolderHangContracts.ps1
# Verifies that icon/overlay caching, single-request guards, and fast icon index resolution
# contracts are strictly implemented in source files.

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$workingRoot = Split-Path -Parent $scriptDir

Write-Host "=== Task 062 Folder Hang Fixes Contract Tests ===" -ForegroundColor Cyan
Write-Host "WorkingRoot: $workingRoot"

$itemDataPath = Join-Path $workingRoot "src\fxfile\item_data.h"
$explorerCtrlPath = Join-Path $workingRoot "src\fxfile\explorer_ctrl.cpp"

if (-not (Test-Path $itemDataPath)) {
    throw "Missing item_data.h at $itemDataPath"
}
if (-not (Test-Path $explorerCtrlPath)) {
    throw "Missing explorer_ctrl.cpp at $explorerCtrlPath"
}

$itemDataContent = Get-Content $itemDataPath -Raw -Encoding UTF8
$explorerCtrlContent = Get-Content $explorerCtrlPath -Raw -Encoding UTF8

$tests = @(
    @{
        Name = "LVITEMDATA struct constructor definition"
        File = "item_data.h"
        Pattern = "LVITEMDATA\(\s*\)"
    },
    @{
        Name = "LVITEMDATA constructor initializes mCachedIconIndex"
        File = "item_data.h"
        Pattern = "mCachedIconIndex\s*\(\s*-1\s*\)"
    },
    @{
        Name = "LVITEMDATA constructor initializes mIconResolved"
        File = "item_data.h"
        Pattern = "mIconResolved\s*\(\s*XPR_FALSE\s*\)"
    },
    @{
        Name = "LVITEMDATA constructor initializes mIconRequestIssued"
        File = "item_data.h"
        Pattern = "mIconRequestIssued\s*\(\s*XPR_FALSE\s*\)"
    },
    @{
        Name = "LVITEMDATA constructor initializes mOverlayResolved"
        File = "item_data.h"
        Pattern = "mOverlayResolved\s*\(\s*XPR_FALSE\s*\)"
    },
    @{
        Name = "LVITEMDATA constructor initializes mOverlayRequestIssued"
        File = "item_data.h"
        Pattern = "mOverlayRequestIssued\s*\(\s*XPR_FALSE\s*\)"
    },
    @{
        Name = "OnGetdispinfoDriveItem checks mIconResolved"
        File = "explorer_ctrl.cpp"
        Pattern = "(aLvItemData->mIconResolved\s*==\s*XPR_TRUE|XPR_IS_TRUE\(aLvItemData->mIconResolved\))"
    },
    @{
        Name = "OnGetdispinfoShellItem checks mIconResolved for fast return"
        File = "explorer_ctrl.cpp"
        Pattern = "if\s*\(\s*(aLvItemData->mIconResolved\s*==\s*XPR_TRUE|XPR_IS_TRUE\(aLvItemData->mIconResolved\))\s*\)\s*\{\s*aLvItem\.iImage\s*=\s*aLvItemData->mCachedIconIndex;"
    },
    @{
        Name = "OnGetdispinfoShellItem guards custom icon request with mIconRequestIssued"
        File = "explorer_ctrl.cpp"
        Pattern = "if\s*\(\s*(aLvItemData->mIconRequestIssued\s*==\s*XPR_FALSE|XPR_IS_FALSE\(aLvItemData->mIconRequestIssued\))\s*\)"
    },
    @{
        Name = "OnGetdispinfoShellItem guards overlay request with mOverlayResolved and mOverlayRequestIssued"
        File = "explorer_ctrl.cpp"
        Pattern = "(XPR_IS_FALSE\(aLvItemData->mOverlayRequestIssued\)\s*&&\s*XPR_IS_FALSE\(aLvItemData->mOverlayResolved\)|aLvItemData->mOverlayResolved\s*==\s*XPR_TRUE)"
    },
    @{
        Name = "OnGetdispinfoShellItem sets mOverlayRequestIssued flag"
        File = "explorer_ctrl.cpp"
        Pattern = "aLvItemData->mOverlayRequestIssued\s*=\s*XPR_TRUE;"
    },
    @{
        Name = "OnShellAsyncIcon caches resolved icon index"
        File = "explorer_ctrl.cpp"
        Pattern = "sLvItemData->mCachedIconIndex\s*=\s*sAsyncIcon->mResult\.mIconIndex;"
    },
    @{
        Name = "OnShellAsyncIcon marks icon resolved and sets request issued flag"
        File = "explorer_ctrl.cpp"
        Pattern = "sLvItemData->mIconResolved\s*=\s*XPR_TRUE;"
    },
    @{
        Name = "OnShellAsyncIcon marks overlay resolved and sets request issued flag"
        File = "explorer_ctrl.cpp"
        Pattern = "sLvItemData->mOverlayResolved\s*=\s*XPR_TRUE;"
    },
    @{
        Name = "getFileIconIndex avoids unrealized sparse extension icons and uses a realized generic fallback"
        File = "explorer_ctrl.cpp"
        Pattern = "Do NOT call GetFileExtIconIndex here[\s\S]*sCachedDefaultFileIconIndex[\s\S]*SHGFI_SYSICONINDEX\s*\|\s*SHGFI_USEFILEATTRIBUTES"
    }
)

$passed = 0
$failed = 0

foreach ($t in $tests) {
    $content = if ($t.File -eq "item_data.h") { $itemDataContent } else { $explorerCtrlContent }
    $matched = [regex]::IsMatch($content, $t.Pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if ($matched) {
        Write-Host "  [PASS] $($t.Name)" -ForegroundColor Green
        $passed++
    }
    else {
        Write-Host "  [FAIL] $($t.Name) (Pattern: $($t.Pattern))" -ForegroundColor Red
        $failed++
    }
}

Write-Host ""
Write-Host "Summary: Passed=$passed, Failed=$failed, Total=$($tests.Count)" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Red" })

if ($failed -gt 0) {
    exit 1
}

exit 0
