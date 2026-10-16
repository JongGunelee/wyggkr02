param()

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$itemDataPath = Join-Path $root 'src\fxfile\item_data.h'
$explorerPath = Join-Path $root 'src\fxfile\explorer_ctrl.cpp'
$fileOpPath = Join-Path $root 'src\fxfile\file_op_thread.cpp'
$adaptivePath = Join-Path $root 'src\fxfile\adaptive_file_operation.cpp'
$mainFramePath = Join-Path $root 'src\fxfile\main_frame.cpp'
$languagePath = Join-Path $root 'src\fxfile\Languages\Korean.xml'

$itemData = [IO.File]::ReadAllText($itemDataPath)
$explorer = [IO.File]::ReadAllText($explorerPath)
$fileOp = [IO.File]::ReadAllText($fileOpPath)
$adaptive = [IO.File]::ReadAllText($adaptivePath)
$mainFrame = [IO.File]::ReadAllText($mainFramePath)

$passed = 0
$failed = 0

function Test-Contract([string]$Name, [bool]$Condition) {
    if ($Condition) {
        $script:passed++
        Write-Host "PASS: $Name"
    }
    else {
        $script:failed++
        Write-Host "FAIL: $Name" -ForegroundColor Red
    }
}

Test-Contract 'per-item icon cache exists' ($itemData.Contains('mCachedIconIndex'))
Test-Contract 'per-item icon resolved state exists' ($itemData.Contains('mIconResolved'))
Test-Contract 'per-item icon request guard exists' ($itemData.Contains('mIconRequestIssued'))
Test-Contract 'per-item overlay cache exists' ($itemData.Contains('mCachedOverlayState'))
Test-Contract 'per-item overlay resolved state exists' ($itemData.Contains('mOverlayResolved'))
Test-Contract 'per-item overlay request guard exists' ($itemData.Contains('mOverlayRequestIssued'))

Test-Contract 'shell insertion initializes cache sentinel twice' (([regex]::Matches($explorer, 'mCachedIconIndex\s*=\s*-1')).Count -eq 2)
$displayStart = $explorer.IndexOf('xpr_bool_t ExplorerCtrl::OnGetdispinfoShellItem(')
$displayEnd = $explorer.IndexOf('void ExplorerCtrl::OnGetdispinfo(', $displayStart)
$displayBody = if ($displayStart -ge 0 -and $displayEnd -gt $displayStart) {
    $explorer.Substring($displayStart, $displayEnd - $displayStart)
} else { '' }
Test-Contract 'reentrant display callback uses local path buffers' ($displayBody.Length -gt 0 -and -not $displayBody.Contains('static xpr_tchar_t sPath[XPR_MAX_PATH + 1]'))
Test-Contract 'resolved icon is returned from cache' ($explorer.Contains('aLvItem.iImage = aLvItemData->mCachedIconIndex'))
Test-Contract 'icon request is marked before enqueue' ($explorer.Contains('aLvItemData->mIconRequestIssued = XPR_TRUE;') -and $explorer.Contains('aLvItemData->mIconRequestIssued = XPR_FALSE;'))
Test-Contract 'overlay request is guarded' ($explorer.Contains('XPR_IS_FALSE(aLvItemData->mOverlayRequestIssued)') -and $explorer.Contains('XPR_IS_FALSE(aLvItemData->mOverlayResolved)'))
Test-Contract 'async icon result updates cache' ($explorer.Contains('sLvItemData->mCachedIconIndex = sAsyncIcon->mResult.mIconIndex'))
Test-Contract 'async overlay result updates cache' ($explorer.Contains('sLvItemData->mCachedOverlayState = sState'))

Test-Contract 'copy completion uses nonblocking shell notification' ($fileOp.Contains('SHCNF_PATH | SHCNF_FLUSHNOWAIT'))
Test-Contract 'adaptive rollback is verified before fallback' ($adaptive.Contains('bool rollbackTargets(') -and $adaptive.Contains('sRollbackSucceeded && canRetryWithModernShell(sError)'))
Test-Contract 'volatile source failures retry through modern shell' ($adaptive.Contains('case ERROR_FILE_NOT_FOUND:') -and $adaptive.Contains('case ERROR_FILE_INVALID:') -and $adaptive.Contains('return ResultNotApplicable;'))
Test-Contract 'copy failure captures the exact source job' ($adaptive.Contains('firstFailureJob.store(sIndex)') -and $adaptive.Contains('sPlan.files[sFailureJob].source.c_str()'))
Test-Contract 'incomplete rollback blocks shell retry' ($adaptive.Contains('aRollbackSucceeded') -and $adaptive.Contains('if (!aRollbackSucceeded)'))
Test-Contract 'file-op buffers are not freed while worker is active' ($fileOp.Contains('sExitCode == STILL_ACTIVE') -and $fileOp.Contains('return XPR_FALSE;'))
Test-Contract 'file-op handles are closed' (($fileOp.Contains('CLOSE_HANDLE(mThread)')) -and ($fileOp.Contains('CLOSE_HANDLE(mStopEvent)')))
Test-Contract 'active file operation blocks unsafe app teardown' ($mainFrame.Contains('main_frame.msg.exit_blocked_file_operation') -and $mainFrame.Contains('MB_ICONWARNING | MB_OK'))

try {
    $languageXml = New-Object System.Xml.XmlDocument
    $languageXml.Load($languagePath)
    Test-Contract 'Korean language XML parses' $true
}
catch {
    Test-Contract 'Korean language XML parses' $false
}
Test-Contract 'blocked-close guidance is translated' (([IO.File]::ReadAllText($languagePath)).Contains('main_frame.msg.exit_blocked_file_operation'))

Write-Host "RESULT: $passed PASS / $failed FAIL"
if ($failed -ne 0) { exit 1 }
exit 0
