param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$cpp = [System.IO.File]::ReadAllText((Join-Path $root 'src\fxfile\cmd\multi_rename.cpp'))
$header = [System.IO.File]::ReadAllText((Join-Path $root 'src\fxfile\cmd\multi_rename.h'))
$batchCpp = [System.IO.File]::ReadAllText((Join-Path $root 'src\fxfile\cmd\batch_rename.cpp'))
$dialogCpp = [System.IO.File]::ReadAllText((Join-Path $root 'src\fxfile\cmd\batch_rename_dlg.cpp'))

$passed = 0
$failed = [System.Collections.Generic.List[string]]::new()

function Assert-Source {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][string]$Pattern,
        [Parameter(Mandatory = $true)][string]$Description
    )

    if ([regex]::IsMatch($Text, $Pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
        $script:passed++
    }
    else {
        $script:failed.Add($Description)
    }
}

function Assert-NotSource {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][string]$Pattern,
        [Parameter(Mandatory = $true)][string]$Description
    )

    if (-not [regex]::IsMatch($Text, $Pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
        $script:passed++
    }
    else {
        $script:failed.Add($Description)
    }
}

Assert-Source $header 'FlagReadOnlyRename\s*=\s*\(1u\s*<<\s*0\)' `
    'FlagReadOnlyRename must be a non-zero bit mask.'
Assert-Source $cpp 'mFlags\(FlagReadOnlyRename\)' `
    'The valid flag must default on to preserve the legacy no-UI behaviour.'
Assert-Source $cpp 'mPreparedCount\s*=\s*0;.*mValidatedCount\s*=\s*0;.*mRenamedCount\s*=\s*0;' `
    'start() must reset every progress counter.'
Assert-Source $cpp 'kMaxBackupCandidateCount\s*=\s*1000' `
    'Backup candidate generation must have a documented finite cap.'
Assert-Source $cpp 'sAttempt\s*<\s*kMaxBackupCandidateCount\s*&&\s*XPR_IS_FALSE\(mThread\.isStop\(\)\)' `
    'Backup probing must be bounded and cancellation-aware.'
Assert-NotSource $cpp 'do\s*\{.*?MoveFile\(sDst\.c_str\(\),\s*sTemp\.c_str\(\)\).*?\}\s*while' `
    'The former unbounded destination backup loop must not return.'
Assert-Source $cpp 'sTemp\.length\(\)\s*>=\s*XPR_MAX_PATH' `
    'Each generated backup path must be length checked.'
Assert-Source $cpp 'isDestinationCollisionError\(sDirectMoveError\)' `
    'Only an actual destination collision may displace the destination.'
Assert-Source $cpp 'MoveFile\(sTemp\.c_str\(\),\s*sDst\.c_str\(\)\)' `
    'A failed source rename must attempt destination rollback.'
Assert-Source $cpp 'rollback itself fails.*?sRenItem2->mOld\s*=\s*sTempFileName' `
    'Rollback failure must update the pending item to its real path.'
Assert-Source $cpp 'GetFileAttributes\(sSrc\.c_str\(\)\)\s*!=\s*INVALID_FILE_ATTRIBUTES.*?SetFileAttributes\(sSrc\.c_str\(\),\s*sAttributes\)' `
    'A failed rename must restore the source read-only attributes.'
Assert-Source $cpp 'getMoveErrorResult\(sDirectMoveError,\s*sAttributes\)' `
    'Direct MoveFile failures must be reported instead of becoming success.'
Assert-Source $cpp 'getMoveErrorResult\(sSecondMoveError,\s*sAttributes\)' `
    'Post-backup source failures must be reported instead of becoming success.'
Assert-Source $cpp 'ERROR_SHARING_VIOLATION.*?ResultShared' `
    'Sharing violations must be distinguishable from unknown failures.'
Assert-Source $header 'StatusRenameCompleted,\s*StatusRenameFailed,' `
    'A filesystem failure must have a status distinct from completion.'
Assert-Source $cpp 'sRenItem->mResult\s*!=\s*ResultSucceeded.*?sRenameFailed\s*=\s*XPR_TRUE' `
    'Every per-item filesystem failure must be aggregated.'
Assert-Source $cpp 'sStatus\s*=\s*StatusRenameFailed' `
    'A failed operation must not be reported as StatusRenameCompleted.'
Assert-Source $cpp 'MoveFile\(sSrc\.c_str\(\),\s*sDst\.c_str\(\)\)\s*==\s*XPR_TRUE.*?sRenItem->mOld\s*=\s*sRenItem->mNew' `
    'Every successful physical rename must update the worker old-name model.'
Assert-Source $batchCpp 'void\s+BatchRename::syncRenamedItems.*?getItemOldName.*?sItem->mOld\s*=\s*sActualOldName' `
    'The preview model must be rebased to physical names after a partial operation.'
Assert-Source $dialogCpp 'StatusRenameFailed.*?StatusStopped.*?syncRenamedItems\(\)' `
    'Failure and stop finalization must synchronize names before retry is enabled.'

if ($failed.Count -ne 0) {
    Write-Host "FAIL: $($failed.Count) contract(s), $passed passed"
    foreach ($failure in $failed) {
        Write-Host " - $failure"
    }
    exit 1
}

Write-Host "PASS: $passed MultiRename safety contracts"
exit 0
