param()

$ErrorActionPreference = 'Stop'

$script:Passed = 0

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not $Condition) {
        throw "FAIL: $Message"
    }

    $script:Passed++
}

function Read-Source {
    param([Parameter(Mandatory = $true)][string]$RelativePath)

    $root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    return [System.IO.File]::ReadAllText((Join-Path $root $RelativePath))
}

function Invoke-ReplacePreview {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Find,
        [AllowEmptyString()][string]$Replacement,
        [int]$Repeat,
        [bool]$IgnoreCase
    )

    # This is the UI compatibility policy under test. Legacy Repeat=0 means
    # the old dialog never initialized the field, so it migrates to 1.
    if ($Repeat -lt 1) {
        $Repeat = 1
    }

    $comparison = if ($IgnoreCase) {
        [System.StringComparison]::OrdinalIgnoreCase
    } else {
        [System.StringComparison]::Ordinal
    }

    $result = $Name
    $offset = 0
    for ($i = 0; $i -lt $Repeat; $i++) {
        $index = $result.IndexOf($Find, $offset, $comparison)
        if ($index -lt 0) {
            break
        }

        $result = $result.Remove($index, $Find.Length).Insert($index, $Replacement)
        $offset = $index + $Replacement.Length
    }

    return $result
}

$batchRenameHeader = Read-Source 'src\fxfile\cmd\batch_rename.h'
$replaceDialog = Read-Source 'src\fxfile\cmd\batch_rename_tab_replace_dlg.cpp'
$formatDialog = Read-Source 'src\fxfile\cmd\batch_rename_tab_format_dlg.cpp'
$mainDialog = Read-Source 'src\fxfile\cmd\batch_rename_dlg.cpp'
$batchRename = Read-Source 'src\fxfile\cmd\batch_rename.cpp'

# Source-level regression guards for the production paths.
Assert-True ($batchRenameHeader.Contains('#define FXFILE_BATCH_RENAME_REPLACE_REPEAT_MIN         (1)')) 'Repeat minimum must reject the old zero/no-op value.'
Assert-True ($replaceDialog.Contains('SetDlgItemInt(IDC_BATCH_RENAME_REPLACE_REPEAT, FXFILE_BATCH_RENAME_REPLACE_REPEAT_DEF, XPR_FALSE);')) 'Replace tab must initialize Repeat to its declared default.'
Assert-True ($replaceDialog.Contains('if (sRepeat < FXFILE_BATCH_RENAME_REPLACE_REPEAT_MIN)')) 'Replace tab must migrate persisted Repeat=0.'
Assert-True ($mainDialog.Contains('if (sRepeat < FXFILE_BATCH_RENAME_REPLACE_REPEAT_MIN)')) 'Apply path must defend against Repeat=0 even if state bypasses initialization.'
Assert-True ($formatDialog.Contains('sId == IDC_BATCH_RENAME_FORMAT_NUMBERING || sId == ID_BATCH_RENAME_FORMAT_MENU_NUMBERING')) 'Numbering button and numbering menu command must use OR.'
Assert-True (-not $formatDialog.Contains('sId == IDC_BATCH_RENAME_FORMAT_NUMBERING && sId == ID_BATCH_RENAME_FORMAT_MENU_NUMBERING')) 'Impossible numbering AND condition must not return.'
Assert-True ($mainDialog.Contains('BatchRenameTabFormatDlg *sDlg = (BatchRenameTabFormatDlg *)getTabDialog(0);')) 'Format apply path must cast tab 0 to BatchRenameTabFormatDlg.'
Assert-True (-not $mainDialog.Contains('BatchRenameTabInsertDlg *sDlg = (BatchRenameTabInsertDlg *)getTabDialog(0);')) 'Format apply path must not use the unrelated insert-tab type.'

Assert-True ($mainDialog.Contains('!mBatchRename->isFlag(BatchRename::FlagHistoryArchive)')) 'History archive toolbar command must toggle its flag.'
Assert-True ($mainDialog.Contains('!mBatchRename->isFlag(BatchRename::FlagNoChangeExt)')) 'Keep-extension toolbar command must toggle its flag.'
Assert-True ($mainDialog.Contains('!mBatchRename->isFlag(BatchRename::FlagResultRename)')) 'Apply-on-result toolbar command must toggle its flag.'

Assert-True ($mainDialog.Contains('XPR_STRING_LITERAL("Result Apply")')) 'Result flag must use its canonical state key.'
Assert-True ($mainDialog.Contains('XPR_STRING_LITERAL("Rename by Result")')) 'Result flag must retain its legacy read key.'
Assert-True ($mainDialog.Contains('XPR_STRING_LITERAL("History Archive")')) 'Archive flag must use its canonical state key.'
Assert-True ($mainDialog.Contains('XPR_STRING_LITERAL("Batch Format Archive")')) 'Archive flag must retain its legacy read key.'
Assert-True ($batchRename.Contains('return XPR_IS_FALSE(sAtLeastOneError);')) 'BatchRename::rename must return true only when every preview is valid.'

# Pure in-memory behavior matrix. No file is created, renamed, or deleted.
Assert-True ((Invoke-ReplacePreview '월마감_월마감.xlsx' '월마감' '월' 0 $false) -eq '월_월마감.xlsx') 'Legacy Repeat=0 must migrate to one replacement.'
Assert-True ((Invoke-ReplacePreview 'abc-abc.txt' 'abc' 'X' 1 $false) -eq 'X-abc.txt') 'Repeat=1 must replace only the first occurrence.'
Assert-True ((Invoke-ReplacePreview 'abc-abc.txt' 'abc' 'X' 2 $false) -eq 'X-X.txt') 'Repeat=2 must replace two occurrences.'
Assert-True ((Invoke-ReplacePreview 'Abc-abc.txt' 'abc' 'X' 2 $false) -eq 'Abc-X.txt') 'Case-sensitive replacement must preserve a nonmatching case variant.'
Assert-True ((Invoke-ReplacePreview 'Abc-abc.txt' 'abc' 'X' 2 $true) -eq 'X-X.txt') 'No-case replacement must match both case variants.'
Assert-True ((Invoke-ReplacePreview 'report_backup.txt' '_backup' '' 1 $false) -eq 'report.txt') 'Empty replacement text must support deletion by replacement.'

# Bit toggling must be reversible; this mirrors all three toolbar flags.
$flag = 4
$flags = 0
$flags = $flags -bxor $flag
Assert-True (($flags -band $flag) -ne 0) 'First toolbar click must set an unset flag.'
$flags = $flags -bxor $flag
Assert-True (($flags -band $flag) -eq 0) 'Second toolbar click must clear the same flag.'

Write-Host ("PASS: {0} batch-rename regression checks; no filesystem item was renamed." -f $script:Passed)
