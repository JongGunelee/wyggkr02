param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# This test deliberately works on strings only.  It never calls BatchRename::start,
# MultiRename, Rename-Item, Move-Item, or any other filesystem mutation API.
$script:Passed = 0
$script:Failures = [System.Collections.Generic.List[string]]::new()

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if ($Condition) {
        $script:Passed++
    }
    else {
        $script:Failures.Add($Message)
    }
}

function Assert-Equal {
    param(
        [AllowNull()]$Actual,
        [AllowNull()]$Expected,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if ([object]::Equals($Actual, $Expected)) {
        $script:Passed++
    }
    else {
        $script:Failures.Add("$Message (expected=[$Expected], actual=[$Actual])")
    }
}

function Assert-SequenceEqual {
    param(
        [Parameter(Mandatory = $true)][object[]]$Actual,
        [Parameter(Mandatory = $true)][object[]]$Expected,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if ($Actual.Count -ne $Expected.Count) {
        $script:Failures.Add("$Message (expected count=$($Expected.Count), actual count=$($Actual.Count))")
        return
    }

    for ($i = 0; $i -lt $Expected.Count; $i++) {
        if (-not [object]::Equals($Actual[$i], $Expected[$i])) {
            $script:Failures.Add("$Message (index=$i, expected=[$($Expected[$i])], actual=[$($Actual[$i])])")
            return
        }
    }

    $script:Passed++
}

function Read-Source {
    param([Parameter(Mandatory = $true)][string]$RelativePath)

    $root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    return [System.IO.File]::ReadAllText((Join-Path $root $RelativePath))
}

function Get-FunctionBody {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Signature
    )

    $start = $Source.IndexOf($Signature, [System.StringComparison]::Ordinal)
    if ($start -lt 0) {
        return ''
    }

    $brace = $Source.IndexOf('{', $start)
    if ($brace -lt 0) {
        return ''
    }

    $depth = 0
    for ($i = $brace; $i -lt $Source.Length; $i++) {
        if ($Source[$i] -eq '{') {
            $depth++
        }
        elseif ($Source[$i] -eq '}') {
            $depth--
            if ($depth -eq 0) {
                return $Source.Substring($start, $i - $start + 1)
            }
        }
    }

    return ''
}

function Split-LeafName {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [bool]$Folder = $false
    )

    if ($Folder) {
        return [pscustomobject]@{ Base = $Name; Ext = '' }
    }

    $dot = $Name.LastIndexOf('.')
    if ($dot -lt 0) {
        return [pscustomobject]@{ Base = $Name; Ext = '' }
    }

    return [pscustomobject]@{
        Base = $Name.Substring(0, $dot)
        Ext  = $Name.Substring($dot)
    }
}

function Invoke-WithExtensionPolicy {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$Operation,
        [bool]$KeepExtension = $false,
        [bool]$Folder = $false
    )

    $parts = Split-LeafName -Name $Name -Folder $Folder
    $inputName = if ($KeepExtension) { $parts.Base } else { $Name }
    $result = [string](& $Operation $inputName)

    if ($KeepExtension) {
        return $result + $parts.Ext
    }

    return $result
}

function Invoke-ReplacePreview {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Find,
        [AllowEmptyString()][string]$Replacement,
        [int]$Repeat,
        [bool]$CaseSensitive,
        [bool]$KeepExtension = $false
    )

    if ($Find.Length -eq 0) {
        throw 'The UI contract rejects an empty find string.'
    }

    # Persisted Repeat=0 came from the old uninitialized dialog.  The fixed UI
    # migrates it to the declared default of one replacement.
    if ($Repeat -lt 1) {
        $Repeat = 1
    }

    $comparison = if ($CaseSensitive) {
        [System.StringComparison]::Ordinal
    }
    else {
        [System.StringComparison]::OrdinalIgnoreCase
    }

    $operation = {
        param([string]$InputName)

        $result = $InputName
        $offset = 0
        for ($i = 0; $i -lt $Repeat; $i++) {
            $found = $result.IndexOf($Find, $offset, $comparison)
            if ($found -lt 0) {
                break
            }

            $result = $result.Remove($found, $Find.Length).Insert($found, $Replacement)
            $offset = $found + $Replacement.Length
        }

        return $result
    }.GetNewClosure()

    return Invoke-WithExtensionPolicy -Name $Name -Operation $operation -KeepExtension $KeepExtension
}

function Invoke-InsertPreview {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [ValidateSet('AtFirst', 'AtLast', 'FromFirst', 'FromLast')][string]$PositionType,
        [int]$Position,
        [Parameter(Mandatory = $true)][string]$Text,
        [bool]$KeepExtension = $false
    )

    $operation = {
        param([string]$InputName)

        switch ($PositionType) {
            'AtFirst' { return $Text + $InputName }
            'AtLast'  { return $InputName + $Text }
            'FromFirst' {
                $offset = [Math]::Max(0, [Math]::Min($Position, $InputName.Length))
                return $InputName.Insert($offset, $Text)
            }
            'FromLast' {
                $offset = if ($Position -ge $InputName.Length) {
                    0
                }
                else {
                    $InputName.Length - [Math]::Max(0, $Position)
                }
                return $InputName.Insert($offset, $Text)
            }
        }
    }.GetNewClosure()

    return Invoke-WithExtensionPolicy -Name $Name -Operation $operation -KeepExtension $KeepExtension
}

function Invoke-DeletePreview {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [ValidateSet('AtFirst', 'AtLast', 'FromFirst', 'FromLast')][string]$PositionType,
        [int]$Position,
        [int]$Length,
        [bool]$KeepExtension = $false
    )

    $operation = {
        param([string]$InputName)

        if ($Length -le 0) {
            return $InputName
        }

        $deleteLength = [Math]::Min($Length, $InputName.Length)
        switch ($PositionType) {
            'AtFirst' {
                return $InputName.Remove(0, $deleteLength)
            }
            'AtLast' {
                return $InputName.Remove($InputName.Length - $deleteLength, $deleteLength)
            }
            'FromFirst' {
                $offset = [Math]::Max(0, [Math]::Min($Position, $InputName.Length))
                $deleteLength = [Math]::Min($deleteLength, $InputName.Length - $offset)
                return $InputName.Remove($offset, $deleteLength)
            }
            'FromLast' {
                $rightRemain = [Math]::Max(0, [Math]::Min($Position, $InputName.Length))
                $deleteLength = [Math]::Min($deleteLength, $InputName.Length - $rightRemain)
                $offset = $InputName.Length - $rightRemain - $deleteLength
                return $InputName.Remove($offset, $deleteLength)
            }
        }
    }.GetNewClosure()

    return Invoke-WithExtensionPolicy -Name $Name -Operation $operation -KeepExtension $KeepExtension
}

function Convert-UpperAtFirst {
    param(
        [Parameter(Mandatory = $true)][string]$Value,
        [AllowEmptyString()][string]$SkipChars
    )

    $result = $Value.ToLowerInvariant()
    for ($i = 0; $i -lt $result.Length; $i++) {
        if ($SkipChars.IndexOf($result[$i]) -lt 0) {
            return $result.Substring(0, $i) + [char]::ToUpperInvariant($result[$i]) + $result.Substring($i + 1)
        }
    }

    return $result
}

function Convert-UpperEveryWord {
    param(
        [Parameter(Mandatory = $true)][string]$Value,
        [AllowEmptyString()][string]$SkipChars
    )

    $result = $Value.ToLowerInvariant().ToCharArray()
    $separators = $SkipChars + ' '
    $capitalize = $true
    for ($i = 0; $i -lt $result.Length; $i++) {
        if ($separators.IndexOf($result[$i]) -ge 0) {
            $capitalize = $true
        }
        elseif ($capitalize) {
            $result[$i] = [char]::ToUpperInvariant($result[$i])
            $capitalize = $false
        }
    }

    return -join $result
}

function Invoke-CasePreview {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [ValidateSet('Base', 'Extension', 'Full')][string]$Target,
        [ValidateSet('Lower', 'Upper', 'UpperFirst', 'UpperEveryWord')][string]$CaseType,
        [AllowEmptyString()][string]$SkipChars = ''
    )

    $parts = Split-LeafName -Name $Name
    $value = switch ($Target) {
        'Base'      { $parts.Base }
        'Extension' { if ($parts.Ext.Length -gt 0) { $parts.Ext.Substring(1) } else { '' } }
        'Full'      { $Name }
    }

    $converted = switch ($CaseType) {
        'Lower'          { $value.ToLowerInvariant() }
        'Upper'          { $value.ToUpperInvariant() }
        'UpperFirst'     { Convert-UpperAtFirst -Value $value -SkipChars $SkipChars }
        'UpperEveryWord' { Convert-UpperEveryWord -Value $value -SkipChars $SkipChars }
    }

    switch ($Target) {
        'Base'      { return $converted + $parts.Ext }
        'Extension' { return $parts.Base + $(if ($converted.Length -gt 0) { '.' + $converted } else { '' }) }
        'Full'      { return $converted }
    }
}

function Format-TimeValue {
    param(
        [Parameter(Mandatory = $true)][datetime]$Value,
        [Parameter(Mandatory = $true)][string]$Pattern
    )

    return $Value.ToString($Pattern, [System.Globalization.CultureInfo]::InvariantCulture)
}

function Invoke-FormatPreview {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Format,
        [int]$Index = 0,
        [bool]$KeepExtension = $false,
        [datetime]$Now = [datetime]'2026-08-13T14:05:06',
        [datetime]$Created = [datetime]'2025-01-02T03:04:05',
        [datetime]$Modified = [datetime]'2026-07-08T09:10:11'
    )

    $original = $Name
    $originalParts = Split-LeafName -Name $original
    $inputName = if ($KeepExtension) { $originalParts.Base } else { $Name }
    $inputParts = Split-LeafName -Name $inputName
    $result = [System.Text.StringBuilder]::new()

    $cursor = 0
    foreach ($match in [regex]::Matches($Format, '<[^>]+>')) {
        if ($match.Index -gt $cursor) {
            [void]$result.Append($Format.Substring($cursor, $match.Index - $cursor))
        }

        $token = $match.Value
        switch -Regex ($token) {
            '^<n>$'  { [void]$result.Append($inputParts.Base); break }
            '^<e>$'  { [void]$result.Append($inputParts.Ext); break }
            '^<\*>$' { [void]$result.Append($inputName); break }
            '^<gn>$' { [void]$result.Append($originalParts.Base); break }
            '^<ge>$' { [void]$result.Append($originalParts.Ext); break }
            '^<g\*>$' { [void]$result.Append($original); break }
            '^<(?<start>[+-]?\d+)(?<step>[+-]\d+)?>$' {
                $startText = $Matches.start
                $startNumber = [int]$startText
                $step = if ($Matches.ContainsKey('step') -and $Matches.step) { [int]$Matches.step } else { 1 }
                $number = $startNumber + ($Index * $step)
                $unsignedStart = $startText.TrimStart('+', '-')
                $zeroFilled = $unsignedStart.Length -gt 1 -and $unsignedStart[0] -eq '0'
                if ($zeroFilled) {
                    [void]$result.Append($number.ToString('D' + $unsignedStart.Length, [System.Globalization.CultureInfo]::InvariantCulture))
                }
                else {
                    [void]$result.Append($number.ToString([System.Globalization.CultureInfo]::InvariantCulture))
                }
                break
            }
            '^<time(?::(?<pattern>[^:]*))?>$' {
                $pattern = if ($Matches.ContainsKey('pattern') -and $Matches.pattern) { $Matches.pattern } else { 'yyyy-MM-dd HH.mm.ss' }
                [void]$result.Append((Format-TimeValue -Value $Now -Pattern $pattern))
                break
            }
            '^<ctime(?::(?<pattern>[^:]*))?>$' {
                $pattern = if ($Matches.ContainsKey('pattern') -and $Matches.pattern) { $Matches.pattern } else { 'yyyy-MM-dd HH.mm.ss' }
                [void]$result.Append((Format-TimeValue -Value $Created -Pattern $pattern))
                break
            }
            '^<mtime(?::(?<pattern>[^:]*))?>$' {
                $pattern = if ($Matches.ContainsKey('pattern') -and $Matches.pattern) { $Matches.pattern } else { 'yyyy-MM-dd HH.mm.ss' }
                [void]$result.Append((Format-TimeValue -Value $Modified -Pattern $pattern))
                break
            }
            default {
                throw "Unsupported/invalid short format token: $token"
            }
        }

        $cursor = $match.Index + $match.Length
    }

    if ($cursor -lt $Format.Length) {
        [void]$result.Append($Format.Substring($cursor))
    }

    $formatted = $result.ToString()
    if ($KeepExtension) {
        $formatted += $originalParts.Ext
    }

    return $formatted
}

function New-HistoryModel {
    param([Parameter(Mandatory = $true)][string[]]$Original)

    return [pscustomobject]@{
        Original = [string[]]$Original.Clone()
        Current = [string[]]$Original.Clone()
        Backward = [System.Collections.Generic.List[scriptblock]]::new()
        Forward = [System.Collections.Generic.List[scriptblock]]::new()
    }
}

function Invoke-HistoryReplay {
    param([Parameter(Mandatory = $true)]$Model)

    $values = [string[]]$Model.Original.Clone()
    foreach ($transform in $Model.Backward) {
        for ($i = 0; $i -lt $values.Count; $i++) {
            $values[$i] = [string](& $transform $values[$i] $i)
        }
    }

    $Model.Current = $values
}

function Add-HistoryTransform {
    param(
        [Parameter(Mandatory = $true)]$Model,
        [Parameter(Mandatory = $true)][scriptblock]$Transform
    )

    # A new branch after Undo must invalidate Redo, matching conventional and
    # deterministic undo-stack semantics.
    $Model.Forward.Clear()
    $Model.Backward.Add($Transform)
    Invoke-HistoryReplay -Model $Model
}

function Undo-HistoryTransform {
    param([Parameter(Mandatory = $true)]$Model)

    if ($Model.Backward.Count -eq 0) {
        return
    }

    $last = $Model.Backward[$Model.Backward.Count - 1]
    $Model.Backward.RemoveAt($Model.Backward.Count - 1)
    $Model.Forward.Add($last)
    Invoke-HistoryReplay -Model $Model
}

function Redo-HistoryTransform {
    param([Parameter(Mandatory = $true)]$Model)

    if ($Model.Forward.Count -eq 0) {
        return
    }

    $last = $Model.Forward[$Model.Forward.Count - 1]
    $Model.Forward.RemoveAt($Model.Forward.Count - 1)
    $Model.Backward.Add($last)
    Invoke-HistoryReplay -Model $Model
}

# ---------------------------------------------------------------------------
# Production source contracts: connect the in-memory model to the actual paths.
# ---------------------------------------------------------------------------
$batchRename = Read-Source 'src\fxfile\cmd\batch_rename.cpp'
$batchRenameHeader = Read-Source 'src\fxfile\cmd\batch_rename.h'
$mainDialog = Read-Source 'src\fxfile\cmd\batch_rename_dlg.cpp'
$formatDialog = Read-Source 'src\fxfile\cmd\batch_rename_tab_format_dlg.cpp'
$replaceDialog = Read-Source 'src\fxfile\cmd\batch_rename_tab_replace_dlg.cpp'
$insertSource = Read-Source 'src\fxfile\cmd\format_insert.cpp'
$deleteSource = Read-Source 'src\fxfile\cmd\format_delete.cpp'
$caseSource = Read-Source 'src\fxfile\cmd\format_case.cpp'
$nameSource = Read-Source 'src\fxfile\cmd\format_name.cpp'

Assert-True ($mainDialog.Contains('case 1: OnFormatApply();')) 'UI dispatcher must route tab 1 to Format.'
Assert-True ($mainDialog.Contains('case 2: OnReplaceApply();')) 'UI dispatcher must route tab 2 to Replace.'
Assert-True ($mainDialog.Contains('case 3: OnInsertApply();')) 'UI dispatcher must route tab 3 to Insert.'
Assert-True ($mainDialog.Contains('case 4: OnDeleteApply();')) 'UI dispatcher must route tab 4 to Delete.'
Assert-True ($mainDialog.Contains('case 5: OnCaseApply();')) 'UI dispatcher must route tab 5 to Case.'
Assert-True ($mainDialog.Contains('BatchRenameTabFormatDlg *sDlg = (BatchRenameTabFormatDlg *)getTabDialog(0);')) 'Format apply must use the Format dialog type and tab index.'
Assert-True ($mainDialog.Contains('mBatchRename->renameReplace(sOld, sNew, sRepeat, sNotCase ? XPR_FALSE : XPR_TRUE);')) 'Replace UI must pass case-sensitivity and repeat state to the core.'
Assert-True ($mainDialog.Contains('mBatchRename->renameInsert((Format::InsertPosType)sType, sPos, sInsert);')) 'Insert UI must pass type, position, and text to the core.'
Assert-True ($mainDialog.Contains('mBatchRename->renameDelete((Format::DeletePosType)sType, sPos, sLength);')) 'Delete UI must pass type, position, and length to the core.'
Assert-True ($mainDialog.Contains('mBatchRename->renameCase((Format::CaseTargetType)sType, (Format::CaseType)sCase, sSkipSpecChar);')) 'Case UI must pass target, case mode, and skip characters to the core.'
Assert-True ($batchRenameHeader.Contains('#define FXFILE_BATCH_RENAME_REPLACE_REPEAT_MIN         (1)')) 'Replace Repeat minimum must be one.'
Assert-True ($replaceDialog.Contains('if (sRepeat < FXFILE_BATCH_RENAME_REPLACE_REPEAT_MIN)')) 'Persisted Repeat=0 must be migrated on dialog load.'
Assert-True ($mainDialog.Contains('if (sRepeat < FXFILE_BATCH_RENAME_REPLACE_REPEAT_MIN)')) 'Replace apply must defensively normalize Repeat=0.'
Assert-True ($formatDialog.Contains('sId == IDC_BATCH_RENAME_FORMAT_NUMBERING || sId == ID_BATCH_RENAME_FORMAT_MENU_NUMBERING')) 'Both numbering commands must enter the numbering branch.'
Assert-True (-not $formatDialog.Contains('sId == IDC_BATCH_RENAME_FORMAT_NUMBERING && sId == ID_BATCH_RENAME_FORMAT_MENU_NUMBERING')) 'Numbering commands must not be joined by an impossible AND.'
Assert-True ($batchRename.Contains('return XPR_IS_FALSE(sAtLeastOneError);')) 'Preview status must report failure if any item is invalid.'

$baseBody = Get-FunctionBody -Source $nameSource -Signature 'void FormatBaseFileName::rename(RenameContext &aContext) const'
$extBody = Get-FunctionBody -Source $nameSource -Signature 'void FormatFileExt::rename(RenameContext &aContext) const'
$fullBody = Get-FunctionBody -Source $nameSource -Signature 'void FormatFileName::rename(RenameContext &aContext) const'
Assert-True ($baseBody.Contains('aContext.mOldFileName')) 'Format <n> must read the current sequence input, not the output cleared by FormatClear.'
Assert-True ($extBody.Contains('aContext.mOldFileName')) 'Format <e> must read the current sequence input, not the output cleared by FormatClear.'
Assert-True ($fullBody.Contains('aContext.mOldFileName')) 'Format <*> must append the current sequence input, not self-append the cleared output.'

$insertBody = Get-FunctionBody -Source $insertSource -Signature 'void FormatInsert::rename(RenameContext &aContext) const'
$insertHasUpperClamp = ($insertBody -match 'sOffset\s*>\s*\(?\s*xpr_sint_t\s*\)?\s*sLength') -or
                       ($insertBody -match 'sOffset\s*>\s*sLength') -or
                       ($insertBody -match '(?:min|std::min)\s*\([^;]*sLength')
Assert-True $insertHasUpperClamp 'Insert FromFirst must clamp a position beyond the leaf-name length instead of throwing out_of_range.'
Assert-True ($deleteSource.Contains('if (sOffset > sNameLength)')) 'Delete FromFirst/FromLast must clamp a position beyond the leaf-name length.'
Assert-True ($caseSource.Contains('CaseTargetTypeBaseFileName')) 'Case engine must support base-name targeting.'
Assert-True ($caseSource.Contains('CaseTargetTypeFileExt')) 'Case engine must support extension targeting.'
Assert-True ($caseSource.Contains('CaseTargetTypeFileName')) 'Case engine must support full-name targeting.'

$renameBody = Get-FunctionBody -Source $batchRename -Signature 'xpr_bool_t BatchRename::rename(const FileNameFormat &aFileNameFormat, xpr_bool_t aAddHistory)'
Assert-True ($renameBody.Contains('clearForward();')) 'Every new user operation must discard the stale Redo branch, even when history recording is disabled.'
Assert-True ($renameBody.Contains('FlagHistoryArchive')) 'The History Archive toolbar option must affect core history recording, not only its check mark.'
Assert-True ($mainDialog.Contains('mBatchRename->goBackward();')) 'Undo command must call the backward-history path.'
Assert-True ($mainDialog.Contains('mBatchRename->goForward();')) 'Redo command must call the forward-history path.'

# ---------------------------------------------------------------------------
# Deterministic behavior matrix.  All subjects below are in-memory strings.
# ---------------------------------------------------------------------------

# Format: current/original name, extension, numbering, date/time, and keep-ext.
Assert-Equal (Invoke-FormatPreview -Name 'Report.Final.TXT' -Format '<n>') 'Report.Final' 'Format <n> must produce the base name.'
Assert-Equal (Invoke-FormatPreview -Name 'Report.Final.TXT' -Format '<n><e>') 'Report.Final.TXT' 'Format <n><e> must reconstruct the full name.'
Assert-Equal (Invoke-FormatPreview -Name 'Report.Final.TXT' -Format '<*>') 'Report.Final.TXT' 'Format <*> must produce the full current name.'
Assert-Equal (Invoke-FormatPreview -Name 'Report.Final.TXT' -Format 'ARCHIVE_<gn><ge>') 'ARCHIVE_Report.Final.TXT' 'Original-name tokens must remain deterministic.'
Assert-Equal (Invoke-FormatPreview -Name 'Report.txt' -Format '<001+2>' -Index 0) '001' 'Zero-filled numbering must honor its start value.'
Assert-Equal (Invoke-FormatPreview -Name 'Report.txt' -Format '<001+2>' -Index 2) '005' 'Numbering must honor item index and increment.'
Assert-Equal (Invoke-FormatPreview -Name 'Report.txt' -Format '<n>_<time:yyyy-MM-dd_HH.mm.ss>' -Now ([datetime]'2026-08-13T14:05:06')) 'Report_2026-08-13_14.05.06' 'Current-time format must use the requested components.'
Assert-Equal (Invoke-FormatPreview -Name 'Report.txt' -Format '<ctime:yyyyMMdd>_<mtime:HHmmss>' -Created ([datetime]'2025-01-02T03:04:05') -Modified ([datetime]'2026-07-08T09:10:11')) '20250102_091011' 'Creation/modification time tokens must use their own timestamps.'
Assert-Equal (Invoke-FormatPreview -Name 'Report.TXT' -Format '<n>_<1>' -KeepExtension $true) 'Report_1.TXT' 'Format keep-extension mode must preserve the original extension once.'

# Replace: legacy zero, repeat count, case mode, deletion, Unicode, keep-ext.
Assert-Equal (Invoke-ReplacePreview -Name '월마감_월마감.xlsx' -Find '월마감' -Replacement '월' -Repeat 0 -CaseSensitive $true) '월_월마감.xlsx' 'Legacy Repeat=0 must migrate to one replacement.'
Assert-Equal (Invoke-ReplacePreview -Name 'abc-abc-abc.txt' -Find 'abc' -Replacement 'X' -Repeat 2 -CaseSensitive $true) 'X-X-abc.txt' 'Repeat=2 must replace exactly two matches.'
Assert-Equal (Invoke-ReplacePreview -Name 'Abc-abc.txt' -Find 'abc' -Replacement 'X' -Repeat 2 -CaseSensitive $true) 'Abc-X.txt' 'Case-sensitive replacement must ignore a different-case match.'
Assert-Equal (Invoke-ReplacePreview -Name 'Abc-abc.txt' -Find 'abc' -Replacement 'X' -Repeat 2 -CaseSensitive $false) 'X-X.txt' 'Case-insensitive replacement must match both variants.'
Assert-Equal (Invoke-ReplacePreview -Name 'report_backup.txt' -Find '_backup' -Replacement '' -Repeat 1 -CaseSensitive $true) 'report.txt' 'An empty replacement must delete the matched text.'
Assert-Equal (Invoke-ReplacePreview -Name 'report.TXT' -Find 'TXT' -Replacement 'bin' -Repeat 1 -CaseSensitive $true -KeepExtension $true) 'report.TXT' 'Keep-extension mode must shield the extension from Replace.'

# Insert: all four positions, out-of-range boundaries, and keep-ext.
Assert-Equal (Invoke-InsertPreview -Name 'abcdef.txt' -PositionType AtFirst -Position 0 -Text 'X') 'Xabcdef.txt' 'Insert AtFirst must prepend.'
Assert-Equal (Invoke-InsertPreview -Name 'abcdef.txt' -PositionType AtLast -Position 0 -Text 'X') 'abcdef.txtX' 'Insert AtLast must append.'
Assert-Equal (Invoke-InsertPreview -Name 'abcdef.txt' -PositionType FromFirst -Position 2 -Text 'X') 'abXcdef.txt' 'Insert FromFirst must use a zero-based offset.'
Assert-Equal (Invoke-InsertPreview -Name 'abcdef.txt' -PositionType FromLast -Position 2 -Text 'X') 'abcdef.tXxt' 'Insert FromLast must retain Position characters on the right.'
Assert-Equal (Invoke-InsertPreview -Name 'abc' -PositionType FromFirst -Position 999 -Text 'X') 'abcX' 'Insert FromFirst must safely clamp an oversized offset to the end.'
Assert-Equal (Invoke-InsertPreview -Name 'abc' -PositionType FromLast -Position 999 -Text 'X') 'Xabc' 'Insert FromLast must safely clamp an oversized offset to the beginning.'
Assert-Equal (Invoke-InsertPreview -Name 'report.TXT' -PositionType AtLast -Position 0 -Text '_v2' -KeepExtension $true) 'report_v2.TXT' 'Keep-extension mode must insert before the extension.'

# Delete: all four positions, length/position boundaries, and keep-ext.
Assert-Equal (Invoke-DeletePreview -Name 'abcdef' -PositionType AtFirst -Position 0 -Length 2) 'cdef' 'Delete AtFirst must remove from the head.'
Assert-Equal (Invoke-DeletePreview -Name 'abcdef' -PositionType AtLast -Position 0 -Length 2) 'abcd' 'Delete AtLast must remove from the tail.'
Assert-Equal (Invoke-DeletePreview -Name 'abcdef' -PositionType FromFirst -Position 2 -Length 2) 'abef' 'Delete FromFirst must remove at a zero-based offset.'
Assert-Equal (Invoke-DeletePreview -Name 'abcdef' -PositionType FromLast -Position 1 -Length 2) 'abcf' 'Delete FromLast must retain Position characters on the right.'
Assert-Equal (Invoke-DeletePreview -Name 'abc' -PositionType FromFirst -Position 999 -Length 2) 'abc' 'Delete FromFirst must safely no-op beyond the end.'
Assert-Equal (Invoke-DeletePreview -Name 'abc' -PositionType AtLast -Position 0 -Length 999) '' 'Delete length must clamp to the name length.'
Assert-Equal (Invoke-DeletePreview -Name 'report.TXT' -PositionType AtLast -Position 0 -Length 3 -KeepExtension $true) 'rep.TXT' 'Keep-extension mode must delete only from the base name.'

# Case: all targets and all case modes, including skip-character behavior.
Assert-Equal (Invoke-CasePreview -Name 'My Report.TXT' -Target Base -CaseType Lower) 'my report.TXT' 'Case Lower/Base must preserve extension case.'
Assert-Equal (Invoke-CasePreview -Name 'My Report.TXT' -Target Extension -CaseType Lower) 'My Report.txt' 'Case Lower/Extension must preserve base-name case.'
Assert-Equal (Invoke-CasePreview -Name 'My Report.TXT' -Target Full -CaseType Upper) 'MY REPORT.TXT' 'Case Upper/Full must transform the whole leaf name.'
Assert-Equal (Invoke-CasePreview -Name '__hELLO.TXT' -Target Base -CaseType UpperFirst -SkipChars '_') '__Hello.TXT' 'UpperFirst must skip configured leading characters.'
Assert-Equal (Invoke-CasePreview -Name '__HELLO-WORLD TEST.TXT' -Target Base -CaseType UpperEveryWord -SkipChars '_-') '__Hello-World Test.TXT' 'UpperEveryWord must treat space and configured characters as separators.'

# Undo/Redo: composition, multiple levels, boundary no-op, and branch invalidation.
$history = New-HistoryModel -Original @('Alpha_Report.TXT', 'Beta_Report.TXT')
$replaceAction = {
    param([string]$Name, [int]$Index)
    Invoke-ReplacePreview -Name $Name -Find 'Report' -Replacement 'Final' -Repeat 1 -CaseSensitive $true -KeepExtension $true
}
$insertAction = {
    param([string]$Name, [int]$Index)
    Invoke-InsertPreview -Name $Name -PositionType AtFirst -Position 0 -Text ('{0:D2}_' -f ($Index + 1)) -KeepExtension $true
}
$caseAction = {
    param([string]$Name, [int]$Index)
    Invoke-CasePreview -Name $Name -Target Base -CaseType Lower
}

Add-HistoryTransform -Model $history -Transform $replaceAction
Assert-SequenceEqual -Actual $history.Current -Expected @('Alpha_Final.TXT', 'Beta_Final.TXT') -Message 'First history operation must apply to every preview row.'
Add-HistoryTransform -Model $history -Transform $insertAction
Assert-SequenceEqual -Actual $history.Current -Expected @('01_Alpha_Final.TXT', '02_Beta_Final.TXT') -Message 'Second history operation must compose on the first.'
Undo-HistoryTransform -Model $history
Assert-SequenceEqual -Actual $history.Current -Expected @('Alpha_Final.TXT', 'Beta_Final.TXT') -Message 'Undo must remove exactly the newest operation.'
Redo-HistoryTransform -Model $history
Assert-SequenceEqual -Actual $history.Current -Expected @('01_Alpha_Final.TXT', '02_Beta_Final.TXT') -Message 'Redo must restore exactly the newest undone operation.'
Undo-HistoryTransform -Model $history
Undo-HistoryTransform -Model $history
Assert-SequenceEqual -Actual $history.Current -Expected @('Alpha_Report.TXT', 'Beta_Report.TXT') -Message 'Multiple Undo operations must restore originals.'
Undo-HistoryTransform -Model $history
Assert-SequenceEqual -Actual $history.Current -Expected @('Alpha_Report.TXT', 'Beta_Report.TXT') -Message 'Undo at the empty-history boundary must be a no-op.'
Redo-HistoryTransform -Model $history
Assert-SequenceEqual -Actual $history.Current -Expected @('Alpha_Final.TXT', 'Beta_Final.TXT') -Message 'Redo after multiple Undo operations must use LIFO order.'
Add-HistoryTransform -Model $history -Transform $caseAction
Assert-SequenceEqual -Actual $history.Current -Expected @('alpha_final.TXT', 'beta_final.TXT') -Message 'A new branch after Undo must compose from the active history.'
Assert-Equal $history.Forward.Count 0 'A new operation after Undo must clear the stale Redo branch.'
Redo-HistoryTransform -Model $history
Assert-SequenceEqual -Actual $history.Current -Expected @('alpha_final.TXT', 'beta_final.TXT') -Message 'Redo after branch invalidation must be a no-op.'

if ($script:Failures.Count -gt 0) {
    Write-Host ("FAIL: {0} passed; {1} failed; no filesystem item was created, renamed, moved, or deleted." -f $script:Passed, $script:Failures.Count)
    foreach ($failure in $script:Failures) {
        Write-Host (" - {0}" -f $failure)
    }
    exit 1
}

Write-Host ("PASS: {0} full batch-rename checks (Format/Replace/Insert/Delete/Case/Undo/Redo); no filesystem item was created, renamed, moved, or deleted." -f $script:Passed)
