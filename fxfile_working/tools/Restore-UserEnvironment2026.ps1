[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$sourceRoot = 'C:\00 소프트웨어\04 Fxfile\01 나의 환경\fxfile\conf'
$installRoot = 'C:\00 소프트웨어\04 Fxfile\fxfile'
$runX64Root = 'C:\Users\PC\Downloads\0000 FxFile\fxfile_run_x64\fxfile'
$runX32Root = 'C:\Users\PC\Downloads\0000 FxFile\fxfile_run_x32\fxfile'
$backupBase = 'C:\Users\PC\Downloads\0000 FxFile\__BUILD_TEMP_BACKUP__'
$roots = @($installRoot, $runX64Root, $runX32Root)

$expectedFiles = @(
    'fxfile-accel.dat',
    'fxfile-bookmark.conf',
    'fxfile-coolbar.dat',
    'fxfile-dlg_state.conf',
    'fxfile-folder_layout.conf',
    'fxfile-main.conf',
    'fxfile-toolbar.dat',
    'fxfile-updater.conf',
    'fxfile-view_set.conf',
    'fxfile.conf'
)

function Assert-ExactPath {
    param([string]$Actual, [string]$Expected)

    $actualFull = [IO.Path]::GetFullPath($Actual).TrimEnd('\')
    $expectedFull = [IO.Path]::GetFullPath($Expected).TrimEnd('\')
    if (-not $actualFull.Equals($expectedFull, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsafe path mismatch: '$actualFull' != '$expectedFull'"
    }
}

function Set-ConfigValue {
    param(
        [string]$Path,
        [string]$Key,
        [string]$Value
    )

    $encoding = [Text.UnicodeEncoding]::new($false, $true, $true)
    $text = [IO.File]::ReadAllText($Path, $encoding)
    $pattern = '(?m)^(?<prefix>[\t ]*' + [Regex]::Escape($Key) + '[\t ]*=[\t ]*)(?<value>[^\r\n]*)(?<ending>\r?)$'
    $matches = [Regex]::Matches($text, $pattern)
    if ($matches.Count -ne 1) {
        throw "Expected exactly one key '$Key' in '$Path'; found $($matches.Count)."
    }

    $updated = [Regex]::Replace(
        $text,
        $pattern,
        { param($match) $match.Groups['prefix'].Value + $Value + $match.Groups['ending'].Value },
        1
    )
    [IO.File]::WriteAllText($Path, $updated, $encoding)
}

function Test-RepairedCoolBar {
    param([byte[]]$Bytes)

    if ($Bytes.Length -ne 88) {
        throw "Unexpected coolbar length: $($Bytes.Length)"
    }
    if ([BitConverter]::ToInt32($Bytes, 0) -ne 8 -or [BitConverter]::ToInt32($Bytes, 4) -ne 4) {
        throw 'Unexpected coolbar header.'
    }

    $bookmarkEntry = $null
    for ($index = 0; $index -lt 4; $index++) {
        $offset = 8 + (20 * $index)
        $id = [BitConverter]::ToUInt32($Bytes, $offset + 4)
        if ($id -eq 54) {
            $bookmarkEntry = [pscustomobject]@{
                Width = [BitConverter]::ToUInt32($Bytes, $offset + 8)
                Style = [BitConverter]::ToUInt32($Bytes, $offset + 12)
            }
            break
        }
    }
    if ($null -eq $bookmarkEntry) {
        throw 'Bookmark band ID 54 is missing from coolbar state.'
    }
    if ($bookmarkEntry.Width -le 0) {
        throw 'Bookmark band width is not repaired.'
    }
    if (($bookmarkEntry.Style -band 0x8) -ne 0) {
        throw 'Bookmark band is marked hidden.'
    }
}

$runningFxFile = @(Get-Process -Name fxfile -ErrorAction SilentlyContinue)
if ($runningFxFile.Count -ne 0) {
    throw 'Close every fxfile.exe process before restoring the configuration.'
}

if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) {
    throw "Source environment is missing: $sourceRoot"
}
foreach ($name in $expectedFiles | Where-Object { $_ -ne 'fxfile-folder_layout.conf' }) {
    if (-not (Test-Path -LiteralPath (Join-Path $sourceRoot $name) -PathType Leaf)) {
        throw "Source file is missing: $name"
    }
}
foreach ($root in $roots) {
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
        throw "Deployment root is missing: $root"
    }
}

# The installed copy has already recalculated the legacy zero-width bands.
# Capture that validated binary before replacing any directory.
$repairedCoolBarPath = Join-Path $installRoot 'fxfile-coolbar.dat'
$repairedCoolBarBytes = [IO.File]::ReadAllBytes($repairedCoolBarPath)
Test-RepairedCoolBar -Bytes $repairedCoolBarBytes

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$backupRoot = Join-Path $backupBase ("user_environment_2026_restore_$timestamp")
New-Item -ItemType Directory -Path $backupRoot | Out-Null

$labels = @('install_x64', 'run_x64', 'run_x32')
for ($index = 0; $index -lt $roots.Count; $index++) {
    $root = $roots[$index]
    Assert-ExactPath -Actual $root -Expected $roots[$index]
    Copy-Item -LiteralPath $root -Destination (Join-Path $backupRoot $labels[$index]) -Recurse
}

$replacements = @(
    @{ File = 'fxfile-main.conf'; Key = 'main.view3.tab1.path'; Value = 'D:\02 기숙사 및 사택\05 기숙사 및 사택 월마감\26년-사택 작업' },
    @{ File = 'fxfile-main.conf'; Key = 'main.view4.tab1.path'; Value = 'D:\02 기숙사 및 사택\06 외부임차 월마감\26년-외부임차' },
    @{ File = 'fxfile.conf'; Key = 'config.view3.file_list.init_folder_path'; Value = 'D:\02 기숙사 및 사택\05 기숙사 및 사택 월마감\26년-사택 작업' },
    @{ File = 'fxfile.conf'; Key = 'config.view4.file_list.init_folder_path'; Value = 'D:\02 기숙사 및 사택\06 외부임차 월마감\26년-외부임차' },
    @{ File = 'fxfile-bookmark.conf'; Key = 'bookmark.item8.path'; Value = 'D:\02 기숙사 및 사택\08 주간회의(업무지원팀)\26년-주간회의' },
    @{ File = 'fxfile-bookmark.conf'; Key = 'bookmark.item9.path'; Value = 'D:\03 금일작업\00 월마감' },
    @{ File = 'fxfile-bookmark.conf'; Key = 'bookmark.item11.path'; Value = 'D:\02 기숙사 및 사택\03 로우데이터\26년_RawData_(기숙사 및 사택 현황).xlsx' },
    @{ File = 'fxfile-bookmark.conf'; Key = 'bookmark.item13.path'; Value = 'D:\02 기숙사 및 사택\05 기숙사 및 사택 월마감\26년-사택 작업' },
    @{ File = 'fxfile-bookmark.conf'; Key = 'bookmark.item14.path'; Value = 'D:\02 기숙사 및 사택\06 외부임차 월마감\26년-외부임차' }
)

$stageRoots = @()
try {
    for ($index = 0; $index -lt $roots.Count; $index++) {
        $root = $roots[$index]
        $parent = Split-Path -Parent $root
        $stage = Join-Path $parent ('.fxfile-config-stage-' + $timestamp + '-' + $index)
        if (Test-Path -LiteralPath $stage) {
            throw "Unexpected pre-existing stage: $stage"
        }
        New-Item -ItemType Directory -Path $stage | Out-Null
        $stageRoots += $stage

        foreach ($name in $expectedFiles | Where-Object { $_ -ne 'fxfile-folder_layout.conf' }) {
            Copy-Item -LiteralPath (Join-Path $sourceRoot $name) -Destination (Join-Path $stage $name)
        }
        # The verified historical environment predates this optional file.
        # Use the application's canonical empty/default representation so the
        # unified package inventory remains complete without importing a
        # different generation's folder-layout rules.
        $folderLayoutText = "# fxfile folder layout file`r`n`r`n[folder_layout]`r`n`r`n"
        [IO.File]::WriteAllText(
            (Join-Path $stage 'fxfile-folder_layout.conf'),
            $folderLayoutText,
            [Text.UnicodeEncoding]::new($false, $true, $true)
        )
        [IO.File]::WriteAllBytes((Join-Path $stage 'fxfile-coolbar.dat'), $repairedCoolBarBytes)

        foreach ($replacement in $replacements) {
            Set-ConfigValue -Path (Join-Path $stage $replacement.File) -Key $replacement.Key -Value $replacement.Value
        }

        $stageFiles = @(Get-ChildItem -LiteralPath $stage -File)
        if ($stageFiles.Count -ne $expectedFiles.Count) {
            throw "Unexpected file count in stage '$stage': $($stageFiles.Count)"
        }
    }

    for ($index = 0; $index -lt $roots.Count; $index++) {
        $root = $roots[$index]
        $stage = $stageRoots[$index]
        $old = $root + '.pre-2026-' + $timestamp

        Assert-ExactPath -Actual $root -Expected $roots[$index]
        Move-Item -LiteralPath $root -Destination $old
        Move-Item -LiteralPath $stage -Destination $root
        Remove-Item -LiteralPath $old -Recurse -Force
    }
}
finally {
    foreach ($stage in $stageRoots) {
        if (Test-Path -LiteralPath $stage) {
            Remove-Item -LiteralPath $stage -Recurse -Force
        }
    }
}

$audit = foreach ($root in $roots) {
    $hashes = foreach ($name in $expectedFiles) {
        $path = Join-Path $root $name
        [pscustomobject]@{
            Name = $name
            SHA256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        }
    }
    [pscustomobject]@{
        Root = $root
        FileCount = @(Get-ChildItem -LiteralPath $root -File).Count
        HasIni = Test-Path -LiteralPath (Join-Path (Split-Path -Parent $root) 'fxfile.ini')
        HasDotFxFile = Test-Path -LiteralPath (Join-Path (Split-Path -Parent $root) '.fxfile')
        Hashes = $hashes
    }
}

[pscustomobject]@{
    BackupRoot = $backupRoot
    Deployment = $audit
} | ConvertTo-Json -Depth 6
