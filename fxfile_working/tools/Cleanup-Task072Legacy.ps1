[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$workspace = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..')).TrimEnd('\')
$buildRoot = Join-Path $workspace '__BUILD_TEMP_BACKUP__'
$backupRoot = Join-Path $workspace '__BACKUP_보존용__'
$keepEvidence = @(
    'unified_deploy_20260827_083131_195',
    'preflight_20260827_082943_577',
    'task072_manual_column_drag_stable_x64_20260827.json'
)

$targets = [Collections.Generic.List[string]]::new()
Get-ChildItem -LiteralPath $buildRoot -Force |
    Where-Object { $_.Name -notin $keepEvidence } |
    ForEach-Object { $targets.Add($_.FullName) }

@(
    (Join-Path $workspace 'fxfile_working\build_cmake'),
    (Join-Path $workspace 'fxfile_working\build_cmake_x32'),
    (Join-Path $workspace 'fxfile_working\bin'),
    (Join-Path $workspace 'fxfile_working\obj'),
    (Join-Path $backupRoot 'fxfile_dev'),
    (Join-Path $backupRoot 'fxfile_working_source_Task060_20260814.zip'),
    (Join-Path $backupRoot 'fxfile_run_x64_Backup(레거시 64bit 빌드)'),
    (Join-Path $backupRoot 'CHANGELOG_HISTORY-1차_bak.md'),
    (Join-Path $backupRoot 'fxfile_original_backup\build')
) | Where-Object { Test-Path -LiteralPath $_ } | ForEach-Object { $targets.Add($_) }

# Remove only byte-identical '(1)' copies from the preserved original tree.
$originalBackup = Join-Path $backupRoot 'fxfile_original_backup'
Get-ChildItem -LiteralPath $originalBackup -Recurse -File -Force |
    Where-Object { $_.BaseName -match '\(1\)$' } |
    ForEach-Object {
        $baseName = $_.BaseName -replace '\(1\)$', ''
        $original = Join-Path $_.DirectoryName ($baseName + $_.Extension)
        if (Test-Path -LiteralPath $original -PathType Leaf) {
            $identical = $_.Length -eq (Get-Item -LiteralPath $original).Length -and
                (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash -eq
                (Get-FileHash -LiteralPath $original -Algorithm SHA256).Hash
            if ($identical) { $targets.Add($_.FullName) }
        }
    }

$uniqueTargets = @($targets | Sort-Object -Unique)
$failures = [Collections.Generic.List[string]]::new()
$deleted = 0
$deletedBytes = [int64]0

foreach ($target in $uniqueTargets) {
    $full = [IO.Path]::GetFullPath($target)
    if (-not $full.StartsWith($workspace + '\', [StringComparison]::OrdinalIgnoreCase) -or
        $full -eq $workspace -or $full -eq $buildRoot -or $full -eq $backupRoot) {
        throw "Unsafe cleanup target rejected: $full"
    }
    if (-not (Test-Path -LiteralPath $full)) { continue }

    $item = Get-Item -LiteralPath $full -Force
    $bytes = if ($item.PSIsContainer) {
        [int64]((Get-ChildItem -LiteralPath $full -Recurse -File -Force -ErrorAction SilentlyContinue |
            Measure-Object Length -Sum).Sum)
    } else { [int64]$item.Length }

    try {
        Remove-Item -LiteralPath $full -Recurse -Force -ErrorAction Stop
        $deleted++
        $deletedBytes += $bytes
    }
    catch {
        $failures.Add("$full :: $($_.Exception.Message)")
    }
}

$result = [ordered]@{
    Result = if ($failures.Count -eq 0) { 'PASS' } else { 'PARTIAL' }
    RequestedTargets = $uniqueTargets.Count
    DeletedTargets = $deleted
    DeletedBytes = $deletedBytes
    DeletedMiB = [math]::Round($deletedBytes / 1MB, 2)
    Failures = @($failures)
    LatestDeploymentKept = Test-Path -LiteralPath (Join-Path $buildRoot 'unified_deploy_20260827_083131_195\deployment_manifest.json')
    LatestPreflightKept = Test-Path -LiteralPath (Join-Path $buildRoot 'preflight_20260827_082943_577\preflight_report.json')
    Task072EvidenceKept = Test-Path -LiteralPath (Join-Path $buildRoot 'task072_manual_column_drag_stable_x64_20260827.json')
    OriginalSourceBackupKept = Test-Path -LiteralPath (Join-Path $backupRoot 'fxfile_original_backup\src')
}

$result | ConvertTo-Json -Depth 5
if ($failures.Count -gt 0) { exit 2 }
