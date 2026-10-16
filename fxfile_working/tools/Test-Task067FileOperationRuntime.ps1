[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$EvidenceRoot,
    [Parameter(Mandatory = $true)] [string]$AdaptiveX64,
    [Parameter(Mandatory = $true)] [string]$AdaptiveX32,
    [Parameter(Mandatory = $true)] [string]$ModernX64,
    [Parameter(Mandatory = $true)] [string]$ModernX32
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Invoke-Engine([string]$Name, [string]$Exe, [string[]]$Arguments) {
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $output = & $Exe @Arguments 2>&1 | Out-String
    $exitCode = $LASTEXITCODE
    $timer.Stop()
    return [pscustomobject]@{
        Name = $Name
        ExitCode = $exitCode
        Seconds = [math]::Round($timer.Elapsed.TotalSeconds, 6)
        Output = $output.Trim()
    }
}

function Write-PatternFile([string]$Path, [int64]$Length, [byte]$Seed) {
    $buffer = [byte[]]::new(65536)
    for ($i = 0; $i -lt $buffer.Length; $i++) {
        $buffer[$i] = [byte](($i + $Seed) -band 0xff)
    }
    $stream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        $remaining = $Length
        while ($remaining -gt 0) {
            $count = [int][math]::Min($buffer.Length, $remaining)
            $stream.Write($buffer, 0, $count)
            $remaining -= $count
        }
        $stream.Flush($true)
    }
    finally {
        $stream.Dispose()
    }
}

function Get-TreeManifest([string]$Root) {
    $fullRoot = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    $rows = @()
    foreach ($item in Get-ChildItem -LiteralPath $fullRoot -Recurse -Force | Sort-Object FullName) {
        $relative = $item.FullName.Substring($fullRoot.Length).TrimStart('\')
        if ($item.PSIsContainer) {
            $rows += [pscustomobject]@{ Path = $relative; Type = 'Directory'; Length = 0; SHA256 = '' }
        }
        else {
            $rows += [pscustomobject]@{
                Path = $relative
                Type = 'File'
                Length = $item.Length
                SHA256 = (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash
            }
        }
    }
    return $rows
}

function Assert-TreeEqual([string]$Expected, [string]$Actual, [string]$Name) {
    $left = @(Get-TreeManifest $Expected | ConvertTo-Json -Depth 3 -Compress)
    $right = @(Get-TreeManifest $Actual | ConvertTo-Json -Depth 3 -Compress)
    if (($left -join '') -cne ($right -join '')) {
        throw "$Name tree/hash comparison failed."
    }
}

$workspace = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$allowed = (Join-Path $workspace '__BUILD_TEMP_BACKUP__').TrimEnd('\') + '\'
$evidence = [IO.Path]::GetFullPath($EvidenceRoot)
if (-not $evidence.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Evidence root must stay below the workspace backup boundary: $evidence"
}
foreach ($exe in @($AdaptiveX64, $AdaptiveX32, $ModernX64, $ModernX32)) {
    if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { throw "Missing test engine: $exe" }
}

$caseRoot = Join-Path $evidence 'cases'
if (Test-Path -LiteralPath $caseRoot) { throw "Refusing to overwrite an existing test root: $caseRoot" }
New-Item -ItemType Directory -Path $caseRoot | Out-Null

$results = @()
$reportPath = Join-Path $evidence 'runtime_engine_report.json'
try {
    # Nested tree, empty directories, many small files and one medium file.
    $folderSourceParent = Join-Path $caseRoot 'folder_source'
    $bundle = Join-Path $folderSourceParent 'bundle'
    $nested = Join-Path $bundle 'nested'
    $empty = Join-Path $bundle 'empty_directory'
    $folderTarget = Join-Path $caseRoot 'folder_target'
    New-Item -ItemType Directory -Path $nested, $empty, $folderTarget | Out-Null
    for ($i = 0; $i -lt 48; $i++) {
        Write-PatternFile (Join-Path $nested ("small_{0:D3}.bin" -f $i)) (1024 + $i) ([byte]$i)
    }
    Write-PatternFile (Join-Path $bundle 'medium_8MiB.bin') (8MB) 91
    $result = Invoke-Engine 'Adaptive x64 nested/many copy' $AdaptiveX64 @('copy', $bundle, $folderTarget)
    $results += $result
    if ($result.ExitCode -ne 0) { throw "$($result.Name) failed: $($result.Output)" }
    Assert-TreeEqual $bundle (Join-Path $folderTarget 'bundle') $result.Name

    # Multi-selection small-file path and the x86 implementation.
    $multiSource = Join-Path $caseRoot 'multi_source'
    $multiTarget = Join-Path $caseRoot 'multi_target'
    New-Item -ItemType Directory -Path $multiSource, $multiTarget | Out-Null
    $multiFiles = @()
    for ($i = 0; $i -lt 64; $i++) {
        $path = Join-Path $multiSource ("multi_{0:D3}.dat" -f $i)
        Write-PatternFile $path (4096 + $i) ([byte](100 + $i))
        $multiFiles += $path
    }
    $result = Invoke-Engine 'Adaptive x32 multi-selection copy' $AdaptiveX32 (@('copymulti', $multiTarget) + $multiFiles)
    $results += $result
    if ($result.ExitCode -ne 0) { throw "$($result.Name) failed: $($result.Output)" }
    Assert-TreeEqual $multiSource $multiTarget $result.Name

    # IFileOperation copy and same-volume metadata move paths.
    $modernSource = Join-Path $caseRoot 'modern_source'
    $modernTarget = Join-Path $caseRoot 'modern_target'
    New-Item -ItemType Directory -Path $modernSource, $modernTarget | Out-Null
    $copySource = Join-Path $modernSource 'modern_copy.bin'
    Write-PatternFile $copySource (2MB) 17
    $copyHash = (Get-FileHash -LiteralPath $copySource -Algorithm SHA256).Hash
    $result = Invoke-Engine 'Modern x64 IFileOperation copy' $ModernX64 @('copy', $copySource, $modernTarget)
    $results += $result
    $copyTarget = Join-Path $modernTarget 'modern_copy.bin'
    if ($result.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $copyTarget) -or
        (Get-FileHash -LiteralPath $copyTarget -Algorithm SHA256).Hash -ne $copyHash) {
        throw "$($result.Name) failed or produced a hash mismatch: $($result.Output)"
    }

    $moveSource = Join-Path $modernSource 'modern_move.bin'
    Write-PatternFile $moveSource (1MB) 23
    $moveHash = (Get-FileHash -LiteralPath $moveSource -Algorithm SHA256).Hash
    $result = Invoke-Engine 'Modern x32 IFileOperation same-volume move' $ModernX32 @('move', $moveSource, $modernTarget)
    $results += $result
    $moveTarget = Join-Path $modernTarget 'modern_move.bin'
    if ($result.ExitCode -ne 0 -or (Test-Path -LiteralPath $moveSource) -or
        -not (Test-Path -LiteralPath $moveTarget) -or
        (Get-FileHash -LiteralPath $moveTarget -Algorithm SHA256).Hash -ne $moveHash) {
        throw "$($result.Name) failed or violated source/target integrity: $($result.Output)"
    }

    # Permanent delete paths operate only inside this evidence root.
    $adaptiveDelete = Join-Path $caseRoot 'adaptive_delete'
    New-Item -ItemType Directory -Path (Join-Path $adaptiveDelete 'child') | Out-Null
    Write-PatternFile (Join-Path $adaptiveDelete 'child\payload.bin') (1MB) 31
    $result = Invoke-Engine 'Adaptive x64 permanent folder delete' $AdaptiveX64 @('delete', $adaptiveDelete)
    $results += $result
    if ($result.ExitCode -ne 0 -or (Test-Path -LiteralPath $adaptiveDelete)) {
        throw "$($result.Name) failed: $($result.Output)"
    }

    $modernDelete = Join-Path $caseRoot 'modern_delete.bin'
    Write-PatternFile $modernDelete 65536 37
    $result = Invoke-Engine 'Modern x64 permanent file delete' $ModernX64 @('delete', $modernDelete)
    $results += $result
    if ($result.ExitCode -ne 0 -or (Test-Path -LiteralPath $modernDelete)) {
        throw "$($result.Name) failed: $($result.Output)"
    }

    # A missing source must never be reported as success.
    $missing = Join-Path $caseRoot 'does_not_exist.bin'
    $before = @(Get-ChildItem -LiteralPath $modernTarget -Force | Select-Object Name, Length | ConvertTo-Json -Compress)
    $result = Invoke-Engine 'Modern x64 missing-source failure reporting' $ModernX64 @('copy', $missing, $modernTarget)
    $results += $result
    $after = @(Get-ChildItem -LiteralPath $modernTarget -Force | Select-Object Name, Length | ConvertTo-Json -Compress)
    if ($result.ExitCode -eq 0 -or ($before -join '') -cne ($after -join '')) {
        throw "$($result.Name) was falsely reported as success or changed the target."
    }

    $report = [ordered]@{
        Result = 'PASS'
        CapturedAt = (Get-Date).ToString('o')
        EvidenceRoot = $evidence
        Cases = $results
        Verified = @(
            'x64 nested folder copy including empty directory and SHA-256 comparison',
            'x86 64-file multi-selection copy and SHA-256 comparison',
            'x64 IFileOperation copy',
            'x86 same-volume IFileOperation move',
            'adaptive permanent folder delete',
            'IFileOperation permanent file delete',
            'IFileOperation missing-source failure is not reported as success'
        )
    }
    $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $reportPath -Encoding utf8
    $report | ConvertTo-Json -Depth 5
}
catch {
    [ordered]@{
        Result = 'FAIL'
        CapturedAt = (Get-Date).ToString('o')
        Message = $_.Exception.Message
        Cases = $results
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $reportPath -Encoding utf8
    throw
}
