param(
    [Parameter(Mandatory = $true)]
    [string]$EvidenceRoot
)

$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath($EvidenceRoot)
$workspace = [IO.Path]::GetFullPath((Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
if (-not $root.StartsWith($workspace, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'EvidenceRoot must stay below the FxFile workspace.'
}
if ([IO.Path]::GetPathRoot($root) -ne 'D:\') {
    throw 'Task068 runtime evidence must use D:.'
}

$source = Join-Path $root 'cases\source_tree'
$targetParent = Join-Path $root 'cases\target'
$target = Join-Path $targetParent 'source_tree'
New-Item -ItemType Directory -Path $source,$targetParent -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $source 'empty'),(Join-Path $source 'nested') -Force | Out-Null

$buffer = New-Object byte[] 257
for ($i = 0; $i -lt $buffer.Length; ++$i) { $buffer[$i] = [byte](($i * 31) % 251) }
for ($i = 0; $i -lt 1001; ++$i) {
    $dir = if (($i % 2) -eq 0) { $source } else { Join-Path $source 'nested' }
    [IO.File]::WriteAllBytes((Join-Path $dir ('item_{0:D4}.bin' -f $i)), $buffer)
}

$started = Get-Date
& "$env:SystemRoot\System32\robocopy.exe" $source $target /E /COPY:DAT /DCOPY:DAT /R:2 /W:1 /MT:4 /XJ /NP /NFL /NDL /NJH /NJS
$exitCode = $LASTEXITCODE
$elapsed = ((Get-Date) - $started).TotalSeconds
if ($exitCode -ge 8) { throw "Robocopy failed with exit code $exitCode" }

$sourceFiles = @(Get-ChildItem -LiteralPath $source -File -Recurse | Sort-Object { $_.FullName.Substring($source.Length) })
$targetFiles = @(Get-ChildItem -LiteralPath $target -File -Recurse | Sort-Object { $_.FullName.Substring($target.Length) })
if ($sourceFiles.Count -ne 1001 -or $targetFiles.Count -ne $sourceFiles.Count) {
    throw "File-count mismatch: source=$($sourceFiles.Count), target=$($targetFiles.Count)"
}
$sourceEmpty = Test-Path -LiteralPath (Join-Path $source 'empty') -PathType Container
$targetEmpty = Test-Path -LiteralPath (Join-Path $target 'empty') -PathType Container
if (-not $sourceEmpty -or -not $targetEmpty) { throw 'Empty-directory replication failed.' }

$hashMismatch = 0
for ($i = 0; $i -lt $sourceFiles.Count; ++$i) {
    if ($sourceFiles[$i].Length -ne $targetFiles[$i].Length -or
        (Get-FileHash -LiteralPath $sourceFiles[$i].FullName -Algorithm SHA256).Hash -ne
        (Get-FileHash -LiteralPath $targetFiles[$i].FullName -Algorithm SHA256).Hash) {
        ++$hashMismatch
    }
}
if ($hashMismatch -ne 0) { throw "SHA-256 mismatch count: $hashMismatch" }

$report = [ordered]@{
    Task = '068'
    Result = 'PASS'
    RobocopyExitCode = $exitCode
    ElapsedSeconds = [math]::Round($elapsed, 3)
    SourceFileCount = $sourceFiles.Count
    TargetFileCount = $targetFiles.Count
    HashMismatchCount = $hashMismatch
    EmptyDirectoryCopied = $targetEmpty
    Arguments = '/E /COPY:DAT /DCOPY:DAT /R:2 /W:1 /MT:4 /XJ /NP /NFL /NDL /NJH /NJS'
}
$reportPath = Join-Path $root 'robocopy_runtime_report.json'
$report | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $reportPath -Encoding UTF8

Remove-Item -LiteralPath (Join-Path $root 'cases') -Recurse -Force
if (Test-Path -LiteralPath (Join-Path $root 'cases')) { throw 'Fixture cleanup failed.' }
$report | Format-List
"Report: $reportPath"
