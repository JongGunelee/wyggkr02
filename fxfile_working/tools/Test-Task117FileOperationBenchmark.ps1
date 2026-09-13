[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$EvidenceRoot,
    [Parameter(Mandatory = $true)] [string]$AdaptiveX64,
    [ValidateRange(67108864, 2147483648)] [int64]$LargeFileBytes = 268435456
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-PatternFile([string]$Path, [int64]$Length, [byte]$Seed) {
    $buffer = [byte[]]::new(65536)
    for ($i = 0; $i -lt $buffer.Length; $i++) { $buffer[$i] = [byte](($i + $Seed) -band 255) }
    $stream = [IO.File]::Open($Path, 'CreateNew', 'Write', 'None')
    try {
        for ($remaining = $Length; $remaining -gt 0; $remaining -= $count) {
            $count = [int][Math]::Min($buffer.Length, $remaining)
            $stream.Write($buffer, 0, $count)
        }
        $stream.Flush($true)
    } finally { $stream.Dispose() }
}

function Get-TreeDigest([string]$Root) {
    $rootPath = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $rows = foreach ($item in Get-ChildItem -LiteralPath $rootPath -Recurse -Force | Sort-Object FullName) {
            $relative = $item.FullName.Substring($rootPath.Length).TrimStart('\')
            if ($item.PSIsContainer) { "D|$relative" }
            else { "F|$relative|$($item.Length)|$((Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash)" }
        }
        $bytes = [Text.Encoding]::UTF8.GetBytes(($rows -join "`n"))
        return [Convert]::ToHexString($sha.ComputeHash($bytes))
    } finally { $sha.Dispose() }
}

function Invoke-Adaptive([string]$Name, [string]$Choice, [string]$Source, [string]$Target) {
    $oldChoice = $env:FXFILE_TEST_ROBOCOPY_CHOICE
    $env:FXFILE_TEST_ROBOCOPY_CHOICE = $Choice
    $timer = [Diagnostics.Stopwatch]::StartNew()
    try { $output = & $AdaptiveX64 copy $Source $Target 2>&1 | Out-String; $exitCode = $LASTEXITCODE }
    finally { $timer.Stop(); $env:FXFILE_TEST_ROBOCOPY_CHOICE = $oldChoice }
    if ($exitCode -ne 0) { throw "$Name failed ($exitCode): $output" }
    [pscustomobject]@{ Name=$Name; Choice=$Choice; Seconds=[Math]::Round($timer.Elapsed.TotalSeconds,6); Output=$output.Trim() }
}

$workspace = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$allowed = (Join-Path $workspace '__BUILD_TEMP_BACKUP__').TrimEnd('\') + '\'
$evidence = [IO.Path]::GetFullPath($EvidenceRoot)
if (-not $evidence.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'EvidenceRoot is outside the approved task boundary.' }
if (-not (Test-Path -LiteralPath $AdaptiveX64 -PathType Leaf)) { throw "Missing probe: $AdaptiveX64" }
if (Test-Path -LiteralPath $evidence) { throw "Refusing to overwrite: $evidence" }

New-Item -ItemType Directory -Path $evidence | Out-Null
$sourceRoot = Join-Path $evidence 'source'
New-Item -ItemType Directory -Path $sourceRoot | Out-Null
$profiles = @(
    @{ Name='single'; Count=1; Bytes=65536 },
    @{ Name='many_1500'; Count=1500; Bytes=16384 },
    @{ Name='tiny_10000'; Count=10000; Bytes=4096 }
)
foreach ($profile in $profiles) {
    $dir = Join-Path $sourceRoot $profile.Name
    New-Item -ItemType Directory -Path $dir | Out-Null
    for ($i=0; $i -lt $profile.Count; $i++) {
        Write-PatternFile (Join-Path $dir ('f_{0:D5}.bin' -f $i)) $profile.Bytes ([byte]($i -band 255))
    }
}
$mixed10 = Join-Path $sourceRoot 'mixed_10'
New-Item -ItemType Directory -Path $mixed10 | Out-Null
$mixedSizes = @(0, 1, 1024, 4096, 16384, 65536, 262144, 1048576, 3145728, 8388608)
for ($i = 0; $i -lt $mixedSizes.Count; $i++) {
    Write-PatternFile (Join-Path $mixed10 ('mixed_{0:D2}.bin' -f $i)) `
        $mixedSizes[$i] ([byte](41 + $i))
}
$nested = Join-Path $sourceRoot 'nested\a\b\c'
New-Item -ItemType Directory -Path $nested, (Join-Path $sourceRoot 'nested\empty') | Out-Null
Write-PatternFile (Join-Path $nested 'payload.bin') 1048576 31
Write-PatternFile (Join-Path $sourceRoot 'large_non_sparse.bin') $LargeFileBytes 73

$results = @()
foreach ($choice in @('adaptive','robocopy')) {
    foreach ($profile in @('many_1500','tiny_10000')) {
        $target = Join-Path $evidence ("target_${choice}_${profile}")
        New-Item -ItemType Directory -Path $target | Out-Null
        $source = Join-Path $sourceRoot $profile
        $results += Invoke-Adaptive "$choice $profile" $choice $source $target
        $actual = Join-Path $target $profile
        if ((Get-TreeDigest $source) -cne (Get-TreeDigest $actual)) { throw "Digest mismatch: $choice $profile" }
    }
}

$singleTarget = Join-Path $evidence 'target_mixed'
New-Item -ItemType Directory -Path $singleTarget | Out-Null
foreach ($name in @('single','nested','large_non_sparse.bin')) {
    $results += Invoke-Adaptive "adaptive $name" 'adaptive' (Join-Path $sourceRoot $name) $singleTarget
}
$results += Invoke-Adaptive 'adaptive mixed_10' 'adaptive' $mixed10 $singleTarget
if ((Get-TreeDigest (Join-Path $sourceRoot 'single')) -cne (Get-TreeDigest (Join-Path $singleTarget 'single'))) { throw 'single digest mismatch' }
if ((Get-TreeDigest (Join-Path $sourceRoot 'nested')) -cne (Get-TreeDigest (Join-Path $singleTarget 'nested'))) { throw 'nested digest mismatch' }
if ((Get-FileHash (Join-Path $sourceRoot 'large_non_sparse.bin')).Hash -cne (Get-FileHash (Join-Path $singleTarget 'large_non_sparse.bin')).Hash) { throw 'large digest mismatch' }
if ((Get-TreeDigest $mixed10) -cne (Get-TreeDigest (Join-Path $singleTarget 'mixed_10'))) { throw 'mixed_10 digest mismatch' }

# Planning cancellation is injected only into the standalone probe.  The
# target must remain untouched because cancellation precedes every write.
$cancelTarget = Join-Path $evidence 'target_cancelled_before_write'
New-Item -ItemType Directory -Path $cancelTarget | Out-Null
$oldCancelAfter = $env:FXFILE_TEST_PLAN_CANCEL_AFTER
$env:FXFILE_TEST_PLAN_CANCEL_AFTER = '20'
$cancelTimer = [Diagnostics.Stopwatch]::StartNew()
try {
    $cancelOutput = & $AdaptiveX64 copy (Join-Path $sourceRoot 'many_1500') $cancelTarget 2>&1 | Out-String
    $cancelExit = $LASTEXITCODE
} finally {
    $cancelTimer.Stop()
    $env:FXFILE_TEST_PLAN_CANCEL_AFTER = $oldCancelAfter
}
if ($cancelExit -eq 0 -or @(Get-ChildItem -LiteralPath $cancelTarget -Force).Count -ne 0) {
    throw 'Planning cancellation was reported successful or wrote target data.'
}
$results += [pscustomobject]@{
    Name='cancel during planning'; Choice='standalone-cancel-hook'
    Seconds=[Math]::Round($cancelTimer.Elapsed.TotalSeconds,6)
    ExitCode=$cancelExit; Output=$cancelOutput.Trim(); TargetItems=0
}

# Hold the standalone probe immediately before verification, mutate its source,
# and prove that the candidate target is rolled back and source is never lost.
$changingSourceParent = Join-Path $sourceRoot 'changing_source'
$changingBundle = Join-Path $changingSourceParent 'bundle'
$changingTarget = Join-Path $evidence 'target_source_changed'
New-Item -ItemType Directory -Path $changingBundle, $changingTarget | Out-Null
$changingFile = Join-Path $changingBundle 'payload.bin'
Write-PatternFile $changingFile 33554432 119
$stdout = Join-Path $evidence 'source_changed.stdout.txt'
$stderr = Join-Path $evidence 'source_changed.stderr.txt'
$oldVerifyDelay = $env:FXFILE_TEST_BEFORE_VERIFY_DELAY_MS
$oldChoice = $env:FXFILE_TEST_ROBOCOPY_CHOICE
$env:FXFILE_TEST_BEFORE_VERIFY_DELAY_MS = '2500'
$env:FXFILE_TEST_ROBOCOPY_CHOICE = 'adaptive'
try {
    $quotedSource = '"' + $changingBundle + '"'
    $quotedTarget = '"' + $changingTarget + '"'
    $changingProcess = Start-Process -FilePath $AdaptiveX64 `
        -ArgumentList @('copy', $quotedSource, $quotedTarget) -PassThru `
        -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    $candidate = Join-Path $changingTarget 'bundle\payload.bin'
    $deadline = [DateTime]::UtcNow.AddSeconds(30)
    while (-not (Test-Path -LiteralPath $candidate -PathType Leaf) -and
           -not $changingProcess.HasExited -and [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 50
        $changingProcess.Refresh()
    }
    if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
        throw 'Source-change probe never reached the pre-verification target state.'
    }
    $stream = [IO.File]::Open($changingFile, 'Append', 'Write', 'Read')
    try { $stream.WriteByte(0x5a); $stream.Flush($true) } finally { $stream.Dispose() }
    if (-not $changingProcess.WaitForExit(30000)) {
        Stop-Process -Id $changingProcess.Id -Force
        throw 'Source-change probe did not finish after verification delay.'
    }
    $changingProcess.Refresh()
    $changingExit = $changingProcess.ExitCode
} finally {
    if (Get-Variable changingProcess -ErrorAction SilentlyContinue) {
        try {
            $changingProcess.Refresh()
            if (-not $changingProcess.HasExited) {
                Stop-Process -Id $changingProcess.Id -Force
            }
        } catch {}
    }
    $env:FXFILE_TEST_BEFORE_VERIFY_DELAY_MS = $oldVerifyDelay
    $env:FXFILE_TEST_ROBOCOPY_CHOICE = $oldChoice
}
if ($changingExit -eq 0 -or -not (Test-Path -LiteralPath $changingFile) -or
    (Test-Path -LiteralPath (Join-Path $changingTarget 'bundle'))) {
    throw 'A changing source was falsely accepted, removed, or left an unverified target.'
}
$results += [pscustomobject]@{
    Name='source changes before verification'; Choice='standalone-delay-hook'
    Seconds=$null; ExitCode=$changingExit
    Output=((Get-Content -LiteralPath $stdout -Raw) + (Get-Content -LiteralPath $stderr -Raw)).Trim()
    SourcePreserved=$true; CandidateRolledBack=$true
}

$comparisons = foreach ($profile in @('many_1500','tiny_10000')) {
    $adaptiveCase = $results | Where-Object Name -eq "adaptive $profile" | Select-Object -First 1
    $robocopyCase = $results | Where-Object Name -eq "robocopy $profile" | Select-Object -First 1
    $gain = if ($adaptiveCase.Seconds -gt 0) {
        (($adaptiveCase.Seconds - $robocopyCase.Seconds) / $adaptiveCase.Seconds) * 100
    } else { 0 }
    [pscustomobject]@{
        Profile=$profile
        AdaptiveSeconds=$adaptiveCase.Seconds
        RobocopySeconds=$robocopyCase.Seconds
        RobocopyGainPercent=[Math]::Round($gain,2)
        MeetsFifteenPercentThreshold=($gain -ge 15)
    }
}

$report = [ordered]@{
    Result='PASS'; CapturedAt=(Get-Date).ToString('o'); LargeFileBytes=$LargeFileBytes
    Profiles=$profiles; Mixed10Sizes=$mixedSizes; Cases=$results
    EngineComparisons=$comparisons
    Integrity='SHA-256 content plus relative-path and empty-directory manifest'
    FailureCoverage=@('cancel before writes','source changes before verification and rollback')
    Note='Timing includes adaptive preparation, execution and verification. Cold/warm interpretation requires repeated runs on an isolated destination; the 15% result is evidence, not an unconditional engine override.'
}
$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $evidence 'benchmark_report.json') -Encoding utf8
$report | ConvertTo-Json -Depth 6
