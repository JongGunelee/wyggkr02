param(
    [string]$ExePath = (Join-Path $PSScriptRoot "__temp_artifacts__\WYGGKR02_Dashboard_Agent_Setup.exe"),
    [string]$ZipPath = (Join-Path $PSScriptRoot "__temp_artifacts__\WYGGKR02_Dashboard_Agent_Setup.zip"),
    [int]$RetryCount = 12,
    [int]$RetryDelayMs = 750
)

$ErrorActionPreference = "Stop"

function Resolve-FullPath {
    param([string]$PathValue)
    return [System.IO.Path]::GetFullPath($PathValue)
}

$exeFullPath = Resolve-FullPath $ExePath
if (-not (Test-Path -LiteralPath $exeFullPath)) {
    throw "EXE 파일을 찾을 수 없습니다: $exeFullPath"
}

$zipFullPath = Resolve-FullPath $ZipPath
$zipDir = Split-Path -Parent $zipFullPath
if (-not (Test-Path -LiteralPath $zipDir)) {
    New-Item -ItemType Directory -Path $zipDir -Force | Out-Null
}

$cleanupScriptPath = Join-Path $PSScriptRoot "cleanup_temp_workdirs.ps1"
if (Test-Path -LiteralPath $cleanupScriptPath) {
    & $cleanupScriptPath -Quiet
}

$tmpZipName = "{0}.__tmp__.{1}.zip" -f ([System.IO.Path]::GetFileNameWithoutExtension($zipFullPath)), ([Guid]::NewGuid().ToString("N"))
$tmpZipPath = Join-Path $zipDir $tmpZipName
$stageDirName = "{0}.__tmp_stage__.{1}" -f ([System.IO.Path]::GetFileNameWithoutExtension($zipFullPath)), ([Guid]::NewGuid().ToString("N"))
$stageDirPath = Join-Path $zipDir $stageDirName
$stagedExePath = Join-Path $stageDirPath ([System.IO.Path]::GetFileName($exeFullPath))

if (Test-Path -LiteralPath $tmpZipPath) {
    Remove-Item -LiteralPath $tmpZipPath -Force
}

if (Test-Path -LiteralPath $stageDirPath) {
    Remove-Item -LiteralPath $stageDirPath -Recurse -Force
}
New-Item -ItemType Directory -Path $stageDirPath -Force | Out-Null

$copied = $false
for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
    try {
        Copy-Item -LiteralPath $exeFullPath -Destination $stagedExePath -Force
        $copied = $true
        break
    }
    catch {
        if ($attempt -ge $RetryCount) {
            throw "Failed to copy EXE (locked): $exeFullPath`nError: $($_.Exception.Message)"
        }
        Start-Sleep -Milliseconds $RetryDelayMs
    }
}

if (-not $copied) {
    throw "Failed to copy EXE: $exeFullPath"
}

Compress-Archive -LiteralPath $stagedExePath -DestinationPath $tmpZipPath -CompressionLevel Optimal -Force

$stageRemoved = $false
for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
    try {
        if (Test-Path -LiteralPath $stageDirPath) {
            Remove-Item -LiteralPath $stageDirPath -Recurse -Force
        }
        $stageRemoved = $true
        break
    }
    catch {
        if ($attempt -ge $RetryCount) {
            break
        }
        Start-Sleep -Milliseconds $RetryDelayMs
    }
}
if (-not $stageRemoved) {
    Write-Warning "Temp stage cleanup skipped (locked): $stageDirPath"
}

$replaced = $false
for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
    try {
        if (Test-Path -LiteralPath $zipFullPath) {
            Remove-Item -LiteralPath $zipFullPath -Force
        }
        Move-Item -LiteralPath $tmpZipPath -Destination $zipFullPath -Force
        $replaced = $true
        break
    }
    catch {
        if ($attempt -ge $RetryCount) {
            break
        }
        Start-Sleep -Milliseconds $RetryDelayMs
    }
}

if (-not $replaced) {
    $fallbackZipPath = [System.IO.Path]::Combine($zipDir, "{0}.__new__.zip" -f [System.IO.Path]::GetFileNameWithoutExtension($zipFullPath))
    $movedToFallback = $false
    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        try {
            if (Test-Path -LiteralPath $fallbackZipPath) {
                Remove-Item -LiteralPath $fallbackZipPath -Force
            }
            Move-Item -LiteralPath $tmpZipPath -Destination $fallbackZipPath -Force
            $movedToFallback = $true
            break
        }
        catch {
            if ($attempt -ge $RetryCount) {
                break
            }
            Start-Sleep -Milliseconds $RetryDelayMs
        }
    }
    if ($movedToFallback) {
        $zipFullPath = $fallbackZipPath
        Write-Warning "Target ZIP was locked. Wrote fallback file: $fallbackZipPath"
    }
    else {
        $zipFullPath = $tmpZipPath
        Write-Warning "Target ZIP was locked. Using temp ZIP directly: $tmpZipPath"
    }
    $replaced = $true
}

$zipFile = Get-Item -LiteralPath $zipFullPath
$zipHash = Get-FileHash -Algorithm SHA256 -LiteralPath $zipFullPath

Write-Host "[OK] ZIP packaging complete"
Write-Host $zipFile.FullName
Write-Host ("Size: {0} bytes" -f $zipFile.Length)
Write-Host ("SHA256: {0}" -f $zipHash.Hash)
Write-Output ("OUTPUT_ZIP={0}" -f $zipFile.FullName)
