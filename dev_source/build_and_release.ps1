param(
    [switch]$PublishRepo,
    [switch]$UploadRelease,
    [string]$Owner = "JongGunelee",
    [string]$Repo = "wyggkr02",
    [string]$Tag = "WYGGKR02_Dashboard_Agent_Setup",
    [string]$AssetName = "WYGGKR02_Dashboard_Agent_Setup.zip",
    [string]$Token = "",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $scriptRoot

$cleanupScriptPath = Join-Path $scriptRoot "cleanup_temp_workdirs.ps1"
if (Test-Path -LiteralPath $cleanupScriptPath) {
    & $cleanupScriptPath -IncludePyCache -Quiet -RetryCount 30 -RetryDelayMs 500
}

if (-not $SkipBuild) {
    & (Join-Path $scriptRoot "build_dashboard_agent.ps1")
}
else {
    Write-Host "[-] SkipBuild enabled: EXE build step skipped"
}

$packageScriptPath = Join-Path $scriptRoot "package_release_zip.ps1"
$packageOutput = & $packageScriptPath
$zipOutputLine = $packageOutput | Where-Object { $_ -like "OUTPUT_ZIP=*" } | Select-Object -Last 1
$packageZipPath = ""
if ($zipOutputLine) {
    $packageZipPath = $zipOutputLine.Substring("OUTPUT_ZIP=".Length)
}
if ([string]::IsNullOrWhiteSpace($packageZipPath)) {
    $packageZipPath = Join-Path $scriptRoot "__temp_artifacts__\WYGGKR02_Dashboard_Agent_Setup.zip"
}

$canonicalZipPath = Join-Path $scriptRoot "__temp_artifacts__\WYGGKR02_Dashboard_Agent_Setup.zip"
if ((Test-Path -LiteralPath $packageZipPath) -and ($packageZipPath -ne $canonicalZipPath)) {
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        try {
            if (Test-Path -LiteralPath $canonicalZipPath) {
                Remove-Item -LiteralPath $canonicalZipPath -Force -ErrorAction Stop
            }
            Move-Item -LiteralPath $packageZipPath -Destination $canonicalZipPath -Force -ErrorAction Stop
            $packageZipPath = $canonicalZipPath
            break
        }
        catch {
            if ($attempt -lt 10) {
                Start-Sleep -Milliseconds 500
            }
        }
    }
}

if ($PublishRepo) {
    & (Join-Path $scriptRoot "publish_repo_assets.ps1") `
        -Owner $Owner `
        -Repo $Repo `
        -Token $Token
}

if ($UploadRelease) {
    & (Join-Path $scriptRoot "release_update_asset.ps1") `
        -Owner $Owner `
        -Repo $Repo `
        -Tag $Tag `
        -AssetPath $packageZipPath `
        -AssetName $AssetName `
        -Token $Token
}
else {
    Write-Host "[-] UploadRelease not set: ZIP packaging only"
}

if (Test-Path -LiteralPath $cleanupScriptPath) {
    & $cleanupScriptPath -IncludePyCache -Quiet -RetryCount 30 -RetryDelayMs 500
}

Write-Host "[OK] Build pipeline complete"
