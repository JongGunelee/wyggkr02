$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
Set-Location -LiteralPath $scriptRoot

$cleanupScriptPath = Join-Path $scriptRoot 'cleanup_temp_workdirs.ps1'
if (Test-Path -LiteralPath $cleanupScriptPath) {
    & $cleanupScriptPath -IncludePyCache -Quiet
}

$artifactDir = Join-Path $scriptRoot '__temp_artifacts__'
$tempRoot = 'C:\Users\Public\Documents\ESTsoft\CreatorTemp\wyggkr02_build'
$distDir = Join-Path $tempRoot 'dist'
$buildDir = Join-Path $tempRoot 'work'
$specPath = Join-Path $scriptRoot 'dashboard_agent_launcher.spec'
$outputExe = Join-Path $artifactDir 'WYGGKR02_Dashboard_Agent_Setup.exe'
$builtExe = Join-Path $distDir 'WYGGKR02_Dashboard_Agent.exe'

if (Test-Path -LiteralPath $tempRoot) {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $artifactDir -Force | Out-Null
New-Item -ItemType Directory -Path $distDir -Force | Out-Null
New-Item -ItemType Directory -Path $buildDir -Force | Out-Null

python -m PyInstaller $specPath --noconfirm --clean --distpath $distDir --workpath $buildDir

if (-not (Test-Path -LiteralPath $builtExe)) {
    throw "빌드 결과 EXE를 찾을 수 없습니다: $builtExe"
}

Copy-Item -LiteralPath $builtExe -Destination $outputExe -Force

Write-Host "[OK] EXE build complete"
Write-Host $outputExe
