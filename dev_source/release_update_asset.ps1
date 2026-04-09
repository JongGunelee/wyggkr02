param(
    [string]$Owner = "JongGunelee",
    [string]$Repo = "wyggkr02",
    [string]$Tag = "WYGGKR02_Dashboard_Agent_Setup",
    [string]$AssetPath = (Join-Path $PSScriptRoot "__temp_artifacts__\WYGGKR02_Dashboard_Agent_Setup.zip"),
    [string]$AssetName = "WYGGKR02_Dashboard_Agent_Setup.zip",
    [string]$Token = ""
)

$ErrorActionPreference = "Stop"

function Resolve-FullPath {
    param([string]$PathValue)
    return [System.IO.Path]::GetFullPath($PathValue)
}

if ([string]::IsNullOrWhiteSpace($Token)) {
    $Token = $env:GITHUB_TOKEN
}

if ([string]::IsNullOrWhiteSpace($Token)) {
    throw "GITHUB_TOKEN이 없습니다. 환경변수 또는 -Token으로 전달하세요."
}

$assetFullPath = Resolve-FullPath $AssetPath
if (-not (Test-Path -LiteralPath $assetFullPath)) {
    throw "업로드 파일이 없습니다: $assetFullPath"
}

$headers = @{
    Authorization = "Bearer $Token"
    "User-Agent" = "WYGGKR02-Release-Uploader"
    Accept = "application/vnd.github+json"
    "X-GitHub-Api-Version" = "2022-11-28"
}

$releaseUri = "https://api.github.com/repos/$Owner/$Repo/releases/tags/$Tag"
$release = Invoke-RestMethod -Method Get -Uri $releaseUri -Headers $headers

$existingAsset = $release.assets | Where-Object { $_.name -eq $AssetName } | Select-Object -First 1
if ($existingAsset) {
    $deleteUri = "https://api.github.com/repos/$Owner/$Repo/releases/assets/$($existingAsset.id)"
    Invoke-RestMethod -Method Delete -Uri $deleteUri -Headers $headers | Out-Null
    Write-Host ("[-] Deleted old asset id={0} ({1})" -f $existingAsset.id, $existingAsset.name)
}
else {
    Write-Host "[-] 기존 동일 에셋 없음, 신규 업로드 진행"
}

$uploadUri = "https://uploads.github.com/repos/$Owner/$Repo/releases/$($release.id)/assets?name=$([uri]::EscapeDataString($AssetName))"
$uploaded = Invoke-RestMethod -Method Post -Uri $uploadUri -Headers $headers -ContentType "application/zip" -InFile $assetFullPath

Write-Host "[OK] Release asset updated"
Write-Host ("Release: {0}" -f $release.html_url)
Write-Host ("Asset ID: {0}" -f $uploaded.id)
Write-Host ("Asset: {0} ({1} bytes)" -f $uploaded.name, $uploaded.size)
Write-Host ("Download: {0}" -f $uploaded.browser_download_url)
