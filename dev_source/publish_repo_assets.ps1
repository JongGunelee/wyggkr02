param(
    [string]$Owner = "JongGunelee",
    [string]$Repo = "wyggkr02",
    [string]$Branch = "main",
    [string]$Token = "",
    [string]$ZipPath = ""
)

$ErrorActionPreference = "Stop"

function Resolve-FullPath {
    param([string]$PathValue)
    if (Test-Path -LiteralPath $PathValue) {
        return (Get-Item -LiteralPath $PathValue).FullName
    }
    return [System.IO.Path]::GetFullPath($PathValue)
}

if ([string]::IsNullOrWhiteSpace($Token)) {
    $Token = $env:GITHUB_TOKEN
}

if ([string]::IsNullOrWhiteSpace($Token)) {
    throw "GITHUB_TOKEN이 없습니다. 환경변수 또는 -Token으로 전달하세요."
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptRoot
if ([string]::IsNullOrWhiteSpace($ZipPath)) {
    $ZipPath = Join-Path $scriptRoot "__temp_artifacts__\WYGGKR02_Dashboard_Agent_Setup.zip"
}

$headers = @{
    Authorization = "Bearer $Token"
    "User-Agent" = "WYGGKR02-Repo-Publisher"
    Accept = "application/vnd.github+json"
    "X-GitHub-Api-Version" = "2022-11-28"
}

function Convert-ToRepoApiPath {
    param([string]$RepoPath)
    $parts = $RepoPath -split "/" | Where-Object { $_ -ne "" }
    return (($parts | ForEach-Object { [uri]::EscapeDataString($_) }) -join "/")
}

function Get-RemoteMetadata {
    param([string]$RepoPath)
    $escapedPath = Convert-ToRepoApiPath $RepoPath
    $uri = "https://api.github.com/repos/$Owner/$Repo/contents/$escapedPath`?ref=$Branch"
    try {
        return Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
    }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 404) {
            return $null
        }
        throw
    }
}

function Publish-RepoFile {
    param(
        [string]$LocalPath,
        [string]$RepoPath
    )

    $localFullPath = Resolve-FullPath $LocalPath
    if (-not (Test-Path -LiteralPath $localFullPath)) {
        throw "업로드 대상 파일이 없습니다: $localFullPath"
    }

    $fileBytes = [System.IO.File]::ReadAllBytes($localFullPath)
    $contentBase64 = [Convert]::ToBase64String($fileBytes)
    $escapedPath = Convert-ToRepoApiPath $RepoPath
    $uri = "https://api.github.com/repos/$Owner/$Repo/contents/$escapedPath"
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $remote = Get-RemoteMetadata -RepoPath $RepoPath
        $payload = @{
            message = "chore: publish $RepoPath"
            content = $contentBase64
            branch  = $Branch
        }
        if ($remote -and $remote.sha) {
            $payload.sha = $remote.sha
        }

        try {
            $jsonBody = $payload | ConvertTo-Json -Depth 8
            Invoke-RestMethod -Method Put -Uri $uri -Headers $headers -Body $jsonBody | Out-Null
            Write-Host ("[OK] Published {0}" -f $RepoPath)
            return
        }
        catch {
            $statusCode = $null
            try { $statusCode = $_.Exception.Response.StatusCode.value__ } catch {}
            if (($statusCode -eq 409) -and ($attempt -lt 3)) {
                Start-Sleep -Milliseconds 400
                continue
            }
            throw
        }
    }
}

function Remove-RepoFileIfExists {
    param([string]$RepoPath)
    $remote = Get-RemoteMetadata -RepoPath $RepoPath
    if (-not $remote -or $remote.type -ne "file") {
        return
    }

    $escapedPath = Convert-ToRepoApiPath $RepoPath
    $uri = "https://api.github.com/repos/$Owner/$Repo/contents/$escapedPath"
    $payload = @{
        message = "chore: remove legacy path $RepoPath"
        sha = $remote.sha
        branch = $Branch
    } | ConvertTo-Json
    Invoke-RestMethod -Method Delete -Uri $uri -Headers $headers -Body $payload | Out-Null
    Write-Host ("[-] Removed legacy file {0}" -f $RepoPath)
}

function Publish-RepoTree {
    param(
        [string]$LocalRoot,
        [string]$RepoRoot
    )

    $localRootFullPath = Resolve-FullPath $LocalRoot
    if (-not (Test-Path -LiteralPath $localRootFullPath)) {
        throw "업로드 대상 폴더가 없습니다: $localRootFullPath"
    }

    Get-ChildItem -LiteralPath $localRootFullPath -Recurse -File |
        Sort-Object FullName |
        ForEach-Object {
            $relativePath = $_.FullName.Substring($localRootFullPath.Length).TrimStart('\')
            $repoPath = (($RepoRoot.TrimEnd('/')) + "/" + ($relativePath -replace "\\", "/")).TrimStart("/")
            Publish-RepoFile -LocalPath $_.FullName -RepoPath $repoPath
        }
}

$singleFiles = @(
    "00 dashboard.html",
    "index.html",
    "manifest.webmanifest",
    "service-worker.js",
    "dev_source/run_dashboard.py",
    "dev_source/dashboard_agent_launcher.py",
    "dev_source/dashboard_agent_launcher.spec",
    "dev_source/build_dashboard_agent.ps1",
    "dev_source/package_release_zip.ps1",
    "dev_source/release_update_asset.ps1",
    "dev_source/build_and_release.ps1",
    "dev_source/publish_repo_assets.ps1",
    "dev_source/cleanup_temp_workdirs.ps1",
    "dev_source/000 Launch_dashboard.bat"
)

foreach ($file in $singleFiles) {
    Publish-RepoFile -LocalPath (Join-Path $projectRoot $file) -RepoPath ($file -replace "\\", "/")
}

$readmePath = Join-Path $projectRoot "README.md"
if (Test-Path -LiteralPath $readmePath) {
    Publish-RepoFile -LocalPath $readmePath -RepoPath "README.md"
}

$webUrlFile = Get-ChildItem -LiteralPath $scriptRoot -File | Where-Object { $_.Extension -eq '.txt' } | Select-Object -First 1
if ($webUrlFile) {
    Publish-RepoFile -LocalPath $webUrlFile.FullName -RepoPath ('dev_source/' + $webUrlFile.Name)
}
else {
    throw "dev_source/웹접속 주소.txt 파일을 찾지 못했습니다."
}

$handoffDir = Get-ChildItem -LiteralPath $scriptRoot -Directory | Where-Object { $_.Name -like '__01*' } | Select-Object -First 1
if (-not $handoffDir) {
    throw "진행현황 폴더를 찾지 못했습니다."
}
$handoffFile = Get-ChildItem -LiteralPath $handoffDir.FullName -File | Where-Object { $_.Extension -eq '.md' -and $_.Name -like '04 *' } | Sort-Object Name -Descending | Select-Object -First 1
if ($handoffFile) {
    Publish-RepoFile -LocalPath $handoffFile.FullName -RepoPath ('dev_source/' + $handoffDir.Name + '/' + $handoffFile.Name)
}
else {
    throw "인수인계서 파일을 찾지 못했습니다."
}

Publish-RepoTree -LocalRoot (Join-Path $projectRoot "automated_scripts") -RepoRoot "automated_scripts"
Publish-RepoTree -LocalRoot (Join-Path $projectRoot "automated_app") -RepoRoot "automated_app"
Publish-RepoTree -LocalRoot (Join-Path $projectRoot "runtime_store") -RepoRoot "runtime_store"
Publish-RepoTree -LocalRoot (Join-Path $projectRoot "system_guides") -RepoRoot "system_guides"
Publish-RepoTree -LocalRoot (Join-Path $projectRoot "dev_source\\runtime_store") -RepoRoot "dev_source/runtime_store"
Publish-RepoFile -LocalPath $ZipPath -RepoPath "dev_source/runtime_store/WYGGKR02_Dashboard_Agent_Setup.zip"
Remove-RepoFileIfExists -RepoPath "WYGGKR02_Dashboard_Agent_Setup.zip"

Write-Host "[OK] Repository publish complete"
