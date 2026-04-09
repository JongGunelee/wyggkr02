param(
    [switch]$IncludePyCache,
    [switch]$Quiet,
    [int]$RetryCount = 30,
    [int]$RetryDelayMs = 500
)

$ErrorActionPreference = "Continue"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$artifactDir = Join-Path $scriptRoot "__temp_artifacts__"
$runtimeDir = Join-Path $scriptRoot "__temp_runtime__"

$removed = 0
$locked = 0

function Remove-IfExists {
    param(
        [string]$LiteralPath,
        [switch]$Recurse
    )

    if (-not (Test-Path -LiteralPath $LiteralPath)) {
        return
    }

    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        try {
            if ($Recurse) {
                Remove-Item -LiteralPath $LiteralPath -Recurse -Force -ErrorAction Stop
            }
            else {
                Remove-Item -LiteralPath $LiteralPath -Force -ErrorAction Stop
            }
            $script:removed++
            if (-not $Quiet) { Write-Host "REMOVED: $LiteralPath" }
            return
        }
        catch {
            if ($attempt -lt $RetryCount) {
                Start-Sleep -Milliseconds $RetryDelayMs
            }
        }
    }

    $script:locked++
    if (-not $Quiet) { Write-Warning "LOCKED: $LiteralPath" }
}

if (Test-Path -LiteralPath $artifactDir) {
    Get-ChildItem -LiteralPath $artifactDir -Force | ForEach-Object {
        if ($_.Name -like "WYGGKR02_Dashboard_Agent_Setup.__tmp__*" -or $_.Name -like "WYGGKR02_Dashboard_Agent_Setup.__tmp_stage__*") {
            Remove-IfExists -LiteralPath $_.FullName -Recurse:$_.PSIsContainer
        }
    }
}

if (Test-Path -LiteralPath $runtimeDir) {
    Get-ChildItem -LiteralPath $runtimeDir -Force | ForEach-Object {
        if ($_.Name -like "*.tmp" -or $_.Name -like "temp_*" -or $_.Name -like "*.bak") {
            Remove-IfExists -LiteralPath $_.FullName -Recurse:$_.PSIsContainer
        }
    }
}

if ($IncludePyCache) {
    Get-ChildItem -LiteralPath $scriptRoot -Recurse -Directory -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -eq "__pycache__" } |
        ForEach-Object { Remove-IfExists -LiteralPath $_.FullName -Recurse }
}

if (-not $Quiet) {
    Write-Host ("[CLEANUP] removed={0}, locked={1}" -f $removed, $locked)
}
