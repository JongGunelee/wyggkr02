[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet('Optimize', 'Restore')]
    [string]$Mode = 'Optimize',

    [string]$BackupPath = (Join-Path $PSScriptRoot 'fxfile-appcompat-backup.json')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$targetExe = 'C:\00 소프트웨어\04 Fxfile\fxfile.exe'
$layersKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers'
$pcaStoreKey = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Store'
$optimizedValue = '~ DISABLEDXMAXIMIZEDWINDOWEDMODE'

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]::new($identity)
$isAdministrator = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ($Mode -eq 'Optimize') {
    $currentValue = (Get-ItemProperty -LiteralPath $layersKey -Name $targetExe -ErrorAction Stop).$targetExe
    $originalValue = $currentValue
    if (Test-Path -LiteralPath $BackupPath -PathType Leaf) {
        $existingBackup = Get-Content -LiteralPath $BackupPath -Raw | ConvertFrom-Json
        if ($existingBackup.ValueName -eq $targetExe -and -not [string]::IsNullOrWhiteSpace($existingBackup.ValueData)) {
            $originalValue = [string]$existingBackup.ValueData
        }
    }

    $pcaStoreValue = $null
    if (Test-Path -LiteralPath $pcaStoreKey) {
        try {
            $pcaStoreValue = (Get-ItemProperty -LiteralPath $pcaStoreKey -Name $targetExe -ErrorAction Stop).$targetExe
        }
        catch {
            $pcaStoreValue = $null
        }
    }

    $backup = [ordered]@{
        CapturedAt = (Get-Date).ToString('o')
        RegistryKey = $layersKey
        ValueName = $targetExe
        ValueData = $originalValue
        OptimizedValueData = $optimizedValue
        PcaStoreKey = $pcaStoreKey
        PcaStoreValueBase64 = if ($null -ne $pcaStoreValue) { [Convert]::ToBase64String([byte[]]$pcaStoreValue) } else { $null }
    }

    $backupDirectory = Split-Path -Parent $BackupPath
    if (-not [string]::IsNullOrWhiteSpace($backupDirectory)) {
        New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
    }
    $backup | ConvertTo-Json | Set-Content -LiteralPath $BackupPath -Encoding UTF8

    if ($currentValue -ne $optimizedValue) {
        if (-not $isAdministrator) {
            throw 'This exact HKLM compatibility setting requires an elevated PowerShell session.'
        }

        if ($PSCmdlet.ShouldProcess($targetExe, "Replace AppCompat '$currentValue' with '$optimizedValue'")) {
            Set-ItemProperty -LiteralPath $layersKey -Name $targetExe -Value $optimizedValue -Type String
        }
    }

    if ($null -ne $pcaStoreValue -and $PSCmdlet.ShouldProcess($targetExe, 'Remove the user PCA compatibility cache entry')) {
        Remove-ItemProperty -LiteralPath $pcaStoreKey -Name $targetExe -ErrorAction Stop
    }
}
else {
    if (-not $isAdministrator) {
        throw 'Restoring the exact HKLM compatibility setting requires an elevated PowerShell session.'
    }

    if (-not (Test-Path -LiteralPath $BackupPath -PathType Leaf)) {
        throw "Compatibility backup not found: $BackupPath"
    }

    $backup = Get-Content -LiteralPath $BackupPath -Raw | ConvertFrom-Json
    if ($backup.ValueName -ne $targetExe) {
        throw "Compatibility backup targets a different executable: $($backup.ValueName)"
    }

    if ($PSCmdlet.ShouldProcess($targetExe, "Restore AppCompat '$($backup.ValueData)'")) {
        Set-ItemProperty -LiteralPath $layersKey -Name $targetExe -Value ([string]$backup.ValueData) -Type String
    }

    if ($null -ne $backup.PcaStoreValueBase64 -and
        -not [string]::IsNullOrWhiteSpace([string]$backup.PcaStoreValueBase64) -and
        $PSCmdlet.ShouldProcess($targetExe, 'Restore the user PCA compatibility cache entry')) {
        $pcaStoreBytes = [Convert]::FromBase64String([string]$backup.PcaStoreValueBase64)
        New-Item -Path $pcaStoreKey -Force | Out-Null
        Set-ItemProperty -LiteralPath $pcaStoreKey -Name $targetExe -Value $pcaStoreBytes -Type Binary
    }
}

$verifiedValue = (Get-ItemProperty -LiteralPath $layersKey -Name $targetExe -ErrorAction Stop).$targetExe
[pscustomobject]@{
    Mode = $Mode
    Target = $targetExe
    AppCompat = $verifiedValue
    PcaStoreEntryPresent = $null -ne (Get-ItemProperty -LiteralPath $pcaStoreKey -Name $targetExe -ErrorAction SilentlyContinue)
    Backup = $BackupPath
}
