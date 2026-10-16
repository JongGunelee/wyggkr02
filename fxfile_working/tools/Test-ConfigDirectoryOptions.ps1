[CmdletBinding()]
param(
    [ValidateSet('Prepare', 'Audit', 'Restore')]
    [string]$Mode,

    [Parameter(Mandatory = $true)]
    [string]$StateRoot,

    [ValidateSet('Program', 'AppData', 'Custom', 'Any')]
    [string]$Expected = 'Any'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$workspaceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$sourcePackage = Join-Path $workspaceRoot 'fxfile_run_x64'
$sandboxRoot = Join-Path $StateRoot 'sandbox_x64'
$sandboxConfig = Join-Path $sandboxRoot 'fxfile'
$customConfig = Join-Path $StateRoot 'custom_config'
$appDataRoot = [IO.Path]::GetFullPath((Join-Path $env:APPDATA 'fxfile'))
$expectedAppDataRoot = [IO.Path]::GetFullPath("C:\Users\$env:USERNAME\AppData\Roaming\fxfile")
$appDataConfig = Join-Path $appDataRoot 'conf'
$statePath = Join-Path $StateRoot 'state.json'
$evidenceRoot = Join-Path $StateRoot 'evidence'
$canonicalNames = @(
    'fxfile-accel.dat',
    'fxfile-bookmark.conf',
    'fxfile-coolbar.dat',
    'fxfile-dlg_state.conf',
    'fxfile-folder_layout.conf',
    'fxfile-main.conf',
    'fxfile-toolbar.dat',
    'fxfile-updater.conf',
    'fxfile-view_set.conf',
    'fxfile.conf'
)

function Assert-NoFxFileProcess {
    $processes = @(Get-Process fxfile, 'fxfile-launcher', 'fxfile-upchecker', 'fxfile-updater' -ErrorAction SilentlyContinue)
    if ($processes.Count -ne 0) {
        throw "Close all FxFile-related processes first: $($processes.Id -join ', ')"
    }
}

function Get-Inventory([string]$Path) {
    $result = [ordered]@{}
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        return $result
    }
    foreach ($file in Get-ChildItem -LiteralPath $Path -Recurse -Force -File | Sort-Object FullName) {
        $relative = $file.FullName.Substring($Path.TrimEnd('\').Length + 1)
        $result[$relative] = [ordered]@{
            Length = $file.Length
            SHA256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
            LastWriteTimeUtc = $file.LastWriteTimeUtc.ToString('o')
        }
    }
    return $result
}

function Get-CanonicalInventory([string]$Path) {
    $result = [ordered]@{}
    foreach ($name in $canonicalNames) {
        $file = Join-Path $Path $name
        if (Test-Path -LiteralPath $file -PathType Leaf) {
            $result[$name] = [ordered]@{
                Length = (Get-Item -LiteralPath $file).Length
                SHA256 = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
            }
        }
    }
    return $result
}

function Get-PointerText([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $null
    }
    return (Get-Content -LiteralPath $Path -Raw)
}

function Get-AuditObject {
    $local = Get-CanonicalInventory $sandboxConfig
    $appData = Get-CanonicalInventory $appDataConfig
    $custom = Get-CanonicalInventory $customConfig
    $localIni = Join-Path $sandboxRoot 'fxfile.ini'
    $localDot = Join-Path $sandboxRoot '.fxfile'
    $appDataPointer = Join-Path $appDataRoot '.fxfile'

    $selected = 'Unknown'
    if ($local.Contains('fxfile.conf') -and $local.Contains('fxfile-main.conf')) {
        $selected = 'Program'
    }
    elseif ($custom.Contains('fxfile.conf') -and $custom.Contains('fxfile-main.conf')) {
        $selected = 'Custom'
    }
    elseif ($appData.Contains('fxfile.conf') -and $appData.Contains('fxfile-main.conf')) {
        $selected = 'AppData'
    }

    return [ordered]@{
        Timestamp = (Get-Date).ToString('o')
        Expected = $Expected
        InferredActiveLocation = $selected
        SandboxRoot = $sandboxRoot
        ProgramConfigPath = $sandboxConfig
        AppDataConfigPath = $appDataConfig
        CustomConfigPath = $customConfig
        ProgramCanonicalCount = $local.Count
        AppDataCanonicalCount = $appData.Count
        CustomCanonicalCount = $custom.Count
        ProgramInventory = $local
        AppDataInventory = $appData
        CustomInventory = $custom
        LocalFxFileIniExists = Test-Path -LiteralPath $localIni -PathType Leaf
        LocalDotFxFileExists = Test-Path -LiteralPath $localDot -PathType Leaf
        LocalFxFileIniText = Get-PointerText $localIni
        LocalDotFxFileText = Get-PointerText $localDot
        AppDataPointerExists = Test-Path -LiteralPath $appDataPointer -PathType Leaf
        AppDataPointerText = Get-PointerText $appDataPointer
    }
}

if ($Mode -eq 'Prepare') {
    Assert-NoFxFileProcess
    if ([IO.Path]::GetFullPath($appDataRoot).TrimEnd('\') -ne $expectedAppDataRoot.TrimEnd('\')) {
        throw "Unexpected AppData target: $appDataRoot"
    }
    if (Test-Path -LiteralPath $StateRoot) {
        throw "StateRoot already exists: $StateRoot"
    }
    New-Item -ItemType Directory -Path $StateRoot, $evidenceRoot | Out-Null

    Copy-Item -LiteralPath $sourcePackage -Destination $sandboxRoot -Recurse
    foreach ($pointer in @((Join-Path $sandboxRoot 'fxfile.ini'), (Join-Path $sandboxRoot '.fxfile'))) {
        if (Test-Path -LiteralPath $pointer) {
            Remove-Item -LiteralPath $pointer -Force
        }
    }
    New-Item -ItemType Directory -Path $customConfig | Out-Null

    $appDataBackup = Join-Path $StateRoot 'appdata_fxfile_original'
    $appDataExisted = Test-Path -LiteralPath $appDataRoot -PathType Container
    if ($appDataExisted) {
        Copy-Item -LiteralPath $appDataRoot -Destination $appDataBackup -Recurse
    }

    $packages = [ordered]@{}
    foreach ($path in @(
        'C:\00 소프트웨어\04 Fxfile',
        (Join-Path $workspaceRoot 'fxfile_run_x64'),
        (Join-Path $workspaceRoot 'fxfile_run_x32')
    )) {
        $packages[$path] = Get-Inventory $path
    }

    $state = [ordered]@{
        Created = (Get-Date).ToString('o')
        AppDataRoot = $appDataRoot
        AppDataExisted = $appDataExisted
        AppDataOriginalInventory = Get-Inventory $appDataRoot
        PackageOriginalInventories = $packages
        SandboxRoot = $sandboxRoot
        CustomConfig = $customConfig
    }
    $state | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $statePath -Encoding utf8
    (Get-AuditObject) | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $evidenceRoot '00-prepared-program.json') -Encoding utf8
    Get-AuditObject | ConvertTo-Json -Depth 12
    exit 0
}

if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) {
    throw "State file not found: $statePath"
}

if ($Mode -eq 'Audit') {
    Assert-NoFxFileProcess
    $audit = Get-AuditObject
    if ($Expected -ne 'Any' -and $audit.InferredActiveLocation -ne $Expected) {
        throw "Expected $Expected but inferred $($audit.InferredActiveLocation)."
    }
    $stamp = Get-Date -Format 'HHmmssfff'
    $audit | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $evidenceRoot "$stamp-$Expected.json") -Encoding utf8
    $audit | ConvertTo-Json -Depth 12
    exit 0
}

if ($Mode -eq 'Restore') {
    Assert-NoFxFileProcess
    # PowerShell 7 otherwise converts ISO timestamp strings into DateTime
    # objects. Keep the original representation so the inventory comparison
    # does not report unchanged files as modified merely because of formatting.
    $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json -DateKind String
    if ([IO.Path]::GetFullPath($state.AppDataRoot).TrimEnd('\') -ne $expectedAppDataRoot.TrimEnd('\')) {
        throw "Refusing to restore unexpected AppData target: $($state.AppDataRoot)"
    }

    if (Test-Path -LiteralPath $appDataRoot) {
        Remove-Item -LiteralPath $appDataRoot -Recurse -Force
    }
    if ($state.AppDataExisted) {
        Copy-Item -LiteralPath (Join-Path $StateRoot 'appdata_fxfile_original') -Destination $appDataRoot -Recurse
    }

    $restored = Get-Inventory $appDataRoot
    $restoredJson = $restored | ConvertTo-Json -Depth 12 -Compress
    $expectedJson = $state.AppDataOriginalInventory | ConvertTo-Json -Depth 12 -Compress
    if ($restoredJson -ne $expectedJson) {
        throw 'AppData inventory did not restore exactly.'
    }

    $packageDiffs = @()
    foreach ($property in $state.PackageOriginalInventories.PSObject.Properties) {
        $actual = Get-Inventory $property.Name
        $actualJson = $actual | ConvertTo-Json -Depth 12 -Compress
        $expectedPackageJson = $property.Value | ConvertTo-Json -Depth 12 -Compress
        if ($actualJson -ne $expectedPackageJson) {
            $packageDiffs += $property.Name
        }
    }
    if ($packageDiffs.Count -ne 0) {
        throw "Production package inventory changed during sandbox test: $($packageDiffs -join ', ')"
    }

    [ordered]@{
        Status = 'Restored'
        AppDataInventoryExact = $true
        ProductionPackagesUnchanged = $true
        EvidenceRoot = $evidenceRoot
        SandboxPreserved = $sandboxRoot
    } | ConvertTo-Json -Depth 5
}
