[CmdletBinding()]
param(
    [ValidateSet('BuildDeployVerify', 'DeployVerify', 'VerifyOnly')]
    [string]$Mode = 'BuildDeployVerify',

    [switch]$SkipSmokeTest,

    [switch]$AllowLowSystemDriveWithDTemp,

    [string]$LowSystemDriveApproval = '',

    [string]$TargetX64 = '',
    [string]$RunX64 = '',
    [string]$RunX32 = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:LowSystemDriveApprovalPhrase = 'I_ACCEPT_LOW_SYSTEM_DRIVE_RISK'
$script:LowSystemDriveOverrideRequested = [bool]$AllowLowSystemDriveWithDTemp
if ($script:LowSystemDriveOverrideRequested -and
    $LowSystemDriveApproval -cne $script:LowSystemDriveApprovalPhrase) {
    throw ("Low-system-drive override requires the exact acknowledgement: -LowSystemDriveApproval '{0}'" -f
        $script:LowSystemDriveApprovalPhrase)
}
if (-not $script:LowSystemDriveOverrideRequested -and
    -not [string]::IsNullOrEmpty($LowSystemDriveApproval)) {
    throw 'LowSystemDriveApproval was supplied without -AllowLowSystemDriveWithDTemp.'
}

$script:ProjectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$script:WorkspaceRoot = [IO.Path]::GetFullPath((Join-Path $script:ProjectRoot '..'))

if ([string]::IsNullOrWhiteSpace($TargetX64)) {
    if (Test-Path 'D:\00 소프트웨어\04 Fxfile') {
        $TargetX64 = 'D:\00 소프트웨어\04 Fxfile'
    } elseif (Test-Path 'C:\00 소프트웨어\04 Fxfile') {
        $TargetX64 = 'C:\00 소프트웨어\04 Fxfile'
    } else {
        $TargetX64 = 'D:\00 소프트웨어\04 Fxfile'
    }
}
if ([string]::IsNullOrWhiteSpace($RunX64)) {
    $RunX64 = Join-Path $script:WorkspaceRoot 'fxfile_run_x64'
}
if ([string]::IsNullOrWhiteSpace($RunX32)) {
    $RunX32 = Join-Path $script:WorkspaceRoot 'fxfile_run_x32'
}

$script:ArtifactRoots = @{
    x64 = Join-Path $script:ProjectRoot 'bin\x64'
    x32 = Join-Path $script:ProjectRoot 'bin\x32'
}

$script:Packages = @(
    [pscustomobject]@{ Name = 'target_x64'; Root = [IO.Path]::GetFullPath($TargetX64); Arch = 'x64' },
    [pscustomobject]@{ Name = 'run_x64';    Root = [IO.Path]::GetFullPath($RunX64);    Arch = 'x64' },
    [pscustomobject]@{ Name = 'run_x32';    Root = [IO.Path]::GetFullPath($RunX32);    Arch = 'x32' }
)

$script:CanonicalPackage = $script:Packages[0]
$script:CanonicalConfigDir = Join-Path $script:CanonicalPackage.Root 'fxfile'
$script:CanonicalLauncherIni = Join-Path $script:CanonicalPackage.Root 'fxfile-launcher\fxfile-launcher.ini'
$script:RequiredConfigFiles = @(
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
$script:OptionalRuntimeConfigFiles = @(
    # Created locally by UpcheckerManager after the package is used. It is
    # machine/runtime state, not part of the portable user-environment set.
    'fxfile-upchecker.conf',
    # Created by FileOperationLockStore, including an empty UTF-16 BOM-only
    # file after the lock manager has been opened.  Lock entries contain
    # package/local absolute paths, so they must remain per-package runtime
    # state and must never be propagated as the canonical portable profile.
    'fxfile-operation-locks.conf'
)

$script:RequiredArtifactFiles = @{
    x64 = @(
        'fxfile.exe', 'fxfile.chm', 'fxfile-launcher.exe', 'fxfile-crash.dll',
        'fxfile-keyhook.dll', 'libxprw.dll', 'libgfl340.dll',
        'libxml2-2.dll', 'mfc140u.dll', 'msvcp140.dll',
        'vcruntime140.dll', 'vcruntime140_1.dll', 'zlib1.dll'
    )
    x32 = @(
        'fxfile.exe', 'fxfile.chm', 'fxfile-launcher.exe', 'fxfile-crash.dll',
        'fxfile-keyhook.dll', 'libxprw.dll', 'libgfl340.dll',
        'libxml2-2.dll', 'mfc140u.dll', 'msvcp140.dll',
        'vcruntime140.dll', 'zlib1.dll'
    )
}

$script:RequiredWorkflowInputs = @(
    'tools\Test-BuildEnvironment.ps1',
    'tools\Build-Deploy-Verify.ps1',
    'tools\Assert-BuildStorage.ps1',
    'build_master.bat',
    'build_deploy_all.bat',
    'CMakeLists.txt',
    '.vsconfig'
)

$script:ArtifactFiles = @{}
$script:Journal = [Collections.Generic.List[object]]::new()
$script:RemovedRootBinaries = [Collections.Generic.List[object]]::new()
$script:BackupRoot = $null
$script:ManifestPath = $null
$script:BuildTempRoot = $null
$script:BuildTempBase = Join-Path $script:WorkspaceRoot '__BUILD_TEMP_BACKUP__'
$script:StorageCheckpoints = [Collections.Generic.List[object]]::new()
$script:BuildStorageActive = $false
$script:OriginalProcessEnvironment = $null
$script:BuildProcessBaseline = @{}
$script:TempCleanupStatus = 'NotStarted'
$script:RemainingBuildProcesses = @()
$script:EnvironmentRestored = $false
$script:RollbackCompleted = $true
$script:FinalStorageSnapshotPassed = $true
$script:PreflightReportPath = $null
$script:PreflightReportHash = $null
$script:PreflightReportCreatedAt = $null
$script:LowSystemDriveOverrideActive = $false
$script:StoragePolicyInitialized = $false
$script:SystemEmergencyFloorBytes = [int64](1GB)
$script:AllowedSystemDriveDecreaseBytes = [int64](1GB)
$script:RequiredProjectFreeBytes = [int64]10GB
$script:InitialSystemFreeBytes = [int64]0
$script:LastCheckpointSystemFreeBytes = [int64]0
$script:LastSystemFreeBytes = [int64]0

function Write-Step([string]$Message) {
    Write-Host ('[FxFile Unified] {0}' -f $Message)
}

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) {
        throw $Message
    }
}

function Get-StorageRecord([string]$Path, [string]$Label) {
    $fullPath = [IO.Path]::GetFullPath($Path)
    $root = [IO.Path]::GetPathRoot($fullPath)
    $drive = [IO.DriveInfo]::new($root)
    Assert-True $drive.IsReady "Storage volume is not ready for $Label`: $root"

    $freePercent = 0.0
    if ($drive.TotalSize -gt 0) {
        $freePercent = 100.0 * $drive.AvailableFreeSpace / $drive.TotalSize
    }

    return [pscustomobject]@{
        Label = $Label
        Path = $fullPath
        Root = $root
        DriveType = $drive.DriveType.ToString()
        TotalBytes = [int64]$drive.TotalSize
        FreeBytes = [int64]$drive.AvailableFreeSpace
        FreeGiB = [math]::Round($drive.AvailableFreeSpace / 1GB, 3)
        FreePercentRaw = [double]$freePercent
        FreePercent = [math]::Round($freePercent, 3)
    }
}

function Assert-SafeBuildBoundary([string]$Path, [string]$Label) {
    $fullPath = [IO.Path]::GetFullPath($Path)
    Assert-True (Test-Path -LiteralPath $fullPath -PathType Container) "$Label does not exist: $fullPath"
    $item = Get-Item -LiteralPath $fullPath -Force
    $unsafeAttributes = [IO.FileAttributes]::ReparsePoint -bor [IO.FileAttributes]::Offline
    Assert-True (($item.Attributes -band $unsafeAttributes) -eq 0) (
        "$Label must not be a reparse/offline/cloud-placeholder directory: $fullPath")
}

function Assert-RecentEnvironmentPreflight {
    $attempts = @(
        Get-ChildItem -LiteralPath $script:BuildTempBase -Directory -Filter 'preflight_*' -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending
    )
    Assert-True ($attempts.Count -gt 0) (
        'No environment preflight report exists. Run preflight_build_environment.bat before the unified workflow.')

    $selectedPath = Join-Path $attempts[0].FullName 'preflight_report.json'
    Assert-True (Test-Path -LiteralPath $selectedPath -PathType Leaf) (
        "The newest preflight attempt is incomplete and has no report: $($attempts[0].FullName)")
    $selected = Get-Item -LiteralPath $selectedPath
    try {
        $report = Get-Content -LiteralPath $selected.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw "The newest preflight report is unreadable: $($selected.FullName); $($_.Exception.Message)"
    }
    Assert-True ($report.Result -eq 'PASS' -and [int]$report.RequiredFailureCount -eq 0 -and
        -not [bool]$report.SkipConfigureSimulation) (
        "The newest preflight is not an authorized PASS with real x64/x32 configure simulation: $($selected.FullName)")
    Assert-True ($report.PSObject.Properties.Name -contains 'LowSystemDriveOverride') (
        'The PASS preflight predates the audited low-system-drive policy. Run preflight again with the current scripts.')
    Assert-True ([bool]$report.LowSystemDriveOverride.Requested -eq $script:LowSystemDriveOverrideRequested) (
        'Preflight low-system-drive override request does not match this workflow invocation.')
    Assert-True ([bool]$report.LowSystemDriveOverride.Active -eq $script:LowSystemDriveOverrideActive) (
        'Preflight low-system-drive override mode does not match the current storage policy.')
    if ($script:LowSystemDriveOverrideActive) {
        Assert-True ([string]$report.LowSystemDriveOverride.Approval -ceq $script:LowSystemDriveApprovalPhrase) (
            'PASS preflight did not record the exact low-system-drive risk acknowledgement.')
        Assert-True ([int64]$report.LowSystemDriveOverride.EmergencyFloorBytes -eq $script:SystemEmergencyFloorBytes) (
            'PASS preflight used a different system-drive emergency floor.')
        Assert-True ([int64]$report.LowSystemDriveOverride.RequiredProjectFreeBytes -ge 20GB) (
            'PASS preflight did not require the 20 GiB D: reserve for low-system-drive operation.')
        Assert-True ([int64]$report.LowSystemDriveOverride.AllowedSystemDriveDecreaseBytes -eq
            $script:AllowedSystemDriveDecreaseBytes) (
            'PASS preflight did not use the exact 1 GiB incidental background-write budget.')
        Assert-True ($script:InitialSystemFreeBytes -ge
            ([int64]$report.LowSystemDriveOverride.FinalSystemFreeBytes - $script:AllowedSystemDriveDecreaseBytes)) (
            'System-drive free space changed beyond the audited preflight/workflow continuity budget.')
    }
    Assert-True ([IO.Path]::GetFullPath([string]$report.ProjectRoot).TrimEnd('\').Equals(
        $script:ProjectRoot.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) (
        "Preflight project root does not match this source tree: $($report.ProjectRoot)")

    $createdAt = [DateTimeOffset]::Parse([string]$report.CreatedAt)
    $age = [DateTimeOffset]::Now - $createdAt
    Assert-True ($age.TotalMinutes -ge -5 -and $age.TotalHours -le 2) (
        "The latest PASS preflight is stale ($([math]::Round($age.TotalHours, 2)) hours). Run it again immediately before building.")
    Assert-True ([bool]$report.Storage.BuildTempProbePassed) 'PASS preflight did not record a successful TEMP write/flush/delete probe.'
    Assert-True ([string]$report.Storage.BuildTempCleanupStatus -eq 'Removed') (
        "PASS preflight did not clean its task TEMP: $($report.Storage.BuildTempCleanupStatus)")
    Assert-True ([bool]$report.Storage.EnvironmentRestored) 'PASS preflight did not restore process TEMP/TMP.'
    Assert-True (@($report.Storage.RemainingBuildProcesses).Count -eq 0) 'PASS preflight left build processes running.'
    $inputRecords = @($report.WorkflowInputs)
    Assert-True ($inputRecords.Count -eq $script:RequiredWorkflowInputs.Count) (
        "PASS preflight must bind exactly $($script:RequiredWorkflowInputs.Count) build workflow inputs; found $($inputRecords.Count).")
    $reportedRelativePaths = @(
        $inputRecords | ForEach-Object { ([string]$_.RelativePath).Replace('/', '\') }
    )
    Assert-True (@($reportedRelativePaths | Select-Object -Unique).Count -eq $reportedRelativePaths.Count) (
        'PASS preflight contains duplicate build workflow input records.')
    foreach ($requiredRelativePath in $script:RequiredWorkflowInputs) {
        $matches = @($inputRecords | Where-Object {
            ([string]$_.RelativePath).Replace('/', '\').Equals(
                $requiredRelativePath, [StringComparison]::OrdinalIgnoreCase)
        })
        Assert-True ($matches.Count -eq 1) "PASS preflight is missing the required workflow input: $requiredRelativePath"
        $inputRecord = $matches[0]
        Assert-True ([string]$inputRecord.SHA256 -match '^[0-9A-Fa-f]{64}$') (
            "PASS preflight has an invalid SHA-256 for workflow input: $requiredRelativePath")
        $inputPath = [IO.Path]::GetFullPath((Join-Path $script:ProjectRoot $requiredRelativePath))
        Assert-True (Test-Path -LiteralPath $inputPath -PathType Leaf) "Preflight workflow input is now missing: $inputPath"
        Assert-True ((Get-FileSha256 $inputPath) -eq [string]$inputRecord.SHA256) (
            "A build workflow input changed after preflight. Run preflight again: $inputPath")
    }

    $script:PreflightReportPath = $selected.FullName
    $script:PreflightReportHash = Get-FileSha256 $selected.FullName
    $script:PreflightReportCreatedAt = $createdAt.ToString('o')
    Write-Step "Using recent PASS preflight: $($selected.FullName)"
}

function Get-BuildProcessSnapshot {
    $snapshot = @{}
    foreach ($process in @(Get-Process cmake, msbuild, cl, link, rc, mspdbsrv -ErrorAction SilentlyContinue)) {
        $snapshot[[int]$process.Id] = [pscustomobject]@{
            Name = $process.ProcessName
            StartTime = $(try { $process.StartTime.ToString('o') } catch { '' })
        }
    }
    return $snapshot
}

function Add-StorageCheckpoint([string]$Stage, [switch]$Enforce) {
    $systemRoot = if ([string]::IsNullOrWhiteSpace($env:SystemDrive)) {
        [IO.Path]::GetPathRoot($env:SystemRoot)
    }
    else {
        $env:SystemDrive + '\'
    }

    $system = Get-StorageRecord $systemRoot 'SystemDrive'
    $project = Get-StorageRecord $script:ProjectRoot 'ProjectDrive'
    $normalSystemGatePassed = $system.FreeBytes -ge 5GB -and $system.FreePercentRaw -ge 5.0
    if (-not $script:StoragePolicyInitialized) {
        $script:LowSystemDriveOverrideActive = $script:LowSystemDriveOverrideRequested -and -not $normalSystemGatePassed
        $script:RequiredProjectFreeBytes = [int64]$(if ($script:LowSystemDriveOverrideActive) { 20GB } else { 10GB })
        $script:InitialSystemFreeBytes = [int64]$system.FreeBytes
        $script:LastCheckpointSystemFreeBytes = [int64]$system.FreeBytes
        $script:StoragePolicyInitialized = $true
    }

    $lastCheckpointSystemFreeBytes = [int64]$script:LastCheckpointSystemFreeBytes
    $systemDeltaBytes = [int64]($system.FreeBytes - $lastCheckpointSystemFreeBytes)
    $cumulativeSystemDeltaBytes = [int64]($system.FreeBytes - $script:InitialSystemFreeBytes)
    $systemEmergencyPassed = $system.FreeBytes -ge $script:SystemEmergencyFloorBytes
    $systemWithinIncidentalBudget = $system.FreeBytes -ge
        ($script:InitialSystemFreeBytes - $script:AllowedSystemDriveDecreaseBytes)
    $projectReservePassed = $project.FreeBytes -ge $script:RequiredProjectFreeBytes
    $projectIsApprovedD = $project.Root.Equals('D:\', [StringComparison]::OrdinalIgnoreCase)
    $effectiveSystemGatePassed = $(if ($script:LowSystemDriveOverrideActive) {
        $systemEmergencyPassed -and $systemWithinIncidentalBudget
    } else {
        $normalSystemGatePassed
    })
    $isSingleVolume = $project.Root.Equals($system.Root, [StringComparison]::OrdinalIgnoreCase)
    $singleVolumeSafePassed = $isSingleVolume -and $system.FreeBytes -ge 10GB -and $system.FreePercentRaw -ge 10.0 -and $projectReservePassed
    $effectiveVolumeLayoutPassed = (-not $isSingleVolume) -or $singleVolumeSafePassed
    $policyPassed = $effectiveSystemGatePassed -and $project.DriveType -eq 'Fixed' -and
        $projectReservePassed -and
        $effectiveVolumeLayoutPassed -and
        (-not $script:LowSystemDriveOverrideActive -or $projectIsApprovedD)
    $record = [pscustomobject]@{
        Stage = $Stage
        CapturedAt = (Get-Date).ToString('o')
        SystemDrive = $system
        ProjectDrive = $project
        Temp = $env:TEMP
        Tmp = $env:TMP
        TempDrive = $(try { Get-StorageRecord $env:TEMP 'TEMP' } catch { $null })
        TmpDrive = $(try { Get-StorageRecord $env:TMP 'TMP' } catch { $null })
        EvidenceDrive = $(try { Get-StorageRecord $script:BuildTempBase 'EvidenceRoot' } catch { $null })
        BuildTempDrive = $(if ([string]::IsNullOrEmpty($script:BuildTempRoot)) { $null } else { try { Get-StorageRecord $script:BuildTempRoot 'BuildTempRoot' } catch { $null } })
        Policy = [pscustomobject]@{
            NormalSystemGatePassed = $normalSystemGatePassed
            LowSystemDriveOverrideRequested = $script:LowSystemDriveOverrideRequested
            LowSystemDriveOverrideActive = $script:LowSystemDriveOverrideActive
            EmergencyFloorBytes = $script:SystemEmergencyFloorBytes
            RequiredProjectFreeBytes = $script:RequiredProjectFreeBytes
            LastCheckpointSystemFreeBytes = $lastCheckpointSystemFreeBytes
            SystemFreeDeltaBytes = $systemDeltaBytes
            InitialSystemFreeBytes = $script:InitialSystemFreeBytes
            CumulativeSystemFreeDeltaBytes = $cumulativeSystemDeltaBytes
            AllowedSystemDriveDecreaseBytes = $script:AllowedSystemDriveDecreaseBytes
            SystemWithinIncidentalBudget = $systemWithinIncidentalBudget
            ProjectIsApprovedD = $projectIsApprovedD
            Passed = $policyPassed
        }
    }
    $script:StorageCheckpoints.Add($record)

    Write-Step ("Storage {0}: system {1} GiB/{2}%; project {3} GiB/{4}%; TEMP={5}" -f
        $Stage, $system.FreeGiB, $system.FreePercent,
        $project.FreeGiB, $project.FreePercent, $env:TEMP)

    if ($Enforce) {
        if ($script:LowSystemDriveOverrideActive) {
            Assert-True ($LowSystemDriveApproval -ceq $script:LowSystemDriveApprovalPhrase) (
                'Low-system-drive D: override is missing the exact explicit acknowledgement.')
            Assert-True $systemEmergencyPassed (
                "LOW-C EMERGENCY FLOOR: system drive must retain at least 1 GiB. Current: $($system.FreeGiB) GiB.")
            Assert-True $systemWithinIncidentalBudget (
                "LOW-C WRITE LEAK: workflow-wide C: loss exceeded the 1 GiB incidental background-write budget. " +
                "InitialBytes=$($script:InitialSystemFreeBytes); CurrentBytes=$($system.FreeBytes); " +
                "CumulativeDeltaBytes=$cumulativeSystemDeltaBytes; AllowedDecreaseBytes=$($script:AllowedSystemDriveDecreaseBytes); Stage=$Stage")
            Assert-True $projectIsApprovedD (
                "Low-system-drive override requires the fixed local D: project/TEMP volume. Current: $($project.Root)")
        }
        else {
            Assert-True $normalSystemGatePassed (
                "DISK SAFETY GATE: the system drive must have at least 5 GiB AND 5% free before build/deploy/smoke. " +
                "Current: $($system.FreeGiB) GiB, $($system.FreePercent)%.")
        }
        Assert-True ($project.DriveType -eq 'Fixed') "Project/build TEMP volume must be a fixed local drive: $($project.Root)"
        Assert-True $projectReservePassed (
            "Project/build TEMP volume must have at least $([math]::Round($script:RequiredProjectFreeBytes / 1GB, 0)) GiB free. Current: $($project.FreeGiB) GiB at $($project.Root)")
        Assert-True $effectiveVolumeLayoutPassed (
            "The FxFile project and build TEMP must be on a non-system fixed drive or safe single system volume (>=10 GiB and >=10% free) to protect system-drive writes.")
        if ($script:BuildStorageActive) {
            Assert-True ($record.TempDrive.Root.Equals($project.Root, [StringComparison]::OrdinalIgnoreCase) -and
                $record.TmpDrive.Root.Equals($project.Root, [StringComparison]::OrdinalIgnoreCase)) (
                "Process TEMP/TMP escaped the validated project volume. TEMP=$env:TEMP; TMP=$env:TMP")
        }
    }

    $script:LastCheckpointSystemFreeBytes = [int64]$system.FreeBytes
    $script:LastSystemFreeBytes = [int64]$system.FreeBytes
    return $record
}

function Enter-BuildStorage {
    Add-StorageCheckpoint 'BeforeTempCreation' -Enforce | Out-Null

    $existingBuildProcesses = @(Get-Process cmake, msbuild, cl, link, rc, mspdbsrv -ErrorAction SilentlyContinue)
    Assert-True ($existingBuildProcesses.Count -eq 0) (
        "Close or wait for existing build-tool processes before the unified FxFile build. " +
        (($existingBuildProcesses | ForEach-Object { '{0}(PID {1})' -f $_.ProcessName, $_.Id }) -join ', '))

    Assert-SafeBuildBoundary $script:ProjectRoot 'Project root'
    Assert-SafeBuildBoundary $script:WorkspaceRoot 'Workspace root'

    New-Item -ItemType Directory -Path $script:BuildTempBase -Force | Out-Null
    Assert-SafeBuildBoundary $script:BuildTempBase 'Build TEMP base'

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $candidate = [IO.Path]::GetFullPath((Join-Path $script:BuildTempBase ("build_temp_{0}_{1}" -f $timestamp, $PID)))
    $basePrefix = [IO.Path]::GetFullPath($script:BuildTempBase).TrimEnd('\') + '\'
    Assert-True $candidate.StartsWith($basePrefix, [StringComparison]::OrdinalIgnoreCase) (
        "Build TEMP escaped the approved evidence root: $candidate")
    Assert-True (-not (Test-Path -LiteralPath $candidate)) "Build TEMP already exists: $candidate"

    $original = [pscustomobject]@{
        TEMP = [Environment]::GetEnvironmentVariable('TEMP', 'Process')
        TMP = [Environment]::GetEnvironmentVariable('TMP', 'Process')
        FXFILE_BUILD_TEMP = [Environment]::GetEnvironmentVariable('FXFILE_BUILD_TEMP', 'Process')
        FXFILE_STORAGE_PREFLIGHT = [Environment]::GetEnvironmentVariable('FXFILE_STORAGE_PREFLIGHT', 'Process')
        FXFILE_LOW_C_OVERRIDE = [Environment]::GetEnvironmentVariable('FXFILE_LOW_C_OVERRIDE', 'Process')
        FXFILE_LOW_C_APPROVAL = [Environment]::GetEnvironmentVariable('FXFILE_LOW_C_APPROVAL', 'Process')
        FXFILE_LOW_C_INITIAL_BYTES = [Environment]::GetEnvironmentVariable('FXFILE_LOW_C_INITIAL_BYTES', 'Process')
        FXFILE_LOW_C_WRITE_BUDGET_BYTES = [Environment]::GetEnvironmentVariable('FXFILE_LOW_C_WRITE_BUDGET_BYTES', 'Process')
    }

    try {
        New-Item -ItemType Directory -Path $candidate | Out-Null
        Assert-SafeBuildBoundary $candidate 'Build TEMP'

        $probePath = Join-Path $candidate '.write_flush_delete.probe'
        $payload = [Text.Encoding]::UTF8.GetBytes('FxFile build storage preflight')
        $stream = [IO.FileStream]::new(
            $probePath,
            [IO.FileMode]::CreateNew,
            [IO.FileAccess]::Write,
            [IO.FileShare]::None,
            4096,
            [IO.FileOptions]::WriteThrough)
        try {
            $stream.Write($payload, 0, $payload.Length)
            $stream.Flush($true)
        }
        finally {
            $stream.Dispose()
        }
        [IO.File]::Delete($probePath)
        Assert-True (-not (Test-Path -LiteralPath $probePath)) "Build TEMP probe could not be deleted: $probePath"

        [Environment]::SetEnvironmentVariable('TEMP', $candidate, 'Process')
        [Environment]::SetEnvironmentVariable('TMP', $candidate, 'Process')
        [Environment]::SetEnvironmentVariable('FXFILE_BUILD_TEMP', $candidate, 'Process')
        if ($script:LowSystemDriveOverrideActive) {
            [Environment]::SetEnvironmentVariable('FXFILE_STORAGE_PREFLIGHT', 'PASS_LOW_C_D_TEMP', 'Process')
            [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_OVERRIDE', 'APPROVED', 'Process')
            [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_APPROVAL', $script:LowSystemDriveApprovalPhrase, 'Process')
            [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_INITIAL_BYTES',
                $script:InitialSystemFreeBytes.ToString([Globalization.CultureInfo]::InvariantCulture), 'Process')
            [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_WRITE_BUDGET_BYTES',
                $script:AllowedSystemDriveDecreaseBytes.ToString([Globalization.CultureInfo]::InvariantCulture), 'Process')
        }
        else {
            [Environment]::SetEnvironmentVariable('FXFILE_STORAGE_PREFLIGHT', 'PASS', 'Process')
            [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_OVERRIDE', $null, 'Process')
            [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_APPROVAL', $null, 'Process')
            [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_INITIAL_BYTES', $null, 'Process')
            [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_WRITE_BUDGET_BYTES', $null, 'Process')
        }
        Assert-True ($env:TEMP -eq $candidate -and $env:TMP -eq $candidate) (
            "TEMP/TMP did not resolve to the approved D: task directory. TEMP=$env:TEMP; TMP=$env:TMP")

        $script:OriginalProcessEnvironment = $original
        $script:BuildTempRoot = $candidate
        $script:BuildProcessBaseline = Get-BuildProcessSnapshot
        $script:BuildStorageActive = $true
        $script:TempCleanupStatus = 'Active'
        Add-StorageCheckpoint 'AfterTempRedirect' -Enforce | Out-Null
    }
    catch {
        [Environment]::SetEnvironmentVariable('TEMP', $original.TEMP, 'Process')
        [Environment]::SetEnvironmentVariable('TMP', $original.TMP, 'Process')
        [Environment]::SetEnvironmentVariable('FXFILE_BUILD_TEMP', $original.FXFILE_BUILD_TEMP, 'Process')
        [Environment]::SetEnvironmentVariable('FXFILE_STORAGE_PREFLIGHT', $original.FXFILE_STORAGE_PREFLIGHT, 'Process')
        [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_OVERRIDE', $original.FXFILE_LOW_C_OVERRIDE, 'Process')
        [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_APPROVAL', $original.FXFILE_LOW_C_APPROVAL, 'Process')
        [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_INITIAL_BYTES', $original.FXFILE_LOW_C_INITIAL_BYTES, 'Process')
        [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_WRITE_BUDGET_BYTES', $original.FXFILE_LOW_C_WRITE_BUDGET_BYTES, 'Process')
        if (Test-Path -LiteralPath $candidate -PathType Container) {
            try {
                $candidateItem = Get-Item -LiteralPath $candidate -Force
                $baseItem = Get-Item -LiteralPath $script:BuildTempBase -Force
                $candidatePrefix = [IO.Path]::GetFullPath($script:BuildTempBase).TrimEnd('\') + '\'
                $safeCandidate = $candidate.StartsWith($candidatePrefix, [StringComparison]::OrdinalIgnoreCase) -and
                    $candidate -ne [IO.Path]::GetPathRoot($candidate) -and
                    (($candidateItem.Attributes -band ([IO.FileAttributes]::ReparsePoint -bor [IO.FileAttributes]::Offline)) -eq 0) -and
                    (($baseItem.Attributes -band ([IO.FileAttributes]::ReparsePoint -bor [IO.FileAttributes]::Offline)) -eq 0) -and
                    @(
                        Get-ChildItem -LiteralPath $candidate -Recurse -Force -ErrorAction Stop | Where-Object {
                            ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0
                        }
                    ).Count -eq 0
                if ($safeCandidate) {
                    Remove-Item -LiteralPath $candidate -Recurse -Force -ErrorAction Stop
                }
                else {
                    Write-Warning "Setup-failed build TEMP was preserved because its cleanup boundary is unsafe: $candidate"
                }
            }
            catch {
                Write-Warning "Setup-failed build TEMP was preserved for audit: $candidate; $($_.Exception.Message)"
            }
        }
        throw
    }
}

function Exit-BuildStorage {
    if (-not $script:BuildStorageActive) {
        return $true
    }

    try {
        $finalCheckpoint = Add-StorageCheckpoint 'WorkflowEnd'
        if (-not [bool]$finalCheckpoint.Policy.Passed) {
            $script:FinalStorageSnapshotPassed = $false
            Write-Warning 'Final storage policy failed. Cleanup will still run inside the validated D: boundary.'
        }
    }
    catch {
        Write-Warning "Final storage snapshot failed: $($_.Exception.Message)"
        $script:FinalStorageSnapshotPassed = $false
    }
    finally {
        [Environment]::SetEnvironmentVariable('TEMP', $script:OriginalProcessEnvironment.TEMP, 'Process')
        [Environment]::SetEnvironmentVariable('TMP', $script:OriginalProcessEnvironment.TMP, 'Process')
        [Environment]::SetEnvironmentVariable('FXFILE_BUILD_TEMP', $script:OriginalProcessEnvironment.FXFILE_BUILD_TEMP, 'Process')
        [Environment]::SetEnvironmentVariable('FXFILE_STORAGE_PREFLIGHT', $script:OriginalProcessEnvironment.FXFILE_STORAGE_PREFLIGHT, 'Process')
        [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_OVERRIDE', $script:OriginalProcessEnvironment.FXFILE_LOW_C_OVERRIDE, 'Process')
        [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_APPROVAL', $script:OriginalProcessEnvironment.FXFILE_LOW_C_APPROVAL, 'Process')
        [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_INITIAL_BYTES', $script:OriginalProcessEnvironment.FXFILE_LOW_C_INITIAL_BYTES, 'Process')
        [Environment]::SetEnvironmentVariable('FXFILE_LOW_C_WRITE_BUDGET_BYTES', $script:OriginalProcessEnvironment.FXFILE_LOW_C_WRITE_BUDGET_BYTES, 'Process')
        $script:EnvironmentRestored = $true
    }

    $remaining = @(
        foreach ($process in @(Get-Process cmake, msbuild, cl, link, rc, mspdbsrv -ErrorAction SilentlyContinue)) {
            if (-not $script:BuildProcessBaseline.ContainsKey([int]$process.Id)) {
                $process
            }
        }
    )
    if ($remaining.Count -gt 0) {
        $script:RemainingBuildProcesses = @($remaining | ForEach-Object { [pscustomobject]@{ Name = $_.ProcessName; Id = $_.Id } })
        $script:TempCleanupStatus = 'PreservedRunningProcesses'
        $description = ($remaining | ForEach-Object { '{0}(PID {1})' -f $_.ProcessName, $_.Id }) -join ', '
        Write-Warning "Build TEMP is preserved because workflow-created build processes may still be running: $description; $($script:BuildTempRoot)"
        return $false
    }

    $tempFull = [IO.Path]::GetFullPath($script:BuildTempRoot)
    $basePrefix = [IO.Path]::GetFullPath($script:BuildTempBase).TrimEnd('\') + '\'
    if (-not $tempFull.StartsWith($basePrefix, [StringComparison]::OrdinalIgnoreCase) -or
        $tempFull -eq [IO.Path]::GetPathRoot($tempFull)) {
        Write-Warning "Build TEMP is preserved because its cleanup boundary is invalid: $tempFull"
        $script:TempCleanupStatus = 'PreservedInvalidBoundary'
        return $false
    }

    try {
        if (Test-Path -LiteralPath $tempFull -PathType Container) {
            $baseItem = Get-Item -LiteralPath $script:BuildTempBase -Force
            $tempItem = Get-Item -LiteralPath $tempFull -Force
            $unsafeReparse = (($baseItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) -or
                (($tempItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) -or
                @(
                    Get-ChildItem -LiteralPath $tempFull -Recurse -Force -ErrorAction Stop | Where-Object {
                        ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0
                    }
                ).Count -gt 0
            if ($unsafeReparse) {
                Write-Warning "Build TEMP is preserved because a reparse point was detected during final cleanup: $tempFull"
                $script:TempCleanupStatus = 'PreservedReparsePoint'
                return $false
            }
            Remove-Item -LiteralPath $tempFull -Recurse -Force -ErrorAction Stop
        }
        if (Test-Path -LiteralPath $tempFull) {
            Write-Warning "Build TEMP cleanup was incomplete and the path is preserved for audit: $tempFull"
            $script:TempCleanupStatus = 'PreservedIncompleteCleanup'
            return $false
        }
    }
    catch {
        Write-Warning "Build TEMP cleanup failed; the original workflow result is preserved. Path=$tempFull; Error=$($_.Exception.Message)"
        $script:TempCleanupStatus = 'PreservedCleanupError'
        return $false
    }
    $script:BuildStorageActive = $false
    $script:TempCleanupStatus = 'Removed'
    return $script:FinalStorageSnapshotPassed
}

function Get-RelativePath([string]$Root, [string]$Path) {
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    $pathFull = [IO.Path]::GetFullPath($Path)
    $prefix = $rootFull + '\'
    Assert-True $pathFull.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase) "Path is outside the expected root: $pathFull"
    return $pathFull.Substring($prefix.Length)
}

function Get-FileSha256([string]$Path) {
    return (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash
}

function Get-DirectoryInventory([string]$Path) {
    $inventory = [ordered]@{}
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        return $inventory
    }

    foreach ($file in Get-ChildItem -LiteralPath $Path -Recurse -File -Force | Sort-Object FullName) {
        $relative = Get-RelativePath $Path $file.FullName
        $inventory[$relative] = [ordered]@{
            Length = $file.Length
            SHA256 = Get-FileSha256 $file.FullName
        }
    }
    return $inventory
}

function Get-InventoryDifferences([System.Collections.IDictionary]$Expected, [System.Collections.IDictionary]$Actual) {
    $keys = @($Expected.Keys) + @($Actual.Keys) | Sort-Object -Unique
    return @($keys | Where-Object {
        -not $Expected.Contains($_) -or
        -not $Actual.Contains($_) -or
        $Expected[$_].Length -ne $Actual[$_].Length -or
        $Expected[$_].SHA256 -ne $Actual[$_].SHA256
    })
}

function Get-PeArchitecture([string]$Path) {
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        $reader = [IO.BinaryReader]::new($stream)
        $stream.Position = 0x3c
        $peOffset = $reader.ReadInt32()
        $stream.Position = $peOffset
        Assert-True ($reader.ReadUInt32() -eq 0x00004550) "Invalid PE signature: $Path"
        $machine = $reader.ReadUInt16()
        switch ($machine) {
            0x8664 { return 'x64' }
            0x014c { return 'x32' }
            default { return ('unknown-0x{0:X4}' -f $machine) }
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Assert-NoFxFileProcesses {
    $processes = @(Get-Process fxfile, 'fxfile-launcher', 'fxfile-upchecker', 'fxfile-updater' -ErrorAction SilentlyContinue)
    if ($processes.Count -gt 0) {
        $description = ($processes | ForEach-Object { '{0}(PID {1})' -f $_.ProcessName, $_.Id }) -join ', '
        throw "Close all FxFile-related processes before build/deploy: $description"
    }
}

function Assert-NoStrayWorkspaceArtifacts {
    # Manual cl.exe probes create .obj files in the current directory unless
    # /Fo is specified.  Those files are not inputs to CMake/MSBuild and can
    # make the workspace look like a deployable artifact set.  Only inspect
    # the two project entry directories; files inside build/obj/bin and
    # evidence/backup directories are managed by their own lifecycle.
    $artifactExtensions = @('.obj', '.pch', '.idb', '.ilk', '.tmp', '.temp', '.tlog')
    $entryRoots = @($script:WorkspaceRoot, $script:ProjectRoot)
    $stray = @(
        foreach ($entryRoot in $entryRoots) {
            Get-ChildItem -LiteralPath $entryRoot -File -Force | Where-Object {
                $_.Extension.ToLowerInvariant() -in $artifactExtensions -or
                $_.Name -match '(^~|~$|\.orig$|\.rej$)'
            }
        }
    )

    if ($stray.Count -gt 0) {
        $description = ($stray | ForEach-Object { $_.FullName }) -join '; '
        throw "Stray compiler/temporary files exist at a workspace entry root. Use a dedicated test output directory and /Fo for manual compilation: $description"
    }
}

function Assert-PackageRoots {
    foreach ($package in $script:Packages) {
        Assert-True (Test-Path -LiteralPath $package.Root -PathType Container) "Package root does not exist: $($package.Root)"
        Assert-True (-not (Test-Path -LiteralPath (Join-Path $package.Root 'fxfile.ini') -PathType Leaf)) "Root fxfile.ini must be absent: $($package.Root)"
        Assert-True (-not (Test-Path -LiteralPath (Join-Path $package.Root '.fxfile') -PathType Leaf)) "Root .fxfile must be absent: $($package.Root)"

        $configDir = Join-Path $package.Root 'fxfile'
        Assert-True (Test-Path -LiteralPath (Join-Path $configDir 'fxfile.conf') -PathType Leaf) "Missing local fxfile.conf: $configDir"
        Assert-True (Test-Path -LiteralPath (Join-Path $configDir 'fxfile-main.conf') -PathType Leaf) "Missing local fxfile-main.conf: $configDir"
    }

    Assert-True (Test-Path -LiteralPath $script:CanonicalLauncherIni -PathType Leaf) "Canonical launcher INI is missing: $script:CanonicalLauncherIni"

    $canonicalNames = @(Get-ChildItem -LiteralPath $script:CanonicalConfigDir -File -Force | Select-Object -ExpandProperty Name | Sort-Object)
    $requiredNames = @($script:RequiredConfigFiles | Sort-Object)
    $missingNames = @($requiredNames | Where-Object { $_ -notin $canonicalNames })
    $allowedNames = @($requiredNames + $script:OptionalRuntimeConfigFiles | Sort-Object -Unique)
    $unexpectedNames = @($canonicalNames | Where-Object { $_ -notin $allowedNames })
    Assert-True ($missingNames.Count -eq 0) "Canonical config directory is missing approved files: $($missingNames -join ', ')"
    Assert-True ($unexpectedNames.Count -eq 0) "Canonical config directory contains unexpected files: $($unexpectedNames -join ', ')"
}

function Assert-ExecutionLevelManifest {
    $manifestPath = Join-Path $script:ProjectRoot 'src\fxfile\res\fxfile.exe.manifest'
    Assert-True (Test-Path -LiteralPath $manifestPath -PathType Leaf) "FxFile application manifest is missing: $manifestPath"

    $manifestText = Get-Content -LiteralPath $manifestPath -Raw
    Assert-True ($manifestText -match '<requestedExecutionLevel\s+level="asInvoker"\s+uiAccess="false"\s*/>') 'FxFile manifest must explicitly declare requestedExecutionLevel=asInvoker.'
    Assert-True ($manifestText -notmatch 'requireAdministrator|highestAvailable') 'FxFile portable runtime must not require administrative elevation.'
}

function Get-InstalledCompatibilityState {
    $targetExe = Join-Path $script:CanonicalPackage.Root 'fxfile.exe'
    $layersKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers'
    $pcaStoreKey = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Compatibility Assistant\Store'

    $machineLayer = $null
    try {
        $machineLayer = (Get-ItemProperty -LiteralPath $layersKey -Name $targetExe -ErrorAction Stop).$targetExe
    }
    catch {
        $machineLayer = $null
    }

    $pcaStorePresent = $false
    try {
        $pcaStorePresent = $null -ne (Get-ItemProperty -LiteralPath $pcaStoreKey -Name $targetExe -ErrorAction Stop).$targetExe
    }
    catch {
        $pcaStorePresent = $false
    }

    return [pscustomobject]@{
        Target = $targetExe
        MachineAppCompatLayer = $machineLayer
        ContainsRunAsAdmin = $null -ne $machineLayer -and $machineLayer -match '(^|\s)RUNASADMIN(\s|$)'
        ContainsWindows7Mode = $null -ne $machineLayer -and $machineLayer -match '(^|\s)WIN7RTM(\s|$)'
        UserPcaStoreEntryPresent = $pcaStorePresent
    }
}

function Assert-InstalledCompatibilityOptimized {
    $state = Get-InstalledCompatibilityState
    Assert-True (-not $state.ContainsRunAsAdmin) "Installed target still has RUNASADMIN compatibility mode: $($state.MachineAppCompatLayer)"
    Assert-True (-not $state.ContainsWindows7Mode) "Installed target still has WIN7RTM compatibility mode: $($state.MachineAppCompatLayer)"
    # Windows recreates a short PCA Store telemetry record after a successful
    # asInvoker run.  Presence alone is not an elevation request, so report it
    # in the manifest but gate only on explicit RUNASADMIN/WIN7RTM layers and
    # the embedded asInvoker manifest.
}

function Get-PortableConfigInventory([string]$ConfigDir) {
    $inventory = [ordered]@{}
    foreach ($configName in $script:RequiredConfigFiles) {
        $path = Join-Path $ConfigDir $configName
        Assert-True (Test-Path -LiteralPath $path -PathType Leaf) "Missing portable configuration file: $path"
        $item = Get-Item -LiteralPath $path
        $inventory[$configName] = [ordered]@{
            Length = $item.Length
            SHA256 = Get-FileSha256 $path
        }
    }
    return $inventory
}

function Invoke-ArchitectureBuild([string]$Arch) {
    $buildScript = Join-Path $script:ProjectRoot 'build_master.bat'
    Assert-True (Test-Path -LiteralPath $buildScript -PathType Leaf) "Build script not found: $buildScript"

    Write-Step "Building Release $Arch"
    Push-Location $script:ProjectRoot
    try {
        if ($Arch -eq 'x32') {
            & $buildScript x32
        }
        else {
            & $buildScript
        }
        $buildExitCode = $LASTEXITCODE
        if (Test-Path -LiteralPath 'Z:\') {
            throw "The $Arch build left the temporary Z: mapping accessible (build exit code $buildExitCode). Do not continue to deployment."
        }
        if ($buildExitCode -ne 0) {
            throw "The $Arch build failed with exit code $buildExitCode."
        }
    }
    finally {
        Pop-Location
    }
}

function Build-HelpArtifact {
    $helpRoot = Join-Path $script:ProjectRoot 'docs\htmlhelp'
    $helpProject = Join-Path $helpRoot 'flyExplorer.hhp'
    $helpSource = Join-Path $helpRoot 'Html\shortkey.htm'
    $compiledHelp = Join-Path $helpRoot 'flyExplorer.chm'
    Assert-True (Test-Path -LiteralPath $helpProject -PathType Leaf) "HTML Help project is missing: $helpProject"
    Assert-True (Test-Path -LiteralPath $helpSource -PathType Leaf) "Shortcut manual source is missing: $helpSource"

    $compiler = Get-Command 'hhc.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $compiler) {
        $knownCompiler = 'C:\Program Files (x86)\HTML Help Workshop\hhc.exe'
        Assert-True (Test-Path -LiteralPath $knownCompiler -PathType Leaf) (
            'HTML Help Workshop is required to build fxfile.chm. Install it or make hhc.exe available in PATH.')
        $compilerPath = $knownCompiler
    }
    else {
        $compilerPath = $compiler.Source
    }

    if (Test-Path -LiteralPath $compiledHelp -PathType Leaf) {
        Remove-Item -LiteralPath $compiledHelp -Force
    }
    Write-Step 'Building fxfile.chm from the audited shortcut manual.'
    Push-Location $helpRoot
    try {
        & $compilerPath $helpProject | Out-Host
        # hhc.exe commonly returns 1 even after a successful compilation, so
        # the generated artifact and its freshness are the authoritative checks.
    }
    finally {
        Pop-Location
    }
    Assert-True (Test-Path -LiteralPath $compiledHelp -PathType Leaf) 'HTML Help compiler did not create flyExplorer.chm.'
    $compiledItem = Get-Item -LiteralPath $compiledHelp
    $sourceItem = Get-Item -LiteralPath $helpSource
    Assert-True ($compiledItem.Length -gt 1KB) 'Compiled CHM is unexpectedly small.'
    Assert-True ($compiledItem.LastWriteTimeUtc -ge $sourceItem.LastWriteTimeUtc) 'Compiled CHM is older than the shortcut manual source.'

    foreach ($arch in @('x64', 'x32')) {
        $destination = Join-Path $script:ArtifactRoots[$arch] 'fxfile.chm'
        Copy-Item -LiteralPath $compiledHelp -Destination $destination -Force
        Assert-True ((Get-FileSha256 $compiledHelp) -eq (Get-FileSha256 $destination)) (
            "Failed to stage the compiled manual for $arch`: $destination")
    }
}

function Get-ArtifactRootFiles([string]$Arch) {
    $root = $script:ArtifactRoots[$Arch]
    Assert-True (Test-Path -LiteralPath $root -PathType Container) "Artifact root does not exist: $root"

    # Isolated Visual Studio/CMake builds can place a newer binary in
    # bin\<arch>\Release while the deployment root still contains an older
    # fxfile.exe.  Never allow a successful build followed by a stale deploy.
    $stagedExe = Join-Path $root 'fxfile.exe'
    $nestedReleaseExe = Join-Path $root 'Release\fxfile.exe'
    if ((Test-Path -LiteralPath $stagedExe -PathType Leaf) -and
        (Test-Path -LiteralPath $nestedReleaseExe -PathType Leaf)) {
        $stagedHash = Get-FileSha256 $stagedExe
        $nestedHash = Get-FileSha256 $nestedReleaseExe
        Assert-True ($stagedHash -eq $nestedHash) (
            "Stale deployment artifact detected for $Arch. " +
            "Promote $nestedReleaseExe to $stagedExe before deploying.")
    }

    $files = @(Get-ChildItem -LiteralPath $root -File -Force | Where-Object {
        $_.Extension -ieq '.dll' -or $_.Name -imatch '^fxfile.*\.exe$' -or $_.Name -ieq 'fxfile.chm'
    } | Sort-Object Name)

    foreach ($required in $script:RequiredArtifactFiles[$Arch]) {
        Assert-True ($files.Name -icontains $required) "Required $Arch artifact is missing: $required"
    }

    foreach ($file in $files | Where-Object { $_.Extension -ieq '.exe' -or $_.Extension -ieq '.dll' }) {
        $actualArch = Get-PeArchitecture $file.FullName
        Assert-True ($actualArch -eq $Arch) "Architecture mismatch in $($file.FullName): expected $Arch, got $actualArch"
    }

    $languageDir = Join-Path $root 'Languages'
    Assert-True (Test-Path -LiteralPath (Join-Path $languageDir 'Korean.xml') -PathType Leaf) "Korean.xml is missing from $Arch artifacts."
    $sourceKorean = Join-Path $script:ProjectRoot 'src\fxfile\Languages\Korean.xml'
    $artifactKorean = Join-Path $languageDir 'Korean.xml'
    Assert-True ((Get-FileSha256 $sourceKorean) -eq (Get-FileSha256 $artifactKorean)) (
        "Stale Korean.xml deployment artifact detected for $Arch. " +
        "Promote $sourceKorean to $artifactKorean before deploying.")
    return $files
}

function Initialize-ArtifactManifests {
    $script:ArtifactFiles.x64 = @(Get-ArtifactRootFiles 'x64')
    $script:ArtifactFiles.x32 = @(Get-ArtifactRootFiles 'x32')
}

function Initialize-BackupRoot {
    $evidenceBase = Join-Path $script:WorkspaceRoot '__BUILD_TEMP_BACKUP__'
    New-Item -ItemType Directory -Path $evidenceBase -Force | Out-Null
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $script:BackupRoot = Join-Path $evidenceBase ("unified_deploy_{0}" -f $timestamp)
    Assert-True (-not (Test-Path -LiteralPath $script:BackupRoot)) "Backup path already exists: $script:BackupRoot"
    New-Item -ItemType Directory -Path $script:BackupRoot | Out-Null
    $script:ManifestPath = Join-Path $script:BackupRoot 'deployment_manifest.json'
}

function Backup-ConfigurationSnapshots {
    foreach ($package in $script:Packages) {
        $snapshotRoot = Join-Path $script:BackupRoot ("configuration_snapshots\{0}" -f $package.Name)
        New-Item -ItemType Directory -Path $snapshotRoot -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $package.Root 'fxfile') -Destination $snapshotRoot -Recurse -Force

        $launcherDir = Join-Path $package.Root 'fxfile-launcher'
        if (Test-Path -LiteralPath $launcherDir -PathType Container) {
            Copy-Item -LiteralPath $launcherDir -Destination $snapshotRoot -Recurse -Force
        }
    }
}

function Deploy-File([string]$Source, [string]$Destination, [pscustomobject]$Package) {
    Assert-True (Test-Path -LiteralPath $Source -PathType Leaf) "Deployment source file is missing: $Source"
    $relative = Get-RelativePath $Package.Root $Destination
    $backupPath = Join-Path $script:BackupRoot ("packages\{0}\{1}" -f $Package.Name, $relative)
    $destinationExisted = Test-Path -LiteralPath $Destination -PathType Leaf
    $originalHash = $(if ($destinationExisted) { Get-FileSha256 $Destination } else { '' })

    if ($destinationExisted) {
        New-Item -ItemType Directory -Path (Split-Path -Parent $backupPath) -Force | Out-Null
        Copy-Item -LiteralPath $Destination -Destination $backupPath -Force
    }

    # Record the rollback entry before overwriting. If the copy itself fails after
    # opening the destination, the catch block can still restore the old file.
    $script:Journal.Add([pscustomobject]@{
        Package = $Package.Name
        PackageRoot = $Package.Root
        RelativePath = $relative
        Destination = $Destination
        Existed = $destinationExisted
        BackupPath = $backupPath
        OriginalSHA256 = $originalHash
    })

    New-Item -ItemType Directory -Path (Split-Path -Parent $Destination) -Force | Out-Null
    Copy-Item -LiteralPath $Source -Destination $Destination -Force
}

function Restore-Deployment {
    if ($script:Journal.Count -eq 0) {
        return
    }

    Write-Step 'A failure occurred. Restoring overwritten files from the deployment backup.'
    $rollbackErrors = [Collections.Generic.List[string]]::new()
    for ($index = $script:Journal.Count - 1; $index -ge 0; --$index) {
        $entry = $script:Journal[$index]
        try {
            if ($entry.Existed) {
                if (-not (Test-Path -LiteralPath $entry.BackupPath -PathType Leaf)) {
                    $expectedHash = $(if ($entry.PSObject.Properties.Name -contains 'OriginalSHA256') { [string]$entry.OriginalSHA256 } else { '' })
                    if ((Test-Path -LiteralPath $entry.Destination -PathType Leaf) -and
                        -not [string]::IsNullOrEmpty($expectedHash) -and
                        (Get-FileSha256 $entry.Destination) -eq $expectedHash) {
                        # The intended mutation did not start; the original is
                        # still present and is verified byte-for-byte.
                        continue
                    }
                    throw "Rollback backup is missing and the original destination is not intact: $($entry.BackupPath)"
                }
                Copy-Item -LiteralPath $entry.BackupPath -Destination $entry.Destination -Force
                if (-not (Test-Path -LiteralPath $entry.Destination -PathType Leaf)) {
                    throw "Rollback destination was not restored: $($entry.Destination)"
                }
                $restoredHash = Get-FileSha256 $entry.Destination
                $expectedHash = $(if ($entry.PSObject.Properties.Name -contains 'OriginalSHA256') { [string]$entry.OriginalSHA256 } else { Get-FileSha256 $entry.BackupPath })
                if ($restoredHash -ne $expectedHash) {
                    throw "Rollback hash mismatch: $($entry.Destination)"
                }
            }
            elseif (Test-Path -LiteralPath $entry.Destination -PathType Leaf) {
                $rollbackPath = Join-Path $script:BackupRoot ("rollback_new_files\{0}\{1}" -f $entry.Package, $entry.RelativePath)
                New-Item -ItemType Directory -Path (Split-Path -Parent $rollbackPath) -Force | Out-Null
                Move-Item -LiteralPath $entry.Destination -Destination $rollbackPath
                if (Test-Path -LiteralPath $entry.Destination) {
                    throw "New deployment file remained after rollback: $($entry.Destination)"
                }
            }
        }
        catch {
            $rollbackErrors.Add("$($entry.Package)\$($entry.RelativePath): $($_.Exception.Message)")
        }
    }
    if ($rollbackErrors.Count -gt 0) {
        throw ("Rollback integrity verification failed: " + ($rollbackErrors -join ' | '))
    }
}

function Deploy-ArtifactSet([pscustomobject]$Package) {
    $artifactRoot = $script:ArtifactRoots[$Package.Arch]
    foreach ($artifact in $script:ArtifactFiles[$Package.Arch]) {
        Deploy-File $artifact.FullName (Join-Path $Package.Root $artifact.Name) $Package
    }

    $languageRoot = Join-Path $artifactRoot 'Languages'
    foreach ($languageFile in Get-ChildItem -LiteralPath $languageRoot -Recurse -File -Force) {
        $relative = Get-RelativePath $languageRoot $languageFile.FullName
        Deploy-File $languageFile.FullName (Join-Path (Join-Path $Package.Root 'Languages') $relative) $Package
    }

    $sevenZipRoot = Join-Path $artifactRoot '7zip'
    if (Test-Path -LiteralPath $sevenZipRoot -PathType Container) {
        foreach ($szFile in Get-ChildItem -LiteralPath $sevenZipRoot -Recurse -File -Force) {
            $relative = Get-RelativePath $sevenZipRoot $szFile.FullName
            Deploy-File $szFile.FullName (Join-Path (Join-Path $Package.Root '7zip') $relative) $Package
        }
    }
}

function Reconcile-PackageRootBinaries([pscustomobject]$Package) {
    $desiredNames = @($script:ArtifactFiles[$Package.Arch] | ForEach-Object { $_.Name.ToLowerInvariant() })
    $currentBinaries = @(Get-ChildItem -LiteralPath $Package.Root -File -Force | Where-Object {
        $_.Extension -ieq '.exe' -or $_.Extension -ieq '.dll' -or $_.Extension -ieq '.pdb' -or $_.Extension -ieq '.map'
    })

    foreach ($file in $currentBinaries) {
        if ($desiredNames -contains $file.Name.ToLowerInvariant()) {
            continue
        }

        $relative = Get-RelativePath $Package.Root $file.FullName
        $backupPath = Join-Path $script:BackupRoot ("packages\{0}\{1}" -f $Package.Name, $relative)
        $originalHash = Get-FileSha256 $file.FullName
        New-Item -ItemType Directory -Path (Split-Path -Parent $backupPath) -Force | Out-Null

        # Unexpected product/runtime binaries are moved, not deleted. Recording
        # first lets rollback recover the file even if a later step fails.
        $script:Journal.Add([pscustomobject]@{
            Package = $Package.Name
            PackageRoot = $Package.Root
            RelativePath = $relative
            Destination = $file.FullName
            Existed = $true
            BackupPath = $backupPath
            OriginalSHA256 = $originalHash
        })
        Move-Item -LiteralPath $file.FullName -Destination $backupPath
        $script:RemovedRootBinaries.Add([pscustomobject]@{
            Package = $Package.Name
            RelativePath = $relative
            BackupPath = $backupPath
            SHA256 = Get-FileSha256 $backupPath
        })
    }
}

function Sync-UserEnvironment {
    foreach ($package in $script:Packages | Where-Object { $_.Name -ne $script:CanonicalPackage.Name }) {
        foreach ($configName in $script:RequiredConfigFiles) {
            Deploy-File (Join-Path $script:CanonicalConfigDir $configName) (Join-Path (Join-Path $package.Root 'fxfile') $configName) $package
        }
        Deploy-File $script:CanonicalLauncherIni (Join-Path $package.Root 'fxfile-launcher\fxfile-launcher.ini') $package
    }
}

function Assert-FileMatches([string]$Expected, [string]$Actual, [string]$Description) {
    Assert-True (Test-Path -LiteralPath $Actual -PathType Leaf) "$Description is missing: $Actual"
    $expectedItem = Get-Item -LiteralPath $Expected
    $actualItem = Get-Item -LiteralPath $Actual
    Assert-True ($expectedItem.Length -eq $actualItem.Length) "$Description length mismatch: $Actual"
    Assert-True ((Get-FileSha256 $Expected) -eq (Get-FileSha256 $Actual)) "$Description hash mismatch: $Actual"
}

function Test-DeploymentState {
    Assert-PackageRoots
    Initialize-ArtifactManifests

    $canonicalConfigInventory = Get-PortableConfigInventory $script:CanonicalConfigDir
    $canonicalLanguageInventory = Get-DirectoryInventory (Join-Path $script:ArtifactRoots.x64 'Languages')
    $results = @()

    foreach ($package in $script:Packages) {
        $artifactRoot = $script:ArtifactRoots[$package.Arch]
        foreach ($artifact in $script:ArtifactFiles[$package.Arch]) {
            Assert-FileMatches $artifact.FullName (Join-Path $package.Root $artifact.Name) ("{0} artifact {1}" -f $package.Name, $artifact.Name)
        }

        $expectedBinaryNames = @($script:ArtifactFiles[$package.Arch] | Where-Object {
            $_.Extension -ieq '.exe' -or $_.Extension -ieq '.dll'
        } | ForEach-Object { $_.Name } | Sort-Object)
        $actualBinaryNames = @(Get-ChildItem -LiteralPath $package.Root -File -Force | Where-Object {
            $_.Extension -ieq '.exe' -or $_.Extension -ieq '.dll'
        } | ForEach-Object { $_.Name } | Sort-Object)
        $binaryNameDiff = @(Compare-Object -ReferenceObject $expectedBinaryNames -DifferenceObject $actualBinaryNames)
        $binaryNameDiffText = (($binaryNameDiff | ForEach-Object { $_.InputObject }) -join ', ')
        Assert-True ($binaryNameDiff.Count -eq 0) "Unexpected or missing root EXE/DLL in $($package.Name): $binaryNameDiffText"

        $packageLanguageInventory = Get-DirectoryInventory (Join-Path $package.Root 'Languages')
        $languageDiff = @(Get-InventoryDifferences $canonicalLanguageInventory $packageLanguageInventory)
        Assert-True ($languageDiff.Count -eq 0) "Language files differ in $($package.Name): $($languageDiff -join ', ')"

        $packageConfigInventory = Get-PortableConfigInventory (Join-Path $package.Root 'fxfile')
        $configDiff = @(Get-InventoryDifferences $canonicalConfigInventory $packageConfigInventory)
        Assert-True ($configDiff.Count -eq 0) "Configuration files differ in $($package.Name): $($configDiff -join ', ')"

        Assert-FileMatches $script:CanonicalLauncherIni (Join-Path $package.Root 'fxfile-launcher\fxfile-launcher.ini') ("{0} launcher INI" -f $package.Name)
        Assert-True ((Get-PeArchitecture (Join-Path $package.Root 'fxfile.exe')) -eq $package.Arch) "fxfile.exe architecture mismatch in $($package.Name)."

        $results += [pscustomobject]@{
            Package = $package.Name
            Architecture = $package.Arch
            FxFileSHA256 = Get-FileSha256 (Join-Path $package.Root 'fxfile.exe')
            ConfigFileCount = $packageConfigInventory.Count
            ConfigMatchesCanonical = $true
            LanguageMatchesArtifacts = $true
            RootIni = Test-Path -LiteralPath (Join-Path $package.Root 'fxfile.ini') -PathType Leaf
            RootDotFxFile = Test-Path -LiteralPath (Join-Path $package.Root '.fxfile') -PathType Leaf
        }
    }

    $targetHash = ($results | Where-Object Package -eq 'target_x64').FxFileSHA256
    $runX64Hash = ($results | Where-Object Package -eq 'run_x64').FxFileSHA256
    Assert-True ($targetHash -eq $runX64Hash) 'target_x64 and run_x64 must contain the identical x64 fxfile.exe.'
    return $results
}

function Add-SmokeNativeMethods {
    if ('FxUnifiedDeploy.NativeMethods' -as [type]) {
        return
    }

    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace FxUnifiedDeploy
{
    public static class NativeMethods
    {
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWnd, EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder className, int maxCount);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr GetProp(IntPtr hWnd, string name);

        public static IntPtr GetLayoutWindow(int processId, out int readyViewCount)
        {
            IntPtr layoutWindow = IntPtr.Zero;
            int maxReadyViewCount = 0;
            EnumWindows(delegate(IntPtr hWnd, IntPtr unused)
            {
                uint ownerProcessId;
                GetWindowThreadProcessId(hWnd, out ownerProcessId);
                if (ownerProcessId != (uint)processId)
                    return true;

                long value = GetProp(hWnd, "FxFile.StartupLayoutReadyViewCount").ToInt64();
                if (value > maxReadyViewCount)
                {
                    maxReadyViewCount = (int)value;
                    layoutWindow = hWnd;
                }
                return true;
            }, IntPtr.Zero);

            readyViewCount = maxReadyViewCount;
            return layoutWindow;
        }

        public static int GetIntProperty(IntPtr hWnd, string name)
        {
            if (hWnd == IntPtr.Zero)
                return 0;
            return (int)GetProp(hWnd, name).ToInt64();
        }

        public static int CountVisibleFileLists(IntPtr frameWindow)
        {
            if (frameWindow == IntPtr.Zero)
                return 0;

            int count = 0;
            EnumChildWindows(frameWindow, delegate(IntPtr hWnd, IntPtr unused)
            {
                if (!IsWindowVisible(hWnd))
                    return true;

                StringBuilder className = new StringBuilder(64);
                GetClassName(hWnd, className, className.Capacity);
                if (String.Equals(className.ToString(), "SysListView32", StringComparison.Ordinal))
                    ++count;
                return true;
            }, IntPtr.Zero);
            return count;
        }
    }
}
'@
}

function Get-SmokeExpectedViewCount([string]$StageRoot) {
    $mainConfig = Join-Path $StageRoot 'fxfile\fxfile-main.conf'
    $rowCount = 1
    $columnCount = 1
    $splitLocked = $false
    $lockedRowCount = 0
    $lockedColumnCount = 0
    foreach ($line in Get-Content -LiteralPath $mainConfig) {
        if ($line -match '^main\.view\.row_count\s*=\s*(\d+)') {
            $rowCount = [int]$Matches[1]
        }
        elseif ($line -match '^main\.view\.column_count\s*=\s*(\d+)') {
            $columnCount = [int]$Matches[1]
        }
        elseif ($line -match '^main\.view\.split_locked\s*=\s*(\d+)') {
            $splitLocked = ([int]$Matches[1] -ne 0)
        }
        elseif ($line -match '^main\.view\.locked_row_count\s*=\s*(\d+)') {
            $lockedRowCount = [int]$Matches[1]
        }
        elseif ($line -match '^main\.view\.locked_column_count\s*=\s*(\d+)') {
            $lockedColumnCount = [int]$Matches[1]
        }
    }
    if ($splitLocked -and $lockedRowCount -ge 1 -and $lockedColumnCount -ge 1) {
        $rowCount = $lockedRowCount
        $columnCount = $lockedColumnCount
    }
    Assert-True ($rowCount -ge 1 -and $columnCount -ge 1) "Invalid saved view layout: ${rowCount}x${columnCount}"
    return $rowCount * $columnCount
}

function New-SmokePackage([string]$Arch) {
    $stageRoot = Join-Path $script:BackupRoot ("smoke\{0}" -f $Arch)
    New-Item -ItemType Directory -Path $stageRoot -Force | Out-Null

    foreach ($artifact in $script:ArtifactFiles[$Arch]) {
        Copy-Item -LiteralPath $artifact.FullName -Destination (Join-Path $stageRoot $artifact.Name) -Force
    }
    Copy-Item -LiteralPath (Join-Path $script:ArtifactRoots[$Arch] 'Languages') -Destination $stageRoot -Recurse -Force

    $stageConfig = Join-Path $stageRoot 'fxfile'
    New-Item -ItemType Directory -Path $stageConfig -Force | Out-Null
    foreach ($configName in $script:RequiredConfigFiles) {
        Copy-Item -LiteralPath (Join-Path $script:CanonicalConfigDir $configName) -Destination (Join-Path $stageConfig $configName) -Force
    }

    $stageLauncher = Join-Path $stageRoot 'fxfile-launcher'
    New-Item -ItemType Directory -Path $stageLauncher -Force | Out-Null
    Copy-Item -LiteralPath $script:CanonicalLauncherIni -Destination (Join-Path $stageLauncher 'fxfile-launcher.ini') -Force
    return $stageRoot
}

function Invoke-SmokeProcess([string]$StageRoot, [string]$Arch) {
    $exe = Join-Path $StageRoot 'fxfile.exe'
    $process = $null
    $forced = $false
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $expectedViewCount = Get-SmokeExpectedViewCount $StageRoot
    $readyViewCount = 0
    $skeletonSeconds = -1.0
    $partialVisibleViewCounts = [Collections.Generic.HashSet[int]]::new()
    try {
        $oldCompat = $env:__COMPAT_LAYER
        try {
            $env:__COMPAT_LAYER = 'RunAsInvoker'
            $process = Start-Process -FilePath $exe -WorkingDirectory $StageRoot -PassThru
        }
        finally {
            $env:__COMPAT_LAYER = $oldCompat
        }

        $deadline = [DateTime]::UtcNow.AddSeconds(180)
        $window = [IntPtr]::Zero
        $responsiveSamples = 0
        do {
            if ($process.HasExited) {
                throw "$Arch smoke process exited before becoming ready. ExitCode=$($process.ExitCode)"
            }
            $process.Refresh()
            $window = [FxUnifiedDeploy.NativeMethods]::GetLayoutWindow($process.Id, [ref]$readyViewCount)
            $frameWindow = if ($process.MainWindowHandle -ne [IntPtr]::Zero) { $process.MainWindowHandle } else { $window }
            if ($frameWindow -ne [IntPtr]::Zero) {
                if ($skeletonSeconds -lt 0 -and
                    [FxUnifiedDeploy.NativeMethods]::GetIntProperty($frameWindow, 'FxFile.StartupLayoutSkeletonPainted') -eq 1) {
                    $skeletonSeconds = [math]::Round($watch.Elapsed.TotalSeconds, 3)
                }
                $visibleFileLists = [FxUnifiedDeploy.NativeMethods]::CountVisibleFileLists($frameWindow)
                if ($readyViewCount -lt $expectedViewCount -and
                    $visibleFileLists -gt 0 -and
                    $visibleFileLists -lt $expectedViewCount) {
                    [void]$partialVisibleViewCounts.Add($visibleFileLists)
                }
            }
            if ($window -ne [IntPtr]::Zero -and $readyViewCount -ge $expectedViewCount -and $process.Responding) {
                ++$responsiveSamples
                if ($responsiveSamples -ge 5) {
                    break
                }
            }
            else {
                $responsiveSamples = 0
            }
            Start-Sleep -Milliseconds 100
        } while ([DateTime]::UtcNow -lt $deadline)

        Assert-True ($responsiveSamples -ge 5) "$Arch smoke process did not finish the saved $expectedViewCount-pane layout within 180 seconds. Ready panes: $readyViewCount."
        Assert-True ($skeletonSeconds -ge 0) "$Arch smoke process did not paint the immediate startup layout skeleton."
        Assert-True ($partialVisibleViewCounts.Count -eq 0) "$Arch smoke process exposed partial file-list counts before atomic layout publication: $([string]::Join(', ', $partialVisibleViewCounts))."
        $readySeconds = [math]::Round($watch.Elapsed.TotalSeconds, 3)
        Assert-True ([FxUnifiedDeploy.NativeMethods]::PostMessage($window, 0x0111, [IntPtr]30120, [IntPtr]::Zero)) "$Arch smoke process did not accept the normal exit command."
        # A full portable profile can contain thousands of history/recent
        # records.  On a heavily loaded x86 host, its normal UTF-16 shutdown
        # save can exceed 30 seconds even though it is making forward progress.
        Assert-True ($process.WaitForExit(90000)) "$Arch smoke process did not exit normally within 90 seconds."
        Assert-True ($process.ExitCode -eq 0) "$Arch smoke process exited with code $($process.ExitCode)."

        return [pscustomobject]@{
            Architecture = $Arch
            SkeletonSeconds = $skeletonSeconds
            ReadySeconds = $readySeconds
            SkeletonToReadySeconds = [math]::Round($readySeconds - $skeletonSeconds, 3)
            ReadinessCriterion = 'AllSavedExplorerViewsRedrawn'
            ExpectedViewCount = $expectedViewCount
            ReadyViewCount = $readyViewCount
            AtomicLayoutPublication = $true
            PartialVisibleViewCounts = @($partialVisibleViewCounts)
            ExitCode = $process.ExitCode
            ForcedTermination = $false
            RootIniCreated = Test-Path -LiteralPath (Join-Path $StageRoot 'fxfile.ini') -PathType Leaf
            RootDotFxFileCreated = Test-Path -LiteralPath (Join-Path $StageRoot '.fxfile') -PathType Leaf
        }
    }
    finally {
        if ($null -ne $process -and -not $process.HasExited) {
            Stop-Process -Id $process.Id -Force
            $process.WaitForExit()
            $forced = $true
        }
        if ($forced) {
            Write-Warning "$Arch smoke process required forced termination."
        }
    }
}

function Invoke-SmokeTests {
    Write-Step 'Running isolated no-INI x64/x32 smoke tests.'
    Add-SmokeNativeMethods
    $appDataPath = Join-Path $env:APPDATA 'fxfile'
    $appDataBefore = Get-DirectoryInventory $appDataPath
    $canonicalBefore = Get-DirectoryInventory $script:CanonicalConfigDir

    $results = @()
    foreach ($arch in @('x64', 'x32')) {
        $stage = New-SmokePackage $arch
        $result = Invoke-SmokeProcess $stage $arch
        Assert-True (-not $result.RootIniCreated) "$arch smoke test created fxfile.ini."
        Assert-True (-not $result.RootDotFxFileCreated) "$arch smoke test created .fxfile."
        $results += $result
    }

    $appDataAfter = Get-DirectoryInventory $appDataPath
    $appDataDiff = @(Get-InventoryDifferences $appDataBefore $appDataAfter)
    Assert-True ($appDataDiff.Count -eq 0) "Smoke tests changed AppData files: $($appDataDiff -join ', ')"

    $canonicalAfter = Get-DirectoryInventory $script:CanonicalConfigDir
    $canonicalDiff = @(Get-InventoryDifferences $canonicalBefore $canonicalAfter)
    Assert-True ($canonicalDiff.Count -eq 0) "Smoke tests changed canonical user settings: $($canonicalDiff -join ', ')"
    return $results
}

function Get-ArtifactManifestRecords {
    $records = @()
    foreach ($arch in @('x64', 'x32')) {
        foreach ($file in $script:ArtifactFiles[$arch]) {
            $records += [pscustomobject]@{
                Architecture = $arch
                Name = $file.Name
                Length = $file.Length
                SHA256 = Get-FileSha256 $file.FullName
            }
        }
    }
    return $records
}

function Save-Manifest([object[]]$PackageResults, [object[]]$SmokeResults, [string]$Status, [string]$FailureMessage = '') {
    if ([string]::IsNullOrEmpty($script:ManifestPath)) {
        return
    }

    $manifest = [ordered]@{
        CreatedAt = (Get-Date).ToString('o')
        Mode = $Mode
        Status = $Status
        FailureMessage = $FailureMessage
        ProjectRoot = $script:ProjectRoot
        PreflightReportPath = $script:PreflightReportPath
        PreflightReportSHA256 = $script:PreflightReportHash
        PreflightReportCreatedAt = $script:PreflightReportCreatedAt
        LowSystemDriveOverride = [ordered]@{
            Requested = $script:LowSystemDriveOverrideRequested
            Active = $script:LowSystemDriveOverrideActive
            Approved = ($script:LowSystemDriveOverrideActive -and
                $LowSystemDriveApproval -ceq $script:LowSystemDriveApprovalPhrase)
            Approval = $(if ($script:LowSystemDriveOverrideActive) { $LowSystemDriveApproval } else { '' })
            EmergencyFloorBytes = $script:SystemEmergencyFloorBytes
            RequiredProjectFreeBytes = $script:RequiredProjectFreeBytes
            AllowedSystemDriveDecreaseBytes = $script:AllowedSystemDriveDecreaseBytes
            InitialSystemFreeBytes = $script:InitialSystemFreeBytes
            LastSystemFreeBytes = $script:LastSystemFreeBytes
            SystemFreeDeltaBytes = [int64]($script:LastSystemFreeBytes - $script:InitialSystemFreeBytes)
            CumulativeSystemFreeDeltaBytes = [int64]($script:LastSystemFreeBytes - $script:InitialSystemFreeBytes)
            TempScope = 'ProcessOnly'
            RequiredTempVolume = $(if ($script:LowSystemDriveOverrideActive) { 'D:\' } else { 'NonSystemProjectVolume' })
        }
        BackupRoot = $script:BackupRoot
        BuildTempRoot = $script:BuildTempRoot
        TempCleanupStatus = $script:TempCleanupStatus
        RemainingBuildProcesses = @($script:RemainingBuildProcesses)
        EnvironmentRestored = $script:EnvironmentRestored
        RollbackCompleted = $script:RollbackCompleted
        FinalStorageSnapshotPassed = $script:FinalStorageSnapshotPassed
        StorageCheckpoints = @($script:StorageCheckpoints)
        CanonicalConfigDirectory = $script:CanonicalConfigDir
        ConfigFiles = $script:RequiredConfigFiles
        Artifacts = @(Get-ArtifactManifestRecords)
        RemovedUnexpectedRootBinaries = @($script:RemovedRootBinaries)
        Packages = @($PackageResults)
        SmokeTests = @($SmokeResults)
        ExecutionLevel = 'asInvoker'
        InstalledCompatibility = Get-InstalledCompatibilityState
        KnownParityExceptions = @(
            'target_x64 may retain DISABLEDXMAXIMIZEDWINDOWEDMODE; RUNASADMIN/WIN7RTM are rejected. Windows may recreate an informational PCA Store record after a successful asInvoker run.',
            'run_x32 is the x86 counterpart and cannot be binary-identical to x64.',
            'The target updater directory is outside the unified core runtime set; update_check.enable is currently 0.',
            'Absolute C: and D: paths in user settings require matching assets on another computer.'
        )
    }
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $script:ManifestPath -Encoding UTF8
}

$packageResults = @()
$smokeResults = @()
$pendingSuccess = $false
$workflowFailure = $null

try {
    if ($Mode -ne 'VerifyOnly') {
        # The very first mutating-workflow action is a read-only drive check.
        # No evidence/task TEMP directory exists before this hard gate passes.
        Add-StorageCheckpoint 'InitialReadOnlyPreflight' -Enforce | Out-Null
        Assert-RecentEnvironmentPreflight
        Add-StorageCheckpoint 'AfterPreflightValidation' -Enforce | Out-Null
    }
    Assert-NoStrayWorkspaceArtifacts
    Assert-NoFxFileProcesses
    Assert-ExecutionLevelManifest
    Assert-PackageRoots
    Assert-InstalledCompatibilityOptimized

    # VerifyOnly is deliberately read-only and remains available for diagnosis
    # on a low-space system. Every mutating mode must pass the storage gate and
    # inherit the task-specific non-system TEMP/TMP directory.
    if ($Mode -ne 'VerifyOnly') {
        Enter-BuildStorage
        Initialize-BackupRoot
    }

    if ($Mode -eq 'BuildDeployVerify') {
        Invoke-ArchitectureBuild 'x64'
        Add-StorageCheckpoint 'AfterX64Build' -Enforce | Out-Null
        Invoke-ArchitectureBuild 'x32'
        Add-StorageCheckpoint 'AfterX32Build' -Enforce | Out-Null
        Build-HelpArtifact
        Add-StorageCheckpoint 'AfterHelpBuild' -Enforce | Out-Null
    }

    Initialize-ArtifactManifests

    if ($Mode -eq 'VerifyOnly') {
        Write-Step 'Verifying the current unified deployment without changing files.'
        $packageResults = @(Test-DeploymentState)
        Write-Step 'Verification completed successfully.'
        $packageResults | Format-Table -AutoSize
        exit 0
    }

    Add-StorageCheckpoint 'BeforeDeployment' -Enforce | Out-Null
    Backup-ConfigurationSnapshots

    try {
        Write-Step "Deploying one controlled runtime set. Backup: $script:BackupRoot"
        foreach ($package in $script:Packages) {
            Deploy-ArtifactSet $package
            Reconcile-PackageRootBinaries $package
        }
        Sync-UserEnvironment

        $packageResults = @(Test-DeploymentState)
        if (-not $SkipSmokeTest) {
            Add-StorageCheckpoint 'BeforeSmoke' -Enforce | Out-Null
            $smokeResults = @(Invoke-SmokeTests)
        }

        Add-StorageCheckpoint 'BeforeSuccessManifest' -Enforce | Out-Null
        $pendingSuccess = $true
        $packageResults | Format-Table -AutoSize
        if ($smokeResults.Count -gt 0) {
            $smokeResults | Format-Table -AutoSize
        }
    }
    catch {
        $failure = $_.Exception.Message
        $workflowFailure = $failure
        throw
    }
}
catch {
    $workflowFailure = $_.Exception.Message
    if ($Mode -ne 'VerifyOnly' -and -not [string]::IsNullOrEmpty($script:ManifestPath)) {
        try {
            Restore-Deployment
        }
        catch {
            $workflowFailure += " | Rollback error: $($_.Exception.Message)"
            $script:RollbackCompleted = $false
        }
    }
    throw
}
finally {
    $cleanupPassed = Exit-BuildStorage
    if ($Mode -ne 'VerifyOnly' -and -not [string]::IsNullOrEmpty($script:ManifestPath)) {
        if ($pendingSuccess -and $cleanupPassed) {
            Save-Manifest $packageResults $smokeResults 'Success'
            Write-Step "Unified build/deploy/verification completed. Manifest: $script:ManifestPath"
        }
        elseif ($pendingSuccess -and -not $cleanupPassed) {
            $workflowFailure = "Build TEMP cleanup/process audit failed: $($script:TempCleanupStatus)"
            Save-Manifest $packageResults $smokeResults 'FailedCleanupIncomplete' $workflowFailure
            throw $workflowFailure
        }
        elseif (-not [string]::IsNullOrEmpty($workflowFailure)) {
            if (-not $cleanupPassed) {
                $workflowFailure += " | Build TEMP cleanup/process audit incomplete: $($script:TempCleanupStatus)"
            }
            $failureStatus = $(if ($script:RollbackCompleted) { 'FailedAndRolledBack' } else { 'FailedRollbackIncomplete' })
            if (-not $cleanupPassed -and $failureStatus -eq 'FailedAndRolledBack') {
                $failureStatus = 'FailedCleanupIncomplete'
            }
            Save-Manifest $packageResults $smokeResults $failureStatus $workflowFailure
        }
    }
}
