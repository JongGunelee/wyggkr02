[CmdletBinding()]
param(
    [string]$ProjectRoot = '',
    [string]$TargetX64 = '',
    [string]$RunX64 = '',
    [string]$RunX32 = '',
    [ValidateRange(30, 900)]
    [int]$ConfigureTimeoutSeconds = 300,
    [switch]$SkipConfigureSimulation,

    [switch]$AllowLowSystemDriveWithDTemp,

    [string]$LowSystemDriveApproval = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$lowSystemDriveApprovalPhrase = 'I_ACCEPT_LOW_SYSTEM_DRIVE_RISK'
$allowedSystemDriveDecreaseBytes = [int64](1GB)
$lowSystemDriveOverrideRequested = [bool]$AllowLowSystemDriveWithDTemp
if ($lowSystemDriveOverrideRequested -and $LowSystemDriveApproval -cne $lowSystemDriveApprovalPhrase) {
    throw ("Low-system-drive override requires the exact acknowledgement: -LowSystemDriveApproval '{0}'" -f
        $lowSystemDriveApprovalPhrase)
}
if (-not $lowSystemDriveOverrideRequested -and -not [string]::IsNullOrEmpty($LowSystemDriveApproval)) {
    throw 'LowSystemDriveApproval was supplied without -AllowLowSystemDriveWithDTemp.'
}

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
}
else {
    $ProjectRoot = [IO.Path]::GetFullPath($ProjectRoot)
}

$workspaceRoot = [IO.Path]::GetFullPath((Join-Path $ProjectRoot '..'))

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
    $RunX64 = Join-Path $workspaceRoot 'fxfile_run_x64'
}
if ([string]::IsNullOrWhiteSpace($RunX32)) {
    $RunX32 = Join-Path $workspaceRoot 'fxfile_run_x32'
}

$checks = [Collections.Generic.List[object]]::new()
$requiredWorkflowInputs = @(
    'tools\Test-BuildEnvironment.ps1',
    'tools\Build-Deploy-Verify.ps1',
    'tools\Assert-BuildStorage.ps1',
    'build_master.bat',
    'build_deploy_all.bat',
    'CMakeLists.txt',
    '.vsconfig'
)

function Add-Check(
    [string]$Category,
    [string]$Name,
    [bool]$Required,
    [bool]$Passed,
    [string]$Detail,
    [string]$Remediation = ''
) {
    $script:checks.Add([pscustomobject]@{
        Category = $Category
        Name = $Name
        Required = $Required
        Passed = $Passed
        Detail = $Detail
        Remediation = $Remediation
    })
}

function Test-RequiredFile([string]$Category, [string]$Path, [string]$Description) {
    $passed = Test-Path -LiteralPath $Path -PathType Leaf
    Add-Check $Category $Description $true $passed $Path 'Copy a complete fxfile_working source tree before continuing.'
}

function Get-StorageRecord([string]$Path) {
    $root = [IO.Path]::GetPathRoot([IO.Path]::GetFullPath($Path))
    $drive = [IO.DriveInfo]::new($root)
    if (-not $drive.IsReady) {
        return [pscustomobject]@{
            Root = $root; DriveType = $drive.DriveType.ToString(); IsReady = $false
            TotalBytes = [int64]0; FreeBytes = [int64]0; FreeGiB = 0.0; FreePercent = 0.0
            FreePercentRaw = [double]0.0
        }
    }
    $freePercent = $(if ($drive.TotalSize -gt 0) { 100.0 * $drive.AvailableFreeSpace / $drive.TotalSize } else { 0.0 })
    return [pscustomobject]@{
        Root = $root
        DriveType = $drive.DriveType.ToString()
        IsReady = $true
        TotalBytes = [int64]$drive.TotalSize
        FreeBytes = [int64]$drive.AvailableFreeSpace
        FreeGiB = [math]::Round($drive.AvailableFreeSpace / 1GB, 3)
        FreePercentRaw = [double]$freePercent
        FreePercent = [math]::Round($freePercent, 3)
    }
}

# The first operational action is a read-only storage inspection.  Evidence
# directories and build-tool subprocesses are created only after this block
# has measured and decided the hard gate.
$systemRoot = $(if ([string]::IsNullOrWhiteSpace($env:SystemDrive)) { [IO.Path]::GetPathRoot($env:SystemRoot) } else { $env:SystemDrive + '\' })
$projectStorage = Get-StorageRecord $ProjectRoot
$systemStorage = Get-StorageRecord $systemRoot
$systemHardGatePassed = $systemStorage.IsReady -and $systemStorage.FreeBytes -ge 5GB -and $systemStorage.FreePercentRaw -ge 5.0
$lowSystemDriveOverrideActive = $lowSystemDriveOverrideRequested -and -not $systemHardGatePassed
$systemEmergencyFloorBytes = [int64]100MB
$requiredProjectFreeBytes = [int64]$(if ($lowSystemDriveOverrideActive) { 20GB } else { 10GB })
$projectCapacityPassed = $projectStorage.IsReady -and $projectStorage.DriveType -eq 'Fixed' -and
    $projectStorage.FreeBytes -ge $requiredProjectFreeBytes
$systemEmergencyFloorPassed = $systemStorage.IsReady -and $systemStorage.FreeBytes -ge $systemEmergencyFloorBytes
$effectiveSystemGatePassed = $(if ($lowSystemDriveOverrideActive) {
    $systemEmergencyFloorPassed
} else {
    $systemHardGatePassed
})
$systemRecommendedPassed = $systemStorage.IsReady -and $systemStorage.FreeBytes -ge 10GB -and $systemStorage.FreePercentRaw -ge 10.0
$isSingleVolume = $projectStorage.Root.Equals($systemStorage.Root, [StringComparison]::OrdinalIgnoreCase)
$separateBuildVolumePassed = -not $isSingleVolume
$singleVolumeSafePassed = $isSingleVolume -and $systemRecommendedPassed -and $projectCapacityPassed
$effectiveVolumeLayoutPassed = $separateBuildVolumePassed -or $singleVolumeSafePassed
$projectIsApprovedD = $projectStorage.Root.Equals('D:\', [StringComparison]::OrdinalIgnoreCase)
$projectItem = Get-Item -LiteralPath $ProjectRoot -Force
$workspaceItem = Get-Item -LiteralPath $workspaceRoot -Force
$unsafePathAttributes = [IO.FileAttributes]::ReparsePoint -bor [IO.FileAttributes]::Offline
$evidenceBase = Join-Path $workspaceRoot '__BUILD_TEMP_BACKUP__'
$existingEvidenceSafe = if (Test-Path -LiteralPath $evidenceBase) {
    $evidenceItem = Get-Item -LiteralPath $evidenceBase -Force
    ($evidenceItem.Attributes -band $unsafePathAttributes) -eq 0
}
else {
    $true
}
$safeProjectBoundaryPassed = (($projectItem.Attributes -band $unsafePathAttributes) -eq 0) -and
    (($workspaceItem.Attributes -band $unsafePathAttributes) -eq 0) -and $existingEvidenceSafe
$originalTemp = [Environment]::GetEnvironmentVariable('TEMP', 'Process')
$originalTmp = [Environment]::GetEnvironmentVariable('TMP', 'Process')

Add-Check 'Capacity' ("Project/build TEMP volume is fixed and has >= {0} GiB free" -f
    [math]::Round($requiredProjectFreeBytes / 1GB, 0)) $true $projectCapacityPassed (
    "Root=$($projectStorage.Root); Type=$($projectStorage.DriveType); Free=$($projectStorage.FreeGiB) GiB ($($projectStorage.FreePercent)%)") (
    "Use a writable fixed project drive with at least $([math]::Round($requiredProjectFreeBytes / 1GB, 0)) GiB free before compiling both architectures.")
Add-Check 'Capacity' 'System drive normal hard gate: >= 5 GiB AND >= 5% free' (-not $lowSystemDriveOverrideActive) $systemHardGatePassed (
    "Root=$($systemStorage.Root); Free=$($systemStorage.FreeGiB) GiB ($($systemStorage.FreePercent)%)") (
    'Free system-drive space safely, or use the explicit audited D: override contract only for an authorized exceptional build.')
if ($lowSystemDriveOverrideActive) {
    Add-Check 'Capacity' 'Low-system-drive D: override explicitly acknowledged' $true (
        $LowSystemDriveApproval -ceq $lowSystemDriveApprovalPhrase) (
        'Authorized by explicit switch and exact acknowledgement; this is not the default path.') (
        "Pass -AllowLowSystemDriveWithDTemp -LowSystemDriveApproval '$lowSystemDriveApprovalPhrase'.")
    Add-Check 'Capacity' 'System drive emergency floor: >= 200 MiB free' $true $systemEmergencyFloorPassed (
        "Root=$($systemStorage.Root); Free=$($systemStorage.FreeGiB) GiB ($($systemStorage.FreePercent)%)") (
        'Stop immediately and safely recover C: above 1 GiB before continuing.')
    Add-Check 'Capacity' 'Exceptional project/TEMP volume is D:' $true $projectIsApprovedD (
        "ProjectRoot=$($projectStorage.Root)") (
        'Move the exceptional workflow to the fixed local D: workspace; another volume is not authorized by this override.')
}
Add-Check 'Capacity' 'System drive recommended state: >= 10 GiB AND >= 10% free' $false $systemRecommendedPassed (
    "Root=$($systemStorage.Root); Free=$($systemStorage.FreeGiB) GiB ($($systemStorage.FreePercent)%)") (
    'Recover the system drive to the recommended state before a long build or stress test.')
Add-Check 'Capacity' 'Project/build TEMP is on a non-system volume or safe single-drive volume (>=10 GiB free)' $true $effectiveVolumeLayoutPassed (
    "System=$($systemStorage.Root); Project=$($projectStorage.Root); SingleVolume=$isSingleVolume; Free=$($projectStorage.FreeGiB) GiB") (
    'Place fxfile_working on a non-system fixed drive or ensure the single system drive has at least 10 GiB free.')
Add-Check 'Capacity' 'Project/workspace boundary is not reparse or offline/cloud' $true $safeProjectBoundaryPassed (
    "Project=$ProjectRoot; Workspace=$workspaceRoot") (
    'Use a direct local project path, not a junction, symlink, offline folder, or cloud placeholder.')

$storageGatePassed = $projectCapacityPassed -and $effectiveSystemGatePassed -and $effectiveVolumeLayoutPassed -and
    $safeProjectBoundaryPassed -and (-not $lowSystemDriveOverrideActive -or $projectIsApprovedD)
$evidenceSafeToWrite = $projectCapacityPassed -and $effectiveVolumeLayoutPassed -and $safeProjectBoundaryPassed -and
    (-not $lowSystemDriveOverrideActive -or $projectIsApprovedD)
if (-not $evidenceSafeToWrite) {
    $checks | Format-Table Category, Name, Required, Passed, Detail -AutoSize -Wrap
    Write-Error 'Required project/evidence storage checks failed before any evidence directory was created.'
    exit 1
}

New-Item -ItemType Directory -Path $evidenceBase -Force | Out-Null
$evidenceRoot = Join-Path $evidenceBase ('preflight_{0}' -f (Get-Date -Format 'yyyyMMdd_HHmmss_fff'))
New-Item -ItemType Directory -Path $evidenceRoot | Out-Null
$createdEvidenceItem = Get-Item -LiteralPath $evidenceRoot -Force
if (($createdEvidenceItem.Attributes -band $unsafePathAttributes) -ne 0) {
    Write-Error "Created preflight evidence root is reparse/offline/cloud-backed: $evidenceRoot"
    exit 1
}

if (-not $storageGatePassed) {
    if ($SkipConfigureSimulation) {
        Add-Check 'Simulation' 'CMake configure x64/x32 was explicitly skipped' $true $false (
            'Diagnostic-only run: -SkipConfigureSimulation cannot authorize a build or deployment.') (
            'Recover storage, then run preflight without -SkipConfigureSimulation.')
    }
    else {
        Add-Check 'Environment' 'Task-specific TEMP/TMP write-flush-delete probe' $true $false (
            'Skipped because a required storage hard gate failed before task TEMP creation.') (
            'Recover the system/project storage gate, then run preflight again.')
        Add-Check 'Simulation' 'CMake configure x64/x32' $true $false (
            'Skipped because a required storage hard gate failed before any build-tool subprocess.') (
            'Recover the storage gate, then run the real x64/x32 configure simulation.')
    }
    $buildTempPlan = [IO.Path]::GetFullPath((Join-Path $evidenceRoot 'build_temp'))
    $requiredStorageFailures = @($checks | Where-Object { $_.Required -and -not $_.Passed })
    $storageWarnings = @($checks | Where-Object { -not $_.Required -and -not $_.Passed })
    $earlyReport = [ordered]@{
        CreatedAt = (Get-Date).ToString('o')
        Mode = 'EnvironmentPreflight'
        HostPowerShellVersion = $PSVersionTable.PSVersion.ToString()
        SkipConfigureSimulation = [bool]$SkipConfigureSimulation
        LowSystemDriveOverride = [ordered]@{
            Requested = $lowSystemDriveOverrideRequested
            Active = $lowSystemDriveOverrideActive
            Approval = $(if ($lowSystemDriveOverrideActive) { $LowSystemDriveApproval } else { '' })
            EmergencyFloorBytes = $systemEmergencyFloorBytes
            RequiredProjectFreeBytes = $requiredProjectFreeBytes
            AllowedSystemDriveDecreaseBytes = $allowedSystemDriveDecreaseBytes
        }
        ProjectRoot = $ProjectRoot
        EvidenceRoot = $evidenceRoot
        Storage = [ordered]@{
            SystemDrive = $systemStorage
            ProjectDrive = $projectStorage
            OriginalTemp = $originalTemp
            OriginalTmp = $originalTmp
            OriginalTempDrive = $(try { Get-StorageRecord $originalTemp } catch { $null })
            OriginalTmpDrive = $(try { Get-StorageRecord $originalTmp } catch { $null })
            EvidenceDrive = Get-StorageRecord $evidenceRoot
            BuildTempRoot = $buildTempPlan
            BuildTempDrive = $projectStorage
            BuildTempCreated = $false
            BuildTempProbePassed = $false
            BuildTempCleanupStatus = 'NotCreated'
            RemainingBuildProcesses = @()
            EnvironmentRestored = $true
        }
        Result = 'FAIL'
        RequiredFailureCount = $requiredStorageFailures.Count
        WarningCount = $storageWarnings.Count
        Checks = @($checks)
    }
    $earlyReportPath = Join-Path $evidenceRoot 'preflight_report.json'
    $earlyReport | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $earlyReportPath -Encoding UTF8
    $checks | Format-Table Category, Name, Required, Passed, Detail -AutoSize -Wrap
    Write-Host "Preflight report: $earlyReportPath"
    Write-Error "$($requiredStorageFailures.Count) required storage preflight check(s) failed. No build-tool subprocess or task TEMP was started."
    exit 1
}

$runningOnWindows = [Environment]::OSVersion.Platform -eq [PlatformID]::Win32NT
Add-Check 'OS' 'Windows operating system' $true $runningOnWindows ([Environment]::OSVersion.VersionString) 'Run this workflow on 64-bit Windows 10/11.'
Add-Check 'OS' '64-bit operating system' $true ([Environment]::Is64BitOperatingSystem) ('Is64BitOS={0}' -f [Environment]::Is64BitOperatingSystem) 'Use 64-bit Windows so both x64 and x86 packages can be built and tested.'

$psPassed = $PSVersionTable.PSVersion -ge [Version]'5.1'
Add-Check 'PowerShell' 'PowerShell 5.1 or later' $true $psPassed ($PSVersionTable.PSVersion.ToString()) 'Use built-in Windows PowerShell 5.1 or install current stable PowerShell 7.'

foreach ($requiredWorkflowInput in $requiredWorkflowInputs) {
    Test-RequiredFile 'Build workflow' (Join-Path $ProjectRoot $requiredWorkflowInput) (
        "Required workflow input: $requiredWorkflowInput")
}

foreach ($required in @(
    @{ Path = (Join-Path $ProjectRoot 'src\fxfile\conf_dir.cpp'); Description = 'No-INI configuration source' },
    @{ Path = (Join-Path $ProjectRoot 'src\fxfile\explorer_view.cpp'); Description = 'Shutdown crash fix source' }
)) {
    Test-RequiredFile 'Source' $required.Path $required.Description
}

$cmakeCommand = Get-Command cmake.exe -ErrorAction SilentlyContinue
$cmakeVersion = $null
if ($null -ne $cmakeCommand) {
    $versionLine = (& $cmakeCommand.Source --version | Select-Object -First 1)
    if ($versionLine -match '([0-9]+\.[0-9]+\.[0-9]+)') {
        $cmakeVersion = [Version]$Matches[1]
    }
}
$cmakePassed = $null -ne $cmakeVersion -and $cmakeVersion -ge [Version]'3.21.0'
Add-Check 'Build tools' 'CMake 3.21 or later' $true $cmakePassed $(if ($null -ne $cmakeVersion) { "$cmakeVersion ($($cmakeCommand.Source))" } else { 'cmake.exe not found' }) 'Install the latest stable x64 CMake and add it to PATH.'

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$vsInstall = ''
if (Test-Path -LiteralPath $vswhere -PathType Leaf) {
    $vsInstall = (& $vswhere -latest -version '[17.0,18.0)' -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 Microsoft.VisualStudio.Component.VC.ATLMFC Microsoft.VisualStudio.Component.Windows11SDK.26100 -property installationPath | Select-Object -First 1)
}
$vsPassed = -not [string]::IsNullOrWhiteSpace($vsInstall)
Add-Check 'Build tools' 'Visual Studio 2022 v143 + x64/x86 + MFC + SDK 26100' $true $vsPassed $(if ($vsPassed) { $vsInstall } else { 'Required Visual Studio 2022 component set not found.' }) 'Import fxfile_working\.vsconfig in Visual Studio Installer or install the documented component IDs.'

if ($vsPassed) {
    $toolsetRoot = Join-Path $vsInstall 'VC\Tools\MSVC'
    $toolset = Get-ChildItem -LiteralPath $toolsetRoot -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
    $toolsetPassed = $null -ne $toolset
    Add-Check 'Build tools' 'MSVC toolset directory' $true $toolsetPassed $(if ($toolsetPassed) { $toolset.FullName } else { $toolsetRoot }) 'Modify Visual Studio and add MSVC v143 x64/x86 build tools.'
    if ($toolsetPassed) {
        foreach ($probe in @(
            @{ Name = 'MFC header afxwin.h'; Path = (Join-Path $toolset.FullName 'atlmfc\include\afxwin.h') },
            @{ Name = 'x64 MFC library'; Path = (Join-Path $toolset.FullName 'atlmfc\lib\x64\mfc140.lib') },
            @{ Name = 'x86 MFC library'; Path = (Join-Path $toolset.FullName 'atlmfc\lib\x86\mfc140.lib') },
            @{ Name = 'x64 compiler'; Path = (Join-Path $toolset.FullName 'bin\Hostx64\x64\cl.exe') },
            @{ Name = 'x86 compiler'; Path = (Join-Path $toolset.FullName 'bin\Hostx64\x86\cl.exe') }
        )) {
            Add-Check 'Build tools' $probe.Name $true (Test-Path -LiteralPath $probe.Path -PathType Leaf) $probe.Path 'Repair or modify the Visual Studio C++ workload.'
        }
    }
}

$sdkRoot = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10'
$sdkInclude = Join-Path $sdkRoot 'Include\10.0.26100.0\um\Windows.h'
$sdkX64Lib = Join-Path $sdkRoot 'Lib\10.0.26100.0\um\x64\User32.Lib'
$sdkX86Lib = Join-Path $sdkRoot 'Lib\10.0.26100.0\um\x86\User32.Lib'
foreach ($sdkProbe in @(
    @{ Name = 'Windows SDK 26100 header'; Path = $sdkInclude },
    @{ Name = 'Windows SDK 26100 x64 library'; Path = $sdkX64Lib },
    @{ Name = 'Windows SDK 26100 x86 library'; Path = $sdkX86Lib }
)) {
    Add-Check 'Build tools' $sdkProbe.Name $true (Test-Path -LiteralPath $sdkProbe.Path -PathType Leaf) $sdkProbe.Path 'Add Windows 11 SDK 10.0.26100 in Visual Studio Installer.'
}

$thirdPartyFiles = @(
    'lib\libxml2\bin64\libxml2-2.dll', 'lib\libxml2\bin\libxml2-2.dll',
    'lib\zlib\bin\zlib-x64.dll', 'lib\zlib\bin\zlib.dll',
    'lib\iconv\bin64\libiconv-2.dll', 'lib\iconv\bin\libiconv-2.dll',
    'lib\gfl\lib64\libgfl340.dll', 'lib\gfl\lib\libgfl340.dll',
    'lib\mingwrt\bin64\libgcc_s_sjlj-1.dll', 'lib\mingwrt\bin\libgcc_s_sjlj-1.dll'
)
$missingThirdParty = @($thirdPartyFiles | Where-Object { -not (Test-Path -LiteralPath (Join-Path $ProjectRoot $_) -PathType Leaf) })
Add-Check 'Source' 'Bundled x64/x86 third-party runtime inputs' $true ($missingThirdParty.Count -eq 0) $(if ($missingThirdParty.Count) { $missingThirdParty -join ', ' } else { 'All required probes are present.' }) 'Restore the complete lib directory; do not download arbitrary replacement DLLs.'

$packages = @(
    @{ Name = 'target_x64'; Root = [IO.Path]::GetFullPath($TargetX64) },
    @{ Name = 'run_x64'; Root = [IO.Path]::GetFullPath($RunX64) },
    @{ Name = 'run_x32'; Root = [IO.Path]::GetFullPath($RunX32) }
)
foreach ($package in $packages) {
    $rootExists = Test-Path -LiteralPath $package.Root -PathType Container
    Add-Check 'Deployment' "$($package.Name) package root" $true $rootExists $package.Root 'Create or restore the approved package folder before unified deployment.'
    if ($rootExists) {
        $corePair = (Test-Path -LiteralPath (Join-Path $package.Root 'fxfile\fxfile.conf') -PathType Leaf) -and (Test-Path -LiteralPath (Join-Path $package.Root 'fxfile\fxfile-main.conf') -PathType Leaf)
        Add-Check 'Deployment' "$($package.Name) local configuration core pair" $true $corePair (Join-Path $package.Root 'fxfile') 'Restore both canonical local configuration files.'
        $noPointers = -not (Test-Path -LiteralPath (Join-Path $package.Root 'fxfile.ini') -PathType Leaf) -and -not (Test-Path -LiteralPath (Join-Path $package.Root '.fxfile') -PathType Leaf)
        Add-Check 'Deployment' "$($package.Name) root pointer files absent" $true $noPointers $package.Root 'Move root fxfile.ini/.fxfile to a recoverable backup before continuing.'
    }
}

$zInUse = Test-Path -LiteralPath 'Z:\'
Add-Check 'Environment' 'Z: drive is available for FxFile SUBST' $true (-not $zInUse) $(if ($zInUse) { ((subst Z: 2>$null) -join ' ') } else { 'Available' }) 'Do not overwrite another drive. Remove only a verified stale FxFile SUBST mapping with: subst Z: /D'

$fxProcesses = @(Get-Process fxfile, 'fxfile-launcher', 'fxfile-upchecker', 'fxfile-updater' -ErrorAction SilentlyContinue)
Add-Check 'Environment' 'FxFile-related processes are closed' $true ($fxProcesses.Count -eq 0) $(if ($fxProcesses.Count) { ($fxProcesses | ForEach-Object { "$($_.ProcessName):$($_.Id)" }) -join ', ' } else { 'No related process is running.' }) 'Close FxFile, launcher, upchecker, and updater before deployment.'

$gitCommand = Get-Command git.exe -ErrorAction SilentlyContinue
if ($null -eq $gitCommand) {
    Add-Check 'Source control' 'Git command and repository health' $false $false 'git.exe not found' 'Install Git or make a complete timestamped source backup before editing.'
}
else {
    $gitVersion = (& $gitCommand.Source --version)
    $gitHealthPath = Join-Path $evidenceRoot 'git_health.txt'
    # Windows PowerShell 5.1 promotes native stderr to ErrorRecord objects;
    # collect Git's diagnostic output without letting a known-bad repository
    # abort the rest of the required storage/tool preflight report.
    $savedPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $gitHealthOutput = @(& $gitCommand.Source -C $ProjectRoot cat-file -e 'HEAD^{commit}' 2>&1)
        $gitObjectHealthy = $LASTEXITCODE -eq 0
        if ($gitObjectHealthy) {
            $gitHealthOutput += @(& $gitCommand.Source -C $ProjectRoot status --porcelain=v1 2>&1)
            $gitStatusHealthy = $LASTEXITCODE -eq 0
        }
        else {
            $gitStatusHealthy = $false
        }
    }
    finally {
        $ErrorActionPreference = $savedPreference
    }
    $gitHealthOutput | Set-Content -LiteralPath $gitHealthPath -Encoding UTF8
    $gitHealthy = $gitObjectHealthy -and $gitStatusHealthy
    Add-Check 'Source control' 'Git command and repository health' $false $gitHealthy "$gitVersion; RepositoryHealthy=$gitHealthy" 'The current repository must be repaired, or a full source snapshot must be kept before every change.'
}

$buildTempRoot = [IO.Path]::GetFullPath((Join-Path $evidenceRoot 'build_temp'))
$buildTempReady = $false
$buildTempCreated = $false
$buildTempCleanupStatus = 'NotCreated'
$remainingBuildProcesses = @()
$buildProcessBaseline = @{}
foreach ($process in @(Get-Process cmake, msbuild, cl, link, rc, mspdbsrv -ErrorAction SilentlyContinue)) {
    $buildProcessBaseline[[int]$process.Id] = $true
}

if (-not $SkipConfigureSimulation -and $cmakePassed -and $vsPassed -and
    $projectCapacityPassed -and $effectiveSystemGatePassed -and $effectiveVolumeLayoutPassed -and $safeProjectBoundaryPassed) {
    try {
        $evidencePrefix = [IO.Path]::GetFullPath($evidenceRoot).TrimEnd('\') + '\'
        if (-not $buildTempRoot.StartsWith($evidencePrefix, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Build TEMP escaped the preflight evidence root: $buildTempRoot"
        }
        New-Item -ItemType Directory -Path $buildTempRoot | Out-Null
        $buildTempCreated = $true
        $buildTempCleanupStatus = 'Active'
        $tempItem = Get-Item -LiteralPath $buildTempRoot -Force
        if (($tempItem.Attributes -band $unsafePathAttributes) -ne 0) {
            throw "Build TEMP must not be a reparse/offline/cloud-placeholder path: $buildTempRoot"
        }

        $probePath = Join-Path $buildTempRoot '.write_flush_delete.probe'
        $payload = [Text.Encoding]::UTF8.GetBytes('FxFile preflight build TEMP probe')
        $stream = [IO.FileStream]::new(
            $probePath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write,
            [IO.FileShare]::None, 4096, [IO.FileOptions]::WriteThrough)
        try {
            $stream.Write($payload, 0, $payload.Length)
            $stream.Flush($true)
        }
        finally {
            $stream.Dispose()
        }
        [IO.File]::Delete($probePath)
        if (Test-Path -LiteralPath $probePath) {
            throw "Build TEMP probe could not be deleted: $probePath"
        }

        [Environment]::SetEnvironmentVariable('TEMP', $buildTempRoot, 'Process')
        [Environment]::SetEnvironmentVariable('TMP', $buildTempRoot, 'Process')
        $buildTempReady = $env:TEMP -eq $buildTempRoot -and $env:TMP -eq $buildTempRoot
        Add-Check 'Environment' 'Task-specific TEMP/TMP write-flush-delete probe' $true $buildTempReady (
            "TEMP=$env:TEMP; TMP=$env:TMP; Volume=$($projectStorage.Root)") (
            'Use the preflight-created task directory on the non-system project volume.')

        foreach ($architecture in @(
            @{ Name = 'x64'; Flag = 'x64' },
            @{ Name = 'x32'; Flag = 'Win32' }
        )) {
            $buildDir = Join-Path $evidenceRoot ("configure_$($architecture.Name)")
            $logPath = Join-Path $evidenceRoot ("configure_$($architecture.Name).log")
            $stdoutPath = Join-Path $evidenceRoot ("configure_$($architecture.Name)_stdout.log")
            $stderrPath = Join-Path $evidenceRoot ("configure_$($architecture.Name)_stderr.log")
            $argumentLine = '-S "{0}" -B "{1}" -G "Visual Studio 17 2022" -A {2}' -f $ProjectRoot, $buildDir, $architecture.Flag
            $configureProcess = Start-Process -FilePath $cmakeCommand.Source -ArgumentList $argumentLine -NoNewWindow -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
            $completed = $configureProcess.WaitForExit($ConfigureTimeoutSeconds * 1000)
            if (-not $completed) {
                $configureProcess.Kill()
                $configureProcess.WaitForExit()
            }
            if ($completed) {
                # Windows PowerShell 5.1 can expose a null ExitCode after the
                # timed WaitForExit overload when stdout/stderr are redirected.
                # The parameterless wait drains both redirected streams and
                # Refresh makes the native exit code observable.
                $configureProcess.WaitForExit()
                $configureProcess.Refresh()
                $configureExit = [int]$configureProcess.ExitCode
            }
            else {
                $configureExit = 124
            }
            @(
                "Command: $($cmakeCommand.Source) $argumentLine",
                "Completed=$completed; ExitCode=$configureExit; TimeoutSeconds=$ConfigureTimeoutSeconds",
                "TEMP=$env:TEMP",
                "TMP=$env:TMP",
                '',
                '--- STDOUT ---',
                (Get-Content -LiteralPath $stdoutPath -ErrorAction SilentlyContinue),
                '',
                '--- STDERR ---',
                (Get-Content -LiteralPath $stderrPath -ErrorAction SilentlyContinue)
            ) | Set-Content -LiteralPath $logPath -Encoding UTF8
            Add-Check 'Simulation' "CMake configure $($architecture.Name)" $true ($completed -and $configureExit -eq 0) "Completed=$completed; ExitCode=$configureExit; Log=$logPath" 'Read the configure log and repair the missing generator, SDK, compiler, library, or source input. A timeout is a failure, not permission to deploy.'
        }
    }
    catch {
        if (-not $buildTempReady) {
            Add-Check 'Environment' 'Task-specific TEMP/TMP write-flush-delete probe' $true $false $_.Exception.Message (
                'Repair the non-system project volume or choose a valid local fixed workspace before building.')
            Add-Check 'Simulation' 'CMake configure x64/x32' $true $false 'Skipped because build TEMP setup failed.' (
                'Fix the required storage/TEMP check, then run preflight again.')
        }
        else {
            Add-Check 'Simulation' 'CMake configure x64/x32 unexpected failure' $true $false $_.Exception.Message (
                'Read the configure logs and repair the failed tool invocation before building.')
        }
    }
    finally {
        [Environment]::SetEnvironmentVariable('TEMP', $originalTemp, 'Process')
        [Environment]::SetEnvironmentVariable('TMP', $originalTmp, 'Process')
        $tempFull = [IO.Path]::GetFullPath($buildTempRoot)
        $evidencePrefix = [IO.Path]::GetFullPath($evidenceRoot).TrimEnd('\') + '\'
        $remainingBuildProcesses = @(
            foreach ($process in @(Get-Process cmake, msbuild, cl, link, rc, mspdbsrv -ErrorAction SilentlyContinue)) {
                if (-not $buildProcessBaseline.ContainsKey([int]$process.Id)) {
                    [pscustomobject]@{ Name = $process.ProcessName; Id = $process.Id }
                }
            }
        )
        if ($remainingBuildProcesses.Count -gt 0) {
            $buildTempCleanupStatus = 'PreservedRunningProcesses'
            Add-Check 'Environment' 'Preflight build-process cleanup' $true $false (
                (($remainingBuildProcesses | ForEach-Object { '{0}(PID {1})' -f $_.Name, $_.Id }) -join ', ')) (
                'Wait for or safely stop only the preflight-created build process tree, then audit the preserved TEMP path.')
        }
        elseif (Test-Path -LiteralPath $tempFull -PathType Container) {
            try {
                if (-not $tempFull.StartsWith($evidencePrefix, [StringComparison]::OrdinalIgnoreCase)) {
                    throw "Refusing to clean a preflight TEMP outside the evidence root: $tempFull"
                }
                $tempItem = Get-Item -LiteralPath $tempFull -Force
                $hasReparse = (($tempItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) -or
                    @(
                        Get-ChildItem -LiteralPath $tempFull -Recurse -Force -ErrorAction Stop | Where-Object {
                            ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0
                        }
                    ).Count -gt 0
                if ($hasReparse) {
                    throw "Refusing to recursively clean a preflight TEMP containing a reparse point: $tempFull"
                }
                Remove-Item -LiteralPath $tempFull -Recurse -Force -ErrorAction Stop
                if (Test-Path -LiteralPath $tempFull) {
                    throw "Preflight TEMP still exists after cleanup: $tempFull"
                }
                $buildTempCleanupStatus = 'Removed'
            }
            catch {
                $buildTempCleanupStatus = 'PreservedCleanupError'
                Add-Check 'Environment' 'Preflight TEMP cleanup' $true $false $_.Exception.Message (
                    'Audit the exact preserved evidence/build_temp path; do not broaden deletion or bypass reparse protections.')
            }
        }
        elseif ($buildTempCreated) {
            $buildTempCleanupStatus = 'Removed'
        }
    }
}
elseif (-not $SkipConfigureSimulation) {
    Add-Check 'Environment' 'Task-specific TEMP/TMP write-flush-delete probe' $true $false (
        "Skipped. EffectiveSystemGate=$effectiveSystemGatePassed; NormalSystemGate=$systemHardGatePassed; LowCOverride=$lowSystemDriveOverrideActive; ProjectGate=$projectCapacityPassed; SeparateVolume=$separateBuildVolumePassed; SafeBoundary=$safeProjectBoundaryPassed") (
        'Pass all required storage and tool checks before creating task TEMP or configuring x64/x32.')
    Add-Check 'Simulation' 'CMake configure x64/x32' $true $false 'Skipped because a required tool or storage hard gate failed.' 'Repair all required checks, then run preflight again.'
}
else {
    Add-Check 'Simulation' 'CMake configure x64/x32 was explicitly skipped' $true $false (
        'Diagnostic-only run: -SkipConfigureSimulation cannot authorize a build or deployment.') (
        'Run preflight again without -SkipConfigureSimulation and require both architectures to configure successfully.')
}

$finalSystemStorage = Get-StorageRecord $systemRoot
$finalProjectStorage = Get-StorageRecord $ProjectRoot
$systemFreeDeltaBytes = [int64]($finalSystemStorage.FreeBytes - $systemStorage.FreeBytes)
if ($lowSystemDriveOverrideActive) {
    Add-Check 'Capacity' 'System drive stayed within the 1 GiB incidental background-write budget' $true (
        $finalSystemStorage.FreeBytes -ge ($systemStorage.FreeBytes - $allowedSystemDriveDecreaseBytes)) (
        "Initial=$($systemStorage.FreeGiB) GiB; Final=$($finalSystemStorage.FreeGiB) GiB; DeltaBytes=$systemFreeDeltaBytes; AllowedDecreaseBytes=$allowedSystemDriveDecreaseBytes") (
        'Stop if cumulative C: loss exceeds 1 GiB; this budget covers incidental OS/antivirus variation only and never authorizes intentional C: writes.')
    Add-Check 'Capacity' 'System drive remained above the 200 MiB emergency floor' $true (
        $finalSystemStorage.FreeBytes -ge $systemEmergencyFloorBytes) (
        "Final=$($finalSystemStorage.FreeGiB) GiB ($($finalSystemStorage.FreePercent)%)") (
        'Safely recover C: above 1 GiB before retrying.')
    Add-Check 'Capacity' 'D: project/TEMP reserve remained >= 20 GiB' $true (
        $finalProjectStorage.FreeBytes -ge 20GB) (
        "Final=$($finalProjectStorage.FreeGiB) GiB ($($finalProjectStorage.FreePercent)%)") (
        'Recover at least 20 GiB on the non-system fixed project/TEMP volume.')
}

$requiredFailures = @($checks | Where-Object { $_.Required -and -not $_.Passed })
$warnings = @($checks | Where-Object { -not $_.Required -and -not $_.Passed })
$report = [ordered]@{
    CreatedAt = (Get-Date).ToString('o')
    Mode = 'EnvironmentPreflight'
    HostPowerShellVersion = $PSVersionTable.PSVersion.ToString()
    SkipConfigureSimulation = [bool]$SkipConfigureSimulation
    LowSystemDriveOverride = [ordered]@{
        Requested = $lowSystemDriveOverrideRequested
        Active = $lowSystemDriveOverrideActive
        Approval = $(if ($lowSystemDriveOverrideActive) { $LowSystemDriveApproval } else { '' })
        EmergencyFloorBytes = $systemEmergencyFloorBytes
        RequiredProjectFreeBytes = $requiredProjectFreeBytes
        AllowedSystemDriveDecreaseBytes = $allowedSystemDriveDecreaseBytes
        InitialSystemFreeBytes = $systemStorage.FreeBytes
        FinalSystemFreeBytes = $finalSystemStorage.FreeBytes
        SystemFreeDeltaBytes = $systemFreeDeltaBytes
        CumulativeSystemFreeDeltaBytes = $systemFreeDeltaBytes
    }
    ProjectRoot = $ProjectRoot
    EvidenceRoot = $evidenceRoot
    WorkflowInputs = @(
        foreach ($relativeInput in $requiredWorkflowInputs) {
            $inputPath = Join-Path $ProjectRoot $relativeInput
            [pscustomobject]@{
                RelativePath = $relativeInput
                Exists = Test-Path -LiteralPath $inputPath -PathType Leaf
                SHA256 = $(if (Test-Path -LiteralPath $inputPath -PathType Leaf) {
                    (Get-FileHash -Algorithm SHA256 -LiteralPath $inputPath).Hash
                } else {
                    $null
                })
            }
        }
    )
    Storage = [ordered]@{
        SystemDrive = $systemStorage
        FinalSystemDrive = $finalSystemStorage
        ProjectDrive = $projectStorage
        FinalProjectDrive = $finalProjectStorage
        Checkpoints = @(
            [pscustomobject]@{
                Stage = 'InitialReadOnly'
                SystemDrive = $systemStorage
                ProjectDrive = $projectStorage
                PreviousSystemFreeDeltaBytes = [int64]0
                CumulativeSystemFreeDeltaBytes = [int64]0
                AllowedSystemDriveDecreaseBytes = $allowedSystemDriveDecreaseBytes
            },
            [pscustomobject]@{
                Stage = 'AfterConfigureAndCleanup'
                SystemDrive = $finalSystemStorage
                ProjectDrive = $finalProjectStorage
                PreviousSystemFreeDeltaBytes = $systemFreeDeltaBytes
                CumulativeSystemFreeDeltaBytes = $systemFreeDeltaBytes
                AllowedSystemDriveDecreaseBytes = $allowedSystemDriveDecreaseBytes
            }
        )
        OriginalTemp = $originalTemp
        OriginalTmp = $originalTmp
        OriginalTempDrive = $(try { Get-StorageRecord $originalTemp } catch { $null })
        OriginalTmpDrive = $(try { Get-StorageRecord $originalTmp } catch { $null })
        EvidenceDrive = Get-StorageRecord $evidenceRoot
        BuildTempRoot = $buildTempRoot
        BuildTempDrive = $projectStorage
        BuildTempCreated = $buildTempCreated
        BuildTempProbePassed = $buildTempReady
        BuildTempCleanupStatus = $buildTempCleanupStatus
        RemainingBuildProcesses = @($remainingBuildProcesses)
        EnvironmentRestored = ($env:TEMP -eq $originalTemp -and $env:TMP -eq $originalTmp)
    }
    Result = $(if ($requiredFailures.Count -eq 0) { 'PASS' } else { 'FAIL' })
    RequiredFailureCount = $requiredFailures.Count
    WarningCount = $warnings.Count
    Checks = @($checks)
}
$reportPath = Join-Path $evidenceRoot 'preflight_report.json'
$report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $reportPath -Encoding UTF8

$checks | Format-Table Category, Name, Required, Passed, Detail -AutoSize -Wrap
Write-Host "Preflight report: $reportPath"
if ($warnings.Count -gt 0) {
    Write-Warning "$($warnings.Count) non-blocking warning(s) were found. Read the report before editing source code."
}
if ($requiredFailures.Count -gt 0) {
    Write-Error "$($requiredFailures.Count) required preflight check(s) failed. Do not build or deploy."
    exit 1
}

Write-Host 'PASS: The environment is ready for the unified FxFile workflow.' -ForegroundColor Green
exit 0
