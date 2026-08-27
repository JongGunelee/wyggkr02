[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectRoot,

    [Parameter(Mandatory = $true)]
    [string]$BuildTempRoot,

    [switch]$AllowLowSystemDriveWithDTemp,

    [string]$LowSystemDriveApproval = '',

    [int64]$InitialSystemFreeBytes = 0,

    [int64]$AllowedSystemDriveDecreaseBytes = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$lowSystemDriveApprovalPhrase = 'I_ACCEPT_LOW_SYSTEM_DRIVE_RISK'
$lowSystemDriveOverrideRequested = [bool]$AllowLowSystemDriveWithDTemp
if ($lowSystemDriveOverrideRequested -and $LowSystemDriveApproval -cne $lowSystemDriveApprovalPhrase) {
    throw ("Low-system-drive override requires the exact acknowledgement: {0}" -f $lowSystemDriveApprovalPhrase)
}
if (-not $lowSystemDriveOverrideRequested -and
    (-not [string]::IsNullOrEmpty($LowSystemDriveApproval) -or $InitialSystemFreeBytes -ne 0 -or
        $AllowedSystemDriveDecreaseBytes -ne 0)) {
    throw 'Low-system-drive approval/baseline was supplied without -AllowLowSystemDriveWithDTemp.'
}

function Assert-Condition([bool]$Condition, [string]$Message) {
    if (-not $Condition) {
        throw $Message
    }
}

function Get-DriveRecord([string]$Path) {
    $fullPath = [IO.Path]::GetFullPath($Path)
    $root = [IO.Path]::GetPathRoot($fullPath)
    $drive = [IO.DriveInfo]::new($root)
    Assert-Condition $drive.IsReady "Volume is not ready: $root"
    $percent = $(if ($drive.TotalSize -gt 0) { 100.0 * $drive.AvailableFreeSpace / $drive.TotalSize } else { 0.0 })
    return [pscustomobject]@{
        Root = $root
        DriveType = $drive.DriveType.ToString()
        FreeBytes = [int64]$drive.AvailableFreeSpace
        FreeGiB = [math]::Round($drive.AvailableFreeSpace / 1GB, 3)
        FreePercentRaw = [double]$percent
        FreePercent = [math]::Round($percent, 3)
    }
}

$project = [IO.Path]::GetFullPath($ProjectRoot).TrimEnd('\')
$buildTemp = [IO.Path]::GetFullPath($BuildTempRoot).TrimEnd('\')
$workspace = [IO.Path]::GetFullPath((Join-Path $project '..')).TrimEnd('\')
$approvedBase = [IO.Path]::GetFullPath((Join-Path $workspace '__BUILD_TEMP_BACKUP__')).TrimEnd('\')
$approvedPrefix = $approvedBase + '\build_temp_'

Assert-Condition (Test-Path -LiteralPath $project -PathType Container) "Project root does not exist: $project"
Assert-Condition (Test-Path -LiteralPath $buildTemp -PathType Container) "Build TEMP does not exist: $buildTemp"
Assert-Condition $buildTemp.StartsWith($approvedPrefix, [StringComparison]::OrdinalIgnoreCase) (
    "Build TEMP must be a new unified-workflow directory under $approvedBase with the build_temp_ prefix: $buildTemp")
Assert-Condition ($env:TEMP -eq $buildTemp -and $env:TMP -eq $buildTemp) (
    "Process TEMP/TMP must exactly match the validated build TEMP. TEMP=$env:TEMP; TMP=$env:TMP")

$systemRoot = $(if ([string]::IsNullOrWhiteSpace($env:SystemDrive)) { [IO.Path]::GetPathRoot($env:SystemRoot) } else { $env:SystemDrive + '\' })
$systemDrive = Get-DriveRecord $systemRoot
$projectDrive = Get-DriveRecord $project
$tempDrive = Get-DriveRecord $buildTemp

$normalSystemGatePassed = $systemDrive.FreeBytes -ge 5GB -and $systemDrive.FreePercentRaw -ge 5.0
$lowSystemDriveOverrideActive = $lowSystemDriveOverrideRequested -and -not $normalSystemGatePassed
if ($lowSystemDriveOverrideActive) {
    Assert-Condition ($InitialSystemFreeBytes -gt 0) (
        'Low-system-drive override requires the unified workflow initial system-drive checkpoint in bytes.')
    Assert-Condition ($AllowedSystemDriveDecreaseBytes -eq 1GB) (
        "Low-system-drive override requires the exact 1 GiB incidental background-write budget; supplied $AllowedSystemDriveDecreaseBytes bytes.")
    Assert-Condition ($systemDrive.FreeBytes -ge 100MB) (
        "System drive emergency floor failed: $($systemDrive.FreeGiB) GiB, $($systemDrive.FreePercent)% free; required >= 100 MiB.")
    Assert-Condition ($systemDrive.FreeBytes -ge ($InitialSystemFreeBytes - $AllowedSystemDriveDecreaseBytes)) (
        "System drive exceeded the workflow-wide 1 GiB incidental background-write budget. InitialBytes=$InitialSystemFreeBytes; CurrentBytes=$($systemDrive.FreeBytes); AllowedDecreaseBytes=$AllowedSystemDriveDecreaseBytes.")
    Assert-Condition ($projectDrive.Root.Equals('D:\', [StringComparison]::OrdinalIgnoreCase) -and
        $tempDrive.Root.Equals('D:\', [StringComparison]::OrdinalIgnoreCase)) (
        "Low-system-drive override requires project and TEMP on fixed local D:. Project=$($projectDrive.Root); TEMP=$($tempDrive.Root)")
}
else {
    Assert-Condition $normalSystemGatePassed (
        "System drive hard gate failed: $($systemDrive.FreeGiB) GiB, $($systemDrive.FreePercent)% free.")
}
Assert-Condition ($projectDrive.DriveType -eq 'Fixed' -and $tempDrive.DriveType -eq 'Fixed') (
    "Project and build TEMP must use a fixed local drive.")
$requiredProjectFreeBytes = [int64]$(if ($lowSystemDriveOverrideActive) { 20GB } else { 10GB })
Assert-Condition ($projectDrive.FreeBytes -ge $requiredProjectFreeBytes -and
    $tempDrive.FreeBytes -ge $requiredProjectFreeBytes) (
    "Project/build TEMP volume must have at least $([math]::Round($requiredProjectFreeBytes / 1GB, 0)) GiB free.")
Assert-Condition ($projectDrive.Root.Equals($tempDrive.Root, [StringComparison]::OrdinalIgnoreCase)) (
    "Build TEMP must be on the same fixed volume as the project.")
$isSingleVolume = $projectDrive.Root.Equals($systemDrive.Root, [StringComparison]::OrdinalIgnoreCase)
$singleVolumeSafe = $isSingleVolume -and $systemDrive.FreeBytes -ge 10GB -and $systemDrive.FreePercentRaw -ge 10.0
Assert-Condition (-not $isSingleVolume -or $singleVolumeSafe) (
    "Project/build TEMP must be on a non-system volume or on a safe single-drive volume with >= 10 GiB and >= 10% free.")

foreach ($boundary in @($project, $workspace, $approvedBase, $buildTemp)) {
    $item = Get-Item -LiteralPath $boundary -Force
    $unsafeAttributes = [IO.FileAttributes]::ReparsePoint -bor [IO.FileAttributes]::Offline
    Assert-Condition (($item.Attributes -band $unsafeAttributes) -eq 0) (
        "Reparse/offline/cloud build boundary is not allowed: $boundary")
}

$probe = Join-Path $buildTemp ('.batch_storage_probe_{0}_{1}.tmp' -f $PID, [Guid]::NewGuid().ToString('N'))
$bytes = [Text.Encoding]::UTF8.GetBytes('FxFile batch storage assertion')
$stream = [IO.FileStream]::new(
    $probe, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write,
    [IO.FileShare]::None, 4096, [IO.FileOptions]::WriteThrough)
try {
    $stream.Write($bytes, 0, $bytes.Length)
    $stream.Flush($true)
}
finally {
    $stream.Dispose()
}
[IO.File]::Delete($probe)
Assert-Condition (-not (Test-Path -LiteralPath $probe)) "Build TEMP write/flush/delete probe cleanup failed: $probe"

$cumulativeDeltaBytes = $(if ($lowSystemDriveOverrideActive) {
    [int64]($systemDrive.FreeBytes - $InitialSystemFreeBytes)
} else { [int64]0 })
Write-Host ("[FxFile Storage] PASS: System={0}GiB/{1}%; Project={2}GiB; TEMP={3}; LowCOverride={4}; CumulativeDeltaBytes={5}" -f
    $systemDrive.FreeGiB, $systemDrive.FreePercent, $projectDrive.FreeGiB, $buildTemp,
    $lowSystemDriveOverrideActive, $cumulativeDeltaBytes)
exit 0
