[CmdletBinding()]
param(
    [ValidateRange(1, 20)]
    [int]$IterationsPerLayout = 3,

    [string]$RunX64 = '',
    [string]$RunX32 = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$workspaceRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot '..'))
$evidenceBase = [IO.Path]::GetFullPath((Join-Path $workspaceRoot '__BUILD_TEMP_BACKUP__')).TrimEnd('\')

if ([string]::IsNullOrWhiteSpace($RunX64)) { $RunX64 = Join-Path $workspaceRoot 'fxfile_run_x64' }
if ([string]::IsNullOrWhiteSpace($RunX32)) { $RunX32 = Join-Path $workspaceRoot 'fxfile_run_x32' }
$RunX64 = [IO.Path]::GetFullPath($RunX64)
$RunX32 = [IO.Path]::GetFullPath($RunX32)

if (@(Get-Process fxfile, 'fxfile-launcher', 'fxfile-upchecker', 'fxfile-updater' -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'Close every FxFile-related process before the repeated shutdown test.'
}

foreach ($root in @($RunX64, $RunX32)) {
    if (-not (Test-Path -LiteralPath (Join-Path $root 'fxfile.exe') -PathType Leaf)) {
        throw "Package root is missing fxfile.exe: $root"
    }
}

if (-not ('Task093NativeMethods' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class Task093NativeMethods
{
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool PostMessage(
        IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam);
}
'@
}

function Assert-TaskPath([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path)
    $prefix = $evidenceBase + '\'
    if (-not $full.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Task path escaped the evidence boundary: $full"
    }
    return $full
}

function Get-CrashReportNames {
    $desktop = [Environment]::GetFolderPath('Desktop')
    return @(
        Get-ChildItem -LiteralPath $desktop -Force -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -like 'fxfile_error_report_*' -and
                ($_.PSIsContainer -or $_.Extension -ieq '.zip')
            } |
            Select-Object -ExpandProperty Name
    )
}

function Get-LaunchArguments([string]$Layout, [int]$ViewCount, [string]$SelectionTarget) {
    $arguments = [Collections.Generic.List[string]]::new()
    $arguments.Add('--window')
    $arguments.Add($Layout)
    for ($view = 1; $view -le $ViewCount; ++$view) {
        $arguments.Add("--dir$view")
        $arguments.Add($SelectionTarget)
    }
    return $arguments.ToArray()
}

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$taskRoot = Assert-TaskPath (Join-Path $evidenceBase "task093_shutdown_stress_$stamp")
$x64Root = Assert-TaskPath (Join-Path $taskRoot 'x64')
$x32Root = Assert-TaskPath (Join-Path $taskRoot 'x32')
New-Item -ItemType Directory -Path $taskRoot -Force | Out-Null

$beforeCrashReports = Get-CrashReportNames
$records = [Collections.Generic.List[object]]::new()
$forcedPids = [Collections.Generic.List[int]]::new()

try {
    Copy-Item -LiteralPath $RunX64 -Destination $x64Root -Recurse
    Copy-Item -LiteralPath $RunX32 -Destination $x32Root -Recurse

    $packages = @(
        [pscustomobject]@{ Architecture = 'x64'; Root = $x64Root },
        [pscustomobject]@{ Architecture = 'x32'; Root = $x32Root }
    )
    $layouts = @(
        [pscustomobject]@{ Name = '1x1'; Views = 1 },
        [pscustomobject]@{ Name = '2x2'; Views = 4 },
        [pscustomobject]@{ Name = '2x3'; Views = 6 }
    )

    foreach ($package in $packages) {
        $exe = Join-Path $package.Root 'fxfile.exe'
        foreach ($layout in $layouts) {
            for ($iteration = 1; $iteration -le $IterationsPerLayout; ++$iteration) {
                # --dirN receives a directory. Passing the executable path here
                # exercises file-selection startup semantics instead of the
                # pane startup/exit lifetime that this test is intended to own.
                $arguments = Get-LaunchArguments $layout.Name $layout.Views $package.Root
                $startedAt = Get-Date
                $process = Start-Process -FilePath $exe -ArgumentList $arguments -WorkingDirectory $package.Root -PassThru
                $inputIdle = $false
                $responsive = $false
                $closeSent = $false
                $forced = $false
                $exitCode = $null
                $mainHwnd = [IntPtr]::Zero

                try {
                    try { $inputIdle = $process.WaitForInputIdle(15000) } catch { $inputIdle = $false }
                    $deadline = (Get-Date).AddSeconds(20)
                    do {
                        $process.Refresh()
                        if ($process.HasExited) { break }
                        $mainHwnd = $process.MainWindowHandle
                        if ($mainHwnd -ne [IntPtr]::Zero) { break }
                        Start-Sleep -Milliseconds 100
                    } while ((Get-Date) -lt $deadline)

                    if (-not $process.HasExited -and $mainHwnd -ne [IntPtr]::Zero) {
                        $process.Refresh()
                        $responsive = $process.Responding
                        # Use FxFile's explicit Exit command. WM_CLOSE can be
                        # configured to hide the frame in the system tray.
                        $closeSent = [Task093NativeMethods]::PostMessage(
                            $mainHwnd, 0x0111, [UIntPtr]30120, [IntPtr]::Zero)
                    }

                    if (-not $process.WaitForExit(90000)) {
                        $forced = $true
                        $forcedPids.Add([int]$process.Id)
                        $process.Kill()
                        $process.WaitForExit(10000) | Out-Null
                    }
                    if ($process.HasExited) { $exitCode = $process.ExitCode }
                }
                finally {
                    if (-not $process.HasExited) {
                        $forced = $true
                        $forcedPids.Add([int]$process.Id)
                        $process.Kill()
                        $process.WaitForExit(10000) | Out-Null
                    }
                    $process.Dispose()
                }

                $records.Add([pscustomobject]@{
                    Architecture = $package.Architecture
                    Layout = $layout.Name
                    Iteration = $iteration
                    InputIdle = $inputIdle
                    Responsive = $responsive
                    MainWindowFound = ($mainHwnd -ne [IntPtr]::Zero)
                    CloseSent = $closeSent
                    ForcedTermination = $forced
                    ExitCode = $exitCode
                    ElapsedSeconds = [math]::Round(((Get-Date) - $startedAt).TotalSeconds, 3)
                })
                Start-Sleep -Milliseconds 250
            }
        }
    }
}
finally {
    foreach ($path in @($x64Root, $x32Root)) {
        $resolved = Assert-TaskPath $path
        if (Test-Path -LiteralPath $resolved -PathType Container) {
            $item = Get-Item -LiteralPath $resolved -Force
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "Refusing to remove a reparse-point test copy: $resolved"
            }
            [IO.Directory]::Delete($resolved, $true)
        }
    }
}

$afterCrashReports = Get-CrashReportNames
$newCrashReports = @($afterCrashReports | Where-Object { $_ -notin $beforeCrashReports })
$failures = @($records | Where-Object {
    -not $_.InputIdle -or -not $_.Responsive -or -not $_.MainWindowFound -or
    -not $_.CloseSent -or $_.ForcedTermination -or $_.ExitCode -ne 0
})
$remainingProcesses = @(Get-Process fxfile, 'fxfile-launcher', 'fxfile-upchecker', 'fxfile-updater' -ErrorAction SilentlyContinue)

$result = [ordered]@{
    CreatedAt = (Get-Date).ToString('o')
    Result = $(if ($failures.Count -eq 0 -and $newCrashReports.Count -eq 0 -and $remainingProcesses.Count -eq 0) { 'PASS' } else { 'FAIL' })
    IterationsPerLayout = $IterationsPerLayout
    TotalRuns = $records.Count
    FailedRuns = $failures.Count
    NewCrashReports = $newCrashReports
    ForcedPids = $forcedPids.ToArray()
    RemainingFxFileProcesses = @($remainingProcesses | ForEach-Object { "{0}:{1}" -f $_.ProcessName, $_.Id })
    Records = $records.ToArray()
}
$resultPath = Join-Path $taskRoot 'runtime_report.json'
$result | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $resultPath -Encoding UTF8
$result | ConvertTo-Json -Depth 6
if ($result.Result -ne 'PASS') { exit 1 }
