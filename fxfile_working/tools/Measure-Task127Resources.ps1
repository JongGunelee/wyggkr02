[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$ExecutablePath,[ValidateRange(10,900)][int]$DurationSeconds=120)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$root=Join-Path $workspace '__BUILD_TEMP_BACKUP__/task127_results_20261012'
$exe=[IO.Path]::GetFullPath($ExecutablePath)
if(-not $exe.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Only isolated task executables may be sampled.'}
$tasks=@(Get-Process fxfile -ErrorAction SilentlyContinue|Where-Object Path -eq $exe)
if($tasks.Count -ne 1){throw 'Expected one task process.'}
$taskPid=$tasks[0].Id;$samples=[Collections.Generic.List[object]]::new();$timer=[Diagnostics.Stopwatch]::StartNew()
while($timer.Elapsed.TotalSeconds -lt $DurationSeconds){
 $taskProcess=Get-Process -Id $taskPid -ErrorAction Stop;$taskProcess.Refresh()
 $samples.Add([pscustomobject]@{Seconds=$timer.Elapsed.TotalSeconds;Responding=$taskProcess.Responding;PrivateBytes=$taskProcess.PrivateMemorySize64;Handles=$taskProcess.HandleCount;Threads=$taskProcess.Threads.Count;CpuSeconds=$taskProcess.TotalProcessorTime.TotalSeconds})
 Start-Sleep -Seconds 2
}
$first=$samples[0];$last=$samples[$samples.Count-1]
$report=[pscustomobject]@{Executable=$exe;DurationSeconds=$timer.Elapsed.TotalSeconds;SampleCount=$samples.Count;UnresponsiveSamples=@($samples|Where-Object{-not $_.Responding}).Count;PrivateDeltaBytes=$last.PrivateBytes-$first.PrivateBytes;HandleDelta=$last.Handles-$first.Handles;Samples=$samples;Limitation='Short controlled sample; not evidence that long-term provider/antivirus stalls cannot occur. No process or security settings changed.'}
$file=Join-Path $root ('resources_'+[IO.Path]::GetFileName([IO.Path]::GetDirectoryName($exe))+'.json')
$report|ConvertTo-Json -Depth 5|Out-File -LiteralPath $file -Encoding utf8
$report|Select-Object SampleCount,UnresponsiveSamples,PrivateDeltaBytes,HandleDelta|ConvertTo-Json
