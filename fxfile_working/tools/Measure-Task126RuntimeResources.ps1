[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$ExecutablePath,
      [ValidateRange(10,900)][int]$DurationSeconds=120)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$boundary=Join-Path $workspace '__BUILD_TEMP_BACKUP__\task126_results_20261012'
$exe=[IO.Path]::GetFullPath($ExecutablePath)
if(-not $exe.StartsWith($boundary+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Only the isolated Task126 executable is allowed.'}
$candidates=@(Get-Process fxfile -ErrorAction SilentlyContinue|Where-Object {$_.Path -eq $exe})
if($candidates.Count -ne 1){throw 'Expected one isolated FxFile process.'}
$taskProcessId=$candidates[0].Id
$samples=[Collections.Generic.List[object]]::new()
$clock=[Diagnostics.Stopwatch]::StartNew()
while($clock.Elapsed.TotalSeconds -lt $DurationSeconds){
    $proc=Get-Process -Id $taskProcessId -ErrorAction Stop
    $proc.Refresh()
    $samples.Add([pscustomobject]@{Seconds=[math]::Round($clock.Elapsed.TotalSeconds,3);Responding=$proc.Responding;PrivateBytes=$proc.PrivateMemorySize64;WorkingSetBytes=$proc.WorkingSet64;Handles=$proc.HandleCount;Threads=$proc.Threads.Count;CpuSeconds=$proc.TotalProcessorTime.TotalSeconds})
    Start-Sleep -Seconds 2
}
$first=$samples[0];$last=$samples[$samples.Count-1]
$report=[pscustomobject]@{Executable=$exe;ProcessId=$taskProcessId;DurationSeconds=$clock.Elapsed.TotalSeconds;SampleCount=$samples.Count;UnresponsiveSamples=@($samples|Where-Object {-not $_.Responding}).Count;PrivateDeltaBytes=$last.PrivateBytes-$first.PrivateBytes;HandleDelta=$last.Handles-$first.Handles;Samples=$samples;Limitation='Short local sample, not proof of absence of long-term leaks or antivirus/provider stalls. No process priority or security setting changed.'}
$reportName='runtime_resources_'+[IO.Path]::GetFileName([IO.Path]::GetDirectoryName($exe))+'_'+(Get-Date -Format 'yyyyMMdd_HHmmss_fff')+'.json'
$report|ConvertTo-Json -Depth 5|Out-File -LiteralPath (Join-Path $boundary $reportName) -Encoding utf8
$report|Select-Object Executable,DurationSeconds,SampleCount,UnresponsiveSamples,PrivateDeltaBytes,HandleDelta|ConvertTo-Json
