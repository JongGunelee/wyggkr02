[CmdletBinding()]
param([ValidateSet('Audit','Delete','Observe')][string]$Mode='Audit')
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$backup=Join-Path $workspace '__BUILD_TEMP_BACKUP__'
$evidence=Join-Path $backup 'task130_results_20261015'
$audit=Join-Path $evidence 'cleanup_audit.json'
function Assert-Idle {
 if(@(Get-Process fxfile,'fxfile-launcher','fxfile-upchecker','fxfile-updater','fxfile-e2e',cl,link,MSBuild,cmake,ole_capture_probe -ErrorAction SilentlyContinue).Count){throw 'Application, probe or build process still active'}
}
function Assert-TaskPath([string]$path) {
 $full=[IO.Path]::GetFullPath($path)
 $inside=$full.StartsWith($workspace+'\',[StringComparison]::OrdinalIgnoreCase)
 $cross=$full.StartsWith('C:\ProgramData\FxFileTask130\',[StringComparison]::OrdinalIgnoreCase) -or $full.StartsWith('C:\Users\ADMIN\AppData\Local\Temp\FxFileTask130\',[StringComparison]::OrdinalIgnoreCase)
 if((!$inside -and !$cross) -or $full -match 'fxfile_backup|__BACKUP_보존용__|task130_before_'){throw "Protected or unscoped path: $full"}
 if(Test-Path -LiteralPath $full){
  $item=Get-Item -LiteralPath $full -Force
  if($item.Attributes -band [IO.FileAttributes]::ReparsePoint){throw "Reparse target: $full"}
  if($item.PSIsContainer -and @(Get-ChildItem -LiteralPath $full -Recurse -Force | Where-Object {$_.Attributes -band [IO.FileAttributes]::ReparsePoint}).Count){throw "Reparse descendant: $full"}
 }
 return $full
}
Assert-Idle
if($Mode -eq 'Audit') {
 $success=@(Get-ChildItem -LiteralPath $backup -Directory -Filter 'unified_deploy_*' | Where-Object { $m=Join-Path $_.FullName 'deployment_manifest.json'; (Test-Path -LiteralPath $m) -and ((Get-Content -LiteralPath $m -Raw|ConvertFrom-Json).Status -eq 'Success') } | Sort-Object Name -Descending)
 if(!$success.Count -or $success[0].Name -notlike 'unified_deploy_20261015_*'){throw 'New Task130 successful deployment required'}
 $keep=$success[0].FullName
 $preflight=Get-ChildItem -LiteralPath $backup -Directory -Filter 'preflight_*'|Sort-Object Name -Descending|Select-Object -First 1
 $candidates=[Collections.Generic.List[string]]::new()
 foreach($name in @('build_cmake','build_cmake_x32','obj','tools/__pycache__','docs/htmlhelp/flyExplorer.chm')){$candidates.Add((Join-Path $workspace "fxfile_working/$name"))}
 foreach($dir in Get-ChildItem -LiteralPath $backup -Directory){
  if(($dir.Name -like 'unified_deploy_*' -and $dir.FullName -ne $keep) -or ($dir.Name -like 'preflight_*' -and $dir.FullName -ne $preflight.FullName) -or $dir.Name -like 'build_temp_*'){$candidates.Add($dir.FullName)}
 }
 $candidates.Add((Join-Path $keep 'smoke'))
 foreach($dir in Get-ChildItem -LiteralPath $evidence -Directory){if($dir.Name -match '^(gui_|fixture_|probe_build$|temp_probe$)'){$candidates.Add($dir.FullName)}}
 foreach($name in @('OleCaptureProbe.ps1','OleCaptureProbe.cpp','compile_probe.cmd','probe_synthetic.txt')){$candidates.Add((Join-Path $evidence $name))}
 foreach($cross in @('C:\ProgramData\FxFileTask130','C:\Users\ADMIN\AppData\Local\Temp\FxFileTask130')){
  if(Test-Path -LiteralPath $cross){foreach($dir in Get-ChildItem -LiteralPath $cross -Directory){if($dir.Name -match '^x(64|32)_c13[0-9]'){$candidates.Add($dir.FullName)}}}
 }
 $rows=foreach($path in $candidates){$full=Assert-TaskPath $path;if(Test-Path -LiteralPath $full){$item=Get-Item -LiteralPath $full;$files=if($item.PSIsContainer){@(Get-ChildItem -LiteralPath $full -File -Recurse -Force)}else{@($item)};[pscustomobject]@{Path=$full;Files=$files.Count;Bytes=[int64](($files|Measure-Object Length -Sum).Sum)}}}
 [pscustomobject]@{CreatedAt=(Get-Date -Format o);ProtectedDeployment=$keep;ProtectedPreflight=$preflight.FullName;Targets=@($rows);EmptyParents=@('C:\ProgramData\FxFileTask130','C:\Users\ADMIN\AppData\Local\Temp\FxFileTask130')}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $audit -Encoding utf8
 $rows|Select-Object Path,Files,Bytes
} elseif($Mode -eq 'Delete') {
 $plan=Get-Content -LiteralPath $audit -Raw|ConvertFrom-Json
 $archive=Join-Path $evidence 'cleanup_archived_manifests';New-Item -ItemType Directory -Path $archive -Force|Out-Null
 $results=foreach($row in $plan.Targets){
  $full=Assert-TaskPath $row.Path
  $failure=''
  if($full -eq $plan.ProtectedDeployment -or $full -eq $plan.ProtectedPreflight){throw 'Protected target'}
  if(Test-Path -LiteralPath $full){
   foreach($file in @('deployment_manifest.json','preflight_report.json')){$src=Join-Path $full $file;if(Test-Path -LiteralPath $src){Copy-Item -LiteralPath $src -Destination (Join-Path $archive ((Split-Path $full -Leaf)+'_'+$file))}}
   try {Remove-Item -LiteralPath $full -Recurse -Force}
   catch {$failure=$_.Exception.Message}
  }
  [pscustomobject]@{Path=$full;Removed=!(Test-Path -LiteralPath $full);Bytes=$row.Bytes;Error=$failure}
 }
 $results|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $evidence 'cleanup_deleted.json') -Encoding utf8
 foreach($parent in $plan.EmptyParents){
  if($parent -notin @('C:\ProgramData\FxFileTask130','C:\Users\ADMIN\AppData\Local\Temp\FxFileTask130')){throw 'Unscoped empty parent'}
  if(Test-Path -LiteralPath $parent){
   $item=Get-Item -LiteralPath $parent -Force
   if($item.Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Reparse parent'}
   if(@(Get-ChildItem -LiteralPath $parent -Force).Count -eq 0){Remove-Item -LiteralPath $parent -Force}
  }
 }
 if(@($results|Where-Object {-not $_.Removed}).Count){throw 'Some audited paths remain locked; see cleanup_deleted.json. No process was stopped.'}
} else {
 $deleted=Join-Path $evidence 'cleanup_deleted.json'
 if(!(Test-Path -LiteralPath $deleted) -or ((Get-Date).ToUniversalTime()-(Get-Item -LiteralPath $deleted).LastWriteTimeUtc).TotalSeconds -lt 30){throw 'Observe at least 30 seconds after the completed deletion ledger'}
 $plan=Get-Content -LiteralPath $audit -Raw|ConvertFrom-Json
 $rows=foreach($row in $plan.Targets){[pscustomobject]@{Path=$row.Path;Exists=(Test-Path -LiteralPath $row.Path)}}
 $rows|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $evidence 'cleanup_observed.json') -Encoding utf8
 if(@($rows|Where-Object Exists).Count){throw 'Cleanup target recreated'}
 'OBSERVATION_PASS'
}
