[CmdletBinding()]
param([ValidateSet('x64','x32')][string]$Architecture='x64',[switch]$Locked,[switch]$AutoOn,[switch]$FixedSizes)
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
if([Diagnostics.Process]::GetProcessesByName('fxfile').Count){throw 'Close FxFile first'}
$result=Join-Path $workspace '__BUILD_TEMP_BACKUP__/task128_results_20261013'
$name="gui_$Architecture";if($Locked){$name+='_locked'}
if($AutoOn){$name+='_on'}
if($FixedSizes){$name+='_fixed'}
$stage=Join-Path $result $name
if(Test-Path -LiteralPath $stage){throw 'Do not overwrite existing fixture'}
New-Item -ItemType Directory -Path $stage -Force|Out-Null
$package=Join-Path $workspace "fxfile_run_$Architecture"
Get-ChildItem -LiteralPath $package -File | Where-Object Extension -in '.exe','.dll','.chm' | Copy-Item -Destination $stage
foreach($dir in @('fxfile','Languages')){Copy-Item -LiteralPath (Join-Path $package $dir) -Destination $stage -Recurse}
$folder=Join-Path $result 'empty';New-Item -ItemType Directory -Path $folder -Force|Out-Null
$main=Join-Path $stage 'fxfile/fxfile-main.conf';$config=Join-Path $stage 'fxfile/fxfile.conf'
function Set-Value($file,$key,$value){$text=[IO.File]::ReadAllText($file);$pattern='(?m)^'+[regex]::Escape($key)+'[^\S\r\n]*=[^\r\n]*';$line="$key = $value";if([regex]::IsMatch($text,$pattern)){$text=[regex]::Replace($text,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m)$line})}else{$text+="`r`n$line`r`n"};[IO.File]::WriteAllText($file,$text,[Text.UnicodeEncoding]::new($false,$true))}
foreach($pair in @(@('main.window.position','0,0,1100,900'),@('main.window.status','0'),@('main.window.position_locked','0'),@('main.view.path_locked','0'),@('main.view.auto_equalize','0'),@('main.view.split_locked',$(if($Locked){'1'}else{'0'})),@('main.view.row_count','2'),@('main.view.column_count','3'),@('main.view.ratio1','0.4'),@('main.view.ratio2','0.25'),@('main.view.ratio3','0.65'),@('main.view.size1','432'),@('main.view.size2','270'),@('main.view.size3','504'),@('main.view.locked_row_count','2'),@('main.view.locked_column_count','3'),@('main.view.locked_ratio1','0.25'),@('main.view.locked_ratio2','0.4'),@('main.view.locked_ratio3','0.35'))){Set-Value $main $pair[0] $pair[1]}
if($AutoOn){Set-Value $main 'main.view.auto_equalize' '1'}
Set-Value $config 'config.display.ui_scale_percent' '100';Set-Value $config 'config.view_split.resize_by_ratio' '1'
if($FixedSizes){Set-Value $config 'config.view_split.resize_by_ratio' '0'}
for($pane=1;$pane -le 6;$pane++){Set-Value $main "main.view$pane.current_tab" '0';Set-Value $main "main.view$pane.tab1.path" $folder;Set-Value $main "main.view$pane.folder_tree.show" '0';Set-Value $config "config.view$pane.file_list.init_folder" '1';Set-Value $config "config.view$pane.file_list.init_folder_path" $folder}
[pscustomobject]@{Stage=$stage;Exe=(Join-Path $stage 'fxfile.exe')}|ConvertTo-Json
