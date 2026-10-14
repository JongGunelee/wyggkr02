[CmdletBinding()]
param([ValidateSet('x64','x32')][string]$Architecture='x64',[switch]$External,[ValidatePattern('^[a-z0-9_]*$')][string]$RunTag='')
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$evidence=Join-Path $workspace '__BUILD_TEMP_BACKUP__/task129_results_20261014'
$suffix=$Architecture;if($External){$suffix+='_external'}
if($RunTag){$suffix+='_'+$RunTag}
$stage=Join-Path $evidence "gui_$suffix";$fixture=Join-Path $evidence "fixture_$suffix"
if(Test-Path -LiteralPath $stage){throw 'Do not overwrite an existing test profile'}
if(@(Get-Process fxfile,'fxfile-launcher' -ErrorAction SilentlyContinue).Count){throw 'Close FxFile first'}
New-Item -ItemType Directory -Path $stage|Out-Null
Get-ChildItem -LiteralPath (Join-Path $workspace "fxfile_run_$Architecture") -File|Where-Object Extension -in '.exe','.dll','.chm'|Copy-Item -Destination $stage
foreach($dir in @('Languages','fxfile')){Copy-Item -LiteralPath (Join-Path $workspace "fxfile_run_$Architecture/$dir") -Destination $stage -Recurse}
function Set-Value($file,$key,$value){
 $text=[IO.File]::ReadAllText($file);$pattern='(?m)^'+[regex]::Escape($key)+'[^\S\r\n]*=[^\r\n]*';$line="$key = $value"
 if([regex]::IsMatch($text,$pattern)){$text=[regex]::Replace($text,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m)$line})}else{$text+="`r`n$line`r`n"}
 [IO.File]::WriteAllText($file,$text,[Text.UnicodeEncoding]::new($false,$true))
}
$main=Join-Path $stage 'fxfile/fxfile-main.conf';$config=Join-Path $stage 'fxfile/fxfile.conf'
foreach($entry in @(@('main.window.position','0,0,1200,900'),@('main.window.status','0'),@('main.window.position_locked','0'),@('main.view.path_locked','0'),@('main.view.split_locked','0'),@('main.view.auto_equalize','0'),@('main.view.row_count','2'),@('main.view.column_count','3'))){Set-Value $main $entry[0] $entry[1]}
Set-Value $config 'config.display.ui_scale_percent' '100'
foreach($op in @('copy','move','delete')){Set-Value $config "config.file_op.external_$op" $(if($External){'1'}else{'0'})}
$paths=@('source','copy_target','move_target','drag_source','drag_target','delete_source')
for($pane=1;$pane -le 6;$pane++){
 $path=Join-Path $fixture $paths[$pane-1];New-Item -ItemType Directory -Path $path -Force|Out-Null
 Set-Value $main "main.view$pane.current_tab" '0';Set-Value $main "main.view$pane.tab1.path" $path;Set-Value $main "main.view$pane.folder_tree.show" '0'
 Set-Value $config "config.view$pane.file_list.init_folder" '1';Set-Value $config "config.view$pane.file_list.init_folder_path" $path
}
$utf8=[Text.UTF8Encoding]::new($false)
foreach($folder in @('source','drag_source','delete_source')){
 $p=Join-Path $fixture $folder;New-Item -ItemType Directory -Path (Join-Path $p '한글 폴더')|Out-Null
 foreach($name in @('alpha.txt','beta.txt','한글 폴더/inside.txt')){[IO.File]::WriteAllText((Join-Path $p $name),"Task129 $suffix $folder $name - synthetic file verification.",$utf8)}
}
[pscustomobject]@{Architecture=$Architecture;External=[bool]$External;Exe=(Join-Path $stage 'fxfile.exe');Fixture=$fixture;Stage=$stage}|ConvertTo-Json
