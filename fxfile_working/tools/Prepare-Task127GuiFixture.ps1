[CmdletBinding()]
param([ValidateSet('x64','x32')][string]$Architecture='x64')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$evidence=Join-Path $workspace '__BUILD_TEMP_BACKUP__/task127_results_20261012'
$stage=Join-Path $evidence "gui_$Architecture"
if(Test-Path -LiteralPath $stage){throw 'Do not overwrite an existing test profile.'}
if(@(Get-Process fxfile,'fxfile-launcher' -ErrorAction SilentlyContinue).Count){throw 'Close FxFile first.'}
$package=Join-Path $workspace "fxfile_run_$Architecture"
New-Item -ItemType Directory -Path $stage|Out-Null
Get-ChildItem -LiteralPath $package -File|Where-Object {$_.Extension -in '.exe','.dll','.chm'}|Copy-Item -Destination $stage
foreach($dir in @('Languages','fxfile')){Copy-Item -LiteralPath (Join-Path $package $dir) -Destination $stage -Recurse}
function Set-Value([string]$file,[string]$key,[string]$value){
 $text=[IO.File]::ReadAllText($file);$pattern='(?m)^'+[regex]::Escape($key)+'[^\S\r\n]*=[^\r\n]*';$line=$key+' = '+$value
 if([regex]::IsMatch($text,$pattern)){$text=[regex]::Replace($text,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m)$line})}else{$text+="`r`n$line`r`n"}
 $text=[regex]::Replace($text,'\r?\n',"`r`n");[IO.File]::WriteAllText($file,$text,[Text.UnicodeEncoding]::new($false,$true))
}
$main=Join-Path $stage 'fxfile/fxfile-main.conf';$config=Join-Path $stage 'fxfile/fxfile.conf'
$source=Join-Path $evidence "fixture_$Architecture/source"
New-Item -ItemType Directory -Path (Join-Path $source 'folder_test') -Force|Out-Null
$utf8=[Text.UTF8Encoding]::new($false)
[IO.File]::WriteAllText((Join-Path $source 'alpha.txt'),'Task127 synthetic alpha.',$utf8)
[IO.File]::WriteAllText((Join-Path $source 'beta.txt'),'Task127 synthetic beta.',$utf8)
[IO.File]::WriteAllText((Join-Path $source 'folder_test/inside.txt'),'Task127 synthetic nested file.',$utf8)
foreach($entry in @(@('main.window.position','0,0,1100,900'),@('main.window.status','0'),@('main.window.position_locked','0'),@('main.view.path_locked','0'),@('main.view.split_locked','0'),@('main.view.row_count','2'),@('main.view.column_count','3'))){Set-Value $main $entry[0] $entry[1]}
# Missing key exercises the production default rather than forcing a test value.
$text=[IO.File]::ReadAllText($main);$text=[regex]::Replace($text,'(?m)^main\.view\.auto_equalize[^\r\n]*\r?\n?','');[IO.File]::WriteAllText($main,$text,[Text.UnicodeEncoding]::new($false,$true))
Set-Value $config 'config.display.ui_scale_percent' '100'
Set-Value $config 'config.file_list.full_row_select' '1'
Set-Value $config 'config.file_list.show_parent_folder' '1'
for($pane=1;$pane -le 6;++$pane){
 $path=$source
 if($pane%2 -eq 0){$path=Join-Path $evidence "fixture_$Architecture/target$pane";New-Item -ItemType Directory -Path $path -Force|Out-Null}
 Set-Value $main "main.view$pane.current_tab" '0';Set-Value $main "main.view$pane.tab1.path" $path;Set-Value $main "main.view$pane.folder_tree.show" '0'
 Set-Value $config "config.view$pane.file_list.init_folder" '1';Set-Value $config "config.view$pane.file_list.init_folder_path" $path
}
[pscustomobject]@{Architecture=$Architecture;Exe=(Join-Path $stage 'fxfile.exe');Source=$source;Stage=$stage}|ConvertTo-Json
