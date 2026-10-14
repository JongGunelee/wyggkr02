[CmdletBinding()]
param([ValidateSet('x64','x32')][string]$Architecture='x64',[switch]$Reconfigure)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$evidence=Join-Path $workspace '__BUILD_TEMP_BACKUP__\task126_results_20261012'
$stage=Join-Path $evidence "gui_$Architecture"
if((Test-Path -LiteralPath $stage) -and -not $Reconfigure){throw "Stage already exists: $stage"}
if(@(Get-Process fxfile,'fxfile-launcher' -ErrorAction SilentlyContinue).Count){throw 'Close FxFile first.'}
$package=Join-Path $workspace "fxfile_run_$Architecture"
if(-not $Reconfigure){
    New-Item -ItemType Directory -Path $stage|Out-Null
    Get-ChildItem -LiteralPath $package -File|Where-Object {$_.Extension -in '.exe','.dll','.chm'}|Copy-Item -Destination $stage
    foreach($dir in @('Languages','fxfile')){Copy-Item -LiteralPath (Join-Path $package $dir) -Destination $stage -Recurse}
}else{
    foreach($name in @('fxfile-main.conf','fxfile.conf')){Copy-Item -LiteralPath (Join-Path $package "fxfile\$name") -Destination (Join-Path $stage "fxfile\$name") -Force}
}
function Set-Value([string]$file,[string]$key,[string]$value){
    $text=[IO.File]::ReadAllText($file)
    $pattern='(?m)^'+[regex]::Escape($key)+'[^\S\r\n]*=[^\r\n]*'
    $line=$key+' = '+$value
    if([regex]::IsMatch($text,$pattern)){$text=[regex]::Replace($text,$pattern,[System.Text.RegularExpressions.MatchEvaluator]{param($m) $line})}
    else{$text+="`r`n$line`r`n"}
    $text=[regex]::Replace($text,'\r?\n',"`r`n")
    [IO.File]::WriteAllText($file,$text,[Text.UnicodeEncoding]::new($false,$true))
}
$main=Join-Path $stage 'fxfile\fxfile-main.conf'
$config=Join-Path $stage 'fxfile\fxfile.conf'
$source=Join-Path $evidence 'fixture\source'
if(-not (Test-Path -LiteralPath $source)){
    New-Item -ItemType Directory -Path (Join-Path $source 'folder_test') -Force|Out-Null
    # Reproducible synthetic inputs; never overwrite an existing fixture.
    $utf8=[Text.UTF8Encoding]::new($false)
    [IO.File]::WriteAllText((Join-Path $source 'alpha_long_name_for_horizontal_scroll_abcdefghijklmnopqrstuvwxyz_0123456789.txt'),('A'*49),$utf8)
    [IO.File]::WriteAllText((Join-Path $source 'beta_rename.txt'),('B'*71),$utf8)
    [IO.File]::WriteAllText((Join-Path $source 'folder_test\inside.txt'),'Task126 safe nested fixture.',$utf8)
}
Set-Value $main 'main.window.position' '0,0,1100,900'
Set-Value $main 'main.window.status' '0'
Set-Value $main 'main.window.position_locked' '1'
Set-Value $main 'main.view.path_locked' '1'
Set-Value $main 'main.view.split_locked' '1'
Set-Value $main 'main.view.locked_row_count' '2'
Set-Value $main 'main.view.locked_column_count' '3'
Set-Value $main 'main.view.locked_size1' '360'
Set-Value $main 'main.view.locked_size2' '360'
Set-Value $main 'main.view.locked_size3' '340'
Set-Value $config 'config.display.ui_scale_percent' '100'
Set-Value $config 'config.file_list.full_row_select' '1'
Set-Value $config 'config.file_list.show_parent_folder' '1'
Set-Value $config 'config.file_list.auto_column_width' '1'
Set-Value $config 'config.file_list.column_ellipsis_name' '0'
Set-Value $config 'config.drag_drop.default_file_op' '0'
Set-Value $config 'config.single_rename_type' '1'
for($pane=1;$pane -le 6;++$pane){
    $path=$source
    if($pane%2 -eq 0){$path=Join-Path $evidence "fixture\target$pane";New-Item -ItemType Directory -Path $path -Force|Out-Null}
    Set-Value $main "main.view$pane.current_tab" '0'
    Set-Value $main "main.view$pane.tab1.path" $path
    Set-Value $main "main.view$pane.locked_path" $path
    Set-Value $main "main.view$pane.folder_tree.show" '0'
    Set-Value $config "config.view$pane.file_list.init_folder" '1'
    Set-Value $config "config.view$pane.file_list.init_folder_path" $path
}
[pscustomobject]@{Architecture=$Architecture;Stage=$stage;Exe=(Join-Path $stage 'fxfile.exe');Source=$source}|ConvertTo-Json
