[CmdletBinding()]
param([switch]$RestoreSynthetic)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if(@(Get-Process fxfile* -ErrorAction SilentlyContinue).Count){throw 'Close FxFile first.'}
$project=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$workspace=Split-Path -Parent $project
$evidence=Join-Path $workspace '__BUILD_TEMP_BACKUP__/task127_results_20261012'
$out=Join-Path $evidence 'recycle_probe'
New-Item -ItemType Directory -Path $out -Force|Out-Null
$source=Join-Path $PSScriptRoot 'recycle_lookup_probe.cpp'
$obj=Join-Path $out 'probe.obj';$exe=Join-Path $out 'recycle_lookup_probe.exe'
$production=Join-Path $project 'build_cmake/src/fxfile/fxfile.dir/Release/bb_info_file.obj'
$pch=Join-Path $project 'build_cmake/src/fxfile/fxfile.dir/Release/cmake_pch.obj'
$includes=@('src/fxfile','src/xpr/include','src/base','src/fxfile-crash','lib/gfl/include','lib/vld/include','lib/iconv/include','lib/libxml2/include/libxml2')|ForEach-Object{'/I"'+(Join-Path $project $_)+'"'}
$lib=Join-Path $project 'obj/lib/x64/Release'
$vs='C:/Program Files/Microsoft Visual Studio/2022/Community/Common7/Tools/VsDevCmd.bat'
Get-ChildItem (Join-Path $workspace 'fxfile_run_x64') -File -Filter '*.dll'|Copy-Item -Destination $out
$temp=Join-Path $out 'temp';New-Item -ItemType Directory -Path $temp -Force|Out-Null
$oldTemp=$env:TEMP;$oldTmp=$env:TMP;$env:TEMP=$temp;$env:TMP=$temp
try{
 $command='call "{0}" -no_logo -arch=x64 -host_arch=x64 && cl /nologo /utf-8 /std:c++17 /MD /EHsc /O2 /DNDEBUG /DUNICODE /D_UNICODE /D_AFXDLL {1} "{2}" "{3}" "{7}" /Fo"{4}" /Fe"{5}" /link /LIBPATH:"{6}" libxprw.lib Shlwapi.lib Advapi32.lib Shell32.lib Ole32.lib' -f $vs,($includes -join ' '),$source,$production,$obj,$exe,$lib,$pch
 $lines=& cmd.exe /d /s /c $command 2>&1;$code=$LASTEXITCODE
 $lines|Tee-Object -FilePath (Join-Path $out 'build.log')
 if($code){throw "Compilation failed: $code"}
 if($RestoreSynthetic){
  & $exe (Join-Path $evidence 'failed_delete_undo_journal.dat') --restore 2>&1 | Tee-Object -FilePath (Join-Path $out 'restore_result.log')
 }else{
  & $exe (Join-Path $evidence 'failed_delete_undo_journal.dat') 2>&1 | Tee-Object -FilePath (Join-Path $out 'result.log')
 }
 exit $LASTEXITCODE
}finally{$env:TEMP=$oldTemp;$env:TMP=$oldTmp}
