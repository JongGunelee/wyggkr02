[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$evidence=Join-Path $workspace '__BUILD_TEMP_BACKUP__/task127_results_20261012'
$temp=Join-Path $evidence 'probe_temp'
New-Item -ItemType Directory -Path $temp -Force|Out-Null
$savedTemp=$env:TEMP;$savedTmp=$env:TMP
$env:TEMP=$temp;$env:TMP=$temp
$vs='C:/Program Files/Microsoft Visual Studio/2022/Community/Common7/Tools/VsDevCmd.bat'
try{
 foreach($arch in @('x64','x86')){
  $out=Join-Path $evidence "probe_$arch"
  New-Item -ItemType Directory -Path $out -Force|Out-Null
  $source=Join-Path $PSScriptRoot 'equal_split_layout_test.cpp'
  $exe=Join-Path $out 'layout_test.exe';$obj=Join-Path $out 'layout_test.obj'
  $command='call "{0}" -no_logo -arch={1} -host_arch=x64 && cl /nologo /utf-8 /std:c++17 /EHsc /O2 /UNDEBUG "{2}" /Fo"{3}" /Fe"{4}" user32.lib && "{4}"' -f $vs,$arch,$source,$obj,$exe
  $lines=& cmd.exe /d /s /c $command 2>&1
  $code=$LASTEXITCODE
  $lines|Tee-Object -FilePath (Join-Path $out 'result.log')
  if($code){throw "Probe $arch failed: $code"}
 }
}finally{$env:TEMP=$savedTemp;$env:TMP=$savedTmp}
