[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string]$PackageRoot,
    [ValidateRange(5,60)] [int]$CaptureSeconds=15
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if(@(Get-Process -Name fxfile -ErrorAction SilentlyContinue).Count){throw 'Close FxFile before the trace.'}
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class FxDbWinTrace {
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern IntPtr CreateFileMapping(IntPtr f,IntPtr a,uint p,uint hi,uint lo,string n);
 [DllImport("kernel32.dll",SetLastError=true)] static extern IntPtr MapViewOfFile(IntPtr h,uint a,uint hi,uint lo,UIntPtr size);
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern IntPtr CreateEvent(IntPtr a,bool manual,bool initial,string n);
 [DllImport("kernel32.dll")] static extern uint WaitForSingleObject(IntPtr h,uint ms);
 [DllImport("kernel32.dll")] static extern bool SetEvent(IntPtr h);
 [DllImport("kernel32.dll")] static extern bool UnmapViewOfFile(IntPtr p);
 [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
 const uint PAGE_READWRITE=4, FILE_MAP_READ=4, WAIT_OBJECT_0=0;
 public static string[] Capture(int seconds,Action ready) {
  var map=CreateFileMapping(new IntPtr(-1),IntPtr.Zero,PAGE_READWRITE,0,4096,"DBWIN_BUFFER");
  var view=MapViewOfFile(map,FILE_MAP_READ,0,0,new UIntPtr(4096));
  var bufferReady=CreateEvent(IntPtr.Zero,false,false,"DBWIN_BUFFER_READY");
  var dataReady=CreateEvent(IntPtr.Zero,false,false,"DBWIN_DATA_READY");
  if(map==IntPtr.Zero||view==IntPtr.Zero||bufferReady==IntPtr.Zero||dataReady==IntPtr.Zero) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
  var lines=new List<string>(); ready();
  DateTime end=DateTime.UtcNow.AddSeconds(seconds);
  try { while(DateTime.UtcNow<end){SetEvent(bufferReady);if(WaitForSingleObject(dataReady,100)==WAIT_OBJECT_0){int pid=Marshal.ReadInt32(view);string msg=Marshal.PtrToStringAnsi(IntPtr.Add(view,4));if(msg!=null&&msg.StartsWith("FXFILE_STARTUP"))lines.Add(pid+"\t"+msg.Trim());}} }
  finally {UnmapViewOfFile(view);CloseHandle(map);CloseHandle(bufferReady);CloseHandle(dataReady);}
  return lines.ToArray();
 }
}
'@
$process=$null
$old=$env:FXFILE_STARTUP_TRACE
try {
 $env:FXFILE_STARTUP_TRACE='1'
 $lines=[FxDbWinTrace]::Capture($CaptureSeconds,[Action]{
   $script:process=Start-Process -FilePath (Join-Path $PackageRoot 'fxfile.exe') -WorkingDirectory $PackageRoot -PassThru
 })
 $lines
} finally {
 if($null -ne $process){try{$process.Refresh();if(-not $process.HasExited){$process.CloseMainWindow()|Out-Null;if(-not $process.WaitForExit(10000)){Stop-Process -Id $process.Id -Force}}}catch{}}
 $env:FXFILE_STARTUP_TRACE=$old
}
