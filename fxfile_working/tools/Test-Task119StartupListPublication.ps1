[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)] [string]$EvidenceRoot,
    [Parameter(Mandatory=$true)] [string]$PackageRoot,
    [ValidateRange(1,6)] [int]$ExpectedPaneCount=6,
    [ValidateRange(10,120)] [int]$TimeoutSeconds=60,
    [ValidateRange(10,500)] [int]$PollMilliseconds=20,
    [switch]$RunDirect
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$allowed=(Join-Path $workspace '__BUILD_TEMP_BACKUP__').TrimEnd('\')+'\'
$evidence=[IO.Path]::GetFullPath($EvidenceRoot)
$package=[IO.Path]::GetFullPath($PackageRoot)
if(-not $evidence.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw "EvidenceRoot outside workspace: $evidence"}
if(Test-Path -LiteralPath $evidence){throw "Refusing overwrite: $evidence"}
if(-not (Test-Path -LiteralPath (Join-Path $package 'fxfile.exe') -PathType Leaf)){throw "Incomplete package: $package"}
if(@(Get-Process -Name fxfile -ErrorAction SilentlyContinue).Count){throw 'Close FxFile before the test.'}
New-Item -ItemType Directory -Path $evidence | Out-Null
$stage=$package
if(-not $RunDirect){
 $stage=Join-Path $evidence 'package'
 New-Item -ItemType Directory -Path $stage | Out-Null
 Get-ChildItem -LiteralPath $package -Force | ForEach-Object {Copy-Item -LiteralPath $_.FullName -Destination $stage -Recurse -Force}
}
$oldTemp=$env:TEMP; $oldTmp=$env:TMP
$env:TEMP=Join-Path $evidence 'temp'; $env:TMP=$env:TEMP
New-Item -ItemType Directory -Path $env:TEMP | Out-Null
try {
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class FxTask119Native {
 public delegate bool EnumProc(IntPtr hwnd,IntPtr data);
 [StructLayout(LayoutKind.Sequential)] public struct RECT {public int Left,Top,Right,Bottom;}
 [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb,IntPtr data);
 [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr hwnd,EnumProc cb,IntPtr data);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
 [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
 [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd,out RECT rect);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd,StringBuilder value,int max);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern IntPtr GetProp(IntPtr hwnd,string name);
 [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hwnd,uint msg,IntPtr w,IntPtr l);
 [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd,uint msg,IntPtr w,IntPtr l);
 public const uint WM_CLOSE=0x0010, LVM_GETITEMCOUNT=0x1004;
 static string Cls(IntPtr hwnd){var s=new StringBuilder(64);GetClassName(hwnd,s,s.Capacity);return s.ToString();}
 public static IntPtr Frame(int processId){IntPtr best=IntPtr.Zero;int area=0;EnumWindows(delegate(IntPtr h,IntPtr d){uint p;GetWindowThreadProcessId(h,out p);RECT r;if(p==(uint)processId&&IsWindowVisible(h)&&GetWindowRect(h,out r)){int a=(r.Right-r.Left)*(r.Bottom-r.Top);if(a>area){area=a;best=h;}}return true;},IntPtr.Zero);return best;}
 public static IntPtr[] Lists(IntPtr frame){var a=new List<IntPtr>();EnumChildWindows(frame,delegate(IntPtr h,IntPtr d){if(String.Equals(Cls(h),"SysListView32",StringComparison.OrdinalIgnoreCase))a.Add(h);return true;},IntPtr.Zero);return a.ToArray();}
 public static int Count(IntPtr hwnd){return (int)SendMessage(hwnd,LVM_GETITEMCOUNT,IntPtr.Zero,IntPtr.Zero).ToInt64();}
 public static long Prop(IntPtr hwnd,string name){return GetProp(hwnd,name).ToInt64();}
}
'@
$process=$null
try {
 $watch=[Diagnostics.Stopwatch]::StartNew()
 $process=Start-Process -FilePath (Join-Path $stage 'fxfile.exe') -WorkingDirectory $stage -PassThru
 $frame=[IntPtr]::Zero; $frameMs=$null; $skeletonMs=$null; $firstContentPropertyMs=$null; $readyPropertyMs=$null
 $firstAnyMs=$null; $allListsMs=$null; $firstAnyRealMs=$null; $allListsRealMs=$null
 $lists=@(); $firstByHandle=@{}; $firstRealByHandle=@{}
 while($watch.Elapsed.TotalSeconds -lt $TimeoutSeconds){
  $process.Refresh(); if($process.HasExited){throw "FxFile exited: $($process.ExitCode)"}
  if($frame -eq [IntPtr]::Zero){$frame=[FxTask119Native]::Frame($process.Id);if($frame -ne [IntPtr]::Zero){$frameMs=$watch.ElapsedMilliseconds}}
  if($frame -ne [IntPtr]::Zero){
   if($null -eq $skeletonMs -and [FxTask119Native]::Prop($frame,'FxFile.StartupLayoutSkeletonPainted') -eq 1){$skeletonMs=$watch.ElapsedMilliseconds}
   if($null -eq $firstContentPropertyMs -and [FxTask119Native]::Prop($frame,'FxFile.StartupLayoutFirstContentViewCount') -ge $ExpectedPaneCount){$firstContentPropertyMs=$watch.ElapsedMilliseconds}
   if($null -eq $readyPropertyMs -and [FxTask119Native]::Prop($frame,'FxFile.StartupLayoutReadyViewCount') -ge $ExpectedPaneCount){$readyPropertyMs=$watch.ElapsedMilliseconds}
   $lists=@([FxTask119Native]::Lists($frame))
   if($lists.Count -eq $ExpectedPaneCount){
    $nonEmpty=0; $withRealItem=0
    foreach($list in $lists){
     $count=[FxTask119Native]::Count($list)
     if($count -gt 0){
      $nonEmpty++
      $key=[string]$list.ToInt64()
      if(-not $firstByHandle.ContainsKey($key)){$firstByHandle[$key]=[pscustomobject]@{Handle=$list.ToInt64();FirstItemMilliseconds=$watch.ElapsedMilliseconds;FirstObservedCount=$count}}
     }
     # In the saved non-Desktop startup folders row zero is the synthetic
     # [..] parent row. A count greater than one is the first proof that a real
     # filesystem folder/file has been published; accepting count=1 was the
     # Task119 false positive reported by the user.
     if($count -gt 1){
      $withRealItem++
      $key=[string]$list.ToInt64()
      if(-not $firstRealByHandle.ContainsKey($key)){$firstRealByHandle[$key]=[pscustomobject]@{Handle=$list.ToInt64();FirstRealItemMilliseconds=$watch.ElapsedMilliseconds;FirstObservedCount=$count}}
     }
    }
    if($nonEmpty -gt 0 -and $null -eq $firstAnyMs){$firstAnyMs=$watch.ElapsedMilliseconds}
    if($nonEmpty -eq $ExpectedPaneCount -and $null -eq $allListsMs){$allListsMs=$watch.ElapsedMilliseconds}
    if($withRealItem -gt 0 -and $null -eq $firstAnyRealMs){$firstAnyRealMs=$watch.ElapsedMilliseconds}
    if($withRealItem -eq $ExpectedPaneCount -and $null -eq $allListsRealMs){$allListsRealMs=$watch.ElapsedMilliseconds}
    if($null -ne $allListsRealMs -and $null -ne $firstContentPropertyMs -and $null -ne $readyPropertyMs){break}
   }
  }
  Start-Sleep -Milliseconds $PollMilliseconds
 }
 if($null -eq $allListsMs){throw "Not all $ExpectedPaneCount lists published a parent row before timeout; observed=$($firstByHandle.Count)"}
 if($null -eq $allListsRealMs){throw "Not all $ExpectedPaneCount lists published a real folder/file before timeout; observed=$($firstRealByHandle.Count)"}
 $report=[ordered]@{
  Result='PASS';CapturedAt=(Get-Date).ToString('o');RunDirect=[bool]$RunDirect;ExecutableSha256=(Get-FileHash -LiteralPath (Join-Path $stage 'fxfile.exe') -Algorithm SHA256).Hash
  FrameVisibleMilliseconds=$frameMs;SkeletonPaintedMilliseconds=$skeletonMs;FirstContentPropertyMilliseconds=$firstContentPropertyMs;ReadyPropertyMilliseconds=$readyPropertyMs
  FirstAnyListItemMilliseconds=$firstAnyMs;AllListsFirstItemMilliseconds=$allListsMs
  FirstAnyRealItemMilliseconds=$firstAnyRealMs;AllListsRealItemMilliseconds=$allListsRealMs
  AllListsBlankAfterFrameMilliseconds=([double]$allListsRealMs-[double]$frameMs)
  ParentOnlyWindowMilliseconds=([double]$allListsRealMs-[double]$allListsMs)
  ReadyPropertyBeforeVisibleContent=($null -ne $readyPropertyMs -and $readyPropertyMs -lt $allListsRealMs)
  PaneFirstItems=@($firstByHandle.Values|Sort-Object FirstItemMilliseconds)
  PaneFirstRealItems=@($firstRealByHandle.Values|Sort-Object FirstRealItemMilliseconds)
 }
 $report|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $evidence 'startup_list_publication.json') -Encoding utf8
 $report|ConvertTo-Json -Depth 5
} finally {
 if($null -ne $process){try{$process.Refresh();if(-not $process.HasExited){$h=if($process.MainWindowHandle){[IntPtr]$process.MainWindowHandle}else{$frame};if($h -ne [IntPtr]::Zero){[FxTask119Native]::PostMessage($h,[FxTask119Native]::WM_CLOSE,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null};if(-not $process.WaitForExit(10000)){Stop-Process -Id $process.Id -Force}}}catch{}}
}
} finally {$env:TEMP=$oldTemp;$env:TMP=$oldTmp}
