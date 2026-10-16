[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [Parameter(Mandatory=$true)][string]$PackageRoot,
    [ValidateRange(1,6)][int]$ExpectedPaneCount=6,
    [ValidateRange(10,180)][int]$TimeoutSeconds=90
)

Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$allowed=(Join-Path $workspace '__BUILD_TEMP_BACKUP__').TrimEnd('\')+'\'
$evidence=[IO.Path]::GetFullPath($EvidenceRoot)
$package=[IO.Path]::GetFullPath($PackageRoot)
if(-not $evidence.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw "EvidenceRoot is outside the approved task boundary: $evidence"}
if(Test-Path -LiteralPath $evidence){throw "Refusing to overwrite existing evidence: $evidence"}
if(-not (Test-Path -LiteralPath (Join-Path $package 'fxfile.exe') -PathType Leaf)){throw "FxFile package is incomplete: $package"}

New-Item -ItemType Directory -Path $evidence | Out-Null
$stage=Join-Path $evidence 'package'
New-Item -ItemType Directory -Path $stage | Out-Null
Get-ChildItem -LiteralPath $package -Force | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $stage -Recurse -Force }
$oldTemp=$env:TEMP; $oldTmp=$env:TMP
$testTemp=Join-Path $evidence 'temp'; New-Item -ItemType Directory -Path $testTemp | Out-Null
$env:TEMP=$testTemp; $env:TMP=$testTemp

try {
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class FxTask121Native {
 public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr data);
 [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left,Top,Right,Bottom; }
 [StructLayout(LayoutKind.Sequential)] public struct GUITHREADINFO { public int cbSize,flags; public IntPtr hwndActive,hwndFocus,hwndCapture,hwndMenuOwner,hwndMoveSize,hwndCaret; public RECT rcCaret; }
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
 [DllImport("user32.dll")] public static extern bool GetGUIThreadInfo(uint tid,ref GUITHREADINFO info);
 [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent,EnumWindowsProc callback,IntPtr data);
 [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd,StringBuilder text,int max);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd,StringBuilder text,int max);
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
 [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd,int command);
 [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hwnd,uint message,IntPtr wParam,IntPtr lParam);
 [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd,uint message,IntPtr wParam,IntPtr lParam);
 [DllImport("user32.dll")] public static extern void keybd_event(byte key,byte scan,uint flags,UIntPtr extra);
 public const uint WM_CLOSE=0x0010,LVM_FIRST=0x1000,LVM_GETITEMCOUNT=LVM_FIRST+4,LVM_GETNEXTITEM=LVM_FIRST+12;
 public const int LVNI_FOCUSED=1,LVNI_SELECTED=2;
 public const byte VK_HOME=0x24,VK_DOWN=0x28,VK_RETURN=0x0D; public const uint KEYEVENTF_KEYUP=2;
 public static string ClassName(IntPtr hwnd){var s=new StringBuilder(128);GetClassName(hwnd,s,s.Capacity);return s.ToString();}
 public static string Title(IntPtr hwnd){var s=new StringBuilder(2048);GetWindowText(hwnd,s,s.Capacity);return s.ToString();}
 public static IntPtr FocusOf(IntPtr main){uint pid;uint tid=GetWindowThreadProcessId(main,out pid);var i=new GUITHREADINFO();i.cbSize=Marshal.SizeOf(i);return GetGUIThreadInfo(tid,ref i)?i.hwndFocus:IntPtr.Zero;}
 public static IntPtr[] VisibleLists(IntPtr main){var v=new List<IntPtr>();EnumChildWindows(main,delegate(IntPtr h,IntPtr d){if(IsWindowVisible(h)&&String.Equals(ClassName(h),"SysListView32",StringComparison.OrdinalIgnoreCase))v.Add(h);return true;},IntPtr.Zero);return v.ToArray();}
 public static long ItemCount(IntPtr list){return SendMessage(list,LVM_GETITEMCOUNT,IntPtr.Zero,IntPtr.Zero).ToInt64();}
 public static long Next(IntPtr list,int flag){return SendMessage(list,LVM_GETNEXTITEM,new IntPtr(-1),new IntPtr(flag)).ToInt64();}
 public static void Key(byte key){keybd_event(key,0,0,UIntPtr.Zero);keybd_event(key,0,KEYEVENTF_KEYUP,UIntPtr.Zero);}
}
'@

$process=Start-Process -FilePath (Join-Path $stage 'fxfile.exe') -WorkingDirectory $stage -PassThru
$clock=[Diagnostics.Stopwatch]::StartNew(); $main=[IntPtr]::Zero; $list=[IntPtr]::Zero
while($clock.Elapsed.TotalSeconds -lt $TimeoutSeconds){
 $process.Refresh(); if($process.HasExited){throw "FxFile exited during Task121 test: $($process.ExitCode)"}
 if($process.MainWindowHandle -ne 0){$main=$process.MainWindowHandle; break}; Start-Sleep -Milliseconds 50
}
if($main -eq [IntPtr]::Zero){throw 'FxFile main window did not appear.'}
[FxTask121Native]::ShowWindow($main,9)|Out-Null; [FxTask121Native]::SetForegroundWindow($main)|Out-Null
while($clock.Elapsed.TotalSeconds -lt $TimeoutSeconds){
 $lists=@([FxTask121Native]::VisibleLists($main)); $candidate=[FxTask121Native]::FocusOf($main)
 if($lists.Count -eq $ExpectedPaneCount -and [FxTask121Native]::ClassName($candidate) -eq 'SysListView32' -and [FxTask121Native]::ItemCount($candidate) -gt 1){$list=$candidate;break}
 Start-Sleep -Milliseconds 50
}
if($list -eq [IntPtr]::Zero){throw "A ready focused list with at least one real item was not found."}

$startupSelected=[FxTask121Native]::Next($list,[FxTask121Native]::LVNI_SELECTED)
$startupFocused=[FxTask121Native]::Next($list,[FxTask121Native]::LVNI_FOCUSED)
if($startupSelected -lt 0 -or $startupFocused -lt 0){throw "Startup list lacks committed selection/focus (selected=$startupSelected focused=$startupFocused)."}

$beforeEnter=[FxTask121Native]::Title($main)
[FxTask121Native]::Key([FxTask121Native]::VK_HOME); Start-Sleep -Milliseconds 40
[FxTask121Native]::Key([FxTask121Native]::VK_DOWN); Start-Sleep -Milliseconds 40
[FxTask121Native]::Key([FxTask121Native]::VK_RETURN)
$entered=$false
while($clock.Elapsed.TotalSeconds -lt $TimeoutSeconds){
 $title=[FxTask121Native]::Title($main)
 if($title -ne $beforeEnter -and [FxTask121Native]::ItemCount($list) -gt 0){$entered=$true;break}
 Start-Sleep -Milliseconds 10
}
if(-not $entered){throw 'Folder entry did not complete in time; the configured row 1 may not be a folder.'}
$entrySelected=[FxTask121Native]::Next($list,[FxTask121Native]::LVNI_SELECTED)
$entryFocused=[FxTask121Native]::Next($list,[FxTask121Native]::LVNI_FOCUSED)
if($entrySelected -lt 0 -or $entryFocused -lt 0){throw "Folder entry required an arrow key to expose selection (selected=$entrySelected focused=$entryFocused)."}

$beforeUp=[FxTask121Native]::Title($main)
[FxTask121Native]::Key([FxTask121Native]::VK_HOME); Start-Sleep -Milliseconds 40
[FxTask121Native]::Key([FxTask121Native]::VK_RETURN)
$returned=$false
while($clock.Elapsed.TotalSeconds -lt $TimeoutSeconds){
 $title=[FxTask121Native]::Title($main)
 if($title -ne $beforeUp -and [FxTask121Native]::ItemCount($list) -gt 0){$returned=$true;break}
 Start-Sleep -Milliseconds 10
}
if(-not $returned){throw 'Parent-folder return did not complete in time.'}
$upSelected=[FxTask121Native]::Next($list,[FxTask121Native]::LVNI_SELECTED)
$upFocused=[FxTask121Native]::Next($list,[FxTask121Native]::LVNI_FOCUSED)
if($upSelected -lt 0 -or $upFocused -lt 0){throw "Parent return required an arrow key to expose selection (selected=$upSelected focused=$upFocused)."}

$report=[ordered]@{Result='PASS';CapturedAt=(Get-Date).ToString('o');PackageRoot=$package;VisiblePaneCount=$ExpectedPaneCount;StartupSelectedRow=$startupSelected;StartupFocusedRow=$startupFocused;FolderEntrySelectedRow=$entrySelected;FolderEntryFocusedRow=$entryFocused;ParentReturnSelectedRow=$upSelected;ParentReturnFocusedRow=$upFocused;DirectionKeyAfterNavigationInjected=$false;ElapsedMilliseconds=$clock.ElapsedMilliseconds}
$report|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $evidence 'navigation_selection_report.json') -Encoding utf8
$report|ConvertTo-Json -Depth 4
}
finally {
 if(Get-Variable process -ErrorAction SilentlyContinue){try{if(-not $process.HasExited){[FxTask121Native]::PostMessage($process.MainWindowHandle,[FxTask121Native]::WM_CLOSE,[IntPtr]::Zero,[IntPtr]::Zero)|Out-Null;if(-not $process.WaitForExit(10000)){Stop-Process -Id $process.Id -Force}}}catch{}}
 $env:TEMP=$oldTemp; $env:TMP=$oldTmp
}
