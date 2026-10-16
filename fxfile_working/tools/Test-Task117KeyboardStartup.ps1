[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$EvidenceRoot,
    [Parameter(Mandatory = $true)] [string]$PackageRoot,
    [ValidateRange(1, 6)] [int]$ExpectedPaneCount = 6,
    [ValidateRange(10, 180)] [int]$TimeoutSeconds = 90
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$workspace = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$allowed = (Join-Path $workspace '__BUILD_TEMP_BACKUP__').TrimEnd('\') + '\'
$evidence = [IO.Path]::GetFullPath($EvidenceRoot)
$package = [IO.Path]::GetFullPath($PackageRoot)
if (-not $evidence.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {
    throw "EvidenceRoot is outside the approved task boundary: $evidence"
}
if (-not (Test-Path -LiteralPath (Join-Path $package 'fxfile.exe') -PathType Leaf)) {
    throw "FxFile package is incomplete: $package"
}
if (Test-Path -LiteralPath $evidence) {
    throw "Refusing to overwrite existing evidence: $evidence"
}

New-Item -ItemType Directory -Path $evidence | Out-Null
$stage = Join-Path $evidence 'package'
New-Item -ItemType Directory -Path $stage | Out-Null
Get-ChildItem -LiteralPath $package -Force | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination $stage -Recurse -Force
}

$oldTemp = $env:TEMP
$oldTmp = $env:TMP
$testTemp = Join-Path $evidence 'temp'
New-Item -ItemType Directory -Path $testTemp | Out-Null
$env:TEMP = $testTemp
$env:TMP = $testTemp

try {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class FxKeyboardNative {
    public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct GUITHREADINFO {
        public int cbSize;
        public int flags;
        public IntPtr hwndActive;
        public IntPtr hwndFocus;
        public IntPtr hwndCapture;
        public IntPtr hwndMenuOwner;
        public IntPtr hwndMoveSize;
        public IntPtr hwndCaret;
        public RECT rcCaret;
    }

    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll")] public static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO info);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr data);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, StringBuilder text, int max);
    [DllImport("user32.dll")] public static extern bool IsChild(IntPtr parent, IntPtr child);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd, int command);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] public static extern void keybd_event(byte key, byte scan, uint flags, UIntPtr extra);

    public const uint WM_CLOSE = 0x0010;
    public const uint LVM_FIRST = 0x1000;
    public const uint LVM_GETNEXTITEM = LVM_FIRST + 12;
    public const int LVNI_FOCUSED = 0x0001;
    public const byte VK_TAB = 0x09;
    public const byte VK_SHIFT = 0x10;
    public const byte VK_DOWN = 0x28;
    public const uint KEYEVENTF_KEYUP = 0x0002;

    public static string ClassName(IntPtr hwnd) {
        var value = new StringBuilder(128);
        GetClassName(hwnd, value, value.Capacity);
        return value.ToString();
    }

    public static IntPtr FocusOf(IntPtr mainWindow) {
        uint processId;
        uint threadId = GetWindowThreadProcessId(mainWindow, out processId);
        var info = new GUITHREADINFO();
        info.cbSize = Marshal.SizeOf(info);
        return GetGUIThreadInfo(threadId, ref info) ? info.hwndFocus : IntPtr.Zero;
    }

    public static IntPtr[] VisibleLists(IntPtr mainWindow) {
        var values = new List<IntPtr>();
        EnumChildWindows(mainWindow, delegate(IntPtr hwnd, IntPtr data) {
            if (IsWindowVisible(hwnd) &&
                String.Equals(ClassName(hwnd), "SysListView32", StringComparison.OrdinalIgnoreCase))
                values.Add(hwnd);
            return true;
        }, IntPtr.Zero);
        return values.ToArray();
    }

    public static void Key(byte key) {
        keybd_event(key, 0, 0, UIntPtr.Zero);
        keybd_event(key, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }

    public static void ShiftTab() {
        keybd_event(VK_SHIFT, 0, 0, UIntPtr.Zero);
        // Give the foreground thread one scheduling quantum to observe the
        // modifier before VK_TAB.  Without this boundary GetAsyncKeyState can
        // intermittently see a forward Tab and wrap to the first pane.
        System.Threading.Thread.Sleep(20);
        Key(VK_TAB);
        System.Threading.Thread.Sleep(20);
        keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
}
'@

    $process = Start-Process -FilePath (Join-Path $stage 'fxfile.exe') `
        -WorkingDirectory $stage -PassThru
    $started = [Diagnostics.Stopwatch]::StartNew()
    $main = [IntPtr]::Zero
    while ($started.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $process.Refresh()
        if ($process.HasExited) { throw "FxFile exited during keyboard startup test: $($process.ExitCode)" }
        if ($process.MainWindowHandle -ne 0) { $main = $process.MainWindowHandle; break }
        Start-Sleep -Milliseconds 100
    }
    if ($main -eq [IntPtr]::Zero) { throw 'FxFile main window did not appear.' }

    [FxKeyboardNative]::ShowWindow($main, 9) | Out-Null
    [FxKeyboardNative]::SetForegroundWindow($main) | Out-Null

    $initialFocus = [IntPtr]::Zero
    $lists = @()
    while ($started.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $lists = @([FxKeyboardNative]::VisibleLists($main))
        $candidate = [FxKeyboardNative]::FocusOf($main)
        if ($lists.Count -eq $ExpectedPaneCount -and
            $candidate -ne [IntPtr]::Zero -and
            [FxKeyboardNative]::ClassName($candidate) -eq 'SysListView32') {
            $initialFocus = $candidate
            break
        }
        [FxKeyboardNative]::SetForegroundWindow($main) | Out-Null
        Start-Sleep -Milliseconds 100
    }
    if ($lists.Count -ne $ExpectedPaneCount) {
        throw "Expected $ExpectedPaneCount visible file lists, observed $($lists.Count)."
    }
    if ($initialFocus -eq [IntPtr]::Zero) {
        $actual = [FxKeyboardNative]::ClassName([FxKeyboardNative]::FocusOf($main))
        throw "Initial keyboard focus was not committed to a file list (class='$actual')."
    }
    # Capture application readiness before injecting any navigation keys.  The
    # former report sampled the stopwatch after the entire Tab loop and thus
    # mixed startup latency with test latency.
    $startupReadyMilliseconds = $started.ElapsedMilliseconds

    # A first arrow key must be meaningful even when asynchronous enumeration
    # finished after the frame obtained keyboard focus.
    [FxKeyboardNative]::Key([FxKeyboardNative]::VK_DOWN)
    Start-Sleep -Milliseconds 150
    $focusedRow = [FxKeyboardNative]::SendMessage(
        $initialFocus, [FxKeyboardNative]::LVM_GETNEXTITEM,
        [IntPtr](-1), [IntPtr][FxKeyboardNative]::LVNI_FOCUSED).ToInt64()
    if ($focusedRow -lt 0) { throw 'Down-arrow did not establish a focused row.' }

    $forwardLists = [Collections.Generic.HashSet[Int64]]::new()
    $forwardSequence = [Collections.Generic.List[object]]::new()
    [void]$forwardLists.Add($initialFocus.ToInt64())
    $forwardSequence.Add([pscustomobject]@{
        Step = 0
        Handle = $initialFocus.ToInt64()
        Class = [FxKeyboardNative]::ClassName($initialFocus)
        FocusedRow = $focusedRow
    })
    for ($i = 1; $i -lt $ExpectedPaneCount; $i++) {
        [FxKeyboardNative]::Key([FxKeyboardNative]::VK_TAB)
        Start-Sleep -Milliseconds 80
        $focus = [FxKeyboardNative]::FocusOf($main)
        $focusClass = [FxKeyboardNative]::ClassName($focus)
        if ($focusClass -ne 'SysListView32') {
            throw "Tab step $i entered '$focusClass' instead of the next file pane."
        }
        $focusRow = [FxKeyboardNative]::SendMessage(
            $focus, [FxKeyboardNative]::LVM_GETNEXTITEM,
            [IntPtr](-1), [IntPtr][FxKeyboardNative]::LVNI_FOCUSED).ToInt64()
        if ($focusRow -ne 0) {
            throw "Tab step $i focused row $focusRow instead of the [..] parent row (0)."
        }
        if (-not $forwardLists.Add($focus.ToInt64())) {
            throw "Tab step $i did not advance to a distinct file pane."
        }
        $forwardSequence.Add([pscustomobject]@{
            Step = $i
            Handle = $focus.ToInt64()
            Class = $focusClass
            FocusedRow = $focusRow
        })
    }
    if ($forwardLists.Count -ne $ExpectedPaneCount) {
        throw "Tab reached only $($forwardLists.Count)/$ExpectedPaneCount file panes."
    }

    $beforeReverse = [FxKeyboardNative]::FocusOf($main)
    [FxKeyboardNative]::ShiftTab()
    Start-Sleep -Milliseconds 80
    $reverseFocus = [FxKeyboardNative]::FocusOf($main)
    $reverseClass = [FxKeyboardNative]::ClassName($reverseFocus)
    $reverseRow = [FxKeyboardNative]::SendMessage(
        $reverseFocus, [FxKeyboardNative]::LVM_GETNEXTITEM,
        [IntPtr](-1), [IntPtr][FxKeyboardNative]::LVNI_FOCUSED).ToInt64()
    $reverseChanged = ($reverseClass -eq 'SysListView32' -and
                       $reverseFocus -ne $beforeReverse -and
                       $reverseRow -eq 0)
    if (-not $reverseChanged) {
        throw "Shift+Tab did not directly reach the previous file pane parent row (class='$reverseClass', row=$reverseRow)."
    }

    $report = [ordered]@{
        Result = 'PASS'
        CapturedAt = (Get-Date).ToString('o')
        ProcessId = $process.Id
        StartupReadyMilliseconds = $startupReadyMilliseconds
        TotalTestMilliseconds = $started.ElapsedMilliseconds
        VisiblePaneCount = $lists.Count
        InitialFocusClass = [FxKeyboardNative]::ClassName($initialFocus)
        ArrowFocusedRow = $focusedRow
        ForwardDistinctFilePanes = $forwardLists.Count
        ForwardOneKeyPerPane = $true
        ForwardSequence = @($forwardSequence)
        ReversePaneReached = $reverseChanged
        ReverseFocusedRow = $reverseRow
        MouseInputInjected = $false
        KeyboardInput = @('Down', 'Tab', 'Shift+Tab')
    }
    $report | ConvertTo-Json -Depth 5 | Set-Content `
        -LiteralPath (Join-Path $evidence 'keyboard_startup_report.json') -Encoding utf8
    $report | ConvertTo-Json -Depth 5
}
finally {
    if (Get-Variable process -ErrorAction SilentlyContinue) {
        try {
            if (-not $process.HasExited -and $process.MainWindowHandle -ne 0) {
                [FxKeyboardNative]::PostMessage($process.MainWindowHandle,
                    [FxKeyboardNative]::WM_CLOSE, [IntPtr]::Zero,
                    [IntPtr]::Zero) | Out-Null
                if (-not $process.WaitForExit(10000)) { Stop-Process -Id $process.Id -Force }
            }
        } catch {}
    }
    $env:TEMP = $oldTemp
    $env:TMP = $oldTmp
}
