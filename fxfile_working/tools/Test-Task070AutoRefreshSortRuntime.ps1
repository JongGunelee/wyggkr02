[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$StageRoot,

    [Parameter(Mandatory = $true)]
    [string]$EvidencePath,

    [ValidateSet('Sorted', 'RefreshOnly', 'NoRefresh')]
    [string]$Mode = 'Sorted',

    [ValidateSet(4, 6)]
    [int]$PaneCount = 4
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

public static class FxTask070ListView
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct LVITEM
    {
        public uint mask;
        public int iItem;
        public int iSubItem;
        public uint state;
        public uint stateMask;
        public IntPtr pszText;
        public int cchTextMax;
        public int iImage;
        public IntPtr lParam;
        public int iIndent;
        public int iGroupId;
        public uint cColumns;
        public IntPtr puColumns;
        public IntPtr piColFmt;
        public int iGroup;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT { public int Left, Top, Right, Bottom; }

    public sealed class View
    {
        public long Handle;
        public int Left;
        public int Top;
        public string[] Names;
    }

    private delegate bool EnumProc(IntPtr hwnd, IntPtr data);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr hwnd, EnumProc callback, IntPtr data);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder text, int maximum);
    [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr OpenProcess(uint access, bool inherit, int processId);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr VirtualAllocEx(IntPtr process, IntPtr address, UIntPtr size, uint allocationType, uint protection);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool VirtualFreeEx(IntPtr process, IntPtr address, UIntPtr size, uint freeType);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool WriteProcessMemory(IntPtr process, IntPtr address, byte[] buffer, int size, out IntPtr written);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool ReadProcessMemory(IntPtr process, IntPtr address, byte[] buffer, int size, out IntPtr read);
    [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);

    private static string[] ReadNames(IntPtr list, int processId)
    {
        const uint processAccess = 0x0438;
        IntPtr process = OpenProcess(processAccess, false, processId);
        if (process == IntPtr.Zero)
            throw new InvalidOperationException("OpenProcess failed: " + Marshal.GetLastWin32Error());

        int structureBytes = Marshal.SizeOf(typeof(LVITEM));
        int textBytes = 2048;
        IntPtr remote = VirtualAllocEx(process, IntPtr.Zero,
            (UIntPtr)(structureBytes + textBytes), 0x3000, 0x04);
        if (remote == IntPtr.Zero)
        {
            CloseHandle(process);
            throw new InvalidOperationException("VirtualAllocEx failed: " + Marshal.GetLastWin32Error());
        }

        var names = new List<string>();
        try
        {
            int count = SendMessage(list, 0x1004, IntPtr.Zero, IntPtr.Zero).ToInt32();
            for (int index = 0; index < count; ++index)
            {
                LVITEM item = new LVITEM();
                item.mask = 0x0001;
                item.iItem = index;
                item.iSubItem = 0;
                item.pszText = IntPtr.Add(remote, structureBytes);
                item.cchTextMax = textBytes / 2;

                IntPtr local = Marshal.AllocHGlobal(structureBytes);
                try
                {
                    Marshal.StructureToPtr(item, local, false);
                    byte[] structure = new byte[structureBytes];
                    Marshal.Copy(local, structure, 0, structureBytes);
                    IntPtr transferred;
                    if (!WriteProcessMemory(process, remote, structure, structureBytes, out transferred))
                        throw new InvalidOperationException("WriteProcessMemory failed: " + Marshal.GetLastWin32Error());
                    SendMessage(list, 0x1073, (IntPtr)index, remote);
                    byte[] text = new byte[textBytes];
                    if (!ReadProcessMemory(process, IntPtr.Add(remote, structureBytes), text, text.Length, out transferred))
                        throw new InvalidOperationException("ReadProcessMemory failed: " + Marshal.GetLastWin32Error());
                    names.Add(Encoding.Unicode.GetString(text).Split('\0')[0]);
                }
                finally
                {
                    Marshal.FreeHGlobal(local);
                }
            }
        }
        finally
        {
            VirtualFreeEx(process, remote, UIntPtr.Zero, 0x8000);
            CloseHandle(process);
        }
        return names.ToArray();
    }

    public static View[] GetViews(IntPtr frame, int processId, int maximumViews)
    {
        var handles = new List<IntPtr>();
        EnumChildWindows(frame, delegate(IntPtr hwnd, IntPtr unused)
        {
            var className = new StringBuilder(64);
            RECT client;
            GetClassName(hwnd, className, className.Capacity);
            if (IsWindowVisible(hwnd) && className.ToString() == "SysListView32" &&
                GetClientRect(hwnd, out client) && client.Right - client.Left >= 150 &&
                client.Bottom - client.Top >= 80)
            {
                IntPtr header = SendMessage(hwnd, 0x101F, IntPtr.Zero, IntPtr.Zero);
                if (header != IntPtr.Zero && SendMessage(header, 0x1200, IntPtr.Zero, IntPtr.Zero).ToInt32() >= 2)
                    handles.Add(hwnd);
            }
            return true;
        }, IntPtr.Zero);

        return handles.Select(delegate(IntPtr hwnd)
        {
            RECT window;
            GetWindowRect(hwnd, out window);
            return new View {
                Handle = hwnd.ToInt64(), Left = window.Left, Top = window.Top,
                Names = ReadNames(hwnd, processId)
            };
        }).OrderBy(view => view.Top).ThenBy(view => view.Left).Take(maximumViews).ToArray();
    }
}
'@

function Get-ViewNames([Diagnostics.Process]$Process) {
    $Process.Refresh()
    if ($Process.HasExited -or $Process.MainWindowHandle -eq [IntPtr]::Zero) { return @() }
    @([FxTask070ListView]::GetViews($Process.MainWindowHandle, $Process.Id, $PaneCount))
}

function Test-ExpectedOrder([object[]]$Views, [string[]]$Expected) {
    if ($Views.Count -ne $PaneCount) { return $false }
    $tracked = @('a_new.txt', 'b_renamed.txt', 'm_middle.txt', 'z_anchor.txt')
    foreach ($view in $Views) {
        $actual = @($view.Names | Where-Object { $_ -in $tracked })
        if (($actual -join '|') -cne ($Expected -join '|')) { return $false }
    }
    return $true
}

function Wait-ExpectedOrder([Diagnostics.Process]$Process, [string[]]$Expected, [int]$Seconds = 15) {
    $deadline = [DateTime]::UtcNow.AddSeconds($Seconds)
    do {
        $views = @(Get-ViewNames $Process)
        if (Test-ExpectedOrder $views $Expected) { return $views }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline -and -not $Process.HasExited)
    $snapshot = @(Get-ViewNames $Process)
    $summary = @($snapshot | ForEach-Object { $_.Names -join ', ' }) -join ' / '
    throw "The $PaneCount panes did not reach expected order '$($Expected -join ', ')' within $Seconds seconds. Actual: $summary"
}

$stage = [IO.Path]::GetFullPath($StageRoot)
$exe = Join-Path $stage 'fxfile.exe'
$config = Join-Path $stage 'fxfile\fxfile.conf'
if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { throw "Staged executable missing: $exe" }
$expectedNoRefresh = if ($Mode -eq 'NoRefresh') { 1 } else { 0 }
$expectedRefreshSort = if ($Mode -eq 'Sorted') { 1 } else { 0 }
if (-not (Select-String -LiteralPath $config -Pattern ("^config\.refresh\.no\s*=\s*{0}\s*$" -f $expectedNoRefresh) -Quiet)) {
    throw "The staged no-refresh value does not match mode $Mode."
}
if (-not (Select-String -LiteralPath $config -Pattern ("^config\.refresh\.sort\s*=\s*{0}\s*$" -f $expectedRefreshSort) -Quiet)) {
    throw "The staged refresh-sort value does not match mode $Mode."
}
if (Select-String -LiteralPath $config -Pattern '^config\.file_list\.no_sort\s*=\s*1\s*$' -Quiet) {
    throw 'The staged profile disables sorting.'
}

$evidenceFull = [IO.Path]::GetFullPath($EvidencePath)
$evidenceRoot = Split-Path -Parent $evidenceFull
New-Item -ItemType Directory -Path $evidenceRoot -Force | Out-Null
$testRoot = Join-Path $evidenceRoot ('task070_runtime_data_' + (Get-Date -Format 'yyyyMMdd_HHmmss_fff'))
$process = $null
$forced = $false
$report = [ordered]@{
    Result = 'FAIL'
    CapturedAt = (Get-Date).ToString('o')
    Mode = $Mode
    PaneCount = $PaneCount
    StageRoot = $stage
    ExecutableSha256 = (Get-FileHash -LiteralPath $exe -Algorithm SHA256).Hash
    ViewCount = 0
    InitialOrder = @()
    CreateOrder = @()
    RenameOrder = @()
    DeleteOrder = @()
    ForcedTermination = $false
    TestDataRemoved = $false
}

try {
    New-Item -ItemType Directory -Path $testRoot | Out-Null
    $panePaths = @()
    foreach ($index in 1..$PaneCount) {
        $panePath = Join-Path $testRoot "pane$index"
        New-Item -ItemType Directory -Path $panePath | Out-Null
        New-Item -ItemType File -Path (Join-Path $panePath 'm_middle.txt') | Out-Null
        New-Item -ItemType File -Path (Join-Path $panePath 'z_anchor.txt') | Out-Null
        $panePaths += $panePath
    }

    $split = if ($PaneCount -eq 6) { '2x3' } else { '2x2' }
    $directoryArguments = for ($index = 0; $index -lt $PaneCount; ++$index) {
        '--dir{0} "{1}"' -f ($index + 1), $panePaths[$index]
    }
    $arguments = '-w {0} {1}' -f $split, ($directoryArguments -join ' ')
    $oldCompat = $env:__COMPAT_LAYER
    try {
        $env:__COMPAT_LAYER = 'RunAsInvoker'
        $process = Start-Process -FilePath $exe -WorkingDirectory $stage -ArgumentList $arguments -PassThru
    }
    finally {
        $env:__COMPAT_LAYER = $oldCompat
    }

    $initial = @(Wait-ExpectedOrder $process @('m_middle.txt', 'z_anchor.txt') 45)
    $report.ViewCount = $initial.Count
    $report.InitialOrder = @($initial | ForEach-Object { @($_.Names) })

    foreach ($panePath in $panePaths) {
        New-Item -ItemType File -Path (Join-Path $panePath 'a_new.txt') | Out-Null
    }

    if ($Mode -eq 'NoRefresh') {
        Start-Sleep -Seconds 2
        $unchanged = @(Get-ViewNames $process)
        if (-not (Test-ExpectedOrder $unchanged @('m_middle.txt', 'z_anchor.txt'))) {
            throw 'NoRefresh mode unexpectedly changed one or more visible panes.'
        }
        $report.CreateOrder = @($unchanged | ForEach-Object { @($_.Names) })
        $report.RenameOrder = @($unchanged | ForEach-Object { @($_.Names) })
        $report.DeleteOrder = @($unchanged | ForEach-Object { @($_.Names) })
    }
    else {
        $createExpected = if ($Mode -eq 'Sorted') {
            @('a_new.txt', 'm_middle.txt', 'z_anchor.txt')
        } else {
            @('m_middle.txt', 'z_anchor.txt', 'a_new.txt')
        }
        $created = @(Wait-ExpectedOrder $process $createExpected)
        $report.CreateOrder = @($created | ForEach-Object { @($_.Names) })

        foreach ($panePath in $panePaths) {
            Move-Item -LiteralPath (Join-Path $panePath 'z_anchor.txt') -Destination (Join-Path $panePath 'b_renamed.txt')
        }
        $renameExpected = if ($Mode -eq 'Sorted') {
            @('a_new.txt', 'b_renamed.txt', 'm_middle.txt')
        } else {
            @('m_middle.txt', 'b_renamed.txt', 'a_new.txt')
        }
        $renamed = @(Wait-ExpectedOrder $process $renameExpected)
        $report.RenameOrder = @($renamed | ForEach-Object { @($_.Names) })

        foreach ($panePath in $panePaths) {
            Remove-Item -LiteralPath (Join-Path $panePath 'a_new.txt')
        }
        $deleteExpected = if ($Mode -eq 'Sorted') {
            @('b_renamed.txt', 'm_middle.txt')
        } else {
            @('m_middle.txt', 'b_renamed.txt')
        }
        $deleted = @(Wait-ExpectedOrder $process $deleteExpected)
        $report.DeleteOrder = @($deleted | ForEach-Object { @($_.Names) })
    }

    if (-not [FxTask070ListView]::PostMessage($process.MainWindowHandle, 0x0111, [IntPtr]30120, [IntPtr]::Zero)) {
        throw 'The isolated FxFile process did not accept its normal exit command.'
    }
    if (-not $process.WaitForExit(90000)) { throw 'The isolated FxFile process did not exit normally.' }
    if ($process.ExitCode -ne 0) { throw "The isolated FxFile exit code was $($process.ExitCode)." }
    if ((Test-Path -LiteralPath (Join-Path $stage 'fxfile.ini')) -or
        (Test-Path -LiteralPath (Join-Path $stage '.fxfile'))) {
        throw 'The isolated run created a forbidden root configuration pointer.'
    }

    $report.Result = 'PASS'
}
finally {
    if ($null -ne $process -and -not $process.HasExited) {
        if ($process.MainWindowHandle -ne [IntPtr]::Zero) {
            [void][FxTask070ListView]::PostMessage($process.MainWindowHandle, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
        }
        if (-not $process.WaitForExit(10000)) {
            Stop-Process -Id $process.Id -Force
            $process.WaitForExit()
            $forced = $true
        }
    }
    $report.ForcedTermination = $forced
    if (Test-Path -LiteralPath $testRoot -PathType Container) {
        $resolved = [IO.Path]::GetFullPath($testRoot)
        if ($resolved.StartsWith(([IO.Path]::GetFullPath($evidenceRoot) + [IO.Path]::DirectorySeparatorChar),
                                 [StringComparison]::OrdinalIgnoreCase) -and
            [IO.Path]::GetFileName($resolved).StartsWith('task070_runtime_data_', [StringComparison]::Ordinal)) {
            [IO.Directory]::Delete($resolved, $true)
        }
    }
    $report.TestDataRemoved = -not (Test-Path -LiteralPath $testRoot)
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $evidenceFull -Encoding UTF8
}

$report | ConvertTo-Json -Depth 8
if ($report.Result -ne 'PASS' -or $report.ForcedTermination -or -not $report.TestDataRemoved) { exit 1 }
