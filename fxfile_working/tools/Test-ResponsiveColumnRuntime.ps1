[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SmokeRoot,

    [Parameter(Mandatory = $true)]
    [string]$EvidencePath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class FxColumnWin32 {
    public delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr data);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr data);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, StringBuilder text, int max);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
}
'@

function Get-Inventory([string[]]$Roots) {
    $records = @{}
    foreach ($root in $Roots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        $rootItem = Get-Item -LiteralPath $root -Force
        $files = if ($rootItem.PSIsContainer) {
            @(Get-ChildItem -LiteralPath $root -File -Recurse -Force -ErrorAction SilentlyContinue)
        }
        else {
            @($rootItem)
        }
        foreach ($file in $files) {
            $records[$file.FullName.ToLowerInvariant()] = [pscustomobject]@{
                Length = $file.Length
                SHA256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
            }
        }
    }
    return $records
}

function Compare-Inventory($Before, $After) {
    $keys = @($Before.Keys + $After.Keys | Sort-Object -Unique)
    return @($keys | Where-Object {
        -not $Before.ContainsKey($_) -or -not $After.ContainsKey($_) -or
        $Before[$_].Length -ne $After[$_].Length -or $Before[$_].SHA256 -ne $After[$_].SHA256
    })
}

function Get-MainWindow([int]$ProcessId) {
    $found = [IntPtr]::Zero
    $callback = [FxColumnWin32+EnumProc]{
        param([IntPtr]$hwnd, [IntPtr]$data)
        [uint32]$ownerProcessId = 0
        [void][FxColumnWin32]::GetWindowThreadProcessId($hwnd, [ref]$ownerProcessId)
        if ($ownerProcessId -eq $ProcessId -and [FxColumnWin32]::IsWindowVisible($hwnd)) {
            $script:foundMain = $hwnd
            return $false
        }
        return $true
    }
    $script:foundMain = [IntPtr]::Zero
    [void][FxColumnWin32]::EnumWindows($callback, [IntPtr]::Zero)
    return $script:foundMain
}

function Get-ViewSnapshot([IntPtr]$MainWindow) {
    $handles = [Collections.Generic.List[IntPtr]]::new()
    $callback = [FxColumnWin32+EnumProc]{
        param([IntPtr]$hwnd, [IntPtr]$data)
        $name = [Text.StringBuilder]::new(128)
        [void][FxColumnWin32]::GetClassName($hwnd, $name, $name.Capacity)
        if ($name.ToString() -eq 'SysListView32' -and [FxColumnWin32]::IsWindowVisible($hwnd)) {
            $client = [FxColumnWin32+RECT]::new()
            [void][FxColumnWin32]::GetClientRect($hwnd, [ref]$client)
            $header = [FxColumnWin32]::SendMessage($hwnd, 0x101F, [IntPtr]::Zero, [IntPtr]::Zero)
            $columnCount = if ($header -eq [IntPtr]::Zero) { 0 } else {
                [int][FxColumnWin32]::SendMessage($header, 0x1200, [IntPtr]::Zero, [IntPtr]::Zero)
            }
            if (($client.Right - $client.Left) -ge 150 -and ($client.Bottom - $client.Top) -ge 80 -and $columnCount -ge 2) {
                $handles.Add($hwnd)
            }
        }
        return $true
    }
    [void][FxColumnWin32]::EnumChildWindows($MainWindow, $callback, [IntPtr]::Zero)

    return @($handles | ForEach-Object {
        $hwnd = $_
        $window = [FxColumnWin32+RECT]::new()
        $client = [FxColumnWin32+RECT]::new()
        [void][FxColumnWin32]::GetWindowRect($hwnd, [ref]$window)
        [void][FxColumnWin32]::GetClientRect($hwnd, [ref]$client)
        $header = [FxColumnWin32]::SendMessage($hwnd, 0x101F, [IntPtr]::Zero, [IntPtr]::Zero)
        $columnCount = [int][FxColumnWin32]::SendMessage($header, 0x1200, [IntPtr]::Zero, [IntPtr]::Zero)
        $widths = @(for ($i = 0; $i -lt $columnCount; $i++) {
            [int][FxColumnWin32]::SendMessage($hwnd, 0x101D, [IntPtr]$i, [IntPtr]::Zero)
        })
        [pscustomobject]@{
            Handle = $hwnd.ToInt64()
            Left = $window.Left
            Top = $window.Top
            ClientWidth = $client.Right - $client.Left
            ClientHeight = $client.Bottom - $client.Top
            ColumnWidths = $widths
            ColumnTotal = ($widths | Measure-Object -Sum).Sum
        }
    } | Sort-Object Top, Left | Select-Object -First 4)
}

$smoke = [IO.Path]::GetFullPath($SmokeRoot)
$exe = Join-Path $smoke 'fxfile.exe'
$config = Join-Path $smoke 'fxfile\fxfile.conf'
if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { throw "Smoke executable missing: $exe" }
if (-not (Select-String -LiteralPath $config -Pattern '^config\.file_list\.auto_column_width\s*=\s*1\s*$' -Quiet)) {
    throw 'Smoke configuration does not enable automatic column width.'
}
if ((Test-Path -LiteralPath (Join-Path $smoke 'fxfile.ini')) -or
    (Test-Path -LiteralPath (Join-Path $smoke '.fxfile'))) {
    throw 'Smoke package has a forbidden root pointer file.'
}

$productionRoots = @(
    'C:\00 소프트웨어\04 Fxfile\fxfile.exe',
    'C:\00 소프트웨어\04 Fxfile\fxfile',
    'C:\Users\PC\Downloads\0000 FxFile\fxfile_run_x64\fxfile.exe',
    'C:\Users\PC\Downloads\0000 FxFile\fxfile_run_x64\fxfile',
    'C:\Users\PC\Downloads\0000 FxFile\fxfile_run_x32\fxfile.exe',
    'C:\Users\PC\Downloads\0000 FxFile\fxfile_run_x32\fxfile',
    (Join-Path $env:APPDATA 'fxfile')
)
$before = Get-Inventory $productionRoots
$process = $null
$main = [IntPtr]::Zero
try {
    $process = Start-Process -FilePath $exe -WorkingDirectory $smoke -PassThru
    $deadline = (Get-Date).AddSeconds(45)
    do {
        Start-Sleep -Milliseconds 250
        $main = Get-MainWindow $process.Id
        $views = @(if ($main -ne [IntPtr]::Zero) { Get-ViewSnapshot $main })
    } while (($main -eq [IntPtr]::Zero -or $views.Count -ne 4) -and (Get-Date) -lt $deadline -and -not $process.HasExited)
    if ($process.HasExited) { throw "FxFile exited before the runtime test: $($process.ExitCode)" }
    if ($views.Count -ne 4) { throw "Expected four Explorer list views, found $($views.Count)." }

    [void][FxColumnWin32]::SetWindowPos($main, [IntPtr]::Zero, 40, 40, 1100, 800, 0x0014)
    Start-Sleep -Milliseconds 800
    $narrow = @(Get-ViewSnapshot $main)
    [void][FxColumnWin32]::SetWindowPos($main, [IntPtr]::Zero, 40, 40, 1700, 900, 0x0014)
    Start-Sleep -Milliseconds 800
    $wide = @(Get-ViewSnapshot $main)

    if ($narrow.Count -ne 4 -or $wide.Count -ne 4) { throw 'The four-view layout was not preserved during resize.' }
    $changed = 0
    $rows = for ($i = 0; $i -lt 4; $i++) {
        $n = $narrow[$i]
        $w = $wide[$i]
        $sameWidths = (($n.ColumnWidths -join ',') -eq ($w.ColumnWidths -join ','))
        if (-not $sameWidths -and $w.ClientWidth -gt $n.ClientWidth) { $changed++ }
        [pscustomobject]@{
            View = $i + 1
            NarrowClientWidth = $n.ClientWidth
            NarrowColumnWidths = $n.ColumnWidths
            NarrowColumnTotal = $n.ColumnTotal
            WideClientWidth = $w.ClientWidth
            WideColumnWidths = $w.ColumnWidths
            WideColumnTotal = $w.ColumnTotal
            ResponsiveChange = (-not $sameWidths -and $w.ClientWidth -gt $n.ClientWidth)
        }
    }
    if ($changed -lt 3) { throw "Responsive column reflow was observed in only $changed of four views." }

    [void][FxColumnWin32]::PostMessage($main, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
    if (-not $process.WaitForExit(20000)) { throw 'FxFile did not close normally after WM_CLOSE.' }
    if ($process.ExitCode -ne 0) { throw "FxFile exit code was $($process.ExitCode)." }

    $after = Get-Inventory $productionRoots
    $productionChanges = @(Compare-Inventory $before $after)
    if ($productionChanges.Count -ne 0) {
        throw "The isolated GUI test changed production/AppData files: $($productionChanges -join '; ')"
    }
    if ((Test-Path -LiteralPath (Join-Path $smoke 'fxfile.ini')) -or
        (Test-Path -LiteralPath (Join-Path $smoke '.fxfile'))) {
        throw 'The no-INI smoke package created a forbidden root pointer.'
    }

    $report = [ordered]@{
        Result = 'PASS'
        CapturedAt = (Get-Date).ToString('o')
        SmokeRoot = $smoke
        ProcessExitCode = $process.ExitCode
        ViewCount = 4
        ResponsiveViewCount = $changed
        NarrowWindow = '1100x800'
        WideWindow = '1700x900'
        Views = @($rows)
        ProductionAndAppDataChanges = @()
        RootPointerFilesCreated = $false
    }
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $EvidencePath -Encoding UTF8
    $report | ConvertTo-Json -Depth 8
}
finally {
    if ($process -and -not $process.HasExited) {
        if ($main -ne [IntPtr]::Zero) { [void][FxColumnWin32]::PostMessage($main, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) }
        [void]$process.WaitForExit(5000)
    }
}
