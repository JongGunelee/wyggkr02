[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TriggerPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [int]$ExpectedViewCount = 4,
    [int]$TimeoutSeconds = 360,

    [ValidateRange(20, 1000)]
    [int]$PollIntervalMilliseconds = 50
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace FxFileShortcutTiming
{
    public static class NativeMethods
    {
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr GetProp(IntPtr hWnd, string name);

        public static IntPtr FindVisibleFrame(int processId)
        {
            IntPtr result = IntPtr.Zero;
            int largestArea = 0;
            EnumWindows(delegate(IntPtr hWnd, IntPtr unused)
            {
                uint ownerProcessId;
                GetWindowThreadProcessId(hWnd, out ownerProcessId);
                if (ownerProcessId != (uint)processId || !IsWindowVisible(hWnd))
                    return true;

                RECT rect;
                if (!GetWindowRect(hWnd, out rect))
                    return true;

                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;
                int area = width * height;
                if (width >= 600 && height >= 400 && area > largestArea)
                {
                    largestArea = area;
                    result = hWnd;
                }
                return true;
            }, IntPtr.Zero);
            return result;
        }

        public static int GetReadyViewCount(int processId)
        {
            int maxReady = 0;
            EnumWindows(delegate(IntPtr hWnd, IntPtr unused)
            {
                uint ownerProcessId;
                GetWindowThreadProcessId(hWnd, out ownerProcessId);
                if (ownerProcessId != (uint)processId)
                    return true;

                long value = GetProp(hWnd, "FxFile.StartupLayoutReadyViewCount").ToInt64();
                if (value > maxReady)
                    maxReady = (int)value;
                return true;
            }, IntPtr.Zero);
            return maxReady;
        }

        public static bool IsSkeletonPainted(int processId)
        {
            bool painted = false;
            EnumWindows(delegate(IntPtr hWnd, IntPtr unused)
            {
                uint ownerProcessId;
                GetWindowThreadProcessId(hWnd, out ownerProcessId);
                if (ownerProcessId == (uint)processId &&
                    GetProp(hWnd, "FxFile.StartupLayoutSkeletonPainted").ToInt64() == 1)
                {
                    painted = true;
                    return false;
                }
                return true;
            }, IntPtr.Zero);
            return painted;
        }
    }
}
'@

$monitorStartedUtc = [DateTime]::UtcNow
$deadlineUtc = $monitorStartedUtc.AddSeconds($TimeoutSeconds)

while (-not (Test-Path -LiteralPath $TriggerPath -PathType Leaf)) {
    if ([DateTime]::UtcNow -ge $deadlineUtc) {
        throw "Shortcut timing trigger did not arrive within $TimeoutSeconds seconds: $TriggerPath"
    }
    Start-Sleep -Milliseconds $PollIntervalMilliseconds
}

$triggerText = Get-Content -LiteralPath $TriggerPath -Raw
$trigger = $triggerText | ConvertFrom-Json
# PowerShell 7's ConvertFrom-Json can materialize an ISO-8601 string as a local
# DateTime.  Coercing that value back to text loses the original trailing `Z`
# and applies the local UTC offset a second time.  Extract the original JSON
# token so shortcut-click timings remain timezone-neutral.
$clickMatch = [regex]::Match($triggerText, '"ClickIssuedUtc"\s*:\s*"([^"]+)"')
if (-not $clickMatch.Success) {
    throw "ClickIssuedUtc is missing from shortcut timing trigger: $TriggerPath"
}
$clickIssuedUtc = [DateTimeOffset]::Parse(
    $clickMatch.Groups[1].Value,
    [Globalization.CultureInfo]::InvariantCulture,
    [Globalization.DateTimeStyles]::AssumeUniversal).UtcDateTime
$existingPids = @($trigger.ExistingFxFilePids | ForEach-Object { [int]$_ })
$ignoredPids = [Collections.Generic.HashSet[int]]::new()
foreach ($existingPid in $existingPids) {
    [void]$ignoredPids.Add($existingPid)
}
$bootstrapProcesses = [Collections.Generic.List[object]]::new()

$process = $null
$processDetectedUtc = $null
$processStartUtc = $null
$frameVisibleUtc = $null
$skeletonPaintedUtc = $null
$viewReadyUtc = [ordered]@{}
$lastReadyCount = 0

while ([DateTime]::UtcNow -lt $deadlineUtc -and $lastReadyCount -lt $ExpectedViewCount) {
    if ($null -eq $process) {
        $candidates = @(Get-Process -Name 'fxfile' -ErrorAction SilentlyContinue |
            Where-Object { -not $ignoredPids.Contains($_.Id) } |
            Sort-Object StartTime)
        if ($candidates.Count -eq 0) {
            Start-Sleep -Milliseconds $PollIntervalMilliseconds
            continue
        }

        $process = $candidates[0]
        $processDetectedUtc = [DateTime]::UtcNow
        $processStartUtc = $process.StartTime.ToUniversalTime()
        $frameVisibleUtc = $null
        $skeletonPaintedUtc = $null
        $viewReadyUtc = [ordered]@{}
        $lastReadyCount = 0
    }

    $process.Refresh()
    if ($process.HasExited) {
        $bootstrapProcesses.Add([pscustomobject]@{
            ProcessId = $process.Id
            ProcessStartUtc = $processStartUtc.ToString('o')
            ProcessDetectedUtc = $processDetectedUtc.ToString('o')
            ClickToProcessStartMs = [math]::Round(($processStartUtc - $clickIssuedUtc).TotalMilliseconds, 1)
            FrameVisible = $null -ne $frameVisibleUtc
            ReadyViewCount = $lastReadyCount
        })
        [void]$ignoredPids.Add($process.Id)
        $process = $null
        Start-Sleep -Milliseconds $PollIntervalMilliseconds
        continue
    }

    if ($null -eq $frameVisibleUtc) {
        $frame = [FxFileShortcutTiming.NativeMethods]::FindVisibleFrame($process.Id)
        if ($frame -ne [IntPtr]::Zero) {
            $frameVisibleUtc = [DateTime]::UtcNow
        }
    }

    if ($null -eq $skeletonPaintedUtc -and
        [FxFileShortcutTiming.NativeMethods]::IsSkeletonPainted($process.Id)) {
        $skeletonPaintedUtc = [DateTime]::UtcNow
    }

    $readyCount = [FxFileShortcutTiming.NativeMethods]::GetReadyViewCount($process.Id)
    if ($readyCount -gt $lastReadyCount) {
        for ($view = $lastReadyCount + 1; $view -le $readyCount; ++$view) {
            $viewReadyUtc[[string]$view] = [DateTime]::UtcNow
        }
        $lastReadyCount = $readyCount
    }

    Start-Sleep -Milliseconds $PollIntervalMilliseconds
}

$completedUtc = [DateTime]::UtcNow
$processExited = $null -eq $process -or $process.HasExited
$timings = [ordered]@{
    ClickToProcessStartMs = if ($null -ne $processStartUtc) { [math]::Round(($processStartUtc - $clickIssuedUtc).TotalMilliseconds, 1) } else { $null }
    ClickToProcessDetectedMs = if ($null -ne $processDetectedUtc) { [math]::Round(($processDetectedUtc - $clickIssuedUtc).TotalMilliseconds, 1) } else { $null }
    ClickToFrameVisibleMs = if ($null -ne $frameVisibleUtc) { [math]::Round(($frameVisibleUtc - $clickIssuedUtc).TotalMilliseconds, 1) } else { $null }
    ClickToSkeletonPaintedMs = if ($null -ne $skeletonPaintedUtc) { [math]::Round(($skeletonPaintedUtc - $clickIssuedUtc).TotalMilliseconds, 1) } else { $null }
}
foreach ($entry in $viewReadyUtc.GetEnumerator()) {
    $timings["ClickToView$($entry.Key)ReadyMs"] = [math]::Round((([DateTime]$entry.Value) - $clickIssuedUtc).TotalMilliseconds, 1)
}

if ($null -ne $process -and -not $process.HasExited) {
    $process.Refresh()
}
$result = [ordered]@{
    Status = if ($lastReadyCount -ge $ExpectedViewCount) { 'Success' } elseif ($processExited) { 'NoFinalProcess' } else { 'Timeout' }
    MonitorStartedUtc = $monitorStartedUtc.ToString('o')
    ClickIssuedUtc = $clickIssuedUtc.ToString('o')
    ProcessId = if ($null -ne $process) { $process.Id } else { $null }
    ProcessStartUtc = if ($null -ne $processStartUtc) { $processStartUtc.ToString('o') } else { $null }
    BootstrapProcesses = @($bootstrapProcesses)
    ExpectedViewCount = $ExpectedViewCount
    ReadyViewCount = $lastReadyCount
    FrameVisible = $null -ne $frameVisibleUtc
    SkeletonPainted = $null -ne $skeletonPaintedUtc
    RespondingAtEnd = if ($processExited) { $false } else { $process.Responding }
    ProcessCpuSecondsAtEnd = if ($processExited) { $null } else { [math]::Round($process.TotalProcessorTime.TotalSeconds, 3) }
    WorkingSetMBAtEnd = if ($processExited) { $null } else { [math]::Round($process.WorkingSet64 / 1MB, 1) }
    Timings = $timings
    CompletedUtc = $completedUtc.ToString('o')
}

$outputDirectory = Split-Path -Parent $OutputPath
if (-not [string]::IsNullOrEmpty($outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}
$result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $OutputPath -Encoding UTF8

if ($result.Status -ne 'Success') {
    exit 2
}
