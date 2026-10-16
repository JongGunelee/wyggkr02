$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$rc = Get-Content -LiteralPath (Join-Path $root 'src\fxfile\fxfile.rc')
$resource = Get-Content -LiteralPath (Join-Path $root 'src\fxfile\resource.h') -Raw
$main = Get-Content -LiteralPath (Join-Path $root 'src\fxfile\main_frame.cpp') -Raw
$table = Get-Content -LiteralPath (Join-Path $root 'src\fxfile\accel_table.cpp') -Raw
$dialog = Get-Content -LiteralPath (Join-Path $root 'src\fxfile\cmd\accel_table_dlg.cpp') -Raw
$commandStrings = Get-Content -LiteralPath (Join-Path $root 'src\fxfile\command_string_table.cpp') -Raw
$hook = Get-Content -LiteralPath (Join-Path $root 'src\fxfile-keyhook\fxfile-keyhook.cpp') -Raw
$launcher = Get-Content -LiteralPath (Join-Path $root 'src\fxfile-launcher\MainWnd.cpp') -Raw

$start = ($rc | Select-String '^\s*IDR_MAINFRAME\s+ACCELERATORS\s*$').LineNumber
if (-not $start) { throw 'IDR_MAINFRAME accelerator table was not found.' }
$entries = @()
for ($i = $start; $i -lt $rc.Count -and $rc[$i] -notmatch '^END\s*$'; $i++) {
    if ($rc[$i] -match '^\s*(.+?),\s*(ID_[A-Z0-9_]+),\s*(.+?)\s*$') {
        $entries += [pscustomobject]@{ Key=$matches[1].Trim(); Command=$matches[2]; Flags=$matches[3].Trim() }
    }
}
$chords = $entries | ForEach-Object { "$($_.Key)|$($_.Flags -replace '\s','')" }
$duplicateChords = @($chords | Group-Object | Where-Object Count -gt 1)
$commands = @($entries.Command | Sort-Object -Unique)
$missingDefinitions = @($commands | Where-Object {
    $_ -notin @('ID_APP_EXIT','ID_EDIT_COPY','ID_EDIT_CUT','ID_EDIT_PASTE','ID_EDIT_UNDO','ID_FILE_PRINT') -and
    $resource -notmatch "(?m)^\s*#define\s+$([regex]::Escape($_))\s+\d+"
})

$checks = [ordered]@{
    'default accelerator table contains the canonical 60 entries' = ($entries.Count -eq 60)
    'default accelerator table has no conflicting key chords' = ($duplicateChords.Count -eq 0)
    'every non-framework accelerator command has a resource id' = ($missingDefinitions.Count -eq 0)
    'loader rejects a null count output pointer' = $table.Contains('XPR_IS_NULL(aCount)')
    'loader clears count before reading untrusted data' = $table.Contains('*aCount = 0;')
    'loader rejects negative and oversized persisted counts' = $table.Contains('0 <= sLoadedCount && sLoadedCount <= aMaxCount')
    'loader verifies every fixed and variable read length' = (([regex]::Matches($table,'sReadSize !=')).Count -ge 4)
    'loader rejects invalid flags and empty command or key values' = $table.Contains('kAllowedFlags') -and $table.Contains('aAccel[i].cmd == 0') -and $table.Contains('aAccel[i].key == 0')
    'loader rejects persisted duplicate chords' = $table.Contains('aAccel[i].key == aAccel[j].key')
    'writer verifies all write lengths' = (([regex]::Matches($table,'sWrittenSize !=')).Count -eq 4)
    'main frame rejects empty and overflowing accelerator tables' = $main.Contains('aCount <= 0 || aCount > MAX_ACCEL')
    'assignment rejects an empty key and respects MAX_ACCEL' = $dialog.Contains('sCurSel >= 0 && sVirtualKeyCode != 0') -and $dialog.Contains('if (mCount >= MAX_ACCEL)')
    'assignment replaces a conflicting chord instead of duplicating it' = $dialog.Contains('A key chord must have exactly one owner')
    'remove and per-command reset reject missing selections' = $dialog.Contains('if (nComCurSel < 0 || nKeyCurSel < 0)') -and $dialog.Contains('if (sComCurSel < 0)')
    'all five picture dock targets are assignable in shortcut settings' = @(
        'ID_VIEW_PIC_DOCK_ACTIVE','ID_VIEW_PIC_DOCK_PANE_1','ID_VIEW_PIC_DOCK_PANE_2',
        'ID_VIEW_PIC_DOCK_PANE_3','ID_VIEW_PIC_DOCK_PANE_4'
    ).Where({ $commandStrings -notmatch "mCommandString\[$([regex]::Escape($_))\]" }).Count -eq 0
    'text-entry controls remain exempt from application accelerators' = $main.Contains('_tcsicmp(sClassName, XPR_STRING_LITERAL("Edit")) == 0')
    'launcher installs a low-level keyboard hook and checks the configured key' = $hook.Contains('SetWindowsHookEx(WH_KEYBOARD_LL') -and $hook.Contains('pkbdhs->vkCode == g_wVirtualKeyCode')
    'launcher limits the global shortcut to a Windows-key chord' = $hook.Contains('GetAsyncKeyState(VK_LWIN)') -and $hook.Contains('GetAsyncKeyState(VK_RWIN)')
    'launcher loads the configured global shortcut before enabling it' = $launcher.IndexOf('loadSetting();') -lt $launcher.IndexOf('::SetHookKey(mSetting.mVirtualKeyCode);')
}

$failed = @($checks.GetEnumerator() | Where-Object { -not $_.Value })
$checks.GetEnumerator() | ForEach-Object { "$(if ($_.Value) {'PASS'} else {'FAIL'}): $($_.Key)" }
if ($failed.Count) { throw "Task122 shortcut contracts failed: $($failed.Key -join '; ')" }
"Task122 shortcut contracts: PASS ($($checks.Count)/$($checks.Count)); defaults=$($entries.Count); commands=$($commands.Count)"
