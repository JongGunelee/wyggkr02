[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
function Read([string]$path){[IO.File]::ReadAllText((Join-Path $root $path))}
$checks=[Collections.Generic.List[object]]::new()
function Check([string]$name,[bool]$pass){$checks.Add([pscustomobject]@{Name=$name;Passed=$pass})}
$option=Read 'src/fxfile/option.cpp'
$frame=Read 'src/fxfile/main_frame.cpp'
$rc=Read 'src/fxfile/fxfile.rc'
$cmd=Read 'src/fxfile/cmd/cmd_cfg.cpp'
$batch=Read 'src/fxfile/window_layout_batch.h'
Check 'Persistent default ON key' ($option -match 'main.view.auto_equalize[^\r\n]+mViewAutoEqualize[^\r\n]+XPR_TRUE')
Check 'Menu directly above placement lock' ($rc -match 'ID_TOOL_AUTO_EQUALIZE_VIEWS\s*\r?\n\s*MENUITEM[^\r\n]+ID_TOOL_WINDOW_PLACEMENT_LOCK')
Check 'Checkbox state and command toggle' ($cmd.Contains('isViewAutoEqualize()') -and $cmd.Contains('setViewAutoEqualize(') -and $cmd.Contains('StateCheck'))
Check 'Command router binding' ((Read 'src/fxfile/cmd/router/cmd_command_map.cpp').Contains('AutoEqualizeViewsCommand'))
Check 'Localized command and customizable shortcut exposure' ((Read 'src/fxfile/command_string_table.cpp').Contains('cmd.auto_equalize_views') -and (Read 'src/fxfile/Languages/Korean.xml').Contains('모든 창 크기 맞추기'))
Check 'Explicit refresh requests fit' ((Read 'src/fxfile/cmd/cmd_refresh.cpp').Contains('sMainFrame->autoFitViewSizes();'))
Check 'Background refresh does not equalize' (-not (Read 'src/fxfile/explorer_ctrl.cpp').Contains('autoFitViewSizes'))
Check 'No automatic fit while dragging' ($frame.Contains('XPR_IS_FALSE(mSplitter.isTracking())'))
Check 'Single pure gap-aware geometry implementation' ($frame.Contains('calculateEqualSplitGeometry(') -and (Read 'src/fxfile/equal_split_geometry.h').Contains('gap) * (columns - 1)'))
Check 'Batch propagates native handle and never ends invalid handle' ($batch.Contains('batch = ::DeferWindowPos(') -and $batch.Contains('if (batch && ::EndDeferWindowPos(batch))'))
Check 'Batch native fallback and unchanged rectangle suppression' ($batch.Contains('::SetWindowPos(') -and $batch.Contains('::EqualRect(&current, &rect)'))
Check 'Batch preserves active window' ($batch.Contains('SWP_NOZORDER | SWP_NOACTIVATE'))
foreach($p in @('main_frame.cpp','explorer_view.cpp','splitter.cpp')){Check "Shared safe batch: $p" ((Read "src/fxfile/$p").Contains('WindowLayoutBatch') -and -not (Read "src/fxfile/$p").Contains('::DeferWindowPos('))}
$clipboard=Read 'src/fxfile/clipboard.cpp'
$undo=Read 'src/fxfile/file_op_undo.cpp'
Check 'Paste prefers paths and respects STGMEDIUM ownership' ($clipboard.Contains('Existing filesystem objects already expose paths.') -and -not $clipboard.Contains('::DragFinish(sDrop)'))
Check 'Versioned bounded recycle metadata' ((Read 'src/fxfile/bb_info_file.cpp').Contains('parseRecycleMetadata(') -and (Read 'src/fxfile/recycle_metadata.h').Contains('version == 2'))
Check 'Undo menu never waits for journal worker' ($undo.Contains('std::try_to_lock'))
Check 'Failed undo retains journal' ($undo.Contains('XPR_IS_TRUE(aSucceeded) && mIndex == aIndex'))
Check 'Move journal stores each destination' ($undo.Contains('sFileIo.write(sParsing,'))
Check 'Recycle lookup is prepared on the worker' ((Read 'src/fxfile/file_op_thread.cpp').Contains('FileOpUndo::resolveRecycleUndo(mShFileOpStruct, mRecycleUndoTimes)') -and $undo.Contains('setRecycleUndoTimes(sRecycleTimes)'))
Check 'Undo path arrays preserve embedded and final NUL' ($undo.Contains('OperationPathPairs<xpr_tchar_t> sPaths;') -and (Read 'src/fxfile/operation_path_pairs.h').Contains('std::basic_string<Char>'))
$checks|ForEach-Object{'{0} {1}' -f $(if($_.Passed){'PASS'}else{'FAIL'}),$_.Name}
$bad=@($checks|Where-Object{-not $_.Passed});"SUMMARY PASS=$($checks.Count-$bad.Count) FAIL=$($bad.Count)"
if($bad.Count){exit 1}
