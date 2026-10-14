[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$text=[IO.File]::ReadAllText((Join-Path $PSScriptRoot '../src/fxfile/main_frame.cpp'))
$start=$text.IndexOf('void MainFrame::setViewAutoEqualize(')
$end=$text.IndexOf('void MainFrame::autoFitViewSizes(', $start)
$block=$text.Substring($start,$end-$start)
$lockStart=$text.IndexOf('void MainFrame::setViewSplitLocked(')
$lockEnd=$text.IndexOf('xpr_bool_t MainFrame::isClockLocked(', $lockStart)
$lockBlock=$text.Substring($lockStart,$lockEnd-$lockStart)
$commands=[IO.File]::ReadAllText((Join-Path $PSScriptRoot '../src/fxfile/cmd/cmd_cfg.cpp'))
$checks=@(
 @('Automatic/manual mode releases split lock',$block.Contains('gOpt->mMain.mViewSplitLocked = XPR_FALSE;')),
 @('Already OFF with locked split can still release',$block.Contains('== aEnabled && XPR_IS_FALSE(sUnlockSplit)')),
 @('Live layout shape saved with unlock',$block.Contains('mSplitter.getPaneCount(&gOpt->mMain.mViewSplitRowCount,')),
 @('Lock snapshots are never overwritten',-not($block -match 'mLockedViewSplit\w*\s*(\[.*\])?\s*=')),
 @('Position/path/clock locks unaffected',-not($block -match 'mWindowPlacementLocked|mViewPathLocked|mClockLocked')),
 @('Pending fit cleared before persisted option',($block.IndexOf('mPendingAutomaticPaneFit = XPR_FALSE') -lt $block.IndexOf('saveMainOption()'))),
 @('Background fit remains ON-only',$text.Contains('if (XPR_IS_TRUE(isViewAutoEqualize())) setSplitEqually();')),
 @('Auto fit menu check reads actual option',$commands.Contains('isViewAutoEqualize() ? StateCheck : 0')),
 @('Split lock menu check reads actual option',$commands.Contains('isViewSplitLocked() ? StateCheck : 0')),
 @('Lock mode disables auto fitting before snapshot',($lockBlock.IndexOf('mViewAutoEqualize = XPR_FALSE') -lt $lockBlock.IndexOf('mSplitter.getPaneCount('))),
 @('Lock mode cancels queued fit',$lockBlock.Contains('mPendingAutomaticPaneFit = XPR_FALSE')),
 @('Legacy conflicting flags prefer locked snapshot',$text.Contains('// Legacy conflicting checks: preserve the explicitly locked snapshot.')),
 @('Linker PDB stays in architecture build cache',([IO.File]::ReadAllText((Join-Path $PSScriptRoot '../src/fxfile/CMakeLists.txt')).Contains('PDB_OUTPUT_DIRECTORY "${CMAKE_CURRENT_BINARY_DIR}/link_symbols"')))
)
$failed=0
foreach($check in $checks){if($check[1]){"PASS $($check[0])"}else{"FAIL $($check[0])";$failed++}}
"SUMMARY PASS=$($checks.Count-$failed) FAIL=$failed"
if($failed){exit 1}
