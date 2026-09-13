$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$worker=Get-Content -LiteralPath (Join-Path $root 'src\fxfile\directory_enumeration_worker.cpp') -Raw
$frameH=Get-Content -LiteralPath (Join-Path $root 'src\fxfile\main_frame.h') -Raw
$frame=Get-Content -LiteralPath (Join-Path $root 'src\fxfile\main_frame.cpp') -Raw
$ctrl=Get-Content -LiteralPath (Join-Path $root 'src\fxfile\explorer_ctrl.cpp') -Raw
$probe=Get-Content -LiteralPath (Join-Path $root 'tools\Test-Task119StartupListPublication.ps1') -Raw
$parentBlock=[regex]::Match($ctrl,'(?s)if \(XPR_IS_TRUE\(mOption\.mParentFolder\).*?watchFileChange\(\);').Value
$checks=[ordered]@{
 'first visual batch is one resolved item'=$worker.Contains('const xpr_size_t kFirstEnumerationBatchSize = 1;')
 'steady-state batch remains bounded at 128'=$worker.Contains('const xpr_size_t kEnumerationBatchSize = 128;')
 'worker emits a finite eight-item startup burst before steady state'=$worker.Contains('const xpr_size_t kInitialEnumerationBatchCount = 8;') -and $worker.Contains('sInitialBatchCount < kInitialEnumerationBatchCount ?') -and $worker.Contains('++sInitialBatchCount;')
 'frame publishes separate first-content progress'=$frame.Contains('FxFile.StartupLayoutFirstContentViewCount') -and $frameH.Contains('notifyStartupExplorerViewFirstContent')
 'frame ready progress is completion-owned'=$frameH.Contains('notifyStartupExplorerViewReady') -and -not $frame.Contains('(HANDLE)(INT_PTR)sReadyViewCount')
 'async first content requires native ListView count growth'=$ctrl -match '(?s)sItemCountBefore = GetItemCount\(\).*?GetItemCount\(\) > sItemCountBefore.*?notifyStartupExplorerViewFirstContent'
 'non-desktop async panes publish parent row without claiming real content'=$parentBlock.Length -gt 0 -and $parentBlock.Contains('addParentItem();') -and $parentBlock.Contains('mDirectoryEnumerationParentPublished = XPR_TRUE;') -and -not $parentBlock.Contains('notifyStartupExplorerViewFirstContent')
 'async parent publication advances real item insertion index'=$ctrl.Contains('mDirectoryEnumerationInsertIndex = GetItemCount();')
 'completion prevents duplicate parent row'=$ctrl -match '(?s)mOption\.mParentFolder\) &&\s*XPR_IS_FALSE\(mDirectoryEnumerationParentPublished\)'
 'startup history yields between pane restorations'=$frame.Contains('++mStartupHistoryViewIndex;') -and $frame.Contains('PostMessage(WM_DEFERRED_STARTUP_HISTORY, 0, 0)')
 'startup history starts only after every current list completes and is painted'=$frame -match '(?s)if \(sCount == getViewCount\(\) &&\s*XPR_IS_FALSE\(mStartupHistoryPosted\)\).*?RDW_UPDATENOW.*?PostMessage\(WM_DEFERRED_STARTUP_HISTORY'
 'keyboard readiness follows completed current lists before history'=$frame -match '(?s)if \(sCount == getViewCount\(\).*?mStartupKeyboardFocusReady = XPR_TRUE;\s*requestStartupKeyboardFocus\(\);.*?TM_ID_DEFERRED_STARTUP_HISTORY'
 'async completion publishes final readiness'=$ctrl -match '(?s)postEnumeration\(mDirectoryEnumerationUpdateBuddy\).*?notifyStartupExplorerViewReady'
 'empty folders close both progress contracts'=$ctrl.Contains('if (XPR_IS_TRUE(sWasFirstBatch))')
 'runtime probe distinguishes parent-only from real rows in direct packages'=$probe.Contains('AllListsRealItemMilliseconds') -and $probe.Contains('ParentOnlyWindowMilliseconds') -and $probe.Contains('[switch]$RunDirect')
}
$failed=@($checks.GetEnumerator()|Where-Object {-not $_.Value})
$checks.GetEnumerator()|ForEach-Object {"$($(if($_.Value){'PASS'}else{'FAIL'})): $($_.Key)"}
if($failed.Count){throw "Task119 contracts failed: $($failed.Key -join '; ')"}
"Task119 contracts: PASS ($($checks.Count)/$($checks.Count))"
