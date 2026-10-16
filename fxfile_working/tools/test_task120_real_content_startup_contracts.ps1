$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$worker=Get-Content -LiteralPath (Join-Path $root 'src\fxfile\directory_enumeration_worker.cpp') -Raw
$frame=Get-Content -LiteralPath (Join-Path $root 'src\fxfile\main_frame.cpp') -Raw
$frameH=Get-Content -LiteralPath (Join-Path $root 'src\fxfile\main_frame.h') -Raw
$ctrl=Get-Content -LiteralPath (Join-Path $root 'src\fxfile\explorer_ctrl.cpp') -Raw
$probe=Get-Content -LiteralPath (Join-Path $root 'tools\Test-Task119StartupListPublication.ps1') -Raw

$parentBlock=[regex]::Match($ctrl,'(?s)if \(XPR_IS_TRUE\(mOption\.mParentFolder\).*?watchFileChange\(\);').Value
$readyBlock=[regex]::Match($frame,'(?s)void MainFrame::notifyStartupExplorerViewReady.*?\n\}').Value
$checks=[ordered]@{
 'synthetic parent publication is not reported as real content'=($parentBlock.Length -gt 0 -and -not $parentBlock.Contains('notifyStartupExplorerViewFirstContent'))
 'filtered first shell item cannot consume the first-content transition'=$ctrl.Contains('const xpr_sint_t sItemCountBefore = GetItemCount();') -and $ctrl.Contains('GetItemCount() > sItemCountBefore')
 'startup burst can pass a filtered desktop.ini without waiting for batch completion'=$worker.Contains('kInitialEnumerationBatchCount = 8') -and $worker.Contains('kFirstEnumerationBatchSize = 1') -and $worker.Contains('kEnumerationBatchSize = 128')
 'history has a one-shot gate stored in the frame'=$frameH.Contains('mStartupHistoryPosted') -and $frame.Contains('mStartupHistoryPosted = XPR_FALSE;')
 'history cannot be posted by initial pane skeleton publication'= -not ([regex]::Match($frame,'(?s)LRESULT MainFrame::OnDeferredStartupViews.*?\n\}').Value.Contains('WM_DEFERRED_STARTUP_HISTORY'))
 'all completed lists are synchronously painted before history restoration'=($readyBlock.Contains('sCount == getViewCount()') -and $readyBlock.Contains('RDW_UPDATENOW') -and $readyBlock.Contains('PostMessage(WM_DEFERRED_STARTUP_HISTORY'))
 'current lists become keyboard-ready before secondary history work'=($readyBlock.Contains('mStartupKeyboardFocusReady = XPR_TRUE;') -and $readyBlock.IndexOf('mStartupKeyboardFocusReady = XPR_TRUE;') -lt $readyBlock.IndexOf('TM_ID_DEFERRED_STARTUP_HISTORY'))
 'history steps use one-shot timers so input is not starved'=$frame.Contains('kInitialStartupHistoryDelayMilliseconds = 250') -and $frame.Contains('kNextStartupHistoryDelayMilliseconds = 25') -and $frame.Contains('KillTimer(TM_ID_DEFERRED_STARTUP_HISTORY);')
 'direct runtime probe requires a second native row for current non-Desktop fixtures'=$probe.Contains('if($count -gt 1)') -and $probe.Contains('Not all $ExpectedPaneCount lists published a real folder/file')
 'reported blank interval ends at real-row publication'=$probe.Contains('AllListsBlankAfterFrameMilliseconds=([double]$allListsRealMs-[double]$frameMs)')
}
$failed=@($checks.GetEnumerator()|Where-Object {-not $_.Value})
$checks.GetEnumerator()|ForEach-Object {"$($(if($_.Value){'PASS'}else{'FAIL'})): $($_.Key)"}
if($failed.Count){throw "Task120 contracts failed: $($failed.Key -join '; ')"}
"Task120 contracts: PASS ($($checks.Count)/$($checks.Count))"
