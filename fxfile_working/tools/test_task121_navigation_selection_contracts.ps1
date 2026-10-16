$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$ctrl=Get-Content -LiteralPath (Join-Path $root 'src\fxfile\explorer_ctrl.cpp') -Raw
$ctrlH=Get-Content -LiteralPath (Join-Path $root 'src\fxfile\explorer_ctrl.h') -Raw

$commit=[regex]::Match($ctrl,'(?s)xpr_bool_t ExplorerCtrl::commitNavigationSelection.*?\n\}').Value
$parent=[regex]::Match($ctrl,'(?s)xpr_bool_t ExplorerCtrl::focusParentFolderRow.*?\n\}').Value
$async=[regex]::Match($ctrl,'(?s)if \(XPR_IS_TRUE\(mOption\.mParentFolder\).*?watchFileChange\(\);').Value
$post=[regex]::Match($ctrl,'(?s)void ExplorerCtrl::postEnumeration.*?mDirectoryEnumerationParentPublished = XPR_FALSE;\s*\}').Value
$focus=[regex]::Match($ctrl,'(?s)void ExplorerCtrl::OnSetFocus.*?if \(XPR_IS_NOT_NULL\(gFrame->mPictureViewer\)\)').Value

$checks=[ordered]@{
 'shared navigation commit is declared once on ExplorerCtrl'=$ctrlH.Contains('xpr_bool_t  commitNavigationSelection(xpr_sint_t aItemIndex);')
 'commit rejects missing rows before native state access'=$commit.Contains('aItemIndex < 0 || aItemIndex >= GetItemCount()')
 'commit clears stale selected and focused states'=$commit.Contains('SetItemState(-1, 0, LVIS_SELECTED | LVIS_FOCUSED);')
 'commit sets selected and focused states together'=$commit.Contains('SetItemState(aItemIndex, LVIS_SELECTED | LVIS_FOCUSED,')
 'commit synchronizes selection mark and cached focus'=$commit.Contains('SetSelectionMark(aItemIndex);') -and $commit.Contains('mFocusedItemIndex = aItemIndex;')
 'commit invalidates old and new focus rows without synchronous repaint'=$commit.Contains('redrawFocusItemChange(sOldFocusedItem, aItemIndex);') -and -not $commit.Contains('UpdateWindow')
 'parent landing validates the synthetic parent item type'=$parent.Contains('sLvItemData->mItemType != IDT_PARENT')
 'async parent publication commits selection before redraw and visibility'=$async.Contains('focusParentFolderRow();') -and $async.IndexOf('focusParentFolderRow();') -lt $async.IndexOf('SetRedraw();')
 'same-folder refresh is excluded from provisional parent selection'=$async.Contains('XPR_IS_FALSE(mRefreshViewStatePending)')
 'completion selection is independent of who inserted the parent row'= -not $post.Contains('sAddedParentItem')
 'completion prefers the go-up child and otherwise falls back to parent or row zero'=$post.Contains('sCommittedSelection = commitNavigationSelection(sFind);') -and $post.Contains('focusParentFolderRow()') -and $post.Contains('commitNavigationSelection(0);')
 'refresh restoration remains the exclusive selection owner'=$post.Contains('if (XPR_IS_FALSE(sRestoredRefreshState))')
 'focus entry commits a visible selection instead of focus-only state'=$focus.Contains('commitNavigationSelection(0);') -and -not $focus.Contains('SetItemState(0, LVIS_FOCUSED, LVIS_FOCUSED);')
 'all panes inherit the fix from the shared ExplorerCtrl implementation'=([regex]::Matches($ctrl,'xpr_bool_t ExplorerCtrl::commitNavigationSelection').Count -eq 1)
}

$failed=@($checks.GetEnumerator()|Where-Object {-not $_.Value})
$checks.GetEnumerator()|ForEach-Object {"$($(if($_.Value){'PASS'}else{'FAIL'})): $($_.Key)"}
if($failed.Count){throw "Task121 contracts failed: $($failed.Key -join '; ')"}
"Task121 contracts: PASS ($($checks.Count)/$($checks.Count))"
