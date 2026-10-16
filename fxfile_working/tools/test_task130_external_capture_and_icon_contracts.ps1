[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$root=Join-Path $PSScriptRoot '../src/fxfile'
$drop=[IO.File]::ReadAllText((Join-Path $root 'gui/DropSource.cpp'))
$icon=[IO.File]::ReadAllText((Join-Path $root 'shell_icon.cpp'))
$shell=[IO.File]::ReadAllText((Join-Path $root 'shell.cpp'))
$confirm=$drop.IndexOf('bool accepted = fxfile::FileOperationConfirmation::confirm')
$release=$drop.IndexOf('if (!::ReleaseCapture()')
$checks=@(
 @('Physical mouse capture released before external modal',($release -ge 0 -and $release -lt $confirm)),
 @('Failed capture release refuses operation', $drop.Contains('if (!::ReleaseCapture() || ::GetCapture() != NULL)')),
 @('No capture reacquisition over consent', -not $drop.Contains('::SetCapture(')),
 @('One native drag; no replay or injected acceptance',(-not $drop.Contains('::DoDragDrop(') -and -not $drop.Contains('BM_CLICK'))),
 @('Consent remains required for external targets',$drop.Contains('return DRAGDROP_S_CANCEL;')),
 @('Frozen window, process, thread and effect revalidated',($drop.Contains('currentThread == targetThread && currentProcess == process') -and $drop.Contains('::WindowFromPoint(point) == target && mEffect == effect'))),
 @('Cancellation never returns drop success',$drop.Contains('hr = accepted ? DRAGDROP_S_DROP : DRAGDROP_S_CANCEL;')),
 @('Worker realizes sparse slot pixels',$icon.Contains('list->ForceImagePresent(index, ILFIP_ALWAYS)')),
 @('Drawable icon verified without optional original-size metadata',($icon.Contains('list->GetIcon(index, ILD_NORMAL, &image)') -and $icon.Contains('if (image) ::DestroyIcon(image);') -and -not $icon.Contains('list->GetOriginalSize('))),
 @('Failed realization does not publish sparse index',($icon.Contains('if (!realizeSystemIconForPainting(sAsyncIcon->mResult.mIconIndex))') -and $icon.Contains('sAsyncIcon->mResult.mIconIndex = -1;'))),
 @('COM image list reference released',$icon.Contains('list->Release();')),
 @('System icon structures zero initialized',([regex]::Matches($shell,'SHFILEINFO sShFileInfo = \{\};').Count -ge 2))
)
$failed=0
foreach($c in $checks){if($c[1]){"PASS $($c[0])"}else{"FAIL $($c[0])";$failed++}}
"SUMMARY PASS=$($checks.Count-$failed) FAIL=$failed"
if($failed){exit 1}
