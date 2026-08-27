param(
    [string]$SourceRoot = (Split-Path -Parent $PSScriptRoot)
)

$ErrorActionPreference = 'Stop'
$source = Get-Content -LiteralPath (Join-Path $SourceRoot 'src\fxfile\adaptive_file_operation.cpp') -Raw
$checks = [ordered]@{
    'copy-only policy' = $source.Contains('aOperation->wFunc != FO_COPY')
    'all top-level sources must be directories' = $source.Contains('allTopLevelSourcesAreDirectories')
    'rename-on-collision excluded' = $source.Contains('FOF_RENAMEONCOLLISION')
    'large file-count threshold' = $source.Contains('kRobocopyManyFileThreshold = 1000')
    'large directory-count threshold' = $source.Contains('kRobocopyManyDirectoryThreshold = 128')
    'large byte threshold' = $source.Contains('kRobocopyLargeTreeThreshold = 2ULL * 1024ULL * 1024ULL * 1024ULL')
    'robocopy resolved from system directory' = $source.Contains('GetSystemDirectoryW')
    'explicit user choice' = $source.Contains('MB_YESNOCANCEL')
    'safe retry limit' = $source.Contains('/R:2 /W:1')
    'storage-aware thread count' = $source.Contains('chooseRobocopyThreads')
    'large-file unbuffered policy' = $source.Contains('aUnbuffered ? L" /J" : L""')
    'junction recursion excluded' = $source.Contains('/XJ')
    'no console window' = $source.Contains('CREATE_NO_WINDOW | CREATE_SUSPENDED')
    'inheritable NUL handles' = $source.Contains('sSecurity.bInheritHandle = TRUE')
    'job kill-on-close' = $source.Contains('JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE')
    'cancel waits for process termination' = $source.Contains('WaitForSingleObject(sProcess.hProcess, 5000)')
    'no unbounded child-process wait' = -not $source.Contains('WaitForSingleObject(sProcess.hProcess, INFINITE)')
    'exit code read only after confirmed termination' = $source.Contains('if (sProcessTerminated &&')
    'robocopy documented success range' = $source.Contains('sExitCode >= 8')
    'post-copy source/target verification' = $source.Contains('sourceSnapshotUnchanged(aPlan, aError)')
    'empty-directory target verification' = $source.Contains('aPlan.directories[i].target.c_str()')
    'created root identity captured before robocopy' = $source.Contains('prepareRobocopyRoots(aOperation, aOwnedRoots, aError)')
    'rollback evidence uses volume and file id' = ($source.Contains('dwVolumeSerialNumber') -and $source.Contains('nFileIndexHigh') -and $source.Contains('nFileIndexLow'))
    'CopyFile2 ownership comes from live destination handle' = ($source.Contains('Info.ChunkStarted.hDestinationFile') -and $source.Contains('readTargetEvidence(sDestination, *sContext->targetPath'))
    'rollback deletes only identity-matched handles' = ($source.Contains('handleMatchesEvidence(sHandle, aEvidence)') -and $source.Contains('SetFileInformationByHandle'))
    'failure rollback collects verified evidence' = ($source.Contains('collectRobocopyTargetEvidence(sPlan, sRobocopyOwnedRoots') -and $source.Contains('rollbackTargets(sPlan, sRobocopyCreatedTargets)'))
    'Robocopy partial children are never path-reopened for deletion' = $source.Contains('Robocopy does not expose destination handles')
    'uncertain termination blocks rollback and fallback' = ($source.IndexOf('if (sRobocopyResult == RobocopyTerminationUncertain)') -ge 0 -and $source.IndexOf('if (sRobocopyResult == RobocopyTerminationUncertain)') -lt $source.IndexOf('collectRobocopyTargetEvidence(sPlan, sRobocopyOwnedRoots'))
    'safe shell fallback after rollback' = $source.Contains('FxFile의 Windows 셸 엔진으로 다시 시도합니다.')
}

$failed = @($checks.GetEnumerator() | Where-Object { -not $_.Value })
foreach ($check in $checks.GetEnumerator()) {
    '{0} {1}' -f ($(if ($check.Value) { 'PASS' } else { 'FAIL' }), $check.Key)
}
if ($failed.Count -ne 0) {
    throw "Task068 Robocopy policy contract failed: $($failed.Count)"
}
"Task068 Robocopy policy contract: $($checks.Count)/$($checks.Count) PASS"
