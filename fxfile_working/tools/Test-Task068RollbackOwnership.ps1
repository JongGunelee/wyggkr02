[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$EvidenceRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$workspace = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$allowed = (Join-Path $workspace '__BUILD_TEMP_BACKUP__').TrimEnd('\') + '\'
$evidence = [IO.Path]::GetFullPath($EvidenceRoot)
if (-not $evidence.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Evidence root must stay below the workspace backup boundary: $evidence"
}
if ([IO.Path]::GetPathRoot($evidence) -ne 'D:\') {
    throw 'Task068 ownership evidence must use D:.'
}
if (Test-Path -LiteralPath $evidence) {
    throw "Refusing to overwrite an existing evidence root: $evidence"
}
New-Item -ItemType Directory -Path $evidence | Out-Null

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere -PathType Leaf)) {
    throw "vswhere.exe is missing: $vswhere"
}
$vsInstall = (& $vswhere -latest -version '[17.0,18.0)' -products * `
    -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 `
    -property installationPath | Select-Object -First 1)
if (-not $vsInstall) { throw 'Visual Studio 2022 C++ toolset was not found.' }
$vsDevCmd = Join-Path $vsInstall 'Common7\Tools\VsDevCmd.bat'
if (-not (Test-Path -LiteralPath $vsDevCmd -PathType Leaf)) {
    throw "VsDevCmd.bat is missing: $vsDevCmd"
}

$source = Join-Path $PSScriptRoot 'adaptive_rollback_ownership_test.cpp'
$build = Join-Path $evidence 'build'
New-Item -ItemType Directory -Path $build | Out-Null

$results = @()
$reportPath = Join-Path $evidence 'rollback_ownership_report.json'
try {
    foreach ($architecture in @('x64', 'x86')) {
        $architectureBuild = Join-Path $build $architecture
        $caseRoot = Join-Path $evidence ("case_{0}" -f $architecture)
        $exe = Join-Path $architectureBuild 'adaptive_rollback_ownership_test.exe'
        New-Item -ItemType Directory -Path $architectureBuild | Out-Null
        $compileOutput = ''
        $runOutput = ''
        Push-Location $architectureBuild
        try {
            $compileCommand = 'call "{0}" -no_logo -arch={1} -host_arch=x64 >nul && cl.exe /nologo /std:c++17 /EHsc /utf-8 /W4 /Fe:"{2}" "{3}" /link Ole32.lib Shell32.lib Shlwapi.lib User32.lib' -f `
                $vsDevCmd, $architecture, $exe, $source
            $compileOutput = (& cmd.exe /d /s /c $compileCommand 2>&1 | Out-String).Trim()
            if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $exe -PathType Leaf)) {
                throw "$architecture ownership probe compilation failed.`n$compileOutput"
            }
        }
        finally {
            Pop-Location
        }

        $runOutput = (& $exe $caseRoot 2>&1 | Out-String).Trim()
        if ($LASTEXITCODE -ne 0) {
            throw "$architecture ownership probe failed.`n$runOutput"
        }
        if (Test-Path -LiteralPath $caseRoot) {
            throw "$architecture ownership probe left its fixture root behind."
        }
        $results += [ordered]@{
            Architecture = $architecture
            Result = 'PASS'
            Output = $runOutput
        }
    }

    $report = [ordered]@{
        Task = '068-rollback-ownership'
        Result = 'PASS'
        CapturedAt = (Get-Date).ToString('o')
        Architectures = $results
        Verified = @(
            'x64 and x86 operation-owned file identity is deleted',
            'x64 and x86 same path replaced by a different file identity is preserved',
            'x64 and x86 untracked child prevents owned-directory deletion and shell fallback'
        )
    }
    $report | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $reportPath -Encoding utf8
    $report | Format-List
    "Report: $reportPath"
}
catch {
    [ordered]@{
        Task = '068-rollback-ownership'
        Result = 'FAIL'
        CapturedAt = (Get-Date).ToString('o')
        Message = $_.Exception.Message
        CompletedArchitectures = $results
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $reportPath -Encoding utf8
    throw
}
finally {
    if (Test-Path -LiteralPath $build) {
        Remove-Item -LiteralPath $build -Recurse -Force
    }
}
