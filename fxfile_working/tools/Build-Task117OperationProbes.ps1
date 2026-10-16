[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$EvidenceRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$workspace = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$allowed = (Join-Path $workspace '__BUILD_TEMP_BACKUP__').TrimEnd('\') + '\'
$evidence = [IO.Path]::GetFullPath($EvidenceRoot)
if (-not $evidence.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {
    throw "EvidenceRoot is outside the approved task boundary: $evidence"
}
if (Test-Path -LiteralPath $evidence) {
    throw "Refusing to overwrite existing probe evidence: $evidence"
}

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere -PathType Leaf)) {
    throw "vswhere.exe was not found: $vswhere"
}
$installation = (& $vswhere -latest -products * `
    -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 `
    -property installationPath | Select-Object -First 1)
if ([string]::IsNullOrWhiteSpace($installation)) {
    throw 'Visual Studio C++ x64/x86 tools were not found.'
}
$vsDevCmd = Join-Path $installation 'Common7\Tools\VsDevCmd.bat'
if (-not (Test-Path -LiteralPath $vsDevCmd -PathType Leaf)) {
    throw "VsDevCmd.bat was not found: $vsDevCmd"
}

$source = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\src\fxfile'))
$adaptiveTest = Join-Path $PSScriptRoot 'adaptive_file_operation_test.cpp'
$modernTest = Join-Path $PSScriptRoot 'modern_shell_file_operation_test.cpp'
$adaptiveSource = Join-Path $source 'adaptive_file_operation.cpp'
$modernSource = Join-Path $source 'modern_shell_file_operation.cpp'

New-Item -ItemType Directory -Path $evidence | Out-Null
$tempRoot = Join-Path $evidence 'temp'
New-Item -ItemType Directory -Path $tempRoot | Out-Null
$oldTemp = $env:TEMP
$oldTmp = $env:TMP
$env:TEMP = $tempRoot
$env:TMP = $tempRoot

$outputs = [ordered]@{}
try {
    foreach ($target in @(
        [pscustomobject]@{ Name='x64'; Arch='x64' },
        [pscustomobject]@{ Name='x32'; Arch='x86' }
    )) {
        $out = Join-Path $evidence $target.Name
        $obj = Join-Path $out 'obj'
        New-Item -ItemType Directory -Path $obj | Out-Null

        $adaptiveImplObj = Join-Path $obj 'adaptive_impl.obj'
        $adaptiveTestObj = Join-Path $obj 'adaptive_test.obj'
        $modernImplObj = Join-Path $obj 'modern_impl.obj'
        $modernTestObj = Join-Path $obj 'modern_test.obj'
        $adaptiveExe = Join-Path $out 'adaptive_file_operation_test.exe'
        $modernExe = Join-Path $out 'modern_shell_file_operation_test.exe'

        $commands = @(
            ('call "{0}" -no_logo -arch={1} -host_arch=x64' -f $vsDevCmd, $target.Arch),
            ('cl.exe /nologo /utf-8 /std:c++17 /EHsc /O2 /DNDEBUG /DFXFILE_ADAPTIVE_STANDALONE /c "{0}" /Fo"{1}"' -f $adaptiveSource, $adaptiveImplObj),
            ('cl.exe /nologo /utf-8 /std:c++17 /EHsc /O2 /DNDEBUG /c "{0}" /Fo"{1}"' -f $adaptiveTest, $adaptiveTestObj),
            ('link.exe /nologo /OUT:"{0}" "{1}" "{2}" Ole32.lib Shell32.lib Shlwapi.lib Advapi32.lib User32.lib' -f $adaptiveExe, $adaptiveImplObj, $adaptiveTestObj),
            ('cl.exe /nologo /utf-8 /std:c++17 /EHsc /O2 /DNDEBUG /DFXFILE_MODERN_SHELL_STANDALONE /c "{0}" /Fo"{1}"' -f $modernSource, $modernImplObj),
            ('cl.exe /nologo /utf-8 /std:c++17 /EHsc /O2 /DNDEBUG /c "{0}" /Fo"{1}"' -f $modernTest, $modernTestObj),
            ('link.exe /nologo /OUT:"{0}" "{1}" "{2}" Ole32.lib Shell32.lib Shlwapi.lib Advapi32.lib User32.lib' -f $modernExe, $modernImplObj, $modernTestObj)
        )
        & $env:ComSpec /d /s /c ($commands -join ' && ')
        if ($LASTEXITCODE -ne 0) {
            throw "Task117 $($target.Name) probe compilation failed with exit code $LASTEXITCODE."
        }
        foreach ($exe in @($adaptiveExe, $modernExe)) {
            if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) {
                throw "Probe output is missing: $exe"
            }
        }
        $outputs[$target.Name] = [ordered]@{
            Adaptive=$adaptiveExe
            AdaptiveSHA256=(Get-FileHash -LiteralPath $adaptiveExe -Algorithm SHA256).Hash
            Modern=$modernExe
            ModernSHA256=(Get-FileHash -LiteralPath $modernExe -Algorithm SHA256).Hash
        }
    }

    $report = [ordered]@{
        Result='PASS'
        CapturedAt=(Get-Date).ToString('o')
        VisualStudio=$installation
        TempRoot=$tempRoot
        Outputs=$outputs
    }
    $report | ConvertTo-Json -Depth 6 | Set-Content `
        -LiteralPath (Join-Path $evidence 'probe_build_report.json') -Encoding utf8
    $report | ConvertTo-Json -Depth 6
}
finally {
    $env:TEMP = $oldTemp
    $env:TMP = $oldTmp
}
