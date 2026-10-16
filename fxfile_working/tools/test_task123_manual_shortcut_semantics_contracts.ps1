[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$rcPath = Join-Path $root 'src\fxfile\fxfile.rc'
$manualPath = Join-Path $root 'docs\htmlhelp\Html\shortkey.htm'
$workflowPath = Join-Path $root 'tools\Build-Deploy-Verify.ps1'

function Assert-Contract([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

function Get-KeyLabel([string]$Token) {
    $token = $Token.Trim()
    if ($token -match '^"(.*)"$') { return ($Matches[1] -replace '\\\\', '\') }
    if ($token -match '^VK_F([1-9]|1[0-2])$') { return "F$($Matches[1])" }
    $map = @{
        VK_INSERT='Insert'; VK_DELETE='Delete'; VK_RETURN='Enter'; VK_BACK='Backspace'; VK_TAB='Tab'
        VK_MULTIPLY='Num *'; VK_ADD='Num +'; VK_SUBTRACT='Num -'
        VK_LEFT='왼쪽 화살표'; VK_RIGHT='오른쪽 화살표'; VK_UP='위 화살표'; VK_DOWN='아래 화살표'
    }
    Assert-Contract $map.ContainsKey($token) "Unknown accelerator key token: $token"
    return $map[$token]
}

$rc = Get-Content -LiteralPath $rcPath -Raw -Encoding UTF8
$blockMatch = [regex]::Match($rc, '(?ms)^IDR_MAINFRAME\s+ACCELERATORS\s*\r?\nBEGIN\s*\r?\n(?<body>.*?)^END\s*\r?$')
Assert-Contract $blockMatch.Success 'IDR_MAINFRAME accelerator block was not found.'
$actual = [Collections.Generic.List[string]]::new()
foreach ($line in ($blockMatch.Groups['body'].Value -split "`r?`n")) {
    if ($line -notmatch '^\s*(?<key>"[^"]+"|VK_[A-Z0-9]+)\s*,\s*(?<command>ID_[A-Z0-9_]+)\s*,\s*(?<flags>.+?)\s*$') { continue }
    $keyToken = $Matches.key
    $command = $Matches.command
    $flags = $Matches.flags
    $parts = [Collections.Generic.List[string]]::new()
    if ($flags -match 'CONTROL') { $parts.Add('Ctrl') }
    if ($flags -match 'SHIFT') { $parts.Add('Shift') }
    if ($flags -match 'ALT') { $parts.Add('Alt') }
    $parts.Add((Get-KeyLabel $keyToken))
    $actual.Add((($parts -join '+') + '|' + $command))
}

$manual = Get-Content -LiteralPath $manualPath -Raw -Encoding UTF8
Assert-Contract ($manual -match '<meta\s+charset="utf-8"') 'Shortcut manual must declare UTF-8.'
$documented = [Collections.Generic.List[string]]::new()
foreach ($match in [regex]::Matches($manual, '(?is)<tr\s+data-command="(?<command>ID_[A-Z0-9_]+)"[^>]*>\s*<td>(?<shortcut>.*?)</td>')) {
    $shortcut = [Net.WebUtility]::HtmlDecode(($match.Groups['shortcut'].Value -replace '<[^>]+>', '')).Trim()
    $documented.Add(($shortcut + '|' + $match.Groups['command'].Value))
}

Assert-Contract ($actual.Count -eq 60) "Expected 60 current default accelerators; found $($actual.Count)."
Assert-Contract ($documented.Count -eq $actual.Count) "Manual/RC shortcut count mismatch: manual=$($documented.Count), RC=$($actual.Count)."
$diff = @(Compare-Object -ReferenceObject @($actual | Sort-Object) -DifferenceObject @($documented | Sort-Object))
Assert-Contract ($diff.Count -eq 0) ("Manual/RC shortcut mapping mismatch: " + (($diff | ForEach-Object { "$($_.SideIndicator)$($_.InputObject)" }) -join ', '))

$stale = @('Shift+F2', 'Shift+Ctrl+V')
foreach ($shortcut in $stale) {
    Assert-Contract ($manual -notmatch [regex]::Escape(">$shortcut<")) "Stale shortcut remains in manual: $shortcut"
}

$semanticChecks = @(
    @('Ctrl+G','ID_GO_PATH','경로'), @('Ctrl+T','ID_WINDOW_TAB_NEW','새 탭'),
    @('Ctrl+W','ID_WINDOW_TAB_CLOSE','탭 닫기'), @('Shift+Alt+C','ID_EDIT_FILENAME_COPY','파일명'),
    @('Ctrl+Shift+Alt+C','ID_EDIT_DEV_PATH_COPY','개발자'), @('Alt+왼쪽 화살표','ID_GO_BACK','뒤로'),
    @('Alt+오른쪽 화살표','ID_GO_FORWARD','앞으로'), @('Alt+아래 화살표','ID_GO_SIBLING_DOWN','다음'),
    @('Alt+위 화살표','ID_GO_SIBLING_UP','이전'), @('F2','ID_FILE_RENAME','이름')
)
foreach ($check in $semanticChecks) {
    $pattern = '(?is)<tr\s+data-command="' + [regex]::Escape($check[1]) + '"[^>]*>\s*<td>' +
        [regex]::Escape($check[0]) + '</td>.*?<td>.*?' + [regex]::Escape($check[2]) + '.*?</td>\s*</tr>'
    Assert-Contract ($manual -match $pattern) "Semantic description mismatch for $($check[0]) / $($check[1])."
}
Assert-Contract ($manual -match '사용자.*단축키') 'Manual must explain that user-defined shortcuts override defaults.'

$workflow = Get-Content -LiteralPath $workflowPath -Raw -Encoding UTF8
Assert-Contract ($workflow -match "'fxfile\.chm'") 'Unified deployment does not require fxfile.chm.'
Assert-Contract ($workflow -match 'function\s+Build-HelpArtifact') 'Unified workflow does not build the audited CHM.'
Assert-Contract ($workflow -match "Build-HelpArtifact\s*\r?\n\s*Add-StorageCheckpoint 'AfterHelpBuild'") 'CHM build is not part of the final build workflow.'

Write-Host '[PASS] Task123 manual shortcut semantics contracts (60/60 mappings and CHM workflow).'
