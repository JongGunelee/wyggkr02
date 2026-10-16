[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PackageRoot,

    [Parameter(Mandatory = $true)]
    [string]$EvidenceRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$workspaceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$allowedBase = [IO.Path]::GetFullPath((Join-Path $workspaceRoot '__BUILD_TEMP_BACKUP__')).TrimEnd('\')
$allowedPrefix = $allowedBase + '\'
$resolvedEvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
$resolvedPackageRoot = [IO.Path]::GetFullPath($PackageRoot)

if (-not ($resolvedEvidenceRoot.Equals($allowedBase, [StringComparison]::OrdinalIgnoreCase) -or
          $resolvedEvidenceRoot.StartsWith($allowedPrefix, [StringComparison]::OrdinalIgnoreCase))) {
    throw "EvidenceRoot must stay below the workspace backup boundary: $allowedPrefix"
}

if (@(Get-Process fxfile, 'fxfile-launcher', 'fxfile-upchecker', 'fxfile-updater' -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'Close every FxFile-related process before preparing the isolated profile.'
}

$sourceExe = Join-Path $resolvedPackageRoot 'fxfile.exe'
if (-not (Test-Path -LiteralPath $sourceExe -PathType Leaf)) {
    throw "Package root is missing fxfile.exe: $resolvedPackageRoot"
}

function Get-TextEncoding([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xff -and $bytes[1] -eq 0xfe) {
        return [Text.Encoding]::Unicode
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xfe -and $bytes[1] -eq 0xff) {
        return [Text.Encoding]::BigEndianUnicode
    }
    return [Text.UTF8Encoding]::new($false)
}

function Set-ConfigValue([string]$Path, [string]$Key, [string]$Value) {
    $encoding = Get-TextEncoding $Path
    $text = [IO.File]::ReadAllText($Path, $encoding)
    # Keep the original CRLF terminator intact.  In .NET multiline mode, `.*$`
    # also consumes the carriage return, which turns every replaced CRLF line
    # into LF-only text.  FxFile's legacy reader then sees adjacent replaced
    # lines as one logical line and only the first setting is loaded.
    $pattern = '(?m)^' + [regex]::Escape($Key) + '\s*=[^\r\n]*'
    $replacement = "$Key = $Value"
    if ([regex]::IsMatch($text, $pattern)) {
        $text = [regex]::Replace($text, $pattern, $replacement)
    }
    else {
        if (-not $text.EndsWith("`r`n")) { $text += "`r`n" }
        $text += "$replacement`r`n"
    }
    [IO.File]::WriteAllText($Path, $text, $encoding)
}

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$runRoot = Join-Path $resolvedEvidenceRoot "task092_row_focus_visual_$stamp"
$scenarioRoot = Join-Path $runRoot 'x64_distinct_six_pane_colors'
New-Item -ItemType Directory -Path $runRoot -Force | Out-Null
Copy-Item -LiteralPath $resolvedPackageRoot -Destination $scenarioRoot -Recurse

$exePath = Join-Path $scenarioRoot 'fxfile.exe'
$configPath = Join-Path $scenarioRoot 'fxfile\fxfile.conf'
$mainConfigPath = Join-Path $scenarioRoot 'fxfile\fxfile-main.conf'
if (-not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
    throw "Staged profile is missing fxfile.conf: $configPath"
}
if (-not (Test-Path -LiteralPath $mainConfigPath -PathType Leaf)) {
    throw "Staged profile is missing fxfile-main.conf: $mainConfigPath"
}
$pointerPath = Join-Path $scenarioRoot 'fxfile.ini'
$isolatedConfigRoot = Join-Path $scenarioRoot 'fxfile'
$pointerText = "# Task092 isolated visual-test configuration pointer`r`n`r`n[.fxfile]`r`nconf_home = $isolatedConfigRoot`r`n"
[IO.File]::WriteAllText($pointerPath, $pointerText, [Text.Encoding]::Unicode)

$colors = @(
    '224,64,64',
    '64,160,64',
    '48,96,224',
    '224,160,48',
    '160,64,192',
    '32,176,176'
)

Set-ConfigValue $configPath 'config.file_list.full_row_select' '1'
Set-ConfigValue $configPath 'config.activate_bar.active_color' '255,0,255'
Set-ConfigValue $mainConfigPath 'main.view.path_locked' '0'
Set-ConfigValue $mainConfigPath 'main.view.split_locked' '0'
Set-ConfigValue $mainConfigPath 'main.view.row_count' '2'
Set-ConfigValue $mainConfigPath 'main.view.column_count' '3'
Set-ConfigValue $mainConfigPath 'main.view.locked_row_count' '2'
Set-ConfigValue $mainConfigPath 'main.view.locked_column_count' '3'
Set-ConfigValue $mainConfigPath 'main.view.ratio1' '0.333000'
Set-ConfigValue $mainConfigPath 'main.view.ratio2' '0.333000'
Set-ConfigValue $mainConfigPath 'main.view.ratio3' '0.500000'
Set-ConfigValue $mainConfigPath 'main.view.size1' '500'
Set-ConfigValue $mainConfigPath 'main.view.size2' '500'
Set-ConfigValue $mainConfigPath 'main.view.size3' '350'
Set-ConfigValue $mainConfigPath 'main.view.locked_ratio1' '0.333000'
Set-ConfigValue $mainConfigPath 'main.view.locked_ratio2' '0.333000'
Set-ConfigValue $mainConfigPath 'main.view.locked_ratio3' '0.500000'
Set-ConfigValue $mainConfigPath 'main.view.locked_size1' '500'
Set-ConfigValue $mainConfigPath 'main.view.locked_size2' '500'
Set-ConfigValue $mainConfigPath 'main.view.locked_size3' '350'

for ($view = 1; $view -le 6; ++$view) {
    Set-ConfigValue $configPath "config.view$view.file_list.row_focus_color" $colors[$view - 1]
    Set-ConfigValue $configPath "config.view$view.file_list.init_folder" '1'
    Set-ConfigValue $configPath "config.view$view.file_list.init_folder_path" 'C:\Windows\Temp'
    Set-ConfigValue $mainConfigPath "main.view$view.folder_tree.show" '0'
    Set-ConfigValue $mainConfigPath "main.view$view.current_tab" '0'
    Set-ConfigValue $mainConfigPath "main.view$view.tab1.path" 'C:\Windows\Temp'
}

foreach ($path in @($configPath, $mainConfigPath)) {
    $encoding = Get-TextEncoding $path
    $text = [IO.File]::ReadAllText($path, $encoding)
    if ([regex]::IsMatch($text, "(?<!`r)`n")) {
        throw "Generated visual-test profile contains an LF-only line: $path"
    }
}

for ($view = 1; $view -le 6; ++$view) {
    $key = "config.view$view.file_list.row_focus_color"
    $text = [IO.File]::ReadAllText($configPath, (Get-TextEncoding $configPath))
    if ([regex]::Matches($text, '(?m)^' + [regex]::Escape($key) + '\s*=').Count -ne 1) {
        throw "Generated visual-test profile must contain exactly one '$key' entry."
    }
}

$launchArguments = [Collections.Generic.List[string]]::new()
$launchArguments.Add('--window')
$launchArguments.Add('2x3')
for ($view = 1; $view -le 6; ++$view) {
    $launchArguments.Add("--dir$view")
    $launchArguments.Add($exePath)
}

$result = [ordered]@{
    Result = 'PASS'
    CreatedAt = (Get-Date).ToString('o')
    RunRoot = $runRoot
    ScenarioRoot = $scenarioRoot
    Executable = $exePath
    ExecutableSha256 = (Get-FileHash -LiteralPath $exePath -Algorithm SHA256).Hash
    Layout = '2x3'
    FullRowFocus = $true
    ExpectedColors = $colors
    SelectionTarget = $exePath
    LaunchArguments = $launchArguments.ToArray()
    IsolatedConfigPointer = $pointerPath
    UserPackagesModified = $false
}
$result | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $runRoot 'profile_manifest.json') -Encoding UTF8
$result | ConvertTo-Json -Depth 4
