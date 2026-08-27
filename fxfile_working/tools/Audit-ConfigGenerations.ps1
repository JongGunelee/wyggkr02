param(
    [string]$WorkspaceRoot = 'C:\Users\PC\Downloads\0000 FxFile'
)

$ErrorActionPreference = 'Stop'

$paths = @()
$paths += & rg --files -g 'fxfile-main.conf' $WorkspaceRoot

$extraRoots = @(
    'C:\00 소프트웨어\04 Fxfile',
    (Join-Path $env:APPDATA 'fxfile')
)

foreach ($root in $extraRoots) {
    if (Test-Path -LiteralPath $root) {
        $paths += Get-ChildItem -LiteralPath $root -Recurse -File -Filter 'fxfile-main.conf' -ErrorAction SilentlyContinue |
            ForEach-Object { $_.FullName }
    }
}

$paths = $paths | ForEach-Object { [IO.Path]::GetFullPath($_) } | Sort-Object -Unique

function Read-KeyMapHead([string]$Path) {
    $map = @{}
    $reader = [IO.StreamReader]::new($Path, [Text.Encoding]::Unicode, $true)
    try {
        while (($line = $reader.ReadLine()) -ne $null) {
            if ($line -eq '[backward]') {
                break
            }
            if ($line -match '^\s*([^#;\[].*?)\s*=\s*(.*)$') {
                $map[$matches[1].Trim()] = $matches[2]
            }
        }
    }
    finally {
        $reader.Dispose()
    }
    return $map
}

function Read-ConfigMap([string]$Path) {
    $map = @{}
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $map
    }
    $reader = [IO.StreamReader]::new($Path, [Text.Encoding]::Unicode, $true)
    try {
        while (($line = $reader.ReadLine()) -ne $null) {
            if ($line -match '^\s*([^#;\[].*?)\s*=\s*(.*)$') {
                $map[$matches[1].Trim()] = $matches[2]
            }
        }
    }
    finally {
        $reader.Dispose()
    }
    return $map
}

$canonical = @(
    'fxfile.conf',
    'fxfile-main.conf',
    'fxfile-bookmark.conf',
    'fxfile-accel.dat',
    'fxfile-coolbar.dat',
    'fxfile-toolbar.dat',
    'fxfile-folder_layout.conf',
    'fxfile-dlg_state.conf',
    'fxfile-view_set.conf',
    'fxfile-updater.conf'
)

$rows = @()
foreach ($path in $paths) {
    $file = Get-Item -LiteralPath $path
    $dir = $file.DirectoryName
    $mainHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    $main = Read-KeyMapHead $path
    $configPath = Join-Path $dir 'fxfile.conf'
    $config = Read-ConfigMap $configPath

    $views = @()
    foreach ($view in 1..4) {
        $viewPaths = @()
        $pattern = '^main\.view' + $view + '\.tab\d+\.path$'
        foreach ($key in $main.Keys | Where-Object { $_ -match $pattern } | Sort-Object) {
            if ($main[$key]) {
                $viewPaths += $main[$key]
            }
        }
        if ($viewPaths.Count -eq 0) {
            $viewPaths += 'INIT:' + $config['config.view' + $view + '.file_list.init_folder_path']
        }
        $views += 'V' + $view + '=' + ($viewPaths -join ' || ')
    }

    $supportFiles = @()
    foreach ($name in $canonical) {
        $supportPath = Join-Path $dir $name
        if (Test-Path -LiteralPath $supportPath -PathType Leaf) {
            $supportFile = Get-Item -LiteralPath $supportPath
            $supportHash = (Get-FileHash -LiteralPath $supportPath -Algorithm SHA256).Hash.Substring(0, 10)
            $supportFiles += $name + ':' + $supportFile.Length + ':' + $supportHash
        }
    }

    $bookmarkPath = Join-Path $dir 'fxfile-bookmark.conf'
    $bookmarkCount = 0
    if (Test-Path -LiteralPath $bookmarkPath) {
        $bookmarkCount = (Select-String -LiteralPath $bookmarkPath -Encoding Unicode -Pattern '^\s*bookmark\d+\.name\s*=' -ErrorAction SilentlyContinue).Count
        if ($bookmarkCount -eq 0) {
            $bookmarkCount = (Select-String -LiteralPath $bookmarkPath -Encoding Unicode -Pattern '^\s*item\d+\.name\s*=' -ErrorAction SilentlyContinue).Count
        }
    }

    $rows += [PSCustomObject]@{
        Hash          = $mainHash
        Length        = $file.Length
        Time          = $file.LastWriteTime
        Dir           = $dir
        Version       = $main['version']
        Window        = $main['main.window.position']
        Status        = $main['main.window.status']
        Grid          = $main['main.view.row_count'] + 'x' + $main['main.view.column_count']
        Ratio         = $main['main.view.ratio1'] + '/' + $main['main.view.ratio2']
        BookmarkText  = $main['main.bookmark.show_text']
        BookmarkMulti = $main['main.bookmark.multiple_line']
        DriveShow     = $main['main.drive_bar.show']
        DriveMulti    = $main['main.drive_bar.multiple_line']
        Views         = $views -join ' | '
        BookmarkCount = $bookmarkCount
        Support       = $supportFiles -join ';'
    }
}

Write-Output ('MAIN_FILE_COUNT=' + $rows.Count)
$groups = $rows | Group-Object Hash | Sort-Object { $_.Group[0].Time } -Descending
Write-Output ('UNIQUE_MAIN_HASHES=' + $groups.Count)

foreach ($group in $groups) {
    $row = $group.Group | Sort-Object Time -Descending | Select-Object -First 1
    Write-Output '---GENERATION---'
    [PSCustomObject]@{
        Hash          = $row.Hash
        Copies        = $group.Count
        Length        = $row.Length
        Newest        = $row.Time
        Version       = $row.Version
        Window        = $row.Window
        Status        = $row.Status
        Grid          = $row.Grid
        Ratio         = $row.Ratio
        BookmarkText  = $row.BookmarkText
        BookmarkMulti = $row.BookmarkMulti
        DriveShow     = $row.DriveShow
        DriveMulti    = $row.DriveMulti
        Views         = $row.Views
        BookmarkCount = $row.BookmarkCount
        Representative = $row.Dir
        OtherCopies   = (($group.Group | Sort-Object Dir | Select-Object -ExpandProperty Dir -First 7) -join ' | ')
        Support       = $row.Support
    } | Format-List
}
