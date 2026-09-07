param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$script:Passed = 0
$script:Failed = [Collections.Generic.List[string]]::new()

function Assert-True([bool]$Condition, [string]$Message) {
    if ($Condition) { $script:Passed++ }
    else { $script:Failed.Add($Message) }
}

function Read-Source([string]$RelativePath) {
    $root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    return [IO.File]::ReadAllText((Join-Path $root $RelativePath))
}

$optionH = Read-Source 'src\fxfile\option.h'
$optionCpp = Read-Source 'src\fxfile\option.cpp'
$explorerH = Read-Source 'src\fxfile\explorer_ctrl.h'
$explorerCpp = Read-Source 'src\fxfile\explorer_ctrl.cpp'
$paneCpp = Read-Source 'src\fxfile\explorer_pane.cpp'
$cfgCpp = Read-Source 'src\fxfile\cfg\cfg_appearance_thumbnail_dlg.cpp'
$thumbH = Read-Source 'src\fxfile\thumbnail.h'
$thumbCpp = Read-Source 'src\fxfile\thumbnail.cpp'
$listH = Read-Source 'src\fxfile\thumb_img_list.h'
$listCpp = Read-Source 'src\fxfile\thumb_img_list.cpp'
$layoutH = Read-Source 'src\fxfile\folder_layout.h'
$rc = Read-Source 'src\fxfile\fxfile.rc'
$resourceH = Read-Source 'src\fxfile\resource.h'

[xml]$korean = Read-Source 'src\fxfile\Languages\Korean.xml'
Assert-True ($null -ne $korean.DocumentElement) 'Korean language XML must parse.'

# Cache path option: persistence, bounded load, UI and runtime propagation.
Assert-True ($optionH.Contains('mThumbnailCachePath[XPR_MAX_PATH + 1]')) 'Config must own a bounded thumbnail cache path.'
Assert-True ($optionCpp.Contains('config.thumbnail.cache_path')) 'Thumbnail cache path must be persisted.'
Assert-True ([regex]::IsMatch($optionCpp, 'config\.thumbnail\.cache_path[^\r\n]*mThumbnailCachePath[^\r\n]*XPR_MAX_PATH\s*\+\s*1')) 'Thumbnail cache key must declare its destination capacity.'
Assert-True ($optionCpp.Contains('_tcsncpy_s(sValue, sOptionKey->mCapacity, sLoadedValue, _TRUNCATE)')) 'Bounded string option load must prevent corrupt-config overflow.'
Assert-True ($explorerH.Contains('mThumbnailCachePath[XPR_MAX_PATH + 1]')) 'Explorer option must carry the cache path.'
Assert-True ($paneCpp.Contains('mThumbnailCachePath')) 'ExplorerPane must copy the cache path.'
Assert-True (($explorerCpp.Split('setCacheDir(').Count - 1) -ge 2) 'Pending and applied explorer options must configure the cache path.'
Assert-True ($cfgCpp.Contains('OnCachePathBrowse')) 'Thumbnail settings must expose a folder browser.'
Assert-True ($cfgCpp.Contains('GetDriveType(sRoot) != DRIVE_FIXED')) 'Custom cache path must require a fixed local drive.'
Assert-True ($cfgCpp.Contains('FILE_ATTRIBUTE_REPARSE_POINT')) 'Custom cache path must reject reparse/cloud redirection.'
Assert-True ($cfgCpp.Contains('1024ULL * 1024ULL * 1024ULL')) 'Custom cache path must require at least 1 GiB free.'
Assert-True ($cfgCpp.Contains('FXFILE_CACHE_PATH_PROBE')) 'Custom cache path must pass a write/flush/delete probe.'
Assert-True ($rc.Contains('IDC_CFG_THUMBNAIL_CACHE_PATH_BROWSE')) 'Thumbnail settings resource must include the browse button.'
Assert-True ($resourceH.Contains('#define IDC_CFG_THUMBNAIL_CACHE_PATH_BROWSE 1890')) 'Cache browse control ID must be stable.'

# Cache corruption and transactional persistence contracts.
Assert-True ($thumbH.Contains('setCacheDir')) 'Thumbnail cache directory API must be declared.'
Assert-True ($listH.Contains('deleteCacheFiles')) 'Explicit cache initialization must delete the persisted pair.'
Assert-True ($listCpp.Contains('kThumbnailDataMagic') -and $listCpp.Contains('kThumbnailIndexMagic')) 'Data and index files must have independent magic values.'
Assert-True ($listCpp.Contains('kThumbnailCacheVersion')) 'Cache format must be versioned.'
Assert-True ($listCpp.Contains('sDataGeneration == sGeneration')) 'Data/index generations must match on load.'
Assert-True ($listCpp.Contains('sThumbElement.mImageIndex != (xpr_sint_t)sLoadedThumbs.size()')) 'Index order must match image-list order.'
Assert-True ($listCpp.Contains('sUsedThumbImageIds')) 'Thumbnail IDs must be unique.'
Assert-True ($listCpp.Contains('kInvalidThumbImageId')) 'Reserved invalid thumbnail ID must be rejected.'
Assert-True ($listCpp.Contains('FlushFileBuffers')) 'Cache writes must be flushed before commit.'
Assert-True ($listCpp.Contains('MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH')) 'Cache commit must use durable replacement.'
Assert-True ($listCpp.Contains('return deleteCacheFiles();')) 'Saving an empty cache must not resurrect an old pair.'
Assert-True ($listCpp.Contains('kMaxThumbnailRecords   = 4096')) 'Cache record count must have a conservative hard cap.'
Assert-True ($listCpp.Contains('#if defined(_WIN64)')) 'x86 and x64 cache size limits must be architecture-aware.'
Assert-True ($listCpp.Contains('sEstimatedPayload > kMaxThumbnailDataSize')) 'Oversized image generations must be rejected before temporary cache serialization.'
Assert-True ($listCpp.Contains('kCacheSafetyReserve = 512ULL * 1024ULL * 1024ULL')) 'Cache save must preserve a fixed free-space reserve.'
Assert-True ($listCpp.Contains('sAvailable.QuadPart < sRequiredFreeSpace')) 'Cache save must re-check actual target free space at serialization time.'
Assert-True ($thumbCpp.Contains('ImageList_GetIconSize')) 'Loaded cache dimensions must match configured thumbnail size.'
Assert-True ($listCpp.Contains('ImageList_SetIconSize')) 'Live size changes must preserve the shared HIMAGELIST handle.'
Assert-True ($thumbCpp.Contains('A queued or already decoded thumbnail must not repopulate the cache')) 'Explicit cache initialization must stop and drain pending thumbnail work.'

# Per-column full/ellipsis policy and legacy layout integrity.
$columnMembers = @('Name','Size','Type','Date','Attr','Ext')
foreach ($name in $columnMembers) {
    Assert-True ($optionCpp.Contains("config.file_list.column_ellipsis_$($name.ToLowerInvariant())")) "Missing persisted $name ellipsis policy."
    Assert-True ($explorerH.Contains("mColumnEllipsis$name")) "Missing Explorer $name ellipsis policy."
    Assert-True ($paneCpp.Contains("mColumnEllipsis$name")) "Missing ExplorerPane $name ellipsis transfer."
    Assert-True ($resourceH.Contains("IDC_CFG_FOLDERLAYOUT_ELLIPSIS_$($name.ToUpperInvariant())")) "Missing $name checkbox resource ID."
}
Assert-True ($explorerCpp.Contains('void ExplorerCtrl::adjustAutomaticColumnWidths')) 'AutoFull column adjustment must be centralized.'
Assert-True ($explorerCpp.Contains('const xpr_sint_t kMaximumSamples = 64')) 'Large folders must not trigger full synchronous auto-width scans.'
Assert-True ($explorerCpp.Contains('restoreSavedColumnWidths(aNewOption)')) 'Switching to ellipsis must restore the retained manual width.'
Assert-True ($explorerCpp.Contains('adjustAutomaticColumnWidths();')) 'AutoFull must apply to the current pane without restart.'
Assert-True ($explorerCpp.Contains('Windows Shell extension columns keep their saved width')) 'Dynamic shell columns must avoid synchronous full scans.'
Assert-True ($explorerCpp.Contains('return ((GetStyle() & LVS_TYPEMASK) == LVS_REPORT)')) 'Details and Content report views must both support auto width.'
Assert-True ($explorerCpp.Contains('scheduleAutomaticColumnReflow();')) 'Pane/window resizing must schedule responsive column reflow.'
Assert-True ($explorerCpp.Contains('mInAutomaticColumnLayout')) 'Programmatic widths must not be persisted as user header drags.'
Assert-True ($layoutH.Contains('return !(aColumnId1 == aColumnId2);')) 'ColumnId inequality must be the exact negation of equality.'

if ($script:Failed.Count -gt 0) {
    foreach ($failure in $script:Failed) { Write-Error "FAIL: $failure" -ErrorAction Continue }
    throw "$($script:Failed.Count) Task056 feature contract checks failed; $($script:Passed) passed."
}

Write-Host "PASS: $($script:Passed) Task056 cache/column contracts; no user file or cache was created, moved, or deleted."
exit 0
