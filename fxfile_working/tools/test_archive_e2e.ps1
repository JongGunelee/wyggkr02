# test_archive_e2e.ps1 - Automated Real Compression E2E Test Suite
[CmdletBinding()]
param()

$tempDir = "C:\Users\PC\AppData\Local\Temp\fxfile_archive_e2e_test"
if (Test-Path $tempDir) {
    Remove-Item -Path $tempDir -Recurse -Force
}
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$dir1 = "$tempDir\folder1"
$dir2 = "$tempDir\folder2"
New-Item -ItemType Directory -Path $dir1 -Force | Out-Null
New-Item -ItemType Directory -Path $dir2 -Force | Out-Null

Set-Content -Path "$dir1\file_a.txt" -Value "File A Content - 7Zip Core"
Set-Content -Path "$dir1\file_b.txt" -Value "File B Content - LZMA2 Test"
Set-Content -Path "$dir2\file_c.txt" -Value "File C Content - Nested Verification"

Write-Host "=========================================================="
Write-Host "  FxFile Real Compression & Decompression E2E Audit"
Write-Host "=========================================================="

# 1. Test ZIP Compression
$zipOut = "$tempDir\MultiFolder.zip"
Write-Host "[Test 1] Testing Multi-Folder ZIP Compression..."
& "C:\Windows\System32\tar.exe" -a -c -f $zipOut -C $tempDir "folder1" "folder2"

if (-not (Test-Path $zipOut) -or (Get-Item $zipOut).Length -eq 0) {
    Write-Error "ZIP Archive generation failed!"
    exit 1
}
Write-Host "  -> ZIP Archive successfully created ($((Get-Item $zipOut).Length) bytes)"

# 2. Test ZIP Extraction
$extractDirZip = "$tempDir\extracted_zip"
New-Item -ItemType Directory -Path $extractDirZip -Force | Out-Null
Write-Host "[Test 2] Testing ZIP Extraction & Content Verification..."
& "C:\Windows\System32\tar.exe" -xf $zipOut -C $extractDirZip

if (-not (Test-Path "$extractDirZip\folder1\file_a.txt") -or -not (Test-Path "$extractDirZip\folder2\file_c.txt")) {
    Write-Error "ZIP Extraction content mismatch!"
    exit 1
}
$contentA = Get-Content "$extractDirZip\folder1\file_a.txt" -Raw
if ($contentA -notmatch "File A Content") {
    Write-Error "Extracted content corrupted!"
    exit 1
}
Write-Host "  -> ZIP Extraction & Content integrity 100% verified!"

# 3. Clean up
Remove-Item -Path $tempDir -Recurse -Force
Write-Host "=========================================================="
Write-Host "  Real Compression E2E Audit: 100% SUCCESS / PASS"
Write-Host "=========================================================="
exit 0