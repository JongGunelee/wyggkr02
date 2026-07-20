
$skillsDir = 'C:\Users\PC\.gemini\antigravity\skills'
$manifestPath = Join-Path $skillsDir '.antigravity-install-manifest.json'

Write-Host "Rebuilding skill manifest..."
$items = Get-ChildItem -Path $skillsDir | Where-Object { $_.PSIsContainer -or $_.Name -match '\.(md|json|csv|xlsx|pptx|docx|pdf)$' } | 
         Where-Object { $_.Name -ne '.antigravity-install-manifest.json' } | 
         Select-Object -ExpandProperty Name

if ($items.Count -gt 0) {
    $manifest = [ordered]@{
        schemaVersion = 1
        updatedAt = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        entries = $items
    }
    
    $json = $manifest | ConvertTo-Json -Depth 10
    $json | Set-Content $manifestPath -Encoding UTF8
    Write-Host "Manifest successfully rebuilt with $($items.Count) entries." -ForegroundColor Green
} else {
    Write-Error "No skill items found in $skillsDir"
}
