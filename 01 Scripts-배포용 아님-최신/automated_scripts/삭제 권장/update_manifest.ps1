
$manifestPath = 'C:\Users\PC\.gemini\antigravity\skills\.antigravity-install-manifest.json'
$newSkills = @(
    'api-and-interface-design',
    'browser-testing-with-devtools',
    'ci-cd-and-automation',
    'code-review-and-quality',
    'code-simplification',
    'context-engineering',
    'debugging-and-error-recovery',
    'deprecation-and-migration',
    'documentation-and-adrs',
    'frontend-ui-engineering',
    'git-workflow-and-versioning',
    'idea-refine',
    'incremental-implementation',
    'performance-optimization',
    'planning-and-task-breakdown',
    'security-and-hardening',
    'shipping-and-launch',
    'source-driven-development',
    'spec-driven-development',
    'test-driven-development',
    'using-agent-skills',
    'references',
    'agents'
)

if (Test-Path $manifestPath) {
    Write-Host "Updating manifest at $manifestPath..."
    $json = Get-Content $manifestPath -Raw | ConvertFrom-Json
    $existing = [System.Collections.Generic.HashSet[string]]::new($json.entries)
    
    foreach ($skill in $newSkills) {
        if ($existing.Add($skill)) {
            Write-Host "Adding $skill..."
        }
    }
    
    $json.entries = [string[]]($existing | Sort-Object)
    $json.updatedAt = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    
    $json | ConvertTo-Json -Depth 10 | Set-Content $manifestPath
    Write-Host "Manifest updated successfully. Please restart the IDE or wait a moment for the search to update." -ForegroundColor Green
} else {
    Write-Error "Manifest file not found at $manifestPath"
}
