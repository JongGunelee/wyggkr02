
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
    $manifestContent = Get-Content $manifestPath -Raw | ConvertFrom-Json
    $entries = [System.Collections.ArrayList]$manifestContent.entries
    
    foreach ($skill in $newSkills) {
        if ($entries -notcontains $skill) {
            $null = $entries.Add($skill)
            Write-Host "Added: $skill"
        }
    }
    
    # Sort and convert back to array
    $manifestContent.entries = $entries | Sort-Object
    $manifestContent.updatedAt = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    
    $manifestContent | ConvertTo-Json -Depth 10 | Set-Content $manifestPath -Encoding UTF8
    Write-Host "Sync Complete! All skills are now indexed." -ForegroundColor Green
}
