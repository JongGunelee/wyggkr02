param (
    [Parameter(Mandatory=$true)]
    [ValidateSet("x64", "x32")]
    [string]$Arch
)

$message = @"
AutoBuild-And-Cleanup.ps1 is a retired historical workflow and is intentionally blocked.
It copied the source to a D:\ root sandbox, built only one architecture, and manually copied
an unverified bin tree, so it cannot satisfy the current three-package rollback/manifest or
the Task 057 drive/TEMP hard gate.

Requested legacy architecture: $Arch
Run this current workflow from fxfile_working instead:
  .\preflight_build_environment.bat
  .\build_deploy_all.bat
"@

Write-Error $message
exit 1
