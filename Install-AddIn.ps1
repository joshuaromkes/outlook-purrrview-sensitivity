<#
.SYNOPSIS
    Installs or uninstalls the Outlook Purview Sensitivity VSTO add-in
    using the ClickOnce setup.exe produced by the publish step.

.DESCRIPTION
    This is a thin wrapper around the ClickOnce bootstrapper (setup.exe).
    All install/uninstall logic and VSTO runtime management is handled
    by setup.exe -- this script just finds it and passes the right flags.

.PARAMETER SourcePath
    Path to the published output folder containing setup.exe and .vsto files.
    Defaults to ".\bin\Release\app.publish".

.PARAMETER Uninstall
    Uninstall the add-in via setup.exe. Alternatively, uninstall from
    Windows Settings -> Apps & Features -> Outlook Purview Sensitivity.

.EXAMPLE
    .\Install-AddIn.ps1

.EXAMPLE
    .\Install-AddIn.ps1 -SourcePath .\bin\Release\app.publish

.EXAMPLE
    .\Install-AddIn.ps1 -Uninstall
#>

param(
    [string]$SourcePath = ".\bin\Release\app.publish",
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"
$SetupExe = Join-Path $SourcePath "setup.exe"

if (-not (Test-Path $SetupExe)) {
    Write-Host "ERROR: setup.exe not found at $SetupExe" -ForegroundColor Red
    Write-Host "Publish the project first: Build -> Publish -> FolderProfile -> Publish" -ForegroundColor Gray
    exit 1
}

if ($Uninstall) {
    Write-Host "Uninstalling via setup.exe..." -ForegroundColor Yellow
    Start-Process -FilePath $SetupExe -ArgumentList "/uninstall" -Wait
    Write-Host "Done. You can also uninstall from Settings -> Apps & Features." -ForegroundColor Green
}
else {
    Write-Host "Installing via setup.exe..." -ForegroundColor Cyan
    Start-Process -FilePath $SetupExe -ArgumentList "/install" -Wait
    Write-Host "Done. Start Outlook to see the Sensitivity Label column." -ForegroundColor Green
}
