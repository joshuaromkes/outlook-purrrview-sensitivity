<#
.SYNOPSIS
    Installs or uninstalls the Outlook Purview Sensitivity VSTO add-in.

.DESCRIPTION
    Install: copies published VSTO files to %LocalAppData% and registers
    the add-in under HKCU with LoadBehavior=3 (load at startup).

    Uninstall: removes registry key and installed files.

.PARAMETER SourcePath
    Path to the published output folder containing the .vsto file.
    Required for -Install.

.PARAMETER Uninstall
    Remove the add-in from the registry and delete installed files.

.EXAMPLE
    .\Install-AddIn.ps1 -SourcePath .\bin\Release\app.publish

.EXAMPLE
    .\Install-AddIn.ps1 -Uninstall
#>

param(
    [string]$SourcePath,
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"
$AddInName = "Outlook-Purview-Sensitivity"
$RegKeyPath = "HKCU:\Software\Microsoft\Office\Outlook\AddIns\$AddInName"
$InstallDir = "$env:LocalAppData\Outlook-Purview-Sensitivity"

function Uninstall-AddIn {
    Write-Host "Uninstalling $AddInName..." -ForegroundColor Yellow

    # Close Outlook if running
    $outlook = Get-Process outlook -ErrorAction SilentlyContinue
    if ($outlook) {
        Write-Host "  Closing Outlook..." -ForegroundColor Yellow
        $outlook | Stop-Process -Force
        Start-Sleep -Seconds 2
    }

    # Remove registry key
    if (Test-Path $RegKeyPath) {
        Remove-Item -Path $RegKeyPath -Recurse -Force
        Write-Host "  Registry key removed." -ForegroundColor Green
    }
    else {
        Write-Host "  Registry key not found." -ForegroundColor Gray
    }

    # Remove installed files
    if (Test-Path $InstallDir) {
        Remove-Item -Path $InstallDir -Recurse -Force
        Write-Host "  Files removed from $InstallDir" -ForegroundColor Green
    }
    else {
        Write-Host "  No installed files found." -ForegroundColor Gray
    }

    Write-Host "Uninstall complete." -ForegroundColor Green
}

function Install-AddIn {
    if (-not $SourcePath) {
        Write-Host "ERROR: -SourcePath is required." -ForegroundColor Red
        Write-Host "Usage: .\Install-AddIn.ps1 -SourcePath path-to-publish-folder" -ForegroundColor Gray
        exit 1
    }

    if (-not (Test-Path $SourcePath)) {
        Write-Host "ERROR: SourcePath '$SourcePath' does not exist." -ForegroundColor Red
        Write-Host "Publish the project first: Build -> Publish -> FolderProfile -> Publish" -ForegroundColor Gray
        exit 1
    }

    Write-Host "Installing $AddInName from $SourcePath" -ForegroundColor Cyan

    # Uninstall any existing version first
    Uninstall-AddIn

    # Copy files
    Write-Host "  Copying files to $InstallDir ..." -ForegroundColor Gray
    New-Item -Path $InstallDir -ItemType Directory -Force | Out-Null
    Copy-Item -Path "$SourcePath\*" -Destination $InstallDir -Recurse -Force

    # Find the .vsto file
    $vstoFile = Get-ChildItem -Path $InstallDir -Filter *.vsto -Recurse | Select-Object -First 1
    if (-not $vstoFile) {
        Write-Host "ERROR: No .vsto file found in $InstallDir." -ForegroundColor Red
        Write-Host "Was the project published? Run: Build -> Publish -> FolderProfile -> Publish" -ForegroundColor Gray
        exit 1
    }

    # Build manifest URL with vstolocal suffix
    $manifestUrl = "file:///" + $vstoFile.FullName.Replace('\', '/') + "|vstolocal"
    Write-Host "  Manifest: $manifestUrl" -ForegroundColor Gray

    # Register in HKCU
    Write-Host "  Registering in HKCU..." -ForegroundColor Gray
    New-Item -Path $RegKeyPath -Force | Out-Null
    Set-ItemProperty -Path $RegKeyPath -Name "Description" -Value "Displays Microsoft Purview sensitivity labels in Outlook" -Type String
    Set-ItemProperty -Path $RegKeyPath -Name "FriendlyName" -Value "Outlook Purview Sensitivity" -Type String
    Set-ItemProperty -Path $RegKeyPath -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty -Path $RegKeyPath -Name "Manifest" -Value $manifestUrl -Type String

    # Verify
    $verify = Get-ItemProperty -Path $RegKeyPath -ErrorAction SilentlyContinue
    if ($verify -and $verify.LoadBehavior -eq 3) {
        Write-Host "Installation successful." -ForegroundColor Green
        Write-Host "Start Outlook to see the Sensitivity Label column." -ForegroundColor Cyan
    }
    else {
        Write-Host "WARNING: Registry verification failed. Check: $RegKeyPath" -ForegroundColor Red
    }
}

# --- Main ---
if ($Uninstall) {
    Uninstall-AddIn
}
else {
    Install-AddIn
}
