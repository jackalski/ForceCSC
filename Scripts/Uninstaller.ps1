# Cisco Secure Client Monitor - Uninstaller Script
# Author: Piotr Szmitkowski @ DSV
# License: MIT

# Ensure we're running with admin rights for task removal
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "This script requires administrative privileges to remove scheduled tasks." -ForegroundColor Yellow
    Write-Host "The script will attempt to continue, but some features may not work correctly." -ForegroundColor Yellow
    Write-Host "Press any key to continue or CTRL+C to abort..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "     Cisco Secure Client Monitor - Uninstaller      " -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host ""

# Confirm uninstallation
$confirm = Read-Host "Are you sure you want to uninstall Cisco Secure Client Monitor? (y/n)"
if ($confirm -ne "y" -and $confirm -ne "Y") {
    Write-Host "Uninstallation cancelled."
    exit 0
}

Write-Host "Starting uninstallation of Cisco Secure Client Monitor..."

$taskName = "CiscoSecureClientMonitor"
$logonTaskName = "$taskName-Logon"

# Remove scheduled tasks
Write-Host "Removing scheduled tasks..."
try {
    schtasks.exe /Delete /TN $taskName /F
    Write-Host "Successfully removed task: $taskName" -ForegroundColor Green
} catch {
    Write-Host "Could not remove task $taskName. It may not exist or you may not have permission." -ForegroundColor Yellow
}

try {
    schtasks.exe /Delete /TN $logonTaskName /F
    Write-Host "Successfully removed task: $logonTaskName" -ForegroundColor Green
} catch {
    Write-Host "Could not remove task $logonTaskName. It may not exist or you may not have permission." -ForegroundColor Yellow
}

# Remove startup shortcuts
$startupFolder = [Environment]::GetFolderPath("Startup")
Write-Host "Checking for VPN connection shortcuts in startup folder..."
$startupShortcutsRemoved = 0
Get-ChildItem -Path $startupFolder -Filter "VPN Connection*.lnk" | ForEach-Object {
    try {
        Remove-Item -Path $_.FullName -Force
        Write-Host "Removed startup shortcut: $($_.Name)" -ForegroundColor Green
        $startupShortcutsRemoved++
    } catch {
        $errorMsg = $_.Exception.Message
        Write-Host ("Could not remove startup shortcut {0}: {1}" -f $_.Name, $errorMsg) -ForegroundColor Yellow
    }
}
if ($startupShortcutsRemoved -eq 0) {
    Write-Host "No startup shortcuts found." -ForegroundColor Yellow
}

# Remove desktop shortcuts
$desktopFolder = [Environment]::GetFolderPath("Desktop")
Write-Host "Checking for VPN connection shortcuts on desktop..."
$desktopShortcutsRemoved = 0
Get-ChildItem -Path $desktopFolder -Filter "Connect to*.lnk" | ForEach-Object {
    try {
        Remove-Item -Path $_.FullName -Force
        Write-Host "Removed desktop shortcut: $($_.Name)" -ForegroundColor Green
        $desktopShortcutsRemoved++
    } catch {
        $errorMsg = $_.Exception.Message
        Write-Host ("Could not remove desktop shortcut {0}: {1}" -f $_.Name, $errorMsg) -ForegroundColor Yellow
    }
}
if ($desktopShortcutsRemoved -eq 0) {
    Write-Host "No desktop shortcuts found." -ForegroundColor Yellow
}

# Define script paths
$scriptsFolder = Join-Path $env:USERPROFILE "Scripts"
$monitorScript = Join-Path $scriptsFolder "CiscoSecureClientMonitor.ps1"
$registrationScript = Join-Path $scriptsFolder "RegisterCiscoMonitorTask.ps1"
$logFile = Join-Path $scriptsFolder "CiscoSecureClientMonitor.log"

# Remove script files
Write-Host "Removing script files..."
$filesRemoved = 0
$filesToRemove = @($monitorScript, $registrationScript, $logFile)
foreach ($file in $filesToRemove) {
    if (Test-Path $file) {
        try {
            Remove-Item -Path $file -Force
            Write-Host ("Removed file: {0}" -f $file) -ForegroundColor Green
            $filesRemoved++
        } catch {
            $errorMsg = $_.Exception.Message
            Write-Host ("Could not remove file {0}: {1}" -f $file, $errorMsg) -ForegroundColor Yellow
        }
    }
}
if ($filesRemoved -eq 0) {
    Write-Host "No script files found." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "====================================================" -ForegroundColor Green
Write-Host "           Uninstallation completed!                " -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Green
Write-Host "The following items were removed:" -ForegroundColor Green
Write-Host "- Scheduled tasks: $taskName and $logonTaskName" -ForegroundColor Green
Write-Host "- Startup shortcuts: $startupShortcutsRemoved found" -ForegroundColor Green
Write-Host "- Desktop shortcuts: $desktopShortcutsRemoved found" -ForegroundColor Green
Write-Host "- Script files: $filesRemoved found" -ForegroundColor Green
Write-Host ""
Write-Host "Note: The Scripts folder was not removed in case it contains other files." -ForegroundColor Yellow
Write-Host "====================================================" -ForegroundColor Green 
