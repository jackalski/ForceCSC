# Cisco Secure Client Monitor - Installer Script
# Author: Piotr Szmitkowski @ DSV
# License: MIT

# Ensure we're running with admin rights for task scheduling
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "This script requires administrative privileges to create scheduled tasks." -ForegroundColor Yellow
    Write-Host "The script will attempt to continue, but some features may not work correctly." -ForegroundColor Yellow
    Write-Host "Press any key to continue or CTRL+C to abort..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Create the Scripts directory in user profile if it doesn't exist
$scriptsFolder = Join-Path $env:USERPROFILE "Scripts"
if (-not (Test-Path $scriptsFolder)) {
    New-Item -Path $scriptsFolder -ItemType Directory | Out-Null
}

# Display welcome message
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "      Cisco Secure Client Monitor - Installer       " -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host "This tool will help you set up automatic monitoring" -ForegroundColor Cyan
Write-Host "and connection for your Cisco Secure Client VPN." -ForegroundColor Cyan
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host ""

# Check if Cisco Secure Client is installed
$vpnCliPath = "C:\Program Files (x86)\Cisco\Cisco Secure Client\vpncli.exe"
if (-not (Test-Path $vpnCliPath)) {
    Write-Host "WARNING: Cisco Secure Client does not appear to be installed at the expected location." -ForegroundColor Red
    Write-Host "The script will continue, but may not work correctly." -ForegroundColor Red
    Write-Host "Press any key to continue or CTRL+C to abort..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Prompt for configuration details
Write-Host "Please provide the following details for your VPN connection:" -ForegroundColor Cyan
Write-Host ""

# Ask whether to use UI or CLI client
Write-Host "VPN Client Mode" -ForegroundColor Green
Write-Host "  1. UI Client (recommended)" -ForegroundColor Green
Write-Host "  2. CLI Client" -ForegroundColor Green
$clientModeSelection = Read-Host "Select VPN client mode (default: 1)"

if ($clientModeSelection -eq "2") {
    $useUiClient = $false
    Write-Host "Using CLI client for VPN connections" -ForegroundColor Yellow
} else {
    $useUiClient = $true
    Write-Host "Using UI client for VPN connections (default)" -ForegroundColor Yellow
}

# VPN Connection Name - try to get available connections
$availableConnections = @()
try {
    $output = & "$vpnCliPath" "hosts" 2>&1
    $connectionLines = $output | Where-Object { $_ -match "^\s+\d+\.\s+" }
    foreach ($line in $connectionLines) {
        if ($line -match "^\s+\d+\.\s+(.+?)\s*$") {
            $availableConnections += $matches[1]
        }
    }
} catch {
    Write-Host "Could not retrieve available VPN connections: $_" -ForegroundColor Yellow
}

# Display available connections if found
if ($availableConnections.Count -gt 0) {
    Write-Host "Available VPN connections:" -ForegroundColor Green
    for ($i = 0; $i -lt $availableConnections.Count; $i++) {
        Write-Host "  $($i+1). $($availableConnections[$i])" -ForegroundColor Green
    }
    
    $selection = Read-Host "Enter the number of your VPN connection or type a custom name"
    if ([int]::TryParse($selection, [ref]$null) -and [int]$selection -ge 1 -and [int]$selection -le $availableConnections.Count) {
        $vpnConnectionName = $availableConnections[[int]$selection-1]
    } else {
        $vpnConnectionName = $selection
    }
} else {
    $vpnConnectionName = Read-Host "Enter your VPN connection name (e.g., 'EMEA Copenhagen')"
}

if ([string]::IsNullOrWhiteSpace($vpnConnectionName)) {
    $vpnConnectionName = "EMEA Copenhagen"  # Default value
    Write-Host "Using default connection name: $vpnConnectionName" -ForegroundColor Yellow
}

# DNS Suffix for verification - try to detect current DNS suffixes
$currentDnsSuffixes = @()
try {
    $networkInterfaces = Get-NetAdapter | Where-Object { $_.Status -eq "Up" } | Get-DnsClient -ErrorAction SilentlyContinue
    foreach ($interface in $networkInterfaces) {
        if (-not [string]::IsNullOrWhiteSpace($interface.ConnectionSpecificSuffix)) {
            $currentDnsSuffixes += $interface.ConnectionSpecificSuffix
        }
    }
} catch {
    Write-Host "Could not retrieve current DNS suffixes: $_" -ForegroundColor Yellow
}

# Display current DNS suffixes if found
if ($currentDnsSuffixes.Count -gt 0) {
    Write-Host "Current DNS suffixes on your system:" -ForegroundColor Green
    for ($i = 0; $i -lt $currentDnsSuffixes.Count; $i++) {
        Write-Host "  $($i+1). $($currentDnsSuffixes[$i])" -ForegroundColor Green
    }
}

$dnsSuffix = Read-Host "Enter the DNS suffix to verify connection (e.g., 'example.com')"
if ([string]::IsNullOrWhiteSpace($dnsSuffix)) {
    $dnsSuffix = "example.com"  # Default value
    Write-Host "Using default DNS suffix: $dnsSuffix" -ForegroundColor Yellow
}

# Ping host for verification
$pingHost = Read-Host "Enter a host to ping for connection verification (e.g., 'someserver.example.com')"
if ([string]::IsNullOrWhiteSpace($pingHost)) {
    $pingHost = "someserver.$dnsSuffix"  # Default based on DNS suffix
    Write-Host "Using default ping host: $pingHost" -ForegroundColor Yellow
}

# Check interval
$checkInterval = Read-Host "Enter the interval in minutes to check connection (default: 5)"
if (-not [int]::TryParse($checkInterval, [ref]$null) -or [int]$checkInterval -lt 1) {
    $checkInterval = 5  # Default value
    Write-Host "Using default check interval: $checkInterval minutes" -ForegroundColor Yellow
}

# Get the template script
$templatePath = Join-Path (Get-Location) "CiscoSecureClientMonitor.ps1"
if (-not (Test-Path $templatePath)) {
    $templatePath = Join-Path (Get-Location) "Scripts\CiscoSecureClientMonitor.ps1"
    if (-not (Test-Path $templatePath)) {
        Write-Host "ERROR: Could not find the CiscoSecureClientMonitor.ps1 template." -ForegroundColor Red
        exit 1
    }
}

# Read the template content
$scriptContent = Get-Content -Path $templatePath -Raw

# Replace placeholders with actual values
$scriptContent = $scriptContent.Replace("##VPN_CONNECTION_NAME##", $vpnConnectionName)
$scriptContent = $scriptContent.Replace("##DNS_SUFFIX##", $dnsSuffix)
$scriptContent = $scriptContent.Replace("##PING_HOST##", $pingHost)
# Correctly handle boolean values
if ($useUiClient) {
    $scriptContent = $scriptContent.Replace("##USE_UI_CLIENT##", '$true')
} else {
    $scriptContent = $scriptContent.Replace("##USE_UI_CLIENT##", '$false')
}

# Save the customized script
$scriptPath = Join-Path $scriptsFolder "CiscoSecureClientMonitor.ps1"
Set-Content -Path $scriptPath -Value $scriptContent

# Create a VBScript wrapper to run the PowerShell script invisibly
Write-Host "Creating VBScript wrapper for invisible execution..."
$vbsWrapperPath = Join-Path $scriptsFolder "RunVpnMonitor.vbs"
$vbsContent = @"
' VBScript wrapper to run PowerShell script without any visible window
Option Explicit
Dim shell, powershellPath, scriptPath, command

' Get PowerShell path and script path
powershellPath = "powershell.exe"

' Set default script path
scriptPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%USERPROFILE%") & "\Scripts\CiscoSecureClientMonitor.ps1"

' Use argument if provided
If WScript.Arguments.Count > 0 Then
    scriptPath = WScript.Arguments(0)
End If

' Build command to run PowerShell invisibly
command = powershellPath & " -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """"

' Create shell object and run command invisibly (0 = hide window completely)
Set shell = CreateObject("WScript.Shell")
shell.Run command, 0, False

Set shell = Nothing
"@
Set-Content -Path $vbsWrapperPath -Value $vbsContent

# Create direct scheduled tasks and shortcuts - more reliable than registration script
Write-Host ""
Write-Host "Creating scheduled tasks and shortcuts..."

# Current user for task credentials
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

# Task names
$taskName = "CiscoSecureClientMonitor"
$logonTaskName = "$taskName-Logon"

# Update the task creation to use the VBS wrapper
Write-Host "Creating main scheduled task that runs every $checkInterval minutes..."
$escapedVbsPath = $vbsWrapperPath
if ($vbsWrapperPath.Contains(" ")) {
    $escapedVbsPath = "`"$vbsWrapperPath`""
}

# Build command with proper escaping
$mainTaskCmd = "schtasks.exe /Create /TN `"$taskName`" /TR `"wscript.exe $escapedVbsPath`" /SC MINUTE /MO $checkInterval /RU `"$currentUser`" /F"

# Execute the command
try {
    cmd.exe /c "$mainTaskCmd 2>&1"
    
    # Create logon task
    Write-Host "Creating logon task..."
    $logonTaskCmd = "schtasks.exe /Create /TN `"$logonTaskName`" /TR `"wscript.exe $escapedVbsPath`" /SC ONLOGON /RU `"$currentUser`" /F"
    
    $logonTaskResult = cmd.exe /c "$logonTaskCmd 2>&1"
    if ($logonTaskResult -match "ERROR") {
        Write-Host "Could not create logon task: $logonTaskResult" -ForegroundColor Yellow
    } else {
        Write-Host "Logon task created successfully!" -ForegroundColor Green
    }
} catch {
    Write-Host "Exception occurred creating tasks: $_" -ForegroundColor Red
}

# Create desktop shortcut
Write-Host "Creating desktop shortcut..."
try {
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $shortcutPath = Join-Path $desktopPath "Connect to VPN.lnk"
    
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($shortcutPath)
    $Shortcut.TargetPath = "wscript.exe"
    $Shortcut.Arguments = "`"$vbsWrapperPath`""
    $Shortcut.IconLocation = "C:\Program Files (x86)\Cisco\Cisco Secure Client\UI\csc_ui.exe,0"
    $Shortcut.Description = "Connect to VPN"
    $Shortcut.Save()
    
    Write-Host "Desktop shortcut created at: $shortcutPath" -ForegroundColor Green
} catch {
    Write-Host "Error creating desktop shortcut: $_" -ForegroundColor Red
}

# Create startup shortcut
Write-Host "Creating startup shortcut..."
try {
    $startupFolder = [Environment]::GetFolderPath("Startup")
    $startupPath = Join-Path $startupFolder "VPN Connection.lnk"
    
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($startupPath)
    $Shortcut.TargetPath = "wscript.exe" 
    $Shortcut.Arguments = "`"$vbsWrapperPath`""
    $Shortcut.IconLocation = "C:\Program Files (x86)\Cisco\Cisco Secure Client\UI\csc_ui.exe,0"
    $Shortcut.Description = "Connect to VPN"
    $Shortcut.Save()
    
    Write-Host "Startup shortcut created at: $startupPath" -ForegroundColor Green
} catch {
    Write-Host "Error creating startup shortcut: $_" -ForegroundColor Red
}

Write-Host ""
Write-Host "====================================================" -ForegroundColor Green
Write-Host "            Installation completed!                 " -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Green
Write-Host "Scripts installed to: $scriptsFolder" -ForegroundColor Green
Write-Host ""
Write-Host "Configuration Summary:" -ForegroundColor Green
Write-Host "- VPN Connection: $vpnConnectionName" -ForegroundColor Green
Write-Host "- DNS Suffix: $dnsSuffix" -ForegroundColor Green
Write-Host "- Ping Host: $pingHost" -ForegroundColor Green
Write-Host "- Check Interval: $checkInterval minutes" -ForegroundColor Green
Write-Host ""
Write-Host "You can manually connect by clicking the desktop shortcut" -ForegroundColor Green
Write-Host "or run the task from Task Scheduler." -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Green

# Offer to run the script immediately
$runNow = Read-Host "Would you like to run the connection script now? (y/n)"
if ($runNow -eq "y" -or $runNow -eq "Y") {
    Write-Host "Running connection script..."
    Start-Process "wscript.exe" -ArgumentList "`"$vbsWrapperPath`"" -WindowStyle Hidden
    Write-Host "Script launched invisibly. Check the log file for results."
}
