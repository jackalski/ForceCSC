# Script to register the Cisco Secure Client monitor as a scheduled task

# Create the Scripts directory in user profile if it doesn't exist
$scriptsFolder = Join-Path $env:USERPROFILE "Scripts"
if (-not (Test-Path $scriptsFolder)) {
    New-Item -Path $scriptsFolder -ItemType Directory | Out-Null
}

# Get current user for task credentials
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

# Create the scheduled task
$taskName = "CiscoSecureClientMonitor"
$scriptPath = Join-Path $scriptsFolder "CiscoSecureClientMonitor.ps1"

# Use schtasks.exe directly - this is the most compatible approach
Write-Host "Registering scheduled task using schtasks.exe..."

# Create a task that runs every 5 minutes (or as configured)
$scriptPathEscaped = $scriptPath.Replace('\', '\\').Replace('"', '\"')
$taskCmd = "schtasks.exe /Create /TN `"$taskName`" /TR `"PowerShell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPathEscaped`"`" /SC MINUTE /MO 5 /RU `"$currentUser`" /F"

try {
    # Execute the command
    Write-Host "Running command: $taskCmd"
    cmd.exe /c $taskCmd
    
    # If successful, try to add a logon trigger (more reliable than startup for regular users)
    try {
        $logonCmd = "schtasks.exe /Create /TN `"$taskName-Logon`" /TR `"PowerShell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPathEscaped`"`" /SC ONLOGON /RU `"$currentUser`" /F"
        Write-Host "Running command: $logonCmd"
        cmd.exe /c $logonCmd
        Write-Host "Logon task created successfully."
    } catch {
        $errorMsg = $_.Exception.Message
        Write-Host ("Could not create logon task: {0}" -f $errorMsg) -ForegroundColor Yellow
    }
    
    Write-Host "Task '$taskName' has been registered. It will run every 5 minutes and at logon."
} catch {
    $errorMsg = $_.Exception.Message
    Write-Host ("Error registering task: {0}" -f $errorMsg) -ForegroundColor Yellow
    
    # Try with system account if user account fails
    Write-Host "Attempting to register task with SYSTEM account..."
    $taskCmdSystem = "schtasks.exe /Create /TN `"$taskName`" /TR `"PowerShell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPathEscaped`"`" /SC MINUTE /MO 5 /RU SYSTEM /F"
    cmd.exe /c $taskCmdSystem
    
    Write-Host "Task '$taskName' has been registered with SYSTEM account. It will run every 5 minutes."
}

# Create a shortcut in the Startup folder to run at login
try {
    $startupFolder = [Environment]::GetFolderPath("Startup")
    $shortcutPath = Join-Path $startupFolder "VPN Connection.lnk"
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($shortcutPath)
    $Shortcut.TargetPath = "PowerShell.exe"
    $Shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""
    $Shortcut.IconLocation = "C:\Program Files (x86)\Cisco\Cisco Secure Client\UI\csc_ui.exe,0"
    $Shortcut.Description = "Connect to VPN"
    $Shortcut.Save()
    
    Write-Host "Created startup shortcut for automatic connection at login"
} catch {
    $errorMsg = $_.Exception.Message
    Write-Host ("Could not create startup shortcut: {0}" -f $errorMsg) -ForegroundColor Yellow
}

Write-Host "You can manually run the task from Task Scheduler or by running: Start-ScheduledTask -TaskName '$taskName'"

# Create a desktop shortcut to run the script manually
$desktopPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "Connect to VPN.lnk"
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($desktopPath)
$Shortcut.TargetPath = "PowerShell.exe"
$Shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
$Shortcut.IconLocation = "C:\Program Files (x86)\Cisco\Cisco Secure Client\UI\csc_ui.exe,0"
$Shortcut.Description = "Connect to VPN"
$Shortcut.Save()

Write-Host "Created desktop shortcut: 'Connect to VPN'"
