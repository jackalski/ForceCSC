# Cisco Secure Client Monitor

A PowerShell solution to automatically maintain VPN connection using Cisco Secure Client.

## Features

- Automatically monitors VPN connection status at configurable intervals
- Automatically connects to your specified VPN network when disconnected
- Handles "Connect capability is unavailable" error by resetting Cisco Secure Client
- Runs at system startup to ensure connection is established
- Provides desktop shortcut for manual connection
- Detailed logging for troubleshooting
- Customizable configuration during installation

## Installation

1. Download the package and extract it to a folder
2. Right-click on `Installer.ps1` and select "Run with PowerShell"
3. Follow the prompts to configure:
   - VPN connection name (e.g., "EMEA Copenhagen")
   - DNS suffix for connection verification (e.g., "example.com")
   - Host to ping for additional verification
   - Check interval in minutes
4. The installer will:
   - Create a customized monitoring script with your settings
   - Create scheduled tasks to run the script at your specified interval and at logon
   - Create a startup shortcut to ensure the script runs at login
   - Create a desktop shortcut for manual connection

## Manual Connection

If you need to manually connect to the VPN:
- Double-click the "Connect to [Your VPN]" shortcut on your desktop

## Logs

Logs are stored in `%USERPROFILE%\Scripts\CiscoSecureClientMonitor.log`

## Uninstallation

1. Right-click on `Uninstaller.ps1` and select "Run with PowerShell"
2. The uninstaller will remove:
   - All scheduled tasks
   - Desktop and startup shortcuts
   - Script files

## Troubleshooting

If you encounter issues:

1. Check the log file at `%USERPROFILE%\Scripts\CiscoSecureClientMonitor.log`
2. Ensure Cisco Secure Client is properly installed
3. Try running the desktop shortcut manually to see any error messages
4. Verify that the scheduled task exists in Task Scheduler

## Requirements

- Windows 10 or later
- PowerShell 5.1 or later
- Cisco Secure Client installed

## License

This project is licensed under the MIT License - see the LICENSE file for details. 
