# Windows Speaker Diagnostic and Fix Tool

This tool helps diagnose and fix common laptop speaker issues on Windows.

## Files Included

- **fix-speaker.ps1** - Main PowerShell diagnostic script
- **fix-speaker.bat** - Easy launcher (double-click to run)
- **fix-speaker-admin.bat** - Launcher with admin rights (recommended for full fixes)

## How to Use

### Option 1: Simple Run (Recommended for first try)
1. Double-click `fix-speaker.bat`
2. Follow the prompts

### Option 2: Full Administrator Mode (Recommended for driver fixes)
1. Right-click `fix-speaker-admin.bat`
2. Select "Run as administrator"
3. Click "Yes" when prompted
4. Follow the prompts

### Option 3: Direct PowerShell
1. Right-click `fix-speaker.ps1`
2. Select "Run with PowerShell"
3. Select "Yes" if prompted for admin rights

## What This Tool Does

The script performs the following diagnostic and repair steps:

1. **Restarts Windows Audio Services** - Restarts essential audio services
2. **Checks Audio Devices** - Verifies all audio devices are detected and enabled
3. **Checks Volume and Mute Status** - Ensures nothing is muted
4. **Resets Audio Drivers** - Disables and re-enables audio drivers
5. **Runs Windows Troubleshooter** - Launches built-in Windows audio troubleshooter
6. **Checks Driver Status** - Shows current audio driver information
7. **Checks Default Audio Device** - Provides guidance on setting default device
8. **Provides Manual Steps** - Lists additional troubleshooting steps

## Common Issues Fixed

- Audio services not running
- Audio drivers disabled or not working
- Volume muted
- Wrong default audio device
- Driver conflicts

## Requirements

- Windows 10/11
- PowerShell (included with Windows)
- Administrator rights (for full functionality)

## Notes

- Some fixes require Administrator rights
- The script will launch Windows Audio Troubleshooter automatically
- All steps are safe and reversible
- No personal data is collected or transmitted

## If Issues Persist

If speakers still don't work after running this tool:

1. Check Device Manager for error codes
2. Download latest audio drivers from your laptop manufacturer's website
3. Check Windows Update for pending driver updates
4. Consider system restore if issue started recently
5. Contact laptop manufacturer support

## Troubleshooting the Script

If the script won't run:
- Right-click the file > Properties > Unblock > OK
- Try running as Administrator
- Check PowerShell execution policy: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`
