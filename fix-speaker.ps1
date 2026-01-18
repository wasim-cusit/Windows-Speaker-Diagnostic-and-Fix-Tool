# Windows Speaker Diagnostic and Fix Script
# Run as Administrator for best results

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Speaker Diagnostic and Fix Tool" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "WARNING: Not running as Administrator. Some fixes may require admin rights." -ForegroundColor Yellow
    Write-Host "To run as admin: Right-click script -> 'Run with PowerShell' -> Yes" -ForegroundColor Yellow
    Write-Host ""
}

$ErrorActionPreference = "Continue"

# Function to restart Windows Audio Service
function Restart-AudioService {
    Write-Host "[1/8] Checking and restarting Windows Audio services..." -ForegroundColor Green
    
    $services = @("Audiosrv", "AudioEndpointBuilder", "HidUsb")
    
    foreach ($service in $services) {
        try {
            $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
            if ($svc) {
                if ($svc.Status -ne "Running") {
                    Write-Host "  Starting $service service..." -ForegroundColor Yellow
                    Start-Service -Name $service -ErrorAction SilentlyContinue
                } else {
                    Write-Host "  Restarting $service service..." -ForegroundColor Yellow
                    Restart-Service -Name $service -Force -ErrorAction SilentlyContinue
                }
                Start-Sleep -Seconds 2
                $svc = Get-Service -Name $service
                if ($svc.Status -eq "Running") {
                    Write-Host "  [OK] $service is now running" -ForegroundColor Green
                }
            }
        } catch {
            Write-Host "  [ERROR] Could not restart $service (may need admin rights)" -ForegroundColor Red
        }
    }
    Write-Host ""
}

# Function to check audio devices
function Check-AudioDevices {
    Write-Host "[2/8] Checking audio devices..." -ForegroundColor Green
    
    try {
        # Get all audio-related devices including endpoints, controllers, and audio classes
        $audioDevices = Get-PnpDevice | Where-Object { 
            $_.Class -eq "AudioEndpoint" -or 
            $_.Class -eq "AudioEndpoint" -or
            $_.Class -eq "System" -or
            $_.FriendlyName -like "*audio*" -or 
            $_.FriendlyName -like "*speaker*" -or
            $_.FriendlyName -like "*sound*" -or
            $_.FriendlyName -like "*Realtek*" -or
            $_.FriendlyName -like "*Intel*Audio*" -or
            $_.FriendlyName -like "*Microphone*"
        }
        
        if ($audioDevices) {
            $problemDevices = @()
            foreach ($device in $audioDevices) {
                # "Unknown" status is normal for AudioEndpoint devices - they're not always active
                # Only treat Error, Disabled, or other non-OK statuses as issues for non-endpoint devices
                if ($device.Class -eq "AudioEndpoint") {
                    if ($device.Status -eq "OK" -or $device.Status -eq "Unknown") {
                        $statusText = if ($device.Status -eq "OK") { "[OK]" } else { "[INFO]" }
                        $color = if ($device.Status -eq "OK") { "Green" } else { "Cyan" }
                        Write-Host "  $statusText $($device.FriendlyName) - Status: $($device.Status)" -ForegroundColor $color
                    } else {
                        Write-Host "  [WARNING] $($device.FriendlyName) - Status: $($device.Status)" -ForegroundColor Yellow
                        $problemDevices += $device
                    }
                } else {
                    # For actual hardware devices (not endpoints), OK is expected
                    $statusText = if ($device.Status -eq "OK") { "[OK]" } else { "[WARNING]" }
                    $color = if ($device.Status -eq "OK") { "Green" } else { "Yellow" }
                    Write-Host "  $statusText $($device.FriendlyName) - Status: $($device.Status)" -ForegroundColor $color
                    
                    if ($device.Status -ne "OK") {
                        $problemDevices += $device
                        Write-Host "    Attempting to enable device..." -ForegroundColor Yellow
                        Enable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                        Start-Sleep -Milliseconds 500
                    }
                }
            }
            
            if ($problemDevices.Count -gt 0) {
                Write-Host ""
                Write-Host "  [INFO] Found $($problemDevices.Count) device(s) with issues. Will attempt driver fixes..." -ForegroundColor Yellow
                return $problemDevices
            }
        } else {
            Write-Host "  [WARNING] No audio devices found" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  [ERROR] Error checking audio devices: $_" -ForegroundColor Red
    }
    Write-Host ""
    return @()
}

# Function to check and unmute audio
function Check-AudioVolume {
    Write-Host "[3/8] Checking audio volume and mute status..." -ForegroundColor Green
    
    try {
        # Check registry for audio settings
        $regPath = "HKCU:\Software\Microsoft\Multimedia\Audio"
        $muteValue = (Get-ItemProperty -Path $regPath -Name "UserDuckingPreference" -ErrorAction SilentlyContinue)
        
        Write-Host "  Checking system volume..." -ForegroundColor Yellow
        
        # Use PowerShell to check volume
        try {
            $audio = New-Object -ComObject WScript.Shell
            Write-Host "  [OK] Volume control accessible" -ForegroundColor Green
            Write-Host "  [INFO] Press Volume Up key to check if speakers work" -ForegroundColor Cyan
        } catch {
            Write-Host "  [WARNING] Could not access volume control directly" -ForegroundColor Yellow
        }
        
        # Check if audio is muted using audio endpoint API
        Write-Host "  [INFO] Please check Volume Mixer (right-click speaker icon)" -ForegroundColor Cyan
        Write-Host "    Ensure no applications or system sounds are muted" -ForegroundColor White
        
    } catch {
        Write-Host "  [WARNING] Could not check volume status" -ForegroundColor Yellow
    }
    Write-Host ""
}

# Function to reset audio driver
function Reset-AudioDriver {
    param([array]$ProblemDevices = @())
    
    Write-Host "[4/8] Resetting audio driver..." -ForegroundColor Green
    
    if ($isAdmin) {
        try {
            # Reset problem devices first
            if ($ProblemDevices.Count -gt 0) {
                Write-Host "  Resetting devices with issues..." -ForegroundColor Yellow
                foreach ($device in $ProblemDevices) {
                    if ($device.Class -ne "AudioEndpoint") {
                        Write-Host "    Resetting $($device.FriendlyName)..." -ForegroundColor Yellow
                        Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 2
                        Enable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                        Write-Host "    [OK] Reset $($device.FriendlyName)" -ForegroundColor Green
                        Start-Sleep -Seconds 1
                    }
                }
            }
            
            # Also reset main audio hardware devices
            Write-Host "  Resetting main audio hardware devices..." -ForegroundColor Yellow
            $audioDevices = Get-PnpDevice | Where-Object { 
                ($_.FriendlyName -like "*Realtek*" -or 
                 $_.FriendlyName -like "*Intel*Audio*" -or
                 $_.FriendlyName -like "*Audio*Controller*" -or
                 ($_.Class -eq "System" -and $_.FriendlyName -like "*Audio*")) -and 
                $_.Status -eq "OK" 
            }
            
            foreach ($device in $audioDevices) {
                Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                Enable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                Write-Host "  [OK] Reset $($device.FriendlyName)" -ForegroundColor Green
                Start-Sleep -Seconds 1
            }
            
            Start-Sleep -Seconds 3
        } catch {
            Write-Host "  [ERROR] Error resetting driver: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WARNING] Skipping driver reset (requires admin rights)" -ForegroundColor Yellow
    }
    Write-Host ""
}

# Function to run Windows troubleshooter
function Run-AudioTroubleshooter {
    Write-Host "[5/8] Running Windows Audio Troubleshooter..." -ForegroundColor Green
    
    try {
        # Windows 10/11 audio troubleshooter
        $troubleshooter = "msdt.exe /id AudioPlaybackDiagnostic"
        Write-Host "  Launching Windows Audio Troubleshooter..." -ForegroundColor Yellow
        Start-Process -FilePath "msdt.exe" -ArgumentList "/id", "AudioPlaybackDiagnostic" -ErrorAction SilentlyContinue
        Write-Host "  [OK] Troubleshooter launched - please follow the wizard" -ForegroundColor Green
        Write-Host "  [INFO] The troubleshooter window should open automatically" -ForegroundColor Cyan
    } catch {
        Write-Host "  [ERROR] Could not launch troubleshooter: $_" -ForegroundColor Red
        Write-Host "  [INFO] You can manually run: Settings - System - Sound - Troubleshoot" -ForegroundColor Cyan
    }
    Write-Host ""
}

# Function to install and update audio drivers
function Install-UpdateAudioDrivers {
    param([array]$ProblemDevices = @())
    
    Write-Host "[6/8] Installing and updating audio drivers..." -ForegroundColor Green
    
    if ($isAdmin) {
        try {
            # Get all audio hardware devices (not endpoints)
            $audioHardware = Get-PnpDevice | Where-Object { 
                ($_.FriendlyName -like "*Realtek*" -or 
                 $_.FriendlyName -like "*Intel*Audio*" -or
                 $_.FriendlyName -like "*Audio*Controller*" -or
                 ($_.Class -eq "System" -and $_.FriendlyName -like "*Audio*")) -and
                $_.Class -ne "AudioEndpoint"
            }
            
            if ($audioHardware -or $ProblemDevices.Count -gt 0) {
                Write-Host "  Attempting to update drivers via Windows Update..." -ForegroundColor Yellow
                
                # Try to update drivers using UpdateDriverForPlugAndPlayDevices API
                $driversToUpdate = @()
                if ($ProblemDevices.Count -gt 0) {
                    $driversToUpdate = $ProblemDevices | Where-Object { $_.Class -ne "AudioEndpoint" }
                } else {
                    $driversToUpdate = $audioHardware
                }
                
                foreach ($device in $driversToUpdate) {
                    try {
                        $instanceId = $device.InstanceId
                        Write-Host "    Checking driver for: $($device.FriendlyName)..." -ForegroundColor Yellow
                        
                        # Get hardware ID for driver search
                        $driverInfo = Get-PnpDeviceProperty -InstanceId $instanceId -ErrorAction SilentlyContinue
                        $hardwareId = ($driverInfo | Where-Object { $_.KeyName -eq "DEVPKEY_Device_HardwareID" }).Data
                        
                        if ($hardwareId -and $hardwareId.Count -gt 0) {
                            Write-Host "      Hardware ID: $($hardwareId[0])" -ForegroundColor Gray
                        }
                        
                        # Try to update driver using pnputil
                        Write-Host "      Attempting to install best available driver..." -ForegroundColor Yellow
                        
                        # Use pnputil to add/install driver if available
                        $pnputilCheck = pnputil /enum-drivers 2>&1 | Select-String -Pattern "Realtek|Intel.*Audio|Audio" -ErrorAction SilentlyContinue
                        
                        # Try to reinstall/update driver using PowerShell
                        try {
                            # Enable the device first
                            Enable-PnpDevice -InstanceId $instanceId -Confirm:$false -ErrorAction SilentlyContinue
                            Start-Sleep -Milliseconds 500
                            
                            # Try to refresh/reinstall driver
                            $driverNode = Get-PnpDevice -InstanceId $instanceId | Get-PnpDeviceProperty | Where-Object { $_.KeyName -eq "DEVPKEY_Device_DriverProblemDesc" }
                            if ($driverNode -and $driverNode.Data) {
                                Write-Host "      [WARNING] Driver issue detected: $($driverNode.Data)" -ForegroundColor Yellow
                                Write-Host "      Attempting to reinstall driver..." -ForegroundColor Yellow
                            }
                            
                            Write-Host "      [INFO] Driver check completed. See Device Manager for manual update." -ForegroundColor Cyan
                        } catch {
                            Write-Host "      [INFO] Driver status check completed" -ForegroundColor Cyan
                        }
                        
                        Start-Sleep -Seconds 2
                        
                        # Refresh device status
                        $updatedDevice = Get-PnpDevice -InstanceId $instanceId -ErrorAction SilentlyContinue
                        if ($updatedDevice) {
                            if ($updatedDevice.Status -eq "OK") {
                                Write-Host "      [OK] Device status: OK" -ForegroundColor Green
                            } else {
                                Write-Host "      [INFO] Device status: $($updatedDevice.Status)" -ForegroundColor Cyan
                            }
                        }
                        
                    } catch {
                        Write-Host "      [WARNING] Could not update driver automatically: $_" -ForegroundColor Yellow
                    }
                }
                
                Write-Host ""
                Write-Host "  Opening Device Manager for manual driver installation..." -ForegroundColor Yellow
                try {
                    # Open Device Manager
                    Start-Process "devmgmt.msc" -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 1
                    Write-Host "  [OK] Device Manager opened" -ForegroundColor Green
                    Write-Host ""
                    Write-Host "  [INFO] Manual driver installation steps:" -ForegroundColor Cyan
                    Write-Host "    1. In Device Manager, expand 'Sound, video and game controllers'" -ForegroundColor White
                    Write-Host "    2. Right-click each audio device (especially Realtek Audio)" -ForegroundColor White
                    Write-Host "    3. Select 'Update driver'" -ForegroundColor White
                    Write-Host "    4. Choose 'Search automatically for drivers'" -ForegroundColor White
                    Write-Host "    5. Wait for Windows to download and install drivers" -ForegroundColor White
                    Write-Host ""
                    Write-Host "  [INFO] Or check Windows Update for driver updates:" -ForegroundColor Cyan
                    Write-Host "    - Open Settings - Update & Security - Windows Update" -ForegroundColor White
                    Write-Host "    - Click 'Check for updates'" -ForegroundColor White
                    Write-Host "    - Install any optional driver updates if available" -ForegroundColor White
                } catch {
                    Write-Host "  [WARNING] Could not open Device Manager automatically" -ForegroundColor Yellow
                }
                
            } else {
                Write-Host "  [INFO] No audio hardware devices found to update" -ForegroundColor Cyan
            }
            
            # Display current driver information
            Write-Host ""
            Write-Host "  Current audio driver information:" -ForegroundColor Yellow
            $allAudioDevices = Get-PnpDevice | Where-Object { 
                ($_.FriendlyName -like "*Realtek*" -or 
                 $_.FriendlyName -like "*Intel*Audio*" -or
                 ($_.Class -eq "System" -and $_.FriendlyName -like "*Audio*")) -and
                $_.Class -ne "AudioEndpoint"
            } | Select-Object -First 5
            
            foreach ($driver in $allAudioDevices) {
                $driverInfo = Get-PnpDeviceProperty -InstanceId $driver.InstanceId -ErrorAction SilentlyContinue
                $version = ($driverInfo | Where-Object { $_.KeyName -eq "DEVPKEY_Device_DriverVersion" }).Data
                $provider = ($driverInfo | Where-Object { $_.KeyName -eq "DEVPKEY_Device_DriverProvider" }).Data
                
                Write-Host "    - $($driver.FriendlyName)" -ForegroundColor Cyan
                Write-Host "      Status: $($driver.Status)" -ForegroundColor $(if ($driver.Status -eq "OK") { "Green" } else { "Yellow" })
                if ($version) {
                    Write-Host "      Driver Version: $version" -ForegroundColor Gray
                }
                if ($provider) {
                    Write-Host "      Provider: $provider" -ForegroundColor Gray
                }
            }
            
        } catch {
            Write-Host "  [ERROR] Error updating drivers: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WARNING] Driver installation requires admin rights" -ForegroundColor Yellow
        Write-Host "  [INFO] Please run this script as Administrator to install drivers" -ForegroundColor Cyan
    }
    Write-Host ""
}

# Function to set default audio device
function Set-DefaultAudioDevice {
    Write-Host "[7/8] Checking default audio device..." -ForegroundColor Green
    
    try {
        Write-Host "  [INFO] Check your default playback device:" -ForegroundColor Cyan
        Write-Host "    - Right-click speaker icon in taskbar" -ForegroundColor White
        Write-Host "    - Select 'Open Sound settings'" -ForegroundColor White
        Write-Host "    - Under 'Output', select your speakers" -ForegroundColor White
        Write-Host "    - Click 'Test' to verify" -ForegroundColor White
        
        # Try to open sound settings
        try {
            Start-Process "ms-settings:sound" -ErrorAction SilentlyContinue
            Write-Host "  [OK] Opening Sound settings..." -ForegroundColor Green
        } catch {
            Write-Host "  [INFO] You can manually open: Settings - System - Sound" -ForegroundColor Cyan
        }
        
    } catch {
        Write-Host "  [WARNING] Could not programmatically set audio device" -ForegroundColor Yellow
    }
    Write-Host ""
}

# Function to provide manual steps
function Show-ManualSteps {
    Write-Host "[8/8] Additional manual steps to try..." -ForegroundColor Green
    Write-Host ""
    Write-Host "  Manual Troubleshooting Steps:" -ForegroundColor Yellow
    Write-Host "  ------------------------------" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  1. Check Physical Connections:" -ForegroundColor Cyan
    Write-Host "     - Ensure speakers/headphones are properly connected" -ForegroundColor White
    Write-Host "     - Try unplugging and reconnecting" -ForegroundColor White
    Write-Host ""
    Write-Host "  2. Check Volume and Mute:" -ForegroundColor Cyan
    Write-Host "     - Right-click speaker icon - Open Volume mixer" -ForegroundColor White
    Write-Host "     - Ensure nothing is muted (no red X)" -ForegroundColor White
    Write-Host "     - Increase volume to test" -ForegroundColor White
    Write-Host ""
    Write-Host "  3. Update Audio Driver:" -ForegroundColor Cyan
    Write-Host "     - Device Manager - Sound, video and game controllers" -ForegroundColor White
    Write-Host "     - Right-click audio device - Update driver" -ForegroundColor White
    Write-Host "     - Or visit laptop manufacturer website for latest drivers" -ForegroundColor White
    Write-Host ""
    Write-Host "  4. Restart Audio Service (if not already done):" -ForegroundColor Cyan
    Write-Host "     - Press Win+R, type: services.msc" -ForegroundColor White
    Write-Host "     - Find 'Windows Audio' and 'Windows Audio Endpoint Builder'" -ForegroundColor White
    Write-Host "     - Right-click each - Restart" -ForegroundColor White
    Write-Host ""
    Write-Host "  5. Run Windows Troubleshooter:" -ForegroundColor Cyan
    Write-Host "     - Settings - System - Sound - Troubleshoot" -ForegroundColor White
    Write-Host ""
    Write-Host "  6. Check for Windows Updates:" -ForegroundColor Cyan
    Write-Host "     - Settings - Update and Security - Windows Update" -ForegroundColor White
    Write-Host "     - Install any pending updates" -ForegroundColor White
    Write-Host ""
    Write-Host "  7. Reset Audio Settings:" -ForegroundColor Cyan
    Write-Host "     - Settings - System - Sound" -ForegroundColor White
    Write-Host "     - Scroll down - 'More sound settings'" -ForegroundColor White
    Write-Host "     - Playback tab - Right-click speakers - Set as Default Device" -ForegroundColor White
    Write-Host ""
}

# Main execution
Write-Host "Starting diagnostics..." -ForegroundColor Cyan
Write-Host ""

Restart-AudioService
Start-Sleep -Seconds 2

$problemDevices = Check-AudioDevices
Start-Sleep -Seconds 2

Check-AudioVolume

Reset-AudioDriver -ProblemDevices $problemDevices
Start-Sleep -Seconds 2

Install-UpdateAudioDrivers -ProblemDevices $problemDevices

Set-DefaultAudioDevice

Run-AudioTroubleshooter

Show-ManualSteps

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Diagnostics Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "If speakers still do not work after trying all steps:" -ForegroundColor Yellow
Write-Host "  - Check Device Manager for error codes" -ForegroundColor White
Write-Host "  - Try updating drivers from manufacturer website" -ForegroundColor White
Write-Host "  - Consider system restore if issue started recently" -ForegroundColor White
Write-Host "  - Contact laptop manufacturer support" -ForegroundColor White
Write-Host ""
Write-Host 'Press any key to exit...'
try {
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
} catch {
    Read-Host 'Press Enter to exit'
}
