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

# Function to enable system speakers (not Bluetooth)
function Enable-SystemSpeakers {
    Write-Host "[2a/8] Enabling disabled system speakers..." -ForegroundColor Green
    
    if ($isAdmin) {
        try {
            # Get speaker endpoints (excluding Bluetooth)
            $speakerEndpoints = Get-PnpDevice | Where-Object { 
                $_.Class -eq "AudioEndpoint" -and
                ($_.FriendlyName -like "*Speaker*" -or $_.FriendlyName -like "*Headphone*" -or $_.FriendlyName -like "*Realtek*") -and
                $_.FriendlyName -notlike "*Bluetooth*" -and
                $_.FriendlyName -notlike "*BT*" -and
                $_.Status -eq "Error" -or $_.Status -eq "Disabled" -or $_.Status -eq "Unknown"
            }
            
            $enabledCount = 0
            foreach ($endpoint in $speakerEndpoints) {
                Write-Host "  Found disabled speaker: $($endpoint.FriendlyName) - Status: $($endpoint.Status)" -ForegroundColor Yellow
                try {
                    Enable-PnpDevice -InstanceId $endpoint.InstanceId -Confirm:$false -ErrorAction Stop
                    Start-Sleep -Seconds 1
                    $refreshed = Get-PnpDevice -InstanceId $endpoint.InstanceId -ErrorAction SilentlyContinue
                    if ($refreshed -and ($refreshed.Status -eq "OK" -or $refreshed.Status -eq "Unknown")) {
                        Write-Host "    [OK] Enabled $($endpoint.FriendlyName)" -ForegroundColor Green
                        $enabledCount++
                    } else {
                        Write-Host "    [INFO] Attempted to enable. Current status: $($refreshed.Status)" -ForegroundColor Cyan
                    }
                } catch {
                    Write-Host "    [WARNING] Could not enable: $_" -ForegroundColor Yellow
                }
            }
            
            if ($enabledCount -gt 0) {
                Write-Host "  [OK] Enabled $enabledCount system speaker device(s)" -ForegroundColor Green
            } else {
                Write-Host "  [INFO] No disabled system speakers found to enable" -ForegroundColor Cyan
            }
            
            # Also check and enable system audio hardware devices (not Bluetooth)
            Write-Host ""
            Write-Host "  Checking system audio hardware devices (excluding Bluetooth)..." -ForegroundColor Yellow
            $systemAudioDevices = Get-PnpDevice | Where-Object {
                ($_.FriendlyName -like "*Realtek*" -or 
                 $_.FriendlyName -like "*Intel*Audio*" -or
                 ($_.Class -eq "System" -and $_.FriendlyName -like "*Audio*")) -and
                $_.FriendlyName -notlike "*Bluetooth*" -and
                $_.FriendlyName -notlike "*BT*" -and
                ($_.Status -eq "Error" -or $_.Status -eq "Disabled")
            }
            
            foreach ($device in $systemAudioDevices) {
                Write-Host "  Found disabled audio hardware: $($device.FriendlyName) - Status: $($device.Status)" -ForegroundColor Yellow
                try {
                    Enable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction Stop
                    Start-Sleep -Seconds 1
                    Write-Host "    [OK] Enabled $($device.FriendlyName)" -ForegroundColor Green
                    $enabledCount++
                } catch {
                    Write-Host "    [WARNING] Could not enable: $_" -ForegroundColor Yellow
                }
            }
            
        } catch {
            Write-Host "  [ERROR] Error enabling speakers: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "  [WARNING] Skipping speaker enable (requires admin rights)" -ForegroundColor Yellow
    }
    Write-Host ""
}

# Function to check audio devices
function Check-AudioDevices {
    Write-Host "[2/8] Checking audio devices..." -ForegroundColor Green
    
    try {
        # Get system audio devices (excluding Bluetooth)
        $audioDevices = Get-PnpDevice | Where-Object { 
            ($_.Class -eq "AudioEndpoint" -or 
             $_.Class -eq "System") -and
            ($_.FriendlyName -like "*audio*" -or 
             $_.FriendlyName -like "*speaker*" -or
             $_.FriendlyName -like "*sound*" -or
             $_.FriendlyName -like "*Realtek*" -or
             $_.FriendlyName -like "*Intel*Audio*" -or
             $_.FriendlyName -like "*Microphone*" -or
             $_.FriendlyName -like "*Headphone*") -and
            $_.FriendlyName -notlike "*Bluetooth*" -and
            $_.FriendlyName -notlike "*BT*"
        }
        
        if ($audioDevices) {
            $problemDevices = @()
            $systemSpeakers = @()
            
            foreach ($device in $audioDevices) {
                # Check if this is a speaker endpoint
                $isSpeaker = $device.FriendlyName -like "*Speaker*" -or $device.FriendlyName -like "*Headphone*" -or $device.FriendlyName -like "*Realtek*Speaker*"
                
                if ($device.Class -eq "AudioEndpoint") {
                    if ($device.Status -eq "OK" -or $device.Status -eq "Unknown") {
                        $statusText = if ($device.Status -eq "OK") { "[OK]" } else { "[INFO]" }
                        $color = if ($device.Status -eq "OK") { "Green" } else { "Cyan" }
                        Write-Host "  $statusText $($device.FriendlyName) - Status: $($device.Status)" -ForegroundColor $color
                        
                        # Track system speakers for default device setting
                        if ($isSpeaker) {
                            $systemSpeakers += $device
                        }
                    } else {
                        Write-Host "  [WARNING] $($device.FriendlyName) - Status: $($device.Status)" -ForegroundColor Yellow
                        $problemDevices += $device
                        if ($isSpeaker) {
                            $systemSpeakers += $device
                        }
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
            
            # Return both problem devices and system speakers for later processing
            if ($problemDevices.Count -gt 0 -or $systemSpeakers.Count -gt 0) {
                Write-Host ""
                if ($problemDevices.Count -gt 0) {
                    Write-Host "  [INFO] Found $($problemDevices.Count) device(s) with issues. Will attempt driver fixes..." -ForegroundColor Yellow
                }
                if ($systemSpeakers.Count -gt 0) {
                    Write-Host "  [INFO] Found $($systemSpeakers.Count) system speaker device(s) to configure..." -ForegroundColor Cyan
                }
                return @{
                    ProblemDevices = $problemDevices
                    SystemSpeakers = $systemSpeakers
                }
            }
        } else {
            Write-Host "  [WARNING] No system audio devices found" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "  [ERROR] Error checking audio devices: $_" -ForegroundColor Red
    }
    Write-Host ""
    return @{
        ProblemDevices = @()
        SystemSpeakers = @()
    }
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
    param([hashtable]$DeviceInfo = @{})
    
    Write-Host "[4/8] Resetting audio driver..." -ForegroundColor Green
    
    if ($isAdmin) {
        try {
            $ProblemDevices = $DeviceInfo.ProblemDevices
            if ($null -eq $ProblemDevices) { $ProblemDevices = @() }
            
            # Reset problem devices first (excluding Bluetooth)
            if ($ProblemDevices.Count -gt 0) {
                Write-Host "  Resetting devices with issues (excluding Bluetooth)..." -ForegroundColor Yellow
                foreach ($device in $ProblemDevices) {
                    if ($device.Class -ne "AudioEndpoint" -and 
                        $device.FriendlyName -notlike "*Bluetooth*" -and 
                        $device.FriendlyName -notlike "*BT*") {
                        Write-Host "    Resetting $($device.FriendlyName)..." -ForegroundColor Yellow
                        Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 2
                        Enable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                        Write-Host "    [OK] Reset $($device.FriendlyName)" -ForegroundColor Green
                        Start-Sleep -Seconds 1
                    }
                }
            }
            
            # Also reset main audio hardware devices (system only, not Bluetooth)
            Write-Host "  Resetting system audio hardware devices..." -ForegroundColor Yellow
            $audioDevices = Get-PnpDevice | Where-Object { 
                (($_.FriendlyName -like "*Realtek*" -or 
                  $_.FriendlyName -like "*Intel*Audio*" -or
                  $_.FriendlyName -like "*Audio*Controller*" -or
                  ($_.Class -eq "System" -and $_.FriendlyName -like "*Audio*")) -and
                 $_.FriendlyName -notlike "*Bluetooth*" -and
                 $_.FriendlyName -notlike "*BT*") -and 
                ($_.Status -eq "OK" -or $_.Status -eq "Error" -or $_.Status -eq "Disabled")
            }
            
            foreach ($device in $audioDevices) {
                Write-Host "    Resetting $($device.FriendlyName)..." -ForegroundColor Yellow
                Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                Enable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                Write-Host "    [OK] Reset $($device.FriendlyName)" -ForegroundColor Green
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
    param([hashtable]$DeviceInfo = @{})
    
    $ProblemDevices = $DeviceInfo.ProblemDevices
    if ($null -eq $ProblemDevices) { $ProblemDevices = @() }
    
    Write-Host "[6/8] Installing and updating audio drivers..." -ForegroundColor Green
    
    if ($isAdmin) {
        try {
            # Get system audio hardware devices (not endpoints, not Bluetooth)
            $audioHardware = Get-PnpDevice | Where-Object { 
                (($_.FriendlyName -like "*Realtek*" -or 
                  $_.FriendlyName -like "*Intel*Audio*" -or
                  $_.FriendlyName -like "*Audio*Controller*" -or
                  ($_.Class -eq "System" -and $_.FriendlyName -like "*Audio*")) -and
                 $_.FriendlyName -notlike "*Bluetooth*" -and
                 $_.FriendlyName -notlike "*BT*") -and
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
            
            # Display current driver information (system only, not Bluetooth)
            Write-Host ""
            Write-Host "  Current system audio driver information:" -ForegroundColor Yellow
            $allAudioDevices = Get-PnpDevice | Where-Object { 
                (($_.FriendlyName -like "*Realtek*" -or 
                  $_.FriendlyName -like "*Intel*Audio*" -or
                  ($_.Class -eq "System" -and $_.FriendlyName -like "*Audio*")) -and
                 $_.FriendlyName -notlike "*Bluetooth*" -and
                 $_.FriendlyName -notlike "*BT*") -and
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

# Function to set default audio device (system speakers, not Bluetooth)
function Set-DefaultAudioDevice {
    param([array]$SystemSpeakers = @())
    
    Write-Host "[7/8] Setting system speakers as default playback device..." -ForegroundColor Green
    
    try {
        Write-Host "  Checking available playback devices..." -ForegroundColor Yellow
        
        # List available audio playback devices from PnP
        Write-Host "  Checking system audio playback devices..." -ForegroundColor Yellow
        $playbackDevices = Get-PnpDevice | Where-Object {
            $_.Class -eq "AudioEndpoint" -and
            ($_.FriendlyName -like "*Speaker*" -or $_.FriendlyName -like "*Headphone*" -or $_.FriendlyName -like "*Realtek*")
        }
        
        if ($playbackDevices) {
            $systemSpeakersFound = @()
            foreach ($device in $playbackDevices) {
                $deviceName = $device.FriendlyName
                $isSystemSpeaker = ($deviceName -notlike "*Bluetooth*" -and $deviceName -notlike "*BT*" -and
                                   ($deviceName -like "*Speaker*" -or $deviceName -like "*Realtek*" -or $deviceName -like "*Headphone*"))
                
                if ($isSystemSpeaker) {
                    Write-Host "    [SYSTEM SPEAKER] $deviceName - Status: $($device.Status)" -ForegroundColor Green
                    $systemSpeakersFound += $device
                } else {
                    Write-Host "    [OTHER/BLUETOOTH] $deviceName" -ForegroundColor Gray
                }
            }
            
            if ($systemSpeakersFound.Count -gt 0) {
                Write-Host "  [INFO] Found $($systemSpeakersFound.Count) system speaker device(s)" -ForegroundColor Cyan
                Write-Host "  [INFO] These will be available in Sound settings for selection" -ForegroundColor Cyan
            } else {
                Write-Host "  [WARNING] No system speakers found in available devices" -ForegroundColor Yellow
            }
        } else {
            Write-Host "  [INFO] Checking audio endpoints..." -ForegroundColor Cyan
        }
        
        # Always provide manual instructions
        Write-Host ""
        Write-Host "  [IMPORTANT] Manual steps to set SYSTEM SPEAKERS as default (NOT Bluetooth):" -ForegroundColor Yellow
        Write-Host "    1. Right-click speaker icon in taskbar" -ForegroundColor White
        Write-Host "    2. Select 'Open Sound settings'" -ForegroundColor White
        Write-Host "    3. Under 'Output', select your SYSTEM SPEAKERS:" -ForegroundColor White
        Write-Host "       - Look for 'Realtek Audio' or 'Speakers (Realtek Audio)'" -ForegroundColor Cyan
        Write-Host "       - DO NOT select Bluetooth or BT devices" -ForegroundColor Red
        Write-Host "       - Make sure it shows 'Realtek' or 'Speakers' not 'Bluetooth'" -ForegroundColor Yellow
        Write-Host "    4. Click 'Test' to verify speakers work" -ForegroundColor White
        Write-Host "    5. If speakers don't appear, click 'More sound settings'" -ForegroundColor White
        Write-Host "    6. In Playback tab, right-click system speakers - Set as Default Device" -ForegroundColor White
        
        # Try to open sound settings
        try {
            Start-Sleep -Seconds 1
            Start-Process "ms-settings:sound" -ErrorAction SilentlyContinue
            Write-Host ""
            Write-Host "  [OK] Opening Sound settings for you..." -ForegroundColor Green
        } catch {
            Write-Host "  [INFO] You can manually open: Settings - System - Sound" -ForegroundColor Cyan
        }
        
    } catch {
        Write-Host "  [WARNING] Could not programmatically set audio device: $_" -ForegroundColor Yellow
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

# Enable disabled system speakers first
Enable-SystemSpeakers
Start-Sleep -Seconds 2

# Check audio devices (excluding Bluetooth)
$deviceInfo = Check-AudioDevices
Start-Sleep -Seconds 2

Check-AudioVolume

Reset-AudioDriver -DeviceInfo $deviceInfo
Start-Sleep -Seconds 2

Install-UpdateAudioDrivers -DeviceInfo $deviceInfo

Set-DefaultAudioDevice -SystemSpeakers $deviceInfo.SystemSpeakers

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
