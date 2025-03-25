# Dell Command Update Script
# This script installs or updates Dell Command Update via Chocolatey and runs it to update Dell systems.
# Enhanced to detect and remove old versions before installation

# Script parameters
param (
    [switch]$BiosOnly,
    [switch]$Force,
    [switch]$SuspendBitLocker
)

# Dell Command Update Download URL and Checksum Last updated 20-2-24 To Version 5.4.0
$downloadUrl = "https://dl.dell.com/FOLDER11914128M/1/Dell-Command-Update-Windows-Universal-Application_9M35M_WIN_5.4.0_A00.EXE"
$expectedMD5 = "20650f194900e205848a04f0d2d4d947"
$minimumRequiredVersion = [version]"5.0.0"  # Define minimum acceptable version

# Function to check if running as administrator
function Test-Administrator {
    $user = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($user)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to log messages
function Write-Log {
    param($Message)
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    Write-Host $logMessage
    Add-Content -Path "$PSScriptRoot\DellUpdate.log" -Value $logMessage
}

# Function to verify MD5 checksum
function Verify-Checksum {
    param (
        [string]$filePath,
        [string]$expectedHash
    )
    $hash = Get-FileHash -Path $filePath -Algorithm MD5
    return $hash.Hash.ToLower() -eq $expectedHash.ToLower()
}

# Function to get DCU version
function Get-DCUVersion {
    $dellCommandPath = "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe"
    
    if (Test-Path $dellCommandPath) {
        try {
            $versionInfo = (Get-Item $dellCommandPath).VersionInfo
            $version = [version]"$($versionInfo.FileMajorPart).$($versionInfo.FileMinorPart).$($versionInfo.FileBuildPart)"
            return $version
        }
        catch {
            Write-Log "Error getting DCU version: $($_.Exception.Message)"
            return $null
        }
    }
    else {
        # Also check older installation paths
        $legacyDCUPath = "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe"
        if (Test-Path $legacyDCUPath) {
            try {
                $versionInfo = (Get-Item $legacyDCUPath).VersionInfo
                $version = [version]"$($versionInfo.FileMajorPart).$($versionInfo.FileMinorPart).$($versionInfo.FileBuildPart)"
                return $version
            }
            catch {
                Write-Log "Error getting legacy DCU version: $($_.Exception.Message)"
                return $null
            }
        }
    }
    return $null
}

# Function to remove old DCU installations
function Remove-OldDCU {
    Write-Log "Removing old Dell Command Update installation..."
    
    # Try to uninstall using Windows Installer
    try {
        $installers = @(
            "Dell Command | Update",
            "Dell Command | Update for Windows Universal"
        )
        
        foreach ($installer in $installers) {
            $app = Get-WmiObject -Class Win32_Product -Filter "Name LIKE '$installer%'"
            if ($app) {
                Write-Log "Uninstalling $($app.Name) via Windows Installer..."
                $uninstallResult = $app.Uninstall()
                if ($uninstallResult.ReturnValue -eq 0) {
                    Write-Log "Successfully uninstalled $($app.Name)"
                } else {
                    Write-Log "Failed to uninstall $($app.Name) with error code: $($uninstallResult.ReturnValue)"
                }
            }
        }
    }
    catch {
        Write-Log "Error during WMI uninstall: $($_.Exception.Message)"
    }
    
    # Look for uninstall strings in registry
    try {
        $uninstallPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        
        foreach ($path in $uninstallPaths) {
            $keys = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
            foreach ($key in $keys) {
                if ($key.DisplayName -like "*Dell Command*Update*") {
                    Write-Log "Found in registry: $($key.DisplayName)"
                    
                    if ($key.UninstallString) {
                        $uninstallCmd = $key.UninstallString
                        Write-Log "Uninstalling via: $uninstallCmd"
                        
                        # Handle different uninstall methods
                        if ($uninstallCmd -like "*msiexec*") {
                            $uninstallCmd = $uninstallCmd -replace "/I", "/X"
                            $uninstallCmd += " /quiet /norestart"
                            Start-Process "cmd.exe" -ArgumentList "/c $uninstallCmd" -Wait -NoNewWindow
                        } 
                        else {
                            # For non-MSI installers, usually needs /S for silent
                            $uninstallProcess = Start-Process $uninstallCmd -ArgumentList "/S" -Wait -NoNewWindow -PassThru
                            Write-Log "Uninstall process exited with code: $($uninstallProcess.ExitCode)"
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-Log "Error during registry uninstall: $($_.Exception.Message)"
    }
    
    # Clean up directories if they still exist
    $dcuPaths = @(
        "C:\Program Files\Dell\CommandUpdate",
        "C:\Program Files (x86)\Dell\CommandUpdate",
        "C:\Program Files\Dell\UpdateService",
        "C:\Program Files (x86)\Dell\UpdateService"
    )
    
    foreach ($path in $dcuPaths) {
        if (Test-Path $path) {
            Write-Log "Removing directory: $path"
            try {
                Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
            }
            catch {
                Write-Log "Failed to remove directory: $path - $($_.Exception.Message)"
            }
        }
    }
}

# Function to check battery level for BIOS updates
function Get-BatteryStatus {
    $battery = Get-WmiObject Win32_Battery
    if ($battery) {
        return $battery.EstimatedChargeRemaining
    }
    return 100  # Return 100 if no battery (desktop)
}

# Function to check AC power status
function Is-ACPowerConnected {
    $powerStatus = Get-WmiObject -Class BatteryStatus -Namespace root\wmi
    if ($powerStatus) {
        return $powerStatus.PowerOnline
    }
    return $true  # Return true if no battery (desktop)
}

# Function to handle BitLocker
function Manage-BitLocker {
    param (
        [string]$Action  # 'Suspend' or 'Resume'
    )
    
    try {
        # Get BitLocker status for all drives
        $bitlockerVolumes = Get-BitLockerVolume
        $systemDrive = $env:SystemDrive
        $systemVolume = $bitlockerVolumes | Where-Object { $_.MountPoint -eq $systemDrive }
        
        if ($systemVolume) {
            Write-Log "BitLocker status for system drive: $($systemVolume.ProtectionStatus)"
            
            if ($systemVolume.ProtectionStatus -eq "On") {
                if ($Action -eq "Suspend") {
                    Write-Log "Suspending BitLocker..."
                    Suspend-BitLocker -MountPoint $systemDrive -RebootCount 1
                    Write-Log "BitLocker suspended for one reboot"
                    return $true
                }
                elseif ($Action -eq "Resume" -and $systemVolume.ProtectionStatus -eq "Off") {
                    Write-Log "Resuming BitLocker..."
                    Resume-BitLocker -MountPoint $systemDrive
                    Write-Log "BitLocker resumed"
                    return $true
                }
            }
            else {
                Write-Log "BitLocker is not enabled on the system drive"
            }
        }
        else {
            Write-Log "BitLocker is not configured on this system"
        }
    }
    catch {
        Write-Log "Error managing BitLocker: $($_.Exception.Message)"
        return $false
    }
    return $false
}

# Function to check if Chocolatey is installed
function Test-ChocolateyInstalled {
    return (Test-Path -Path "$env:ProgramData\chocolatey\choco.exe")
}

# Function to install Chocolatey if needed
function Install-Chocolatey {
    Write-Log "Installing Chocolatey..."
    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        return $true
    }
    catch {
        Write-Log "Error installing Chocolatey: $($_.Exception.Message)"
        return $false
    }
}

# Check for administrative privileges
if (-not (Test-Administrator)) {
    Write-Host "This script requires administrative privileges. Please run as administrator." -ForegroundColor Red
    exit 1
}

# Path to Dell Command Update executable
$dellCommandPath = "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe"

try {
    # Check if Dell Command Update is installed and get version
    $currentVersion = Get-DCUVersion
    $needsUpdate = $false
    
    if ($currentVersion) {
        Write-Log "Current Dell Command Update version: $currentVersion"
        
        # Check if the version is too old
        if ($currentVersion -lt $minimumRequiredVersion) {
            Write-Log "Current version is below minimum required version ($minimumRequiredVersion). Will remove and reinstall."
            Remove-OldDCU
            $needsUpdate = $true
        }
    } else {
        Write-Log "Dell Command Update not found or version couldn't be determined."
        $needsUpdate = $true
    }

    # Install DCU if needed
    if ($needsUpdate -or -not (Test-Path $dellCommandPath)) {
        Write-Log "Installing Dell Command Update..."
        
        # Check if Chocolatey is installed, install if not
        if (-not (Test-ChocolateyInstalled)) {
            if (-not (Install-Chocolatey)) {
                Write-Log "Failed to install Chocolatey. Falling back to manual installation."
                
                # Create temp directory if it doesn't exist
                $tempDir = "$env:TEMP\DellCommandUpdate"
                if (-not (Test-Path $tempDir)) {
                    New-Item -ItemType Directory -Path $tempDir | Out-Null
                }
                
                # Download the installer
                $installerPath = "$tempDir\DellCommandUpdate.exe"
                Write-Log "Downloading from: $downloadUrl"
                Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath
                
                # Verify checksum
                Write-Log "Verifying download integrity..."
                if (-not (Verify-Checksum -filePath $installerPath -expectedHash $expectedMD5)) {
                    Write-Log "Checksum verification failed! The downloaded file may be corrupted."
                    exit 1
                }
                
                # Install Dell Command Update
                Write-Log "Installing Dell Command Update..."
                Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait -NoNewWindow
            } else {
                # Chocolatey installed successfully, now install DCU
                Write-Log "Installing Dell Command Update via Chocolatey..."
                $chocoProcess = Start-Process -FilePath "choco" -ArgumentList "install dellcommandupdate -y" -Wait -NoNewWindow -PassThru
                if ($chocoProcess.ExitCode -ne 0) {
                    Write-Log "Chocolatey installation failed with exit code: $($chocoProcess.ExitCode)"
                    exit 1
                }
            }
        } else {
            # Chocolatey is already installed, use it to install DCU
            Write-Log "Installing Dell Command Update via Chocolatey..."
            $chocoProcess = Start-Process -FilePath "choco" -ArgumentList "install dellcommandupdate -y" -Wait -NoNewWindow -PassThru
            if ($chocoProcess.ExitCode -ne 0) {
                Write-Log "Chocolatey installation failed with exit code: $($chocoProcess.ExitCode)"
                exit 1
            }
        }
        
        # Verify installation
        Start-Sleep -Seconds 10  # Wait for installation to complete
        if (-not (Test-Path $dellCommandPath)) {
            Write-Log "Installation failed!"
            exit 1
        }
        Write-Log "Dell Command Update installed successfully"
    } else {
        # DCU is installed with acceptable version, but let's update it via Chocolatey if available
        if (Test-ChocolateyInstalled) {
            Write-Log "Updating Dell Command Update via Chocolatey..."
            $chocoProcess = Start-Process -FilePath "choco" -ArgumentList "upgrade dellcommandupdate -y" -Wait -NoNewWindow -PassThru
            if ($chocoProcess.ExitCode -eq 0) {
                Write-Log "Dell Command Update updated successfully"
            } else {
                Write-Log "Note: Chocolatey update returned code: $($chocoProcess.ExitCode) - continuing anyway"
            }
        } else {
            Write-Log "Dell Command Update is already installed. Chocolatey not available for updates."
        }
    }

    # Handle BitLocker if requested
    if ($SuspendBitLocker) {
        if (-not (Manage-BitLocker -Action "Suspend")) {
            if (-not $Force) {
                Write-Log "Failed to suspend BitLocker. Use -Force to continue anyway."
                exit 1
            }
        }
    }

    # BIOS update specific checks
    if ($BiosOnly) {
        Write-Log "BIOS update mode selected..."
        
        # Check battery level and AC power for BIOS updates
        $batteryLevel = Get-BatteryStatus
        $isACConnected = Is-ACPowerConnected

        if (-not $Force) {
            if ($batteryLevel -lt 50) {
                Write-Log "Battery level is below 50% ($batteryLevel%). BIOS update not recommended."
                if (-not $Force) {
                    Write-Log "Use -Force to override this check."
                    exit 1
                }
            }

            if (-not $isACConnected) {
                Write-Log "AC power not connected. BIOS update not recommended."
                if (-not $Force) {
                    Write-Log "Use -Force to override this check."
                    exit 1
                }
            }
        }

        # Scan for BIOS updates only
        Write-Log "Scanning for BIOS updates..."
        $scanProcess = Start-Process -FilePath $dellCommandPath -ArgumentList "/scan -updateType=bios -silent" -Wait -NoNewWindow -PassThru
    } else {
        # Set policies to allow scanning and updates - remove these if they cause errors
        Write-Log "Scanning for updates..."
        $scanProcess = Start-Process -FilePath $dellCommandPath -ArgumentList "/scan -silent" -Wait -NoNewWindow -PassThru
    }
    
    if ($scanProcess.ExitCode -eq 0) {
        Write-Log "Scan completed successfully"
        
        # Download and install updates
        $updateArgs = if ($BiosOnly) { "/applyUpdates -updateType=bios -silent" } else { "/applyUpdates -silent" }
        Write-Log "Downloading and installing updates..."
        $updateProcess = Start-Process -FilePath $dellCommandPath -ArgumentList $updateArgs -Wait -NoNewWindow -PassThru
        
        switch ($updateProcess.ExitCode) {
            0 { 
                Write-Log "Updates installed successfully" 
                $result = "Success - Updates installed"
            }
            1 { 
                Write-Log "Reboot required to complete installation" 
                $result = "Success - Reboot required"
            }
            3 { 
                Write-Log "No updates available" 
                $result = "Success - No updates needed"
            }
            2 { 
                Write-Log "Fatal error occurred" 
                $result = "Error - Fatal error"
            }
            4 { 
                Write-Log "Updates completed with errors" 
                $result = "Warning - Completed with errors"
            }
            5 { 
                Write-Log "Operation cancelled" 
                $result = "Warning - Operation cancelled"
            }
            default { 
                Write-Log "Unknown error occurred (Exit code: $($updateProcess.ExitCode))" 
                $result = "Error - Unknown (Code: $($updateProcess.ExitCode))"
            }
        }
    } else {
        Write-Log "Scan failed with exit code: $($scanProcess.ExitCode)"
        $result = "Error - Scan failed (Code: $($scanProcess.ExitCode))"
    }
} catch {
    Write-Log "Error occurred: $($_.Exception.Message)"
    $result = "Error - Exception: $($_.Exception.Message)"
    exit 1
}

# Check if reboot is pending
$rebootPending = Start-Process -FilePath $dellCommandPath -ArgumentList "/rebootpending" -Wait -NoNewWindow -PassThru
if ($rebootPending.ExitCode -eq 0) {
    Write-Log "System restart is required to complete the update process"
    $result += " (Reboot Pending)"
}
