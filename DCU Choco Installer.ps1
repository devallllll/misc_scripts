<#
.SYNOPSIS
    Manages Dell Command Update installation and updates via Chocolatey
.DESCRIPTION
    Checks current DCU status, ensures Chocolatey is available, removes incompatible software,
    and installs/updates Dell Command Update via Chocolatey following industry best practices.
.PARAMETER ForceReinstall
    Forces removal and reinstallation even if current version is acceptable
.PARAMETER ChocolateySource
    Chocolatey source to use for installation (defaults to community repository)
.EXAMPLE
    .\Install-DellCommandUpdate.ps1
    .\Install-DellCommandUpdate.ps1 -ForceReinstall
    .\Install-DellCommandUpdate.ps1 -ChocolateySource "https://internal.repo.url/api/v2"
.NOTES
    Author: Modified for best practices
    Requires: PowerShell 5.1+ and Administrator privileges
    Based on industry standards from Dell/Chocolatey documentation
#>

[CmdletBinding()]
param(
    [switch]$ForceReinstall,
    [string]$ChocolateySource = "https://community.chocolatey.org/api/v2/"
)

#Requires -RunAsAdministrator

# Script configuration
$Script:LogPath = "$env:SystemRoot\Logs\DellCommandUpdate"
$Script:DCUMinVersion = [version]"5.0.0"
$Script:ValidChocolateyExitCodes = @(0, 1605, 1614, 1641, 3010)

# Initialize logging
function Initialize-Logging {
    if (-not (Test-Path $Script:LogPath)) {
        New-Item -Path $Script:LogPath -ItemType Directory -Force | Out-Null
    }
    $Script:LogFile = Join-Path $Script:LogPath "DCU-Install-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
}

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Output to console with colors
    switch ($Level) {
        'ERROR'   { Write-Host $logEntry -ForegroundColor Red }
        'WARN'    { Write-Host $logEntry -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
        default   { Write-Host $logEntry }
    }
    
    # Write to log file
    Add-Content -Path $Script:LogFile -Value $logEntry
}

function Test-DellSystem {
    Write-Log "Checking if system is Dell hardware"
    
    try {
        $manufacturer = (Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop).Manufacturer
        if ($manufacturer -like '*Dell*') {
            Write-Log "Dell system detected: $manufacturer" -Level 'SUCCESS'
            return $true
        } else {
            Write-Log "Non-Dell system detected: $manufacturer" -Level 'WARN'
            return $false
        }
    }
    catch {
        Write-Log "Failed to determine system manufacturer: $($_.Exception.Message)" -Level 'ERROR'
        return $false
    }
}

function Get-InstalledDCU {
    Write-Log "Checking for existing Dell Command Update installations"
    
    $registryPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    
    $dcuApps = Get-ItemProperty -Path $registryPaths -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like '*Dell Command*Update*' }
    
    if ($dcuApps) {
        foreach ($app in $dcuApps) {
            Write-Log "Found: $($app.DisplayName) v$($app.DisplayVersion)"
        }
        return $dcuApps
    } else {
        Write-Log "No Dell Command Update installations found"
        return $null
    }
}

function Test-DCUVersion {
    $currentDCU = Get-InstalledDCU
    
    if (-not $currentDCU) {
        Write-Log "No DCU installed - installation required"
        return $false
    }
    
    # Check if any DCU version is below 5.0 (these need to be removed)
    $oldDCU = $currentDCU | Where-Object {
        $_.DisplayVersion -and [version]$_.DisplayVersion -lt $Script:DCUMinVersion
    }
    
    if ($oldDCU) {
        Write-Log "Found DCU version(s) below 5.0 - these must be removed as updates don't work properly" -Level 'WARN'
        foreach ($old in $oldDCU) {
            Write-Log "  - $($old.DisplayName) v$($old.DisplayVersion) (requires removal)"
        }
        return $false
    }
    
    # Check for modern DCU version (5.0+)
    $modernDCU = $currentDCU | Where-Object {
        $_.DisplayName -like '*Dell Command | Update*' -and
        $_.DisplayVersion -and
        [version]$_.DisplayVersion -ge $Script:DCUMinVersion
    }
    
    if ($modernDCU -and -not $ForceReinstall) {
        Write-Log "Current DCU version acceptable: $($modernDCU.DisplayVersion)" -Level 'SUCCESS'
        return $true
    } elseif ($ForceReinstall) {
        Write-Log "Force reinstall requested - will update DCU" -Level 'INFO'
        return $false
    } else {
        Write-Log "DCU installation incomplete or incompatible - update required" -Level 'WARN'
        return $false
    }
}

function Test-ChocolateyInstalled {
    Write-Log "Checking for Chocolatey installation"
    
    $chocoCommand = Get-Command choco.exe -ErrorAction SilentlyContinue
    if ($chocoCommand) {
        try {
            $chocoVersion = & choco.exe --version 2>$null
            Write-Log "Chocolatey found: v$chocoVersion" -Level 'SUCCESS'
            return $true
        }
        catch {
            Write-Log "Chocolatey executable found but not functional" -Level 'WARN'
            return $false
        }
    } else {
        Write-Log "Chocolatey not found"
        return $false
    }
}

function Install-Chocolatey {
    Write-Log "Installing Chocolatey"
    
    try {
        # Following Chocolatey's official installation method
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        
        $installScript = (New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')
        Invoke-Expression $installScript
        
        # Refresh environment to pick up choco
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
        
        if (Test-ChocolateyInstalled) {
            Write-Log "Chocolatey installed successfully" -Level 'SUCCESS'
            return $true
        } else {
            throw "Chocolatey installation failed verification"
        }
    }
    catch {
        Write-Log "Cannot access community.chocolatey.org or install Chocolatey: $($_.Exception.Message)" -Level 'ERROR'
        return $false
    }
}

function Stop-DCUProcesses {
    Write-Log "Stopping Dell Command Update processes"
    
    $processesToStop = @(
        'DellCommandUpdate',
        'dcu-cli',
        'DCU',
        'SupportAssistAgent',
        'SupportAssistClientUI',
        'DellClientManagementService'
    )
    
    foreach ($processName in $processesToStop) {
        $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
        if ($processes) {
            Write-Log "Stopping process: $processName"
            try {
                $processes | Stop-Process -Force -ErrorAction Stop
                Write-Log "Successfully stopped $processName" -Level 'SUCCESS'
            }
            catch {
                Write-Log "Failed to stop $processName`: $($_.Exception.Message)" -Level 'WARN'
            }
        }
    }
    
    # Wait for processes to fully terminate
    Start-Sleep -Seconds 3
}

function Remove-IncompatibleSoftware {
    Write-Log "Identifying and removing incompatible Dell software"
    
    $registryPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    
    # Find applications to remove
    $appsToRemove = @()
    
    # Get all Dell applications that must be removed
    $allDellApps = Get-ItemProperty -Path $registryPaths -ErrorAction SilentlyContinue |
        Where-Object { 
            # Remove SupportAssist (always conflicts)
            $_.DisplayName -like '*SupportAssist*' -or
            # Remove old Dell Update variants
            $_.DisplayName -like '*Dell Update*' -or
            # Remove ANY DCU version below 5.0 (updates don't work properly)
            ($_.DisplayName -like '*Dell Command*Update*' -and 
             $_.DisplayVersion -and 
             [version]$_.DisplayVersion -lt $Script:DCUMinVersion) -or
            # Remove DCU without version info (likely corrupted)
            ($_.DisplayName -like '*Dell Command*Update*' -and 
             (-not $_.DisplayVersion -or $_.DisplayVersion -eq ''))
        }
    
    foreach ($app in $allDellApps) {
        if ($app.UninstallString) {
            $appsToRemove += $app
            if ($app.DisplayName -like '*Dell Command*Update*') {
                Write-Log "Marking DCU v$($app.DisplayVersion) for removal (below v5.0 - updates don't work)" -Level 'WARN'
            }
        }
    }
    
    if ($appsToRemove.Count -eq 0) {
        Write-Log "No incompatible software found"
        return
    }
    
    Write-Log "Found $($appsToRemove.Count) applications to remove"
    
    foreach ($app in $appsToRemove) {
        Write-Log "Removing: $($app.DisplayName) v$($app.DisplayVersion)"
        
        try {
            if ($app.UninstallString -match 'msiexec') {
                # MSI uninstall
                $guid = [regex]::Match($app.UninstallString, '\{[0-9a-fA-F-]{36}\}').Value
                if ($guid) {
                    $process = Start-Process -FilePath 'msiexec.exe' -ArgumentList "/x `"$guid`" /quiet /norestart" -Wait -PassThru -NoNewWindow
                    if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 1605) {
                        Write-Log "Successfully removed $($app.DisplayName)" -Level 'SUCCESS'
                    } else {
                        Write-Log "Uninstall returned exit code: $($process.ExitCode)" -Level 'WARN'
                    }
                }
            } else {
                # EXE uninstall
                $uninstallArgs = '/quiet /norestart'
                if ($app.UninstallString -match '"([^"]+)"(.*)') {
                    $exePath = $matches[1]
                    $existingArgs = $matches[2].Trim()
                    if ($existingArgs -and $existingArgs -notmatch '/quiet') {
                        $uninstallArgs = "$existingArgs /quiet"
                    }
                } else {
                    $exePath = $app.UninstallString
                }
                
                if (Test-Path $exePath) {
                    $process = Start-Process -FilePath $exePath -ArgumentList $uninstallArgs -Wait -PassThru -NoNewWindow
                    Write-Log "Uninstall process completed with exit code: $($process.ExitCode)"
                }
            }
        }
        catch {
            Write-Log "Failed to remove $($app.DisplayName)`: $($_.Exception.Message)" -Level 'WARN'
        }
    }
    
    # Clean up any remaining processes
    Stop-DCUProcesses
}

function Install-DCUViaChocolatey {
    Write-Log "Installing Dell Command Update via Chocolatey"
    
    try {
        # Following Chocolatey best practices for scripting
        $chocoArgs = @(
            'upgrade'
            'dellcommandupdate'
            '-y'
            "--source=`"$ChocolateySource`""
        )
        
        Write-Log "Executing: choco $($chocoArgs -join ' ')"
        
        $process = Start-Process -FilePath 'choco' -ArgumentList $chocoArgs -Wait -PassThru -NoNewWindow
        $exitCode = $process.ExitCode
        
        Write-Log "Chocolatey process completed with exit code: $exitCode"
        
        if ($Script:ValidChocolateyExitCodes -contains $exitCode) {
            Write-Log "Dell Command Update installed successfully via Chocolatey" -Level 'SUCCESS'
            
            # Verify installation
            Start-Sleep -Seconds 5
            $dcuPath = Get-ChildItem -Path "$env:SystemDrive\Program Files*\Dell\CommandUpdate\dcu-cli.exe" -ErrorAction SilentlyContinue
            if ($dcuPath) {
                Write-Log "DCU CLI verified at: $($dcuPath.FullName)" -Level 'SUCCESS'
                
                # Test DCU functionality
                try {
                    $versionOutput = & $dcuPath.FullName /version 2>$null
                    Write-Log "DCU version check successful: $versionOutput" -Level 'SUCCESS'
                }
                catch {
                    Write-Log "DCU installed but version check failed" -Level 'WARN'
                }
            } else {
                Write-Log "DCU CLI not found after installation" -Level 'WARN'
            }
            
            return $true
        } else {
            Write-Log "Chocolatey installation failed with exit code: $exitCode" -Level 'ERROR'
            return $false
        }
    }
    catch {
        Write-Log "Chocolatey installation failed: $($_.Exception.Message)" -Level 'ERROR'
        return $false
    }
}

function Main {
    Initialize-Logging
    Write-Log "=== Dell Command Update Management Script Started ===" -Level 'INFO'
    
    # Set PowerShell preferences
    $ProgressPreference = 'SilentlyContinue'
    $ErrorActionPreference = 'Stop'
    
    try {
        # Step 1: Verify Dell system
        if (-not (Test-DellSystem)) {
            Write-Log "Script not applicable to non-Dell systems. Exiting." -Level 'WARN'
            return 0
        }
        
        # Step 2: Check current DCU status
        if (Test-DCUVersion) {
            Write-Log "Dell Command Update is already current. No action needed." -Level 'SUCCESS'
            return 0
        }
        
        # Step 3: Ensure Chocolatey is available
        if (-not (Test-ChocolateyInstalled)) {
            Write-Log "Installing Chocolatey..."
            if (-not (Install-Chocolatey)) {
                Write-Log "Cannot proceed without Chocolatey. Exiting." -Level 'ERROR'
                return 1
            }
        }
        
        # Step 4: Stop processes and remove incompatible software
        Stop-DCUProcesses
        Remove-IncompatibleSoftware
        
        # Step 5: Install DCU via Chocolatey
        if (Install-DCUViaChocolatey) {
            Write-Log "=== Dell Command Update management completed successfully ===" -Level 'SUCCESS'
            return 0
        } else {
            Write-Log "=== Dell Command Update management failed ===" -Level 'ERROR'
            return 1
        }
    }
    catch {
        Write-Log "Script execution failed: $($_.Exception.Message)" -Level 'ERROR'
        return 1
    }
    finally {
        # Keep PowerShell window open when run interactively
        if ($Host.Name -eq "ConsoleHost") {
            Write-Host "`nScript completed. Press any key to close this window..." -ForegroundColor Cyan
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    }
}

# Execute main function
Main
