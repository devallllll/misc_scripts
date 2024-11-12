# To Run to test: PowerShell.exe -ExecutionPolicy Bypass -NoProfile -File "Deploy-IntuneEnrollment.ps1"

#The script will:

#Only run on non-domain-joined devices
#Install prerequisites for all users
#Create reminders accessible to all users
#Log all actions for troubleshooting
#Handle running as SYSTEM account properly

# Function to check if device is already domain joined
function Check-DomainStatus {
    try {
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem
        $isDomainJoined = $computerSystem.PartOfDomain
        $isAADJoined = (dsregcmd /status | Select-String "AzureAdJoined").ToString().Contains("YES")
        
        Write-Output "Domain Join Status: $isDomainJoined"
        Write-Output "Azure AD Join Status: $isAADJoined"
        
        if ($isDomainJoined -or $isAADJoined) {
            return $true
        }
        return $false
    }
    catch {
        Write-Error "Error checking domain status: $_"
        return $true # Fail safe - if we can't check, assume it's joined
    }
}

# Function to get current interactive user when running as SYSTEM
function Get-InteractiveUser {
    try {
        $explorerProcesses = Get-WmiObject -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'"
        if ($explorerProcesses) {
            foreach ($explorer in $explorerProcesses) {
                $owner = $explorer.GetOwner()
                if ($owner.User -ne "SYSTEM") {
                    return @{
                        Username = $owner.User
                        Domain = $owner.Domain
                    }
                }
            }
        }
        Write-Output "No interactive user found"
        return $null
    }
    catch {
        Write-Error "Error getting interactive user: $_"
        return $null
    }
}

# Create registry values for all users
function Set-RegistryForAllUsers {
    try {
        # Default user profile hive
        $defaultUserPath = "C:\Users\Default\NTUSER.DAT"
        
        # Load the default user hive
        reg load "HKU\DefaultUser" $defaultUserPath

        # Set registry for default profile
        reg add "HKU\DefaultUser\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\MDM" /v "EnableMDMEnrollment" /t REG_DWORD /d 1 /f

        # Unload the hive
        [gc]::Collect()
        reg unload "HKU\DefaultUser"

        # Set for HKLM
        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\MDM"
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        New-ItemProperty -Path $regPath -Name "EnableMDMEnrollment" -Value 1 -PropertyType DWORD -Force
    }
    catch {
        Write-Error "Error setting registry: $_"
    }
}

# Install Company Portal for all users
function Install-CompanyPortal {
    try {
        # Using winget to install Company Portal silently
        $env:PROGRAMDATA + "\Microsoft\Windows\Start Menu\Programs\StartUp"
        winget install "Company Portal" --silent --scope machine --accept-package-agreements --accept-source-agreements
        
        # Create shortcut in public desktop
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\Enroll Device.lnk")
        $Shortcut.TargetPath = "ms-windows-store://pdp/?productid=9WZDNCRFJ3PZ"
        $Shortcut.Description = "Enroll your device for required updates"
        $Shortcut.Save()
    }
    catch {
        Write-Error "Error installing Company Portal: $_"
    }
}

# Create a scheduled task to remind users
function Create-EnrollmentReminder {
    try {
        # Create a scheduled task that runs for all users
        $taskName = "DeviceEnrollmentReminder"
        $taskDescription = "Reminder to complete device enrollment"
        
        # Delete existing task if it exists
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        
        # Create the task action
        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
            -Argument "-WindowStyle Hidden -Command Start-Process 'ms-windows-store://pdp/?productid=9WZDNCRFJ3PZ'"
        
        # Create triggers - at logon and every 4 hours
        $logonTrigger = New-ScheduledTaskTrigger -AtLogOn
        $repeatTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 4)
        
        # Specify that task can run on battery and doesn't stop on battery
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 1)
        
        # Register the task to run for all users
        Register-ScheduledTask -TaskName $taskName `
            -Description $taskDescription `
            -Trigger @($logonTrigger, $repeatTrigger) `
            -Action $action `
            -Settings $settings `
            -Force
    }
    catch {
        Write-Error "Error creating reminder task: $_"
    }
}

# Main execution block
$logPath = "$env:ProgramData\IntuneMigration.log"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

try {
    # Start logging
    Add-Content -Path $logPath -Value "[$timestamp] Starting enrollment script"
    
    # Check if already domain joined
    if (Check-DomainStatus) {
        Add-Content -Path $logPath -Value "[$timestamp] Device is already domain/AAD joined. Exiting."
        exit 0
    }
    
    # Get interactive user
    $currentUser = Get-InteractiveUser
    Add-Content -Path $logPath -Value "[$timestamp] Current interactive user: $($currentUser.Username)"
    
    # Set registry values
    Set-RegistryForAllUsers
    Add-Content -Path $logPath -Value "[$timestamp] Registry values set"
    
    # Install Company Portal
    Install-CompanyPortal
    Add-Content -Path $logPath -Value "[$timestamp] Company Portal installation attempted"
    
    # Create reminder task
    Create-EnrollmentReminder
    Add-Content -Path $logPath -Value "[$timestamp] Reminder task created"
    
    Add-Content -Path $logPath -Value "[$timestamp] Script completed successfully"
}
catch {
    Add-Content -Path $logPath -Value "[$timestamp] ERROR: $($_.Exception.Message)"
    throw
}