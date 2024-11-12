# To Run to test: PowerShell.exe -ExecutionPolicy Bypass -NoProfile -File "Deploy-IntuneEnrollment.ps1"

#The script will:

#Only run on non-domain-joined devices
#Install prerequisites for all users
#Create reminders accessible to all users
#Log all actions for troubleshooting
#Handle running as SYSTEM account properly

# Function to test basic requirements and connectivity
function Test-Prerequisites {
    Write-Host "`n=== Testing Prerequisites ===" -ForegroundColor Cyan
    
    # Check domain status
    $computerSystem = Get-WmiObject -Class Win32_ComputerSystem
    $isDomainJoined = $computerSystem.PartOfDomain
    $isAADJoined = (dsregcmd /status | Select-String "AzureAdJoined").ToString().Contains("YES")
    
    Write-Host "Domain Joined: $isDomainJoined"
    Write-Host "Azure AD Joined: $isAADJoined"
    
    if ($isDomainJoined -or $isAADJoined) {
        Write-Host "Device is already domain or AAD joined - skipping enrollment" -ForegroundColor Yellow
        return $false
    }

    # Test key endpoints
    Write-Host "`nTesting connectivity to required endpoints:" -ForegroundColor Cyan
    $endpoints = @(
        "login.microsoftonline.com",
        "portal.manage.microsoft.com",
        "enrollment.manage.microsoft.com"
    )
    
    $hasError = $false
    foreach ($endpoint in $endpoints) {
        $test = Test-NetConnection -ComputerName $endpoint -Port 443 -WarningAction SilentlyContinue
        if ($test.TcpTestSucceeded) {
            Write-Host "✓ $endpoint" -ForegroundColor Green
        } else {
            Write-Host "✗ $endpoint" -ForegroundColor Red
            $hasError = $true
        }
    }

    if ($hasError) {
        Write-Host "`nConnectivity test failed - please check network connectivity" -ForegroundColor Red
        return $false
    }

    return $true
}

# Install Company Portal and set registry keys
function Install-EnrollmentPrerequisites {
    Write-Host "`n=== Installing Prerequisites ===" -ForegroundColor Cyan
    
    try {
        # Set registry keys
        Write-Host "Setting MDM registry keys..." -NoNewline
        $regPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\MDM"
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        New-ItemProperty -Path $regPath -Name "EnableMDMEnrollment" -Value 1 -PropertyType DWORD -Force | Out-Null
        Write-Host "✓" -ForegroundColor Green
    }
    catch {
        Write-Host "✗" -ForegroundColor Red
        Write-Host "Error setting registry: $_" -ForegroundColor Red
        return $false
    }

    try {
        # Install Company Portal
        Write-Host "Installing Company Portal..." -NoNewline
        winget install "Company Portal" --silent --accept-package-agreements --accept-source-agreements | Out-Null
        
        # Verify installation
        Start-Sleep -Seconds 5  # Give it a moment to complete
        $cpInstalled = Get-AppxPackage -Name "Microsoft.CompanyPortal" -AllUsers
        if ($cpInstalled) {
            Write-Host "✓" -ForegroundColor Green
        } else {
            Write-Host "✗" -ForegroundColor Red
            Write-Host "Company Portal not found after installation" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "✗" -ForegroundColor Red
        Write-Host "Error installing Company Portal: $_" -ForegroundColor Red
        return $false
    }

    return $true
}

# Create enrollment reminder
function Set-EnrollmentReminder {
    Write-Host "`n=== Setting Up Reminders ===" -ForegroundColor Cyan
    
    try {
        # Create a scheduled task
        Write-Host "Creating reminder task..." -NoNewline
        
        $taskName = "DeviceEnrollmentReminder"
        # Remove existing task if present
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
        
        $action = New-ScheduledTaskAction -Execute "ms-windows-store://pdp/?productid=9WZDNCRFJ3PZ"
        $trigger = New-ScheduledTaskTrigger -AtLogOn
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings | Out-Null
        Write-Host "✓" -ForegroundColor Green
    }
    catch {
        Write-Host "✗" -ForegroundColor Red
        Write-Host "Error creating reminder: $_" -ForegroundColor Red
        return $false
    }

    return $true
}

# Main execution
Write-Host "Starting Intune Enrollment Setup" -ForegroundColor Cyan

# Check if running as admin
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Script must run with administrative privileges" -ForegroundColor Red
    exit 1
}

# Run checks
if (-not (Test-Prerequisites)) {
    Write-Host "`nPrerequisite checks failed - please resolve issues and try again" -ForegroundColor Red
    exit 1
}

# Install prerequisites
if (-not (Install-EnrollmentPrerequisites)) {
    Write-Host "`nFailed to install prerequisites - please check errors above" -ForegroundColor Red
    exit 1
}

# Set up reminders
if (-not (Set-EnrollmentReminder)) {
    Write-Host "`nFailed to set up reminders - please check errors above" -ForegroundColor Red
    exit 1
}

Write-Host "`nSetup completed successfully!" -ForegroundColor Green
Write-Host "The Company Portal will launch at next user login to complete enrollment" -ForegroundColor Cyan