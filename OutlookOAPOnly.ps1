#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Configures Outlook to use Classic mode by default for all user profiles.
.DESCRIPTION
    This script configures registry settings for all local user profiles to ensure
    Outlook launches in Classic mode by default and prevents automatic migration to
    the new Outlook experience while still allowing manual switching.
.NOTES
    - Requires Administrator privileges
    - Runs silently without prompts
    - Does not close Outlook if running
    - Safe to run on machines with or without Outlook installed
#>

# Ensure we're running as Administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
    Write-Error "This script must be run as Administrator"
    exit 1
}

# Install/update RunAsUser module if needed
$minimumVersion = [Version]"2.4.0"
$moduleName = "RunAsUser"

$module = Get-Module -ListAvailable -Name $moduleName | Where-Object { $_.Version -ge $minimumVersion }
if ($null -eq $module) {
    Write-Output "Installing/updating $moduleName module (v$minimumVersion or higher)..."
    
    # Ensure TLS 1.2 is used for PowerShell Gallery
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    
    # Install NuGet package provider if needed
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -Force -Scope CurrentUser | Out-Null
    }
    
    # Set PSGallery as trusted
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    
    # Install the module
    Install-Module -Name $moduleName -MinimumVersion $minimumVersion -Force -AllowClobber
}

Import-Module $moduleName

# Create a function to set Outlook registry settings
function Set-OutlookClassicMode {
    param (
        [string]$RegistryPath,
        [string]$UserName = ""
    )
    
    try {
        # Create Outlook Preferences path if it doesn't exist
        $outlookPreferencesPath = "$RegistryPath\Software\Microsoft\Office\16.0\Outlook\Preferences"
        if (-not (Test-Path -Path $outlookPreferencesPath)) {
            New-Item -Path $outlookPreferencesPath -Force | Out-Null
        }
        
        # Create Outlook Options\General path if it doesn't exist
        $outlookOptionsGeneralPath = "$RegistryPath\Software\Microsoft\Office\16.0\Outlook\Options\General"
        if (-not (Test-Path -Path $outlookOptionsGeneralPath)) {
            New-Item -Path $outlookOptionsGeneralPath -Force | Out-Null
        }
        
        # Set UseNewOutlook = 0 in Preferences
        Set-ItemProperty -Path $outlookPreferencesPath -Name "UseNewOutlook" -Value 0 -Type DWORD -Force
        
        # Set DoNewOutlookAutoMigration = 0 in Options\General
        Set-ItemProperty -Path $outlookOptionsGeneralPath -Name "DoNewOutlookAutoMigration" -Value 0 -Type DWORD -Force
        
        if ($UserName) {
            Write-Output "Successfully configured Outlook Classic mode for $UserName"
        }
        return $true
    }
    catch {
        if ($UserName) {
            Write-Warning "Failed to set registry for $UserName`: $_"
        }
        return $false
    }
}

# Script block for RunAsUser
$scriptBlock = {
    # Create registry keys if they don't exist and set values
    $outlookPreferencesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Preferences"
    if (-not (Test-Path -Path $outlookPreferencesPath)) {
        New-Item -Path $outlookPreferencesPath -Force | Out-Null
    }
    
    $outlookOptionsGeneralPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General"
    if (-not (Test-Path -Path $outlookOptionsGeneralPath)) {
        New-Item -Path $outlookOptionsGeneralPath -Force | Out-Null
    }
    
    # Set UseNewOutlook = 0 in Preferences
    Set-ItemProperty -Path $outlookPreferencesPath -Name "UseNewOutlook" -Value 0 -Type DWORD -Force
    
    # Set DoNewOutlookAutoMigration = 0 in Options\General
    Set-ItemProperty -Path $outlookOptionsGeneralPath -Name "DoNewOutlookAutoMigration" -Value 0 -Type DWORD -Force
}

# Step 1: Process active sessions directly from HKEY_USERS
Write-Output "Step 1: Processing active user sessions..."
$activeSids = Get-ChildItem -Path "Registry::HKEY_USERS" | Where-Object { $_.PSChildName -match '^S-1-5-21-' }

foreach ($sidKey in $activeSids) {
    try {
        $sid = $sidKey.PSChildName
        $registryPath = "Registry::HKEY_USERS\$sid"
        
        # Try to get the username for logging purposes
        $username = "Unknown"
        try {
            $sidObj = New-Object System.Security.Principal.SecurityIdentifier($sid)
            $account = $sidObj.Translate([System.Security.Principal.NTAccount])
            $username = ($account.Value -split '\\')[1]
        } catch {
            # Continue with SID if username can't be resolved
        }
        
        Write-Output "Processing active user session for $username (SID: $sid)..."
        $result = Set-OutlookClassicMode -RegistryPath $registryPath -UserName $username
        
        # If direct registry edit failed, try RunAsUser as backup
        if (-not $result) {
            Write-Output "Attempting to use RunAsUser for $username..."
            try {
                Invoke-AsCurrentUser -ScriptBlock $scriptBlock -UseWindowsPowerShell
                Write-Output "Successfully configured Outlook Classic mode via RunAsUser for $username"
            } catch {
                Write-Warning "RunAsUser also failed for $username`: $_"
            }
        }
    } catch {
        Write-Warning "Error processing SID $($sidKey.PSChildName)`: $_"
    }
}

# Step 2: Process offline users by loading their registry hives
Write-Output "`nStep 2: Processing offline user profiles..."
$usersFolder = "C:\Users"
$processedUsers = @{}

# First record which users were already processed in Step 1
foreach ($sidKey in $activeSids) {
    try {
        $sidObj = New-Object System.Security.Principal.SecurityIdentifier($sidKey.PSChildName)
        $account = $sidObj.Translate([System.Security.Principal.NTAccount])
        $username = ($account.Value -split '\\')[1]
        $processedUsers[$username] = $true
    } catch {
        # Skip if username can't be resolved
    }
}

# Now process offline users
Get-ChildItem -Path $usersFolder -Directory | 
Where-Object { 
    $_.Name -ne "Public" -and 
    $_.Name -ne "Default" -and 
    $_.Name -ne "Default User" -and
    -not $processedUsers.ContainsKey($_.Name)
} | ForEach-Object {
    $username = $_.Name
    Write-Output "Processing offline user profile for $username..."
    
    # Check if NTUSER.DAT exists
    $userHivePath = Join-Path -Path $_.FullName -ChildPath "NTUSER.DAT"
    if (Test-Path -Path $userHivePath) {
        # Generate a unique name for the registry load point
        $hiveName = "HKU_$($username -replace '[^a-zA-Z0-9]', '')"
        
        # Unload hive if previously loaded
        reg unload "HKU\$hiveName" 2>$null
        
        try {
            # Load user's registry hive
            Write-Output "Loading registry hive for $username..."
            $result = reg load "HKU\$hiveName" "$userHivePath" 2>&1
            
            if ($LASTEXITCODE -eq 0) {
                # Modify registry settings
                $registryPath = "Registry::HKEY_USERS\$hiveName"
                Set-OutlookClassicMode -RegistryPath $registryPath -UserName $username
                
                # Force garbage collection to ensure file handles are released
                [gc]::Collect()
                [gc]::WaitForPendingFinalizers()
                
                # Unload the hive
                reg unload "HKU\$hiveName" 2>$null
            } else {
                if ($result -like "*because it is being used by another process*") {
                    Write-Output "User $username is logged in. This profile was or will be processed in Step 1."
                } else {
                    Write-Warning "Failed to load registry hive for $username`: $result"
                }
            }
        } catch {
            Write-Warning "Error processing offline user $username`: $_"
            # Attempt to unload in case of failure
            reg unload "HKU\$hiveName" 2>$null
        }
    } else {
        Write-Warning "NTUSER.DAT not found for user $username"
    }
}

# Step 3: Use RunAsUser as final attempt to catch any active users we might have missed
Write-Output "`nStep 3: Final pass using RunAsUser for any remaining active sessions..."
try {
    # This will run for the currently active user context(s)
    Invoke-AsCurrentUser -ScriptBlock $scriptBlock -UseWindowsPowerShell
    Write-Output "Successfully ran final RunAsUser pass"
} catch {
    Write-Warning "Error in final RunAsUser pass: $_"
}

Write-Output "`nCompleted configuring Outlook Classic mode for all users."
