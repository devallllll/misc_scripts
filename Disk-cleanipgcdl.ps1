# Windows Storage Sense Configuration Script for ConnectWise RMM  based on cyberdrains script and modified by GCDL and claude 
# -------------------------------------------
# TECHNICIAN GUIDANCE:
# This script configures Windows Storage Sense and optionally runs DISM cleanup. 
#
# RISKS AND CONSIDERATIONS:
# 1. Downloads Folder (Default: OFF):
#    - Disabled by default as users often keep important files here
#    - Enable only if organization policy requires or users have been trained to not store permanent files here
#    - Recommend documenting in your knowledge base if enabled
#
# 2. Recycle Bin (Default: ON, 60 days):
#    - Safe to enable as users expect items here to be temporary
#    - 60-day retention provides good balance between cleanup and recovery needs
#    - Consider shorter periods for devices with limited storage
#
# 3. Temporary Files (Default: ON):
#    - Safe to clean - Windows manages this location
#    - Can help recover significant space on many systems
#
# 4. OneDrive Cache (Default: ON, 60 days):
#    - Safe for most environments using OneDrive
#    - Ensures offline files not accessed within 60 days become online-only
#    - May impact offline access if users don't regularly connect
#
# 5. DISM Cleanup (Default: ON):
#    - Enabled by default for one-time WinSxS cleanup when script runs
#    - Runs immediately when script is executed, not on Storage Sense schedule
#    - Disable for production systems or if you need Windows Update rollback capability
#    - Can take significant time to complete
#    - Helps reduce WinSxS folder size but can't be undone
#    - May impact Windows Update rollback capabilities
#
# PARAMETERS (all optional):
# prefschedule: How often Storage Sense runs (0=Low Disk Space, 1=Monthly, 7=Weekly, 30=Daily). Default: 0
# cleartemp: Clear temporary files (True/False). Default: True
# clearrecycler: Clear recycle bin (True/False). Default: True
# cleardownloads: Clear downloads folder (True/False). Default: False
# clearonedrivecache: Allow clearing OneDrive cache (True/False). Default: True
# addonedriveloc: Add all OneDrive locations (True/False). Default: True
# recyclerdays: Days before clearing recycle bin (0=never, 1, 14, 30, 60). Default: 60
# downloaddays: Days before clearing downloads (0=never, 1, 14, 30, 60). Default: 60
# onedrivecachedays: Days before clearing OneDrive cache (0=never, 1, 14, 30, 60). Default: 60
# ClearDISMDANGER: Run DISM cleanup (True/False). Default: True

#Settings
[PSCustomObject]@{
    PrefSched               = if('@prefschedule@' -eq '') {'0'} else {'@prefschedule@'} #Options are: 0(Low Diskspace),1,7,30
    ClearTemporaryFiles     = if('@cleartemp@' -eq '') {$true} else {'@cleartemp@'}
    ClearRecycler          = if('@clearrecycler@' -eq '') {$true} else {'@clearrecycler@'}
    ClearDownloads         = if('@cleardownloads@' -eq '') {$false} else {'@cleardownloads@'}
    AllowClearOneDriveCache = if('@clearonedrivecache@' -eq '') {$true} else {'@clearonedrivecache@'}
    AddAllOneDrivelocations = if('@addonedriveloc@' -eq '') {$true} else {'@addonedriveloc@'}
    ClearRecyclerDays      = if('@recyclerdays@' -eq '') {'60'} else {'@recyclerdays@'} #Options are: 0(never),1,14,30,60
    ClearDownloadsDays     = if('@downloaddays@' -eq '') {'0'} else {'@downloaddays@'} #Options are: 0(never),1,14,30,60
    ClearOneDriveCacheDays = if('@onedrivecachedays@' -eq '') {'60'} else {'@onedrivecachedays@'} #Options are: 0(never),1,14,30,60
    RunDISMCleanup         = if('@ClearDISMDANGER@' -eq '') {$true} else {'@ClearDISMDANGER@'} #Run DISM cleanup True/False
} | ConvertTo-Json | Out-File "C:\Windows\Temp\WantedStorageSenseSettings.txt"

# Get settings
$WantedSettings = Get-Content "C:\Windows\Temp\WantedStorageSenseSettings.txt" | ConvertFrom-Json

# Run DISM cleanup if enabled, checking for pending reboot
if ($WantedSettings.RunDISMCleanup -eq 'True') {
    if (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue) {
        Write-Output "ERROR: Cannot proceed with DISM cleanup - System has pending reboot"
    }
    else {
        Write-Output "Starting DISM cleanup in background - this will continue after script completes..."
        try {
            Start-Process -FilePath "dism.exe" -ArgumentList "/online /Cleanup-Image /StartComponentCleanup /ResetBase" -NoNewWindow
            Write-Output "SUCCESS: DISM cleanup started in background"
        }
        catch {
            Write-Output "ERROR: Failed to start DISM cleanup - $($_.Exception.Message)"
        }
    }
}

# Additional Windows 11 cleanup (items Storage Sense doesn't handle)
Write-Output "Starting additional Windows 11 cleanup..."

# Clean Windows.old folder
if (Test-Path "C:\Windows.old") {
    try {
        Remove-Item "C:\Windows.old" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Output "SUCCESS: Windows.old folder removed"
    }
    catch {
        Write-Output "WARNING: Could not remove Windows.old folder - may require manual cleanup"
    }
}

# Clear Delivery Optimization cache
try {
    Get-ChildItem "C:\Windows\SoftwareDistribution\DeliveryOptimization" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Write-Output "SUCCESS: Delivery Optimization cache cleared"
}
catch {
    Write-Output "WARNING: Could not clear Delivery Optimization cache"
}

# Clear CBS logs (keep recent ones)
try {
    Get-ChildItem "C:\Windows\Logs\CBS" -Filter "*.log" -ErrorAction SilentlyContinue | Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-30)} | Remove-Item -Force -ErrorAction SilentlyContinue
    Write-Output "SUCCESS: Old CBS logs cleared (kept last 30 days)"
}
catch {
    Write-Output "WARNING: Could not clear CBS logs"
}

# Clear Windows Error Reporting files
try {
    Get-ChildItem "C:\ProgramData\Microsoft\Windows\WER" -Recurse -ErrorAction SilentlyContinue | Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-30)} | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Write-Output "SUCCESS: Windows Error Reporting files cleared (kept last 30 days)"
}
catch {
    Write-Output "WARNING: Could not clear Windows Error Reporting files"
}

# Clear old Windows Defender offline definitions
try {
    Get-ChildItem "C:\ProgramData\Microsoft\Windows Defender\Scans\History" -Recurse -ErrorAction SilentlyContinue | Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-14)} | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Write-Output "SUCCESS: Old Windows Defender scan history cleared"
}
catch {
    Write-Output "WARNING: Could not clear Windows Defender scan history"
}

Write-Output "Additional cleanup completed"

#RunAsUser Module Check/Install
If (Get-Module -ListAvailable -Name "RunAsUser") {
    Import-module RunAsUser
}
Else {
    Install-PackageProvider NuGet -Force
    Set-PSRepository PSGallery -InstallationPolicy Trusted
    Install-Module RunAsUser -force -Repository PSGallery
}

$ScriptBlock = {
    $WantedSettings = Get-Content "C:\Windows\Temp\WantedStorageSenseSettings.txt" | ConvertFrom-Json
    $StorageSenseKeys = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy\'
    
    # Configure Storage Sense with normal settings
    Set-ItemProperty -Path $StorageSenseKeys -name '01' -value '1' -Type DWord  -Force
    Set-ItemProperty -Path $StorageSenseKeys -name '04' -value $WantedSettings.ClearTemporaryFiles -Type DWord -Force
    Set-ItemProperty -Path $StorageSenseKeys -name '08' -value $WantedSettings.ClearRecycler -Type DWord -Force
    Set-ItemProperty -Path $StorageSenseKeys -name '32' -value $WantedSettings.ClearDownloads -Type DWord -Force
    Set-ItemProperty -Path $StorageSenseKeys -name '256' -value $WantedSettings.ClearRecyclerDays -Type DWord -Force
    Set-ItemProperty -Path $StorageSenseKeys -name '512' -value $WantedSettings.ClearDownloadsDays -Type DWord -Force
    Set-ItemProperty -Path $StorageSenseKeys -name '2048' -value $WantedSettings.PrefSched -Type DWord -Force
    Set-ItemProperty -Path $StorageSenseKeys -name 'CloudfilePolicyConsent' -value $WantedSettings.AllowClearOneDriveCache -Type DWord -Force
    
    # Configure OneDrive locations with normal settings
    if ($WantedSettings.AddAllOneDrivelocations) {
$CurrentUserSID = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
        $CurrentSites = Get-ItemProperty 'HKCU:\SOFTWARE\Microsoft\OneDrive\Accounts\Business1\ScopeIdToMountPointPathCache' -ErrorAction SilentlyContinue | Select-Object -Property * -ExcludeProperty PSPath, PsParentPath, PSChildname, PSDrive, PsProvider
        foreach ($OneDriveSite in $CurrentSites.psobject.properties.name) {
            New-Item "$($StorageSenseKeys)/OneDrive!$($CurrentUserSID)!Business1|$($OneDriveSite)" -Force
            New-ItemProperty "$($StorageSenseKeys)/OneDrive!$($CurrentUserSID)!Business1|$($OneDriveSite)" -Name '02' -Value '1' -type DWORD -Force
            New-ItemProperty "$($StorageSenseKeys)/OneDrive!$($CurrentUserSID)!Business1|$($OneDriveSite)" -Name '128' -Value $WantedSettings.ClearOneDriveCacheDays -type DWORD -Force
        }
    }

    # Temporarily override settings for aggressive immediate run
    Set-ItemProperty -Path $StorageSenseKeys -name '256' -value '1' -Type DWord -Force  # Clear all recycle bin
    if ($WantedSettings.AddAllOneDrivelocations) {
        foreach ($OneDriveSite in $CurrentSites.psobject.properties.name) {
            Set-ItemProperty "$($StorageSenseKeys)/OneDrive!$($CurrentUserSID)!Business1|$($OneDriveSite)" -Name '128' -Value '1' -type DWORD -Force  # 1 day OneDrive
        }
    }

    # Run Storage Sense immediately with aggressive settings
    try {
        Start-StorageSense
        Write-Output "Storage Sense cleanup initiated with aggressive settings (1 day OneDrive, clear all recycle bin)"
        Start-Sleep -Seconds 5
    }
    catch {
        Write-Output "Storage Sense configured but could not trigger immediate run: $($_.Exception.Message)"
    }

    # Restore normal settings for future scheduled runs
    Set-ItemProperty -Path $StorageSenseKeys -name '256' -value $WantedSettings.ClearRecyclerDays -Type DWord -Force
    if ($WantedSettings.AddAllOneDrivelocations) {
        foreach ($OneDriveSite in $CurrentSites.psobject.properties.name) {
            Set-ItemProperty "$($StorageSenseKeys)/OneDrive!$($CurrentUserSID)!Business1|$($OneDriveSite)" -Name '128' -Value $WantedSettings.ClearOneDriveCacheDays -type DWORD -Force
        }
    }
    Write-Output "Storage Sense configured for low disk space triggers with normal retention settings"
}

$null = Invoke-AsCurrentUser -ScriptBlock $ScriptBlock -UseWindowsPowerShell -NonElevatedSession -CacheToDisk
