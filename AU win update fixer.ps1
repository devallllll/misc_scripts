# Windows Update Registry Cleanup for RMM
# Remove all WindowsUpdate policies except needed ones for RMM control

$regPath = "HKLM:\SOFTWARE\WOW6432Node\Policies\Microsoft\Windows\WindowsUpdate"

# Show current state - check all known Windows Update registry keys
Write-Output "BEFORE CLEANUP - ALL WINDOWS UPDATE REGISTRY KEYS:"
Write-Output "=============================================="

if (Test-Path $regPath) {
    Write-Output "Main WindowsUpdate Keys:"
    $mainKeys = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
    if ($mainKeys) {
        $knownMainKeys = @("SetDisableUXWUAccess", "WUServer", "WUStatusServer", "TargetGroup", "TargetGroupEnabled", "ElevateNonAdmins")
        foreach ($key in $knownMainKeys) {
            if ($mainKeys.$key -ne $null) { Write-Output "  $key = $($mainKeys.$key)" }
        }
        # Show any other keys not in our known list
        $mainKeys.PSObject.Properties | Where-Object { 
            $_.Name -notin @("PSPath", "PSParentPath", "PSChildName", "PSDrive", "PSProvider") -and
            $_.Name -notin $knownMainKeys
        } | ForEach-Object { Write-Output "  $($_.Name) = $($_.Value) [UNKNOWN KEY]" }
    }
    
    if (Test-Path "$regPath\AU") {
        Write-Output "AU Subkey:"
        $auKeys = Get-ItemProperty -Path "$regPath\AU" -ErrorAction SilentlyContinue
        if ($auKeys) {
            $knownAUKeys = @("NoAutoUpdate", "AUOptions", "ScheduledInstallDay", "ScheduledInstallTime", 
                             "UseWUServer", "AutoInstallMinorUpdates", "NoAutoRebootWithLoggedOnUsers",
                             "DetectionFrequency", "DetectionFrequencyEnabled", "RebootRelaunchTimeout", 
                             "RebootRelaunchTimeoutEnabled", "RebootWarningTimeout", "RebootWarningTimeoutEnabled",
                             "RescheduleWaitTime", "RescheduleWaitTimeEnabled", "ScheduledInstallEveryWeek")
            foreach ($key in $knownAUKeys) {
                if ($auKeys.$key -ne $null) { Write-Output "  $key = $($auKeys.$key)" }
            }
            # Show any other keys not in our known list
            $auKeys.PSObject.Properties | Where-Object { 
                $_.Name -notin @("PSPath", "PSParentPath", "PSChildName", "PSDrive", "PSProvider") -and
                $_.Name -notin $knownAUKeys
            } | ForEach-Object { Write-Output "  $($_.Name) = $($_.Value) [UNKNOWN KEY]" }
        }
    } else {
        Write-Output "AU Subkey: Does not exist"
    }
} else {
    Write-Output "WindowsUpdate key does not exist"
}

# Remove and recreate
if (Test-Path $regPath) {
    Remove-Item -Path $regPath -Recurse -Force
    Write-Output "Removed existing WindowsUpdate key"
}

# Create clean structure
New-Item -Path $regPath -Force | Out-Null
New-ItemProperty -Path $regPath -Name "SetDisableUXWUAccess" -Value 1 -PropertyType DWORD -Force | Out-Null
New-Item -Path "$regPath\AU" -Force | Out-Null
New-ItemProperty -Path "$regPath\AU" -Name "NoAutoUpdate" -Value 1 -PropertyType DWORD -Force | Out-Null

# Show final state
Write-Output "AFTER CLEANUP:"
Get-ItemProperty -Path $regPath | Select-Object * -ExcludeProperty PS*
Write-Output "AU Subkey:"
Get-ItemProperty -Path "$regPath\AU" | Select-Object * -ExcludeProperty PS*

Write-Output "Cleanup completed - RMM has full control"
