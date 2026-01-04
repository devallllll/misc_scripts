#requires -RunAsAdministrator
# Disk Cleanup - Admin/SYSTEM - GoodChoice IT Ltd

$log = @()
$SpaceFreed = 0
$SFCClean = $false

function Get-FolderSize { param([string]$Path)
    if (Test-Path $Path) {
        try { return [math]::Round(((Get-ChildItem $Path -Recurse -Force -EA 0 | Measure-Object Length -Sum -EA 0).Sum / 1GB), 2) } catch { return 0 }
    }
    return 0
}

# Health Check
$log += "Running SFC..."
$SFCResult = & sfc /scannow 2>&1 | Out-String

if ($SFCResult -match "did not find any integrity violations") {
    $HealthStatus = "Clean"; $SFCClean = $true; $log += "SFC: Clean"
} elseif ($SFCResult -match "successfully repaired") {
    $HealthStatus = "Repaired"; $SFCClean = $true; $log += "SFC: Repaired"
} elseif ($SFCResult -match "unable to fix") {
    $log += "SFC: Running DISM..."
    & DISM /Online /Cleanup-Image /RestoreHealth 2>&1 | Out-Null
    $SFCResult2 = & sfc /scannow 2>&1 | Out-String
    if ($SFCResult2 -match "did not find any integrity violations") {
        $HealthStatus = "Repaired"; $SFCClean = $true; $log += "SFC: Fixed after DISM"
    } else {
        $HealthStatus = "Failed"; $log += "SFC: Still has errors"
    }
}

# Cleanup
$log += "Starting cleanup..."
& DISM /Online /Cleanup-Image /StartComponentCleanup /Quiet 2>&1 | Out-Null
$log += "DISM: Component cleanup done"

if ($SFCClean) {
    & DISM /Online /Cleanup-Image /StartComponentCleanup /ResetBase /Quiet 2>&1 | Out-Null
    $log += "DISM: ResetBase done"
} else {
    $log += "DISM: Skipped ResetBase (errors present)"
}

# Clean folders
@{
    "$env:windir\Logs\DISM" = "DISM Logs"
    "$env:windir\SoftwareDistribution\Download" = "Update Cache"
    "$env:windir\Temp" = "Windows Temp"
    "$env:ProgramData\Microsoft\Windows\WER" = "Error Reports"
}.GetEnumerator() | ForEach-Object {
    if (Test-Path $_.Key) {
        $size = Get-FolderSize $_.Key
        Get-ChildItem $_.Key -Recurse -Force -EA 0 | Remove-Item -Force -Recurse -EA 0
        if ($size -gt 0) { $SpaceFreed += $size; $log += "Cleaned: $($_.Value) - $size GB" }
    }
}

# CBS.log
$cbsLog = "$env:windir\Logs\CBS\CBS.log"
if (Test-Path $cbsLog) {
    $cbsSize = [math]::Round((Get-Item $cbsLog).Length / 1GB, 2)
    if ($cbsSize -gt 0.1) {
        Remove-Item $cbsLog -Force -EA 0
        $SpaceFreed += $cbsSize; $log += "Cleaned: CBS.log - $cbsSize GB"
    }
}

# Windows.old
if (Test-Path "$env:SystemDrive\Windows.old") {
    $size = Get-FolderSize "$env:SystemDrive\Windows.old"
    & takeown /F "$env:SystemDrive\Windows.old" /R /A /D Y 2>&1 | Out-Null
    & icacls "$env:SystemDrive\Windows.old" /grant administrators:F /T /C /Q 2>&1 | Out-Null
    Remove-Item "$env:SystemDrive\Windows.old" -Recurse -Force -EA 0
    $SpaceFreed += $size; $log += "Removed: Windows.old - $size GB"
}

# Upgrade folders
'$Windows.~BT','$Windows.~WS','$WinREAgent' | ForEach-Object {
    $path = "$env:SystemDrive\$_"
    if (Test-Path $path) {
        $size = Get-FolderSize $path
        & takeown /F $path /R /A /D Y 2>&1 | Out-Null
        & icacls $path /grant administrators:F /T /C /Q 2>&1 | Out-Null
        Remove-Item $path -Recurse -Force -EA 0
        if ($size -gt 0) { $SpaceFreed += $size; $log += "Removed: $_ - $size GB" }
    }
}

# Profile Analysis
$users = Get-LocalUser | Where-Object Enabled -eq $true
$profiles = Get-ChildItem "$env:SystemDrive\Users" -Directory -EA 0 | Where-Object { $_.Name -notin @('Public','Default','Default User','All Users') }

$enabled = @()
$users | ForEach-Object {
    $prof = $profiles | Where-Object Name -eq $_.Name
    if ($prof) {
        $ntuser = Join-Path $prof.FullName "NTUSER.DAT"
        $lastLogin = "Never"
        if (Test-Path $ntuser -EA 0) {
            try { $lastLogin = (Get-Item $ntuser -Force -EA 0).LastWriteTime.ToString("yyyy-MM-dd") } catch { $lastLogin = "Unknown" }
        }
        
        $enabled += [PSCustomObject]@{
            Name = $_.Name
            LastLogin = $lastLogin
            SizeGB = Get-FolderSize $prof.FullName
            NoPassword = if (-not $_.PasswordRequired -or $null -eq $_.PasswordLastSet) { "YES" } else { "No" }
        }
    }
}

$orphaned = @()
$profiles | ForEach-Object {
    if (-not ($users | Where-Object Name -eq $_.Name)) {
        $ntuser = Join-Path $_.FullName "NTUSER.DAT"
        $lastLogin = "Unknown"
        if (Test-Path $ntuser -EA 0) {
            try { $lastLogin = (Get-Item $ntuser -Force -EA 0).LastWriteTime.ToString("yyyy-MM-dd") } catch { $lastLogin = "Unknown" }
        }
        
        $orphaned += [PSCustomObject]@{
            Name = $_.Name
            LastLogin = $lastLogin
            SizeGB = Get-FolderSize $_.FullName
        }
    }
}

# Output
Write-Host "`nDISK CLEANUP REPORT"
Write-Host "===================="
Write-Host "Health: $HealthStatus"
Write-Host "Space Freed: $([math]::Round($SpaceFreed, 2)) GB`n"

if ($enabled.Count -gt 0) {
    Write-Host "ENABLED ACCOUNTS:"
    Write-Host "PROFILENAME          LASTLOGIN      SIZE(GB)    NOPASSWORD"
    Write-Host "------------------------------------------------------------"
    $enabled | Sort-Object SizeGB -Descending | ForEach-Object {
        Write-Host ("{0,-20} {1,-14} {2,-11} {3}" -f $_.Name, $_.LastLogin, $_.SizeGB.ToString("0.0"), $_.NoPassword)
    }
    Write-Host ""
}

if ($orphaned.Count -gt 0) {
    Write-Host "ORPHANED PROFILES:"
    Write-Host "PROFILENAME          LASTLOGIN      SIZE(GB)"
    Write-Host "------------------------------------------------"
    $orphaned | Sort-Object SizeGB -Descending | ForEach-Object {
        Write-Host ("{0,-20} {1,-14} {2}" -f $_.Name, $_.LastLogin, $_.SizeGB.ToString("0.0"))
    }
    Write-Host "`nWARNING: Check orphaned profiles for user data before deletion`n"
}

Write-Host "CLEANUP LOG:"
$log | ForEach-Object { Write-Host "- $_" }
