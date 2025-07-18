<#
.SYNOPSIS
    Silently uninstalls Adobe Acrobat 9 (Standard, Pro, and Pro Extended) and runs Adobe's Cleaner Tool to remove remnants.

.DESCRIPTION
    This script:
      - Uninstalls Acrobat 9 using known MSI product GUIDs
      - Downloads and executes the AdobeAcroCleaner_DC2021 tool silently
      - Removes any orphaned registry entries (Add/Remove Programs)
      - Deletes temporary files

.VERSION
    1.0
.AUTHOR
    Dave Lane, GoodChoice IT Ltd
.COPYRIGHT
    © 2025 GoodChoice IT Ltd. All rights reserved.
    For internal use only. UnTested on Windows 10/11.

.NOTES
    Must be run with administrative privileges.
#>


# Setup
$workingDir = "$env:TEMP\AdobeCleanup"
$cleanerUrl = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/2100120135/x64/AdobeAcroCleaner_DC2021.exe"
$cleanerExe = "$workingDir\AdobeAcroCleaner.exe"

# Acrobat 9 Product GUIDs (Standard, Pro, Pro Extended)
$acrobatGUIDs = @(
    "{AC76BA86-1033-F400-7760-000000000004}",
    "{AC76BA86-1033-0000-BA7E-000000000004}",
    "{AC76BA86-1033-F400-7761-000000000004}"
)

# Create working directory
if (-not (Test-Path $workingDir)) {
    New-Item -ItemType Directory -Path $workingDir | Out-Null
}

# Download AcroCleaner
Invoke-WebRequest -Uri $cleanerUrl -OutFile $cleanerExe

# Uninstall all known Acrobat 9 versions
foreach ($guid in $acrobatGUIDs) {
    Start-Process "msiexec.exe" -ArgumentList "/x $guid /qn /norestart" -Wait -NoNewWindow
}

# Run the AcroCleaner tool silently
Start-Process -FilePath $cleanerExe -ArgumentList "/silent", "/product=0", "/cleanlevel=1", "/scanforothers=1" -Wait -NoNewWindow

# Remove orphaned Add/Remove Programs entry for Pro Extended (if exists)
$arpKey = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{AC76BA86-1033-F400-7761-000000000004}"
if (Test-Path $arpKey) {
    Remove-Item -Path $arpKey -Recurse -Force -ErrorAction SilentlyContinue
}

# Final cleanup
Remove-Item -Path $workingDir -Recurse -Force -ErrorAction SilentlyContinue
