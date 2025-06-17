<#
  .SYNOPSIS
    Runs Dell Command Update to install available updates
  .DESCRIPTION
    Assumes Dell Command Update is already installed and runs it to apply all available Dell updates silently.
  .NOTES
    Author: Aaron J. Stevenson
    Modified to only run updates, not install DCU
#>

function Invoke-DellCommandUpdate {
  # Check for DCU CLI
  $DCU = Get-ChildItem -Path "$env:SystemDrive\Program Files*\Dell\CommandUpdate\dcu-cli.exe" -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
  
  if ($null -eq $DCU) {
    Write-Warning 'Dell Command Update CLI was not detected.'
    Write-Output 'Please install Dell Command Update first.'
    exit 1
  }
  
  Write-Output "Found Dell Command Update at: $DCU"
  
  try {
    Write-Output 'Configuring Dell Command Update...'
    # Configure DCU automatic updates
    $ConfigArgs = '/configure -scheduleAction=DownloadInstallAndNotify -updatesNotification=disable -forceRestart=disable -scheduleAuto -silent'
    Start-Process -NoNewWindow -Wait -FilePath $DCU -ArgumentList $ConfigArgs
    
    Write-Output 'Scanning for and applying Dell updates...'
    # Install updates
    $UpdateArgs = '/applyUpdates -forceUpdate=enable -autoSuspendBitLocker=enable -reboot=disable'
    Start-Process -NoNewWindow -Wait -FilePath $DCU -ArgumentList $UpdateArgs
    
    Write-Output 'Dell updates completed successfully.'
  }
  catch {
    Write-Warning 'Unable to apply updates using the dcu-cli.'
    Write-Warning $_
    exit 1
  }
}

# Set PowerShell preferences
Set-Location -Path $env:SystemRoot
$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'

# Check device manufacturer
$Manufacturer = (Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue).Manufacturer
if ($Manufacturer -notlike '*Dell*') {
  Write-Output "`nNot a Dell system. Manufacturer: $Manufacturer"
  Write-Output "Aborting..."
  exit 0
}

Write-Output "Dell system detected. Manufacturer: $Manufacturer"

# Run Dell Command Update
Invoke-DellCommandUpdate

Write-Output "`nDell Command Update process completed."
