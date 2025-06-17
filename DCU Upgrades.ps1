<#
  .SYNOPSIS
    Runs Dell Command Update to install available updates
  .DESCRIPTION
    Removes SupportAssist bloatware, then runs Dell Command Update to apply all available Dell updates silently.
  .NOTES
    Author: Aaron J. Stevenson
    Modified to remove SupportAssist and run updates only
#>

function Get-InstalledApps {
  param(
    [Parameter(Mandatory)][String[]]$DisplayNames,
    [String[]]$Exclude
  )
  
  $RegPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
  )
  
  # Get applications matching criteria
  $BroadMatch = @()
  foreach ($DisplayName in $DisplayNames) {
    $AppsWithBundledVersion = Get-ChildItem -Path $RegPaths | Get-ItemProperty | Where-Object { $_.DisplayName -like "*$DisplayName*" -and $null -ne $_.BundleVersion }
    if ($AppsWithBundledVersion) { $BroadMatch += $AppsWithBundledVersion }
    else { $BroadMatch += Get-ChildItem -Path $RegPaths | Get-ItemProperty | Where-Object { $_.DisplayName -like "*$DisplayName*" } }
  }
  
  # Remove excluded apps
  $MatchedApps = @()
  foreach ($App in $BroadMatch) {
    if ($Exclude -notcontains $App.DisplayName) { $MatchedApps += $App }
  }

  return $MatchedApps | Sort-Object { [version]$_.BundleVersion } -Descending
}

function Remove-DellSupportAssist {
  # Stop SupportAssist processes first (not services)
  $ProcessesToStop = @(
    'SupportAssistClientUI',
    'SupportAssistAgent',
    'DellClientManagementService'
  )
  
  foreach ($ProcessName in $ProcessesToStop) {
    $Process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($Process) {
      Write-Output "Stopping process: $ProcessName"
      try {
        $Process | Stop-Process -Force -ErrorAction Stop
        Write-Output "Successfully stopped $ProcessName"
      }
      catch {
        Write-Warning "Failed to stop process $ProcessName: $_"
      }
    }
  }
  
  # Find SupportAssist using registry method (more reliable)
  $SupportAssistApps = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' | 
    Where-Object { $_.DisplayName -like '*SupportAssist*' }
  
  if ($SupportAssistApps) { 
    Write-Output 'Dell SupportAssist bloatware detected - removing...' 
  }
  
  foreach ($App in $SupportAssistApps) {
    Write-Output "Attempting to remove [$($App.DisplayName)] version $($App.DisplayVersion)"
    try {
      if ($App.UninstallString -match 'msiexec') {
        $Guid = [regex]::Match($App.UninstallString, '\{[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}\}').Value
        Start-Process -NoNewWindow -Wait -FilePath 'msiexec.exe' -ArgumentList "/x $Guid /quiet /qn"
      }
      else { Start-Process -NoNewWindow -Wait -FilePath $App.UninstallString -ArgumentList '/quiet' }
      Write-Output "Successfully removed $($App.DisplayName)"
    }
    catch { 
      Write-Warning "Failed to remove $($App.DisplayName): $_"
      # Don't exit - continue with updates even if SupportAssist removal fails
    }
  }
}

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
    # Install updates (without forcing or BitLocker suspension)
    $UpdateArgs = '/applyUpdates -reboot=disable'
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

# Remove SupportAssist bloatware first
Remove-DellSupportAssist

# Run Dell Command Update
Invoke-DellCommandUpdate

Write-Output "`nDell Command Update process completed."
