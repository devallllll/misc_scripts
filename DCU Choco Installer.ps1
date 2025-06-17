<#
  .SYNOPSIS
    Removes old Dell Command Update and installs latest via Chocolatey
  .DESCRIPTION
    Removes incompatible/old DCU versions and installs the latest via Chocolatey. Does not run updates.
  .NOTES
    Author: Aaron J. Stevenson
    Modified to only install DCU via Chocolatey, not run updates
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

function Test-ChocolateyInstalled {
  return (Test-Path -Path "$env:ProgramData\chocolatey\choco.exe")
}

function Install-Chocolatey {
  Write-Output 'Installing Chocolatey...'
  try {
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    return $true
  }
  catch {
    Write-Warning "Error installing Chocolatey: $($_.Exception.Message)"
    return $false
  }
}

function Remove-IncompatibleApps {
  # Check for incompatible products (including old DCU versions that need upgrading)
  $IncompatibleApps = Get-InstalledApps -DisplayNames 'Dell Update', 'Dell Command | Update' `
    -Exclude 'Dell SupportAssist OS Recovery Plugin for Dell Update'
  
  # Filter out current versions we want to keep (5.0+)
  $AppsToRemove = @()
  foreach ($App in $IncompatibleApps) {
    if ($App.DisplayName -like '*Dell Command | Update*') {
      # Check version - remove if less than 5.0.0
      if ($App.DisplayVersion -and [version]$App.DisplayVersion -lt [version]'5.0.0') {
        $AppsToRemove += $App
      }
    } else {
      # Remove all other Dell Update variants
      $AppsToRemove += $App
    }
  }
  
  if ($AppsToRemove) { Write-Output 'Incompatible or outdated Dell applications detected' }
  foreach ($App in $AppsToRemove) {
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
      Write-Warning "Failed to remove $($App.DisplayName)"
      Write-Warning $_
      exit 1
    }
  }
}

function Install-DellCommandUpdateViaChocolatey {
  if (-not (Test-ChocolateyInstalled)) {
    Write-Output 'Chocolatey not found, installing...'
    if (-not (Install-Chocolatey)) {
      Write-Warning 'Failed to install Chocolatey'
      exit 1
    }
  }

  Write-Output 'Installing/updating Dell Command Update via Chocolatey...'
  try {
    $ChocoProcess = Start-Process -FilePath 'choco' -ArgumentList 'install dellcommandupdate -y' -Wait -NoNewWindow -PassThru
    Write-Output "Chocolatey process completed with exit code: $($ChocoProcess.ExitCode)"
    
    if ($ChocoProcess.ExitCode -eq 0) {
      Write-Output 'Successfully installed/updated Dell Command Update via Chocolatey'
      
      # Verify installation
      Start-Sleep -Seconds 5
      $DCU = Get-ChildItem -Path "$env:SystemDrive\Program Files*\Dell\CommandUpdate\dcu-cli.exe" -ErrorAction SilentlyContinue
      if ($DCU) {
        Write-Output "Dell Command Update CLI found at: $($DCU.FullName)"
      } else {
        Write-Warning 'Dell Command Update CLI not found after installation'
      }
    } else {
      Write-Warning "Chocolatey failed with exit code: $($ChocoProcess.ExitCode)"
      exit 1
    }
  }
  catch {
    Write-Warning "Chocolatey installation failed: $($_.Exception.Message)"
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

# Remove old/incompatible DCU versions
Remove-IncompatibleApps

# Install latest DCU via Chocolatey
Install-DellCommandUpdateViaChocolatey

Write-Output "`nDell Command Update installation completed. Ready for manual updates."
