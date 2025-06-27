<#
.SYNOPSIS
   Service Management Script for Clearing RMM Alerts
   
.DESCRIPTION
   Temporarily enables/starts then stops/disables Windows services to clear RMM monitoring alerts.
   
.PARAMETER ServiceName
   Windows service name to manage. Defaults to "RemoteAccess".
#>

param(
   [string]$ServiceName = "RemoteAccess"
)

try {
   Set-Service -Name $ServiceName -StartupType Automatic
   Start-Service -Name $ServiceName
   Start-Sleep -Seconds 10
   
   Stop-Service -Name $ServiceName -Force
   Start-Sleep -Seconds 10
   
   Set-Service -Name $ServiceName -StartupType Disabled
   
   $service = Get-Service -Name $ServiceName
   Write-Host "Final Status: $($service.Status), Startup: $((Get-WmiObject -Class Win32_Service -Filter "Name='$ServiceName'").StartMode)"
   
} catch {
   Write-Error "Error: $($_.Exception.Message)"
}
