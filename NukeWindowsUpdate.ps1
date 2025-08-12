# GCIT - Schedule One-Time Windows Update Repair + Single Run (CWRMM) v1.1
# - RMM-safe: schedules work, starts it, and exits quickly (no long hold-open)
# - Task runs as SYSTEM, Highest Privilege
# - No PSWindowsUpdate / API usage: UsoClient + built-ins only
# - Safe for WSUS/Intune (doesn't remove WSUS policy)
# - Self-clean: removes task + payload when done
# Exit codes: 0 = scheduled/started OK, 1 = failed to schedule/start

$TaskName   = 'GCIT_WU_Repair_Once'
$WorkDir    = 'C:\ProgramData\GCIT'
$Payload    = Join-Path $WorkDir 'wu-repair-once.ps1'
$LogPath    = 'C:\Windows\Logs\GCIT-WU-Repair-Once.log'

New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null

# --- Payload that the Scheduled Task will run (as SYSTEM) ---
$payloadContent = @'
param([int]$MaxMinutes=45,[switch]$IncludeSFC)

$ErrorActionPreference = 'Stop'
$start = Get-Date
"[$(Get-Date -Format s)] Start WU repair (MaxMinutes=$MaxMinutes, IncludeSFC=$IncludeSFC)" | Out-File -FilePath "$env:SystemRoot\Logs\GCIT-WU-Repair-Once.log" -Append -Encoding utf8

function Write-Log($m){ "[{0}] {1}" -f (Get-Date -Format s),$m | Out-File -Append -FilePath "$env:SystemRoot\Logs\GCIT-WU-Repair-Once.log" -Encoding utf8 }

function Stop-WUServices{ 'wuauserv','bits','cryptsvc','dosvc' | % { sc.exe stop $_ | Out-Null 2>&1 } ; Start-Sleep 3 }
function Start-WUServices{ 'cryptsvc','bits','wuauserv','dosvc' | % { sc.exe start $_ | Out-Null 2>&1 } }

try{
  Write-Log "Stopping services"
  Stop-WUServices

  $sd = Join-Path $env:SystemRoot 'SoftwareDistribution'
  if(Test-Path $sd){ $new="$sd.old_{0:yyyyMMddHHmmss}" -f (Get-Date); Rename-Item $sd $new -Force ; Write-Log "Renamed SoftwareDistribution -> $new" }

  $cr = Join-Path $env:SystemRoot 'System32\catroot2'
  if(Test-Path $cr){ $new="$cr.old_{0:yyyyMMddHHmmss}" -f (Get-Date); Rename-Item $cr $new -Force ; Write-Log "Renamed Catroot2 -> $new" }

  Write-Log "Starting services"
  Start-WUServices

  Write-Log "DISM /RestoreHealth"
  $p = Start-Process dism.exe -ArgumentList '/Online','/Cleanup-Image','/RestoreHealth' -Wait -PassThru -NoNewWindow
  if($p.ExitCode -ne 0){ throw "DISM exit code $($p.ExitCode)" }

  if($IncludeSFC){
    Write-Log "SFC /scannow"
    $p = Start-Process sfc.exe -ArgumentList '/scannow' -Wait -PassThru -NoNewWindow
    if($p.ExitCode -gt 1){ throw "SFC exit code $($p.ExitCode)" }
  }

# (A) After stopping services, clear BITS queue (modern path)
$bitsDir = Join-Path $env:ALLUSERSPROFILE 'Microsoft\Network\Downloader'
if (Test-Path $bitsDir) {
  Get-ChildItem -Path $bitsDir -Filter 'qmgr*.dat' -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue
  Log "BITS queue (qmgr*.dat) cleared"
}

# (B) Optional: Winsock reset (only if you suspect socket stack issues)
# NOTE: This REQUIRES a reboot to fully take effect; schedule outside business hours.
$DoWinsockReset = $false  # set to $false  don’t want it by default
if ($DoWinsockReset) {
  Log "netsh winsock reset"
  Start-Process -FilePath netsh.exe -ArgumentList 'winsock','reset' -Wait -NoNewWindow
  # Leave reboot to your normal policy; the RebootRequired flags will likely be set.
}


  Write-Log "Kick one Windows Update cycle (scan/download/install)"
  usoClient StartScan     | Out-Null 2>&1
  Start-Sleep 5
  usoClient StartDownload | Out-Null 2>&1
  Start-Sleep 5
  usoClient StartInstall  | Out-Null 2>&1

  # Timebox (let installs run a bit but not forever). Adjust as needed.
  $limit = (Get-Date).AddMinutes($MaxMinutes)
  while((Get-Date) -lt $limit){
    Start-Sleep -Seconds 30
    # exit early if a reboot is queued; let RMM handle it later
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'){
      Write-Log "RebootRequired detected; leaving for RMM."
      break
    }
  }
  Write-Log "WU cycle finished or timeboxed."

} catch {
  Write-Log "ERROR: $($_.Exception.Message)"
} finally {
  Write-Log ("Completed in {0} minutes." -f ([int](New-TimeSpan -Start $start -End (Get-Date)).TotalMinutes))
  # Self-clean: remove scheduled task and payload
Get-ChildItem "$env:SystemRoot\SoftwareDistribution.old_*","$env:SystemRoot\System32\catroot2.old_*" -Directory -ErrorAction SilentlyContinue  |
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

  
  try{
    schtasks /Delete /TN "GCIT_WU_Repair_Once" /F | Out-Null 2>&1
    Write-Log "Scheduled task removed."
  } catch { Write-Log "Could not remove task: $($_.Exception.Message)" }
  try{
    Remove-Item -Path $PSCommandPath -Force
  } catch { Write-Log "Could not remove payload: $($_.Exception.Message)" }
}
'@

Set-Content -Path $Payload -Value $payloadContent -Encoding UTF8 -Force

# Build a start time a minute from now (local)
$startTime = (Get-Date).AddMinutes(1).ToString('HH:mm')
$startDate = (Get-Date).ToString('dd/MM/yyyy')

# Create the scheduled task
$create = schtasks /Create /TN $TaskName /RU "SYSTEM" /RL HIGHEST /SC ONCE /SD $startDate /ST $startTime /F /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$Payload`" -MaxMinutes 45"
if($LASTEXITCODE -ne 0){ Write-Host "Failed to create task ($LASTEXITCODE)"; exit 1 }

# Start it immediately so we don't wait for the clock tick
schtasks /Run /TN $TaskName | Out-Null 2>&1

Write-Host "Scheduled task '$TaskName' created and started. Payload: $Payload  Log: $LogPath"
exit 0
