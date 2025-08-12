# GCIT - Schedule One-Time Windows Update Repair + MU Bypass (CWRMM) v1.4
# - Schedules a SYSTEM task, starts it, then exits quickly (avoids RMM timeouts)
# - Task flow:
#     stop services -> clear BITS qmgr -> rename SoftwareDistribution/Catroot2
#     -> TEMPORARY WSUS BYPASS + enable Microsoft Update (drivers/Office)
#     -> start services -> DISM (optional SFC)
#     -> one WU scan/download/install -> self-clean
# - No PSWindowsUpdate usage (USOClient + built-ins only)
# - WSUS bypass affects THIS RUN ONLY; GPO will restore WSUS later
# - Cleanup: deletes any *.old_* caches immediately after run
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
$logPath = "$env:SystemRoot\Logs\GCIT-WU-Repair-Once.log"
$start = Get-Date
"[{0}] Start WU repair (MaxMinutes={1}, IncludeSFC={2})" -f (Get-Date -Format s),$MaxMinutes,$IncludeSFC | Out-File -FilePath $logPath -Append -Encoding utf8

function Write-Log($m){ "[{0}] {1}" -f (Get-Date -Format s),$m | Out-File -Append -FilePath $logPath -Encoding utf8 }
function Stop-WUServices{ 'wuauserv','bits','cryptsvc','dosvc' | % { sc.exe stop $_ | Out-Null 2>&1 } ; Start-Sleep 3 }
function Start-WUServices{ 'cryptsvc','bits','wuauserv','dosvc' | % { sc.exe start $_ | Out-Null 2>&1 } }

try{
  Write-Log "Stopping services"
  Stop-WUServices

  # Clear BITS queue (with services stopped)
  $bitsDir = Join-Path $env:ALLUSERSPROFILE 'Microsoft\Network\Downloader'
  if (Test-Path $bitsDir) {
    Get-ChildItem -Path $bitsDir -Filter 'qmgr*.dat' -ErrorAction SilentlyContinue |
      Remove-Item -Force -ErrorAction SilentlyContinue
    Write-Log "BITS queue (qmgr*.dat) cleared"
  }

  # Rename caches
  $sd = Join-Path $env:SystemRoot 'SoftwareDistribution'
  if(Test-Path $sd){
    $new="$sd.old_{0:yyyyMMddHHmmss}" -f (Get-Date)
    Rename-Item $sd $new -Force
    Write-Log "Renamed SoftwareDistribution -> $new"
  }
  $cr = Join-Path $env:SystemRoot 'System32\catroot2'
  if(Test-Path $cr){
    $new="$cr.old_{0:yyyyMMddHHmmss}" -f (Get-Date)
    Rename-Item $cr $new -Force
    Write-Log "Renamed Catroot2 -> $new"
  }

  # --- Temporary WSUS + Microsoft Update bypass (this run only) ---
  try {
    $WU = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $AU = Join-Path $WU 'AU'

    # Optional: backup current WSUS policy branch for inspection
    $Backup = Join-Path $env:ProgramData 'GCIT-WSUS-Backup.reg'
    reg.exe export "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" $Backup /y | Out-Null 2>&1
    Write-Log "WSUS policy backed up to $Backup"

    New-Item -Path $WU -Force | Out-Null
    New-Item -Path $AU -Force | Out-Null

    # Disable WSUS for this session and clear WSUS URLs
    Set-ItemProperty -Path $AU -Name UseWUServer -Type DWord -Value 0 -Force
    Remove-ItemProperty -Path $WU -Name WUServer -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path $WU -Name WUStatusServer -ErrorAction SilentlyContinue

    # Enable Microsoft Update service (includes drivers/Office if present)
    New-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services' -Force | Out-Null
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Services' `
                     -Name DefaultService `
                     -Value '7971f918-a847-4430-9279-4a52d1efe18d' -Force

    Write-Log "Full WSUS bypass applied and Microsoft Update enabled for this run"
  } catch {
    Write-Log "Failed to apply full WSUS/MU bypass: $($_.Exception.Message)"
  }

  Write-Log "Starting services"
  Start-WUServices

  # Ensure client picks up policy change before scan
  try {
    Write-Log "Refreshing WU client settings"
    UsoClient.exe RefreshSettings | Out-Null 2>&1
  } catch { }

  Write-Log "DISM /RestoreHealth"
  $p = Start-Process dism.exe -ArgumentList '/Online','/Cleanup-Image','/RestoreHealth' -Wait -PassThru -NoNewWindow
  if($p.ExitCode -ne 0){ throw "DISM exit code $($p.ExitCode)" }

  if($IncludeSFC){
    Write-Log "SFC /scannow"
    $p = Start-Process sfc.exe -ArgumentList '/scannow' -Wait -PassThru -NoNewWindow
    if($p.ExitCode -gt 1){ throw "SFC exit code $($p.ExitCode)" }
  }

  Write-Log "Kick one Windows Update cycle (scan/download/install) via Microsoft Update"
  UsoClient StartScan     | Out-Null 2>&1
  Start-Sleep 5
  UsoClient StartDownload | Out-Null 2>&1
  Start-Sleep 5
  UsoClient StartInstall  | Out-Null 2>&1

  # Timebox / early-exit on reboot flag
  $limit = (Get-Date).AddMinutes($MaxMinutes)
  while((Get-Date) -lt $limit){
    Start-Sleep -Seconds 30
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'){
      Write-Log "RebootRequired detected; leaving for RMM."
      break
    }
  }
  Write-Log "WU cycle finished or timeboxed."
}
catch {
  Write-Log "ERROR: $($_.Exception.Message)"
}
finally {
  Write-Log ("Completed in {0} minutes." -f ([int](New-TimeSpan -Start $start -End (Get-Date)).TotalMinutes))

  # Cleanup: delete renamed caches immediately (no retention)
  Get-ChildItem @("$env:SystemRoot\SoftwareDistribution.old_*","$env:SystemRoot\System32\catroot2.old_*") -Directory -ErrorAction SilentlyContinue |
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

# Create the scheduled task (SYSTEM, highest privs)
schtasks /Create /TN $TaskName /RU "SYSTEM" /RL HIGHEST /SC ONCE /SD $startDate /ST $startTime /F /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$Payload`" -MaxMinutes 45"
if($LASTEXITCODE -ne 0){ Write-Host "Failed to create task ($LASTEXITCODE)"; exit 1 }

# Start it immediately so we don't wait for the clock tick
schtasks /Run /TN $TaskName | Out-Null 2>&1

Write-Host "Scheduled task '$TaskName' created and started. Payload: $Payload  Log: $LogPath"
exit 0
