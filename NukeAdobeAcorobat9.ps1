<#
.SYNOPSIS
    Schedules dynamic Acrobat 9 removal on shutdown using a self-deleting task.

.VERSION
    1.4
.AUTHOR
    Dave Lane / GoodChoice IT Ltd
#>

# === CONFIG ===
$scriptFolder = "C:\Scripts"
$taskScript  = "$scriptFolder\RemoveAcrobat9.ps1"
$cleanerUrl  = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/2100120135/x64/AdobeAcroCleaner_DC2021.exe"
$cleanerExe  = "$scriptFolder\AdobeAcroCleaner.exe"
$taskName    = "RemoveAdobeAcrobat9"

# === Create folder ===
if (-not (Test-Path $scriptFolder)) {
  New-Item -ItemType Directory -Path $scriptFolder | Out-Null
}

# === Write the deferred uninstall script ===
@"
# Acrobat 9 Dynamic Uninstall + Cleanup (shutdown-safe)

# Find Acrobat 9 GUIDs from registry (32-bit uninstall hive on x64)
`$keys = Get-ChildItem 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall' -ErrorAction SilentlyContinue |
Where-Object {
  try {
    `$p = Get-ItemProperty \$_ -ErrorAction Stop
    (`$p.DisplayName -like '*Acrobat*9*') -or (\$_.'PSChildName' -like '{AC76BA86*}')
  } catch { `$false }
}

`$guids = `$keys | Select-Object -ExpandProperty PSChildName -ErrorAction SilentlyContinue

foreach (`$guid in `$guids) {
  Start-Process msiexec.exe -ArgumentList "/x `$guid /qn /norestart" -Wait -NoNewWindow
}

# Run Acrobat Cleaner
Start-Process -FilePath `"$cleanerExe`" -ArgumentList "/silent","/product=0","/cleanlevel=1","/scanforothers=1" -Wait

# Kill Acrobat processes before removing folder
Get-Process -Name "Acrobat","AdobeARM","AcroRd32" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

# Remove folder
`$installPath = "C:\Program Files (x86)\Adobe\Acrobat 9.0"
if (Test-Path `$installPath) {
  try {
    Remove-Item -Path `$installPath -Recurse -Force -ErrorAction Stop
    Write-Host "Removed leftover Acrobat 9.0 folder."
  } catch {
    Write-Warning "Failed to remove Acrobat folder: `$($_.Exception.Message)"
  }
}

# Remove orphaned Pro Extended uninstall key
`$orphan = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{AC76BA86-1033-F400-7761-000000000004}"
if (Test-Path `$orphan) {
  Remove-Item -Path `$orphan -Recurse -Force -ErrorAction SilentlyContinue
}

# Self-delete task and script
schtasks /Delete /TN "$taskName" /F
Remove-Item -Path `"$taskScript`" -Force
"@ | Set-Content -Encoding UTF8 -Path $taskScript

# === Download Cleaner Tool (TLS 1.2 + robust fallback) ===
try {
  # Ensure TLS 1.2 is enabled for this process (older OS defaults)
  [Net.ServicePointManager]::SecurityProtocol = `
      [Net.SecurityProtocolType]::Tls12 -bor `
      [Net.SecurityProtocolType]::Tls11 -bor `
      [Net.SecurityProtocolType]::Tls

  $headers = @{ "User-Agent" = "Mozilla/5.0 (Windows NT; PowerShell)" }
  Invoke-WebRequest -Uri $cleanerUrl -OutFile $cleanerExe -Headers $headers -UseBasicParsing -ErrorAction Stop
}
catch {
  Write-Warning "Invoke-WebRequest failed (`$($_.Exception.Message)`), trying BITS…"
  try {
    Start-BitsTransfer -Source $cleanerUrl -Destination $cleanerExe -ErrorAction Stop
  } catch {
    throw "Failed to download AdobeAcroCleaner via both IWR and BITS: `$($_.Exception.Message)`"
  }
}

# === Register Shutdown Task (safe XML build + UTF-16) ===
# Build XML as a single-quoted here-string (no expansion/HTML-encoding), then inject $taskScript with -f
$taskXml = @'
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <Triggers>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>
        <QueryList>
          <Query Id="0" Path="System">
            <Select Path="System">
              *[System[(EventID=1074)]]
            </Select>
          </Query>
        </QueryList>
      </Subscription>
    </EventTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <RunLevel>HighestAvailable</RunLevel>
      <UserId>S-1-5-18</UserId>
      <LogonType>ServiceAccount</LogonType>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <AllowStartOnDemand>false</AllowStartOnDemand>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT30M</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe</Command>
      <Arguments>-NoProfile -ExecutionPolicy Bypass -File "{0}"</Arguments>
    </Exec>
  </Actions>
</Task>
'@ -f $taskScript

# Write and register the task (Task Scheduler expects UTF-16LE)
$taskXmlPath = "$env:TEMP\RemoveAdobeTask.xml"
$taskXml | Out-File -Encoding Unicode -FilePath $taskXmlPath
schtasks.exe /Create /TN $taskName /XML $taskXmlPath /F
Remove-Item -Path $taskXmlPath -Force -ErrorAction SilentlyContinue
