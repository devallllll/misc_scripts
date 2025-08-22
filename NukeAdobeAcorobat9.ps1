<#
.SYNOPSIS
    Schedule Acrobat 9 removal at STARTUP and SHUTDOWN (no XML), re-run until gone. Adds MSI logging and kills AcroTray.
.VERSION
    1.5.0
.AUTHOR
    Dave Lane / GoodChoice IT Ltd (with assist)
#>

# =========================
# 0) Prefer 64-bit host
# =========================
if (-not [Environment]::Is64BitProcess) {
    $sysNativePS = "$env:WINDIR\SysNative\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path $sysNativePS) {
        & $sysNativePS -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath @args
        exit $LASTEXITCODE
    }
}

# =========================
# 1) Config
# =========================
$scriptFolder = "C:\Scripts"
$taskScript   = Join-Path $scriptFolder 'RemoveAcrobat9.ps1'

# Two tasks: one for startup, one for shutdown
$taskNameStart    = "RemoveAdobeAcrobat9_Start"
$taskNameShutdown = "RemoveAdobeAcrobat9_Shutdown"

$workRoot     = "C:\ProgramData\GCIT"
$logFolder    = Join-Path $workRoot "Logs"
$logFile      = Join-Path $logFolder "RemoveAcrobat9.log"

$cleanerUrl   = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/2100120135/x64/AdobeAcroCleaner_DC2021.exe"
$cleanerExe   = Join-Path $scriptFolder 'AdobeAcroCleaner.exe'

# Optional: also create a fallback trigger on EventID 6006 (Event Log service stopped)
$AddEvent6006Fallback = $false

# =========================
# 2) Ensure folders
# =========================
foreach ($p in @($scriptFolder, $workRoot, $logFolder)) {
    if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p | Out-Null }
}

# =========================
# 3) Write the deferred uninstall script (runs at startup/shutdown)
# =========================
$deferred = @"
# Acrobat 9 Dynamic Uninstall + Cleanup (startup/shutdown-safe)
# Version 1.5.0
`$ErrorActionPreference = 'Stop'

`$LogFolder = '$logFolder'
`$LogFile   = '$logFile'
`$MsiLog    = Join-Path `$LogFolder 'Acrobat9-uninstall.msi.log'

# Task names to remove when done
`$TaskNames = @('$taskNameStart', '$taskNameShutdown')

function Write-Log([string]`$msg, [string]`$level='INFO') {
    try { if (-not (Test-Path `$LogFolder)) { New-Item -ItemType Directory -Path `$LogFolder -Force | Out-Null } } catch {}
    `$stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    `$line  = "`$stamp [`$level] `$msg"
    # Shared write with retries
    for (`$i=1; `$i -le 10; `$i++) {
        try {
            `$fs = [System.IO.File]::Open(`$LogFile,
                [System.IO.FileMode]::Append,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::ReadWrite)
            `$sw = New-Object System.IO.StreamWriter(`$fs)
            `$sw.WriteLine(`$line)
            `$sw.Close(); `$fs.Close(); break
        } catch { Start-Sleep -Milliseconds 200 }
    }
}

function Get-Acrobat9Targets {
    # Returns a hashtable: @{ ProductCodes = @(...); UninstallStrings = @(...) }
    `$targets = @{
        ProductCodes    = New-Object System.Collections.Generic.List[string]
        UninstallStrings= New-Object System.Collections.Generic.List[string]
    }

    # 1) Adobe Installer keys (often contain ProductCode)
    `$installerKeys = @(
        'HKLM:\SOFTWARE\WOW6432Node\Adobe\Adobe Acrobat\9.0\Installer',
        'HKLM:\SOFTWARE\Adobe\Adobe Acrobat\9.0\Installer'
    )
    foreach (`$k in `$installerKeys) {
        try {
            if (Test-Path `$k) {
                `$p = Get-ItemProperty -Path `$k -ErrorAction Stop
                `$pc = `$p.ProductCode
                if (`$pc -and (`$pc -match '^\{[0-9A-F-]+\}$')) {
                    if (-not `$targets.ProductCodes.Contains(`$pc)) { `$targets.ProductCodes.Add(`$pc) }
                }
            }
        } catch { Write-Log "Installer key read warning: `$($_.Exception.Message)" 'WARN' }
    }

    # 2) Uninstall keys (exclude Reader; match Acrobat 9 by name or GUID pattern)
    `$uninstRoots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    `$guidPattern = '^\{AC76BA86-[0-9A-F]{4}-[0-9A-F]{4}-(7760|7761)-[0-9A-F]{12}\}$' # common Acrobat 9 pattern

    foreach (`$root in `$uninstRoots) {
        try {
            Get-ChildItem `$root -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    `$p = Get-ItemProperty -Path `$_.PSPath -ErrorAction Stop
                    `$name = `$p.DisplayName
                    if (-not `$name) { return }
                    # Must contain "Adobe Acrobat" and version 9, and not "Reader"
                    if ((`$name -match 'Adobe\s+Acrobat\b') -and (`$name -notmatch 'Reader') -and (`$name -match '\b9(\.|\b)')) {
                        # Prefer ProductCode if key is a GUID
                        if (`$_.PSChildName -match '^\{[0-9A-F-]+\}$') {
                            if (-not `$targets.ProductCodes.Contains(`$_.PSChildName)) { `$targets.ProductCodes.Add(`$_.PSChildName) }
                        }
                        # Capture UninstallString as fallback
                        if (`$p.UninstallString) {
                            if (-not `$targets.UninstallStrings.Contains(`$p.UninstallString)) { `$targets.UninstallStrings.Add(`$p.UninstallString) }
                        }
                    } elseif (`$_.PSChildName -match `$guidPattern) {
                        if (-not `$targets.ProductCodes.Contains(`$_.PSChildName)) { `$targets.ProductCodes.Add(`$_.PSChildName) }
                    }
                } catch { }
            }
        } catch { Write-Log "Uninstall hive read warning: `$($_.Exception.Message)" 'WARN' }
    }

    return `$targets
}

function Test-Acrobat9Present {
    `$t = Get-Acrobat9Targets
    return (`$t.ProductCodes.Count -gt 0 -or `$t.UninstallStrings.Count -gt 0 -or
            (Test-Path 'C:\Program Files (x86)\Adobe\Acrobat 9.0') -or
            (Test-Path 'C:\Program Files\Adobe\Acrobat 9.0'))
}

function Remove-Tasks-And-Self {
    foreach (`$tn in `$TaskNames) {
        try { schtasks.exe /Delete /TN "`$tn" /F | Out-Null } catch { Write-Log "Task delete warning for `$tn: `$($_.Exception.Message)" 'WARN' }
    }
    try {
        Write-Log "Self-deleting script '$taskScript'"
        Remove-Item -Path '$taskScript' -Force -ErrorAction Stop
    } catch { Write-Log "Self-delete warning: `$($_.Exception.Message)" 'WARN' }
}

Write-Log "=== START === (triggered at `$((Get-Date).ToString('HH:mm:ss'))) "

# 0) If Acrobat 9 is already gone, tidy up and exit
if (-not (Test-Acrobat9Present)) {
    Write-Log "Acrobat 9 not detected. Cleaning tasks and exiting."
    Remove-Tasks-And-Self
    Write-Log "=== END (nothing to do) ==="
    exit 0
}

# 1) Stop Acrobat-related processes (incl. AcroTray)
try {
    Get-Process -Name 'Acrobat','AdobeARM','AcroRd32','AcroTray' -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Write-Log "Stopped Acrobat/Tray/ARM processes (if present)."
} catch { Write-Log "Process stop warning: `$($_.Exception.Message)" 'WARN' }

# 2) MSI uninstalls for discovered Acrobat 9 products (with logging)
`$targets = Get-Acrobat9Targets
try {
    foreach (`$guid in (`$targets.ProductCodes | Select-Object -Unique)) {
        Write-Log "Uninstalling MSI product `$guid"
        Start-Process msiexec.exe -ArgumentList "/x `$guid /qn REBOOT=ReallySuppress /L*v `"`$MsiLog`"" -Wait -NoNewWindow
    }
} catch { Write-Log "MSI uninstall error: `$($_.Exception.Message)" 'ERROR' }

# 3) If only UninstallString(s) exist, normalize and run with quiet+log
try {
    foreach (`$cmd in (`$targets.UninstallStrings | Select-Object -Unique)) {
        # Normalize /I to /X and ensure quiet+log
        `$norm = `$cmd -replace '(/I|/i)\b','/X'
        if (`$norm -notmatch '/X') { `$norm = `$norm + ' /X' }
        if (`$norm -notmatch '/qn') { `$norm = `$norm + ' /qn' }
        if (`$norm -notmatch 'REBOOT=ReallySuppress') { `$norm = `$norm + ' REBOOT=ReallySuppress' }
        if (`$norm -notmatch '/L\*v') { `$norm = `$norm + " /L*v `"`$MsiLog`"" }
        Write-Log "Uninstall via UninstallString: `$norm"
        Start-Process cmd.exe -ArgumentList "/c `$norm" -Wait -NoNewWindow
    }
} catch { Write-Log "UninstallString error: `$($_.Exception.Message)" 'ERROR' }

# 4) Run Adobe Cleaner (post-uninstall sweep)
try {
    if (Test-Path 'C:\Scripts\AdobeAcroCleaner.exe') {
        Write-Log "Running Acrobat Cleaner."
        Start-Process -FilePath 'C:\Scripts\AdobeAcroCleaner.exe' -ArgumentList '/silent','/product=0','/cleanlevel=1','/scanforothers=1' -Wait -NoNewWindow
        Write-Log "Cleaner completed."
    } else {
        Write-Log "Cleaner missing at 'C:\Scripts\AdobeAcroCleaner.exe'." 'WARN'
    }
} catch { Write-Log "Cleaner error: `$($_.Exception.Message)" 'ERROR' }

# 5) Remove leftover folders (best-effort)
foreach (`$p in @('C:\Program Files (x86)\Adobe\Acrobat 9.0','C:\Program Files\Adobe\Acrobat 9.0')) {
    try {
        if (Test-Path `$p) {
            Write-Log "Removing folder `$p"
            Remove-Item -Path `$p -Recurse -Force -ErrorAction Stop
        }
    } catch { Write-Log "Folder removal warning for `$p : `$($_.Exception.Message)" 'WARN' }
}

# 6) Remove orphaned uninstall entries matching Acrobat 9 pattern (only after uninstall attempt)
try {
    `$uninstRoots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    `$pattern = '^\{AC76BA86-[0-9A-F]{4}-[0-9A-F]{4}-(7760|7761)-[0-9A-F]{12}\}$'
    foreach (`$root in `$uninstRoots) {
        Get-ChildItem `$root -ErrorAction SilentlyContinue | Where-Object {
            `$_.PSChildName -match `$pattern -and
            ((Get-ItemProperty `$_.PSPath -ea SilentlyContinue).DisplayName -match 'Adobe\s+Acrobat\b.*\b9(\.|\b)')
        } | ForEach-Object {
            try {
                Remove-Item -Path `$_.PsPath -Recurse -Force -ErrorAction Stop
                Write-Log "Removed orphan uninstall key: `$($_.PSChildName)"
            } catch { Write-Log "Orphan key removal warning: `$($_.Exception.Message)" 'WARN' }
        }
    }
} catch { Write-Log "Registry cleanup warning: `$($_.Exception.Message)" 'WARN' }

# 7) Decide whether to keep rescheduling or self-clean
if (-not (Test-Acrobat9Present)) {
    Write-Log "Acrobat 9 no longer detected. Removing tasks and self-deleting."
    Remove-Tasks-And-Self
    Write-Log "=== END (removed) ==="
    exit 0
} else {
    Write-Log "Acrobat 9 still detected; will run again on next startup/shutdown."
    Write-Log "=== END (retry pending) ==="
    exit 0
}
"@

$deferred | Set-Content -Path $taskScript -Encoding UTF8 -Force

# =========================
# 4) Download Cleaner (BITS first, fallback to IWR)
# =========================
$downloaded = $false
try {
    Start-BitsTransfer -Source $cleanerUrl -Destination $cleanerExe -TransferType Download -ErrorAction Stop
    if ((Test-Path $cleanerExe) -and ((Get-Item $cleanerExe).Length -gt 0)) { $downloaded = $true }
} catch { }

if (-not $downloaded) {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
    try {
        Invoke-WebRequest -Uri $cleanerUrl -OutFile $cleanerExe -UseBasicParsing -TimeoutSec 120
        if ((Test-Path $cleanerExe) -and ((Get-Item $cleanerExe).Length -gt 0)) { $downloaded = $true }
    } catch { }
}

if (-not $downloaded) {
    Write-Host "WARN: Cleaner download failed via BITS and IWR; continuing (script can still run without cleaner)."
}

# =========================
# 5) Register the startup & shutdown tasks (NO XML)
# =========================
$SchTasks = Join-Path $env:WINDIR 'System32\schtasks.exe'
if (-not (Test-Path $SchTasks)) { throw "schtasks.exe not found at $SchTasks" }

# Build action
$Cmd  = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
$Args = "-NoProfile -ExecutionPolicy Bypass -File `"$taskScript`""

# Remove any existing tasks quietly
& $SchTasks /Delete /TN $taskNameStart /F 2>$null | Out-Null
& $SchTasks /Delete /TN $taskNameShutdown /F 2>$null | Out-Null
if ($AddEvent6006Fallback) {
    & $SchTasks /Delete /TN "${taskNameShutdown}_6006" /F 2>$null | Out-Null
}

# Startup task: delay to let services settle
& $SchTasks /Create /TN $taskNameStart /SC ONSTART /DELAY 0001:30 /TR "$Cmd $Args" /RU SYSTEM /RL HIGHEST /F

# Shutdown task: use cmd.exe to preserve XPath quoting
$XPath1074 = "*[System[Provider[@Name='USER32'] and (EventID=1074)]]"
$createShutdown = @"
"$SchTasks" /Create /TN "$taskNameShutdown" /SC ONEVENT /EC System /MO "$XPath1074" /TR "$Cmd $Args" /RU SYSTEM /RL HIGHEST /F
"@
cmd.exe /c $createShutdown

# Optional 6006 fallback (Event Log service stopped)
if ($AddEvent6006Fallback) {
    $XPath6006 = "*[System[(EventID=6006)]]"
    $create6006 = @"
"$SchTasks" /Create /TN "${taskNameShutdown}_6006" /SC ONEVENT /EC System /MO "$XPath6006" /TR "$Cmd $Args" /RU SYSTEM /RL HIGHEST /F
"@
    cmd.exe /c $create6006
}

# =========================
# 6) Output a quick sanity summary
# =========================
Write-Host "Tasks registered as SYSTEM. Details:"
& $SchTasks /Query /TN $taskNameStart /V /FO LIST
& $SchTasks /Query /TN $taskNameShutdown /V /FO LIST

Write-Host "`nManual tests:"
Write-Host "  schtasks /Run /TN $taskNameStart"
Write-Host "  schtasks /Run /TN $taskNameShutdown"
Write-Host "Log (when the deferred script runs): $logFile"
Write-Host "MSI log: $logFolder\Acrobat9-uninstall.msi.log"
