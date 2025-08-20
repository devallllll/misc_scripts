<#
.SYNOPSIS
    Schedules Acrobat 9 removal on shutdown (ONEVENT trigger, no XML) and self-cleans.
.VERSION
    1.4.2
.AUTHOR
    Dave Lane / GoodChoice IT Ltd
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
$taskName     = "RemoveAdobeAcrobat9"

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
# 3) Write the deferred uninstall script (runs at shutdown)
# =========================
$deferred = @"
# Acrobat 9 Dynamic Uninstall + Cleanup (shutdown-safe)
# Version 1.4.2
`$ErrorActionPreference = 'Stop'

`$LogFile = '$logFile'

function Write-Log([string]`$msg, [string]`$level='INFO') {
    `$stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    `$line  = "`$stamp [`$level] `$msg"

    # Write with shared access + retries to avoid lock issues
    for (`$i=1; `$i -le 10; `$i++) {
        try {
            `$fs = [System.IO.File]::Open(`$LogFile,
                [System.IO.FileMode]::Append,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::ReadWrite)
            `$sw = New-Object System.IO.StreamWriter(`$fs)
            `$sw.WriteLine(`$line)
            `$sw.Close()
            `$fs.Close()
            break
        } catch {
            Start-Sleep -Milliseconds 200
        }
    }
}

Write-Log "=== START ==="

# 1) Stop Acrobat-related processes
try {
    Get-Process -Name 'Acrobat','AdobeARM','AcroRd32' -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Write-Log "Stopped Acrobat processes (if present)."
} catch { Write-Log "Process stop warning: `$($_.Exception.Message)" 'WARN' }

# 2) MSI uninstalls for Acrobat 9 GUIDs
try {
    `$keys = Get-ChildItem 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall' -ErrorAction SilentlyContinue |
             Where-Object {
                try {
                    `$p = Get-ItemProperty -Path `$_.PSPath -ErrorAction Stop
                    (`$p.DisplayName -like '*Acrobat*9*') -or (`$_.PSChildName -like '{AC76BA86*}')
                } catch { `$false }
             }
    if (`$keys) {
        `$guids = `$keys | Select-Object -ExpandProperty PSChildName
        foreach (`$guid in `$guids) {
            Write-Log "Uninstalling MSI product `$guid"
            Start-Process msiexec.exe -ArgumentList "/x `$guid /qn /norestart" -Wait -NoNewWindow
        }
    } else {
        Write-Log "No Acrobat 9 MSI entries found."
    }
} catch { Write-Log "MSI uninstall error: `$($_.Exception.Message)" 'ERROR' }

# 3) Run Adobe Cleaner (if present)
try {
    if (Test-Path 'C:\Scripts\AdobeAcroCleaner.exe') {
        Write-Log "Running Acrobat Cleaner."
        Start-Process -FilePath 'C:\Scripts\AdobeAcroCleaner.exe' -ArgumentList '/silent','/product=0','/cleanlevel=1','/scanforothers=1' -Wait -NoNewWindow
        Write-Log "Cleaner completed."
    } else {
        Write-Log "Cleaner missing at 'C:\Scripts\AdobeAcroCleaner.exe'." 'WARN'
    }
} catch { Write-Log "Cleaner error: `$($_.Exception.Message)" 'ERROR' }

# 4) Remove leftover folders
foreach (`$p in @('C:\Program Files (x86)\Adobe\Acrobat 9.0','C:\Program Files\Adobe\Acrobat 9.0')) {
    try {
        if (Test-Path `$p) {
            Write-Log "Removing folder `$p"
            Remove-Item -Path `$p -Recurse -Force -ErrorAction Stop
        }
    } catch { Write-Log "Folder removal warning for `$p : `$($_.Exception.Message)" 'WARN' }
}

# 5) Remove known orphaned uninstall key
try {
    `$orphan = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{AC76BA86-1033-F400-7761-000000000004}'
    if (Test-Path `$orphan) {
        Remove-Item -Path `$orphan -Recurse -Force -ErrorAction Stop
        Write-Log "Removed orphaned uninstall key."
    }
} catch { Write-Log "Registry cleanup warning: `$($_.Exception.Message)" 'WARN' }

# 6) Self-delete task and script
try {
    Write-Log "Deleting scheduled task 'RemoveAdobeAcrobat9'"
    schtasks.exe /Delete /TN "RemoveAdobeAcrobat9" /F | Out-Null
} catch { Write-Log "Task delete warning: `$($_.Exception.Message)" 'WARN' }

try {
    Write-Log "Self-deleting script 'C:\Scripts\RemoveAcrobat9.ps1'"
    Remove-Item -Path 'C:\Scripts\RemoveAcrobat9.ps1' -Force -ErrorAction Stop
} catch { Write-Log "Self-delete warning: `$($_.Exception.Message)" 'WARN' }

Write-Log "=== END ==="
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
# 5) Register the shutdown task (ONEVENT; NO XML)
# =========================
$SchTasks = Join-Path $env:WINDIR 'System32\schtasks.exe'
if (-not (Test-Path $SchTasks)) { throw "schtasks.exe not found at $SchTasks" }

# Build action
$Cmd  = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
$Args = "-NoProfile -ExecutionPolicy Bypass -File `"$taskScript`""

# Remove any existing task quietly
& $SchTasks /Delete /TN $taskName /F 2>$null | Out-Null

# Use cmd.exe to preserve XPath quoting exactly
$XPath1074 = "*[System[Provider[@Name='USER32'] and (EventID=1074)]]"
$createCmd = @"
"$SchTasks" /Create /TN "$taskName" /SC ONEVENT /EC System /MO "$XPath1074" /TR "$Cmd $Args" /RU SYSTEM /RL HIGHEST /F
"@
cmd.exe /c $createCmd

if ($AddEvent6006Fallback) {
    $fallbackName = "${taskName}_6006"
    & $SchTasks /Delete /TN $fallbackName /F 2>$null | Out-Null
    $XPath6006 = "*[System[(EventID=6006)]]"
    $createCmd2 = @"
"$SchTasks" /Create /TN "$fallbackName" /SC ONEVENT /EC System /MO "$XPath6006" /TR "$Cmd $Args" /RU SYSTEM /RL HIGHEST /F
"@
    cmd.exe /c $createCmd2
}

# =========================
# 6) Output a quick sanity summary
# =========================
Write-Host "Task registered as SYSTEM. Details:"
& $SchTasks /Query /TN $taskName /V /FO LIST

Write-Host "`nManual test: schtasks /Run /TN $taskName"
Write-Host "Log (when script runs): $logFile"
