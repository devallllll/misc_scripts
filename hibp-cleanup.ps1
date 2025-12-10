param(
    [ValidateSet("Prep", "Audit", "Cleanup")]
    [string]$Action
)

$ErrorActionPreference = "Stop"

# Core paths
$auditPath  = "C:\AD-Audit"
$ntdsPath   = Join-Path $auditPath "ntds.dit"
$systemPath = Join-Path $auditPath "SYSTEM"
$hibpPath   = Join-Path $auditPath "pwnedpasswords_ntlm.txt"

function Ensure-AuditFolder {
    if (-not (Test-Path $auditPath)) {
        Write-Host "Creating audit folder at $auditPath" -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $auditPath -Force | Out-Null
    } else {
        Write-Host "Audit folder already exists at $auditPath" -ForegroundColor Green
    }
    Set-Location $auditPath
}

function Do-Prep {
    Write-Host "=== PREP MODE ===" -ForegroundColor Cyan
    Ensure-AuditFolder

    # Disk space check
    $drive = Get-PSDrive -Name C
    $freeGB = [math]::Round($drive.Free / 1GB, 2)
    Write-Host "Free space on C: is $freeGB GB" -ForegroundColor Cyan
    if ($freeGB -lt 80) {
        Write-Warning "Less than 80 GB free on C:. The HIBP NTLM file is ~45–50 GB."
    }

    # DSInternals check/install
    Write-Host "Checking for DSInternals module..." -ForegroundColor Yellow
    try {
        $dsModule = Get-Module -ListAvailable -Name DSInternals -ErrorAction SilentlyContinue
        if (-not $dsModule) {
            Write-Host "DSInternals not found. Attempting to install from PowerShell Gallery..." -ForegroundColor Yellow
            Install-Module DSInternals -Scope AllUsers -Force
        } else {
            Write-Host "DSInternals is already installed." -ForegroundColor Green
        }
        Import-Module DSInternals
        Write-Host "DSInternals module loaded." -ForegroundColor Green
    } catch {
        Write-Warning "❌ DSInternals failed to install or import."
        Write-Host "➡ Manually run:  Install-Module DSInternals -Scope AllUsers" -ForegroundColor Cyan
        Write-Host "➡ Ensure this server has internet access and PowerShellGet configured." -ForegroundColor Cyan
    }

    # ActiveDirectory module
    Write-Host "Checking for ActiveDirectory module..." -ForegroundColor Yellow
    try {
        $adModule = Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue
        if (-not $adModule) {
            Write-Warning "ActiveDirectory module not found."
            Write-Host "➡ Install RSAT / AD tools on this machine (Server: Add-WindowsFeature RSAT-AD-PowerShell, Client: RSAT)." -ForegroundColor Cyan
        } else {
            Import-Module ActiveDirectory
            Write-Host "ActiveDirectory module loaded." -ForegroundColor Green
        }
    } catch {
        Write-Warning "Failed to load ActiveDirectory module. Install RSAT/AD tools manually."
    }

    # Check for NTDS + SYSTEM
    $ntdsExists = Test-Path $ntdsPath
    $sysExists  = Test-Path $systemPath

    if ($ntdsExists -and $sysExists) {
        Write-Host "✔ NTDS + SYSTEM files found in $auditPath" -ForegroundColor Green
    }
    else {
        Write-Warning "❌ NTDS and/or SYSTEM hive not found in $auditPath."
        Write-Host ""
        Write-Host "To run the audit, you MUST provide an offline copy of:" -ForegroundColor Yellow
        Write-Host "   • ntds.dit" -ForegroundColor White
        Write-Host "   • SYSTEM hive" -ForegroundColor White
        Write-Host ""
        Write-Host "Recommended ways to obtain them:" -ForegroundColor Cyan
        Write-Host "  1) Restore from a DC *backup* (preferred)." -ForegroundColor White
        Write-Host "  2) Restore via a VSS snapshot of a DC." -ForegroundColor White
        Write-Host "  3) Use DSInternals offline backup cmdlet on a DC:" -ForegroundColor White
        Write-Host ""
        Write-Host "       Create-NtdsOfflineBackup -Destination $auditPath" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Would you like to attempt an automated offline NTDS backup using DSInternals *now*?" -ForegroundColor Cyan
        Write-Host "(Only do this if you're running as Domain Admin on a domain controller.)" -ForegroundColor Cyan
        $choice = Read-Host "Enter Y to run Create-NtdsOfflineBackup, any other key to skip"

        if ($choice -match "^[Yy]") {
            try {
                Write-Host "Running Create-NtdsOfflineBackup..." -ForegroundColor Yellow
                Create-NtdsOfflineBackup -Destination $auditPath
                Write-Host "✔ Offline NTDS backup created (check for ntds.dit and SYSTEM in $auditPath)." -ForegroundColor Green
            } catch {
                Write-Warning "❌ Failed to create offline backup."
                Write-Host "➡ Ensure you are running as Domain Admin *on a domain controller*." -ForegroundColor Cyan
            }
        } else {
            Write-Host "Skipping automatic NTDS backup. Copy ntds.dit and SYSTEM into $auditPath manually later." -ForegroundColor Yellow
        }
    }

    # HIBP downloader tool
    Write-Host "Checking for haveibeenpwned-downloader (.NET global tool)..." -ForegroundColor Yellow
    $hibpToolInstalled = $false
    try {
        $toolList = dotnet tool list -g 2>$null
        if ($toolList -match "haveibeenpwned-downloader") {
            $hibpToolInstalled = $true
        }
    } catch {
        Write-Warning "❌ Unable to query dotnet tools. Ensure .NET SDK is installed and 'dotnet' is in PATH."
        Write-Host "➡ Install .NET SDK manually and re-run Prep." -ForegroundColor Cyan
    }

    if (-not $hibpToolInstalled) {
        try {
            Write-Host "Installing haveibeenpwned-downloader..." -ForegroundColor Yellow
            dotnet tool install --global haveibeenpwned-downloader
            Write-Host "✔ haveibeenpwned-downloader installed." -ForegroundColor Green
        } catch {
            Write-Warning "❌ Failed to install haveibeenpwned-downloader."
            Write-Host "➡ Manually run: dotnet tool install --global haveibeenpwned-downloader" -ForegroundColor Cyan
            Write-Host "➡ Ensure .NET SDK is installed and accessible." -ForegroundColor Cyan
        }
    } else {
        Write-Host "✔ haveibeenpwned-downloader is already installed." -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "Prep complete. Next steps:" -ForegroundColor Green
    Write-Host "  1) Ensure NTDS + SYSTEM exist in $auditPath (either from backup or Create-NtdsOfflineBackup)." -ForegroundColor White
    Write-Host "  2) From $auditPath, download HIBP NTLM hashes:" -ForegroundColor White
    Write-Host "       haveibeenpwned-downloader.exe -n pwnedpasswords_ntlm" -ForegroundColor Yellow
    Write-Host "  3) Then run this script with -Action Audit" -ForegroundColor White
}

function Do-Audit {
    Write-Host "=== AUDIT MODE ===" -ForegroundColor Cyan
    Ensure-AuditFolder

    # Check required files
    if (-not (Test-Path $ntdsPath)) {
        Write-Error "ntds.dit not found at $ntdsPath"
        Write-Host "➡ Copy ntds.dit from a DC backup or run Create-NtdsOfflineBackup into $auditPath, then re-run Audit." -ForegroundColor Cyan
        return
    }
    if (-not (Test-Path $systemPath)) {
        Write-Error "SYSTEM hive not found at $systemPath"
        Write-Host "➡ Copy SYSTEM from the same backup as ntds.dit into $auditPath, then re-run Audit." -ForegroundColor Cyan
        return
    }
    if (-not (Test-Path $hibpPath)) {
        Write-Error "HIBP NTLM file not found at $hibpPath"
        Write-Host "➡ From $auditPath, run: haveibeenpwned-downloader.exe -n pwnedpasswords_ntlm" -ForegroundColor Cyan
        return
    }

    # Load modules
    try { Import-Module DSInternals } catch { Write-Error "DSInternals not available. Run Prep first."; return }
    try { Import-Module ActiveDirectory } catch { Write-Warning "ActiveDirectory module not available. Email/UPN enrichment may fail."; }

    $timestamp  = Get-Date -Format "yyyyMMdd-HHmmss"
    $rawReport  = Join-Path $auditPath "PasswordQualityReport-$timestamp.csv"
    $hibpReport = Join-Path $auditPath "HIBP-Compromised-Users-$timestamp.csv"

    Write-Host "Running Get-ADDBAccount | Test-PasswordQuality..." -ForegroundColor Yellow
    Write-Host "This may take a while on large domains." -ForegroundColor Yellow

    $results = Get-ADDBAccount -All `
        -NtdsPath $ntdsPath `
        -SystemHivePath $systemPath |
        Test-PasswordQuality `
            -WeakPasswordHashesFile $hibpPath `
            -IncludeDisabledAccounts:$false `
            -IncludeServiceAccounts:$false

    Write-Host "Total accounts analysed: $($results.Count)" -ForegroundColor Cyan

    # Save full DSInternals output (raw)
    $results | Export-Csv $rawReport -NoTypeInformation -Encoding UTF8
    Write-Host "Full password quality report saved to: $rawReport" -ForegroundColor Green

    # Only accounts whose password hash appears in HIBP (WeakPassword = True)
    $weak = $results | Where-Object { $_.WeakPassword -eq $true }
    Write-Host "Accounts with HIBP-compromised passwords: $($weak.Count)" -ForegroundColor Yellow

    # Enrich with AD info
    $final = foreach ($u in $weak) {
        $ad = $null
        try {
            $ad = Get-ADUser -Identity $u.SamAccountName -Properties mail, userPrincipalName, displayName -ErrorAction Stop
        } catch {
            # ignore lookup failures and still output basic info
        }

        [PSCustomObject]@{
            SamAccountName     = $u.SamAccountName
            DisplayName        = $ad.DisplayName
            UserPrincipalName  = $ad.UserPrincipalName
            Email              = $ad.mail
            DistinguishedName  = $u.DistinguishedName
            WeakPassword       = $u.WeakPassword
        }
    }

    $final | Export-Csv $hibpReport -NoTypeInformation -Encoding UTF8
    Write-Host "HIBP-compromised user report saved to: $hibpReport" -ForegroundColor Green
    Write-Host "Use this CSV to drive password resets / user communication before enabling writeback/CIPP onboarding." -ForegroundColor White
}

function Do-Cleanup {
    Write-Host "=== CLEANUP MODE ===" -ForegroundColor Cyan
    Ensure-AuditFolder

    # Unload modules
    Write-Host "Unloading DSInternals and ActiveDirectory modules..." -ForegroundColor Yellow
    Remove-Module DSInternals -ErrorAction SilentlyContinue
    Remove-Module ActiveDirectory -ErrorAction SilentlyContinue

    # Uninstall DSInternals
    Write-Host "Uninstalling DSInternals module (if installed)..." -ForegroundColor Yellow
    try {
        Get-InstalledModule DSInternals -ErrorAction SilentlyContinue | Uninstall-Module -Force
    } catch {
        Write-Warning "DSInternals uninstall failed or it was not installed. You can manually run: Uninstall-Module DSInternals"
    }

    # Uninstall HIBP downloader
    Write-Host "Uninstalling haveibeenpwned-downloader (.NET global tool)..." -ForegroundColor Yellow
    try {
        dotnet tool uninstall --global haveibeenpwned-downloader
    } catch {
        Write-Warning "Failed to uninstall haveibeenpwned-downloader. You can manually run: dotnet tool uninstall --global haveibeenpwned-downloader"
    }

    # Remove AD DB copies (keep HIBP + CSV reports)
    Write-Host "Removing AD database copies (ntds.dit + SYSTEM) from $auditPath..." -ForegroundColor Yellow
    foreach ($file in @("ntds.dit", "SYSTEM")) {
        $full = Join-Path $auditPath $file
        if (Test-Path $full) {
            Remove-Item $full -Force
            Write-Host "Deleted: $full" -ForegroundColor Green
        }
    }

    # Clean temp files but preserve HIBP dataset and CSV reports
    Write-Host "Cleaning temp files but preserving pwnedpasswords_ntlm.txt and CSV reports..." -ForegroundColor Yellow

    Get-ChildItem $auditPath | Where-Object {
        $_.Name -ne "pwnedpasswords_ntlm.txt" -and
        $_.Extension -ne ".csv"
    } | ForEach-Object {
        if (-not $_.PSIsContainer) {
            Remove-Item $_.FullName -Force
            Write-Host "Removed temp file: $($_.Name)" -ForegroundColor DarkGray
        }
    }

    # Optional: clear PowerShell history
    Write-Host "Clearing PowerShell history..." -ForegroundColor Yellow
    try { Clear-History } catch { }

    Write-Host "Cleanup complete. AD copies and tools removed, HIBP hashes and reports retained." -ForegroundColor Green
}

# -------- Main entry --------
if (-not $Action) {
    Write-Host "Select Action: [1] Prep  [2] Audit  [3] Cleanup" -ForegroundColor Cyan
    $choice = Read-Host "Enter 1, 2, or 3"
    switch ($choice) {
        "1" { $Action = "Prep" }
        "2" { $Action = "Audit" }
        "3" { $Action = "Cleanup" }
        default {
            Write-Error "Invalid choice. Use -Action Prep | Audit | Cleanup"
            exit 1
        }
    }
}

switch ($Action) {
    "Prep"    { Do-Prep }
    "Audit"   { Do-Audit }
    "Cleanup" { Do-Cleanup }
}
