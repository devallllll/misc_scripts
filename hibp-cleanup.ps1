Write-Host "Starting cleanup..." -ForegroundColor Cyan

# Paths
$auditPath = "C:\AD-Audit"
$hibpHash = Join-Path $auditPath "pwnedpasswords_ntlm.txt"

Write-Host "Audit path: $auditPath" -ForegroundColor Cyan

# 1. Unload modules
Write-Host "Unloading DSInternals and AD modules..." -ForegroundColor Yellow
Remove-Module DSInternals -ErrorAction SilentlyContinue
Remove-Module ActiveDirectory -ErrorAction SilentlyContinue

# 2. Uninstall DSInternals
Write-Host "Uninstalling DSInternals PowerShell module..." -ForegroundColor Yellow
Get-InstalledModule DSInternals -ErrorAction SilentlyContinue | Uninstall-Module -Force

# 3. Uninstall the HIBP downloader
Write-Host "Uninstalling HaveIBeenPwned NTLM downloader..." -ForegroundColor Yellow
dotnet tool uninstall --global haveibeenpwned-downloader

# 4. Remove only AD database copies (NOT the HIBP hashes!)
Write-Host "Removing AD database copies (ntds.dit + SYSTEM)..." -ForegroundColor Yellow

$filesToRemove = @(
    "ntds.dit",
    "SYSTEM"
)

foreach ($file in $filesToRemove) {
    $fullPath = Join-Path $auditPath $file
    if (Test-Path $fullPath) {
        Remove-Item $fullPath -Force
        Write-Host "Deleted: $fullPath" -ForegroundColor Green
    }
}

# 5. Keep HIBP hash file but clean any leftover temp files
Write-Host "Preserving HIBP NTLM hash file: $hibpHash" -ForegroundColor Green

# Remove stray files except logs and hashes
Write-Host "Cleaning temp files but keeping CSV logs and HIBP dataset..." -ForegroundColor Yellow

Get-ChildItem $auditPath | Where-Object {
    # Keep hash + CSV reports
    $_.Name -ne "pwnedpasswords_ntlm.txt" -and
    $_.Extension -ne ".csv"
} | ForEach-Object {
    if ($_.PSIsContainer -eq $false -and $_.Name -notlike "*log*") {
        Remove-Item $_.FullName -Force
        Write-Host "Removed temp file: $($_.Name)" -ForegroundColor DarkGray
    }
}

# 6. Optional: Clear PowerShell history
Write-Host "Clearing PowerShell history..." -ForegroundColor Yellow
Clear-History

Write-Host "Cleanup complete. AD database and tools removed, HIBP hashes + logs retained." -ForegroundColor Green
