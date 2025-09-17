#Requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$Uninstall,
    [switch]$Reconfigure,
    [switch]$ForceNewCert,
    [string]$TaskName = "PowerShell AutoSigner",
    [string]$WorkingDir = "C:\ProgramData\PS-AutoSigner"
)

$ErrorActionPreference = 'Stop'

if ($Uninstall) {
    Write-Host "Uninstalling PowerShell AutoSigner..." -ForegroundColor Yellow
    
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($task) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Removed scheduled task: $TaskName"
    }
    
    if (Test-Path $WorkingDir) {
        Remove-Item $WorkingDir -Recurse -Force
        Write-Host "Removed directory: $WorkingDir"
    }
    
    Write-Host "Uninstall complete!" -ForegroundColor Green
    return
}

Write-Host "Setting up PowerShell AutoSigner..." -ForegroundColor Cyan

if (!(Test-Path $WorkingDir)) {
    New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
    Write-Host "Created working directory: $WorkingDir"
}

$ConfigFile = Join-Path $WorkingDir "config.json"
$LogFile = Join-Path $WorkingDir "autosigner.log"
$WorkerScript = Join-Path $WorkingDir "AutoSigner-Worker.ps1"
$VersionFile = Join-Path $WorkingDir "file-versions.json"

$config = @{}

if ((Test-Path $ConfigFile) -and !$Reconfigure) {
    try {
        $configData = Get-Content $ConfigFile -Raw | ConvertFrom-Json
        $config.ToSignFolder = $configData.ToSignFolder
        $config.SignedFolder = $configData.SignedFolder
        Write-Host "Loaded existing configuration"
    } catch {
        Write-Warning "Config file corrupted, will reconfigure"
        $config = @{}
    }
}

if (!$config.ToSignFolder -or !$config.SignedFolder -or $Reconfigure) {
    Write-Host "`nFolder Configuration:" -ForegroundColor Yellow
    
    $defaultToSign = Join-Path $env:USERPROFILE "Documents\PowerShell\ToSign"
    $defaultSigned = Join-Path $env:USERPROFILE "Documents\PowerShell\Signed"
    
    do {
        $toSignInput = Read-Host "ToSign folder [$defaultToSign]"
        $config.ToSignFolder = if ($toSignInput) { $toSignInput } else { $defaultToSign }
    } while (!(Test-Path $config.ToSignFolder -IsValid))
    
    do {
        $signedInput = Read-Host "Signed folder [$defaultSigned]"  
        $config.SignedFolder = if ($signedInput) { $signedInput } else { $defaultSigned }
    } while (!(Test-Path $config.SignedFolder -IsValid))
}

foreach ($folder in @($config.ToSignFolder, $config.SignedFolder)) {
    if (!(Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Write-Host "Created folder: $folder"
    }
}

$config.LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$config | ConvertTo-Json | Set-Content $ConfigFile -Encoding UTF8
Write-Host "Configuration saved"

Write-Host "`nSetting up code signing certificate..." -ForegroundColor Cyan

function Get-CodeSigningCert {
    param([switch]$ForceNew)
    
    if (!$ForceNew) {
        $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object {
            $_.HasPrivateKey -and 
            $_.NotAfter -gt (Get-Date).AddDays(30) -and
            ($_.EnhancedKeyUsageList | Where-Object { $_.ObjectId -eq "1.3.6.1.5.5.7.3.3" })
        } | Sort-Object NotAfter -Descending | Select-Object -First 1
        
        if ($cert) {
            Write-Host "Using existing certificate: $($cert.Subject)"
            return $cert
        }
    }
    
    Write-Host "Creating new code signing certificate..."
    
    # Create certificate with explicit code signing EKU
    $cert = New-SelfSignedCertificate `
        -Subject "CN=PowerShell AutoSigner" `
        -Type CodeSigningCert `
        -KeyAlgorithm RSA `
        -KeyLength 2048 `
        -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" `
        -KeyExportPolicy NonExportable `
        -KeyUsage DigitalSignature `
        -KeyUsageProperty Sign `
        -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3") `
        -CertStoreLocation Cert:\CurrentUser\My `
        -NotAfter (Get-Date).AddYears(5)
    
    Write-Host "Created new certificate: $($cert.Thumbprint)"
    
    # Verify the certificate is suitable for code signing
    $hasCodeSigning = $cert.EnhancedKeyUsageList | Where-Object { $_.ObjectId -eq "1.3.6.1.5.5.7.3.3" }
    if (!$hasCodeSigning) {
        Write-Warning "Certificate may not be properly configured for code signing"
    } else {
        Write-Host "Certificate verified for code signing"
    }
    
    return $cert
}

$signingCert = Get-CodeSigningCert -ForceNew:$ForceNewCert

# Update config with certificate thumbprint
$config.CertThumbprint = $signingCert.Thumbprint
$config.LastUpdated = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Save updated config
$config | ConvertTo-Json | Set-Content $ConfigFile -Encoding UTF8
Write-Host "Configuration updated with certificate thumbprint"

# Add certificate to Trusted Root store for self-signed certificates
try {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
    $store.Open("ReadWrite")
    
    # Check if certificate is already in the Root store
    $existingCert = $store.Certificates | Where-Object { $_.Thumbprint -eq $signingCert.Thumbprint }
    if (!$existingCert) {
        $store.Add($signingCert)
        Write-Host "Added certificate to Trusted Root Certification Authorities"
    } else {
        Write-Host "Certificate already in Trusted Root store"
    }
    $store.Close()
} catch {
    Write-Warning "Could not add certificate to Trusted Root store: $($_.Exception.Message)"
    Write-Host "You may need to manually trust the certificate for code signing to work"
}

$publicCertPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "PowerShell-AutoSigner.cer"
Export-Certificate -Cert $signingCert -FilePath $publicCertPath -Force | Out-Null
Write-Host "Public certificate exported to desktop"

Write-Host "`nCreating worker script..." -ForegroundColor Cyan

$workerContent = '#Requires -Version 5.1
param(
    [Parameter(Mandatory)][string]$ConfigFile
)

# Set error handling and create log immediately
$ErrorActionPreference = "Continue"
$LogFile = Join-Path (Split-Path $ConfigFile) "autosigner.log"

function Write-AutoSignLog {
    param([string]$Message, [string]$Level = "INFO")
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp] [$Level] $Message"
        Add-Content $LogFile $logEntry -Encoding UTF8
        Write-Host $logEntry
    } catch {
        # Fallback if logging fails
        Write-Host "[$Level] $Message"
    }
}

Write-AutoSignLog "=== AutoSigner Worker Started ==="
Write-AutoSignLog "ConfigFile: $ConfigFile"
Write-AutoSignLog "LogFile: $LogFile"

# Test if config file exists
if (!(Test-Path $ConfigFile)) {
    Write-AutoSignLog "ERROR: Config file not found: $ConfigFile" "ERROR"
    exit 1
}

# Load configuration with better error handling
try {
    Write-AutoSignLog "Loading configuration..."
    $configContent = Get-Content $ConfigFile -Raw
    Write-AutoSignLog "Config content loaded, length: $($configContent.Length)"
    
    $configData = $configContent | ConvertFrom-Json
    Write-AutoSignLog "JSON parsed successfully"
    
    $config = @{
        ToSignFolder = $configData.ToSignFolder
        SignedFolder = $configData.SignedFolder
        CertThumbprint = $configData.CertThumbprint
    }
    
    Write-AutoSignLog "ToSignFolder: $($config.ToSignFolder)"
    Write-AutoSignLog "SignedFolder: $($config.SignedFolder)"
    Write-AutoSignLog "CertThumbprint: $($config.CertThumbprint)"
    
} catch {
    Write-AutoSignLog "ERROR loading config: $($_.Exception.Message)" "ERROR"
    Write-AutoSignLog "Config file contents: $(Get-Content $ConfigFile -Raw)" "ERROR"
    exit 1
}

# Test folders exist
if (!(Test-Path $config.ToSignFolder)) {
    Write-AutoSignLog "ERROR: ToSign folder not found: $($config.ToSignFolder)" "ERROR"
    exit 1
}

if (!(Test-Path $config.SignedFolder)) {
    Write-AutoSignLog "ERROR: Signed folder not found: $($config.SignedFolder)" "ERROR"
    exit 1
}

# Test certificate - handle empty thumbprint
if ([string]::IsNullOrWhiteSpace($config.CertThumbprint)) {
    Write-AutoSignLog "ERROR: Certificate thumbprint is empty in config. Please run installer again." "ERROR"
    exit 1
}

$certs = Get-ChildItem "Cert:\CurrentUser\My\$($config.CertThumbprint)" -ErrorAction SilentlyContinue
if (!$certs -or $certs.Count -eq 0) {
    Write-AutoSignLog "ERROR: Certificate not found: $($config.CertThumbprint)" "ERROR"
    Write-AutoSignLog "Available certificates:" "ERROR"
    Get-ChildItem "Cert:\CurrentUser\My" | ForEach-Object { 
        Write-AutoSignLog "  $($_.Thumbprint) - $($_.Subject)" "ERROR"
    }
    exit 1
}

# Take the first certificate if multiple found
$cert = $certs[0]
Write-AutoSignLog "Certificate found: $($cert.Subject) (Thumbprint: $($cert.Thumbprint))"
if ($certs.Count -gt 1) {
    Write-AutoSignLog "WARNING: Multiple certificates found with same thumbprint, using first one" "WARN"
}

# Debug certificate properties
Write-AutoSignLog "Certificate debugging:"
Write-AutoSignLog "  HasPrivateKey: $($cert.HasPrivateKey)"
Write-AutoSignLog "  NotAfter: $($cert.NotAfter)"
$keyUsageExt = $cert.Extensions | Where-Object {$_.Oid.FriendlyName -eq "Key Usage"}
Write-AutoSignLog "  KeyUsage: $keyUsageExt"
Write-AutoSignLog "  Enhanced Key Usage count: $($cert.EnhancedKeyUsageList.Count)"

foreach ($eku in $cert.EnhancedKeyUsageList) {
    Write-AutoSignLog "    EKU: $($eku.FriendlyName) - OID: $($eku.ObjectId)"
}

# Try to test the certificate with a simple signing operation first
try {
    Write-AutoSignLog "Testing certificate with temporary file..."
    $tempFile = [System.IO.Path]::GetTempFileName() + ".ps1"
    "# Test file" | Set-Content $tempFile
    $testResult = Set-AuthenticodeSignature -FilePath $tempFile -Certificate $cert
    Write-AutoSignLog "Test signing result: $($testResult.Status)"
    Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
    
    if ($testResult.Status -ne "Valid") {
        Write-AutoSignLog "Certificate test failed. Looking for alternative certificates..." "WARN"
        
        # Find any code signing certificate
        $altCerts = Get-ChildItem "Cert:\CurrentUser\My" | Where-Object {
            $_.HasPrivateKey -and 
            $_.NotAfter -gt (Get-Date) -and
            ($_.EnhancedKeyUsageList | Where-Object { $_.ObjectId -eq "1.3.6.1.5.5.7.3.3" -or $_.FriendlyName -eq "Code Signing" })
        }
        
        if ($altCerts) {
            $cert = $altCerts[0]
            Write-AutoSignLog "Using alternative certificate: $($cert.Subject) (Thumbprint: $($cert.Thumbprint))"
        } else {
            Write-AutoSignLog "No suitable code signing certificates found" "ERROR"
            Write-AutoSignLog "Available certificates with Enhanced Key Usage:" "ERROR"
            Get-ChildItem "Cert:\CurrentUser\My" | Where-Object { $_.EnhancedKeyUsageList.Count -gt 0 } | ForEach-Object {
                Write-AutoSignLog "  $($_.Thumbprint) - $($_.Subject)" "ERROR"
                foreach ($eku in $_.EnhancedKeyUsageList) {
                    Write-AutoSignLog "    $($eku.FriendlyName) - $($eku.ObjectId)" "ERROR"
                }
            }
            exit 1
        }
    }
} catch {
    Write-AutoSignLog "Certificate test failed: $($_.Exception.Message)" "ERROR"
    exit 1
}

# Look for PowerShell files
$VersionFile = Join-Path (Split-Path $ConfigFile) "file-versions.json"
Write-AutoSignLog "VersionFile: $VersionFile"

Write-AutoSignLog "Scanning for PowerShell files in: $($config.ToSignFolder)"
$psFiles = Get-ChildItem $config.ToSignFolder -Filter "*.ps*1" -Recurse -File -ErrorAction SilentlyContinue
Write-AutoSignLog "Found $($psFiles.Count) PowerShell files"

if ($psFiles.Count -eq 0) {
    Write-AutoSignLog "No PowerShell files found to process"
    Write-AutoSignLog "=== AutoSigner Worker Completed (No Files) ==="
    exit 0
}

# Load file version tracking
$fileVersions = @{}
if (Test-Path $VersionFile) {
    try {
        $versionData = Get-Content $VersionFile -Raw | ConvertFrom-Json
        $versionData.PSObject.Properties | ForEach-Object {
            $fileVersions[$_.Name] = @{
                Version = $_.Value.Version
                Hash = $_.Value.Hash
                LastSigned = $_.Value.LastSigned
            }
        }
        Write-AutoSignLog "Loaded version data for $($fileVersions.Count) files"
    } catch {
        Write-AutoSignLog "Version file corrupted, starting fresh: $($_.Exception.Message)" "WARN"
    }
}

$processedCount = 0

foreach ($file in $psFiles) {
    try {
        Write-AutoSignLog "Processing: $($file.Name)"
        
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $fileContent = Get-Content $file.FullName -Raw -ErrorAction Stop
        
        # Calculate content hash (ignore version comments)
        $cleanContent = $fileContent -replace "(?m)^#\s*Version:.*$", "" -replace "(?m)^#\s*Signed:.*$", ""
        $contentBytes = [System.Text.Encoding]::UTF8.GetBytes($cleanContent)
        $hashAlgorithm = [System.Security.Cryptography.SHA256]::Create()
        $hashBytes = $hashAlgorithm.ComputeHash($contentBytes)
        $contentHash = [System.BitConverter]::ToString($hashBytes) -replace "-", ""
        
        Write-AutoSignLog "File hash: $contentHash"
        
        # Check if file changed
        $currentVersion = "1.0.0"
        if ($fileVersions[$fileName]) {
            $currentVersion = $fileVersions[$fileName].Version
            if ($fileVersions[$fileName].Hash -eq $contentHash) {
                Write-AutoSignLog "No changes detected for: $($file.Name)"
                continue
            }
        }
        
        Write-AutoSignLog "Current version: $currentVersion"
        
        # Increment version
        $versionParts = $currentVersion.Split(".")
        $major = [int]$versionParts[0]
        $minor = [int]$versionParts[1] 
        $patch = [int]$versionParts[2]
        
        # Check for version bump indicators in comments
        if ($fileContent -match "(?m)^#.*bump:\s*major") {
            $major++; $minor = 0; $patch = 0
            Write-AutoSignLog "Major version bump detected"
        } elseif ($fileContent -match "(?m)^#.*bump:\s*minor") {
            $minor++; $patch = 0
            Write-AutoSignLog "Minor version bump detected"
        } else {
            $patch++
            Write-AutoSignLog "Patch version bump"
        }
        
        $newVersion = "$major.$minor.$patch"
        Write-AutoSignLog "New version: $newVersion"
        
        # Add/update version header
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $versionHeader = "# Version: $newVersion`r`n# Signed: $timestamp"
        
        # Remove old version headers and add new ones
        $updatedContent = $fileContent -replace "(?m)^#\s*Version:.*$", "" -replace "(?m)^#\s*Signed:.*$", "" -replace "(?m)^#\s*bump:\s*(major|minor).*$", ""
        $updatedContent = $versionHeader + "`r`n" + $updatedContent.TrimStart()
        
        # Write updated content
        Set-Content $file.FullName $updatedContent -Encoding UTF8
        Write-AutoSignLog "Updated file content with version header"
        
        # Sign the file
        Write-AutoSignLog "Signing file..."
        $signResult = Set-AuthenticodeSignature -FilePath $file.FullName -Certificate $cert -TimestampServer "http://timestamp.sectigo.com"
        if ($signResult.Status -ne "Valid") {
            Write-AutoSignLog "Failed to sign: $($file.Name) - Status: $($signResult.Status) - Message: $($signResult.StatusMessage)" "ERROR"
            continue
        }
        Write-AutoSignLog "File signed successfully"
        
        # Move to signed folder with version suffix
        $signedFileName = "$fileName" + "_v$newVersion" + $file.Extension
        $signedPath = Join-Path $config.SignedFolder $signedFileName
        
        # Handle duplicate names
        $counter = 1
        while (Test-Path $signedPath) {
            $signedFileName = "$fileName" + "_v$newVersion" + "_$counter" + $file.Extension
            $signedPath = Join-Path $config.SignedFolder $signedFileName
            $counter++
        }
        
        Write-AutoSignLog "Moving to: $signedPath"
        Move-Item $file.FullName $signedPath -Force
        
        # Update version tracking
        $fileVersions[$fileName] = @{
            Version = $newVersion
            Hash = $contentHash
            LastSigned = $timestamp
        }
        
        Write-AutoSignLog "SUCCESS: $($file.Name) -> $signedFileName (v$newVersion)"
        $processedCount++
        
    } catch {
        Write-AutoSignLog "ERROR processing $($file.Name): $($_.Exception.Message)" "ERROR"
        Write-AutoSignLog "Stack trace: $($_.ScriptStackTrace)" "ERROR"
    }
}

# Save version tracking
try {
    $fileVersions | ConvertTo-Json -Depth 3 | Set-Content $VersionFile -Encoding UTF8
    Write-AutoSignLog "Version tracking saved"
} catch {
    Write-AutoSignLog "ERROR saving version tracking: $($_.Exception.Message)" "ERROR"
}

Write-AutoSignLog "=== AutoSigner Worker Completed - Processed $processedCount files ==="'

Set-Content $WorkerScript $workerContent -Encoding UTF8
Write-Host "Worker script created"

Write-Host "`nSetting up scheduled task..." -ForegroundColor Cyan

# Create XML template for scheduled task
$startBoundary = (Get-Date).AddMinutes(1).ToString("yyyy-MM-ddTHH:mm:ss")

$xmlContent = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>PowerShell AutoSigner - Signs and versions PowerShell scripts every 5 minutes</Description>
  </RegistrationInfo>
  <Triggers>
    <TimeTrigger>
      <Repetition>
        <Interval>PT5M</Interval>
      </Repetition>
      <StartBoundary>$startBoundary</StartBoundary>
      <Enabled>true</Enabled>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$env:USERNAME</UserId>
      <LogonType>S4U</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
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
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>PowerShell.exe</Command>
      <Arguments>-NoProfile -ExecutionPolicy Bypass -File &quot;$WorkerScript&quot; -ConfigFile &quot;$ConfigFile&quot;</Arguments>
    </Exec>
  </Actions>
</Task>
"@

$xmlFile = Join-Path $WorkingDir "task.xml"
$xmlContent | Out-File $xmlFile -Encoding Unicode

# Create the task using XML with proper quoting
Write-Host "Creating task: $TaskName"
$result = schtasks.exe /Create /TN "`"$TaskName`"" /XML "`"$xmlFile`"" /F
if ($LASTEXITCODE -eq 0) {
    Write-Host "Scheduled task created successfully: $TaskName"
    Remove-Item $xmlFile -Force -ErrorAction SilentlyContinue
} else {
    Write-Host "Task creation failed with exit code: $LASTEXITCODE"
    Write-Host "Error output: $result"
    Write-Host "XML file saved at: $xmlFile"
    Write-Host "You can manually import this XML file into Task Scheduler"
}

Write-Host "Scheduled task created: $TaskName"

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host "PowerShell AutoSigner Setup Complete!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Configuration:" -ForegroundColor Yellow
Write-Host "  ToSign Folder: $($config.ToSignFolder)"
Write-Host "  Signed Folder: $($config.SignedFolder)" 
Write-Host "  Working Dir:   $WorkingDir"
Write-Host "  Task Name:     $TaskName"
Write-Host "  Log File:      $LogFile"
Write-Host ""
Write-Host "Usage:" -ForegroundColor Yellow
Write-Host "  1. Drop PowerShell files (.ps1, .psm1, .psd1) into: $($config.ToSignFolder)"
Write-Host "  2. Files are automatically signed and moved to: $($config.SignedFolder)"
Write-Host "  3. Add '# bump: major' or '# bump: minor' comments to control version increments"
Write-Host "  4. Check logs at: $LogFile"
Write-Host ""
Write-Host "The system will run every 5 minutes automatically." -ForegroundColor Cyan
