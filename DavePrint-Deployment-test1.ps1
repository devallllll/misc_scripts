#Daves Printer Magic Voodoo Driver installer
#Designed to replace printers with AI
<#
.SYNOPSIS
    Daves Printer Magic Voodoo Driver Installer
.DESCRIPTION
    Automated printer deployment script with support for multiple manufacturers and models
.VERSION
    1.0.0
.NOTES
    Requires admin rights and Chocolatey for some features
	Copyright David Lane 2025
#>

# RMM Wrapper for Printer Installation
$printerIP = "@printerIP@"
$printerModel = "@printerModel@"  # This should match our driver keys (e.g., "hp-generic", "epson-t5100m")

# Basic validation
if ($printerIP -eq "" -or $printerModel -eq "") {
    Write-Error "Printer IP and Model are required"
    exit 1
}

# Validate IP format
if ($printerIP -notmatch "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$") {
    Write-Error "Invalid IP address format"
    exit 1
}

# Validate model exists in our configuration
if (-not ($knownDrivers.ContainsKey($printerModel) -or $specificPrinters.ContainsKey($printerModel))) {
    Write-Error "Unknown printer model: $printerModel"
    Write-Host "Available models:"
    $knownDrivers.Keys | ForEach-Object { Write-Host "- $_" }
    $specificPrinters.Keys | ForEach-Object { Write-Host "- $_" }
    exit 1
}

# Set printer name based on model's friendly name
$driverConfig = if ($knownDrivers.ContainsKey($printerModel)) {
    $knownDrivers[$printerModel]
} else {
    $specificPrinters[$printerModel]
}

$printerName = "$($driverConfig.FriendlyName) ($printerIP)"

# Install printer
Add-NetworkPrinter -PrinterName $printerName `
    -IPAddress $printerIP `
    -DriverKey $printerModel `
    -Force

# Output result for RMM
if ($?) {
    Write-Host "Printer installation successful"
    exit 0
} else {
    Write-Error "Printer installation failed"
    exit 1
}

# Verify running as admin
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be run as Administrator"
    exit 1
}

# Check if Chocolatey is installed
$chocoInstalled = Get-Command choco -ErrorAction SilentlyContinue
if (-not $chocoInstalled) {
    Write-Warning "Chocolatey is not installed. Some driver installations may fail."
}


# Core configuration for printer deployment
$printerConfig = @{
    # Installation paths
    TempPath = "$env:TEMP\PrinterInstall"
    LogPath = "$env:ProgramData\PrinterDeployment\Logs"
    ExtractPath = "$env:TEMP\printer_driver"
    
    # Default printer settings
    DefaultSettings = @{
        PaperSize = "A4"                # Change to "Letter" for US
        Duplex = $true                  # Enable duplex by default
        Color = $true                   # Enable color by default
        InputBin = "Auto"
        Orientation = "Portrait"
        DPI = "600"
    }
    
    # Queue management
    QueueSettings = @{
        CreateBWQueue = $false          # Don't create B&W queue by default
        SetBWAsDefault = $false         # Don't set B&W as default queue
        BWSuffix = " - BW"              # Suffix for B&W queue
        ColorSuffix = " - Color"        # Suffix for color queue
    }

    # Network and Port Settings
    PortConfig = @{
        Protocol = "TCP/IP"
        PortNumber = "9100"             # Default printer port
        SNMPEnabled = $true
        SNMPCommunity = "public"
        EnableLPR = $false
        LPRQueue = "print"
        PortMonitorDLL = "tcpmon.dll"
    }

    # Deployment Options
    Deployment = @{
        Shared = $false                 # Whether to share the printer
        ShareName = ""                  # Name if shared
        RemoveExisting = $true          # Remove existing printer with same name
        CleanupFiles = $true           # Remove installation files after setup
        SetDefaultPrinter = $false      # Don't set as default printer
        WaitForSpooler = $true         # Wait for spooler after installation
        SpoolerTimeout = 300           # Timeout in seconds
    }

    # Error Handling
    ErrorHandling = @{
        RetryAttempts = 3              # Number of retry attempts
        RetryDelay = 10                # Seconds between retries
        IgnoreWarnings = $false        # Continue on non-critical warnings
        VerboseLogging = $true         # Enable detailed logging
    }
}

# Known printer drivers with installation details
$knownDrivers = @{
    # HP Drivers
    "hp-generic" = @{
        FriendlyName = "HP Generic"
        Manufacturer = "HP Inc."
        DriverName = "HP Universal Printing PCL6"
        ChocoPackage = "hp-universal-print-driver"
        DirectURL = "https://ftp.hp.com/pub/softlib/software13/COL14290/HPUPD64_64.zip"
        SilentArgs = "/s /v`"/qn INSTALLTYPE=130`""
        SearchPattern = "*HP*Universal*PCL*"
        PreInstall = { Stop-Service -Name "HP Print Sandbox Service" -ErrorAction SilentlyContinue }
        BWModeSupported = $true
        ColorModeSupported = $true
        DefaultQuality = "Normal"
    }
    
    "hp-generic-ps" = @{
        FriendlyName = "HP Generic PS"
        DriverName = "HP Universal Printing PS"
        DirectURL = "https://ftp.hp.com/pub/softlib/software13/COL40842/ds-99374-14/upd-ps-x64-7.0.1.24923.exe"
        SilentArgs = "/s /v`"/qn INSTALLTYPE=130`""
        SearchPattern = "*HP*Universal*PS*"
        BWModeSupported = $true
        ColorModeSupported = $true
        DefaultQuality = "Normal"
    }

    # Canon Drivers
    "canon-generic" = @{
        FriendlyName = "Canon Generic"
        Manufacturer = "Canon Inc."
        DriverName = "Canon Generic Plus UFR II"
        DirectURL = "https://gdlp01.c-wss.com/gds/7/0100003577/01/UFRII_PrinterDriver_Win64_V21.exe"
        SilentArgs = "/quiet /norestart"
        SearchPattern = "*Canon*UFR*II*"
        BWModeSupported = $true
        ColorModeSupported = $true
        DefaultQuality = "Standard"
        ColorProfiles = @{
            Office = "sRGB"
            Professional = "AdobeRGB"
        }
    }
    
    "canon-generic-ps" = @{
        FriendlyName = "Canon Generic PS"
        DriverName = "Canon Generic Plus PS3"
        DirectURL = "https://gdlp01.c-wss.com/gds/0/0100011130/01/PS3_PrinterDriver_V320_00_W64.exe"
        SilentArgs = "/quiet /norestart"
        SearchPattern = "*Canon*PS3*"
        BWModeSupported = $true
        ColorModeSupported = $true
    }

    # Epson Drivers
    "epson-generic" = @{
        FriendlyName = "Epson Generic"
        Manufacturer = "Epson"
        DriverName = "EPSON Universal Print Driver"
        DirectURL = "https://download.epson-europe.com/pub/download/6583/epson658359eu.exe"
        SilentArgs = "/s /f"
        SearchPattern = "*EPSON*Universal*"
        BWModeSupported = $true
        ColorModeSupported = $true
        QualitySettings = @{
            Draft = "-2"
            Normal = "0"
            High = "2"
        }
    }

    # Brother Drivers
    "brother-generic" = @{
        FriendlyName = "Brother Generic"
        Manufacturer = "Brother Industries, Ltd."
        DriverName = "Brother Universal Printer Driver"
        DirectURL = "https://download.brother.com/pub/driver/print/uni/pcl/UniversalPrinterDriver_Win64.exe"
        SilentArgs = "/q /r"
        SearchPattern = "*Brother*Universal*"
        BWModeSupported = $true
        ColorModeSupported = $true
    }

    # Xerox Drivers
    "xerox-generic" = @{
        FriendlyName = "Xerox Generic"
        Manufacturer = "Xerox Corporation"
        DriverName = "Xerox Global Print Driver"
        ChocoPackage = "xerox-global-print-driver"
        DirectURL = "https://download.support.xerox.com/pub/drivers/GLOBALPRINT/drivers/win10x64/ar/XeroxGlobalDriverPCL6_X.exe"
        SilentArgs = "/install /quiet"
        SearchPattern = "*Xerox*Global*"
        BWModeSupported = $true
        ColorModeSupported = $true
        DefaultTrays = @{
            Letter = "Tray1"
            A4 = "Tray2"
        }
    }
    
    "xerox-generic-ps" = @{
        FriendlyName = "Xerox Generic PS"
        DriverName = "Xerox Global Print Driver PS"
        DirectURL = "https://download.support.xerox.com/pub/drivers/GLOBALPRINT/drivers/win10x64/ar/XeroxGlobalDriverPS_X.exe"
        SilentArgs = "/install /quiet"
        SearchPattern = "*Xerox*Global*PS*"
        BWModeSupported = $true
        ColorModeSupported = $true
    }

    # Dymo Drivers
    "dymo-generic" = @{
        FriendlyName = "DYMO Generic"
        Manufacturer = "Dymo"
        DriverName = "DYMO LabelWriter Driver"
        DirectURL = "https://download.dymo.com/dymo/Software/Win/DCD8Setup.8.7.5.exe"
        SilentArgs = "/quiet /norestart"
        SearchPattern = "*DYMO*"
        BWModeSupported = $true
        ColorModeSupported = $false
    }

    # Ricoh Drivers
    "ricoh-generic" = @{
        FriendlyName = "Ricoh Generic"
        Manufacturer = "Ricoh"
        DriverName = "Ricoh Universal Print Driver"
        DirectURL = "https://support.ricoh.com/bb/pub_e/dr_ut_e/0001331/0001331748/V440/z88894L16.exe"
        SilentArgs = "/silent"
        SearchPattern = "*Ricoh*Universal*"
        BWModeSupported = $true
        ColorModeSupported = $true
    }

    # Lexmark Drivers
    "lexmark-generic" = @{
        FriendlyName = "Lexmark Generic"
        Manufacturer = "Lexmark"
        DriverName = "Lexmark Universal v2"
        DirectURL = "https://downloads.lexmark.com/downloads/drivers/universal_print_driver_x64.exe"
        SilentArgs = "/silent /norestart"
        SearchPattern = "*Lexmark*Universal*"
        BWModeSupported = $true
        ColorModeSupported = $true
    }

    # Konica Minolta Drivers
    "konica-generic" = @{
        FriendlyName = "Konica Generic"
        Manufacturer = "Konica Minolta"
        DriverName = "Konica Minolta Universal Print Driver"
        DirectURL = "https://dl.konicaminolta.eu/en/?tx_kmanacondaimport_downloadproxy[fileId]=9417971e551f4f5252b2ca6027f44c3b&tx_kmanacondaimport_downloadproxy[documentId]=128235&tx_kmanacondaimport_downloadproxy[system]=KonicaMinolta&tx_kmanacondaimport_downloadproxy[language]=EN&type=1558521685"
        SilentArgs = "/silent"
        SearchPattern = "*Konica*Universal*"
        BWModeSupported = $true
        ColorModeSupported = $true
    }

    # Samsung Drivers
    "samsung-generic" = @{
        FriendlyName = "Samsung Generic"
        Manufacturer = "Samsung Electronics Co., Ltd."
        DriverName = "Samsung Universal Print Driver"
        DirectURL = "https://downloadcenter.samsung.com/downloadfile/002/0000021490/SUPD64.exe"
        SilentArgs = "/s /v`"/qn`""
        SearchPattern = "*Samsung*Universal*"
        BWModeSupported = $true
        ColorModeSupported = $true
    }

    # Toshiba Drivers
    "toshiba-generic" = @{
        FriendlyName = "Toshiba Generic"
        Manufacturer = "Toshiba Tec Corporation"
        DriverName = "Toshiba Universal Print Driver"
        DirectURL = "https://download.toshibatec.com/Drivers/Toshiba_UPD_Win64.exe"
        SilentArgs = "/quiet /norestart"
        SearchPattern = "*Toshiba*Universal*"
        BWModeSupported = $true
        ColorModeSupported = $true
    }

    # Sharp Drivers
    "sharp-generic" = @{
        FriendlyName = "Sharp Generic"
        Manufacturer = "Sharp Corporation"
        DriverName = "Sharp Universal Print Driver"
        DirectURL = "https://download.sharpusa.com/Drivers/Sharp_UPD_Win64.exe"
        SilentArgs = "/s /v`"/qn`""
        SearchPattern = "*Sharp*Universal*"
        BWModeSupported = $true
        ColorModeSupported = $true
    }
}

# Required Windows features and services
$requiredFeatures = @{
    WindowsFeatures = @(
        "Printing-PrintToPDFServices-Features",
        "Printing-XPSServices-Features"
    )
    Services = @(
        @{ Name = "Spooler"; StartupType = "Automatic" },
        @{ Name = "PrintNotify"; StartupType = "Automatic" }
    )
    FirewallRules = @{
        "TCP" = 9100,
        "UDP" = 161  # SNMP
    }
}

# Known specific printer models
$specificPrinters = @{
    # Epson Specific Models
    "epson-t5100m" = @{
        FriendlyName = "Epson SureColor T5100M"
        Manufacturer = "Epson"
        DriverName = "EPSON SC-T5100M Series"
        DirectURL = "https://download.epson-europe.com/pub/download/6583/epson658359eu.exe"
        SilentArgs = "/s /f"
        SearchPattern = "*EPSON*T5100*"
        BWModeSupported = $true
        ColorModeSupported = $true
        DefaultPaperSize = "A1"
        QualitySettings = @{
            Draft = "-2"
            Normal = "0"
            High = "2"
        }
    }

    # Add more specific models here...
}

# Example usage:
$printerSetup = @{
    Name = "Epson SC-T5100M"
    IPAddress = "192.168.2.231"
    Location = "Design Office"
    DriverConfig = $specificPrinters["epson-t5100m"]
    QueueSettings = @{
        CreateBWQueue = $false
        SetBWAsDefault = $false
    }
}


# Function to write to log file
function Write-PrinterLog {
    param(
        [string]$Message,
        [string]$Type = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Type] $Message"
    
    # Create log directory if it doesn't exist
    $logDir = Split-Path -Parent $printerConfig.LogPath
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    # Write to log file
    $logMessage | Out-File -FilePath $printerConfig.LogPath -Append
    
    # Output to console based on message type
    switch ($Type) {
        "ERROR"   { Write-Error $Message }
        "WARNING" { Write-Warning $Message }
        default   { Write-Host $Message }
    }
}

# Function to download and install printer driver
function Install-PrinterDriver {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DriverKey,
        [switch]$Force
    )
    
    Write-PrinterLog "Starting driver installation for: $DriverKey"
    
    # Validate driver exists
    if (-not $knownDrivers.ContainsKey($DriverKey)) {
        Write-PrinterLog "Unknown driver key: $DriverKey" -Type "ERROR"
        return $false
    }
    
    $driver = $knownDrivers[$DriverKey]
    
    # Create temp directory
    if (-not (Test-Path $printerConfig.TempPath)) {
        New-Item -ItemType Directory -Path $printerConfig.TempPath -Force | Out-Null
    }
    
    # Remove existing driver if force is specified
    if ($Force) {
        Write-PrinterLog "Removing existing drivers matching: $($driver.SearchPattern)"
        Get-PrinterDriver | Where-Object { $_.Name -like $driver.SearchPattern } | ForEach-Object {
            try {
                Remove-PrinterDriver -Name $_.Name -RemoveFromDriverStore -ErrorAction Stop
                Write-PrinterLog "Removed driver: $($_.Name)"
            }
            catch {
                Write-PrinterLog "Could not remove driver $($_.Name): $_" -Type "WARNING"
            }
        }
    }
    
    # Try Chocolatey first if package is specified
    if ($driver.ChocoPackage) {
        Write-PrinterLog "Attempting installation via Chocolatey: $($driver.ChocoPackage)"
        try {
            choco install $driver.ChocoPackage -y
            if ($LASTEXITCODE -eq 0) {
                Write-PrinterLog "Chocolatey installation successful"
                return $true
            }
            Write-PrinterLog "Chocolatey installation failed, trying direct download" -Type "WARNING"
        }
        catch {
            Write-PrinterLog "Chocolatey installation failed: $_" -Type "WARNING"
        }
    }
    
    # Download and install driver
    try {
        $installerPath = Join-Path $printerConfig.TempPath "$($DriverKey)_installer$(([IO.Path]::GetExtension($driver.DirectURL)))"
        Write-PrinterLog "Downloading driver from: $($driver.DirectURL)"
        
        # Download with retry
        $maxRetries = $printerConfig.ErrorHandling.RetryAttempts
        $retryCount = 0
        $success = $false
        
        do {
            try {
                Invoke-WebRequest -Uri $driver.DirectURL -OutFile $installerPath
                $success = $true
            }
            catch {
                $retryCount++
                if ($retryCount -ge $maxRetries) {
                    throw
                }
                Write-PrinterLog "Download failed, attempt $retryCount of $maxRetries. Retrying..." -Type "WARNING"
                Start-Sleep -Seconds $printerConfig.ErrorHandling.RetryDelay
            }
        } while (-not $success -and $retryCount -lt $maxRetries)
        
        # Handle ZIP files
        if ($installerPath.EndsWith('.zip')) {
            $extractPath = Join-Path $printerConfig.TempPath "$($DriverKey)_extracted"
            Write-PrinterLog "Extracting ZIP file to: $extractPath"
            Expand-Archive -Path $installerPath -DestinationPath $extractPath -Force
            $installerPath = Get-ChildItem -Path $extractPath -Filter "*.exe" -Recurse | 
                            Select-Object -First 1 -ExpandProperty FullName
        }
        
        # Run pre-install steps if defined
        if ($driver.PreInstall) {
            Write-PrinterLog "Running pre-installation steps"
            & $driver.PreInstall
        }
        
        # Install driver
        Write-PrinterLog "Installing driver with arguments: $($driver.SilentArgs)"
        $process = Start-Process -FilePath $installerPath -ArgumentList $driver.SilentArgs -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-PrinterLog "Driver installation completed successfully"
            
            # Run post-install steps if defined
            if ($driver.PostInstall) {
                Write-PrinterLog "Running post-installation steps"
                & $driver.PostInstall
            }
            
            return $true
        }
        else {
            Write-PrinterLog "Driver installation failed with exit code: $($process.ExitCode)" -Type "ERROR"
            return $false
        }
    }
    catch {
        Write-PrinterLog "Failed to install driver: $_" -Type "ERROR"
        return $false
    }
    finally {
        # Cleanup
        if ($printerConfig.Deployment.CleanupFiles) {
            Remove-Item -Path $printerConfig.TempPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# Function to verify printer port
function Test-PrinterPort {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Address,
        [int]$Port = $printerConfig.PortConfig.PortNumber
    )
    
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $connect = $tcp.BeginConnect($Address, $Port, $null, $null)
        $wait = $connect.AsyncWaitHandle.WaitOne(3000, $false)
        
        if (-not $wait) {
            $tcp.Close()
            return $false
        }
        
        $tcp.EndConnect($connect)
        $tcp.Close()
        return $true
    }
    catch {
        Write-PrinterLog "Port test failed: $_" -Type "WARNING"
        return $false
    }
}

# Function to create printer port
function New-PrinterPortSafe {
    param (
        [Parameter(Mandatory=$true)]
        [string]$IPAddress,
        [string]$PortName = "IP_$IPAddress"
    )
    
    try {
        # Remove existing port if it exists
        Remove-PrinterPort -Name $PortName -ErrorAction SilentlyContinue
        
        # Create new port
        Add-PrinterPort -Name $PortName `
            -PrinterHostAddress $IPAddress `
            -PortNumber $printerConfig.PortConfig.PortNumber
            
        Write-PrinterLog "Created printer port: $PortName"
        return $true
    }
    catch {
        Write-PrinterLog "Failed to create printer port: $_" -Type "ERROR"
        return $false
    }
}


# Main printer deployment function
function Add-NetworkPrinter {
    param (
        [Parameter(Mandatory=$true)]
        [string]$PrinterName,
        
        [Parameter(Mandatory=$true)]
        [string]$IPAddress,
        
        [Parameter(Mandatory=$true)]
        [string]$DriverKey,
        
        [string]$Location = "",
        [switch]$Force,
        [switch]$CreateBWQueue,
        [switch]$SetBWAsDefault
    )
    
    Write-PrinterLog "Starting printer deployment for $PrinterName ($IPAddress)"
    
    # Test network connectivity
    if (-not (Test-PrinterPort -Address $IPAddress)) {
        Write-PrinterLog "Cannot connect to printer at $IPAddress" -Type "ERROR"
        return $false
    }
    
    # Install driver
    if (-not (Install-PrinterDriver -DriverKey $DriverKey -Force:$Force)) {
        Write-PrinterLog "Driver installation failed" -Type "ERROR"
        return $false
    }
    
    # Create printer port
    if (-not (New-PrinterPortSafe -IPAddress $IPAddress)) {
        Write-PrinterLog "Port creation failed" -Type "ERROR"
        return $false
    }
    
    $portName = "IP_$IPAddress"
    $driver = $knownDrivers[$DriverKey]
    
    # Remove existing printers if Force is specified
    if ($Force) {
        Get-Printer | Where-Object { $_.Name -like "*$PrinterName*" } | ForEach-Object {
            try {
                Remove-Printer -Name $_.Name
                Write-PrinterLog "Removed existing printer: $($_.Name)"
            }
            catch {
                Write-PrinterLog "Failed to remove printer $($_.Name): $_" -Type "WARNING"
            }
        }
    }
    
    try {
        # Create color queue
        $colorName = $PrinterName
        Add-Printer -Name $colorName `
            -DriverName $driver.DriverName `
            -PortName $portName `
            -Location $Location
            
        Write-PrinterLog "Created color queue: $colorName"
        
        # Configure default settings
        Set-PrintConfiguration -PrinterName $colorName `
            -PaperSize $printerConfig.DefaultSettings.PaperSize `
            -DuplexingMode $(if ($printerConfig.DefaultSettings.Duplex) { "TwoSidedLongEdge" } else { "OneSided" })
        
        # Create B&W queue if requested and supported
        if ($CreateBWQueue -and $driver.BWModeSupported) {
            $bwName = $PrinterName + $printerConfig.QueueSettings.BWSuffix
            Add-Printer -Name $bwName `
                -DriverName $driver.DriverName `
                -PortName $portName `
                -Location $Location
                
            Write-PrinterLog "Created B&W queue: $bwName"
            
            # Configure B&W settings
            Set-PrintConfiguration -PrinterName $bwName `
                -PaperSize $printerConfig.DefaultSettings.PaperSize `
                -DuplexingMode $(if ($printerConfig.DefaultSettings.Duplex) { "TwoSidedLongEdge" } else { "OneSided" })
            
            # Set as default if requested
            if ($SetBWAsDefault) {
                (Get-CimInstance -Class Win32_Printer -Filter "Name='$bwName'").SetDefaultPrinter()
                Write-PrinterLog "Set $bwName as default printer"
            }
        }
        
        Write-PrinterLog "Printer deployment completed successfully"
        return $true
    }
    catch {
        Write-PrinterLog "Failed to create printer: $_" -Type "ERROR"
        return $false
    }
}

# Example usage:
<#
# Install HP printer with B&W queue
Add-NetworkPrinter -PrinterName "Office HP" `
    -IPAddress "192.168.1.100" `
    -DriverKey "hp-generic" `
    -Location "Main Office" `
    -CreateBWQueue `
    -SetBWAsDefault `
    -Force

# Install Epson plotter without B&W queue
Add-NetworkPrinter -PrinterName "Plotter" `
    -IPAddress "192.168.1.101" `
    -DriverKey "epson-t5100m" `
    -Location "Design Office" `
    -Force
#>
