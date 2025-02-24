function Get-NetworkPrinters {
    [CmdletBinding()]
    param()
    
    $discoveredPrinters = @{}
    
    # Method 1: Check common printer ports
    $commonPorts = @(9100, 515, 631, 80, 443)
    $networkPrefix = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.PrefixOrigin -eq 'Dhcp' }).IPAddress
    if ($networkPrefix) {
        $subnet = $networkPrefix -replace "\.\d+$", ""
        
        foreach ($i in 1..254) {
            $ip = "$subnet.$i"
            foreach ($port in $commonPorts) {
                try {
                    $tcpClient = New-Object System.Net.Sockets.TcpClient
                    $asyncResult = $tcpClient.BeginConnect($ip, $port, $null, $null)
                    $success = $asyncResult.AsyncWaitHandle.WaitOne(100) # 100ms timeout
                    if ($success) {
                        $tcpClient.EndConnect($asyncResult)
                        if (-not $discoveredPrinters.ContainsKey($ip)) {
                            $printerInfo = @{
                                DetectionMethod = "Port Scan"
                                OpenPorts = @($port)
                            }
                            $discoveredPrinters[$ip] = $printerInfo
                        } else {
                            $discoveredPrinters[$ip].OpenPorts += $port
                        }
                    }
                }
                catch {}
                finally {
                    if ($tcpClient) { $tcpClient.Close() }
                }
            }
        }
    }

    # Method 2: Windows Print Server Discovery
    try {
        Get-PrinterPort | Where-Object { $_.Description -eq "Standard TCP/IP Port" } | ForEach-Object {
            $ip = $_.PrinterHostAddress
            if ($ip -and -not $discoveredPrinters.ContainsKey($ip)) {
                $printerInfo = @{
                    DetectionMethod = "Windows Print Server"
                    PrinterPort = $_.Name
                }
                $discoveredPrinters[$ip] = $printerInfo
            }
        }
    }
    catch {}

    # Method 3: WSD (Web Services for Devices) Discovery
    try {
        $wsdSearch = New-Object -ComObject WSDDiscovery.WSDDiscoveryPublisher
        $wsdSearch.SearchById("urn:schemas-microsoft-com:device:PrintDeviceType:1")
        Start-Sleep -Seconds 2 # Give time for responses
        if (-not $discoveredPrinters.ContainsKey($ip)) {
            $printerInfo = @{
                DetectionMethod = "WSD"
            }
            $discoveredPrinters[$ip] = $printerInfo
        }
    }
    catch {}

    # Method 4: mDNS/Bonjour Discovery (IPP)
    try {
        $ipps = [System.Net.Dns]::GetHostEntry("_ipp._tcp.local")
        foreach ($address in $ipps.AddressList) {
            if (-not $discoveredPrinters.ContainsKey($address.ToString())) {
                $printerInfo = @{
                    DetectionMethod = "mDNS/Bonjour"
                }
                $discoveredPrinters[$address.ToString()] = $printerInfo
            }
        }
    }
    catch {}

    # Method 5: Check for common printer HTTP endpoints
    $webEndpoints = @("/", "/hp/device/info", "/web/info.html", "/printer", "/Printer", "/status.html")
    foreach ($ip in $discoveredPrinters.Keys) {
        foreach ($endpoint in $webEndpoints) {
            try {
                $webRequest = [System.Net.WebRequest]::Create("http://$ip$endpoint")
                $webRequest.Timeout = 1000
                $response = $webRequest.GetResponse()
                $stream = $response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $content = $reader.ReadToEnd()
                
                # Look for printer-specific keywords in response
                if ($content -match "printer|copier|scanner|HP|Epson|Canon|Xerox|Brother|Lexmark") {
                    $discoveredPrinters[$ip].WebInterface = $true
                    $discoveredPrinters[$ip].WebEndpoint = $endpoint
                    break
                }
            }
            catch {}
        }
    }

    # Get MAC addresses for discovered printers
    foreach ($ip in $discoveredPrinters.Keys) {
        try {
            $arpResult = arp -a | Select-String -Pattern $ip
            if ($arpResult -match "([0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2})") {
                $discoveredPrinters[$ip].MACAddress = $matches[1]
                
                # Check MAC address prefixes against known printer manufacturers
                $macPrefix = ($matches[1] -split '[:-]')[0..2] -join ':'
                $discoveredPrinters[$ip].PossibleManufacturer = switch -Wildcard ($macPrefix.ToUpper()) {
                    "00:17:C8" { "HP" }
                    "00:80:77" { "Brother" }
                    "08:00:37" { "Xerox" }
                    "00:00:74" { "Ricoh" }
                    "00:26:73" { "Epson" }
                    "00:00:85" { "Canon" }
                    default { "---" }
                }
            }
        }
        catch {}
    }

    return $discoveredPrinters
}

# Main printer information collection function
function Get-PrinterDetails {
    [CmdletBinding()]
    param()
    
    # Get configured printers
    $printerDetails = @()
    $allPrinters = Get-Printer | Sort-Object Type, Name
    
    # Get network printers
    Write-Host "Discovering network printers... This may take a few moments."
    $networkPrinters = Get-NetworkPrinters
    
    # Process configured printers
    foreach ($printer in $allPrinters) {
        $printerName = $printer.Name
        $printerType = $printer.Type
        $printerMake = if ($printer.Manufacturer) { $printer.Manufacturer } else { "---" }
        $printerModel = if ($printer.DriverName) { $printer.DriverName } else { "---" }
        $printerPortName = $printer.PortName
        $printerShared = $printer.Shared
        
        # Get detailed driver information
        try {
            $driverInfo = Get-PrinterDriver -Name $printer.DriverName | Select-Object -Property *
            $driverVersion = if ($driverInfo.DriverVersion) {
                $driverInfo.DriverVersion.ToString()
            } elseif ($driverInfo.MajorVersion -or $driverInfo.MinorVersion) {
                "$($driverInfo.MajorVersion).$($driverInfo.MinorVersion)"
            } else {
                "---"
            }
        }
        catch {
            $driverVersion = "---"
        }
        
        # Get port info including IP address
        try {
            $portInfo = Get-PrinterPort -Name $printerPortName -ErrorAction SilentlyContinue
            $ipAddress = $portInfo.PrinterHostAddress
            
            if ([string]::IsNullOrEmpty($ipAddress) -and $printerPortName -match "^(TCP|IP)_") {
                $ipAddress = $printerPortName -replace "^(TCP|IP)_", ""
            }
            
            if ([string]::IsNullOrEmpty($ipAddress)) {
                if ($printerPortName -match "^(COM|LPT)") {
                    $ipAddress = "LOCAL"
                } else {
                    $ipAddress = "---"
                }
            }
        }
        catch {
            $ipAddress = "---"
        }
        
        # Create custom object for this printer
        $printerInfo = [PSCustomObject]@{
            PrinterName = $printerName
            Type = $printerType
            Manufacturer = $printerMake
            Model = $printerModel
            DriverVersion = $driverVersion
            PortName = $printerPortName
            IPAddress = $ipAddress
            MACAddress = "---"
            DetectionMethod = "Windows Config"
            OpenPorts = "---"
            Shared = $printerShared
        }
        
        $printerDetails += $printerInfo
    }
    
    # Add discovered network printers that aren't configured
    foreach ($netPrinter in $networkPrinters.GetEnumerator()) {
        if (-not ($printerDetails | Where-Object { $_.IPAddress -eq $netPrinter.Key })) {
            $manufacturer = if ($netPrinter.Value.PossibleManufacturer) {
                $netPrinter.Value.PossibleManufacturer
            } else {
                "---"
            }
            
            $macAddress = if ($netPrinter.Value.MACAddress) {
                $netPrinter.Value.MACAddress
            } else {
                "---"
            }
            
            $openPorts = if ($netPrinter.Value.OpenPorts) {
                $netPrinter.Value.OpenPorts -join ", "
            } else {
                "---"
            }
            
            $printerInfo = [PSCustomObject]@{
                PrinterName = "UNCONFIGURED"
                Type = "Network"
                Manufacturer = $manufacturer
                Model = "---"
                DriverVersion = "---"
                PortName = "---"
                IPAddress = $netPrinter.Key
                MACAddress = $macAddress
                DetectionMethod = $netPrinter.Value.DetectionMethod
                OpenPorts = $openPorts
                Shared = $false
            }
            $printerDetails += $printerInfo
        }
    }
    
    return $printerDetails
}

# Execute and format output
$printerData = Get-PrinterDetails

# Output results
$printerData | Select-Object PrinterName, Manufacturer, Model, DriverVersion, IPAddress, MACAddress, DetectionMethod, OpenPorts, WebInterface, WebEndpoint | Format-Table -AutoSize
