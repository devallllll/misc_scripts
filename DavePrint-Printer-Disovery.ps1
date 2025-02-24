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
                            $discoveredPrinters[$ip] = @{
                                IPAddress = $ip
                                OpenPorts = @($port)
                                DetectionMethod = "Port Scan"
                                Name = "UNCONFIGURED"
                                Model = "---"
                                Manufacturer = "---"
                                MACAddress = "---"
                            }
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

    # Method 2: Get MAC addresses and manufacturer info
    foreach ($ip in $discoveredPrinters.Keys) {
        try {
            $arp = arp -a | Where-Object { $_ -match $ip }
            if ($arp -match "([0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2})") {
                $mac = $matches[1].Replace('-','').Replace(':','').ToLower()
                $discoveredPrinters[$ip].MACAddress = $mac
                
                # Check MAC prefix for manufacturer
                $macPrefix = $mac.Substring(0,6)
                $discoveredPrinters[$ip].Manufacturer = switch ($macPrefix) {
                    "0017c8" { "HP" }
                    "008077" { "Brother" }
                    "080037" { "Xerox" }
                    "000074" { "Ricoh" }
                    "002673" { "Epson" }
                    "000085" { "Canon" }
                    default { "---" }
                }
            }
        }
        catch {}
    }

    # Method 3: Check web interfaces for additional info
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
                
                if ($content -match "printer|copier|scanner|HP|Epson|Canon|Xerox|Brother|Lexmark") {
                    $discoveredPrinters[$ip].WebInterface = $true
                    break
                }
            }
            catch {}
        }
    }

    return $discoveredPrinters
}

function Get-PrinterInfo {
    [CmdletBinding()]
    param()
    
    # First get all network printers
    $allPrinters = Get-NetworkPrinters
    
    # Now get all installed printers
    $installedPrinters = Get-Printer | Sort-Object Type, Name
    
    # Update network printer info with installed printer details
    foreach ($printer in $installedPrinters) {
        $portInfo = Get-PrinterPort -Name $printer.PortName -ErrorAction SilentlyContinue
        $ip = if ($portInfo.PrinterHostAddress) {
            $portInfo.PrinterHostAddress
        } elseif ($printer.PortName -match "^(TCP|IP)_(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})") {
            $matches[2]
        } else {
            "LOCAL"
        }
        
        # If this is a network printer we discovered
        if ($ip -ne "LOCAL" -and $allPrinters.ContainsKey($ip)) {
            # Get driver information
            try {
                $driverInfo = Get-PrinterDriver -Name $printer.DriverName | Select-Object -Property *
                $driverVersion = if ($driverInfo.DriverVersion) {
                    try {
                        $major = [math]::Floor($driverInfo.DriverVersion / 1000000)
                        $minor = $driverInfo.DriverVersion % 1000000
                        "$major.$minor"
                    } catch {
                        "---"
                    }
                } else {
                    "---"
                }
                
                # Update discovered printer with installed printer info
                $allPrinters[$ip].Name = $printer.Name
                $allPrinters[$ip].Model = $printer.DriverName
                $allPrinters[$ip].DriverName = "$($driverInfo.Name) $driverVersion"
                $allPrinters[$ip].DriverPath = $driverInfo.Path
                $allPrinters[$ip].DriverVersion = $driverVersion
                # Keep existing MAC and detection method if we found it
                if ($allPrinters[$ip].DetectionMethod -ne "Windows Config") {
                    $allPrinters[$ip].DetectionMethod += ", Windows Config"
                }
            }
            catch {}
        }
        # If this is a network printer we didn't discover
        elseif ($ip -ne "LOCAL") {
            try {
                $driverInfo = Get-PrinterDriver -Name $printer.DriverName | Select-Object -Property *
                $driverVersion = if ($driverInfo.DriverVersion) {
                    try {
                        $major = [math]::Floor($driverInfo.DriverVersion / 1000000)
                        $minor = $driverInfo.DriverVersion % 1000000
                        "$major.$minor"
                    } catch {
                        "---"
                    }
                } else {
                    "---"
                }
                
                # Add to our collection
                $allPrinters[$ip] = @{
                    IPAddress = $ip
                    Name = $printer.Name
                    Model = $printer.DriverName
                    Manufacturer = if ($printer.Manufacturer) { $printer.Manufacturer } else { "---" }
                    DriverName = "$($driverInfo.Name) $driverVersion"
                    DriverPath = $driverInfo.Path
                    DriverVersion = $driverVersion
                    DetectionMethod = "Windows Config Only"
                    MACAddress = "---"
                    OpenPorts = "---"
                }
                
                # Try to get MAC address
                $arp = arp -a | Where-Object { $_ -match $ip }
                if ($arp -match "([0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2})") {
                    $allPrinters[$ip].MACAddress = $matches[1].Replace('-','').Replace(':','').ToLower()
                }
            }
            catch {}
        }
    }

    # Convert to output objects
    $results = foreach ($printer in $allPrinters.Values) {
        [PSCustomObject]@{
            'Printer Name' = $printer.Name
            'IP Address' = $printer.IPAddress
            'MAC Address' = $printer.MACAddress
            'Model' = $printer.Model
            'Driver' = $printer.DriverName
            'Detection' = $printer.DetectionMethod
            'Ports' = if ($printer.OpenPorts) { $printer.OpenPorts -join ', ' } else { "---" }
        }
    }

    return $results | Sort-Object 'IP Address'
}

# Execute and display results
Get-PrinterInfo | Format-Table -AutoSize
