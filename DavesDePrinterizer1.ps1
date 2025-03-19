# Printer Manager GUI
# Lists all printers (system and user) with checkboxes for selective removal
# Run as administrator

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Initialize arrays for tracking printers
$script:SystemPrinters = @()
$script:UserPrinters = @()

# Function to ensure spooler service is running
function Ensure-SpoolerRunning {
    try {
        $service = Get-Service "Spooler"
        if (-not $service) {
            [System.Windows.Forms.MessageBox]::Show("Spooler service does not exist.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return $false
        }
        
        if ($service.Status -ne [ServiceProcess.ServiceControllerStatus]::Running) {
            $result = [System.Windows.Forms.MessageBox]::Show("Spooler service is not running. Start it now?", "Warning", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                $timeSpan = New-Object Timespan 0, 0, 60
                try {
                    $service.Start()
                    $service.WaitForStatus([ServiceProcess.ServiceControllerStatus]::Running, $timeSpan)
                    return $true
                }
                catch {
                    [System.Windows.Forms.MessageBox]::Show("Failed to start Spooler service: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    return $false
                }
            }
            else {
                return $false
            }
        }
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error accessing Spooler service: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

# Function to get system printers
function Get-AllSystemPrinters {
    try {
        # Create lookup table for printer port IP addresses
        $hostAddresses = @{}
        Get-WmiObject Win32_TCPIPPrinterPort -ErrorAction SilentlyContinue | ForEach-Object {
            $hostAddresses.Add($_.Name, $_.HostAddress)
        }
        
        # Get all printers using WMI
        $printerStatusCode = @{
            "1" = "Other"; "2" = "Unknown"; "3" = "Idle"; "4" = "Printing";
            "5" = "Warmup"; "6" = "Stopped Printing"; "7" = "Offline"
        }
        
        $wmiPrinters = Get-WmiObject -Class "Win32_Printer" -Namespace "root\CIMV2" -ErrorAction Stop
        
        if (-not $wmiPrinters) {
            return @()
        }
        
        $printerInfo = @()
        foreach ($printer in $wmiPrinters) {
            $isNetwork = $printer.PortName -match '^(IP_|\\\\)'
            $printerObj = [PSCustomObject]@{
                Name = $printer.Name
                DriverName = $printer.DriverName
                PortName = $printer.PortName
                HostAddress = $hostAddresses[$printer.PortName]
                Status = $printerStatusCode[[String]$printer.PrinterStatus]
                Location = "System"
                IsNetwork = $isNetwork
                IsLocal = $printer.Local
                Type = if ($isNetwork) { "Network" } else { "Local" }
                ObjectType = "System"
                FullName = $printer.Name  # Used for removal
            }
            
            $printerInfo += $printerObj
            $script:SystemPrinters += $printerObj
        }
        
        return $printerInfo
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to get system printers: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return @()
    }
}

# Function to get user-specific printers
function Get-AllUserPrinters {
    try {
        # Get all user profiles
        $userProfiles = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" | 
                        Where-Object { $_.ProfileImagePath -like "C:\Users\*" -and $_.ProfileImagePath -notlike "C:\Users\default*" }
        
        if (-not $userProfiles) {
            return @()
        }
        
        # Create PSDrive to access user registry hives if it doesn't exist
        if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
            $null = New-PSDrive -Name "HKU" -PSProvider Registry -Root HKEY_USERS -ErrorAction Stop
        }
        
        $allUserPrinters = @()
        
        foreach ($profile in $userProfiles) {
            $sid = $profile.PSChildName
            $userName = Split-Path -Leaf $profile.ProfileImagePath
            $printerRegPath = "HKU:\$sid\Printers\Connections"
            
            if (Test-Path $printerRegPath) {
                $userPrinters = Get-ChildItem -Path $printerRegPath -ErrorAction SilentlyContinue
                
                foreach ($printer in $userPrinters) {
                    $printerName = $printer.PSChildName -replace ',', '\\'
                    
                    $printerObj = [PSCustomObject]@{
                        Name = $printerName
                        DriverName = "N/A"  # Registry doesn't store this
                        PortName = "N/A"
                        HostAddress = "N/A"
                        Status = "N/A"
                        Location = "User: $userName"
                        IsNetwork = $true  # User printers are always network printers
                        IsLocal = $false
                        Type = "Network (User)"
                        ObjectType = "User"
                        FullName = $printerName
                        UserSID = $sid
                        RegPath = $printer.PSPath
                    }
                    
                    $allUserPrinters += $printerObj
                    $script:UserPrinters += $printerObj
                }
            }
        }
        
        return $allUserPrinters
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to get user printers: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return @()
    }
}

# Function to remove selected printers
function Remove-SelectedPrinters {
    param (
        [System.Windows.Forms.ListView]$ListView
    )
    
    $selectedItems = $ListView.CheckedItems
    
    if ($selectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No printers selected for removal.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove $($selectedItems.Count) selected printer(s)?", "Confirm Removal", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $successCount = 0
        $failCount = 0
        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Minimum = 0
        $progressBar.Maximum = $selectedItems.Count
        $progressBar.Step = 1
        $progressBar.Value = 0
        $progressBar.Width = $ListView.Width
        $progressBar.Height = 20
        $progressBar.Top = $ListView.Bottom + 5
        $progressBar.Left = $ListView.Left
        $form.Controls.Add($progressBar)
        
        foreach ($item in $selectedItems) {
            $printerName = $item.Text
            $printerType = $item.SubItems[1].Text
            
            # Update status label
            $statusLabel.Text = "Removing: $printerName"
            $form.Refresh()
            
            if ($printerType -eq "Network (User)") {
                # Find the user printer object
                $userPrinter = $script:UserPrinters | Where-Object { $_.Name -eq $printerName }
                
                if ($userPrinter) {
                    try {
                        Remove-Item -Path $userPrinter.RegPath -Recurse -Force -ErrorAction Stop
                        $successCount++
                    }
                    catch {
                        $failCount++
                        $statusLabel.Text = "Failed: $printerName - $($_.Exception.Message)"
                        $form.Refresh()
                        Start-Sleep -Milliseconds 1000
                    }
                }
            }
            else {
                # System printer
                try {
                    Remove-Printer -Name $printerName -ErrorAction Stop
                    $successCount++
                }
                catch {
                    $failCount++
                    $statusLabel.Text = "Failed: $printerName - $($_.Exception.Message)"
                    $form.Refresh()
                    Start-Sleep -Milliseconds 1000
                }
            }
            
            $progressBar.PerformStep()
        }
        
        # Refresh printer list
        $ListView.Items.Clear()
        $script:SystemPrinters = @()
        $script:UserPrinters = @()
        Populate-PrinterList -ListView $ListView
        
        $form.Controls.Remove($progressBar)
        $statusLabel.Text = "Removal complete: $successCount succeeded, $failCount failed"
    }
}

# Function to populate the ListView with printer data
function Populate-PrinterList {
    param (
        [System.Windows.Forms.ListView]$ListView
    )
    
    $ListView.Items.Clear()
    $statusLabel.Text = "Loading printers..."
    $form.Refresh()
    
    # Get system and user printers
    $allSystemPrinters = Get-AllSystemPrinters
    $allUserPrinters = Get-AllUserPrinters
    
    # Add system printers to ListView
    foreach ($printer in $allSystemPrinters) {
        $item = New-Object System.Windows.Forms.ListViewItem($printer.Name)
        $item.Checked = $false
        $item.SubItems.Add($printer.Type)
        $item.SubItems.Add($printer.PortName)
        $item.SubItems.Add($printer.HostAddress)
        $item.SubItems.Add($printer.Status)
        $item.SubItems.Add($printer.Location)
        
        # Color network printers differently for better visibility
        if ($printer.IsNetwork) {
            $item.BackColor = [System.Drawing.Color]::LightCyan
        }
        
        $ListView.Items.Add($item)
    }
    
    # Add user printers to ListView
    foreach ($printer in $allUserPrinters) {
        $item = New-Object System.Windows.Forms.ListViewItem($printer.Name)
        $item.Checked = $false
        $item.SubItems.Add($printer.Type)
        $item.SubItems.Add($printer.PortName)
        $item.SubItems.Add($printer.HostAddress)
        $item.SubItems.Add($printer.Status)
        $item.SubItems.Add($printer.Location)
        
        # User printers are always network printers, use a different color
        $item.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
        
        $ListView.Items.Add($item)
    }
    
    $totalCount = $allSystemPrinters.Count + $allUserPrinters.Count
    $statusLabel.Text = "Found $totalCount total printers ($($allSystemPrinters.Count) system, $($allUserPrinters.Count) user)"
}

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    [System.Windows.Forms.MessageBox]::Show("This application requires administrator privileges. Please run as administrator.", "Admin Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Printer Manager"
$form.Size = New-Object System.Drawing.Size(980, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Create checkbox for selecting all printers
$selectAllCheckbox = New-Object System.Windows.Forms.CheckBox
$selectAllCheckbox.Text = "Select All"
$selectAllCheckbox.Location = New-Object System.Drawing.Point(10, 10)
$selectAllCheckbox.Size = New-Object System.Drawing.Size(100, 20)
$selectAllCheckbox.Add_CheckedChanged({
    foreach ($item in $listView.Items) {
        $item.Checked = $selectAllCheckbox.Checked
    }
})
$form.Controls.Add($selectAllCheckbox)

# Create checkbox for selecting all network printers
$selectNetworkCheckbox = New-Object System.Windows.Forms.CheckBox
$selectNetworkCheckbox.Text = "Select Network Only"
$selectNetworkCheckbox.Location = New-Object System.Drawing.Point(120, 10)
$selectNetworkCheckbox.Size = New-Object System.Drawing.Size(150, 20)
$selectNetworkCheckbox.Add_CheckedChanged({
    if ($selectNetworkCheckbox.Checked) {
        $selectAllCheckbox.Checked = $false
        foreach ($item in $listView.Items) {
            # Check if it's a network printer based on type or color
            if ($item.SubItems[1].Text -match "Network") {
                $item.Checked = $true
            } else {
                $item.Checked = $false
            }
        }
    }
})
$form.Controls.Add($selectNetworkCheckbox)

# Create refresh button
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(280, 8)
$refreshButton.Size = New-Object System.Drawing.Size(100, 25)
$refreshButton.Text = "Refresh List"
$refreshButton.Add_Click({
    $script:SystemPrinters = @()
    $script:UserPrinters = @()
    $selectAllCheckbox.Checked = $false
    $selectNetworkCheckbox.Checked = $false
    Populate-PrinterList -ListView $listView
})
$form.Controls.Add($refreshButton)

# Create ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10, 40)
$listView.Size = New-Object System.Drawing.Size(945, 550)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.CheckBoxes = $true
$listView.MultiSelect = $true

# Create columns
$listView.Columns.Add("Printer Name", 300)
$listView.Columns.Add("Type", 120)
$listView.Columns.Add("Port", 100)
$listView.Columns.Add("IP Address", 120)
$listView.Columns.Add("Status", 100)
$listView.Columns.Add("Location", 200)

$form.Controls.Add($listView)

# Create status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 600)
$statusLabel.Size = New-Object System.Drawing.Size(945, 20)
$statusLabel.Text = "Ready"
$form.Controls.Add($statusLabel)

# Create remove button
$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Location = New-Object System.Drawing.Point(750, 630)
$removeButton.Size = New-Object System.Drawing.Size(100, 25)
$removeButton.Text = "Remove Selected"
$removeButton.Add_Click({
    Remove-SelectedPrinters -ListView $listView
})
$form.Controls.Add($removeButton)

# Create close button
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Location = New-Object System.Drawing.Point(855, 630)
$closeButton.Size = New-Object System.Drawing.Size(100, 25)
$closeButton.Text = "Close"
$closeButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($closeButton)

# Start the form
if (Ensure-SpoolerRunning) {
    Populate-PrinterList -ListView $listView
    [void]$form.ShowDialog()
}
