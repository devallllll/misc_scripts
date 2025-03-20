$form.Controls.Add($updateDriversButton)# Printer Manager GUI
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
    
    # Ask about removing associated ports and drivers
    $removePortsChecked = $removePortsCheckbox.Checked
    $removeDriversChecked = $removeDriversCheckbox.Checked
    
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
        
        # Track printer ports and drivers to remove
        $portsToRemove = @()
        $driversToRemove = @()
        
        foreach ($item in $selectedItems) {
            $printerName = $item.Text
            $printerType = $item.SubItems[1].Text
            $portName = $item.SubItems[2].Text
            
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
                    # Record port name for later removal if it's a network port
                    if ($removePortsChecked -and $portName -match '^(IP_|\\\\)' -and $portName -notmatch "PDF|txt|fax|usb|enhanced|epson|microsoft") {
                        $portsToRemove += $portName
                    }
                    
                    # Get driver name for later removal
                    if ($removeDriversChecked) {
                        $driver = (Get-Printer -Name $printerName -ErrorAction SilentlyContinue).DriverName
                        if ($driver) {
                            $driversToRemove += $driver
                        }
                    }
                    
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
        
        # Remove printer ports if requested
        if ($removePortsChecked -and $portsToRemove.Count -gt 0) {
            $statusLabel.Text = "Removing printer ports..."
            $form.Refresh()
            
            # Make ports unique
            $portsToRemove = $portsToRemove | Select-Object -Unique
            
            foreach ($port in $portsToRemove) {
                try {
                    Remove-PrinterPort -Name $port -ErrorAction SilentlyContinue
                }
                catch {
                    # Just continue if port removal fails
                }
            }
        }
        
        # Remove printer drivers if requested
        if ($removeDriversChecked -and $driversToRemove.Count -gt 0) {
            $statusLabel.Text = "Removing printer drivers..."
            $form.Refresh()
            
            # Make drivers unique
            $driversToRemove = $driversToRemove | Select-Object -Unique
            
            foreach ($driver in $driversToRemove) {
                try {
                    Remove-PrinterDriver -Name $driver -ErrorAction SilentlyContinue
                }
                catch {
                    # Just continue if driver removal fails
                }
            }
        }
        
        # Simple restart of the Print Spooler service
        $statusLabel.Text = "Restarting Print Spooler service..."
        $form.Refresh()
        
        try {
            Restart-Service -Name "Spooler" -Force -ErrorAction Stop
            Start-Sleep -Seconds 5  # Simple wait for 5 seconds
        }
        catch {
            $statusLabel.Text = "Failed to restart Print Spooler service. You may need to restart it manually."
        }
        
        # Refresh printer list
        $ListView.Items.Clear()
        $script:SystemPrinters = @()
        $script:UserPrinters = @()
        Populate-PrinterList -ListView $ListView
        
        if ($progressBar -ne $null -and $form.Controls.Contains($progressBar)) {
            $form.Controls.Remove($progressBar)
        }
        
        $statusLabel.Text = "Removal complete: $successCount succeeded, $failCount failed."
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

# Check if running as administrator - This is performed right at script start
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    # Display more prominent warning about admin privileges
    $warnForm = New-Object System.Windows.Forms.Form
    $warnForm.Text = "ADMINISTRATOR PRIVILEGES REQUIRED"
    $warnForm.Size = New-Object System.Drawing.Size(500, 200)
    $warnForm.StartPosition = "CenterScreen"
    $warnForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $warnForm.MaximizeBox = $false
    $warnForm.BackColor = [System.Drawing.Color]::LightYellow
    
    $warnIcon = New-Object System.Windows.Forms.PictureBox
    $warnIcon.Location = New-Object System.Drawing.Point(20, 20)
    $warnIcon.Size = New-Object System.Drawing.Size(48, 48)
    $warnIcon.Image = [System.Drawing.SystemIcons]::Warning.ToBitmap()
    $warnForm.Controls.Add($warnIcon)
    
    $warnLabel = New-Object System.Windows.Forms.Label
    $warnLabel.Location = New-Object System.Drawing.Point(80, 20)
    $warnLabel.Size = New-Object System.Drawing.Size(390, 80)
    $warnLabel.Text = "This application requires administrator privileges to manage printers.`n`nPlease close this window and run the script again by right-clicking and selecting 'Run as administrator'."
    $warnLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $warnForm.Controls.Add($warnLabel)
    
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(200, 120)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.Text = "Close"
    $closeButton.Add_Click({ $warnForm.Close() })
    $warnForm.Controls.Add($closeButton)
    
    [void]$warnForm.ShowDialog()
    exit
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Printer Manager"
$form.Size = New-Object System.Drawing.Size(980, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Create checkboxes for various options
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

# Add checkbox for removing printer ports
$removePortsCheckbox = New-Object System.Windows.Forms.CheckBox
$removePortsCheckbox.Text = "Remove Printer Ports"
$removePortsCheckbox.Location = New-Object System.Drawing.Point(280, 10)
$removePortsCheckbox.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($removePortsCheckbox)

# Add checkbox for removing printer drivers
$removeDriversCheckbox = New-Object System.Windows.Forms.CheckBox
$removeDriversCheckbox.Text = "Remove Printer Drivers"
$removeDriversCheckbox.Location = New-Object System.Drawing.Point(440, 10)
$removeDriversCheckbox.Size = New-Object System.Drawing.Size(170, 20)
$form.Controls.Add($removeDriversCheckbox)

# Create update drivers button
$updateDriversButton = New-Object System.Windows.Forms.Button
$updateDriversButton.Location = New-Object System.Drawing.Point(730, 8)
$updateDriversButton.Size = New-Object System.Drawing.Size(150, 25)
$updateDriversButton.Text = "Update Printer Drivers"
$updateDriversButton.BackColor = [System.Drawing.Color]::LightBlue
$updateDriversButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$updateDriversButton.Add_Click({
    $updateResult = [System.Windows.Forms.MessageBox]::Show(
        "This will search Windows Update for printer drivers and may take 5+ minutes to complete.`n`nContinue?", 
        "Update Printer Drivers", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo, 
        [System.Windows.Forms.MessageBoxIcon]::Warning)
        
    if ($updateResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        $statusLabel.Text = "Updating printer drivers from Windows Update. This may take several minutes..."
        $form.Refresh()
        
        # Create a progress window
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Updating Printer Drivers"
        $progressForm.Size = New-Object System.Drawing.Size(400, 150)
        $progressForm.StartPosition = "CenterScreen"
        $progressForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $progressForm.MaximizeBox = $false
        
        $progressLabel = New-Object System.Windows.Forms.Label
        $progressLabel.Location = New-Object System.Drawing.Point(10, 20)
        $progressLabel.Size = New-Object System.Drawing.Size(360, 40)
        $progressLabel.Text = "Searching Windows Update for printer drivers...`nThis may take 5+ minutes to complete."
        $progressForm.Controls.Add($progressLabel)
        
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(150, 70)
        $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
        $cancelButton.Text = "Close"
        $cancelButton.Add_Click({ $progressForm.Close() })
        $progressForm.Controls.Add($cancelButton)
        
        # Start the update process in the background
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.ApartmentState = "STA"
        $runspace.ThreadOptions = "ReuseThread"
        $runspace.Open()
        
        $updateScript = {
            $drivers = Get-PrinterDriver
            
            # Command to update drivers via Windows Update
            # This can be done via an elevated PowerShell command
            $updateCommand = "pnputil.exe /scan-devices"
            $updateProcess = Start-Process -FilePath "powershell.exe" -ArgumentList "-Command $updateCommand" -Verb RunAs -PassThru
            $updateProcess.WaitForExit()
            
            # Another option is to use Windows Update directly
            $windowsUpdateCommand = "wuauclt.exe /detectnow /updatenow"
            $updateProcess = Start-Process -FilePath "powershell.exe" -ArgumentList "-Command $windowsUpdateCommand" -Verb RunAs -PassThru
            $updateProcess.WaitForExit()
        }
        
        $psCmd = [powershell]::Create().AddScript($updateScript)
        $psCmd.Runspace = $runspace
        
        # Show progress form before starting the task
        $progressForm.Show()
        
        # Start the task
        $handle = $psCmd.BeginInvoke()
        
        # When the form closes, we should clean up
        $progressForm.Add_FormClosed({
            if (-not $handle.IsCompleted) {
                # The user closed the form before the task completed
                # We'll let the task continue in the background
                $statusLabel.Text = "Driver update running in the background. Please wait before making changes."
            } else {
                $statusLabel.Text = "Printer driver update completed."
                # Refresh the printer list
                $script:SystemPrinters = @()
                $script:UserPrinters = @()
                Populate-PrinterList -ListView $listView
            }
        })
    }
})
$form.Controls.Add($updateDriversButton)

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
$removeButton.BackColor = [System.Drawing.Color]::LightCoral
$removeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
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
