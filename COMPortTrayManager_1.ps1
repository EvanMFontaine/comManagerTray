# COM Port System Tray Manager
# Save as: COMPortTrayManager.ps1
# Run with: powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "COMPortTrayManager.ps1"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create application context
$appContext = New-Object System.Windows.Forms.ApplicationContext


# Create NotifyIcon
$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Text = 'COM Port Manager'

# --- Load local comport.ico if it exists ---
$iconPath = Join-Path $PSScriptRoot 'comport.ico'
try {
    if (Test-Path $iconPath) {
        $trayIcon.Icon = New-Object System.Drawing.Icon $iconPath   # use your .ico
    } else {
        $trayIcon.Icon = [System.Drawing.SystemIcons]::Information  # fallback
    }
} catch {
    # corrupted/invalid .ico → fall back quietly
    $trayIcon.Icon = [System.Drawing.SystemIcons]::Information
}

$trayIcon.Visible = $true

# Create context menu
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Function to get COM ports and their status
function Get-COMPortInfo {
    $ports = @()
    
    try {
        # Just use System.IO.Ports - it's the fastest method
        $serialPorts = [System.IO.Ports.SerialPort]::GetPortNames() | Sort-Object { 
            # Extract the numeric part for proper sorting
            if ($_ -match 'COM(\d+)') { [int]$matches[1] } else { $_ }
        }
        Write-Host "Found ports: $($serialPorts -join ', ')"
        
        foreach ($port in $serialPorts) {
            if ($port) {
                $inUse = $false
                $status = "Available"
                
                # Try to check if port is in use (this is fast)
                try {
                    $testPort = New-Object System.IO.Ports.SerialPort $port
                    $testPort.Open()
                    $testPort.Close()
                    $testPort.Dispose()
                } catch {
                    $inUse = $true
                    $status = "In Use"
                }
                
                $ports += [PSCustomObject]@{
                    Name = $port
                    Description = $port  # Simple description for now
                    InUse = $inUse
                    Status = $status
                    DeviceInfo = $null   # Will be loaded on demand
                }
                
                Write-Host "Added port: $port - $status"
            }
        }
    } catch {
        Write-Host "Error getting COM ports: $_"
    }
    
    return $ports
}

# Function to forcibly close a COM port
function Close-COMPort {
    param($portName)
    
    try {
        # Use handle64.exe - check multiple locations
        $handleExe = $null
        $possiblePaths = @(
            "handle64.exe",  # If in PATH
            "$env:WINDIR\System32\handle64.exe",
            "C:\Tools\handle64.exe",
            "C:\ProgramData\Tools\handle64.exe",
            "$PSScriptRoot\handle64.exe"  # Same folder as script
        )
        
        foreach ($path in $possiblePaths) {
            if (Test-Path $path) {
                $handleExe = $path
                break
            }
        }
        
        if ($handleExe) {
            Write-Host "Using handle64.exe to find processes using $portName"
            
            # First, try to force close the port by resetting the device
            try {
                # Try to identify what's using the port by checking all handles
                $allHandles = & $handleExe -a -nobanner 2>$null | Out-String
                
                # Look for references to the COM port
                $portReferences = $allHandles -split "`n" | Where-Object { 
                    $_ -match $portName -or $_ -match "Serial" -or $_ -match "\\Device\\VCP"
                }
                
                $foundProcess = $false
                $processesUsingPort = @()
                
                foreach ($line in $portReferences) {
                    if ($line -match "^(\S+)\s+pid:\s*(\d+)") {
                        $processName = $matches[1]
                        $_pid = $matches[2]
                        
                        # Check if this process actually has the port open
                        try {
                            $_process = Get-Process -Id $_pid -ErrorAction SilentlyContinue
                            if ($_process) {
                                $processesUsingPort += [PSCustomObject]@{
                                    Name = $_process.Name
                                    PID = $_pid
                                }
                                $foundProcess = $true
                            }
                        } catch {}
                    }
                }
                
                if ($foundProcess) {
                    $processList = ($processesUsingPort | ForEach-Object { "$($_.Name) (PID: $($_.PID))" }) -join "`n"
                    $result = [System.Windows.Forms.MessageBox]::Show(
                        "The following processes may be using $portName`:`n`n$processList`n`nClose all these processes?",
                        "COM Port In Use",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Question
                    )
                    
                    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                        foreach ($proc in $processesUsingPort) {
                            try {
                                Stop-Process -Id $proc.PID -Force
                                Write-Host "Closed process $($proc.Name) (PID: $($proc.PID))"
                            } catch {
                                Write-Host "Failed to close process $($proc.Name): $_"
                            }
                        }
                    }
                } else {
                    # If handle64 can't find it, try the device reset approach
                    $result = [System.Windows.Forms.MessageBox]::Show(
                        "Could not identify the specific process using $portName.`n`nWould you like to reset the port device instead?",
                        "COM Port Reset",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Question
                    )
                    
                    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                        # Reset the device
                        Write-Host "Attempting to reset $portName device..."
                        $device = Get-PnpDevice | Where-Object { $_.FriendlyName -match $portName }
                        if ($device) {
                            Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                            Start-Sleep -Milliseconds 500
                            Enable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                            
                            [System.Windows.Forms.MessageBox]::Show(
                                "$portName has been reset. You may need to restart applications using this port.",
                                "Port Reset Complete",
                                [System.Windows.Forms.MessageBoxButtons]::OK,
                                [System.Windows.Forms.MessageBoxIcon]::Information
                            )
                        }
                    }
                }
            } catch {
                Write-Host "Error using handle64.exe: $_"
                [System.Windows.Forms.MessageBox]::Show(
                    "Error checking port usage: $_",
                    "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        } else {
            # Basic reset without handle64.exe
            [System.Windows.Forms.MessageBox]::Show(
                "Port reset attempted. For better functionality, download handle64.exe from Microsoft Sysinternals and place it in C:\Windows\System32 or add to PATH.",
                "COM Port Reset",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        
        # Reset the port in Device Manager
        $device = Get-PnpDevice | Where-Object { $_.FriendlyName -match $portName }
        if ($device) {
            Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 500
            Enable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
        }
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error resetting $portName`: $_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# Function to get USB info for a specific port
function Get-USBInfo {
    param($portName)
    
    try {
        $portInfo = Get-WmiObject Win32_PnPEntity | Where-Object { $_.Name -match $portName } | Select-Object -First 1
        
        if ($portInfo) {
            # Extract device description
            $deviceName = $portInfo.Name
            if ($deviceName -match "\($portName\)") {
                $deviceName = $deviceName -replace "\s*\($portName\)\s*", " "
            }
            
            # Get USB info
            $hardwareID = (Get-PnpDeviceProperty -InstanceId $portInfo.DeviceID -KeyName "DEVPKEY_Device_HardwareIds" -ErrorAction SilentlyContinue).Data | Select-Object -First 1
            $manufacturer = (Get-PnpDeviceProperty -InstanceId $portInfo.DeviceID -KeyName "DEVPKEY_Device_Manufacturer" -ErrorAction SilentlyContinue).Data
            
            $_vid = $null
            $_pid = $null
            if ($hardwareID -match 'VID_([0-9A-F]+)&PID_([0-9A-F]+)') {
                $_vid = $matches[1]
                $_pid = $matches[2]
            }
            
            return [PSCustomObject]@{
                Description = "$portName - $($deviceName.Trim())"
                VID = $_vid
                PID = $_pid
                Manufacturer = $manufacturer
            }
        }
    } catch {
        Write-Host "Error getting USB info for $portName`: $_"
    }
    
    return $null
}
function Open-COMPort {
    param($portName)

    # Create form for settings
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Open $portName - Serial Monitor Settings"
    $form.Size = New-Object System.Drawing.Size(420, 560)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false

    # Baud rate
    $lblBaud = New-Object System.Windows.Forms.Label
    $lblBaud.Text = "Baud Rate:"
    $lblBaud.Location = New-Object System.Drawing.Point(20, 20)
    $lblBaud.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($lblBaud)

    $cmbBaud = New-Object System.Windows.Forms.ComboBox
    $cmbBaud.Location = New-Object System.Drawing.Point(130, 20)
    $cmbBaud.Size = New-Object System.Drawing.Size(240, 20)
    $cmbBaud.DropDownStyle = "DropDown"
    $cmbBaud.Items.AddRange(@("9600", "115200", "256000", "500000", "921600", "1000000", "1500000", "2000000", "3000000", "6000000"))
    $cmbBaud.Text = "921600"
    $form.Controls.Add($cmbBaud)

    # Data bits
    $lblData = New-Object System.Windows.Forms.Label
    $lblData.Text = "Data Bits:"
    $lblData.Location = New-Object System.Drawing.Point(20, 60)
    $lblData.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($lblData)

    $cmbData = New-Object System.Windows.Forms.ComboBox
    $cmbData.Location = New-Object System.Drawing.Point(130, 60)
    $cmbData.Size = New-Object System.Drawing.Size(100, 20)
    $cmbData.DropDownStyle = "DropDownList"
    $cmbData.Items.AddRange(@("5", "6", "7", "8"))
    $cmbData.SelectedItem = "8"
    $form.Controls.Add($cmbData)

    # Parity
    $lblParity = New-Object System.Windows.Forms.Label
    $lblParity.Text = "Parity:"
    $lblParity.Location = New-Object System.Drawing.Point(20, 100)
    $lblParity.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($lblParity)

    $cmbParity = New-Object System.Windows.Forms.ComboBox
    $cmbParity.Location = New-Object System.Drawing.Point(130, 100)
    $cmbParity.Size = New-Object System.Drawing.Size(100, 20)
    $cmbParity.DropDownStyle = "DropDownList"
    $cmbParity.Items.AddRange(@("N", "E", "O", "S", "M"))
    $cmbParity.SelectedItem = "N"
    $form.Controls.Add($cmbParity)

    # Stop bits
    $lblStop = New-Object System.Windows.Forms.Label
    $lblStop.Text = "Stop Bits:"
    $lblStop.Location = New-Object System.Drawing.Point(20, 140)
    $lblStop.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($lblStop)

    $cmbStop = New-Object System.Windows.Forms.ComboBox
    $cmbStop.Location = New-Object System.Drawing.Point(130, 140)
    $cmbStop.Size = New-Object System.Drawing.Size(100, 20)
    $cmbStop.DropDownStyle = "DropDownList"
    $cmbStop.Items.AddRange(@("1", "1.5", "2"))
    $cmbStop.SelectedItem = "1"
    $form.Controls.Add($cmbStop)

    # Monitor mode (base filter)
    $grpMode = New-Object System.Windows.Forms.GroupBox
    $grpMode.Text = "Monitor mode (base filter)"
    $grpMode.Location = New-Object System.Drawing.Point(20, 180)
    $grpMode.Size = New-Object System.Drawing.Size(370, 140)
    $form.Controls.Add($grpMode)

    $rbDirect = New-Object System.Windows.Forms.RadioButton
    $rbDirect.Text = "Direct (raw, keeps CR/ANSI)"
    $rbDirect.Location = New-Object System.Drawing.Point(12, 22)
    $rbDirect.Size = New-Object System.Drawing.Size(250, 20)
    $rbDirect.Checked = $true
    $grpMode.Controls.Add($rbDirect)

    $rbDefault = New-Object System.Windows.Forms.RadioButton
    $rbDefault.Text = "Default (strips control codes)"
    $rbDefault.Location = New-Object System.Drawing.Point(12, 45)
    $rbDefault.Size = New-Object System.Drawing.Size(250, 20)
    $grpMode.Controls.Add($rbDefault)

    $rbPrintable = New-Object System.Windows.Forms.RadioButton
    $rbPrintable.Text = "Printable (drops non-printables)"
    $rbPrintable.Location = New-Object System.Drawing.Point(12, 68)
    $rbPrintable.Size = New-Object System.Drawing.Size(250, 20)
    $grpMode.Controls.Add($rbPrintable)

    $rbNoControl = New-Object System.Windows.Forms.RadioButton
    $rbNoControl.Text = "NoControl (drops control chars)"
    $rbNoControl.Location = New-Object System.Drawing.Point(12, 91)
    $rbNoControl.Size = New-Object System.Drawing.Size(250, 20)
    $grpMode.Controls.Add($rbNoControl)

    $rbHexlify = New-Object System.Windows.Forms.RadioButton
    $rbHexlify.Text = "Hexlify (show bytes as hex)"
    $rbHexlify.Location = New-Object System.Drawing.Point(12, 114)
    $rbHexlify.Size = New-Object System.Drawing.Size(250, 20)
    $grpMode.Controls.Add($rbHexlify)

    # Extras (additional filters/transforms)
    $grpExtras = New-Object System.Windows.Forms.GroupBox
    $grpExtras.Text = "Extras (optional)"
    $grpExtras.Location = New-Object System.Drawing.Point(20, 330)
    $grpExtras.Size = New-Object System.Drawing.Size(370, 100)
    $form.Controls.Add($grpExtras)

    $cbColorize = New-Object System.Windows.Forms.CheckBox
    $cbColorize.Text = "colorize"
    $cbColorize.Location = New-Object System.Drawing.Point(12, 22)
    $cbColorize.Size = New-Object System.Drawing.Size(100, 20)
    $grpExtras.Controls.Add($cbColorize)

    $cbDebug = New-Object System.Windows.Forms.CheckBox
    $cbDebug.Text = "debug"
    $cbDebug.Location = New-Object System.Drawing.Point(120, 22)
    $cbDebug.Size = New-Object System.Drawing.Size(80, 20)
    $grpExtras.Controls.Add($cbDebug)

    $cbTime = New-Object System.Windows.Forms.CheckBox
    $cbTime.Text = "time"
    $cbTime.Location = New-Object System.Drawing.Point(210, 22)
    $cbTime.Size = New-Object System.Drawing.Size(80, 20)
    $grpExtras.Controls.Add($cbTime)

    $cbSendOnEnter = New-Object System.Windows.Forms.CheckBox
    $cbSendOnEnter.Text = "send_on_enter"
    $cbSendOnEnter.Location = New-Object System.Drawing.Point(12, 48)
    $cbSendOnEnter.Size = New-Object System.Drawing.Size(150, 20)
    $grpExtras.Controls.Add($cbSendOnEnter)

    $cbLog2File = New-Object System.Windows.Forms.CheckBox
    $cbLog2File.Text = "log2file"
    $cbLog2File.Location = New-Object System.Drawing.Point(170, 48)
    $cbLog2File.Size = New-Object System.Drawing.Size(100, 20)
    $grpExtras.Controls.Add($cbLog2File)

    # Print chosen settings in terminal
    $cbPrintSelection = New-Object System.Windows.Forms.CheckBox
    $cbPrintSelection.Text = "Print selected settings when monitor opens"
    $cbPrintSelection.Location = New-Object System.Drawing.Point(20, 440)
    $cbPrintSelection.Size = New-Object System.Drawing.Size(370, 20)
    $cbPrintSelection.Checked = $true
    $form.Controls.Add($cbPrintSelection)

    # Buttons
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Open Monitor"
    $btnOK.Location = New-Object System.Drawing.Point(130, 470)
    $btnOK.Size = New-Object System.Drawing.Size(110, 30)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(250, 470)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)

    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $baudRate = $cmbBaud.Text
        $dataBits = $cmbData.SelectedItem
        $parity = $cmbParity.SelectedItem
        $stopBits = $cmbStop.SelectedItem

        # Pick base filter (always pass one explicitly so PlatformIO doesn't fall back to implicit defaults)
        $baseFilter = "direct"
        if ($rbDefault.Checked)   { $baseFilter = "default" }
        if ($rbPrintable.Checked) { $baseFilter = "printable" }
        if ($rbNoControl.Checked) { $baseFilter = "nocontrol" }
        if ($rbHexlify.Checked)   { $baseFilter = "hexlify" }

        $filters = New-Object System.Collections.Generic.List[string]
        $filters.Add($baseFilter)

        # Extras
        if ($cbColorize.Checked)    { $filters.Add("colorize") }
        if ($cbDebug.Checked)       { $filters.Add("debug") }
        if ($cbTime.Checked)        { $filters.Add("time") }
        if ($cbSendOnEnter.Checked) { $filters.Add("send_on_enter") }
        if ($cbLog2File.Checked)    { $filters.Add("log2file") }

        # Build PlatformIO command
        $pioPath = "$env:USERPROFILE\.platformio\penv\Scripts\platformio.exe"

        # Check if PlatformIO is installed
        if (-not (Test-Path $pioPath)) {
            # Try to find it in PATH
            $pioCmd = Get-Command "platformio" -ErrorAction SilentlyContinue
            if ($pioCmd) {
                $pioPath = "platformio"
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "PlatformIO not found. Please install PlatformIO or ensure it's in your PATH.",
                    "PlatformIO Not Found",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return
            }
        }

        # Build arguments for platformio.exe
        $args = New-Object System.Collections.Generic.List[string]
        $args.AddRange(@("device", "monitor", "--port", $portName, "--baud", $baudRate))

        # Add parity/data/stop bits if not default
        if ($dataBits -ne "8" -or $parity -ne "N" -or $stopBits -ne "1") {
            $args.AddRange(@("--parity", $parity, "--databits", $dataBits, "--stopbits", $stopBits))
        }

        # Add all selected filters (repeat --filter)
        foreach ($f in $filters) {
            $args.AddRange(@("--filter", $f))
        }

        # Pretty banner
        $filterStr = ($filters.ToArray() -join ", ")
        $uartFmt = "$baudRate $dataBits-$parity-$stopBits"

        # Launch in a new console and optionally print the chosen settings first
        # Using cmd.exe so we can echo before PlatformIO starts (and keep Ctrl+C behavior).
        $pioQuoted = '"' + $pioPath + '"'
        $argsJoined = ($args | ForEach-Object {
            if ($_ -match '\s') { '"' + $_ + '"' } else { $_ }
        }) -join ' '

        $cmdLine = ""
        if ($cbPrintSelection.Checked) {
            $cmdLine = "echo --- Serial Monitor: $portName [$uartFmt] & echo --- Filters: $filterStr & echo."
            $cmdLine += " & "
        }
        $cmdLine += "call $pioQuoted $argsJoined"

        Write-Host "Launching monitor: cmd.exe /k $cmdLine"
        Start-Process -FilePath "cmd.exe" -ArgumentList @("/k", $cmdLine)
    }

    $form.Dispose()
}

# Function to show detailed port information
function Show-PortInfo {
    param($portName)
    
    # Create form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$portName - Detailed Information"
    $form.Size = New-Object System.Drawing.Size(600, 500)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "Sizable"
    $form.MinimumSize = New-Object System.Drawing.Size(500, 400)
    
    # Create text box for info
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.ScrollBars = "Vertical"
    $textBox.ReadOnly = $true
    $textBox.Font = New-Object System.Drawing.Font("Consolas", 10)
    $textBox.Dock = "Fill"
    $textBox.BackColor = [System.Drawing.Color]::White
    
    # Create panel for buttons
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Height = 40
    $buttonPanel.Dock = "Bottom"
    
    # Copy button
    $copyButton = New-Object System.Windows.Forms.Button
    $copyButton.Text = "Copy All"
    $copyButton.Location = New-Object System.Drawing.Point(10, 8)
    $copyButton.Size = New-Object System.Drawing.Size(100, 25)
    $copyButton.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($textBox.Text)
        [System.Windows.Forms.MessageBox]::Show(
            "Information copied to clipboard!",
            "Copied",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })
    
    # Close button
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(120, 8)
    $closeButton.Size = New-Object System.Drawing.Size(100, 25)
    $closeButton.Add_Click({ $form.Close() })
    
    $buttonPanel.Controls.Add($copyButton)
    $buttonPanel.Controls.Add($closeButton)
    
    # Add controls
    $form.Controls.Add($textBox)
    $form.Controls.Add($buttonPanel)
    
    # Gather information
    $info = "=== COM PORT INFORMATION ===" + "`r`n"
    $info += "Port: $portName" + "`r`n"
    $info += "Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" + "`r`n"
    $info += "============================" + "`r`n`r`n"
    
    # Check if port is available
    try {
        $testPort = New-Object System.IO.Ports.SerialPort $portName
        $testPort.Open()
        $testPort.Close()
        $testPort.Dispose()
        $info += "Status: Available" + "`r`n"
    } catch {
        $info += "Status: In Use" + "`r`n"
        $info += "Error: $($_.Exception.Message)" + "`r`n"
    }
    
    $info += "`r`n--- DEVICE INFORMATION ---" + "`r`n"
    
    # Get detailed device info
    try {
        $portInfo = Get-WmiObject Win32_PnPEntity -Filter "Name LIKE '%$portName%'" | Select-Object -First 1
        
        if ($portInfo) {
            $info += "Device Name: $($portInfo.Name)" + "`r`n"
            $info += "Device ID: $($portInfo.DeviceID)" + "`r`n"
            $info += "Status: $($portInfo.Status)" + "`r`n"
            $info += "Class: $($portInfo.PNPClass)" + "`r`n"
            
            # Get all device properties
            $deviceProps = Get-PnpDeviceProperty -InstanceId $portInfo.DeviceID -ErrorAction SilentlyContinue
            
            $info += "`r`n--- USB INFORMATION ---" + "`r`n"
            
            # Hardware IDs
            $hardwareIds = ($deviceProps | Where-Object { $_.KeyName -eq "DEVPKEY_Device_HardwareIds" }).Data
            if ($hardwareIds) {
                $info += "Hardware IDs:" + "`r`n"
                foreach ($id in $hardwareIds) {
                    $info += "  $id" + "`r`n"
                    if ($id -match 'VID_([0-9A-F]+)&PID_([0-9A-F]+)') {
                        $info += "  -> VID: 0x$($matches[1]) ($([Convert]::ToInt32($matches[1], 16)))" + "`r`n"
                        $info += "  -> PID: 0x$($matches[2]) ($([Convert]::ToInt32($matches[2], 16)))" + "`r`n"
                    }
                }
            }
            
            # Manufacturer
            $manufacturer = ($deviceProps | Where-Object { $_.KeyName -eq "DEVPKEY_Device_Manufacturer" }).Data
            if ($manufacturer) {
                $info += "Manufacturer: $manufacturer" + "`r`n"
            }
            
            # Driver info
            $info += "`r`n--- DRIVER INFORMATION ---" + "`r`n"
            $driverVersion = ($deviceProps | Where-Object { $_.KeyName -eq "DEVPKEY_Device_DriverVersion" }).Data
            $driverDate = ($deviceProps | Where-Object { $_.KeyName -eq "DEVPKEY_Device_DriverDate" }).Data
            $driverProvider = ($deviceProps | Where-Object { $_.KeyName -eq "DEVPKEY_Device_DriverProvider" }).Data
            
            if ($driverVersion) { $info += "Driver Version: $driverVersion" + "`r`n" }
            if ($driverDate) { $info += "Driver Date: $driverDate" + "`r`n" }
            if ($driverProvider) { $info += "Driver Provider: $driverProvider" + "`r`n" }
            
            # Additional properties
            $info += "`r`n--- ADDITIONAL PROPERTIES ---" + "`r`n"
            $busReportedDesc = ($deviceProps | Where-Object { $_.KeyName -eq "DEVPKEY_Device_BusReportedDeviceDesc" }).Data
            $friendlyName = ($deviceProps | Where-Object { $_.KeyName -eq "DEVPKEY_Device_FriendlyName" }).Data
            $locationInfo = ($deviceProps | Where-Object { $_.KeyName -eq "DEVPKEY_Device_LocationInfo" }).Data
            
            if ($busReportedDesc) { $info += "Bus Reported Description: $busReportedDesc" + "`r`n" }
            if ($friendlyName) { $info += "Friendly Name: $friendlyName" + "`r`n" }
            if ($locationInfo) { $info += "Location: $locationInfo" + "`r`n" }
            
        } else {
            $info += "No device information found for $portName" + "`r`n"
        }
        
        # Get serial port settings from registry
        $info += "`r`n--- PORT SETTINGS (Registry) ---" + "`r`n"
        try {
            $regPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\$($portInfo.DeviceID)\Device Parameters"
            if (Test-Path $regPath) {
                $regSettings = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
                if ($regSettings.PortName) { $info += "Registry Port Name: $($regSettings.PortName)" + "`r`n" }
                if ($regSettings.LatencyTimer) { $info += "Latency Timer: $($regSettings.LatencyTimer)" + "`r`n" }
            }
        } catch {}
        
    } catch {
        $info += "Error retrieving device information: $_" + "`r`n"
    }
    
    # Get current port settings using MODE command
    $info += "`r`n--- CURRENT PORT SETTINGS ---" + "`r`n"
    try {
        $modeOutput = & mode $portName 2>&1
        if ($modeOutput -match "Access is denied") {
            $info += "Port is in use - cannot retrieve current settings via MODE" + "`r`n"
            
            # Try to get settings from registry for in-use ports
            $info += "`r`nAttempting to read default settings from registry..." + "`r`n"
            try {
                # Common registry locations for COM port settings
                $commPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports"
                $serialCommPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Serial"
                
                # Try to get default baud rate from device parameters
                if ($portInfo) {
                    $enumPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\$($portInfo.DeviceID)\Device Parameters"
                    if (Test-Path $enumPath) {
                        $params = Get-ItemProperty -Path $enumPath -ErrorAction SilentlyContinue
                        if ($params) {
                            if ($params.MaximumBaudRate) { $info += "Maximum Baud Rate: $($params.MaximumBaudRate)" + "`r`n" }
                            if ($params.LatencyTimer) { $info += "Latency Timer: $($params.LatencyTimer) ms" + "`r`n" }
                            if ($params.MinReadTimeout) { $info += "Min Read Timeout: $($params.MinReadTimeout) ms" + "`r`n" }
                            if ($params.MinWriteTimeout) { $info += "Min Write Timeout: $($params.MinWriteTimeout) ms" + "`r`n" }
                        }
                    }
                }
                
                $info += "`r`nNote: These are device defaults, not current active settings" + "`r`n"
            } catch {
                $info += "Could not retrieve registry settings" + "`r`n"
            }
        } else {
            $info += $modeOutput -join "`r`n"
        }
    } catch {
        $info += "Unable to retrieve port settings: $_" + "`r`n"
    }
    
    $textBox.Text = $info
    $form.ShowDialog()
    $form.Dispose()
}

# Function to refresh menu
function Update-Menu {
    Write-Host "`nUpdating menu..."
    $contextMenu.Items.Clear()
    
    # Add header
    $header = New-Object System.Windows.Forms.ToolStripMenuItem
    $header.Text = "COM Ports"
    $header.Enabled = $false
    $header.Font = New-Object System.Drawing.Font($header.Font, [System.Drawing.FontStyle]::Bold)
    $contextMenu.Items.Add($header) | Out-Null
    
    $contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
    
    # Get COM port info
    $ports = Get-COMPortInfo
    
    if ($ports.Count -eq 0) {
        $noPortsItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $noPortsItem.Text = "No COM ports found"
        $noPortsItem.Enabled = $false
        $contextMenu.Items.Add($noPortsItem) | Out-Null
        Write-Host "No COM ports found"
    } else {
        Write-Host "Adding $($ports.Count) ports to menu"
        foreach ($port in $ports) {
            $portItem = New-Object System.Windows.Forms.ToolStripMenuItem
            
            # Simple display - just port name and status
            $portItem.Text = $port.Name
            
            if ($port.InUse) {
                $portItem.ForeColor = [System.Drawing.Color]::Red
                $portItem.Text += " (In Use)"
            } else {
                $portItem.ForeColor = [System.Drawing.Color]::Green
                $portItem.Text += " (Available)"
            }
            
            # Add submenu for actions
            if (-not $port.InUse) {
                # Open Monitor option (only for available ports)
                $openItem = New-Object System.Windows.Forms.ToolStripMenuItem
                $openItem.Text = "Open Monitor"
                $openItem.Tag = $port.Name
                $openItem.Font = New-Object System.Drawing.Font($openItem.Font, [System.Drawing.FontStyle]::Bold)
                $openItem.Add_Click({
                    param($_sender, $e)
                    Open-COMPort -portName $_sender.Tag
                })
                $portItem.DropDownItems.Add($openItem) | Out-Null
                
                $portItem.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
            }
            
            # Get Info button
            $infoItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $infoItem.Text = "Get Info"
            $infoItem.Tag = $port.Name
            $infoItem.Add_Click({
                param($_sender, $e)
                Show-PortInfo -portName $_sender.Tag
            })
            
            $resetItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $resetItem.Text = "Reset/Close Port"
            $resetItem.Tag = $port.Name
            $resetItem.Add_Click({
                param($_sender, $e)
                Close-COMPort -portName $_sender.Tag
                Update-Menu
            })
            
            $copyItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $copyItem.Text = "Copy Port Name"
            $copyItem.Tag = $port.Name
            $copyItem.Add_Click({
                param($_sender, $e)
                [System.Windows.Forms.Clipboard]::SetText($_sender.Tag)
            })
            
            $portItem.DropDownItems.Add($infoItem) | Out-Null
            $portItem.DropDownItems.Add($resetItem) | Out-Null
            $portItem.DropDownItems.Add($copyItem) | Out-Null
            
            $contextMenu.Items.Add($portItem) | Out-Null
        }
    }
    
    $contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
    
    # Add exit option
    $exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exitItem.Text = "Exit"
    $exitItem.Add_Click({
        $trayIcon.Visible = $false
        $appContext.ExitThread()
    })
    $contextMenu.Items.Add($exitItem) | Out-Null
    
    Write-Host "Menu updated with $($contextMenu.Items.Count) items"
}

# Set up tray icon
$trayIcon.ContextMenuStrip = $contextMenu

# Handle clicks
$trayIcon.Add_MouseClick({
    param($_sender, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left -or $e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        Update-Menu
        $contextMenu.Show([System.Windows.Forms.Cursor]::Position)
    }
})

# Initial menu update
Update-Menu

# Run the application
Write-Host "COM Port Manager started. Look for the icon in your system tray."
[System.Windows.Forms.Application]::Run($appContext)

# Cleanup
$trayIcon.Dispose()