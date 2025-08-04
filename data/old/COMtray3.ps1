# COM Port System Tray Manager
# Save as: COMPortTrayManager.ps1
# Run with: powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "COMPortTrayManager.ps1"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Ports

# Create application context
$appContext = New-Object System.Windows.Forms.ApplicationContext

# Create NotifyIcon
$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Text = "COM Port Manager"
$trayIcon.Icon = [System.Drawing.SystemIcons]::Information
$trayIcon.Visible = $true

# Create context menu
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Function to get COM ports and their status
function Get-COMPortInfo {
    $ports = @()
    
    try {
        # Method 1: Using System.IO.Ports
        $serialPorts = [System.IO.Ports.SerialPort]::GetPortNames() | Sort-Object
        Write-Host "Found ports via System.IO.Ports: $($serialPorts -join ', ')"
        
        # If no ports found, try WMI
        if ($serialPorts.Count -eq 0) {
            Write-Host "No ports found via System.IO.Ports, trying WMI..."
            $wmiPorts = Get-WmiObject Win32_SerialPort | Select-Object -ExpandProperty DeviceID
            $serialPorts = $wmiPorts
        }
        
        # Also check Win32_PnPEntity for all COM ports
        $pnpPorts = Get-WmiObject Win32_PnPEntity | Where-Object { 
            $_.Name -match 'COM\d+' -or $_.Caption -match 'COM\d+'
        } | ForEach-Object {
            if ($_.Name -match 'COM(\d+)') { "COM$($matches[1])" }
            elseif ($_.Caption -match 'COM(\d+)') { "COM$($matches[1])" }
        } | Select-Object -Unique | Sort-Object
        
        Write-Host "Found ports via PnP: $($pnpPorts -join ', ')"
        
        # Combine all discovered ports
        $allPorts = ($serialPorts + $pnpPorts) | Select-Object -Unique | Sort-Object
        
        foreach ($port in $allPorts) {
            if ($port) {
                $portInfo = Get-WmiObject -Query "SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%$port%'" | Select-Object -First 1
                $inUse = $false
                $status = "Available"
                
                # Try to check if port is in use
                try {
                    $testPort = New-Object System.IO.Ports.SerialPort $port
                    $testPort.Open()
                    $testPort.Close()
                    $testPort.Dispose()
                } catch {
                    $inUse = $true
                    $status = "In Use"
                }
                
                $description = if ($portInfo) { 
                    # Extract the device name, removing redundant COM port info
                    $deviceName = $portInfo.Name
                    if ($deviceName -match "\($port\)") {
                        $deviceName = $deviceName -replace "\s*\($port\)\s*", " "
                        $description = "$port - $($deviceName.Trim())"
                    } else {
                        $description = "$port - $deviceName"
                    }
                    $description
                } else { 
                    "$port"
                }
                
                # Clean up description
                if ($description.Length -gt 60) {
                    $description = $description.Substring(0, 57) + "..."
                }
                
                $ports += [PSCustomObject]@{
                    Name = $port
                    Description = $description
                    InUse = $inUse
                    Status = $status
                }
                
                Write-Host "Added port: $port - $status - Description: $description"
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
            $handles = & $handleExe -a -nobanner "\Device\$portName" 2>$null
            
            foreach ($line in $handles) {
                if ($line -match "pid: (\d+).*handle: ([0-9A-F]+)") {
                    $_pid = $matches[1]
                    $_handle = $matches[2]
                    
                    try {
                        $_process = Get-Process -Id $_pid -ErrorAction SilentlyContinue
                        if ($_process) {
                            $result = [System.Windows.Forms.MessageBox]::Show(
                                "Process '$($_process.Name)' (PID: $_pid) is using $portName. Close it?",
                                "COM Port In Use",
                                [System.Windows.Forms.MessageBoxButtons]::YesNo
                            )
                            
                            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                                Stop-Process -Id $_pid -Force
                            }
                        }
                    } catch {}
                }
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

# Function to refresh menu
function Update-Menu {
    Write-Host "`nUpdating menu..."
    $contextMenu.Items.Clear()
    
    # Add header
    $header = New-Object System.Windows.Forms.ToolStripMenuItem
    $header.Text = "COM Ports"
    $header.Enabled = $false
    $header.Font = New-Object System.Drawing.Font($header.Font, [System.Drawing.FontStyle]::Bold)
    $contextMenu.Items.Add($header)
    
    $contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    
    # Get COM port info
    $ports = Get-COMPortInfo
    
    if ($ports.Count -eq 0) {
        $noPortsItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $noPortsItem.Text = "No COM ports found"
        $noPortsItem.Enabled = $false
        $contextMenu.Items.Add($noPortsItem)
        Write-Host "No COM ports found"
    } else {
        Write-Host "Adding $($ports.Count) ports to menu"
        foreach ($port in $ports) {
            $portItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $displayText = "$($port.Name) - $($port.Status)"
            
            if ($port.Description -ne $port.Name) {
                $displayText = $port.Description
            }
            
            $portItem.Text = $displayText
            
            if ($port.InUse) {
                $portItem.ForeColor = [System.Drawing.Color]::Red
            } else {
                $portItem.ForeColor = [System.Drawing.Color]::Green
            }
            
            # Add submenu for actions
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
            
            $portItem.DropDownItems.Add($resetItem)
            $portItem.DropDownItems.Add($copyItem)
            $contextMenu.Items.Add($portItem)
        }
    }
    
    $contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    
    # Add refresh option
    $refreshItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $refreshItem.Text = "Refresh"
    $refreshItem.Add_Click({ Update-Menu })
    $contextMenu.Items.Add($refreshItem)
    
    # Add Device Manager option
    $devMgrItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $devMgrItem.Text = "Open Device Manager"
    $devMgrItem.Add_Click({ Start-Process "devmgmt.msc" })
    $contextMenu.Items.Add($devMgrItem)
    
    # Add Mode command option (shows COM port settings)
    $modeItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $modeItem.Text = "Show Port Settings (MODE)"
    $modeItem.Add_Click({ 
        Start-Process "cmd.exe" -ArgumentList "/k mode"
    })
    $contextMenu.Items.Add($modeItem)
    
    $contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    
    # Add exit option
    $exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exitItem.Text = "Exit"
    $exitItem.Add_Click({
        $trayIcon.Visible = $false
        $appContext.ExitThread()
    })
    $contextMenu.Items.Add($exitItem)
    
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