# COM Port System Tray Manager
# Save as: COMPortTrayManager.ps1
# Run with: powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "COMPortTrayManager.ps1"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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
    
    # Get all COM ports
    $serialPorts = [System.IO.Ports.SerialPort]::GetPortNames() | Sort-Object
    
    foreach ($port in $serialPorts) {
        $portInfo = Get-WmiObject -Query "SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%$port%'" | Select-Object -First 1
        $inUse = $false
        
        # Try to check if port is in use
        try {
            $testPort = New-Object System.IO.Ports.SerialPort $port
            $testPort.Open()
            $testPort.Close()
            $testPort.Dispose()
        } catch {
            $inUse = $true
        }
        
        $ports += [PSCustomObject]@{
            Name = $port
            Description = if ($portInfo) { $portInfo.Name } else { "Unknown Device" }
            InUse = $inUse
        }
    }
    
    return $ports
}

# Function to forcibly close a COM port
function Close-COMPort {
    param($portName)
    
    try {
        # Get processes that might be using COM ports
        $processes = Get-Process | Where-Object {
            $_.Modules | Where-Object { $_.ModuleName -match "System.IO.Ports" }
        }
        
        # More aggressive: Find processes with handles to the port
        $signature = @'
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool CloseHandle(IntPtr hObject);
'@
        
        Add-Type -MemberDefinition $signature -Name Win32Utils -Namespace Win32
        
        # Use handle64.exe - check multiple locations
        $handleExe = $null
        $possiblePaths = @(
            # "handle64.exe",  # If in PATH
            # "$env:WINDIR\System32\handle64.exe",
            # "C:\Tools\handle64.exe",
            # "C:\ProgramData\Tools\handle64.exe",
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
                    $handle = $matches[2]
                    
                    try {
                        $process = Get-Process -Id $_pid -ErrorAction SilentlyContinue
                        if ($process) {
                            [System.Windows.Forms.MessageBox]::Show(
                                "Process '$($process.Name)' (PID: $_pid) is using $portName. Close it?",
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
    $contextMenu.Items.Clear()
    
    # Add header
    $header = New-Object System.Windows.Forms.ToolStripMenuItem
    $header.Text = "COM Ports"
    $header.Enabled = $false
    $contextMenu.Items.Add($header)
    
    $contextMenu.Items.Add("-")
    
    # Get COM port info
    $ports = Get-COMPortInfo
    
    if ($ports.Count -eq 0) {
        $noPortsItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $noPortsItem.Text = "No COM ports found"
        $noPortsItem.Enabled = $false
        $contextMenu.Items.Add($noPortsItem)
    } else {
        foreach ($port in $ports) {
            $portItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $portItem.Text = "$($port.Name) - $($port.Description)"
            if ($port.InUse) {
                $portItem.Text += " [IN USE]"
                $portItem.ForeColor = [System.Drawing.Color]::Red
            } else {
                $portItem.ForeColor = [System.Drawing.Color]::Green
            }
            
            # Add submenu for actions
            $resetItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $resetItem.Text = "Reset/Close Port"
            $resetItem.Tag = $port.Name
            $resetItem.Add_Click({
                Close-COMPort -portName $this.Tag
                Update-Menu
            })
            
            $portItem.DropDownItems.Add($resetItem)
            $contextMenu.Items.Add($portItem)
        }
    }
    
    $contextMenu.Items.Add("-")
    
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
    
    $contextMenu.Items.Add("-")
    
    # Add exit option
    $exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exitItem.Text = "Exit"
    $exitItem.Add_Click({
        $trayIcon.Visible = $false
        $appContext.ExitThread()
    })
    $contextMenu.Items.Add($exitItem)
}

# Set up tray icon
$trayIcon.ContextMenuStrip = $contextMenu

# Handle double-click
$trayIcon.Add_DoubleClick({ Update-Menu; $contextMenu.Show() })

# Initial menu update
Update-Menu

# Timer to refresh port list periodically
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000 # Refresh every 1 second
$timer.Add_Tick({ Update-Menu })
$timer.Start()

# Run the application
[System.Windows.Forms.Application]::Run($appContext)

# Cleanup
$timer.Stop()
$timer.Dispose()
$trayIcon.Dispose()