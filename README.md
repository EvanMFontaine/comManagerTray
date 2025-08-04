# COM Port Tray Manager

A lightweight Windows system tray application for managing COM ports, monitoring serial devices, and providing quick access to PlatformIO device monitoring.

![COM Port Manager](data/image.png)

## Features

- **System Tray Integration**: Runs quietly in the background with a system tray icon
- **Real-time COM Port Detection**: Automatically detects and lists all available COM ports
- **Port Status Monitoring**: Shows whether ports are available or in use
- **PlatformIO Integration**: Quick access to PlatformIO's device monitor with customizable settings
- **Detailed Port Information**: View comprehensive device details including USB VID/PID, drivers, and hardware information
- **Port Reset/Close Functionality**: Force close stuck ports or reset devices
- **Smart Process Detection**: Uses Microsoft Sysinternals `handle64.exe` to identify processes using COM ports

## Requirements

### System Requirements

- **Windows 11** (required)
- PowerShell 5.1 or later
- Administrator privileges (recommended for full functionality)

### Software Dependencies

- **Visual Studio Code** with **PlatformIO Extension** installed
- **PlatformIO Core** (automatically installed with the VS Code extension)

### Optional Tools

- **handle64.exe** from Microsoft Sysinternals (for advanced process detection)
  - Download from: [Microsoft Sysinternals Handle](https://docs.microsoft.com/en-us/sysinternals/downloads/handle)
  - Place in one of these locations:
    - Same folder as the script
    - `C:\Windows\System32\`
    - Any folder in your PATH

## Installation

1. **Clone or download** this repository to your desired location
2. **Ensure dependencies** are installed:
   - Install VS Code from [Visual Studio Code](https://code.visualstudio.com/)
   - Install the PlatformIO extension in VS Code
   - (Optional) Download and place `handle64.exe` for enhanced functionality

3. **Set execution policy** (if needed):

   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

## Usage

### Starting the Application

#### Method 1: Direct PowerShell execution

```powershell
powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File "COMPortTrayManager.ps1"
```

#### Method 2: Using the provided installer

Run `COMtray_INSTALLER.ps1` to set up automatic startup

#### Method 3: Manual execution

Double-click `COMPortTrayManager.ps1` and select "Run with PowerShell"

### Using the Application

1. **Locate the tray icon**: Look for the COM port icon in your system tray (bottom-right corner)

2. **View COM ports**:
   - Left-click or right-click the tray icon
   - The menu will show all detected COM ports with their status:
     - 🟢 **Green**: Port is available
     - 🔴 **Red**: Port is in use

3. **Available Actions**:
   - **Open Monitor**: Launch PlatformIO device monitor with customizable settings
   - **Get Info**: View detailed port and device information
   - **Reset/Close Port**: Force close or reset a stuck port

### PlatformIO Monitor Features

When opening a monitor, you can configure:

- **Baud Rate**: 9600, 115200, 256000, 500000, 921600, 1000000, 1500000, 2000000, 3000000, 6000000
- **Data Bits**: 5, 6, 7, 8
- **Parity**: None (N), Even (E), Odd (O), Space (S), Mark (M)
- **Stop Bits**: 1, 1.5, 2
- **Filters**: colorize, debug, direct, hexlify, time, and more

## File Structure

```text
comManagerTray/
├── COMPortTrayManager.ps1     # Main application script
├── COMtray_INSTALLER.ps1      # Installation script
├── COMtray_UNINSTALLER.ps1    # Uninstallation script
├── comport.ico                # Application icon
├── handle64.exe               # Optional: Sysinternals handle tool
├── data/                      # Additional resources
│   ├── comport - Copy.ico
│   ├── iconFixer.py
│   ├── image.png
│   └── serial_port_icon.ico
└── iconbs/                    # Icon development files
```

## Troubleshooting

### Common Issues

1. **"PlatformIO not found" error**:
   - Ensure VS Code with PlatformIO extension is installed
   - Check that PlatformIO Core is properly installed
   - Restart VS Code and let PlatformIO initialize

2. **Script won't run**:
   - Check PowerShell execution policy: `Get-ExecutionPolicy`
   - Set appropriate policy: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`
   - Run as Administrator if needed

3. **No COM ports detected**:
   - Ensure devices are properly connected
   - Check Device Manager for COM port listings
   - Restart the application

4. **Can't reset/close ports**:
   - Download and install `handle64.exe` for better process detection
   - Run the application as Administrator
   - Check if antivirus is blocking the operation

### Performance Notes

- The application uses efficient .NET System.IO.Ports for port detection
- Menu updates are triggered by user interaction, not continuous polling
- Minimal system resource usage when running in background

## Advanced Features

### Handle64.exe Integration

When `handle64.exe` is available, the application can:

- Identify specific processes using COM ports
- Show process names and PIDs
- Offer to close processes cleanly before resetting ports

### Registry Integration

The application can read and display:

- Current port settings from Windows registry
- Device driver information
- Hardware identification details
- USB Vendor ID (VID) and Product ID (PID)

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve this tool.

## License

This project is provided as-is for educational and utility purposes.

## Acknowledgments

- Uses Microsoft Sysinternals Handle utility for enhanced functionality
- Integrates with PlatformIO for professional serial monitoring
- Built with PowerShell and .NET Windows Forms
