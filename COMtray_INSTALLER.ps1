# Setup script for COM Port Manager
# Run this once to set up automatic startup

$scriptPath = Join-Path $PSScriptRoot "COMPortTrayManager.ps1"
$startupFolder = [Environment]::GetFolderPath("Startup")
$shortcutPath = Join-Path $startupFolder "COM Port Manager.lnk"

# Create WScript Shell object
$WshShell = New-Object -ComObject WScript.Shell

# Create shortcut
$shortcut = $WshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
$shortcut.WorkingDirectory = $PSScriptRoot
$shortcut.Description = "COM Port Manager - System Tray Utility"
$shortcut.WindowStyle = 7  # Minimized
$shortcut.Save()

Write-Host "Startup shortcut created at: $shortcutPath" -ForegroundColor Green
Write-Host "COM Port Manager will now start automatically with Windows." -ForegroundColor Green
Write-Host ""
Write-Host "To remove from startup, delete the shortcut at:" -ForegroundColor Yellow
Write-Host $shortcutPath -ForegroundColor Yellow
Write-Host ""
Write-Host "To create a custom icon:" -ForegroundColor Cyan
Write-Host "1. Place a file named 'comport.ico' in the same folder as the script" -ForegroundColor Cyan
Write-Host "2. The icon will be automatically loaded on next run" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to start COM Port Manager now..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# Start the manager
Start-Process "powershell.exe" -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`"" -WindowStyle Hidden