# Uninstall script for COM Port Manager
# Run this to remove COM Port Manager from startup and stop running instances

Write-Host "COM Port Manager Uninstaller" -ForegroundColor Cyan
Write-Host "=============================" -ForegroundColor Cyan
Write-Host ""

$startupFolder = [Environment]::GetFolderPath("Startup")
$shortcutPath = Join-Path $startupFolder "COM Port Manager.lnk"
$removed = $false

# Remove startup shortcut
if (Test-Path $shortcutPath) {
    try {
        Remove-Item $shortcutPath -Force
        Write-Host "✓ Startup shortcut removed successfully" -ForegroundColor Green
        $removed = $true
    }
    catch {
        Write-Host "✗ Failed to remove startup shortcut: $($_.Exception.Message)" -ForegroundColor Red
    }
}
else {
    Write-Host "i Startup shortcut not found (may already be removed)" -ForegroundColor Yellow
}

# Stop any running instances of the COM Port Manager
Write-Host "Checking for running instances..." -ForegroundColor Cyan
$processes = Get-Process -Name "powershell" -ErrorAction SilentlyContinue | Where-Object {
    $_.CommandLine -like "*COMPortTrayManager.ps1*"
}

if ($processes) {
    foreach ($process in $processes) {
        try {
            $process.Kill()
            Write-Host "✓ Stopped running COM Port Manager process (PID: $($process.Id))" -ForegroundColor Green
            $removed = $true
        }
        catch {
            Write-Host "✗ Failed to stop process (PID: $($process.Id)): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
} else {
    Write-Host "ℹ No running instances found" -ForegroundColor Yellow
}

# Alternative method: Kill by window title (if the tray app sets a specific title)
try {
    $comPortProcesses = Get-Process | Where-Object { $_.MainWindowTitle -like "*COM Port*" -or $_.ProcessName -eq "COMPortTrayManager" }
    foreach ($process in $comPortProcesses) {
        $process.Kill()
        Write-Host "✓ Stopped COM Port Manager process: $($process.ProcessName)" -ForegroundColor Green
        $removed = $true
        }
}
catch {
    # Silently continue if no processes found
}

Write-Host ""
if ($removed) {
    Write-Host "COM Port Manager has been successfully uninstalled!" -ForegroundColor Green
    Write-Host "• Removed from Windows startup" -ForegroundColor Green
    Write-Host "• Stopped any running instances" -ForegroundColor Green
} else {
    Write-Host "COM Port Manager appears to already be uninstalled." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Note: The script files remain in this folder and can be manually deleted if desired." -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")