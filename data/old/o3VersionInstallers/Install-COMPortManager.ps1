<#
    Drops a shortcut into the current user's Startup folder that
    runs COMPortTrayManager.ps1 hidden and bypasses the execution-policy prompt.
#>

$trayScript   = Join-Path $PSScriptRoot 'COMPortTrayManager.ps1'
if (-not (Test-Path $trayScript)) { Write-Error "Tray script not found: $trayScript"; exit 1 }

$startup      = [Environment]::GetFolderPath('Startup')
$lnk          = Join-Path $startup 'COM Port Manager.lnk'

$ws           = New-Object -ComObject WScript.Shell
$sc           = $ws.CreateShortcut($lnk)
$sc.TargetPath      = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$sc.Arguments       = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$trayScript`""
$sc.WorkingDirectory= $PSScriptRoot
$sc.WindowStyle     = 7        # Minimized
$sc.IconLocation    = Join-Path $PSScriptRoot 'comport.ico'
$sc.Description     = 'COM Port Manager - system-tray utility'
$sc.Save()

Write-Host "✔ Installed.  Shortcut created in $startup.`nCOM Port Manager will launch automatically at log-in."
