<#
    Deletes the Startup shortcut and terminates any running tray instances.
#>

$startup  = [Environment]::GetFolderPath('Startup')
$lnk      = Join-Path $startup 'COM Port Manager.lnk'

if (Test-Path $lnk) {
    Remove-Item $lnk -Force
    Write-Host "✔ Startup shortcut removed."
} else {
    Write-Host "ℹ Shortcut already absent."
}

# kill running tray scripts
$procs = Get-Process -Name powershell -ErrorAction SilentlyContinue |
         ? { $_.CommandLine -like '*COMPortTrayManager.ps1*' }

if ($procs) {
    $procs | % { $_.Kill() }
    Write-Host "✔ Stopped $($procs.Count) running tray instance(s)."
} else {
    Write-Host "ℹ No tray instances found."
}
