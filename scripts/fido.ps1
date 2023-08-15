Write-Host ===============================================================
Write-Host Name:           Windows Download PowerShell Script
Write-Host Description:    Download the lastest versions.
Write-Host Version:        1.0
Write-Host Date :          26/7/2023
Write-Host Website:        https://bonguides.com
Write-Host Script by:      Leo Nguyen
Write-Host For detailed script execution: https://bonguides.com
Write-Host ===============================================================

if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}

New-Item -Path $env:temp\tmp -ItemType Directory -Force | Out-Null
Set-Location $env:temp\tmp
