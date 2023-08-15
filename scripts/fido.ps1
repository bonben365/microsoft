#Write-Host ===============================================================
#Write-Host Name:           Windows Download PowerShell Script
#Write-Host Description:    Download the lastest versions.
#Write-Host Version:        1.0
#Write-Host Date :          26/7/2023
#Write-Host Website:        https://bonguides.com
#Write-Host Script by:      Leo Nguyen
#Write-Host For detailed script execution: https://bonguides.com
#Write-Host ===============================================================

if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}

New-Item -Path $env:temp\tmp -ItemType Directory -Force | Out-Null
Set-Location $env:temp\tmp
Invoke-Item $env:temp\tmp
Set-ExecutionPolicy Bypass Process

Invoke-RestMethod 'https://raw.githubusercontent.com/pbatard/Fido/master/Fido.ps1' -OutFile "$env:temp\tmp\fido.ps1"

.\fido.ps1 -win list

Write-Host
Write-Host "Donwload Windows 10: .\fido.ps1 -win 10"
Write-Host "Donwload Windows 11: .\fido.ps1 -win 11"
