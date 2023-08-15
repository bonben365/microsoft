#Write-Host ***************************************************************
#Write-Host Name:           Windows Download PowerShell Script
#Write-Host Description:    Download the lastest versions.
#Write-Host Version:        1.0
#Write-Host Date :          26/7/2023
#Write-Host Website:        https://bonguides.com
#Write-Host Script by:      Leo Nguyen
#Write-Host For detailed script execution: https://bonguides.com
#Write-Host ***************************************************************

if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}

New-Item -Path $env:temp\tmp -ItemType Directory -Force | Out-Null
Set-Location $env:temp\tmp
Invoke-Item $env:temp\tmp
Set-ExecutionPolicy Bypass Process

Invoke-RestMethod 'https://raw.githubusercontent.com/pbatard/Fido/master/Fido.ps1' -OutFile "$env:temp\tmp\fido.ps1"

Write-Host
.\fido.ps1 -win list

Write-Host
Write-Host "*Instructions.................................................."    -ForegroundColor Green
Write-Host
Write-Host "- Donwload Windows 10:        .\fido.ps1 -win 10"                   -ForegroundColor Yellow
Write-Host "- Donwload Windows 11:        .\fido.ps1 -win 11"                   -ForegroundColor Yellow
Write-Host
Write-Host "- Get supported languages:    .\fido.ps1 -win 10 -lang list"        -ForegroundColor Yellow
Write-Host "- Download custom language:   .\fido.ps1 -win 10 -lang Spanish"     -ForegroundColor Yellow
Write-Host
Write-Host "- Get download URL:           .\fido.ps1 -win 10 -GetURL"           -ForegroundColor Yellow
Write-Host
Write-host "*.............................................................."    -ForegroundColor Green
