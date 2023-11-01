if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}

$path64 = "C:\Program Files\Microsoft Office\Office1*"
$path32 = "C:\Program Files (x86)\Microsoft Office\Office1*"
if ((Test-Path -Path "$path32\ospp.vbs")) { Set-Location $path32 -ErrorAction SilentlyContinue }
if ((Test-Path -Path "$path64\ospp.vbs")) { Set-Location $path64 -ErrorAction SilentlyContinue }
