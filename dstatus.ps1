if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}

$path = "C:\Program Files*\Microsoft Office\Office1*\ospp.vbs"
$ospp = Resolve-Path -Path $path | Select-Object -ExpandProperty Path -Last 1
cscript //nologo $ospp /dstatus
