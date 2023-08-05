
Write-Host ===============================================================
Write-Host Name:           Windows Activator
Write-Host Description:    Activate all Windows Editions for free.
Write-Host Version:        1.0
Write-Host Date :          26/7/2023
Write-Host Website:        https://msgang.com
Write-Host Script by:      Leo Nguyen
Write-Host For detailed script execution: https://msgang.com/windows
Write-Host ===============================================================

if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}

$ospp = Resolve-Path -Path "C:\Program Files*\Microsoft Office\Office16\ospp.vbs" | Select-Object -ExpandProperty Path -Last 1
#$ospp = (Resolve-Path -Path "C:\Program Files*\Microsoft Office\Office16\ospp.vbs").Path

cscript $ospp /dstatus
$productkeys = @('NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP';'6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK';'B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B';'C4F7P-NCP8C-6CQPT-MQHV9-JXD2M';'9BGNQ-K37YR-RQHF2-38RQ3-7VCBB';'7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2';'9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT';'TMJWT-YYNMB-3BKTF-644FC-RVXBD';'7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK';'RRNCX-C64HY-W2MM7-MCH9G-TJHMQ';'G2KWX-3NW6P-PY93R-JXK2T-C9Y9V';'NCJ33-JHBBY-HTK98-MYCV8-HMKHJ';'PBX3G-NWMT6-Q7XBW-PYJGG-WXD33')
foreach ($productkey in $productkeys) {
    cscript $ospp /inpkey:$productkey
}
cscript $ospp /sethst:kms.msgang.com
cscript $ospp /act
pause


