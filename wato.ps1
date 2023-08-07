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

#$edition = (Get-CimInstance Win32_OperatingSystem).Caption
$edition = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
Write-Host '---------------------------------------------------------------'
Write-Host "You're using $edition"                 
Write-Host '---------------------------------------------------------------'

If ($edition -eq 'Windows 11 Home' -or $edition -eq 'Windows 10 Home') {$productkey = 'TX9XD-98N7V-6WMQ6-BX7FG-H8Q99'}
If ($edition -eq 'Windows 11 Home N' -or $edition -eq 'Windows 10 Home N') {$productkey = '3KHY7-WNT83-DGQKR-F7HPR-844BM'}
If ($edition -eq 'Windows 11 Home Single Language' -or $edition -eq 'Windows 10 Home Single Language') {$productkey = '7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH'}

If ($edition -eq 'Windows 11 Pro' -or $edition -eq 'Windows 10 Pro') {$productkey = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'}
If ($edition -eq 'Windows 11 Pro N' -or $edition -eq 'Windows 10 Pro N') {$productkey = 'MH37W-N47XK-V7XM9-C7227-GCQG9'}
If ($edition -eq 'Windows 11 Pro for Workstations' -or $edition -eq 'Windows 10 Pro for Workstations') {$productkey = 'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J'}
If ($edition -eq 'Windows 11 Pro N for Workstations' -or $edition -eq 'Windows 10 Pro N for Workstations') {$productkey = '9FNHH-K3HBT-3W4TD-6383H-6XYWF'}

If ($edition -eq 'Windows 11 Enterprise' -or $edition -eq 'Windows 10 Enterprise') {$productkey = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'}
If ($edition -eq 'Windows 11 Enterprise N' -or $edition -eq 'Windows 10 Enterprise N') {$productkey = 'DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4'}

If ($edition -eq 'Windows 11 Education' -or $edition -eq 'Windows 10 Education') {$productkey = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2'}
If ($edition -eq 'Windows 11 Education N' -or $edition -eq 'Windows 10 Education N') {$productkey = '2WH4N-8QGBV-H22JP-CT43Q-MDWWJ'}

If ($edition -eq 'Windows 11 Pro Education' -or $edition -eq 'Windows 10 Pro Education') {$productkey = '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y'}
If ($edition -eq 'Windows 11 Pro N Education' -or $edition -eq 'Windows 10 Pro N Education') {$productkey = 'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC'}

If ($edition -eq 'Windows Server 2012 Standard') {$productkey = 'BN3D2-R7TKB-3YPBD-8DRP2-27GG4'}
If ($edition -eq 'Windows Server 2012 Datacenter') {$productkey = '48HP8-DN98B-MYWDG-T2DCC-8W83P'}
If ($edition -eq 'Windows Server 2012 R2 Standard') {$productkey = 'D2N9P-3P6X9-2R39C-7RTCD-MDVJX'}
If ($edition -eq 'Windows Server 2012 R2 Datacenter') {$productkey = 'W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9'}
If ($edition -eq 'Windows Server 2012 R2 Essentials') {$productkey = 'KNC87-3J2TX-XB4WP-VCPJV-M4FWM'}

If ($edition -eq 'Windows Server 2016 Standard') {$productkey = 'WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY'}
If ($edition -eq 'Windows Server 2016 Datacenter') {$productkey = 'CB7KF-BWN84-R7R2Y-793K2-8XDDG'}
If ($edition -eq 'Windows Server 2016 Essentials') {$productkey = 'JCKRF-N37P4-C2D82-9YXRT-4M63B'}

If ($edition -eq 'Windows Server 2019 Standard') {$productkey = 'N69G4-B89J2-4G8F4-WWYCC-J464C'}
If ($edition -eq 'Windows Server 2019 Datacenter') {$productkey = 'WMDGN-G9PQG-XVVXX-R3X43-63DFG'}
If ($edition -eq 'Windows Server 2019 Essentials') {$productkey = 'WVDHN-86M7X-466P6-VHXV7-YY726'}

If ($edition -eq 'Windows Server 2022 Standard') {$productkey = 'VDYBN-27WPP-V4HQT-9VMD4-VMK7H'}
If ($edition -eq 'Windows Server 2022 Datacenter') {$productkey = 'WX4NM-KYWYW-QJJR4-XV3QB-6VM33'}

If ($edition -eq 'Windows 7 Professional') {$productkey = 'FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4'}
If ($edition -eq 'Windows 7 Enterprise') {$productkey = '33PXH-7Y6KF-2VJC9-XBBR8-HVTHH'}

If ($edition -eq 'Windows 10 Enterprise LTSC 2019' -or $edition -eq 'Windows 10 Enterprise LTSC 2021') {$productkey = 'M7XTQ-FN8P6-TTKYV-9D4CC-J462D'}


cscript $env:windir\system32\slmgr.vbs /ckms | Out-Null
cscript $env:windir\system32\slmgr.vbs /ipk "$productkey" | Out-Null
cscript $env:windir\system32\slmgr.vbs /skms kms.msgang.com | Out-Null
cscript $env:windir\system32\slmgr.vbs /ato

Write-Host
Write-Host "Done............"
Write-Host
Write-Host "Your Windows edition: $edition" -ForegroundColor Yellow
$command = "cscript $env:windir\system32\slmgr.vbs /dlv"
$status = Invoke-Expression -Command $command
Write-Host "$($status | Select-String -SimpleMatch "Product Key Channel")" -ForegroundColor Yellow
Write-Host "$($status | Select-String -SimpleMatch "License Status")" -ForegroundColor Yellow
Write-Host "Visit https://msang.com for more products."
