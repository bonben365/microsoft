Write-Host ===============================================================
Write-Host Name:           Windows Activator
Write-Host Description:    Activate all Windows Editions for free.
Write-Host Version:        1.0
Write-Host Date :          26/7/2023
Write-Host Website:        https://msgang.com
Write-Host Script by:      Leo Nguyen
Write-Host For detailed script execution: https://msgang.com/windows
Write-Host ===============================================================

$edition = (Get-CimInstance Win32_OperatingSystem).Caption
Write-Host '---------------------------------------------------------------'
Write-Host "You're using $edition"                 
Write-Host '---------------------------------------------------------------'

If ($edition -eq 'Microsoft Windows 11 Home' -or $edition -eq 'Microsoft Windows 10 Home') {$productkey = 'TX9XD-98N7V-6WMQ6-BX7FG-H8Q99'}
If ($edition -eq 'Microsoft Windows 11 Home N' -or $edition -eq 'Microsoft Windows 10 Home N') {$productkey = '3KHY7-WNT83-DGQKR-F7HPR-844BM'}
If ($edition -eq 'Microsoft Windows 11 Home Single Language' -or $edition -eq 'Microsoft Windows 10 Home Single Language') {$productkey = '7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH'}

If ($edition -eq 'Microsoft Windows 11 Pro' -or $edition -eq 'Microsoft Windows 10 Pro') {$productkey = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'}
If ($edition -eq 'Microsoft Windows 11 Pro N' -or $edition -eq 'Microsoft Windows 10 Pro N') {$productkey = 'MH37W-N47XK-V7XM9-C7227-GCQG9'}
If ($edition -eq 'Microsoft Windows 11 Pro for Workstations' -or $edition -eq 'Microsoft Windows 10 Pro for Workstations') {$productkey = 'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J'}
If ($edition -eq 'Microsoft Windows 11 Pro for Workstations N' -or $edition -eq 'Microsoft Windows 10 Pro for Workstations N') {$productkey = '9FNHH-K3HBT-3W4TD-6383H-6XYWF'}

If ($edition -eq 'Microsoft Windows 11 Enterprise' -or $edition -eq 'Microsoft Windows 10 Enterprise') {$productkey = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'}
If ($edition -eq 'Microsoft Windows 11 Enterprise N' -or $edition -eq 'Microsoft Windows 10 Enterprise N') {$productkey = 'DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4'}

If ($edition -eq 'Microsoft Windows 11 Education' -or $edition -eq 'Microsoft Windows 10 Education') {$productkey = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2'}
If ($edition -eq 'Microsoft Windows 11 Education N' -or $edition -eq 'Microsoft Windows 10 Education N') {$productkey = '2WH4N-8QGBV-H22JP-CT43Q-MDWWJ'}

If ($edition -eq 'Microsoft Windows 11 Pro Education' -or $edition -eq 'Microsoft Windows 10 Pro Education') {$productkey = '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y'}
If ($edition -eq 'Microsoft Windows 11 Pro Education N' -or $edition -eq 'Microsoft Windows 10 Pro Education N') {$productkey = 'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC'}


If ($edition -eq 'Microsoft Windows Server 2016 Standard') {$productkey = 'WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY'}
If ($edition -eq 'Microsoft Windows Server 2016 Datacenter') {$productkey = 'CB7KF-BWN84-R7R2Y-793K2-8XDDG'}
If ($edition -eq 'Microsoft Windows Server 2016 Essentials') {$productkey = 'JCKRF-N37P4-C2D82-9YXRT-4M63B'}

If ($edition -eq 'Microsoft Windows Server 2019 Standard') {$productkey = 'N69G4-B89J2-4G8F4-WWYCC-J464C'}
If ($edition -eq 'Microsoft Windows Server 2019 Datacenter') {$productkey = 'WMDGN-G9PQG-XVVXX-R3X43-63DFG'}
If ($edition -eq 'Microsoft Windows Server 2019 Essentials') {$productkey = 'WVDHN-86M7X-466P6-VHXV7-YY726'}

If ($edition -eq 'Microsoft Windows Server 2022 Standard') {$productkey = 'VDYBN-27WPP-V4HQT-9VMD4-VMK7H'}
If ($edition -eq 'Microsoft Windows Server 2022 Datacenter') {$productkey = 'WX4NM-KYWYW-QJJR4-XV3QB-6VM33'}



cscript $env:windir\system32\slmgr.vbs /ckms
cscript $env:windir\system32\slmgr.vbs /ipk "$productkey"
cscript $env:windir\system32\slmgr.vbs /skms kms.msgang.com
cscript $env:windir\system32\slmgr.vbs /ato
