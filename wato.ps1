<# Write-Host ========================================================================================
Write-Host Description:    Activating Microsoft software products for FREE without additional software
Write-Host Website:        https://msgang.com
Write-Host Script by:      Leo Nguyen
Write-Host For detailed script execution: https://msgang.com/windows
Write-Host ======================================================================================== #>

if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "irm msgang.com/win | iex"
    break
}

$edition = (Get-CimInstance Win32_OperatingSystem).Caption
$editionx = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName

Write-Host "`nYou're using: $edition" -ForegroundColor Green
Write-Host "Activating the product..." -ForegroundColor Green  

If ($edition -eq 'Microsoft Windows 11 Home' -or $edition -eq 'Microsoft Windows 10 Home') {$productkey = 'TX9XD-98N7V-6WMQ6-BX7FG-H8Q99'}
If ($edition -eq 'Microsoft Windows 11 Home N' -or $edition -eq 'Microsoft Windows 10 Home N') {$productkey = '3KHY7-WNT83-DGQKR-F7HPR-844BM'}
If ($edition -eq 'Microsoft Windows 11 Home Single Language' -or $edition -eq 'Microsoft Windows 10 Home Single Language') {$productkey = '7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH'}

If ($edition -eq 'Microsoft Windows 11 Pro' -or $edition -eq 'Microsoft Windows 10 Pro') {$productkey = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'}
If ($edition -eq 'Microsoft Windows 11 Pro N' -or $edition -eq 'Microsoft Windows 10 Pro N') {$productkey = 'MH37W-N47XK-V7XM9-C7227-GCQG9'}
If ($edition -eq 'Microsoft Windows 11 Pro for Workstations' -or $edition -eq 'Microsoft Windows 10 Pro for Workstations') {$productkey = 'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J'}
If ($edition -eq 'Microsoft Windows 11 Pro N for Workstations' -or $edition -eq 'Microsoft Windows 10 Pro N for Workstations') {$productkey = '9FNHH-K3HBT-3W4TD-6383H-6XYWF'}

If ($edition -eq 'Microsoft Windows 11 Enterprise' -or $edition -eq 'Microsoft Windows 10 Enterprise') {$productkey = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'}
If ($edition -eq 'Microsoft Windows 11 Enterprise N' -or $edition -eq 'Microsoft Windows 10 Enterprise N') {$productkey = 'DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4'}

If ($edition -eq 'Microsoft Windows 11 Education' -or $edition -eq 'Microsoft Windows 10 Education') {$productkey = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2'}
If ($edition -eq 'Microsoft Windows 11 Education N' -or $edition -eq 'Microsoft Windows 10 Education N') {$productkey = '2WH4N-8QGBV-H22JP-CT43Q-MDWWJ'}

If ($edition -eq 'Microsoft Windows 11 Pro Education' -or $edition -eq 'Microsoft Windows 10 Pro Education') {$productkey = '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y'}
If ($edition -eq 'Microsoft Windows 11 Pro N Education' -or $edition -eq 'Microsoft Windows 10 Pro N Education') {$productkey = 'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC'}

If ($edition -eq 'Microsoft Windows Server 2012 Standard') {$productkey = 'BN3D2-R7TKB-3YPBD-8DRP2-27GG4'}
If ($edition -eq 'Microsoft Windows Server 2012 Datacenter') {$productkey = '48HP8-DN98B-MYWDG-T2DCC-8W83P'}
If ($edition -eq 'Microsoft Windows Server 2012 R2 Standard') {$productkey = 'D2N9P-3P6X9-2R39C-7RTCD-MDVJX'}
If ($edition -eq 'Microsoft Windows Server 2012 R2 Datacenter') {$productkey = 'W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9'}
If ($edition -eq 'Microsoft Windows Server 2012 R2 Essentials') {$productkey = 'KNC87-3J2TX-XB4WP-VCPJV-M4FWM'}

If ($edition -eq 'Microsoft Windows Server 2016 Standard') {$productkey = 'WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY'}
If ($edition -eq 'Microsoft Windows Server 2016 Datacenter') {$productkey = 'CB7KF-BWN84-R7R2Y-793K2-8XDDG'}
If ($edition -eq 'Microsoft Windows Server 2016 Essentials') {$productkey = 'JCKRF-N37P4-C2D82-9YXRT-4M63B'}

If ($edition -eq 'Microsoft Windows Server 2019 Standard') {$productkey = 'N69G4-B89J2-4G8F4-WWYCC-J464C'}
If ($edition -eq 'Microsoft Windows Server 2019 Datacenter') {$productkey = 'WMDGN-G9PQG-XVVXX-R3X43-63DFG'}
If ($edition -eq 'Microsoft Windows Server 2019 Essentials') {$productkey = 'WVDHN-86M7X-466P6-VHXV7-YY726'}

If ($edition -eq 'Microsoft Windows Server 2022 Standard') {$productkey = 'VDYBN-27WPP-V4HQT-9VMD4-VMK7H'}
If ($edition -eq 'Microsoft Windows Server 2022 Datacenter') {$productkey = 'WX4NM-KYWYW-QJJR4-XV3QB-6VM33'}

If ($edition -eq 'Microsoft Windows Server 2008 Standard') {$productkey = 'TM24T-X9RMF-VWXK6-X8JC9-BFGM2'}
If ($edition -eq 'Microsoft Windows Server 2008 Enterprise') {$productkey = 'YQGMW-MPWTJ-34KDK-48M3W-X4Q6V'}
If ($edition -eq 'Microsoft Windows Server 2008 Datacenter') {$productkey = '7M67G-PC374-GR742-YH8V4-TCBY3'}

If ($edition -eq 'Microsoft Windows Server 2008 R2 Standard') {$productkey = 'YC6KT-GKW9T-YTKYR-T4X34-R7VHC'}
If ($edition -eq 'Microsoft Windows Server 2008 R2 Enterprise') {$productkey = '489J6-VHDMP-X63PK-3K798-CPX3Y'}
If ($edition -eq 'Microsoft Windows Server 2008 R2 Datacenter') {$productkey = '74YFP-3QFB3-KQT8W-PMXWJ-7M648'}

If ($edition -eq 'Microsoft Windows Server 2012') {$productkey = 'BN3D2-R7TKB-3YPBD-8DRP2-27GG4'}
If ($edition -eq 'Microsoft Windows Server 2012 Essentials') {$productkey = 'HTDQM-NBMMG-KGYDT-2DTKT-J2MPV'}
If ($edition -eq 'Microsoft Windows Server 2012 Standard') {$productkey = 'XC9B7-NBPP2-83J2H-RHMBY-92BT4'}
If ($edition -eq 'Microsoft Windows Server 2012 Datacenter') {$productkey = '48HP8-DN98B-MYWDG-T2DCC-8W83P'}

If ($edition -eq 'Microsoft Windows Server 2012 R2 Essentials') {$productkey = 'KNC87-3J2TX-XB4WP-VCPJV-M4FWM'}
If ($edition -eq 'Microsoft Windows Server 2012 R2 Standard') {$productkey = 'D2N9P-3P6X9-2R39C-7RTCD-MDVJX'}
If ($edition -eq 'Microsoft Windows Server 2012 R2 Datacenter') {$productkey = 'W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9'}
If ($edition -eq 'Microsoft Windows Server 2012 R2 Essentials') {$productkey = 'KNC87-3J2TX-XB4WP-VCPJV-M4FWM'}

If ($edition -eq 'Microsoft Windows 8.1 Pro') {$productkey = 'GCRJD-8NW9H-F2CDX-CCM8D-9D6T9'}
If ($edition -eq 'Microsoft Windows 8.1 Pro N') {$productkey = 'HMCNV-VVBFX-7HMBH-CTY9B-B4FXY'}
If ($edition -eq 'Microsoft Windows 8.1 Enterprise') {$productkey = 'MHF9N-XY6XB-WVXMC-BTDCT-MKKG7'}
If ($edition -eq 'Microsoft Windows 8.1 Enterprise N') {$productkey = 'TT4HM-HN7YT-62K67-RGRQJ-JFFXW'}

If ($edition -eq 'Microsoft Windows 8 Pro') {$productkey = 'NG4HW-VH26C-733KW-K6F98-J8CK4'}
If ($edition -eq 'Microsoft Windows 8 Pro N') {$productkey = 'XCVCF-2NXM9-723PB-MHCB7-2RYQQ'}
If ($edition -eq 'Microsoft Windows 8 Enterprise') {$productkey = '	32JNW-9KQ84-P47T8-D8GGY-CWCK7'}
If ($edition -eq 'Microsoft Windows 8 Enterprise N') {$productkey = 'JMNMF-RHW7P-DMY6X-RF3DR-X2BQT'}

If ($editionx -eq 'Microsoft Windows 10 Enterprise LTSC 2019' -or $editionx -eq 'Microsoft Windows 10 Enterprise LTSC 2021') {$productkey = 'M7XTQ-FN8P6-TTKYV-9D4CC-J462D'}
If ($editionx -eq 'Microsoft Windows 10 Enterprise LTSB 2016') {$productkey = 'DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ'}
If ($editionx -eq 'Microsoft Windows 10 Enterprise 2015 LTSB') {$productkey = 'WNMTR-4C88C-JK8YV-HQ7T2-76DF9'}
If ($editionx -eq 'Microsoft Windows 10 Enterprise Evaluation' -or $editionx -eq 'Microsoft Windows 10 Enterprise Evaluation') {$productkey = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'}

If ($editionx -eq 'Windows 7 Professional') {$productkey = 'FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4'}
If ($editionx -eq 'Windows 7 Enterprise') {$productkey = '33PXH-7Y6KF-2VJC9-XBBR8-HVTHH'}

&$env:windir\system32\cscript.exe $env:windir\system32\slmgr.vbs /ckms | Out-Null
&$env:windir\system32\cscript.exe $env:windir\system32\slmgr.vbs /ipk "$productkey" | Out-Null
&$env:windir\system32\cscript.exe $env:windir\system32\slmgr.vbs /skms kms.msgang.com | Out-Null
&$env:windir\system32\cscript.exe $env:windir\system32\slmgr.vbs /ato | Out-Null

Write-Host "Product activation successful.`n" -ForegroundColor Green
Write-Host "License status:" -ForegroundColor Yellow

if ($edition){

    Write-Host "Windows edition: $edition" -ForegroundColor Cyan

} else {

    Write-Host "Windows edition: $editionx" -ForegroundColor Cyan
}

$command = "&$env:windir\system32\cscript.exe $env:windir\system32\slmgr.vbs /dlv"

$status = Invoke-Expression -Command $command

Write-Host "$($status | Select-String -SimpleMatch "Product Key Channel")" -ForegroundColor Cyan
Write-Host "$($status | Select-String -SimpleMatch "License Status")" -ForegroundColor Cyan

Write-Host "(*)Visit https://msgang.com for more products.`n" -ForegroundColor Cyan

Start-Sleep -Seconds 3
start ms-settings:activation
