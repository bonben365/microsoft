Write-Host ===============================================================
Write-Host Name:           Windows Activator
Write-Host Description:    Activate all Windows Edition for free.
Write-Host Version:        1.0
Write-Host Date :          26/7/2023
Write-Host Website:        https://msgang.com
Write-Host Script by:      Leo Nguyen
Write-Host For detailed script execution: https://msgang.com/wato
Write-Host ===============================================================

$edition = (Get-CimInstance Win32_OperatingSystem).Caption
Write-Host '--------------------------------------------------------------'
Write-Host "You're using $edition"                 
Write-Host '--------------------------------------------------------------'

If ($edition -eq 'Microsoft Windows 11 Pro' -or $edition -eq 'Microsoft Windows 10 Pro') {$productkey = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'}
If ($edition -eq 'Microsoft Windows 11 Enterprise' -or $edition -eq 'Microsoft Windows 10 Enterprise') {$productkey = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'}
If ($edition -eq 'Microsoft Windows 11 Education' -or $edition -eq 'Microsoft Windows 10 Education') {$productkey = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2'}
If ($edition -eq 'Microsoft Windows 11 Education' -or $edition -eq 'Microsoft Windows 10 Education') {$productkey = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2'}

If ($edition -eq 'Microsoft Windows 11 Education' -or $edition -eq 'Microsoft Windows 10 Home') {$productkey = 'TX9XD-98N7V-6WMQ6-BX7FG-H8Q99'}





cscript $env:windir\system32\slmgr.vbs /ckms
cscript $env:windir\system32\slmgr.vbs /ipk "$productkey"
cscript $env:windir\system32\slmgr.vbs /skms kms.msgang.com
cscript $env:windir\system32\slmgr.vbs /ato
