Write-Host ==================================================================================================================
Write-Host Name:           Windows Activator
Write-Host Description:    Activate all Windows Edition for free.
Write-Host Version:        1.0
Write-Host Date :          26/7/2023
Write-Host Website:        https://msgang.com
Write-Host Script by:      Leo Nguyen
Write-Host For detailed script execution: https://msgang.com/wato
Write-Host =================================================================================================================

$edition = (Get-CimInstance Win32_OperatingSystem).Caption
Write-Host '-----------------------------------'
Write-Host "You're using $edition"
Write-Host '-----------------------------------'
If ($edition -eq 'Microsoft Windows 11 Pro' -or $edition -eq 'Microsoft Windows 10 Pro') {
  cscript $env:windir\system32\slmgr.vbs /ckms
  cscript $env:windir\system32\slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
  cscript $env:windir\system32\slmgr.vbs /skms kms.msgang.com
  cscript $env:windir\system32\slmgr.vbs /ato
}

If ($edition -eq 'Microsoft Windows 11 Enterprise' -or $edition -eq 'Microsoft Windows 10 Enterprise') {
  cscript $env:windir\system32\slmgr.vbs /ckms
  cscript $env:windir\system32\slmgr.vbs /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43
  cscript $env:windir\system32\slmgr.vbs /skms kms.msgang.com
  cscript $env:windir\system32\slmgr.vbs /ato
}

If ($edition -eq 'Microsoft Windows 11 Education' -or $edition -eq 'Microsoft Windows 10 Education') {
  cscript $env:windir\system32\slmgr.vbs /ckms
  cscript $env:windir\system32\slmgr.vbs /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
  cscript $env:windir\system32\slmgr.vbs /skms kms.msgang.com
  cscript $env:windir\system32\slmgr.vbs /ato
}
