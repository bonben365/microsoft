$edition = (Get-CimInstance Win32_OperatingSystem).Caption
Write-Host "You're using $edition"
If ($edition -eq 'Microsoft Windows 11 Pro' -or $edition -eq 'Microsoft Windows 10 Pro') {
  cscript $env:windir\system32\slmgr.vbs /ckms
  cscript $env:windir\system32\slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
  cscript $env:windir\system32\slmgr.vbs /skms kms.msgang.com
  cscript $env:windir\system32\slmgr.vbs /ato
}


