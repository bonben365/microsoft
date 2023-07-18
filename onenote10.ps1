#Install and update Desktop framework packages
Add-AppxPackage "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"

#Download and Install OneNote for Windows 10
$uri = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/Files/Microsoft.Office.OneNote_16001.14326.21146.0.AppxBundle'
Write-Host
Write-Host 'Downloading Microsoft OneNote Package...'
(New-Object Net.WebClient).DownloadFile($uri, "$env:temp\Microsoft.Office.OneNote_16001.14326.21146.0.AppxBundle")
Add-AppxPackage -Path "$env:temp\Microsoft.Office.OneNote_16001.14326.21146.0.AppxBundle"
Write-Host
Write-Host 'Installed Package:'
Get-AppxPackage -Name *Microsoft.Office.OneNote*
