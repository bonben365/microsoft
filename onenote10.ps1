$uri = 'https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/Files/Microsoft.Office.OneNote_16001.14326.21146.0.AppxBundle'
(New-Object Net.WebClient).DownloadFile($uri, "$env:temp\Microsoft.Office.OneNote_16001.14326.21146.0.AppxBundle")
Add-AppxPackage -Path "$env:temp\Microsoft.Office.OneNote_16001.14326.21146.0.AppxBundle"
