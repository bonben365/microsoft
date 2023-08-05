$edition = (Get-CimInstance Win32_OperatingSystem).Caption
Write-Host "You're using $edition"


