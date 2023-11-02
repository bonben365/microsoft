if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}
#Find OSPP.vbs path and run the command with the dstatus option (Last 1...)
$OSPP = Resolve-Path -Path "C:\Program Files*\Microsoft Office\Office16\ospp.vbs" | Select-Object -ExpandProperty Path -Last 1
Write-Host "OSPP Location is: $OSPP" -ForegroundColor Green
$Command = "cscript.exe //nologo '$OSPP' /dstatus"
$DStatus = Invoke-Expression -Command $Command
#Get product keys from OSPP.vbs output.
$ProductKeys = $DStatus | Select-String -SimpleMatch "Last 5" | ForEach-Object -Process { $_.tostring().split(" ")[-1]}
if ($ProductKeys) {
    Write-Host "Found $(($ProductKeys | Measure-Object).Count) productkeys, proceeding with deactivation..." -ForegroundColor Green
    #Run OSPP.vbs per key with /unpkey option.
    foreach ($ProductKey in $ProductKeys) {
        Write-Host "Processing productkey $ProductKey" -ForegroundColor Green
        $Command = "cscript.exe //nologo '$OSPP' /unpkey:$ProductKey"
        $Command = "cscript.exe //nologo '$OSPP' /remhst"
        Invoke-Expression -Command $Command
    }
} else {
    Write-Host "Found no keys to remove... " -ForegroundColor Green
}
