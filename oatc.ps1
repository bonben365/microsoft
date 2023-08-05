
Write-Host ===============================================================
Write-Host Name:           Microsoft Office Activator
Write-Host Description:    Activate all Offices Editions for free.
Write-Host Version:        1.0
Write-Host Date :          26/7/2023
Write-Host Website:        https://msgang.com
Write-Host Script by:      Leo Nguyen
Write-Host For detailed script execution: https://msgang.com/office
Write-Host ===============================================================

if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}

#$ospp = (Resolve-Path -Path "C:\Program Files*\Microsoft Office\Office16\ospp.vbs").Path
#Find OSPP.vbs path and run the command with the dstatus option (Last 1...)
$ospp = Resolve-Path -Path "C:\Program Files*\Microsoft Office\Office16\ospp.vbs" | Select-Object -ExpandProperty Path -Last 1
Write-Output -InputObject "OSPP Location is: $OSPP"
$Command = "cscript.exe '$ospp' /dstatus"
$DStatus = Invoke-Expression -Command $Command
#Get product keys from OSPP.vbs output.
$keys = $DStatus | Select-String -SimpleMatch "Last 5" | ForEach-Object -Process { $_.tostring().split(" ")[-1]}
if ($keys) {
    Write-Output -InputObject "Found $(($ProductKeys | Measure-Object).Count) productkeys, proceeding with deactivation..."
    #Run OSPP.vbs per key with /unpkey option.
    foreach ($key in $keys) {
        Write-Output -InputObject "Processing productkey $key"
        $ospp /unpkey:$key
    }
}

$productkeys = @(
    'NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP';
    '6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK';
    'B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B';
    'C4F7P-NCP8C-6CQPT-MQHV9-JXD2M';
    '9BGNQ-K37YR-RQHF2-38RQ3-7VCBB';
    '7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2';
    '9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT';
    'TMJWT-YYNMB-3BKTF-644FC-RVXBD';
    '7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK';
    'RRNCX-C64HY-W2MM7-MCH9G-TJHMQ';
    'G2KWX-3NW6P-PY93R-JXK2T-C9Y9V';
    'NCJ33-JHBBY-HTK98-MYCV8-HMKHJ';
    'PBX3G-NWMT6-Q7XBW-PYJGG-WXD33';
    'FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH';
    'KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3';
    'FTNWT-C6WBT-8HMGF-K9PRX-QV9H8';
    'J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T';
    'KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4';
    'MJVNY-BYWPY-CWV6J-2RKRT-4M8QG';
    'WM8YG-YNGDD-4JHDC-PG3F4-FC4T4';
    'NWG3X-87C9K-TC7YY-BC2G7-G6RVC';
    'C9FM6-3N72F-HFJXB-TM3V9-T86R9';
    'TY7XF-NFRBR-KJ44C-G83KF-GX27K';
    '2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ';
    'HWCXN-K3WBT-WJBKY-R8BD9-XK29P';
    'TN8H9-M34D3-Y64V9-TR72V-X79KV')

foreach ($productkey in $productkeys) {
    $ospp /inpkey:$productkey
}
cscript $ospp /sethst:kms.msgang.com
cscript $ospp /act
pause


