
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

$path64 = "C:\Program Files\Microsoft Office\Office1*"
$path32 = "C:\Program Files (x86)\Microsoft Office\Office1*"
if ("$path64\ospp.vbs") {
    Set-Location $path64
}
else {
Set-Location $path32
}

#$ospp = (Resolve-Path -Path "C:\Program Files*\Microsoft Office\Office16\ospp.vbs").Path
#Find OSPP.vbs path and run the command with the dstatus option (Last 1...)
Write-Host '---------------------------------------------------------------'
Write-Host "Processing...It could take a while, please be patient."                 
Write-Host

$dstatus = Invoke-Expression -Command "cscript.exe ospp.vbs /dstatus"
$keys = $dstatus | Select-String -SimpleMatch "Last 5" | ForEach-Object -Process { $_.tostring().split(" ")[-1]}
foreach ($key in $keys) {
    cscript ospp.vbs /unpkey:$key | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office19").Count -gt 0) {

    $2019keys = @(
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
        'FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH'
    )

    foreach ($2019key in $2019keys) {
        cscript ospp.vbs /inpkey:$2019key | Out-Null
    }
}

if (($dstatus | Select-String -SimpleMatch "Office21").Count -gt 0) {

    $2021keys = @(
        'FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH';    #Office LTSC Professional Plus 2021
        'KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3';    #Office LTSC Standard 2021
        'FTNWT-C6WBT-8HMGF-K9PRX-QV9H8';    #Project Professional 2021
        'J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T';    #Project Standard 2021
        'KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4';    #Visio LTSC Professional 2021
        'MJVNY-BYWPY-CWV6J-2RKRT-4M8QG';    #Visio LTSC Standard 2021
        'WM8YG-YNGDD-4JHDC-PG3F4-FC4T4';    #Access LTSC 2021
        'NWG3X-87C9K-TC7YY-BC2G7-G6RVC';    #Excel LTSC 2021
        'C9FM6-3N72F-HFJXB-TM3V9-T86R9';    #Outlook LTSC 2021
        'TY7XF-NFRBR-KJ44C-G83KF-GX27K';    #PowerPoint LTSC 2021
        '2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ';    #Publisher LTSC 2021
        'TN8H9-M34D3-Y64V9-TR72V-X79KV'     #Word LTSC 2021
    )

    foreach ($2021key in $2021keys) {
        cscript ospp.vbs /inpkey:$2021key | Out-Null
    }
}

if (($dstatus | Select-String -SimpleMatch "Office16ProPlus").Count -gt 0) {
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-pl.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ul-phn.xrm-ms" | Out-Null
    cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office16Standard").Count -gt 0) {
    cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_KMS_Client-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_KMS_Client-ul.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_KMS_Client-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_MAK-pl.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_MAK-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_MAK-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_MAK-ul-phn.xrm-ms" | Out-Null
    cscript ospp.vbs /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office16VisioPro").Count -gt 0) {
    cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_MAK-pl.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_MAK-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_MAK-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_MAK-ul-phn.xrm-ms" | Out-Null
    cscript ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office16ProjectPro").Count -gt 0) {
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_MAK-pl.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_MAK-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_MAK-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_MAK-ul-phn.xrm-ms" | Out-Null
    cscript ospp.vbs /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT | Out-Null
}

#For Office 2013 Retail verion.
if (($dstatus | Select-String -SimpleMatch "OfficeProfessional").Count -gt 0) {
    New-Item -Path $env:temp\tmp -ItemType Directory -Force | Out-Null
    (New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/bonben365/microsoft/main/proplusvl_kms_client-ppd.xrm-ms', "$env:temp\tmp\proplusvl_kms_client-ppd.xrm-ms") | Out-Null
    (New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/bonben365/microsoft/main/proplusvl_kms_client-ul-oob.xrm-ms', "$env:temp\tmp\proplusvl_kms_client-ul-oob.xrm-ms") | Out-Null
    (New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/bonben365/microsoft/main/proplusvl_kms_client-ul.xrm-ms', "$env:temp\tmp\proplusvl_kms_client-ul.xrm-ms") | Out-Null
    cscript ospp.vbs /inslic:"$env:temp\tmp\proplusvl_kms_client-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"$env:temp\tmp\proplusvl_kms_client-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"$env:temp\tmp\proplusvl_kms_client-ul.xrm-ms" | Out-Null
    cscript ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT | Out-Null
}

cscript ospp.vbs /sethst:kms.msgang.com | Out-Null
cscript ospp.vbs /act

