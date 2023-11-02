Write-Host =========================================================================
Write-Host Name:           Microsoft Office Activator by Leo.
Write-Host Description:    Activate all Offices Editions for free without any software.
Write-Host Website:        https://msgang.com
Write-Host Script by:      Leo Nguyen
Write-Host =========================================================================

if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}

function Remove-OfficeRetail {
    $ProductKeys = $DStatus | Select-String -SimpleMatch "Last 5" | ForEach-Object -Process { $_.tostring().split(" ")[-1]}
    if ($ProductKeys) {
        Write-Host "Found $(($ProductKeys | Measure-Object).Count) productkeys, proceeding with deactivation..." -ForegroundColor Green
    
        foreach ($ProductKey in $ProductKeys) {
            Write-Host "Processing productkey $ProductKey" -ForegroundColor Green
            $Command = "cscript.exe //nologo ospp.vbs /unpkey:$ProductKey"
            Invoke-Expression -Command $Command | Out-Null
        }
        Write-Host "Converting Office Retail to Volume..." -ForegroundColor Green
        Write-Host
    } else {}
}

function Download-Library {
    New-Item -Path $env:temp\tmp -ItemType Directory -Force | Out-Null
    (New-Object Net.WebClient).DownloadFile('https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/office2013/library.zip', "$env:temp\tmp\library.zip") | Out-Null
    Expand-Archive "$env:temp\tmp\library.zip" "$env:temp\tmp\library" -Force | Out-Null
}


$path64 = "C:\Program Files\Microsoft Office\Office1*"
$path32 = "C:\Program Files (x86)\Microsoft Office\Office1*"
if ((Test-Path -Path "$path32\ospp.vbs")) { Set-Location $path32 -ErrorAction SilentlyContinue }
if ((Test-Path -Path "$path64\ospp.vbs")) { Set-Location $path64 -ErrorAction SilentlyContinue }

#$ospp = (Resolve-Path -Path "C:\Program Files*\Microsoft Office\Office1*\ospp.vbs").Path
#Find OSPP.vbs path and run the command with the dstatus option (Last 1...)
Write-Host
Write-Host "Checking installed Office editions..." -ForegroundColor Green
$dstatus = Invoke-Expression -Command "cscript.exe ospp.vbs /dstatus"

$apps = $dstatus | Select-String -SimpleMatch "NAME:" | ForEach-Object -Process { $_.tostring().split(" ")[-2]}
foreach ($app in $apps) {
    Write-Host "Installed Office: $app `n" -ForegroundColor Yellow
    Write-Host "Activating $app, please be patient.`n" -ForegroundColor Green
}
#For Office 2019 VL.
if (($dstatus | Select-String -SimpleMatch "Office19" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 19, VOLUME" | Measure-Object).Count -gt 0 ) {

    $2019keys = @(
        'NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP';    #Office Professional Plus 2019
        '6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK';    #Office Standard 2019
        'B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B';    #Project Professional 2019
        'C4F7P-NCP8C-6CQPT-MQHV9-JXD2M';    #Project Standard 2019
        '9BGNQ-K37YR-RQHF2-38RQ3-7VCBB';    #Visio Professional 2019
        '7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2';    #Visio Standard 2019
        '9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT';    #Access 2019
        'TMJWT-YYNMB-3BKTF-644FC-RVXBD';    #Excel 2019
        '7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK';    #Outlook 2019
        'RRNCX-C64HY-W2MM7-MCH9G-TJHMQ';    #PowerPoint 2019
        'G2KWX-3NW6P-PY93R-JXK2T-C9Y9V';    #Publisher 2019
        'PBX3G-NWMT6-Q7XBW-PYJGG-WXD33'     #Word 2019
    )

    foreach ($2019key in $2019keys) {
        cscript ospp.vbs /inpkey:$2019key | Out-Null
    }
}

#For Office 2019 Retail.
if (($dstatus | Select-String -SimpleMatch "Office19Professional" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 19, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    $licenses = Get-Item "..\root\Licenses16\ProPlus2019VL*.xrm-ms"
    foreach ($license in $licenses) {
        cscript ospp.vbs /inslic:"..\root\Licenses16\$($license.Name)" | Out-Null
    }
    cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office19ProPlus" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 19, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    $licenses = Get-Item "..\root\Licenses16\ProPlus2019VL*.xrm-ms"
    foreach ($license in $licenses) {
        cscript ospp.vbs /inslic:"..\root\Licenses16\$($license.Name)" | Out-Null
    }
    cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office19Standard" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 19, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    $licenses = Get-Item "..\root\Licenses16\Standard2019VL*.xrm-ms"
    foreach ($license in $licenses) {
        cscript ospp.vbs /inslic:"..\root\Licenses16\$($license.Name)" | Out-Null
    }
    cscript ospp.vbs /inpkey:6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office19VisioPro" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 19, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    $licenses = Get-Item "..\root\Licenses16\VisioPro2019VL*.xrm-ms"
    foreach ($license in $licenses) {
        cscript ospp.vbs /inslic:"..\root\Licenses16\$($license.Name)" | Out-Null
    }
    cscript ospp.vbs /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office19ProjectPro" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 19, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    $licenses = Get-Item "..\root\Licenses16\ProjectPro2019VL*.xrm-ms"
    foreach ($license in $licenses) {
        cscript ospp.vbs /inslic:"..\root\Licenses16\$($license.Name)" | Out-Null
    }
    cscript ospp.vbs /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B | Out-Null
}

#For Office 2021 VL.
if (($dstatus | Select-String -SimpleMatch "Office21" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 21, VOLUME" | Measure-Object).Count -gt 0 ) {

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

#For Office 2021 Retail.
if (($dstatus | Select-String -SimpleMatch "Office21Professional" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 21, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    $licenses = Get-Item "..\root\Licenses16\ProPlus2021VL*.xrm-ms"
    foreach ($license in $licenses) {
        cscript ospp.vbs /inslic:"..\root\Licenses16\$($license.Name)" | Out-Null
    }
    cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office21ProPlus" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 21, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    $licenses = Get-Item "..\root\Licenses16\ProPlus2021VL*.xrm-ms"
    foreach ($license in $licenses) {
        cscript ospp.vbs /inslic:"..\root\Licenses16\$($license.Name)" | Out-Null
    }
    cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office21Standard" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 21, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    $licenses = Get-Item "..\root\Licenses16\Standard2021VL*.xrm-ms"
    foreach ($license in $licenses) {
        cscript ospp.vbs /inslic:"..\root\Licenses16\$($license.Name)" | Out-Null
    }
    cscript ospp.vbs /inpkey:KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3 | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office21VisioPro" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 21, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    $licenses = Get-Item "..\root\Licenses16\VisioPro2021VL*.xrm-ms"
    foreach ($license in $licenses) {
        cscript ospp.vbs /inslic:"..\root\Licenses16\$($license.Name)" | Out-Null
    }
    cscript ospp.vbs /inpkey:KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4 | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office21ProjectPro" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 21, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    $licenses = Get-Item "..\root\Licenses16\ProjectPro2021VL*.xrm-ms"
    foreach ($license in $licenses) {
        cscript ospp.vbs /inslic:"..\root\Licenses16\$($license.Name)" | Out-Null
    }
    cscript ospp.vbs /inpkey:FTNWT-C6WBT-8HMGF-K9PRX-QV9H8 | Out-Null
}

#For Office 2016 VL.
if (($dstatus | Select-String -SimpleMatch "Office16ProPlus" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 16, VOLUME" | Measure-Object).Count -gt 0 ) {
    cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office16Standard" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 16, VOLUME" | Measure-Object).Count -gt 0 ) {
    cscript ospp.vbs /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM | Out-Null
}

#For Office 2016 Retail (MSDN)
if (($dstatus | Select-String -SimpleMatch "Office16ProPlus" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 16, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-pl.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ul-phn.xrm-ms" | Out-Null
    cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 | Out-Null
}

#For Office 2016 Retail.
if (($dstatus | Select-String -SimpleMatch "Office16Professional" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 16, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-pl.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ul-phn.xrm-ms" | Out-Null
    cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office16Standard" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 16, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_KMS_Client-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_KMS_Client-ul.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_KMS_Client-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_MAK-pl.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_MAK-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_MAK-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\StandardVL_MAK-ul-phn.xrm-ms" | Out-Null
    cscript ospp.vbs /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office16VisioPro" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 16, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_MAK-pl.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_MAK-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_MAK-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_MAK-ul-phn.xrm-ms" | Out-Null
    cscript ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "Office16ProjectPro" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 16, RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_MAK-pl.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_MAK-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_MAK-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_MAK-ul-phn.xrm-ms" | Out-Null
    cscript ospp.vbs /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT | Out-Null
}

#For Office 2013 VL.
if (($dstatus | Select-String -SimpleMatch "OfficeProfessional" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "VOLUME_" | Measure-Object).Count -gt 0 ) {
    cscript ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT | Out-Null
}

if (($dstatus | Select-String -SimpleMatch "OfficeStandard" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "VOLUME_" | Measure-Object).Count -gt 0 ) {
    cscript ospp.vbs /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4 | Out-Null
}

#For Office 2013 Retail (MSDN)
if (($dstatus | Select-String -SimpleMatch "OfficeProPlus" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 15" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    New-Item -Path $env:temp\tmp -ItemType Directory -Force | Out-Null
    (New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/bonben365/microsoft/main/Office2013_Library/proplusvl_kms_client-ppd.xrm-ms', "$env:temp\tmp\proplusvl_kms_client-ppd.xrm-ms") | Out-Null
    (New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/bonben365/microsoft/main/Office2013_Library/proplusvl_kms_client-ul-oob.xrm-ms', "$env:temp\tmp\proplusvl_kms_client-ul-oob.xrm-ms") | Out-Null
    (New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/bonben365/microsoft/main/Office2013_Library/proplusvl_kms_client-ul.xrm-ms', "$env:temp\tmp\proplusvl_kms_client-ul.xrm-ms") | Out-Null
    cscript ospp.vbs /inslic:"$env:temp\tmp\proplusvl_kms_client-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"$env:temp\tmp\proplusvl_kms_client-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"$env:temp\tmp\proplusvl_kms_client-ul.xrm-ms" | Out-Null
    cscript ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT | Out-Null
}

#For Office 2013 Retail.
if (($dstatus | Select-String -SimpleMatch "OfficeProfessional" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 15" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    New-Item -Path $env:temp\tmp -ItemType Directory -Force | Out-Null
    (New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/bonben365/microsoft/main/Office2013_Library/proplusvl_kms_client-ppd.xrm-ms', "$env:temp\tmp\proplusvl_kms_client-ppd.xrm-ms") | Out-Null
    (New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/bonben365/microsoft/main/Office2013_Library/proplusvl_kms_client-ul-oob.xrm-ms', "$env:temp\tmp\proplusvl_kms_client-ul-oob.xrm-ms") | Out-Null
    (New-Object Net.WebClient).DownloadFile('https://raw.githubusercontent.com/bonben365/microsoft/main/Office2013_Library/proplusvl_kms_client-ul.xrm-ms', "$env:temp\tmp\proplusvl_kms_client-ul.xrm-ms") | Out-Null
    cscript ospp.vbs /inslic:"$env:temp\tmp\proplusvl_kms_client-ppd.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"$env:temp\tmp\proplusvl_kms_client-ul-oob.xrm-ms" | Out-Null
    cscript ospp.vbs /inslic:"$env:temp\tmp\proplusvl_kms_client-ul.xrm-ms" | Out-Null
    cscript ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT | Out-Null
}


#For Office 2013 Standalone (Word)
if (($dstatus | Select-String -SimpleMatch "Word" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 15" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    Download-Library
    $inslics = Get-ChildItem -Path "$env:temp\tmp\library" | Where-Object {$_.Name -like 'Word*'}
    foreach ($inslic in $inslics){
        $licname = $inslic.Name
        cscript ospp.vbs /inslic:"$env:temp\tmp\library\$licname" | Out-Null
    }
    cscript ospp.vbs /inpkey:6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7 | Out-Null
}

#For Office 2013 Standalone (Excel)
if (($dstatus | Select-String -SimpleMatch "Excel" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 15" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    Download-Library
    $inslics = Get-ChildItem -Path "$env:temp\tmp\library" | Where-Object {$_.Name -like 'excel*'}
    foreach ($inslic in $inslics){
        $licname = $inslic.Name
        cscript ospp.vbs /inslic:"$env:temp\tmp\library\$licname" | Out-Null
    }
    cscript ospp.vbs /inpkey:VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB | Out-Null
}

#For Office 2013 Standalone (PowerPoint)
if (($dstatus | Select-String -SimpleMatch "PowerPoint" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 15" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    Download-Library
    $inslics = Get-ChildItem -Path "$env:temp\tmp\library" | Where-Object {$_.Name -like 'PowerPoint*'}
    foreach ($inslic in $inslics){
        $licname = $inslic.Name
        cscript ospp.vbs /inslic:"$env:temp\tmp\library\$licname" | Out-Null
    }
    cscript ospp.vbs /inpkey:4NT99-8RJFH-Q2VDH-KYG2C-4RD4F | Out-Null
}

#For Office 2013 Standalone (Outlook)
if (($dstatus | Select-String -SimpleMatch "Outlook" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 15" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    Download-Library
    $inslics = Get-ChildItem -Path "$env:temp\tmp\library" | Where-Object {$_.Name -like 'Outlook*'}
    foreach ($inslic in $inslics){
        $licname = $inslic.Name
        cscript ospp.vbs /inslic:"$env:temp\tmp\library\$licname" | Out-Null
    }
    cscript ospp.vbs /inpkey:QPN8Q-BJBTJ-334K3-93TGY-2PMBT | Out-Null
}

#For Office 2013 Standalone (Access)
if (($dstatus | Select-String -SimpleMatch "Access" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 15" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    Download-Library
    $inslics = Get-ChildItem -Path "$env:temp\tmp\library" | Where-Object {$_.Name -like 'Access*'}
    foreach ($inslic in $inslics){
        $licname = $inslic.Name
        cscript ospp.vbs /inslic:"$env:temp\tmp\library\$licname" | Out-Null
    }
    cscript ospp.vbs /inpkey:NG2JY-H4JBT-HQXYP-78QH9-4JM2D | Out-Null
}

#For Office 2013 Standalone (Publisher)
if (($dstatus | Select-String -SimpleMatch "Publisher" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 15" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    Download-Library
    $inslics = Get-ChildItem -Path "$env:temp\tmp\library" | Where-Object {$_.Name -like 'Publisher*'}
    foreach ($inslic in $inslics){
        $licname = $inslic.Name
        cscript ospp.vbs /inslic:"$env:temp\tmp\library\$licname" | Out-Null
    }
    cscript ospp.vbs /inpkey:PN2WF-29XG2-T9HJ7-JQPJR-FCXK4 | Out-Null
}

#For Office 2013 Standalone (Visio)
if (($dstatus | Select-String -SimpleMatch "Visio" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 15" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    Download-Library
    $inslics = Get-ChildItem -Path "$env:temp\tmp\library" | Where-Object {$_.Name -like 'Visio*'}
    foreach ($inslic in $inslics){
        $licname = $inslic.Name
        cscript ospp.vbs /inslic:"$env:temp\tmp\library\$licname" | Out-Null
    }
    cscript ospp.vbs /inpkey:C2FG9-N6J68-H8BTJ-BW3QX-RM3B3 | Out-Null
    cscript ospp.vbs /inpkey:J484Y-4NKBF-W2HMG-DBMJC-PGWR7 | Out-Null
}

#For Office 2013 Standalone (Project)
if (($dstatus | Select-String -SimpleMatch "Project" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 15" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "RETAIL" | Measure-Object).Count -gt 0 ) {
    Remove-OfficeRetail
    Download-Library
    $inslics = Get-ChildItem -Path "$env:temp\tmp\library" | Where-Object {$_.Name -like 'Project*'}
    foreach ($inslic in $inslics){
        $licname = $inslic.Name
        cscript ospp.vbs /inslic:"$env:temp\tmp\library\$licname" | Out-Null
    }
    cscript ospp.vbs /inpkey:FN8TT-7WMH6-2D4X9-M337T-2342K | Out-Null
    cscript ospp.vbs /inpkey:6NTH3-CW976-3G3Y2-JK3TX-8QHTT | Out-Null
}


cscript ospp.vbs /sethst:kms.msgang.com | Out-Null
cscript //nologo ospp.vbs /act
