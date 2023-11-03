#Write-Host =========================================================================
#Write-Host Name:           Microsoft Office Activator by Leo.
#Write-Host Description:    Activate all Offices Editions for free without any software.
#Write-Host Website:        https://msgang.com
#Write-Host Script by:      Leo Nguyen
#Write-Host =========================================================================

if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}

#Detect Office installations
$path64 = "C:\Program Files\Microsoft Office\Office1*"
$path32 = "C:\Program Files (x86)\Microsoft Office\Office1*"
if ((Test-Path -Path "$path32\ospp.vbs")) { Set-Location $path32 -ErrorAction SilentlyContinue }
if ((Test-Path -Path "$path64\ospp.vbs")) { Set-Location $path64 -ErrorAction SilentlyContinue }

#$ospp = (Resolve-Path -Path "C:\Program Files*\Microsoft Office\Office1*\ospp.vbs").Path
#Find OSPP.vbs path and run the command with the dstatus option (Last 1...)
Write-Host
Write-Host "Checking Installed Office editions..." -ForegroundColor Green
$dstatus = Invoke-Expression -Command "cscript.exe ospp.vbs /dstatus"

#$apps = $dstatus | Select-String -SimpleMatch "NAME:" | ForEach-Object -Process { $_.tostring().split(" ")[-2]}
$retailCount = ($dstatus | Select-String -SimpleMatch "RETAIL" | Measure-Object).Count
$volumeCount = ($dstatus | Select-String -SimpleMatch "VOLUME" | Measure-Object).Count
Write-Host "Number of Installed Office apps: VOLUME: $volumeCount - RETAIL: $retailCount" -ForegroundColor Green
Write-Host "Activating Microsoft Office products..." -ForegroundColor Green

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
function Office2019-V2R {
    Write-Host "Converting from Retail to Volume..." -ForegroundColor Green
    $inslics = Get-ChildItem -Path "..\root\Licenses16" | Where-Object {$_.Name -like "$licName"}
    foreach ($inslic in $inslics){
        cscript ospp.vbs /inslic:"..\root\Licenses16\$($inslic.Name)" | Out-Null
    }
    cscript ospp.vbs /inpkey:$2019kmskey | Out-Null
}

$matchingOffice2019Retail = ($dstatus | Select-String -SimpleMatch "Office 19, RETAIL" | Measure-Object).Count
if (($dstatus | Select-String -SimpleMatch "Office19Professional" | Measure-Object).Count -gt 0 -and $matchingOffice2019Retail -gt 0 ) {$licName = 'ProPlus2019VL*'; $2019kmskey = 'NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP'; Office2019-V2R} #For Office 2019 Retail (Pro)
if (($dstatus | Select-String -SimpleMatch "Office19ProPlus" | Measure-Object).Count -gt 0 -and $matchingOffice2019Retail -gt 0 ) {$licName = 'ProPlus2019VL*'; $2019kmskey = 'NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP'; Office2019-V2R} #For Office 2019 Retail (Pro) (MSDN)
if (($dstatus | Select-String -SimpleMatch "Office19Standard" | Measure-Object).Count -gt 0 -and $matchingOffice2019Retail -gt 0 ) {$licName = 'Standard2019VL*'; $2019kmskey = '6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK'; Office2019-V2R} #For Office 2019 Retail (Standard)
if (($dstatus | Select-String -SimpleMatch "Office19ProjectPro" | Measure-Object).Count -gt 0 -and $matchingOffice2019Retail -gt 0 ) {$licName = 'ProjectPro2019VL*'; $2019kmskey = 'B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B'; Office2019-V2R} #For Office 2019 Standalone (Project Pro)
if (($dstatus | Select-String -SimpleMatch "Office19ProjectStd" | Measure-Object).Count -gt 0 -and $matchingOffice2019Retail -gt 0 ) {$licName = 'ProjectStd2019VL*'; $2019kmskey = 'C4F7P-NCP8C-6CQPT-MQHV9-JXD2M'; Office2019-V2R} #For Office 2019 Standalone (Project Standard)
if (($dstatus | Select-String -SimpleMatch "Office19VisioPro" | Measure-Object).Count -gt 0 -and $matchingOffice2019Retail -gt 0 ) {$licName = 'VisioPro2019VL*'; $2019kmskey = '9BGNQ-K37YR-RQHF2-38RQ3-7VCBB'; Office2019-V2R} #For Office 2019 Standalone (Visio Pro)
if (($dstatus | Select-String -SimpleMatch "Office19VisioStd" | Measure-Object).Count -gt 0 -and $matchingOffice2019Retail -gt 0 ) {$licName = 'VisioStd2019VL*'; $2019kmskey = '7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2'; Office2019-V2R} #For Office 2019 Standalone (Visio Standard)
if (($dstatus | Select-String -SimpleMatch "Office19Word" | Measure-Object).Count -gt 0 -and $matchingOffice2019Retail -gt 0 ) {$licName = 'Word2019VL*'; $2019kmskey = 'PBX3G-NWMT6-Q7XBW-PYJGG-WXD33'; Office2019-V2R} #For Office 2019 Standalone (Word)
if (($dstatus | Select-String -SimpleMatch "Office19Excel" | Measure-Object).Count -gt 0 -and $matchingOffice2019Retail -gt 0 ) {$licName = 'Excel2019VL*'; $2019kmskey = 'TMJWT-YYNMB-3BKTF-644FC-RVXBD'; Office2019-V2R} #For Office 2019 Standalone (Excel)
if (($dstatus | Select-String -SimpleMatch "Office19PowerPoint" | Measure-Object).Count -gt 0 -and $matchingOffice2019Retail -gt 0 ) {$licName = 'PowerPoint2019VL*'; $2019kmskey = 'RRNCX-C64HY-W2MM7-MCH9G-TJHMQ'; Office2019-V2R} #For Office 2019 Standalone (PowerPoint)
if (($dstatus | Select-String -SimpleMatch "Office19Outlook" | Measure-Object).Count -gt 0 -and $matchingOffice2019Retail -gt 0 ) {$licName = 'Outlook2019VL*'; $2019kmskey = '7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK'; Office2019-V2R} #For Office 2019 Standalone (Outlook)
if (($dstatus | Select-String -SimpleMatch "Office19Access" | Measure-Object).Count -gt 0 -and $matchingOffice2019Retail -gt 0 ) {$licName = 'Access2019VL*'; $2019kmskey = '9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT'; Office2019-V2R} #For Office 2019 Standalone (Access)
if (($dstatus | Select-String -SimpleMatch "Office19Publisher" | Measure-Object).Count -gt 0 -and $matchingOffice2019Retail -gt 0 ) {$licName = 'Publisher2019VL*'; $2019kmskey = 'G2KWX-3NW6P-PY93R-JXK2T-C9Y9V'; Office2019-V2R} #For Office 2019 Standalone (Publisher)

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

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
function Office2021-V2R {
    Write-Host "Converting from Retail to Volume..." -ForegroundColor Green
    $inslics = Get-ChildItem -Path "..\root\Licenses16" | Where-Object {$_.Name -like "$licName"}
    foreach ($inslic in $inslics){
        cscript ospp.vbs /inslic:"..\root\Licenses16\$($inslic.Name)" | Out-Null
    }
    cscript ospp.vbs /inpkey:$2021kmskey | Out-Null
}

$matchingOffice2021Retail = ($dstatus | Select-String -SimpleMatch "Office 21, RETAIL" | Measure-Object).Count
if (($dstatus | Select-String -SimpleMatch "Office21Professional" | Measure-Object).Count -gt 0 -and $matchingOffice2021Retail -gt 0 ) {$licName = 'ProPlus2021VL*'; $2021kmskey = 'FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH'; Office2021-V2R} #For Office 2021 Retail (Pro)
if (($dstatus | Select-String -SimpleMatch "Office21ProPlus" | Measure-Object).Count -gt 0 -and $matchingOffice2021Retail -gt 0 ) {$licName = 'ProPlus2021VL*'; $2021kmskey = 'FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH'; Office2021-V2R} #For Office 2021 Retail (Pro) (MSDN)
if (($dstatus | Select-String -SimpleMatch "Office21Standard" | Measure-Object).Count -gt 0 -and $matchingOffice2021Retail -gt 0 ) {$licName = 'Standard2021VL*'; $2021kmskey = 'KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3'; Office2021-V2R} #For Office 2021 Retail (Standard)
if (($dstatus | Select-String -SimpleMatch "Office21ProjectPro" | Measure-Object).Count -gt 0 -and $matchingOffice2021Retail -gt 0 ) {$licName = 'ProjectPro2021VL*'; $2021kmskey = 'FTNWT-C6WBT-8HMGF-K9PRX-QV9H8'; Office2021-V2R} #For Office 2021 Standalone (Project Pro)
if (($dstatus | Select-String -SimpleMatch "Office21ProjectStd" | Measure-Object).Count -gt 0 -and $matchingOffice2021Retail -gt 0 ) {$licName = 'ProjectStd2021VL*'; $2021kmskey = 'J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T'; Office2021-V2R} #For Office 2021 Standalone (Project Standard)
if (($dstatus | Select-String -SimpleMatch "Office21VisioPro" | Measure-Object).Count -gt 0 -and $matchingOffice2021Retail -gt 0 ) {$licName = 'VisioPro2021VL*'; $2021kmskey = 'KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4'; Office2021-V2R} #For Office 2021 Standalone (Visio Pro)
if (($dstatus | Select-String -SimpleMatch "Office21VisioStd" | Measure-Object).Count -gt 0 -and $matchingOffice2021Retail -gt 0 ) {$licName = 'VisioStd2021VL*'; $2021kmskey = 'MJVNY-BYWPY-CWV6J-2RKRT-4M8QG'; Office2021-V2R} #For Office 2021 Standalone (Visio Standard)
if (($dstatus | Select-String -SimpleMatch "Office21Word" | Measure-Object).Count -gt 0 -and $matchingOffice2021Retail -gt 0 ) {$licName = 'Word2021VL*'; $2021kmskey = 'TN8H9-M34D3-Y64V9-TR72V-X79KV'; Office2021-V2R} #For Office 2021 Standalone (Word)
if (($dstatus | Select-String -SimpleMatch "Office21Excel" | Measure-Object).Count -gt 0 -and $matchingOffice2021Retail -gt 0 ) {$licName = 'Excel2021VL*'; $2021kmskey = 'NWG3X-87C9K-TC7YY-BC2G7-G6RVC'; Office2021-V2R} #For Office 2021 Standalone (Excel)
if (($dstatus | Select-String -SimpleMatch "Office21PowerPoint" | Measure-Object).Count -gt 0 -and $matchingOffice2021Retail -gt 0 ) {$licName = 'PowerPoint2021VL*'; $2021kmskey = 'TY7XF-NFRBR-KJ44C-G83KF-GX27K'; Office2021-V2R} #For Office 2021 Standalone (PowerPoint)
if (($dstatus | Select-String -SimpleMatch "Office21Outlook" | Measure-Object).Count -gt 0 -and $matchingOffice2021Retail -gt 0 ) {$licName = 'Outlook2021VL*'; $2021kmskey = 'C9FM6-3N72F-HFJXB-TM3V9-T86R9'; Office2021-V2R} #For Office 2021 Standalone (Outlook)
if (($dstatus | Select-String -SimpleMatch "Office21Access" | Measure-Object).Count -gt 0 -and $matchingOffice2021Retail -gt 0 ) {$licName = 'Access2021VL*'; $2021kmskey = 'WM8YG-YNGDD-4JHDC-PG3F4-FC4T4'; Office2021-V2R} #For Office 2021 Standalone (Access)
if (($dstatus | Select-String -SimpleMatch "Office21Publisher" | Measure-Object).Count -gt 0 -and $matchingOffice2021Retail -gt 0 ) {$licName = 'Publisher2021VL*'; $2021kmskey = '2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ'; Office2021-V2R} #For Office 2021 Standalone (Publisher)

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#For Office 2016 VL
if (($dstatus | Select-String -SimpleMatch "Office16" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "Office 16, VOLUME" | Measure-Object).Count -gt 0 ) {

    $2016keys = @(
        'XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99';    #Office Professional Plus 2016
        'JNRGM-WHDWX-FJJG3-K47QV-DRTFM';    #Office Standard 2016
        'YG9NW-3K39V-2T3HJ-93F3Q-G83KT';    #Project Professional 2016
        'GNFHQ-F6YQM-KQDGJ-327XX-KQBVC';    #Project Standard 2016
        'PD3PC-RHNGV-FXJ29-8JK7D-RJRJK';    #Visio Professional 2016
        '7WHWN-4T7MP-G96JF-G33KR-W8GF4';    #Visio Standard 2016
        'GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW';    #Access 2016
        '9C2PK-NWTVB-JMPW8-BFT28-7FTBF';    #Excel 2016
        'R69KK-NTPKF-7M3Q4-QYBHW-6MT9B';    #Outlook 2016
        'J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6';    #PowerPoint 2016
        'F47MM-N3XJP-TQXJ9-BP99D-8837K';    #Publisher 2016
        'WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6'     #Word 2016
    )

    foreach ($2016key in $2016keys) {
        cscript ospp.vbs /inpkey:$2016key | Out-Null
    }
}

function Office2016-V2R {
    Write-Host "Converting from Retail to Volume..." -ForegroundColor Green
    $inslics = Get-ChildItem -Path "..\root\Licenses16" | Where-Object {$_.Name -like "$licName"}
    foreach ($inslic in $inslics){
        cscript ospp.vbs /inslic:"..\root\Licenses16\$($inslic.Name)" | Out-Null
    }
    cscript ospp.vbs /inpkey:$2016kmskey | Out-Null
}

$matchingOffice2016Retail = ($dstatus | Select-String -SimpleMatch "Office 16, RETAIL" | Measure-Object).Count

if (($dstatus | Select-String -SimpleMatch "Office16Professional" | Measure-Object).Count -gt 0 -and $matchingOffice2016Retail -gt 0 ) {$licName = 'ProPlusVL*'; $2016kmskey = 'XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99'; Office2016-V2R} #For Office 2016 Retail (Pro)
if (($dstatus | Select-String -SimpleMatch "Office16ProPlus" | Measure-Object).Count -gt 0 -and $matchingOffice2016Retail -gt 0 ) {$licName = 'ProPlusVL*'; $2016kmskey = 'XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99'; Office2016-V2R} #For Office 2016 Retail (Pro) (MSDN)
if (($dstatus | Select-String -SimpleMatch "Office16Standard" | Measure-Object).Count -gt 0 -and $matchingOffice2016Retail -gt 0 ) {$licName = 'StandardVL*'; $2016kmskey = 'JNRGM-WHDWX-FJJG3-K47QV-DRTFM'; Office2016-V2R} #For Office 2016 Retail (Standard)
if (($dstatus | Select-String -SimpleMatch "Office16ProjectPro" | Measure-Object).Count -gt 0 -and $matchingOffice2016Retail -gt 0 ) {$licName = 'ProjectProVL*'; $2016kmskey = 'YG9NW-3K39V-2T3HJ-93F3Q-G83KT'; Office2016-V2R} #For Office 2016 Standalone (Project Pro)
if (($dstatus | Select-String -SimpleMatch "Office16ProjectStd" | Measure-Object).Count -gt 0 -and $matchingOffice2016Retail -gt 0 ) {$licName = 'ProjectStdVL*'; $2016kmskey = 'GNFHQ-F6YQM-KQDGJ-327XX-KQBVC'; Office2016-V2R} #For Office 2016 Standalone (Project Standard)
if (($dstatus | Select-String -SimpleMatch "Office16VisioPro" | Measure-Object).Count -gt 0 -and $matchingOffice2016Retail -gt 0 ) {$licName = 'VisioProVL*'; $2016kmskey = 'PD3PC-RHNGV-FXJ29-8JK7D-RJRJK'; Office2016-V2R} #For Office 2016 Standalone (Visio Pro)
if (($dstatus | Select-String -SimpleMatch "Office16VisioStd" | Measure-Object).Count -gt 0 -and $matchingOffice2016Retail -gt 0 ) {$licName = 'VisioStdVL*'; $2016kmskey = '7WHWN-4T7MP-G96JF-G33KR-W8GF4'; Office2016-V2R} #For Office 2016 Standalone (Visio Standard)
if (($dstatus | Select-String -SimpleMatch "Office16Word" | Measure-Object).Count -gt 0 -and $matchingOffice2016Retail -gt 0 ) {$licName = 'WordVL*'; $2016kmskey = 'WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6'; Office2016-V2R} #For Office 2016 Standalone (Word)
if (($dstatus | Select-String -SimpleMatch "Office16Excel" | Measure-Object).Count -gt 0 -and $matchingOffice2016Retail -gt 0 ) {$licName = 'ExcelVL*'; $2016kmskey = '9C2PK-NWTVB-JMPW8-BFT28-7FTBF'; Office2016-V2R} #For Office 2016 Standalone (Excel)
if (($dstatus | Select-String -SimpleMatch "Office16PowerPoint" | Measure-Object).Count -gt 0 -and $matchingOffice2016Retail -gt 0 ) {$licName = 'PowerPointVL*'; $2016kmskey = 'J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6'; Office2016-V2R} #For Office 2016 Standalone (PowerPoint)
if (($dstatus | Select-String -SimpleMatch "Office16Outlook" | Measure-Object).Count -gt 0 -and $matchingOffice2016Retail -gt 0 ) {$licName = 'OutlookVL*'; $2016kmskey = 'R69KK-NTPKF-7M3Q4-QYBHW-6MT9B'; Office2016-V2R} #For Office 2016 Standalone (Outlook)
if (($dstatus | Select-String -SimpleMatch "Office16Access" | Measure-Object).Count -gt 0 -and $matchingOffice2016Retail -gt 0 ) {$licName = 'AccessVL*'; $2016kmskey = 'GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW'; Office2016-V2R} #For Office 2016 Standalone (Access)
if (($dstatus | Select-String -SimpleMatch "Office16Publisher" | Measure-Object).Count -gt 0 -and $matchingOffice2016Retail -gt 0 ) {$licName = 'PublisherVL*'; $2016kmskey = 'F47MM-N3XJP-TQXJ9-BP99D-8837K'; Office2016-V2R} #For Office 2016 Standalone (Publisher)

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#For Office 2013 VL.
if (($dstatus | Select-String -SimpleMatch "Office15" | Measure-Object).Count -gt 0 -and ($dstatus | Select-String -SimpleMatch "VOLUME_" | Measure-Object).Count -gt 0 ) {

    $2013keys = @(
        'YC7DK-G2NP3-2QQC3-J6H88-GVGXT';    #Office Professional Plus 2013
        'KBKQT-2NMXY-JJWGP-M62JB-92CD4';    #Office Standard 2013
        'FN8TT-7WMH6-2D4X9-M337T-2342K';    #Project Professional 2013
        '6NTH3-CW976-3G3Y2-JK3TX-8QHTT';    #Project Standard 2013
        'C2FG9-N6J68-H8BTJ-BW3QX-RM3B3';    #Visio Professional 2013
        'J484Y-4NKBF-W2HMG-DBMJC-PGWR7';    #Visio Standard 2013
        'NG2JY-H4JBT-HQXYP-78QH9-4JM2D';    #Access 2013
        'VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB';    #Excel 2013
        'QPN8Q-BJBTJ-334K3-93TGY-2PMBT';    #Outlook 2013
        '4NT99-8RJFH-Q2VDH-KYG2C-4RD4F';    #PowerPoint 2013
        'PN2WF-29XG2-T9HJ7-JQPJR-FCXK4';    #Publisher 2013
        '6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7'     #Word 2013
    )
    foreach ($2013key in $2013keys) {
        cscript ospp.vbs /inpkey:$2013key | Out-Null
    }
}

function Office2013-V2R {
    Write-Host "Converting from Retail to Volume..." -ForegroundColor Green
    $inslics = Get-ChildItem -Path "$env:temp\tmp\library" | Where-Object {$_.Name -like "$licName"}
    foreach ($inslic in $inslics){
        cscript ospp.vbs /inslic:"$env:temp\tmp\library\$($inslic.Name)" | Out-Null
    }
    cscript ospp.vbs /inpkey:$2013kmskey | Out-Null
}

$matchingOffice2013Retail = ($dstatus | Select-String -SimpleMatch "Office 15, RETAIL" | Measure-Object).Count

if (($dstatus | Select-String -SimpleMatch "OfficeProfessional" | Measure-Object).Count -gt 0 -and $matchingOffice2013Retail -gt 0 ) {$licName = 'ProPlusVL*'; $2013kmskey = 'YC7DK-G2NP3-2QQC3-J6H88-GVGXT'} #For Office 2013Retail (Pro)
if (($dstatus | Select-String -SimpleMatch "OfficeProPlus" | Measure-Object).Count -gt 0 -and $matchingOffice2013Retail -gt 0 ) {$licName = 'ProPlusVL*'; $2013kmskey = 'YC7DK-G2NP3-2QQC3-J6H88-GVGXT'} #For Office 2013Retail (Pro) (MSDN)
if (($dstatus | Select-String -SimpleMatch "OfficeStandard" | Measure-Object).Count -gt 0 -and $matchingOffice2013Retail -gt 0 ) {$licName = 'StandardVL*'; $2013kmskey = 'KBKQT-2NMXY-JJWGP-M62JB-92CD4'} #For Office 2013Retail (Standard)
if (($dstatus | Select-String -SimpleMatch "OfficeProjectPro" | Measure-Object).Count -gt 0 -and $matchingOffice2013Retail -gt 0 ) {$licName = 'ProjectProVL*'; $2013kmskey = 'FN8TT-7WMH6-2D4X9-M337T-2342K'} #For Office 2013Standalone (Project Pro)
if (($dstatus | Select-String -SimpleMatch "OfficeProjectStd" | Measure-Object).Count -gt 0 -and $matchingOffice2013Retail -gt 0 ) {$licName = 'ProjectStdVL*'; $2013kmskey = '6NTH3-CW976-3G3Y2-JK3TX-8QHTT'} #For Office 2013Standalone (Project Standard)
if (($dstatus | Select-String -SimpleMatch "OfficeVisioPro" | Measure-Object).Count -gt 0 -and $matchingOffice2013Retail -gt 0 ) {$licName = 'VisioProVL*'; $2013kmskey = 'C2FG9-N6J68-H8BTJ-BW3QX-RM3B3'} #For Office 2013Standalone (Visio Pro)
if (($dstatus | Select-String -SimpleMatch "OfficeVisioStd" | Measure-Object).Count -gt 0 -and $matchingOffice2013Retail -gt 0 ) {$licName = 'VisioStdVL*'; $2013kmskey = 'J484Y-4NKBF-W2HMG-DBMJC-PGWR7'} #For Office 2013Standalone (Visio Standard)
if (($dstatus | Select-String -SimpleMatch "OfficeWord" | Measure-Object).Count -gt 0 -and $matchingOffice2013Retail -gt 0 ) {$licName = 'WordVL*'; $2013kmskey = '6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7'} #For Office 2013Standalone (Word)
if (($dstatus | Select-String -SimpleMatch "OfficeExcel" | Measure-Object).Count -gt 0 -and $matchingOffice2013Retail -gt 0 ) {$licName = 'ExcelVL*'; $2013kmskey = 'VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB'} #For Office 2013Standalone (Excel)
if (($dstatus | Select-String -SimpleMatch "OfficePowerPoint" | Measure-Object).Count -gt 0 -and $matchingOffice2013Retail -gt 0 ) {$licName = 'PowerPointVL*'; $2013kmskey = '4NT99-8RJFH-Q2VDH-KYG2C-4RD4F'} #For Office 2013Standalone (PowerPoint)
if (($dstatus | Select-String -SimpleMatch "OfficeOutlook" | Measure-Object).Count -gt 0 -and $matchingOffice2013Retail -gt 0 ) {$licName = 'OutlookVL*'; $2013kmskey = 'QPN8Q-BJBTJ-334K3-93TGY-2PMBT'} #For Office 2013Standalone (Outlook)
if (($dstatus | Select-String -SimpleMatch "OfficeAccess" | Measure-Object).Count -gt 0 -and $matchingOffice2013Retail -gt 0 ) {$licName = 'AccessVL*'; $2013kmskey = 'NG2JY-H4JBT-HQXYP-78QH9-4JM2D'} #For Office 2013Standalone (Access)
if (($dstatus | Select-String -SimpleMatch "OfficePublisher" | Measure-Object).Count -gt 0 -and $matchingOffice2013Retail -gt 0 ) {$licName = 'PublisherVL*'; $2013kmskey = 'PN2WF-29XG2-T9HJ7-JQPJR-FCXK4'} #For Office 2013Standalone (Publisher)

if ($matchingOffice2013Retail -gt 0 ) {
    New-Item -Path $env:temp\tmp -ItemType Directory -Force | Out-Null
    (New-Object Net.WebClient).DownloadFile('https://filedn.com/lOX1R8Sv7vhpEG9Q77kMbn0/MSGANG/scripts/office/office2013/library.zip', "$env:temp\tmp\library.zip") | Out-Null
    $null = Expand-Archive "$env:temp\tmp\library.zip" "$env:temp\tmp\library" -Force
    Office2013-V2R 
}
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
cscript ospp.vbs /remhst | Out-Null
cscript ospp.vbs /sethst:kms.msgang.com | Out-Null
cscript //nologo ospp.vbs /act | Out-Null

$dstatus = Invoke-Expression -Command "cscript.exe ospp.vbs /dstatus"
Write-Host "Number of Activated Products: $(($dstatus | Select-String -SimpleMatch "LICENSED").Count)" -ForegroundColor Green