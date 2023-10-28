@echo off
color f0
mode con cols=98 lines=30
title Activate Windows license for FREE - msgang.com
cls
echo .....................................................................................
echo #Project: Activating Microsoft software products for FREE without additional software
echo .....................................................................................
echo #Supported products: Windows 7/8/10/11/2008/2008R2/2012/2012R2/2016/2019/2022
echo .....................................................................................
for /f "tokens=* delims== " %%i in ('"powershell -c (Get-CimInstance Win32_OperatingSystem).Caption"') do (set edition=%%i)
echo You're using: %edition%
echo .....................................................................................

::Microsoft Windows 10
if /i "%edition%" equ "Microsoft Windows 10 Home" (set productkey=TX9XD-98N7V-6WMQ6-BX7FG-H8Q99)
if /i "%edition%" equ "Microsoft Windows 10 Home N" (set productkey=3KHY7-WNT83-DGQKR-F7HPR-844BM)
if /i "%edition%" equ "Microsoft Windows 10 Home Single Language" (set productkey=7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH)
if /i "%edition%" equ "Microsoft Windows 10 Pro" (set productkey=W269N-WFGWX-YVC9B-4J6C9-T83GX)
if /i "%edition%" equ "Microsoft Windows 10 Pro N" (set productkey=MH37W-N47XK-V7XM9-C7227-GCQG9)
if /i "%edition%" equ "Microsoft Windows 10 Pro for Workstations" (set productkey=NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J)
if /i "%edition%" equ "Microsoft Windows 10 Pro N for Workstations" (set productkey=9FNHH-K3HBT-3W4TD-6383H-6XYWF)
if /i "%edition%" equ "Microsoft Windows 10 Enterprise" (set productkey=NPPR9-FWDCX-D2C8J-H872K-2YT43)
if /i "%edition%" equ "Microsoft Windows 10 Enterprise N" (set productkey=DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4)
if /i "%edition%" equ "Microsoft Windows 10 Education" (set productkey=NW6C2-QMPVW-D7KKK-3GKT6-VCFB2)
if /i "%edition%" equ "Microsoft Windows 10 Education N" (set productkey=2WH4N-8QGBV-H22JP-CT43Q-MDWWJ)
if /i "%edition%" equ "Microsoft Windows 10 Enterprise LTSC 2019" (set productkey=M7XTQ-FN8P6-TTKYV-9D4CC-J462D)
if /i "%edition%" equ "Microsoft Windows 10 Enterprise LTSC 2021" (set productkey=M7XTQ-FN8P6-TTKYV-9D4CC-J462D)
if /i "%edition%" equ "Microsoft Windows 10 Enterprise LTSB 2016" (set productkey=DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ)
if /i "%edition%" equ "Microsoft Windows 10 Enterprise 2015 LTSB" (set productkey=WNMTR-4C88C-JK8YV-HQ7T2-76DF9)
if /i "%edition%" equ "Microsoft Windows 10 Enterprise Evaluation" (set productkey=NPPR9-FWDCX-D2C8J-H872K-2YT43)

::Microsoft Windows 11
if /i "%edition%" equ "Microsoft Windows 11 Home" (set productkey=TX9XD-98N7V-6WMQ6-BX7FG-H8Q99)
if /i "%edition%" equ "Microsoft Windows 11 Home N" (set productkey=3KHY7-WNT83-DGQKR-F7HPR-844BM)
if /i "%edition%" equ "Microsoft Windows 11 Home Single Language" (set productkey=7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH)
if /i "%edition%" equ "Microsoft Windows 11 Pro" (set productkey=W269N-WFGWX-YVC9B-4J6C9-T83GX)
if /i "%edition%" equ "Microsoft Windows 11 Pro N" (set productkey=MH37W-N47XK-V7XM9-C7227-GCQG9)
if /i "%edition%" equ "Microsoft Windows 11 Pro for Workstations" (set productkey=NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J)
if /i "%edition%" equ "Microsoft Windows 11 Pro N for Workstations" (set productkey=9FNHH-K3HBT-3W4TD-6383H-6XYWF)
if /i "%edition%" equ "Microsoft Windows 11 Enterprise" (set productkey=NPPR9-FWDCX-D2C8J-H872K-2YT43)
if /i "%edition%" equ "Microsoft Windows 11 Enterprise N" (set productkey=DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4)
if /i "%edition%" equ "Microsoft Windows 11 Education" (set productkey=NW6C2-QMPVW-D7KKK-3GKT6-VCFB2)
if /i "%edition%" equ "Microsoft Windows 11 Education N" (set productkey=2WH4N-8QGBV-H22JP-CT43Q-MDWWJ)
if /i "%edition%" equ "Microsoft Windows 11 Enterprise Evaluation" (set productkey=NPPR9-FWDCX-D2C8J-H872K-2YT43)

::Microsoft Windows Server 2008
if /i "%edition%" equ "Microsoft Windows Server 2008 Standard" (set productkey=TM24T-X9RMF-VWXK6-X8JC9-BFGM2)
if /i "%edition%" equ "Microsoft Windows Server 2008 Enterprise" (set productkey=YQGMW-MPWTJ-34KDK-48M3W-X4Q6V)
if /i "%edition%" equ "Microsoft Windows Server 2008 Datacenter" (set productkey=7M67G-PC374-GR742-YH8V4-TCBY3)

::Microsoft Windows Server 2008 R2
if /i "%edition%" equ "Microsoft Windows Server 2008 R2 Standard" (set productkey=YC6KT-GKW9T-YTKYR-T4X34-R7VHC)
if /i "%edition%" equ "Microsoft Windows Server 2008 R2 Enterprise" (set productkey=489J6-VHDMP-X63PK-3K798-CPX3Y)
if /i "%edition%" equ "Microsoft Windows Server 2008 R2 Datacenter" (set productkey=74YFP-3QFB3-KQT8W-PMXWJ-7M648)

::Microsoft Windows Server 2012
if /i "%edition%" equ "Microsoft Windows Server 2012" (set productkey=BN3D2-R7TKB-3YPBD-8DRP2-27GG4)
if /i "%edition%" equ "Microsoft Windows Server 2012 Essentials" (set productkey=HTDQM-NBMMG-KGYDT-2DTKT-J2MPV)
if /i "%edition%" equ "Microsoft Windows Server 2012 Standard" (set productkey=XC9B7-NBPP2-83J2H-RHMBY-92BT4)
if /i "%edition%" equ "Microsoft Windows Server 2012 Datacenter" (set productkey=48HP8-DN98B-MYWDG-T2DCC-8W83P)

::Microsoft Windows Server 2012 R2
if /i "%edition%" equ "Microsoft Windows Server 2012 Essentials" (set productkey=KNC87-3J2TX-XB4WP-VCPJV-M4FWM)
if /i "%edition%" equ "Microsoft Windows Server 2012 R2 Standard" (set productkey=D2N9P-3P6X9-2R39C-7RTCD-MDVJX)
if /i "%edition%" equ "Microsoft Windows Server 2012 R2 Datacenter" (set productkey=W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9)

::Microsoft Windows Server 2016
if /i "%edition%" equ "Microsoft Windows Server 2016 Essentials" (set productkey=JCKRF-N37P4-C2D82-9YXRT-4M63B)
if /i "%edition%" equ "Microsoft Windows Server 2016 Standard" (set productkey=WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY)
if /i "%edition%" equ "Microsoft Windows Server 2016 Datacenter" (set productkey=CB7KF-BWN84-R7R2Y-793K2-8XDDG)

::Microsoft Windows Server 2019
if /i "%edition%" equ "Microsoft Windows Server 2019 Essentials" (set productkey=WVDHN-86M7X-466P6-VHXV7-YY726)
if /i "%edition%" equ "Microsoft Windows Server 2019 Standard" (set productkey=N69G4-B89J2-4G8F4-WWYCC-J464C)
if /i "%edition%" equ "Microsoft Windows Server 2019 Datacenter" (set productkey=WMDGN-G9PQG-XVVXX-R3X43-63DFG)

::Microsoft Windows Server 2022
if /i "%edition%" equ "Microsoft Windows Server 2022 Standard" (set productkey=VDYBN-27WPP-V4HQT-9VMD4-VMK7H)
if /i "%edition%" equ "Microsoft Windows Server 2022 Datacenter" (set productkey=WX4NM-KYWYW-QJJR4-XV3QB-6VM33)

::Microsoft Windows 8
if /i "%edition%" equ "Microsoft Windows 8 Pro" (set productkey=NG4HW-VH26C-733KW-K6F98-J8CK4)
if /i "%edition%" equ "Microsoft Windows 8 Enterprise" (set productkey=32JNW-9KQ84-P47T8-D8GGY-CWCK7)

::Microsoft Windows 8.1
if /i "%edition%" equ "Microsoft Windows 8.1 Pro" (set productkey=GCRJD-8NW9H-F2CDX-CCM8D-9D6T9)
if /i "%edition%" equ "Microsoft Windows 8.1 Enterprise" (set productkey=MHF9N-XY6XB-WVXMC-BTDCT-MKKG7)

::Windows Server versions 20H2, 2004, 1909, 1903, and 1809 (Semi-Annual Channel versions)
if /i "%edition%" equ "Microsoft Windows Server Standard" (set productkey=VDYBN-27WPP-V4HQT-9VMD4-VMK7H)
if /i "%edition%" equ "Microsoft Windows Server Datacenter" (set productkey=WX4NM-KYWYW-QJJR4-XV3QB-6VM33)

::Microsoft Windows 7
wmic os get caption | find /v "Caption" > %temp%\ver.txt
set /p edition=<%temp%\ver.txt

echo.%edition% | findstr /C:"Microsoft Windows 7 Professional" >nul 2>&1
if not errorlevel 1 (set productkey=FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4) else (break)
echo.%edition% | findstr /C:"Microsoft Windows 7 Enterprise" >nul 2>&1
if not errorlevel 1 (set productkey=33PXH-7Y6KF-2VJC9-XBBR8-HVTHH) else (break)

echo .....................................................................................
echo Activating your Windows using product key: %productkey%

cscript %windir%\system32\slmgr.vbs /upk >nul 2>&1
cscript %windir%\system32\slmgr.vbs /ckms >nul 2>&1
cscript %windir%\system32\slmgr.vbs /cpky >nul 2>&1
cscript %windir%\system32\slmgr.vbs /skms kms.msgang.com >nul 2>&1
cscript %windir%\system32\slmgr.vbs /ipk %productkey% >nul 2>&1
cscript %windir%\system32\slmgr.vbs /ato | find /i "successfully" 

echo .....................................................................................
echo Your Windows license details:
echo.
cscript %windir%\system32\slmgr.vbs /dlv | find /i "Description"
cscript %windir%\system32\slmgr.vbs /dlv | find /i "Licensed"
cscript %windir%\system32\slmgr.vbs /dlv | find /i "Channel:"
cscript %windir%\system32\slmgr.vbs /dlv | find /i "Partial"
cscript %windir%\system32\slmgr.vbs /dlv | find /i "expiration"
echo.

echo Press any key to close this window.
pause >nul
