$MainMenu = {
    Write-Host " ******************************************************"
    Write-Host " * Author: https://msgang.com                         *" 
    Write-Host " * Updated on: 10/02/2023                             *" 
    Write-Host " ******************************************************" 
    Write-Host " 1. Professional (Activate and Convert)" 
    Write-Host " 2. Enterprise (Activate and Convert)" 
    Write-Host " 3. ProfessionalWorkstation (Activate and Convert)" 
    Write-Host " 4. Education (Activate and Convert)" 
    Write-Host " ******************************************************" 
    Write-Host " 5. Windows Server" 
    Write-Host " ******************************************************" 
    Write-Host " Information:"
    Write-Host "  - This script works on both Windows 10 and Windows 11."
    Write-Host "  - You can activate or switch Windows edition."
    Write-Host "  - Press Ctrl + C to Quit."
    Write-Host " ******************************************************" 
    Write-Host 
    Write-Host " Select an option and press Enter: "  -nonewline
    }
    cls

$MenuServer = {
    Write-Host " *******************************************"
    Write-Host " *                  Menu                   *" 
    Write-Host " *******************************************" 
    Write-Host " 1. Standard | Windows Server 2016" 
    Write-Host " 2. Standard | Windows Server 2019" 
    Write-Host " 3. Standard | Windows Server 2022" 
    Write-Host " *******************************************" 
    Write-Host " 4. Datacenter | Windows Server 2016" 
    Write-Host " 5. Datacenter | Windows Server 2019" 
    Write-Host " 6. Datacenter | Windows Server 2022" 
    Write-Host " 7. Quit"
    Write-Host 
    Write-Host " Select an option and press Enter: "  -nonewline
    }
    
$windows = {

    Write-Host
    Write-Host " Processing............"
    $null = New-Item -Path $env:temp\kms -ItemType Directory -Force
    Set-Location $env:temp\kms
    $uri = "https://s3.amazonaws.com/s3.bonben365.com/files/zip/windows-converter/windows10/$edition.zip"
    $ProgressPreference='Silent'
    Invoke-WebRequest -Uri $uri -OutFile "$edition.zip" -ErrorAction:SilentlyContinue
    Expand-Archive "$edition.zip" $edition -Force -ErrorAction:SilentlyContinue
    Copy-Item "$env:temp\kms\$edition\$edition" "$env:windir\system32\spp\tokens\skus\" -ErrorAction:SilentlyContinue
    Copy-Item "$env:temp\kms\$edition\$edition\*" "$env:windir\system32\spp\tokens\skus\$edition" -ErrorAction:SilentlyContinue
    
    Set-Location "$env:windir\system32"
    $null = cmd.exe /c cscript.exe slmgr.vbs /rilc
    $null = cmd.exe /c cscript.exe slmgr.vbs /upk
    $null = cmd.exe /c cscript.exe slmgr.vbs /ckms
    $null = cmd.exe /c cscript.exe slmgr.vbs /cpky
    $null = cmd.exe /c cscript.exe slmgr.vbs /skms kms.msgang.com
    $null = cmd.exe /c cscript.exe slmgr.vbs /ipk $productkey
    $null = cmd.exe /c cscript.exe slmgr.vbs /ato
    Write-Host
    Write-Host " Done............"
    Write-Host
    Write-Host " Your Windows edition: $((Get-ComputerInfo).WindowsProductName)"
    $licStatus = cmd.exe /c cscript.exe slmgr.vbs /dlv | find "License Status"
    $actid = cmd.exe /c cscript.exe slmgr.vbs /dlv | find "Activation ID"
    $pkc = cmd.exe /c cscript.exe slmgr.vbs /dlv | find "Product Key Channel"
    $iid = cmd.exe /c cscript.exe slmgr.vbs /dlv | find "Installation ID"
    Write-Host " $actid"
    Write-Host " $pkc"
    Write-Host " $licStatus"
    Write-Host " $iid"
    Start-Sleep -Seconds 30
}

$windowsserver = {

    Write-Host
    Write-Host " Processing............"
    Set-Location "$env:windir\system32"
    $null = cmd.exe /c cscript.exe slmgr.vbs /rilc
    $null = cmd.exe /c cscript.exe slmgr.vbs /upk
    $null = cmd.exe /c cscript.exe slmgr.vbs /ckms
    $null = cmd.exe /c cscript.exe slmgr.vbs /cpky
    $null = cmd.exe /c cscript.exe slmgr.vbs /skms kms.msgang.com
    $null = cmd.exe /c cscript.exe slmgr.vbs /ipk $productkeystd
    $null = cmd.exe /c cscript.exe slmgr.vbs /ipk $productkeydc
    $null = cmd.exe /c cscript.exe slmgr.vbs /ato
    Write-Host
    Write-Host " Done............"
    Write-Host
    Write-Host " Your Windows edition: $((Get-ComputerInfo).WindowsProductName)"
    $licStatus = cmd.exe /c cscript.exe slmgr.vbs /dlv | find "License Status"
    $actid = cmd.exe /c cscript.exe slmgr.vbs /dlv | find "Activation ID"
    $pkc = cmd.exe /c cscript.exe slmgr.vbs /dlv | find "Product Key Channel"
    $iid = cmd.exe /c cscript.exe slmgr.vbs /dlv | find "Installation ID"
    Write-Host " $actid"
    Write-Host " $pkc"
    Write-Host " $licStatus"
    Write-Host " $iid"
    Start-Sleep -Seconds 30
}
    
    Do { 
    cls
    Invoke-Command $MainMenu
    $select = Read-Host

    if ($select -eq 1) {$edition = "Professional"; $productkey = "W269N-WFGWX-YVC9B-4J6C9-T83GX"}
    if ($select -eq 2) {$edition = "Enterprise"; $productkey ="NPPR9-FWDCX-D2C8J-H872K-2YT43"}
    if ($select -eq 3) {$edition = "ProfessionalWorkstation"; $productkey = "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J"}
    if ($select -eq 4) {$edition = "Education"; $productkey = "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"}

    Switch ($select)
       {
            1 { Invoke-Command $windows }
            2 { Invoke-Command $windows }
            3 { Invoke-Command $windows }
            4 { Invoke-Command $windows }
            5 {
                Do { 
                    cls
                    Invoke-Command $MenuServer
                    $selectServer = Read-Host
  
                    if ($selectServer -eq 1) {$productkeystd = 'WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY';$productkeydc = 'CB7KF-BWN84-R7R2Y-793K2-8XDDG'} #Windows Server 2016 Standard
                    if ($selectServer -eq 2) {$productkey = 'N69G4-B89J2-4G8F4-WWYCC-J464C'} #Windows Server 2019 Standard
                    if ($selectServer -eq 3) {$productkey = 'VDYBN-27WPP-V4HQT-9VMD4-VMK7H'} #Windows Server 2022 Standard
                    if ($selectServer -eq 4) {$productkey = 'CB7KF-BWN84-R7R2Y-793K2-8XDDG'} #Windows Server 2016 Datacenter
                    if ($selectServer -eq 5) {$productkey = 'WMDGN-G9PQG-XVVXX-R3X43-63DFG'} #Windows Server 2019 Datacenter
                    if ($selectServer -eq 6) {$productkey = 'WX4NM-KYWYW-QJJR4-XV3QB-6VM33'} #Windows Server 2022 Datacenter
  
                    Switch ($selectServer)
                       {
                       1 {Invoke-Command $windowsserver}
                       2 {Invoke-Command $windowsserver}
                       3 {Invoke-Command $windowsserver}
                       4 {Invoke-Command $windowsserver}
                       5 {Invoke-Command $windowsserver}
                       6 {Invoke-Command $windowsserver}
                       }
                    }
                    While ($selectServer -ne 7)
                    cls
            }
        }
    }
    While ($select -ne 5)
