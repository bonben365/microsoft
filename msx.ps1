$MainMenu = {
    Write-Host " ******************************************************"
    Write-Host " * Author: https://msgang.com                         *" 
    Write-Host " * Updated on: 10/02/2023                             *" 
    Write-Host " ******************************************************" 
    Write-Host 
    Write-Host " 1. Activate / Convert to Professional Edition" 
    Write-Host " 2. Activate / Convert to Enterprise Edition"
    Write-Host " 3. Activate / Convert to ProfessionalWorkstation Edition"
    Write-Host " 4. Activate / Convert to Education Edition"
    Write-Host
    Write-Host " Information:"
    Write-Host 
    Write-Host "  - This script works on both Windows 10 and Windows 11."
    Write-Host "  - You can activate or switch Windows edition."
    Write-Host "  - Press Ctrl + C to Quit."
    Write-Host
    Write-Host " ******************************************************" 
    Write-Host 
    Write-Host " Select an option and press Enter: "  -nonewline
    }
    cls
    
    Do { 
    cls
    Invoke-Command $MainMenu
    $select = Read-Host

    if ($select -eq 1) {$edition = "Professional"; $productkey = "W269N-WFGWX-YVC9B-4J6C9-T83GX"}
    if ($select -eq 2) {$edition = "Enterprise"; $productkey ="NPPR9-FWDCX-D2C8J-H872K-2YT43"}
    if ($select -eq 3) {$edition = "ProfessionalWorkstation"; $productkey = "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J"}
    if ($select -eq 4) {$edition = "Education"; $productkey = "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"}

    $bonben = {
        
        Write-Host
        Write-Host "Processing............"
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
        Write-Host "Done............"
        Write-Host
        Write-Host "Your Windows edition: $((Get-ComputerInfo).WindowsProductName)"
        Start-Sleep -Seconds 30
    }

    Switch ($select)
       {
            1 { Invoke-Command $bonben }
            2 { Invoke-Command $bonben }
            3 { Invoke-Command $bonben }
            4 { Invoke-Command $bonben }
    
        }
    }
    While ($select -ne 5)
