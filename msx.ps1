$MainMenu = {
    Write-Host " ******************************************************"
    Write-Host " * Author: bonguides.com                              *" 
    Write-Host " * Updated on: 10/16/2022                             *" 
    Write-Host " ******************************************************" 
    Write-Host 
    Write-Host " 1. Convert to Professional" 
    Write-Host " 2. Convert to Enterprise"
    Write-Host " 3. Convert to ProfessionalWorkstation"
    Write-Host " 4. Convert to Education"
    Write-Host
    Write-Host " ******************************************************" 
    Write-Host " Press Ctrl + C to Quit"
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
        Write-Host "Processing............"
        $null = New-Item -Path $env:temp\kms -ItemType Directory -Force
        Set-Location $env:temp\kms
        $uri = "https://s3.amazonaws.com/s3.bonben365.com/files/zip/windows-converter/windows10/$edition.zip"
        $null = Invoke-WebRequest -Uri $uri -OutFile "$edition.zip" -ErrorAction:SilentlyContinue
        $null = Expand-Archive "$edition.zip" $edition -Force -ErrorAction:SilentlyContinue
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
        Write-Host "Done............"
        Write-Host "Your Windows edition: $((Get-ComputerInfo).WindowsProductName)"
        Start-Sleep -Seconds 3
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
