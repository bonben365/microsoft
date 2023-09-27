#===============================================================
#Name:           Windows Converter.
#Description:    Convert all Windows Editions for free.
#Version:        1.2
#Date :          08/8/2023
#Website:        https://msgang.com
#Script by:      Leo Nguyen
#For detailed script execution: https://msgang.com/windows
#===============================================================


if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework")
[void] [Reflection.Assembly]::LoadWithPartialName("PresentationCore")

$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(890,470)
$Form.StartPosition = "CenterScreen" #loads the window in the center of the screen
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow #modifies the window border
$Form.Text = "Microsoft Windows Converter - www.msgang.com" #window description
$Form.ShowInTaskbar = $True
$Form.KeyPreview = $True
$Form.AutoSize = $True
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $False
$Form.MinimizeBox = $False
$Form.ControlBox = $True
$Form.Icon = $Icon

$convert = {
    
   Write-Host ===============================================================
   Write-Host Name:           Windows Converter.
   Write-Host Description:    Convert all Windows Editions for free.
   Write-Host Version:        1.0
   Write-Host Date :          26/7/2023
   Write-Host Website:        https://msgang.com
   Write-Host Script by:      Leo Nguyen
   Write-Host For detailed script execution: https://msgang.com/windows
   Write-Host ===============================================================

   Write-Host '---------------------------------------------------------------'
   Write-Host "Processing...It could take a while, please be patient."                 
   Write-Host
   $version = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
   New-Item -Path $env:temp\temp\$skuid -ItemType Directory -Force
   Set-Location $env:temp\temp\$skuid
   $uri = "https://raw.githubusercontent.com/bonben365/microsoft/main/Files/$skuid.zip"
   $FilePath = "$env:temp\temp\$skuid\$skuid.zip"
   (New-Object Net.WebClient).DownloadFile($uri, $FilePath)
   Expand-Archive .\*.zip -DestinationPath . -Force | Out-Null
   Copy-Item -Path $sku $env:windir\system32\spp\tokens\skus\ -Recurse -Force -ErrorAction SilentlyContinue
   cscript.exe $env:windir\system32\slmgr.vbs /rilc | Out-Null
   cscript.exe $env:windir\system32\slmgr.vbs /upk | Out-Null
   cscript.exe $env:windir\system32\slmgr.vbs /ckms | Out-Null
   cscript.exe $env:windir\system32\slmgr.vbs /cpky | Out-Null
   cscript.exe $env:windir\system32\slmgr.vbs /skms kms.msgang.com | Out-Null
   cscript.exe $env:windir\system32\slmgr.vbs /ipk $key
   cscript.exe $env:windir\system32\slmgr.vbs /ato
   Write-Host
   Write-Host "Done............" -ForegroundColor Green 
   Write-Host
   Write-Host "Before Upgrading : $version" -ForegroundColor Yellow
   Write-Host "After Upgrading  : $((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName)" -ForegroundColor Yellow
   Write-Host

   #Cleanup
   #Set-Location $env:temp
   #Remove-Item -Path $env:temp\temp -Recurse -Force
}
############################################## Start functions
   function microsoftInstaller {
   try {

   if ($10Home.Checked -eq $true) {$sku = 'Core'; $skuid = 'skus'; $key = 'TX9XD-98N7V-6WMQ6-BX7FG-H8Q99'; Invoke-Command $convert}
   if ($10HomeSL.Checked -eq $true) {$sku = 'CoreSingleLanguage'; $skuid = 'skus'; $key = '7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH'; Invoke-Command $convert}
   if ($10Pro.Checked -eq $true) {$sku = 'Professional'; $skuid = 'skus'; $key = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'; Invoke-Command $convert}
   if ($10ProWorkstation.Checked -eq $true) {$sku = 'ProfessionalWorkstation'; $skuid = 'skus'; $key = 'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J'; Invoke-Command $convert}
   if ($10Enterprise.Checked -eq $true) {$sku = 'Enterprise'; $skuid = 'skus'; $key = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'; Invoke-Command $convert}
   if ($10Education.Checked -eq $true) {$sku = 'Education'; $skuid = 'skus'; $key = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2'; Invoke-Command $convert}
   if ($10ProEducation.Checked -eq $true) {$sku = 'ProfessionalEducation'; $skuid = 'skus'; $key = '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y'; Invoke-Command $convert}

   if ($11Home.Checked -eq $true) {$sku = 'Core'; $skuid = 'sku11'; $key = 'TX9XD-98N7V-6WMQ6-BX7FG-H8Q99'; Invoke-Command $convert}
   if ($11HomeSL.Checked -eq $true) {$sku = 'CoreSingleLanguage'; $skuid = 'sku11'; $key = '7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH'; Invoke-Command $convert}
   if ($11Pro.Checked -eq $true) {$sku = 'Professional'; $skuid = 'sku11'; $key = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'; Invoke-Command $convert}
   if ($11ProWorkstation.Checked -eq $true) {$sku = 'ProfessionalWorkstation'; $skuid = 'sku11'; $key = 'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J'; Invoke-Command $convert}
   if ($11Enterprise.Checked -eq $true) {$sku = 'Enterprise'; $skuid = 'sku11'; $key = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'; Invoke-Command $convert}
   if ($11Education.Checked -eq $true) {$sku = 'Education'; $skuid = 'sku11'; $key = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2'; Invoke-Command $convert}
   if ($11ProEducation.Checked -eq $true) {$sku = 'ProfessionalEducation'; $skuid = 'sku11'; $key = '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y'; Invoke-Command $convert}

   if ($10Eval2Pro.Checked -eq $true) {$sku = 'Professional'; $skuid = 'skus'; $key = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'; Invoke-Command $convert}
   if ($10Eval2Enterprise.Checked -eq $true) {$sku = 'Enterprise'; $skuid = 'skus'; $key = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'; Invoke-Command $convert}
   if ($10Eval2ProWorkstation.Checked -eq $true) {$sku = 'ProfessionalWorkstation'; $skuid = 'skus'; $key = 'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J'; Invoke-Command $convert}
   if ($10Eval2LTSC2019Full.Checked -eq $true) {$sku = 'EnterpriseS'; $skuid = 'ltsc2019'; $key = 'M7XTQ-FN8P6-TTKYV-9D4CC-J462D'; Invoke-Command $convert}
   if ($10Eval2LTSC2021Full.Checked -eq $true) {$sku = 'EnterpriseS'; $skuid = 'ltsc2021'; $key = 'M7XTQ-FN8P6-TTKYV-9D4CC-J462D'; Invoke-Command $convert}

   if ($10LTSCEval2Pro.Checked -eq $true) {$sku = 'Professional'; $skuid = 'skus'; $key = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'; Invoke-Command $convert}
   if ($10LTSCEval2Enterprise.Checked -eq $true) {$sku = 'Enterprise'; $skuid = 'skus'; $key = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'; Invoke-Command $convert}
   if ($10LTSCEval2ProWorkstation.Checked -eq $true) {$sku = 'ProfessionalWorkstation'; $skuid = 'skus'; $key = 'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J'; Invoke-Command $convert}
   if ($10LTSCEval2LTSC2019Full.Checked -eq $true) {$sku = 'EnterpriseS'; $skuid = 'ltsc2019'; $key = 'M7XTQ-FN8P6-TTKYV-9D4CC-J462D'; Invoke-Command $convert}
   if ($10LTSCEval2LTSC2021Full.Checked -eq $true) {$sku = 'EnterpriseS'; $skuid = 'ltsc2021'; $key = 'M7XTQ-FN8P6-TTKYV-9D4CC-J462D'; Invoke-Command $convert}

   if ($10Eval2LTSC2019Full.Checked -eq $true) {$sku = 'ServerStandard'; $skuid = 'srv2016'; $key = 'WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY'; Invoke-Command $convert}
   if ($srv2016data.Checked -eq $true) {$sku = 'ServerDatacenter'; $skuid = 'srv2016'; $key = 'CB7KF-BWN84-R7R2Y-793K2-8XDDG'; Invoke-Command $convert}
   if ($srv2016eval2std.Checked -eq $true) {$sku = 'ServerStandard'; $skuid = 'srv2016'; $key = 'WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY'; Invoke-Command $convert}
   if ($srv2016eval2data.Checked -eq $true) {$sku = 'ServerDatacenter'; $skuid = 'srv2016'; $key = 'CB7KF-BWN84-R7R2Y-793K2-8XDDG'; Invoke-Command $convert}

   if ($srv2019std.Checked -eq $true) {$sku = 'ServerStandard'; $skuid = 'srv2019'; $key = 'N69G4-B89J2-4G8F4-WWYCC-J464C'; Invoke-Command $convert}
   if ($srv2019data.Checked -eq $true) {$sku = 'ServerDatacenter'; $skuid = 'srv2019'; $key = 'WMDGN-G9PQG-XVVXX-R3X43-63DFG'; Invoke-Command $convert}
   if ($srv2019eval2std.Checked -eq $true) {$sku = 'ServerStandard'; $skuid = 'srv2019'; $key = 'N69G4-B89J2-4G8F4-WWYCC-J464C'; Invoke-Command $convert}
   if ($srv2019eval2data.Checked -eq $true) {$sku = 'ServerDatacenter'; $skuid = 'srv2019'; $key = 'WMDGN-G9PQG-XVVXX-R3X43-63DFG'; Invoke-Command $convert}

   if ($srv2022std.Checked -eq $true) {$sku = 'ServerStandard'; $skuid = 'srv2022'; $key = 'VDYBN-27WPP-V4HQT-9VMD4-VMK7H'; Invoke-Command $convert}
   if ($srv2022data.Checked -eq $true) {$sku = 'ServerDatacenter'; $skuid = 'srv2022'; $key = 'WX4NM-KYWYW-QJJR4-XV3QB-6VM33'; Invoke-Command $convert}
   if ($srv2022eval2std.Checked -eq $true) {$sku = 'ServerStandard'; $skuid = 'srv2022'; $key = 'VDYBN-27WPP-V4HQT-9VMD4-VMK7H'; Invoke-Command $convert}
   if ($srv2022eval2data.Checked -eq $true) {$sku = 'ServerDatacenter'; $skuid = 'srv2022'; $key = 'WX4NM-KYWYW-QJJR4-XV3QB-6VM33'; Invoke-Command $convert}
   
   } #end try
   catch {$outputBox.text = "`nOperation could not be completed"}
} 
############################################## end functions
############################################## Start group boxes

   $groupbox10 = New-Object System.Windows.Forms.GroupBox
   $groupbox10.Location = New-Object System.Drawing.Size(10,10) 
   $groupbox10.size = New-Object System.Drawing.Size(170,180) 
   $groupbox10.text = "Windows 10"
   $Form.Controls.Add($groupbox10)

   $groupbox11 = New-Object System.Windows.Forms.GroupBox
   $groupbox11.Location = New-Object System.Drawing.Size(190,10) 
   $groupbox11.size = New-Object System.Drawing.Size(170,180) 
   $groupbox11.text = "Windows 11"
   $Form.Controls.Add($groupbox11)
   
   $groupboxeval = New-Object System.Windows.Forms.GroupBox
   $groupboxeval.Location = New-Object System.Drawing.Size(370,10) 
   $groupboxeval.size = New-Object System.Drawing.Size(220,180) 
   $groupboxeval.text = "Windows 10 Enterprise Evaluation to:"
   $Form.Controls.Add($groupboxeval)

   $groupbox10LTSCeval = New-Object System.Windows.Forms.GroupBox
   $groupbox10LTSCeval.Location = New-Object System.Drawing.Size(600,10) 
   $groupbox10LTSCeval.size = New-Object System.Drawing.Size(250,180) 
   $groupbox10LTSCeval.text = "Windows 10 Enterprise LTSC Evaluation to:"
   $Form.Controls.Add($groupbox10LTSCeval)

   $groupboxsrv2016 = New-Object System.Windows.Forms.GroupBox
   $groupboxsrv2016.Location = New-Object System.Drawing.Size(10,200) 
   $groupboxsrv2016.size = New-Object System.Drawing.Size(170,130) 
   $groupboxsrv2016.text = "Windows Server 2016"
   $Form.Controls.Add($groupboxsrv2016)

   $groupboxsrv2019 = New-Object System.Windows.Forms.GroupBox
   $groupboxsrv2019.Location = New-Object System.Drawing.Size(190,200) 
   $groupboxsrv2019.size = New-Object System.Drawing.Size(170,130) 
   $groupboxsrv2019.text = "Windows Server 2019"
   $Form.Controls.Add($groupboxsrv2019)

   $groupboxsrv2022 = New-Object System.Windows.Forms.GroupBox
   $groupboxsrv2022.Location = New-Object System.Drawing.Size(370,200) 
   $groupboxsrv2022.size = New-Object System.Drawing.Size(170,130) 
   $groupboxsrv2022.text = "Windows Server 2022"
   $Form.Controls.Add($groupboxsrv2022)

   # label
   $objLabel = New-Object System.Windows.Forms.label
   $objLabel.Location = New-Object System.Drawing.Size(10,350)
   $objLabel.Size = New-Object System.Drawing.Size(400,15)
   $objLabel.Text = "(*)You can convert, upgrade or switch from or to any editions of Windows."
   $Form.Controls.Add($objLabel)

############################################## end group boxes
############################################## Start Windows 10 checkboxes
   $10Home = New-Object System.Windows.Forms.RadioButton
   $10Home.Location = New-Object System.Drawing.Size(10,20)
   $10Home.Size = New-Object System.Drawing.Size(140,20)
   $10Home.Checked = $false
   $10Home.Text = "Home"
   $groupbox10.Controls.Add($10Home)

   $10HomeSL = New-Object System.Windows.Forms.RadioButton
   $10HomeSL.Location = New-Object System.Drawing.Size(10,40)
   $10HomeSL.Size = New-Object System.Drawing.Size(140,20)
   $10HomeSL.Text = "Home Single Language"
   $groupbox10.Controls.Add($10HomeSL)

   $10Pro = New-Object System.Windows.Forms.RadioButton
   $10Pro.Location = New-Object System.Drawing.Size(10,60)
   $10Pro.Size = New-Object System.Drawing.Size(140,20)
   $10Pro.Text = "Pro"
   $groupbox10.Controls.Add($10Pro)

   $10ProWorkstation = New-Object System.Windows.Forms.RadioButton
   $10ProWorkstation.Location = New-Object System.Drawing.Size(10,80)
   $10ProWorkstation.Size = New-Object System.Drawing.Size(140,20)
   $10ProWorkstation.AutoSize = $true
   $10ProWorkstation.Text = "Pro for Workstation"
   $groupbox10.Controls.Add($10ProWorkstation)

   $10Enterprise = New-Object System.Windows.Forms.RadioButton
   $10Enterprise.Location = New-Object System.Drawing.Size(10,100)
   $10Enterprise.Size = New-Object System.Drawing.Size(140,20)
   $10Enterprise.Text = "Enterprise"
   $groupbox10.Controls.Add($10Enterprise)

   $10Education = New-Object System.Windows.Forms.RadioButton
   $10Education.Location = New-Object System.Drawing.Size(10,120)
   $10Education.Size = New-Object System.Drawing.Size(140,20)
   $10Education.Text = "Education"
   $groupbox10.Controls.Add($10Education)

   $10ProEducation = New-Object System.Windows.Forms.RadioButton
   $10ProEducation.Location = New-Object System.Drawing.Size(10,140)
   $10ProEducation.Size = New-Object System.Drawing.Size(140,20)
   $10ProEducation.Text = "Pro Education"
   $groupbox10.Controls.Add($10ProEducation)
############################################## End Windows 10 checkboxes
############################################## Start Windows 11 checkboxes
   $11Home = New-Object System.Windows.Forms.RadioButton
   $11Home.Location = New-Object System.Drawing.Size(10,20)
   $11Home.Size = New-Object System.Drawing.Size(140,20)
   $11Home.Checked = $false
   $11Home.Text = "Home"
   $groupbox11.Controls.Add($11Home)

   $11HomeSL = New-Object System.Windows.Forms.RadioButton
   $11HomeSL.Location = New-Object System.Drawing.Size(10,40)
   $11HomeSL.Size = New-Object System.Drawing.Size(140,20)
   $11HomeSL.Text = "Home Single Language"
   $groupbox11.Controls.Add($11HomeSL)

   $11Pro = New-Object System.Windows.Forms.RadioButton
   $11Pro.Location = New-Object System.Drawing.Size(10,60)
   $11Pro.Size = New-Object System.Drawing.Size(140,20)
   $11Pro.Text = "Pro"
   $groupbox11.Controls.Add($11Pro)

   $11ProWorskstation = New-Object System.Windows.Forms.RadioButton
   $11ProWorskstation.Location = New-Object System.Drawing.Size(10,80)
   $11ProWorskstation.Size = New-Object System.Drawing.Size(140,20)
   $11ProWorskstation.Text = "Pro for Workstation"
   $11ProWorskstation.AutoSize = $true
   $groupbox11.Controls.Add($11ProWorskstation)

   $11Enterprise = New-Object System.Windows.Forms.RadioButton
   $11Enterprise.Location = New-Object System.Drawing.Size(10,100)
   $11Enterprise.Size = New-Object System.Drawing.Size(140,20)
   $11Enterprise.Text = "Enterprise"
   $groupbox11.Controls.Add($11Enterprise)

   $11Education = New-Object System.Windows.Forms.RadioButton
   $11Education.Location = New-Object System.Drawing.Size(10,120)
   $11Education.Size = New-Object System.Drawing.Size(140,20)
   $11Education.Text = "Education"
   $groupbox11.Controls.Add($11Education)

   $11ProEducation = New-Object System.Windows.Forms.RadioButton
   $11ProEducation.Location = New-Object System.Drawing.Size(10,140)
   $11ProEducation.Size = New-Object System.Drawing.Size(140,20)
   $11ProEducation.Text = "Pro Education"
   $groupbox11.Controls.Add($11ProEducation)

############################################## End Windows 11 checkboxes
############################################## Start Server 2016 checkboxes
   $srv2016std = New-Object System.Windows.Forms.RadioButton
   $srv2016std.Location = New-Object System.Drawing.Size(10,20)
   $srv2016std.Size = New-Object System.Drawing.Size(140,20)
   $srv2016std.Checked = $false
   $srv2016std.Text = "Standard"
   $groupboxsrv2016.Controls.Add($srv2016std)

   $srv2016data = New-Object System.Windows.Forms.RadioButton
   $srv2016data.Location = New-Object System.Drawing.Size(10,40)
   $srv2016data.Size = New-Object System.Drawing.Size(140,20)
   $srv2016data.Text = "Datacenter"
   $groupboxsrv2016.Controls.Add($srv2016data)

   $srv2016eval2std = New-Object System.Windows.Forms.RadioButton
   $srv2016eval2std.Location = New-Object System.Drawing.Size(10,60)
   $srv2016eval2std.Size = New-Object System.Drawing.Size(140,20)
   $srv2016eval2std.Text = "Evaluation to Standard"
   $groupboxsrv2016.Controls.Add($srv2016eval2std)

   $srv2016eval2data = New-Object System.Windows.Forms.RadioButton
   $srv2016eval2data.Location = New-Object System.Drawing.Size(10,80)
   $srv2016eval2data.Size = New-Object System.Drawing.Size(150,20)
   $srv2016eval2data.Text = "Evaluation to Datacenter"
   $groupboxsrv2016.Controls.Add($srv2016eval2data)

############################################## End Server 2016 checkboxes
############################################## Start Server 2019 checkboxes
   $srv2019std = New-Object System.Windows.Forms.RadioButton
   $srv2019std.Location = New-Object System.Drawing.Size(10,20)
   $srv2019std.Size = New-Object System.Drawing.Size(140,20)
   $srv2019std.Checked = $false
   $srv2019std.Text = "Standard"
   $groupboxsrv2019.Controls.Add($srv2019std)

   $srv2019data = New-Object System.Windows.Forms.RadioButton
   $srv2019data.Location = New-Object System.Drawing.Size(10,40)
   $srv2019data.Size = New-Object System.Drawing.Size(140,20)
   $srv2019data.Text = "Datacenter"
   $groupboxsrv2019.Controls.Add($srv2019data)

   $srv2019eval2std = New-Object System.Windows.Forms.RadioButton
   $srv2019eval2std.Location = New-Object System.Drawing.Size(10,60)
   $srv2019eval2std.Size = New-Object System.Drawing.Size(140,20)
   $srv2019eval2std.Text = "Evaluation to Standard"
   $groupboxsrv2019.Controls.Add($srv2019eval2std)

   $srv2019eval2data = New-Object System.Windows.Forms.RadioButton
   $srv2019eval2data.Location = New-Object System.Drawing.Size(10,80)
   $srv2019eval2data.Size = New-Object System.Drawing.Size(150,20)
   $srv2019eval2data.Text = "Evaluation to Datacenter"
   $groupboxsrv2019.Controls.Add($srv2019eval2data)
############################################## End Server 2019 checkboxes
############################################## Start Server 2022 checkboxes
   $srv2022std = New-Object System.Windows.Forms.RadioButton
   $srv2022std.Location = New-Object System.Drawing.Size(10,20)
   $srv2022std.Size = New-Object System.Drawing.Size(140,20)
   $srv2022std.Checked = $false
   $srv2022std.Text = "Standard"
   $groupboxsrv2022.Controls.Add($srv2022std)

   $srv2022data = New-Object System.Windows.Forms.RadioButton
   $srv2022data.Location = New-Object System.Drawing.Size(10,40)
   $srv2022data.Size = New-Object System.Drawing.Size(140,20)
   $srv2022data.Text = "Datacenter"
   $groupboxsrv2022.Controls.Add($srv2022data)

   $srv2022eval2std = New-Object System.Windows.Forms.RadioButton
   $srv2022eval2std.Location = New-Object System.Drawing.Size(10,60)
   $srv2022eval2std.Size = New-Object System.Drawing.Size(140,20)
   $srv2022eval2std.Text = "Evaluation to Standard"
   $groupboxsrv2022.Controls.Add($srv2022eval2std)

   $srv2022eval2data = New-Object System.Windows.Forms.RadioButton
   $srv2022eval2data.Location = New-Object System.Drawing.Size(10,80)
   $srv2022eval2data.Size = New-Object System.Drawing.Size(150,20)
   $srv2022eval2data.Text = "Evaluation to Datacenter"
   $groupboxsrv2022.Controls.Add($srv2022eval2data)
############################################## End Server 2022 checkboxes
############################################## Start Windows 10 Eval checkboxes
   $10Eval2Pro = New-Object System.Windows.Forms.RadioButton
   $10Eval2Pro.Location = New-Object System.Drawing.Size(10,20)
   $10Eval2Pro.Size = New-Object System.Drawing.Size(140,20)
   $10Eval2Pro.Checked = $false
   $10Eval2Pro.Text = "Pro"
   $groupboxeval.Controls.Add($10Eval2Pro)

   $10Eval2Enterprise = New-Object System.Windows.Forms.RadioButton
   $10Eval2Enterprise.Location = New-Object System.Drawing.Size(10,40)
   $10Eval2Enterprise.Size = New-Object System.Drawing.Size(140,20)
   $10Eval2Enterprise.Checked = $false
   $10Eval2Enterprise.Text = "Enterprise"
   $groupboxeval.Controls.Add($10Eval2Enterprise)

   $10Eval2ProWorkstation = New-Object System.Windows.Forms.RadioButton
   $10Eval2ProWorkstation.Location = New-Object System.Drawing.Size(10,60)
   $10Eval2ProWorkstation.Size = New-Object System.Drawing.Size(140,20)
   $10Eval2ProWorkstation.Checked = $false
   $10Eval2ProWorkstation.Text = "Pro for Workstation"
   $groupboxeval.Controls.Add($10Eval2ProWorkstation)

   $10Eval2LTSC2019Full = New-Object System.Windows.Forms.RadioButton
   $10Eval2LTSC2019Full.Location = New-Object System.Drawing.Size(10,80)
   $10Eval2LTSC2019Full.Size = New-Object System.Drawing.Size(300,20)
   $10Eval2LTSC2019Full.Checked = $false
   $10Eval2LTSC2019Full.Text = "Enterprise LTSC 2019 (Full)"
   $groupboxeval.Controls.Add($10Eval2LTSC2019Full)

   $10Eval2LTSC2021Full = New-Object System.Windows.Forms.RadioButton
   $10Eval2LTSC2021Full.Location = New-Object System.Drawing.Size(10,100)
   $10Eval2LTSC2021Full.Size = New-Object System.Drawing.Size(300,20)
   $10Eval2LTSC2021Full.Checked = $false
   $10Eval2LTSC2021Full.Text = "Enterprise LTSC 2021 (Full)"
   $groupboxeval.Controls.Add($10Eval2LTSC2021Full)
############################################## End Windows 10 Eval checkboxes
############################################## Start Windows 10 LTSC Eval checkboxes

   $10LTSCEval2Pro = New-Object System.Windows.Forms.RadioButton
   $10LTSCEval2Pro.Location = New-Object System.Drawing.Size(10,20)
   $10LTSCEval2Pro.Size = New-Object System.Drawing.Size(140,20)
   $10LTSCEval2Pro.Checked = $false
   $10LTSCEval2Pro.Text = "Pro"
   $groupbox10LTSCeval.Controls.Add($10LTSCEval2Pro)


   $10LTSCEval2Enterprise = New-Object System.Windows.Forms.RadioButton
   $10LTSCEval2Enterprise.Location = New-Object System.Drawing.Size(10,40)
   $10LTSCEval2Enterprise.Size = New-Object System.Drawing.Size(140,20)
   $10LTSCEval2Enterprise.Checked = $false
   $10LTSCEval2Enterprise.Text = "Enterprise"
   $groupbox10LTSCeval.Controls.Add($10LTSCEval2Enterprise)

   $10LTSCEval2ProWorkstation = New-Object System.Windows.Forms.RadioButton
   $10LTSCEval2ProWorkstation.Location = New-Object System.Drawing.Size(10,60)
   $10LTSCEval2ProWorkstation.Size = New-Object System.Drawing.Size(140,20)
   $10LTSCEval2ProWorkstation.Checked = $false
   $10LTSCEval2ProWorkstation.Text = "Pro for Workstation"
   $groupbox10LTSCeval.Controls.Add($10LTSCEval2ProWorkstation)

   $10LTSCEval2LTSC2019Full = New-Object System.Windows.Forms.RadioButton
   $10LTSCEval2LTSC2019Full.Location = New-Object System.Drawing.Size(10,80)
   $10LTSCEval2LTSC2019Full.Size = New-Object System.Drawing.Size(300,20)
   $10LTSCEval2LTSC2019Full.Checked = $false
   $10LTSCEval2LTSC2019Full.Text = "Enterprise LTSC 2019 (Full)"
   $groupbox10LTSCeval.Controls.Add($10LTSCEval2LTSC2019Full)

   $10LTSCEval2LTSC2021Full = New-Object System.Windows.Forms.RadioButton
   $10LTSCEval2LTSC2021Full.Location = New-Object System.Drawing.Size(10,100)
   $10LTSCEval2LTSC2021Full.Size = New-Object System.Drawing.Size(300,20)
   $10LTSCEval2LTSC2021Full.Checked = $false
   $10LTSCEval2LTSC2021Full.Text = "Enterprise LTSC 2021 (Full)"
   $groupbox10LTSCeval.Controls.Add($10LTSCEval2LTSC2021Full)
############################################## End Windows 10 LTSC Eval checkboxes
############################################## Start buttons

   $submitButton = New-Object System.Windows.Forms.Button 
   $submitButton.Cursor = [System.Windows.Forms.Cursors]::Hand
   $submitButton.BackColor = [System.Drawing.Color]::LightGreen
   $submitButton.Location = New-Object System.Drawing.Size(10,380) 
   $submitButton.Size = New-Object System.Drawing.Size(110,40) 
   $submitButton.Text = "Submit" 
   $submitButton.Add_Click({microsoftInstaller}) 
   $Form.Controls.Add($submitButton) 

############################################## end buttons

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
