#===============================================================
#Name:           Windows Converter.
#Description:    Convert all Windows Editions for free.
#Version:        1.0
#Date :          26/7/2023
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
$Form.Size = New-Object System.Drawing.Size(540,450)
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
   New-Item -Path $env:temp\temp -ItemType Directory -Force
   Set-Location $env:temp\temp
    
   $uri = 'https://raw.githubusercontent.com/bonben365/microsoft/main/Files/skus.zip'
   (New-Object Net.WebClient).DownloadFile($uri, "$env:temp\temp\skus.zip")
    
   Expand-Archive .\skus.zip -DestinationPath . -Force | Out-Null
    
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

   $command = "cscript $env:windir\system32\slmgr.vbs /dlv"
   $status = Invoke-Expression -Command $command
   Write-Host "$($status | Select-String -SimpleMatch "License Status")" -ForegroundColor Yellow
   Write-Host

   #Cleanup
   Set-Location $env:temp
   Remove-Item -Path $env:temp\temp -Recurse -Force
}

############################################## Start functions

   function microsoftInstaller {
   try {

   if ($10Home.Checked -eq $true) {$sku = 'Professional'; $key = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'; Invoke-Command $convert}
   if ($10HomeSL.Checked -eq $true) {$productId = 'Standard2021Volume';Invoke-Command $convert}
   if ($10Pro.Checked -eq $true) {$sku = 'Professional'; $key = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'; Invoke-Command $convert}
   if ($10ProWorkstation.Checked -eq $true) {$sku = 'ProfessionalWorkstation'; $key = 'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J'; Invoke-Command $convert}
   if ($10Enterprise.Checked -eq $true) {$sku = 'Enterprise'; $key = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'; Invoke-Command $convert}
   if ($10Education.Checked -eq $true) {$sku = 'Education'; $key = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2'; Invoke-Command $convert}
   if ($10ProEducation.Checked -eq $true) {$sku = 'ProfessionalEducation'; $key = '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y'; Invoke-Command $convert}


   if ($11Home.Checked -eq $true) {$productId = 'ProPlus2019Volume';Invoke-Command $convert}
   if ($11HomeSL.Checked -eq $true) {$productId = 'Standard2019Volume';Invoke-Command $convert}
   if ($11Pro.Checked -eq $true) {$productId = 'ProjectPro2019Volume';Invoke-Command $convert}
   if ($11ProWorskstation.Checked -eq $true) {$productId = 'ProjectStd2019Volume';Invoke-Command $convert}
   if ($11Enterprise.Checked -eq $true) {$productId = 'VisioPro2019Volume';Invoke-Command $convert}
   if ($11Education.Checked -eq $true) {$productId = 'VisioStd2019Volume';Invoke-Command $convert}
   if ($11ProEducation.Checked -eq $true) {$productId = 'Word2019Volume';Invoke-Command $convert}


   if ($2016Pro.Checked -eq $true) {$productId = 'ProfessionalRetail';Invoke-Command $convert}
   if ($2016Std.Checked -eq $true) {$productId = 'StandardRetail';Invoke-Command $convert}
   if ($2016ProjectPro.Checked -eq $true) {$productId = 'VisioProRetail';Invoke-Command $convert}
   if ($2016VisioPro.Checked -eq $true) {$productId = 'ProjectProRetail';Invoke-Command $convert}
   if ($2016OneNote.Checked -eq $true) {$productId = 'OneNoteRetail';Invoke-Command $convert}

   } #end try

   catch {$outputBox.text = "`nOperation could not be completed"}

} 
############################################## end functions

############################################## Start group boxes

   $groupbox10 = New-Object System.Windows.Forms.GroupBox
   $groupbox10.Location = New-Object System.Drawing.Size(10,10) 
   $groupbox10.size = New-Object System.Drawing.Size(160,200) 
   $groupbox10.text = "Windows 10"
   $Form.Controls.Add($groupbox10)

   $groupbox11 = New-Object System.Windows.Forms.GroupBox
   $groupbox11.Location = New-Object System.Drawing.Size(180,10) 
   $groupbox11.size = New-Object System.Drawing.Size(160,200) 
   $groupbox11.text = "Windows 11"
   $Form.Controls.Add($groupbox11)

   $groupboxsrv = New-Object System.Windows.Forms.GroupBox
   $groupboxsrv.Location = New-Object System.Drawing.Size(350,10) 
   $groupboxsrv.size = New-Object System.Drawing.Size(160,200) 
   $groupboxsrv.text = "Windows Server"
   $Form.Controls.Add($groupboxsrv)


############################################## end group boxes


############################################## Start Office 2021 checkboxes
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

############################################## End Office 2021 checkboxes

############################################## Start Office 2019 checkboxes
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


############################################## End Office 2019 checkboxes


############################################## Start Office 2016 checkboxes
   $2016Pro = New-Object System.Windows.Forms.RadioButton
   $2016Pro.Location = New-Object System.Drawing.Size(10,20)
   $2016Pro.Size = New-Object System.Drawing.Size(140,20)
   $2016Pro.Checked = $false
   $2016Pro.Text = "Professional"
   $groupboxsrv.Controls.Add($2016Pro)

   $2016Std = New-Object System.Windows.Forms.RadioButton
   $2016Std.Location = New-Object System.Drawing.Size(10,40)
   $2016Std.Size = New-Object System.Drawing.Size(140,20)
   $2016Std.Text = "Standard"
   $groupboxsrv.Controls.Add($2016Std)

   $2016ProjectPro = New-Object System.Windows.Forms.RadioButton
   $2016ProjectPro.Location = New-Object System.Drawing.Size(10,60)
   $2016ProjectPro.Size = New-Object System.Drawing.Size(140,20)
   $2016ProjectPro.Text = "Project Pro"
   $groupboxsrv.Controls.Add($2016ProjectPro)

   $2016VisioPro = New-Object System.Windows.Forms.RadioButton
   $2016VisioPro.Location = New-Object System.Drawing.Size(10,80)
   $2016VisioPro.Size = New-Object System.Drawing.Size(140,20)
   $2016VisioPro.Text = "Visio Pro"
   $groupboxsrv.Controls.Add($2016VisioPro)

   $2016OneNote = New-Object System.Windows.Forms.RadioButton
   $2016OneNote.Location = New-Object System.Drawing.Size(10,100)
   $2016OneNote.Size = New-Object System.Drawing.Size(140,20)
   $2016OneNote.Text = "OneNote"
   $groupboxsrv.Controls.Add($2016OneNote)

############################################## End Office 2016 checkboxes

############################################## Start buttons

   $submitButton = New-Object System.Windows.Forms.Button 
   $submitButton.Cursor = [System.Windows.Forms.Cursors]::Hand
   $submitButton.BackColor = [System.Drawing.Color]::LightGreen
   $submitButton.Location = New-Object System.Drawing.Size(10,220) 
   $submitButton.Size = New-Object System.Drawing.Size(110,40) 
   $submitButton.Text = "Submit" 
   $submitButton.Add_Click({microsoftInstaller}) 
   $Form.Controls.Add($submitButton) 

############################################## end buttons

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
