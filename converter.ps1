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
$Form.Text = "Microsoft Office Installation Toool - www.bonguides.com" #window description
$Form.ShowInTaskbar = $True
$Form.KeyPreview = $True
$Form.AutoSize = $True
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $False
$Form.MinimizeBox = $False
$Form.ControlBox = $True
$Form.Icon = $Icon

$convert = {
    New-Item -Path $env:temp\temp -ItemType Directory -Force
    Set-Location $env:temp\temp
    
    $uri = 'https://raw.githubusercontent.com/bonben365/microsoft/main/Files/skus.zip'
    (New-Object Net.WebClient).DownloadFile($uri, "$env:temp\temp\skus.zip")
    
    Invoke-Item $env:temp\temp
    Expand-Archive .\skus.zip -DestinationPath . | Out-Null
    
    Copy-Item -Path $sku $env:windir\system32\spp\tokens\skus\ -Recurse -Force -ErrorAction SilentlyContinue
    cscript.exe $env:windir\system32\slmgr.vbs /rilc | Out-Null
    cscript.exe $env:windir\system32\slmgr.vbs /upk | Out-Null
    cscript.exe $env:windir\system32\slmgr.vbs /ckms | Out-Null
    cscript.exe $env:windir\system32\slmgr.vbs /cpky | Out-Null
    cscript.exe $env:windir\system32\slmgr.vbs /skms kms.msgang.com | Out-Null
    cscript.exe $env:windir\system32\slmgr.vbs /ipk $key
    cscript.exe $env:windir\system32\slmgr.vbs /ato
    Write-Host
    Write-Host "Done............"
    Write-Host
    Write-Host "Your Windows edition: $((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName)" -ForegroundColor Yellow

    $command = "cscript $env:windir\system32\slmgr.vbs /dlv"
    $status = Invoke-Expression -Command $command
    Write-Host "$($status | Select-String -SimpleMatch "Product Key Channel")" -ForegroundColor Yellow
    Write-Host "$($status | Select-String -SimpleMatch "License Status")" -ForegroundColor Yellow
    Write-Host "$($status | Select-String -SimpleMatch "Volume activation expiration:")"
    Write-Host
    Write-Host "$($status | Select-String -SimpleMatch "Key Management Service client information")"
    Write-Host "$($status | Select-String -SimpleMatch "Registered KMS machine name:")"
    Write-Host "$($status | Select-String -SimpleMatch "KMS machine IP address:")"
    Write-Host "$($status | Select-String -SimpleMatch "Renewal interval:")"
    Write-Host

}

############################################## Start functions

   function microsoftInstaller {
   try {

   if ($10Home.Checked -eq $true) {$sku = 'Professional'; $key = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'; Invoke-Command $convert}
   if ($10HomeSL.Checked -eq $true) {$productId = 'Standard2021Volume';Invoke-Command $convert}
   if ($10Pro.Checked -eq $true) {$sku = 'Professional'; $key = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'; Invoke-Command $convert}
   if ($10ProWorkstation.Checked -eq $true) {$productId = 'ProjectStd2021Volume'; Invoke-Command $convert}
   if ($10Enterprise.Checked -eq $true) {$productId = 'VisioPro2021Volume'; Invoke-Command $convert}
   if ($10Education.Checked -eq $true) {$productId = 'VisioStd2021Volume'; Invoke-Command $convert}
   if ($10ProEducation.Checked -eq $true) {$productId = 'Word2021Volume';Invoke-Command $convert}


   if ($2019Pro.Checked -eq $true) {$productId = 'ProPlus2019Volume';Invoke-Command $convert}
   if ($2019Std.Checked -eq $true) {$productId = 'Standard2019Volume';Invoke-Command $convert}
   if ($2019ProjectPro.Checked -eq $true) {$productId = 'ProjectPro2019Volume';Invoke-Command $convert}
   if ($2019ProjectStd.Checked -eq $true) {$productId = 'ProjectStd2019Volume';Invoke-Command $convert}
   if ($2019VisioPro.Checked -eq $true) {$productId = 'VisioPro2019Volume';Invoke-Command $convert}
   if ($2019VisioStd.Checked -eq $true) {$productId = 'VisioStd2019Volume';Invoke-Command $convert}
   if ($2019Word.Checked -eq $true) {$productId = 'Word2019Volume';Invoke-Command $convert}


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
   $2019Pro = New-Object System.Windows.Forms.RadioButton
   $2019Pro.Location = New-Object System.Drawing.Size(10,20)
   $2019Pro.Size = New-Object System.Drawing.Size(140,20)
   $2019Pro.Checked = $false
   $2019Pro.Text = "Professional"
   $groupbox11.Controls.Add($2019Pro)

   $2019Std = New-Object System.Windows.Forms.RadioButton
   $2019Std.Location = New-Object System.Drawing.Size(10,40)
   $2019Std.Size = New-Object System.Drawing.Size(140,20)
   $2019Std.Text = "Standard"
   $groupbox11.Controls.Add($2019Std)

   $2019ProjectPro = New-Object System.Windows.Forms.RadioButton
   $2019ProjectPro.Location = New-Object System.Drawing.Size(10,60)
   $2019ProjectPro.Size = New-Object System.Drawing.Size(140,20)
   $2019ProjectPro.Text = "Project Pro"
   $groupbox11.Controls.Add($2019ProjectPro)

   $2019ProjectStd = New-Object System.Windows.Forms.RadioButton
   $2019ProjectStd.Location = New-Object System.Drawing.Size(10,80)
   $2019ProjectStd.Size = New-Object System.Drawing.Size(140,20)
   $2019ProjectStd.Text = "Project Standard"
   $2019ProjectStd.AutoSize = $true
   $groupbox11.Controls.Add($2019ProjectStd)

   $2019VisioPro = New-Object System.Windows.Forms.RadioButton
   $2019VisioPro.Location = New-Object System.Drawing.Size(10,100)
   $2019VisioPro.Size = New-Object System.Drawing.Size(140,20)
   $2019VisioPro.Text = "Visio Pro"
   $groupbox11.Controls.Add($2019VisioPro)

   $2019VisioStd = New-Object System.Windows.Forms.RadioButton
   $2019VisioStd.Location = New-Object System.Drawing.Size(10,120)
   $2019VisioStd.Size = New-Object System.Drawing.Size(140,20)
   $2019VisioStd.Text = "Standard"
   $groupbox11.Controls.Add($2019VisioStd)

   $2019Word = New-Object System.Windows.Forms.RadioButton
   $2019Word.Location = New-Object System.Drawing.Size(10,140)
   $2019Word.Size = New-Object System.Drawing.Size(140,20)
   $2019Word.Text = "Word"
   $groupbox11.Controls.Add($2019Word)

   $2019Excel = New-Object System.Windows.Forms.RadioButton
   $2019Excel.Location = New-Object System.Drawing.Size(10,160)
   $2019Excel.Size = New-Object System.Drawing.Size(140,20)
   $2019Excel.Text = "Excel"
   $groupbox11.Controls.Add($2019Excel)


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
