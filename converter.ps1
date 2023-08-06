if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework")
[void] [Reflection.Assembly]::LoadWithPartialName("PresentationCore")

$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(530,450)
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

$install = { 
    $global:ProgressPreference = "SilentlyContinue"$status
    New-Item -Path $env:temp\temp -ItemType Directory -Force
    Set-Location $env:temp\temp
    
    $uri = 'https://raw.githubusercontent.com/bonben365/microsoft/main/Files/skus.zip'
    (New-Object Net.WebClient).DownloadFile($uri, "$env:temp\temp\skus.zip")
    
    Invoke-Item $env:temp\temp
    Expand-Archive .\skus.zip -DestinationPath . | Out-Null
    
    Copy-Item .\Professional\ $env:windir\system32\spp\tokens\skus\ -Recurse -Force
    cscript.exe $env:windir\system32\slmgr.vbs /rilc
    cscript.exe $env:windir\system32\slmgr.vbs /upk
    cscript.exe $env:windir\system32\slmgr.vbs /ckms
    cscript.exe $env:windir\system32\slmgr.vbs /cpky
    cscript.exe $env:windir\system32\slmgr.vbs /skms kms.msgang.com
    cscript.exe $env:windir\system32\slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
    cscript.exe $env:windir\system32\slmgr.vbs /ato
    Write-Host
    Write-Host "Done............"
    Write-Host
    Write-Host "Your Windows edition: $edition" -ForegroundColor Yellow
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

$install2013 = { 
   New-Item -Path $env:temp\c2r -ItemType Directory -Force
   Set-Location $env:temp\c2r
   $fileName = "configuration-x$arch.xml"
   New-Item $fileName -ItemType File -Force | Out-Null
   Add-Content $fileName -Value '<Configuration>'
   Add-content $fileName -Value "<Add OfficeClientEdition=`"$arch`">"
   Add-content $fileName -Value "<Product ID=`"$productId`">"
   Add-content $fileName -Value "<Language ID=`"$languageId`"/>"
   Add-Content $fileName -Value '</Product>'
   Add-Content $fileName -Value '</Add>'
   Add-Content $fileName -Value '</Configuration>'

   $uri = 'https://github.com/bonben365/office-installer/raw/main/bin2013.exe'
   (New-Object Net.WebClient).DownloadFile($uri, "$env:temp\c2r\bin2013.exe")
   .\bin2013.exe /configure .\$fileName
}


############################################## Start functions

   function microsoftInstaller {
   try {

   if ($2021Pro.Checked -eq $true) {$productId = 'ProPlus2021Volume'; Invoke-Command $install}
   if ($2021Std.Checked -eq $true) {$productId = 'Standard2021Volume';Invoke-Command $install}
   if ($2021ProjectPro.Checked -eq $true) {$productId = 'ProjectPro2021Volume'; Invoke-Command $install}
   if ($2021ProjectStd.Checked -eq $true) {$productId = 'ProjectStd2021Volume'; Invoke-Command $install}
   if ($2021VisioPro.Checked -eq $true) {$productId = 'VisioPro2021Volume'; Invoke-Command $install}
   if ($2021VisioStd.Checked -eq $true) {$productId = 'VisioStd2021Volume'; Invoke-Command $install}
   if ($2021Word.Checked -eq $true) {$productId = 'Word2021Volume';Invoke-Command $install}
   if ($2021Excel.Checked -eq $true) {$productId = 'Excel2021Volume';Invoke-Command $install}
   if ($2021PowerPoint.Checked -eq $true) {$productId = 'PowerPoint2021Volume'; Invoke-Command $install}
   if ($2021Outlook.Checked -eq $true) {$productId = 'Outlook2021Volume'; Invoke-Command $install}
   if ($2021Publisher.Checked -eq $true) {$productId = 'Publisher2021Volume';Invoke-Command $install}
   if ($2021Access.Checked -eq $true) {$productId = 'Access2021Volume'; Invoke-Command $install}
   if ($2021HomeBusiness.Checked -eq $true) {$productId = 'HomeBusiness2021Retail';Invoke-Command $install}
   if ($2021HomeStudent.Checked -eq $true) {$productId = 'HomeStudent2021Retail'; Invoke-Command $install}

   if ($2019Pro.Checked -eq $true) {$productId = 'ProPlus2019Volume';Invoke-Command $install}
   if ($2019Std.Checked -eq $true) {$productId = 'Standard2019Volume';Invoke-Command $install}
   if ($2019ProjectPro.Checked -eq $true) {$productId = 'ProjectPro2019Volume';Invoke-Command $install}
   if ($2019ProjectStd.Checked -eq $true) {$productId = 'ProjectStd2019Volume';Invoke-Command $install}
   if ($2019VisioPro.Checked -eq $true) {$productId = 'VisioPro2019Volume';Invoke-Command $install}
   if ($2019VisioStd.Checked -eq $true) {$productId = 'VisioStd2019Volume';Invoke-Command $install}
   if ($2019Word.Checked -eq $true) {$productId = 'Word2019Volume';Invoke-Command $install}
   if ($2019Excel.Checked -eq $true) {$productId = 'Excel2019Volume';Invoke-Command $install}
   if ($2019PowerPoint.Checked -eq $true) {$productId = 'PowerPoint2019Volume';Invoke-Command $install}
   if ($2019Outlook.Checked -eq $true) {$productId = 'Outlook2019Volume';Invoke-Command $install}
   if ($2019Publisher.Checked -eq $true) {$productId = 'Publisher2019Volume';Invoke-Command $install}
   if ($2019Access.Checked -eq $true) {$productId = 'Access2019Volume';Invoke-Command $install}
   if ($2019HomeBusiness.Checked -eq $true) {$productId = 'HomeBusiness2019Retail';Invoke-Command $install}
   if ($2019HomeStudent.Checked -eq $true) {$productId = 'HomeStudent2019Retail'; Invoke-Command $install}

   if ($2016Pro.Checked -eq $true) {$productId = 'ProfessionalRetail';Invoke-Command $install}
   if ($2016Std.Checked -eq $true) {$productId = 'StandardRetail';Invoke-Command $install}
   if ($2016ProjectPro.Checked -eq $true) {$productId = 'VisioProRetail';Invoke-Command $install}
   if ($2016VisioPro.Checked -eq $true) {$productId = 'ProjectProRetail';Invoke-Command $install}
   if ($2016OneNote.Checked -eq $true) {$productId = 'OneNoteRetail';Invoke-Command $install}

   } #end try

   catch {$outputBox.text = "`nOperation could not be completed"}

} 
############################################## end functions

############################################## Start group boxes

   $groupbox10 = New-Object System.Windows.Forms.GroupBox
   $groupbox10.Location = New-Object System.Drawing.Size(10,10) 
   $groupbox10.size = New-Object System.Drawing.Size(160,310) 
   $groupbox10.text = "Windows 10"
   $Form.Controls.Add($groupbox10)

   $groupbox11 = New-Object System.Windows.Forms.GroupBox
   $groupbox11.Location = New-Object System.Drawing.Size(180,10) 
   $groupbox11.size = New-Object System.Drawing.Size(160,310) 
   $groupbox11.text = "Windows 11"
   $Form.Controls.Add($groupbox11)

   $groupboxsrv = New-Object System.Windows.Forms.GroupBox
   $groupboxsrv.Location = New-Object System.Drawing.Size(360,10) 
   $groupboxsrv.size = New-Object System.Drawing.Size(130,130) 
   $groupboxsrv.text = "Windows Server"
   $Form.Controls.Add($groupboxsrv)


############################################## end group boxes


############################################## Start Office 2021 checkboxes
   $2021Pro = New-Object System.Windows.Forms.RadioButton
   $2021Pro.Location = New-Object System.Drawing.Size(10,20)
   $2021Pro.Size = New-Object System.Drawing.Size(100,20)
   $2021Pro.Checked = $false
   $2021Pro.Text = "Professional"
   $groupbox10.Controls.Add($2021Pro)

   $2021Std = New-Object System.Windows.Forms.RadioButton
   $2021Std.Location = New-Object System.Drawing.Size(10,40)
   $2021Std.Size = New-Object System.Drawing.Size(100,20)
   $2021Std.Text = "Standard"
   $groupbox10.Controls.Add($2021Std)

   $2021ProjectPro = New-Object System.Windows.Forms.RadioButton
   $2021ProjectPro.Location = New-Object System.Drawing.Size(10,60)
   $2021ProjectPro.Size = New-Object System.Drawing.Size(100,20)
   $2021ProjectPro.Text = "Project Pro"
   $groupbox10.Controls.Add($2021ProjectPro)

   $2021ProjectStd = New-Object System.Windows.Forms.RadioButton
   $2021ProjectStd.Location = New-Object System.Drawing.Size(10,80)
   $2021ProjectStd.Size = New-Object System.Drawing.Size(100,20)
   $2021ProjectStd.AutoSize = $true
   $2021ProjectStd.Text = "Project Standard"
   $groupbox10.Controls.Add($2021ProjectStd)

   $2021VisioPro = New-Object System.Windows.Forms.RadioButton
   $2021VisioPro.Location = New-Object System.Drawing.Size(10,100)
   $2021VisioPro.Size = New-Object System.Drawing.Size(100,20)
   $2021VisioPro.Text = "Visio Pro"
   $groupbox10.Controls.Add($2021VisioPro)

   $2021VisioStd = New-Object System.Windows.Forms.RadioButton
   $2021VisioStd.Location = New-Object System.Drawing.Size(10,120)
   $2021VisioStd.Size = New-Object System.Drawing.Size(100,20)
   $2021VisioStd.Text = "Standard"
   $groupbox10.Controls.Add($2021VisioStd)

   $2021Word = New-Object System.Windows.Forms.RadioButton
   $2021Word.Location = New-Object System.Drawing.Size(10,140)
   $2021Word.Size = New-Object System.Drawing.Size(100,20)
   $2021Word.Text = "Word"
   $groupbox10.Controls.Add($2021Word)

   $2021Excel = New-Object System.Windows.Forms.RadioButton
   $2021Excel.Location = New-Object System.Drawing.Size(10,160)
   $2021Excel.Size = New-Object System.Drawing.Size(100,20)
   $2021Excel.Text = "Excel"
   $groupbox10.Controls.Add($2021Excel)

   $2021PowerPoint = New-Object System.Windows.Forms.RadioButton
   $2021PowerPoint.Location = New-Object System.Drawing.Size(10,180)
   $2021PowerPoint.Size = New-Object System.Drawing.Size(100,20)
   $2021PowerPoint.Text = "PowerPoint"
   $groupbox10.Controls.Add($2021PowerPoint)

   $2021Outlook = New-Object System.Windows.Forms.RadioButton
   $2021Outlook.Location = New-Object System.Drawing.Size(10,200)
   $2021Outlook.Size = New-Object System.Drawing.Size(100,20)
   $2021Outlook.Text = "Outlook"
   $groupbox10.Controls.Add($2021Outlook)

   $2021Publisher = New-Object System.Windows.Forms.RadioButton
   $2021Publisher.Location = New-Object System.Drawing.Size(10,220)
   $2021Publisher.Size = New-Object System.Drawing.Size(100,20)
   $2021Publisher.Text = "Publisher"
   $groupbox10.Controls.Add($2021Publisher)

   $2021Access = New-Object System.Windows.Forms.RadioButton
   $2021Access.Location = New-Object System.Drawing.Size(10,240)
   $2021Access.Size = New-Object System.Drawing.Size(100,20)
   $2021Access.Text = "Access"
   $groupbox10.Controls.Add($2021Access)

   $2021HomeBusiness = New-Object System.Windows.Forms.RadioButton
   $2021HomeBusiness.Location = New-Object System.Drawing.Size(10,260)
   $2021HomeBusiness.Size = New-Object System.Drawing.Size(100,20)
   $2021HomeBusiness.Text = "HomeBusiness"
   $groupbox10.Controls.Add($2021HomeBusiness)

   $2021HomeStudent = New-Object System.Windows.Forms.RadioButton
   $2021HomeStudent.Location = New-Object System.Drawing.Size(10,280)
   $2021HomeStudent.Size = New-Object System.Drawing.Size(100,20)
   $2021HomeStudent.Text = "HomeStudent"
   $groupbox10.Controls.Add($2021HomeStudent)
############################################## End Office 2021 checkboxes

############################################## Start Office 2019 checkboxes
   $2019Pro = New-Object System.Windows.Forms.RadioButton
   $2019Pro.Location = New-Object System.Drawing.Size(10,20)
   $2019Pro.Size = New-Object System.Drawing.Size(100,20)
   $2019Pro.Checked = $false
   $2019Pro.Text = "Professional"
   $groupbox11.Controls.Add($2019Pro)

   $2019Std = New-Object System.Windows.Forms.RadioButton
   $2019Std.Location = New-Object System.Drawing.Size(10,40)
   $2019Std.Size = New-Object System.Drawing.Size(100,20)
   $2019Std.Text = "Standard"
   $groupbox11.Controls.Add($2019Std)

   $2019ProjectPro = New-Object System.Windows.Forms.RadioButton
   $2019ProjectPro.Location = New-Object System.Drawing.Size(10,60)
   $2019ProjectPro.Size = New-Object System.Drawing.Size(100,20)
   $2019ProjectPro.Text = "Project Pro"
   $groupbox11.Controls.Add($2019ProjectPro)

   $2019ProjectStd = New-Object System.Windows.Forms.RadioButton
   $2019ProjectStd.Location = New-Object System.Drawing.Size(10,80)
   $2019ProjectStd.Size = New-Object System.Drawing.Size(100,20)
   $2019ProjectStd.Text = "Project Standard"
   $2019ProjectStd.AutoSize = $true
   $groupbox11.Controls.Add($2019ProjectStd)

   $2019VisioPro = New-Object System.Windows.Forms.RadioButton
   $2019VisioPro.Location = New-Object System.Drawing.Size(10,100)
   $2019VisioPro.Size = New-Object System.Drawing.Size(100,20)
   $2019VisioPro.Text = "Visio Pro"
   $groupbox11.Controls.Add($2019VisioPro)

   $2019VisioStd = New-Object System.Windows.Forms.RadioButton
   $2019VisioStd.Location = New-Object System.Drawing.Size(10,120)
   $2019VisioStd.Size = New-Object System.Drawing.Size(100,20)
   $2019VisioStd.Text = "Standard"
   $groupbox11.Controls.Add($2019VisioStd)

   $2019Word = New-Object System.Windows.Forms.RadioButton
   $2019Word.Location = New-Object System.Drawing.Size(10,140)
   $2019Word.Size = New-Object System.Drawing.Size(100,20)
   $2019Word.Text = "Word"
   $groupbox11.Controls.Add($2019Word)

   $2019Excel = New-Object System.Windows.Forms.RadioButton
   $2019Excel.Location = New-Object System.Drawing.Size(10,160)
   $2019Excel.Size = New-Object System.Drawing.Size(100,20)
   $2019Excel.Text = "Excel"
   $groupbox11.Controls.Add($2019Excel)

   $2019PowerPoint = New-Object System.Windows.Forms.RadioButton
   $2019PowerPoint.Location = New-Object System.Drawing.Size(10,180)
   $2019PowerPoint.Size = New-Object System.Drawing.Size(100,20)
   $2019PowerPoint.Text = "PowerPoint"
   $groupbox11.Controls.Add($2019PowerPoint)

   $2019Outlook = New-Object System.Windows.Forms.RadioButton
   $2019Outlook.Location = New-Object System.Drawing.Size(10,200)
   $2019Outlook.Size = New-Object System.Drawing.Size(100,20)
   $2019Outlook.Text = "Outlook"
   $groupbox11.Controls.Add($2019Outlook)

   $2019Publisher = New-Object System.Windows.Forms.RadioButton
   $2019Publisher.Location = New-Object System.Drawing.Size(10,220)
   $2019Publisher.Size = New-Object System.Drawing.Size(100,20)
   $2019Publisher.Text = "Publisher"
   $groupbox11.Controls.Add($2019Publisher)

   $2019Access = New-Object System.Windows.Forms.RadioButton
   $2019Access.Location = New-Object System.Drawing.Size(10,240)
   $2019Access.Size = New-Object System.Drawing.Size(100,20)
   $2019Access.Text = "Access"
   $groupbox11.Controls.Add($2019Access)

   $2019HomeBusiness = New-Object System.Windows.Forms.RadioButton
   $2019HomeBusiness.Location = New-Object System.Drawing.Size(10,260)
   $2019HomeBusiness.Size = New-Object System.Drawing.Size(100,20)
   $2019HomeBusiness.Text = "HomeBusiness"
   $groupbox11.Controls.Add($2019HomeBusiness)

   $2019HomeStudent = New-Object System.Windows.Forms.RadioButton
   $2019HomeStudent.Location = New-Object System.Drawing.Size(10,280)
   $2019HomeStudent.Size = New-Object System.Drawing.Size(100,20)
   $2019HomeStudent.Text = "HomeStudent"
   $groupbox11.Controls.Add($2019HomeStudent)
############################################## End Office 2019 checkboxes


############################################## Start Office 2016 checkboxes
   $2016Pro = New-Object System.Windows.Forms.RadioButton
   $2016Pro.Location = New-Object System.Drawing.Size(10,20)
   $2016Pro.Size = New-Object System.Drawing.Size(100,20)
   $2016Pro.Checked = $false
   $2016Pro.Text = "Professional"
   $groupboxsrv.Controls.Add($2016Pro)

   $2016Std = New-Object System.Windows.Forms.RadioButton
   $2016Std.Location = New-Object System.Drawing.Size(10,40)
   $2016Std.Size = New-Object System.Drawing.Size(100,20)
   $2016Std.Text = "Standard"
   $groupboxsrv.Controls.Add($2016Std)

   $2016ProjectPro = New-Object System.Windows.Forms.RadioButton
   $2016ProjectPro.Location = New-Object System.Drawing.Size(10,60)
   $2016ProjectPro.Size = New-Object System.Drawing.Size(100,20)
   $2016ProjectPro.Text = "Project Pro"
   $groupboxsrv.Controls.Add($2016ProjectPro)

   $2016VisioPro = New-Object System.Windows.Forms.RadioButton
   $2016VisioPro.Location = New-Object System.Drawing.Size(10,80)
   $2016VisioPro.Size = New-Object System.Drawing.Size(100,20)
   $2016VisioPro.Text = "Visio Pro"
   $groupboxsrv.Controls.Add($2016VisioPro)

   $2016OneNote = New-Object System.Windows.Forms.RadioButton
   $2016OneNote.Location = New-Object System.Drawing.Size(10,100)
   $2016OneNote.Size = New-Object System.Drawing.Size(100,20)
   $2016OneNote.Text = "OneNote"
   $groupboxsrv.Controls.Add($2016OneNote)

############################################## End Office 2016 checkboxes

############################################## Start buttons

   $submitButton = New-Object System.Windows.Forms.Button 
   $submitButton.Cursor = [System.Windows.Forms.Cursors]::Hand
   $submitButton.BackColor = [System.Drawing.Color]::LightGreen
   $submitButton.Location = New-Object System.Drawing.Size(10,280) 
   $submitButton.Size = New-Object System.Drawing.Size(110,40) 
   $submitButton.Text = "Submit" 
   $submitButton.Add_Click({microsoftInstaller}) 
   $Form.Controls.Add($submitButton) 

############################################## end buttons

$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
