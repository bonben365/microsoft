if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    break
}

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework")
[void] [Reflection.Assembly]::LoadWithPartialName("PresentationCore")

$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(785,450)
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

   $uri = 'https://github.com/bonben365/office-installer/raw/main/setup.exe'
   (New-Object Net.WebClient).DownloadFile($uri, "$env:temp\c2r\setup.exe")
   .\setup.exe /configure .\$fileName
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

$uninstall = {
   New-Item -Path $env:temp\c2r -ItemType Directory -Force
   Set-Location $env:temp\c2r
   $fileName = 'configuration.xml'
   New-Item $fileName -ItemType File -Force
   Add-Content $fileName -Value '<Configuration>'
   Add-Content $fileName -Value '<Remove All="True"/>'
   Add-Content $fileName -Value '</Configuration>'
   $uri = 'https://github.com/bonben365/office-installer/raw/main/setup.exe'
   (New-Object Net.WebClient).DownloadFile($uri, "$env:temp\c2r\setup.exe")
   .\setup.exe /configure .\configuration.xml
}

############################################## Start functions

   function microsoftInstaller {
   try {
   if ($arch32.Checked -eq $true) {$arch='32'}
   if ($arch64.Checked -eq $true) {$arch='64'}

   if ($English.Checked -eq $true) {$languageId='en-US'}
   if ($Japanese.Checked -eq $true) {$languageId='ja-JP'}
   if ($Korean.Checked -eq $true) {$languageId='ko-KR'}
   if ($Chinese.Checked -eq $true) {$languageId='zh-TW'}
   if ($French.Checked -eq $true) {$languageId='fr-FR'}
   if ($Spanish.Checked -eq $true) {$languageId='es-ES'}
   if ($Vietnamese.Checked -eq $true) {$languageId='vi-VN'}


   if ($m365Home.Checked -eq $true) {$productId = 'O365HomePremRetail'; Invoke-Command $install}
   if ($m365Business.Checked -eq $true) {$productId = 'O365BusinessRetail'; Invoke-Command $install}
   if ($m365Enterprise.Checked -eq $true) {$productId = 'O365ProPlusRetail'; Invoke-Command $install}

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


   if ($2013Pro.Checked -eq $true) {$productId = 'ProfessionalRetail';Invoke-Command $install2013}
   if ($2013Std.Checked -eq $true) {$productId = 'StandardRetail';Invoke-Command $install2013}
   if ($2013ProjectPro.Checked -eq $true) {$productId = 'VisioProRetail';Invoke-Command $install2013}
   if ($2013VisioPro.Checked -eq $true) {$productId = 'ProjectProRetail';Invoke-Command $install2013}

   if ($uninstallcb.Checked -eq $true) {Invoke-Command $uninstall}


   } #end try

   catch {$outputBox.text = "`nOperation could not be completed"}

} 
############################################## end functions

############################################## Start group boxes

   $arch = New-Object System.Windows.Forms.GroupBox
   $arch.Location = New-Object System.Drawing.Size(10,10) 
   $arch.size = New-Object System.Drawing.Size(130,70) 
   $arch.text = "Arch:"
   $Form.Controls.Add($arch) 

   $language = New-Object System.Windows.Forms.GroupBox
   $language.Location = New-Object System.Drawing.Size(10,90) 
   $language.size = New-Object System.Drawing.Size(130,170) 
   $language.text = "Language:"
   $Form.Controls.Add($language) 

   $groupBox365 = New-Object System.Windows.Forms.GroupBox
   $groupBox365.Location = New-Object System.Drawing.Size(150,10) 
   $groupBox365.size = New-Object System.Drawing.Size(130,90) 
   $groupBox365.text = "Microsoft 365:"
   $Form.Controls.Add($groupBox365) 

   $groupBox2021 = New-Object System.Windows.Forms.GroupBox
   $groupBox2021.Location = New-Object System.Drawing.Size(290,10) 
   $groupBox2021.size = New-Object System.Drawing.Size(130,310) 
   $groupBox2021.text = "Office 2021 Apps:"
   $Form.Controls.Add($groupBox2021)

   $groupBox2019 = New-Object System.Windows.Forms.GroupBox
   $groupBox2019.Location = New-Object System.Drawing.Size(430,10) 
   $groupBox2019.size = New-Object System.Drawing.Size(130,310) 
   $groupBox2019.text = "Office 2019 Apps:"
   $Form.Controls.Add($groupBox2019)

   $groupBox2016 = New-Object System.Windows.Forms.GroupBox
   $groupBox2016.Location = New-Object System.Drawing.Size(570,10) 
   $groupBox2016.size = New-Object System.Drawing.Size(130,130) 
   $groupBox2016.text = "Office 2016 Apps:"
   $Form.Controls.Add($groupBox2016)

   $groupBox2013 = New-Object System.Windows.Forms.GroupBox
   $groupBox2013.Location = New-Object System.Drawing.Size(570,150) 
   $groupBox2013.size = New-Object System.Drawing.Size(130,110) 
   $groupBox2013.text = "Office 2013 Apps:"
   $Form.Controls.Add($groupBox2013)

   $groupBoxUninstall = New-Object System.Windows.Forms.GroupBox
   $groupBoxUninstall.Location = New-Object System.Drawing.Size(290,330) 
   $groupBoxUninstall.size = New-Object System.Drawing.Size(280,50) 
   $groupBoxUninstall.text = "Remove All Office Apps:"
   $Form.Controls.Add($groupBoxUninstall)

############################################## end group boxes

############################################## Start Arch checkboxes

   $arch64 = New-Object System.Windows.Forms.RadioButton
   $arch64.Location = New-Object System.Drawing.Size(10,20)
   $arch64.Size = New-Object System.Drawing.Size(100,20)
   $arch64.Checked = $true
   $arch64.Text = "64 bit"
   $arch.Controls.Add($arch64)

   $arch32 = New-Object System.Windows.Forms.RadioButton
   $arch32.Location = New-Object System.Drawing.Size(10,40)
   $arch32.Size = New-Object System.Drawing.Size(100,20)
   $arch32.Checked = $false
   $arch32.Text = "32 bit"
   $arch.Controls.Add($arch32)

############################################## End Arch checkboxes 

############################################## Start Arch checkboxes

   $English = New-Object System.Windows.Forms.RadioButton
   $English.Location = New-Object System.Drawing.Size(10,20)
   $English.Size = New-Object System.Drawing.Size(100,20)
   $English.Checked = $true
   $English.Text = "English"
   $language.Controls.Add($English)

   $Japanese = New-Object System.Windows.Forms.RadioButton
   $Japanese.Location = New-Object System.Drawing.Size(10,40)
   $Japanese.Size = New-Object System.Drawing.Size(100,20)
   $Japanese.Text = "Japanese"
   $language.Controls.Add($Japanese)

   $Korean = New-Object System.Windows.Forms.RadioButton
   $Korean.Location = New-Object System.Drawing.Size(10,60)
   $Korean.Size = New-Object System.Drawing.Size(100,20)
   $Korean.Text = "Korean"
   $language.Controls.Add($Korean)

   $Chinese = New-Object System.Windows.Forms.RadioButton
   $Chinese.Location = New-Object System.Drawing.Size(10,80)
   $Chinese.Size = New-Object System.Drawing.Size(100,20)
   $Chinese.Text = "Chinese"
   $language.Controls.Add($Chinese)

   $French = New-Object System.Windows.Forms.RadioButton
   $French.Location = New-Object System.Drawing.Size(10,100)
   $French.Size = New-Object System.Drawing.Size(100,20)
   $French.Text = "French"
   $language.Controls.Add($French)

   $Spanish = New-Object System.Windows.Forms.RadioButton
   $Spanish.Location = New-Object System.Drawing.Size(10,120)
   $Spanish.Size = New-Object System.Drawing.Size(100,20)
   $Spanish.Text = "Spanish"
   $language.Controls.Add($Spanish)

   $Vietnamese = New-Object System.Windows.Forms.RadioButton
   $Vietnamese.Location = New-Object System.Drawing.Size(10,140)
   $Vietnamese.Size = New-Object System.Drawing.Size(100,20)
   $Vietnamese.Text = "Vietnamese"
   $language.Controls.Add($Vietnamese)

############################################## End Arch checkboxes 

############################################## Start Microsoft 365 checkboxes
   $m365Home = New-Object System.Windows.Forms.RadioButton
   $m365Home.Location = New-Object System.Drawing.Size(10,20)
   $m365Home.Size = New-Object System.Drawing.Size(100,20)
   $m365Home.Checked = $false
   $m365Home.Text = "Home"
   $groupBox365.Controls.Add($m365Home)

   $m365Business = New-Object System.Windows.Forms.RadioButton
   $m365Business.Location = New-Object System.Drawing.Size(10,40)
   $m365Business.Size = New-Object System.Drawing.Size(100,20)
   $m365Business.Text = "Business"
   $groupBox365.Controls.Add($m365Business)

   $m365Enterprise = New-Object System.Windows.Forms.RadioButton
   $m365Enterprise.Location = New-Object System.Drawing.Size(10,60)
   $m365Enterprise.Size = New-Object System.Drawing.Size(100,20)
   $m365Enterprise.Text = "Enterprise"
   $groupBox365.Controls.Add($m365Enterprise)
############################################## End Microsoft 365 checkboxes

############################################## Start Office 2021 checkboxes
   $2021Pro = New-Object System.Windows.Forms.RadioButton
   $2021Pro.Location = New-Object System.Drawing.Size(10,20)
   $2021Pro.Size = New-Object System.Drawing.Size(100,20)
   $2021Pro.Checked = $false
   $2021Pro.Text = "Professional"
   $groupBox2021.Controls.Add($2021Pro)

   $2021Std = New-Object System.Windows.Forms.RadioButton
   $2021Std.Location = New-Object System.Drawing.Size(10,40)
   $2021Std.Size = New-Object System.Drawing.Size(100,20)
   $2021Std.Text = "Standard"
   $groupBox2021.Controls.Add($2021Std)

   $2021ProjectPro = New-Object System.Windows.Forms.RadioButton
   $2021ProjectPro.Location = New-Object System.Drawing.Size(10,60)
   $2021ProjectPro.Size = New-Object System.Drawing.Size(100,20)
   $2021ProjectPro.Text = "Project Pro"
   $groupBox2021.Controls.Add($2021ProjectPro)

   $2021ProjectStd = New-Object System.Windows.Forms.RadioButton
   $2021ProjectStd.Location = New-Object System.Drawing.Size(10,80)
   $2021ProjectStd.Size = New-Object System.Drawing.Size(100,20)
   $2021ProjectStd.AutoSize = $true
   $2021ProjectStd.Text = "Project Standard"
   $groupBox2021.Controls.Add($2021ProjectStd)

   $2021VisioPro = New-Object System.Windows.Forms.RadioButton
   $2021VisioPro.Location = New-Object System.Drawing.Size(10,100)
   $2021VisioPro.Size = New-Object System.Drawing.Size(100,20)
   $2021VisioPro.Text = "Visio Pro"
   $groupBox2021.Controls.Add($2021VisioPro)

   $2021VisioStd = New-Object System.Windows.Forms.RadioButton
   $2021VisioStd.Location = New-Object System.Drawing.Size(10,120)
   $2021VisioStd.Size = New-Object System.Drawing.Size(100,20)
   $2021VisioStd.Text = "Standard"
   $groupBox2021.Controls.Add($2021VisioStd)

   $2021Word = New-Object System.Windows.Forms.RadioButton
   $2021Word.Location = New-Object System.Drawing.Size(10,140)
   $2021Word.Size = New-Object System.Drawing.Size(100,20)
   $2021Word.Text = "Word"
   $groupBox2021.Controls.Add($2021Word)

   $2021Excel = New-Object System.Windows.Forms.RadioButton
   $2021Excel.Location = New-Object System.Drawing.Size(10,160)
   $2021Excel.Size = New-Object System.Drawing.Size(100,20)
   $2021Excel.Text = "Excel"
   $groupBox2021.Controls.Add($2021Excel)

   $2021PowerPoint = New-Object System.Windows.Forms.RadioButton
   $2021PowerPoint.Location = New-Object System.Drawing.Size(10,180)
   $2021PowerPoint.Size = New-Object System.Drawing.Size(100,20)
   $2021PowerPoint.Text = "PowerPoint"
   $groupBox2021.Controls.Add($2021PowerPoint)

   $2021Outlook = New-Object System.Windows.Forms.RadioButton
   $2021Outlook.Location = New-Object System.Drawing.Size(10,200)
   $2021Outlook.Size = New-Object System.Drawing.Size(100,20)
   $2021Outlook.Text = "Outlook"
   $groupBox2021.Controls.Add($2021Outlook)

   $2021Publisher = New-Object System.Windows.Forms.RadioButton
   $2021Publisher.Location = New-Object System.Drawing.Size(10,220)
   $2021Publisher.Size = New-Object System.Drawing.Size(100,20)
   $2021Publisher.Text = "Publisher"
   $groupBox2021.Controls.Add($2021Publisher)

   $2021Access = New-Object System.Windows.Forms.RadioButton
   $2021Access.Location = New-Object System.Drawing.Size(10,240)
   $2021Access.Size = New-Object System.Drawing.Size(100,20)
   $2021Access.Text = "Access"
   $groupBox2021.Controls.Add($2021Access)

   $2021HomeBusiness = New-Object System.Windows.Forms.RadioButton
   $2021HomeBusiness.Location = New-Object System.Drawing.Size(10,260)
   $2021HomeBusiness.Size = New-Object System.Drawing.Size(100,20)
   $2021HomeBusiness.Text = "HomeBusiness"
   $groupBox2021.Controls.Add($2021HomeBusiness)

   $2021HomeStudent = New-Object System.Windows.Forms.RadioButton
   $2021HomeStudent.Location = New-Object System.Drawing.Size(10,280)
   $2021HomeStudent.Size = New-Object System.Drawing.Size(100,20)
   $2021HomeStudent.Text = "HomeStudent"
   $groupBox2021.Controls.Add($2021HomeStudent)
############################################## End Office 2021 checkboxes

############################################## Start Office 2019 checkboxes
   $2019Pro = New-Object System.Windows.Forms.RadioButton
   $2019Pro.Location = New-Object System.Drawing.Size(10,20)
   $2019Pro.Size = New-Object System.Drawing.Size(100,20)
   $2019Pro.Checked = $false
   $2019Pro.Text = "Professional"
   $groupBox2019.Controls.Add($2019Pro)

   $2019Std = New-Object System.Windows.Forms.RadioButton
   $2019Std.Location = New-Object System.Drawing.Size(10,40)
   $2019Std.Size = New-Object System.Drawing.Size(100,20)
   $2019Std.Text = "Standard"
   $groupBox2019.Controls.Add($2019Std)

   $2019ProjectPro = New-Object System.Windows.Forms.RadioButton
   $2019ProjectPro.Location = New-Object System.Drawing.Size(10,60)
   $2019ProjectPro.Size = New-Object System.Drawing.Size(100,20)
   $2019ProjectPro.Text = "Project Pro"
   $groupBox2019.Controls.Add($2019ProjectPro)

   $2019ProjectStd = New-Object System.Windows.Forms.RadioButton
   $2019ProjectStd.Location = New-Object System.Drawing.Size(10,80)
   $2019ProjectStd.Size = New-Object System.Drawing.Size(100,20)
   $2019ProjectStd.Text = "Project Standard"
   $2019ProjectStd.AutoSize = $true
   $groupBox2019.Controls.Add($2019ProjectStd)

   $2019VisioPro = New-Object System.Windows.Forms.RadioButton
   $2019VisioPro.Location = New-Object System.Drawing.Size(10,100)
   $2019VisioPro.Size = New-Object System.Drawing.Size(100,20)
   $2019VisioPro.Text = "Visio Pro"
   $groupBox2019.Controls.Add($2019VisioPro)

   $2019VisioStd = New-Object System.Windows.Forms.RadioButton
   $2019VisioStd.Location = New-Object System.Drawing.Size(10,120)
   $2019VisioStd.Size = New-Object System.Drawing.Size(100,20)
   $2019VisioStd.Text = "Standard"
   $groupBox2019.Controls.Add($2019VisioStd)

   $2019Word = New-Object System.Windows.Forms.RadioButton
   $2019Word.Location = New-Object System.Drawing.Size(10,140)
   $2019Word.Size = New-Object System.Drawing.Size(100,20)
   $2019Word.Text = "Word"
   $groupBox2019.Controls.Add($2019Word)

   $2019Excel = New-Object System.Windows.Forms.RadioButton
   $2019Excel.Location = New-Object System.Drawing.Size(10,160)
   $2019Excel.Size = New-Object System.Drawing.Size(100,20)
   $2019Excel.Text = "Excel"
   $groupBox2019.Controls.Add($2019Excel)

   $2019PowerPoint = New-Object System.Windows.Forms.RadioButton
   $2019PowerPoint.Location = New-Object System.Drawing.Size(10,180)
   $2019PowerPoint.Size = New-Object System.Drawing.Size(100,20)
   $2019PowerPoint.Text = "PowerPoint"
   $groupBox2019.Controls.Add($2019PowerPoint)

   $2019Outlook = New-Object System.Windows.Forms.RadioButton
   $2019Outlook.Location = New-Object System.Drawing.Size(10,200)
   $2019Outlook.Size = New-Object System.Drawing.Size(100,20)
   $2019Outlook.Text = "Outlook"
   $groupBox2019.Controls.Add($2019Outlook)

   $2019Publisher = New-Object System.Windows.Forms.RadioButton
   $2019Publisher.Location = New-Object System.Drawing.Size(10,220)
   $2019Publisher.Size = New-Object System.Drawing.Size(100,20)
   $2019Publisher.Text = "Publisher"
   $groupBox2019.Controls.Add($2019Publisher)

   $2019Access = New-Object System.Windows.Forms.RadioButton
   $2019Access.Location = New-Object System.Drawing.Size(10,240)
   $2019Access.Size = New-Object System.Drawing.Size(100,20)
   $2019Access.Text = "Access"
   $groupBox2019.Controls.Add($2019Access)

   $2019HomeBusiness = New-Object System.Windows.Forms.RadioButton
   $2019HomeBusiness.Location = New-Object System.Drawing.Size(10,260)
   $2019HomeBusiness.Size = New-Object System.Drawing.Size(100,20)
   $2019HomeBusiness.Text = "HomeBusiness"
   $groupBox2019.Controls.Add($2019HomeBusiness)

   $2019HomeStudent = New-Object System.Windows.Forms.RadioButton
   $2019HomeStudent.Location = New-Object System.Drawing.Size(10,280)
   $2019HomeStudent.Size = New-Object System.Drawing.Size(100,20)
   $2019HomeStudent.Text = "HomeStudent"
   $groupBox2019.Controls.Add($2019HomeStudent)
############################################## End Office 2019 checkboxes


############################################## Start Office 2016 checkboxes
   $2016Pro = New-Object System.Windows.Forms.RadioButton
   $2016Pro.Location = New-Object System.Drawing.Size(10,20)
   $2016Pro.Size = New-Object System.Drawing.Size(100,20)
   $2016Pro.Checked = $false
   $2016Pro.Text = "Professional"
   $groupBox2016.Controls.Add($2016Pro)

   $2016Std = New-Object System.Windows.Forms.RadioButton
   $2016Std.Location = New-Object System.Drawing.Size(10,40)
   $2016Std.Size = New-Object System.Drawing.Size(100,20)
   $2016Std.Text = "Standard"
   $groupBox2016.Controls.Add($2016Std)

   $2016ProjectPro = New-Object System.Windows.Forms.RadioButton
   $2016ProjectPro.Location = New-Object System.Drawing.Size(10,60)
   $2016ProjectPro.Size = New-Object System.Drawing.Size(100,20)
   $2016ProjectPro.Text = "Project Pro"
   $groupBox2016.Controls.Add($2016ProjectPro)

   $2016VisioPro = New-Object System.Windows.Forms.RadioButton
   $2016VisioPro.Location = New-Object System.Drawing.Size(10,80)
   $2016VisioPro.Size = New-Object System.Drawing.Size(100,20)
   $2016VisioPro.Text = "Visio Pro"
   $groupBox2016.Controls.Add($2016VisioPro)

   $2016OneNote = New-Object System.Windows.Forms.RadioButton
   $2016OneNote.Location = New-Object System.Drawing.Size(10,100)
   $2016OneNote.Size = New-Object System.Drawing.Size(100,20)
   $2016OneNote.Text = "OneNote"
   $groupBox2016.Controls.Add($2016OneNote)

############################################## End Office 2016 checkboxes

############################################## Start Office 2013 checkboxes
   $2013Pro = New-Object System.Windows.Forms.RadioButton
   $2013Pro.Location = New-Object System.Drawing.Size(10,20)
   $2013Pro.Size = New-Object System.Drawing.Size(100,20)
   $2013Pro.Checked = $false
   $2013Pro.Text = "Professional"
   $groupBox2013.Controls.Add($2013Pro)

   $2013Std = New-Object System.Windows.Forms.RadioButton
   $2013Std.Location = New-Object System.Drawing.Size(10,40)
   $2013Std.Size = New-Object System.Drawing.Size(100,20)
   $2013Std.Text = "Standard"
   $groupBox2013.Controls.Add($2013Std)

   $2013ProjectPro = New-Object System.Windows.Forms.RadioButton
   $2013ProjectPro.Location = New-Object System.Drawing.Size(10,60)
   $2013ProjectPro.Size = New-Object System.Drawing.Size(100,20)
   $2013ProjectPro.Text = "Project Pro"
   $groupBox2013.Controls.Add($2013ProjectPro)

   $2013VisioPro = New-Object System.Windows.Forms.RadioButton
   $2013VisioPro.Location = New-Object System.Drawing.Size(10,80)
   $2013VisioPro.Size = New-Object System.Drawing.Size(100,20)
   $2013VisioPro.Text = "Visio Pro"
   $groupBox2013.Controls.Add($2013VisioPro)

############################################## End Office 2013 checkboxes


############################################## Start uninstall checkbox
   $uninstallcb = New-Object System.Windows.Forms.RadioButton
   $uninstallcb.Location = New-Object System.Drawing.Size(10,20)
   $uninstallcb.Size = New-Object System.Drawing.Size(200,20)
   $uninstallcb.Text = "I Agree (Be careful)"
   $groupBoxUninstall.Controls.Add($uninstallcb)
############################################## End uninstall checkbox

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
