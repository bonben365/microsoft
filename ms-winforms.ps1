<#=============================================================================================
Script by    : Leo Nguyen
Website      : www.msgang.com
Description  : Activate Microsoft products for FREE without cracking tools

Script Highlights:
~~~~~~~~~~~~~~~~~
#. Activate Microsoft products for FREE without cracking tools.
#. Download/ Install Microsoft Office (C2R) all versions.
============================================================================================#>

if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You need to have Administrator rights to run this script!`nPlease re-run this script as an Administrator in an elevated powershell prompt!"
    Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "irm msgang.com/ms | iex"
    break
}

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Long running task with runspace
function MicrosoftWindowsAct {
    # The main script block
    $scriptBlock = {
        # Disable the button 
        $sync.button.Enabled = $false
        $sync.button1.Enabled = $false
        $sync.textbox.Text = ""
        $sync.textbox.AppendText([Environment]::NewLine)

        $edition = (Get-CimInstance Win32_OperatingSystem).Caption
        $version = (Get-CimInstance Win32_OperatingSystem).Version

        $sync.textbox.AppendText("System information:")
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("---------------------------------------------------------------------")
        $sync.textbox.AppendText([Environment]::NewLine)

        $sync.textbox.AppendText("Edition  : " + $($edition))
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("OS build : " + $($version))
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("---------------------------------------------------------------------")
        $sync.textbox.AppendText([Environment]::NewLine)

        $sync.textbox.AppendText("Activating $($edition)...")
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("Please wait...")

        Invoke-RestMethod msgang.com/win | Invoke-Expression

        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("---------------------------------------------------------------------")
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText([Environment]::NewLine)

        $sync.textbox.AppendText("License status:")
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("---------------------------------------------------------------------")

        $command = "&$env:windir\system32\cscript.exe $env:windir\system32\slmgr.vbs /dlv"
        $status = Invoke-Expression -Command $command
        $productKey = $status | Select-String -SimpleMatch "Product Key Channel"
        $licenseStatus = $status | Select-String -SimpleMatch "License Status"
        # $expiration = $status | Select-String -SimpleMatch "Volume activation expiration"
        # $kmsServer = $status | Select-String -SimpleMatch "Registered KMS machine name"


        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("$($productKey)")

        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("$($licenseStatus)")

        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("---------------------------------------------------------------------")
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("(*) More Microsoft products for free at https://msgang.com")
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("---------------------------------------------------------------------")

        # Re-enable the button when the loop finishes.
        $sync.button.Enabled = $true
        $sync.button1.Enabled = $true
    }

    # Create a PowerShell instance with the script block
    $psInstance = [PowerShell]::Create().AddScript($scriptBlock)
    
    # Create then open a new runspace
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()

    # Add shared data to the runspace
    $runspace.SessionStateProxy.SetVariable("sync", $sync)
    
    # Associate the PowerShell instance with the opened runspace
    $psInstance.Runspace = $runspace

    # Execute the script asynchronously using BeginInvoke() method
    $psInstance.BeginInvoke()
}

function MicrosoftOfficeAct {
    # The main script block
    $scriptBlock = {
        # Disable the button 
        $sync.button.Enabled = $false
        $sync.button1.Enabled = $false
        $sync.textbox.Text = ""
        $sync.textbox.AppendText([Environment]::NewLine)


        # Detect Office installations
        $path64 = "C:\Program Files\Microsoft Office\Office1*"
        $path32 = "C:\Program Files (x86)\Microsoft Office\Office1*"
        if ((Test-Path -Path "$path32\ospp.vbs")) { Set-Location $path32 -ErrorAction SilentlyContinue }
        if ((Test-Path -Path "$path64\ospp.vbs")) { Set-Location $path64 -ErrorAction SilentlyContinue }

        $dstatus = Invoke-Expression -Command "cscript.exe ospp.vbs /dstatus"

        $retailCount = ($dstatus | Select-String -SimpleMatch "RETAIL" | Measure-Object).Count
        $volumeCount = ($dstatus | Select-String -SimpleMatch "VOLUME" | Measure-Object).Count

        $sync.textbox.AppendText("Checking Installed Office editions...")
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("---------------------------------------------------------------------")
        $sync.textbox.AppendText([Environment]::NewLine)

        $sync.textbox.AppendText("Number of Installed Office apps: VOLUME: $($volumeCount) - RETAIL: $($retailCount)")
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("Activating Microsoft Office products...")
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("Please wait....")
        $sync.textbox.AppendText([Environment]::NewLine)

        Invoke-RestMethod 'https://msgang.com/office' | Invoke-Expression

        $sync.textbox.AppendText("---------------------------------------------------------------------")

        $dstatus = Invoke-Expression -Command "cscript.exe ospp.vbs /dstatus"
        $activated = ($dstatus | Select-String -SimpleMatch "LICENSED").Count

        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("Number of Activated Products: $($activated)")
        $sync.textbox.AppendText([Environment]::NewLine)

        $sync.textbox.AppendText("---------------------------------------------------------------------")
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("(*) More Microsoft products for free at https://msgang.com")
        $sync.textbox.AppendText([Environment]::NewLine)
        $sync.textbox.AppendText("---------------------------------------------------------------------")

        # Re-enable the button when the loop finishes.
        $sync.button.Enabled = $true
        $sync.button1.Enabled = $true
    }

    # Create a PowerShell instance with the script block
    $psInstance = [PowerShell]::Create().AddScript($scriptBlock)
    
    # Create then open a new runspace
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()

    # Add shared data to the runspace
    $runspace.SessionStateProxy.SetVariable("sync", $sync)
    
    # Associate the PowerShell instance with the opened runspace
    $psInstance.Runspace = $runspace

    # Execute the script asynchronously using BeginInvoke() method
    $psInstance.BeginInvoke()
}


# Create the form.
    $form                   = New-Object Windows.Forms.Form
    $form.ClientSize        = New-Object Drawing.Size(520, 380)
    $form.Text              = "Free Microsoft Products Activation - www.msgang.com"
    $form.StartPosition     = "CenterScreen"
    $form.FormBorderStyle   = "FixedSingle"

# Create the button.
    $button             = New-Object Windows.Forms.Button
    $button.Location    = New-Object Drawing.Point(10, 15)
    $button.BackColor   = [System.Drawing.Color]::Green
    $button.ForeColor   = [System.Drawing.Color]::White
    $button.Size        = New-Object System.Drawing.Size(150,40) 
    $button.Text        = "Activate Windows"
    $button.Font        = New-Object System.Drawing.Font("Consolas",10,[System.Drawing.FontStyle]::Bold)
    $button.Add_Click({MicrosoftWindowsAct})

# Create the button1.
    $button1             = New-Object Windows.Forms.Button
    $button1.Location    = New-Object Drawing.Point(175, 15)
    $button1.BackColor   = [System.Drawing.Color]::Blue
    $button1.ForeColor   = [System.Drawing.Color]::White
    $button1.Size        = New-Object System.Drawing.Size(150,40) 
    $button1.Text        = "Activate Office"
    $button1.Font        = New-Object System.Drawing.Font("Consolas",10,[System.Drawing.FontStyle]::Bold)
    $button1.Add_Click({MicrosoftOfficeAct})    

# Create the button2.
    $button2             = New-Object Windows.Forms.Button
    $button2.Location    = New-Object Drawing.Point(340, 15)
    $button2.BackColor   = [System.Drawing.Color]::Red
    $button2.ForeColor   = [System.Drawing.Color]::White
    $button2.Size        = New-Object System.Drawing.Size(160,40) 
    $button2.Text        = "Install Office"
    $button2.Font        = New-Object System.Drawing.Font("Consolas",10,[System.Drawing.FontStyle]::Bold)
    $button2.Add_Click({Invoke-RestMethod msgang.com/install| Invoke-Expression})   
# Create a textbox to display the output
    $textbox                      = New-Object system.Windows.Forms.TextBox
    $textbox.Multiline            = $true
    $textbox.Text                 = ""
    $textbox.AppendText([Environment]::NewLine)
    $textbox.AppendText("(*) Activate Microsoft products for FREE without cracking tools:")
    $textbox.AppendText([Environment]::NewLine)
    $textbox.AppendText("(*)  - Windows 7/8/8.1/10/11.")
    $textbox.AppendText([Environment]::NewLine)
    $textbox.AppendText("(*)  - Windows Server 2012/2012 R2/2016/2019/2022.")
    $textbox.AppendText([Environment]::NewLine)
    $textbox.AppendText("(*)  - Microsoft Office 2013/2016/2019/2021/365.")
    $textbox.AppendText([Environment]::NewLine)
    $textbox.AppendText("(*)  - Microsoft Visio / Project all versions.")
    $textbox.AppendText([Environment]::NewLine)
    $textbox.AppendText("(*)  - Download / Install all Microsoft Office versions.")
    $textbox.AppendText([Environment]::NewLine)
    $textbox.AppendText([Environment]::NewLine)
    $textbox.AppendText("---------------------------------------------------------------------")
    $textbox.AppendText([Environment]::NewLine)
    $textbox.AppendText("> Script by : Leo Nguyen")
    $textbox.AppendText([Environment]::NewLine)
    $textbox.AppendText("> Website   : www.msgang.com")
    $textbox.AppendText([Environment]::NewLine)
    $textbox.AppendText("> Copyright 2023 All rights Reserved. Design by Leo with <3<3<3")

    $textbox.Font                 = New-Object System.Drawing.Font("Consolas",10,[System.Drawing.FontStyle]::Bold)
    $textbox.Size                 = New-Object System.Drawing.Size(500,300)
    $textbox.Location             = New-Object System.Drawing.Point(10,70)
    $textbox.BackColor            = "#0D1117"
    $textbox.ForeColor            = 'White'

# For talking across runspaces (share info between runspaces)
    $sync           = [hashtable]::Synchronized(@{})
    $sync.runspace  = $runspace
    $sync.host      = $host
    $sync.form      = $form
    $sync.button    = $button
    $sync.button1   = $button1
    $sync.button2   = $button2
    $sync.label     = $label
    $sync.textbox   = $textbox

# Add controls to the form.
    $form.Controls.AddRange(@($sync.button, $sync.button1,$sync.button2, $sync.label, $sync.textbox))

# Show the form.
    [Windows.Forms.Application]::Run($form)
