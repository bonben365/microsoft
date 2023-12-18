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
    Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "irm msgang.com/install | iex"
    break
}

# Load ddls to the current session.
Add-Type -AssemblyName PresentationFramework, System.Drawing, PresentationFramework, System.Windows.Forms, WindowsFormsIntegration, PresentationCore
[System.Windows.Forms.Application]::EnableVisualStyles()

# Place your xaml code from Visual Studio in here string (between @ symbols)
# $xamlinput = @'<xaml code here'@

$xamlInput = @'
<Window x:Class="ms.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:ms"
        mc:Ignorable="d"
        Title="Free Microsoft Products Activation - www.msgang.com" Height="441" Width="529" Background="#FFFDF2F2">
    <Grid Margin="0,0,25,20">
        <Button x:Name="button" Content="Active Windows " HorizontalAlignment="Left" Margin="24,31,0,0" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold" Background="#FF19B74E" Height="36" Foreground="White" Width="141"/>
        <TextBox x:Name="textbox" HorizontalAlignment="Left" Margin="25,91,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="463" Height="287" Background="#FF0D1117" FontFamily="Consolas" Foreground="White"/>
        <Button x:Name="button1" Content="Active Office" HorizontalAlignment="Left" Margin="181,31,0,0" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold" Background="#FF1923B7" Height="36" Foreground="White" Width="151"/>
        <Button x:Name="button2" Content="Install Office" HorizontalAlignment="Left" Margin="345,31,0,0" VerticalAlignment="Top" FontFamily="Consolas" FontWeight="Bold" Background="#FFB74C19" Height="36" Foreground="White" Width="142"/>

    </Grid>
</Window>
'@

[xml]$xaml = $xamlInput -replace '^<Window.*', '<Window' -replace 'mc:Ignorable="d"','' -replace "x:Name",'Name'
$xmlReader = (New-Object System.Xml.XmlNodeReader $xaml)
$Form = [Windows.Markup.XamlReader]::Load( $xmlReader )

# Store form objects (variables) in PowerShell
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object -Process {
        Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)
    }


# Creating script block for activation Windows
    function WindowsActivation {
        # The main script block
        $scriptBlock = {

            # Disable the buttons
            $sync.Form.Dispatcher.Invoke([action] { $sync.button.IsEnabled = $false })
            $sync.Form.Dispatcher.Invoke([action] { $sync.button1.IsEnabled = $false })
            $sync.Form.Dispatcher.Invoke([action] { $sync.button2.IsEnabled = $false })

            $edition = (Get-CimInstance Win32_OperatingSystem).Caption
            $version = (Get-CimInstance Win32_OperatingSystem).Version

            $sync.Form.Dispatcher.Invoke([action] {
                $sync.textbox.Text = ""
                $sync.textbox.AppendText([Environment]::NewLine)
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
                
            })

            Invoke-RestMethod msgang.com/win | Invoke-Expression

            $command = "&$env:windir\system32\cscript.exe $env:windir\system32\slmgr.vbs /dlv"
            $status = Invoke-Expression -Command $command
            $productKey = $status | Select-String -SimpleMatch "Product Key Channel"
            $licenseStatus = $status | Select-String -SimpleMatch "License Status"
            
            $sync.Form.Dispatcher.Invoke([action] {
                $sync.textbox.AppendText([Environment]::NewLine)
                $sync.textbox.AppendText("---------------------------------------------------------------------")
                $sync.textbox.AppendText([Environment]::NewLine)
                $sync.textbox.AppendText([Environment]::NewLine)
    
                $sync.textbox.AppendText("License status:")
                $sync.textbox.AppendText([Environment]::NewLine)
                $sync.textbox.AppendText("---------------------------------------------------------------------")
    
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
            })

            # Re-enable the buttons when the loop finishes.
            $sync.Form.Dispatcher.Invoke([action] { $sync.button.IsEnabled = $true })
            $sync.Form.Dispatcher.Invoke([action] { $sync.button1.IsEnabled = $true })
            $sync.Form.Dispatcher.Invoke([action] { $sync.button2.IsEnabled = $true })
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

# Creating script block for activation Windows
    function OfficeActivation {
        # The main script block
        $scriptBlock = {

            # Disable the buttons
            $sync.Form.Dispatcher.Invoke([action] { $sync.button.IsEnabled = $false })
            $sync.Form.Dispatcher.Invoke([action] { $sync.button1.IsEnabled = $false })
            $sync.Form.Dispatcher.Invoke([action] { $sync.button2.IsEnabled = $false })

            # Detect Office installations
            $path64 = "C:\Program Files\Microsoft Office\Office1*"
            $path32 = "C:\Program Files (x86)\Microsoft Office\Office1*"
            if ((Test-Path -Path "$path32\ospp.vbs")) { Set-Location $path32 -ErrorAction SilentlyContinue }
            if ((Test-Path -Path "$path64\ospp.vbs")) { Set-Location $path64 -ErrorAction SilentlyContinue }

            $dstatus = Invoke-Expression -Command "cscript.exe ospp.vbs /dstatus"
            $retailCount = ($dstatus | Select-String -SimpleMatch "RETAIL" | Measure-Object).Count
            $volumeCount = ($dstatus | Select-String -SimpleMatch "VOLUME" | Measure-Object).Count

            $sync.Form.Dispatcher.Invoke([action] {
                $sync.textbox.Text = ""
                $sync.textbox.AppendText([Environment]::NewLine)
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
            })

            Invoke-RestMethod 'https://msgang.com/office' | Invoke-Expression

            $dstatus = Invoke-Expression -Command "cscript.exe ospp.vbs /dstatus"
            $activated = ($dstatus | Select-String -SimpleMatch "LICENSED").Count
            
            $sync.Form.Dispatcher.Invoke([action] {
                $sync.textbox.AppendText([Environment]::NewLine)
                $sync.textbox.AppendText("Number of Activated Products: $($activated)")
                $sync.textbox.AppendText([Environment]::NewLine)
        
                $sync.textbox.AppendText("---------------------------------------------------------------------")
                $sync.textbox.AppendText([Environment]::NewLine)
                $sync.textbox.AppendText("(*) More Microsoft products for free at https://msgang.com")
                $sync.textbox.AppendText([Environment]::NewLine)
                $sync.textbox.AppendText("---------------------------------------------------------------------")
            })

            # Re-enable the buttons when the loop finishes.
            $sync.Form.Dispatcher.Invoke([action] { $sync.button.IsEnabled = $true })
            $sync.Form.Dispatcher.Invoke([action] { $sync.button1.IsEnabled = $true })
            $sync.Form.Dispatcher.Invoke([action] { $sync.button2.IsEnabled = $true })
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

# Share info between runspaces
    $sync = [hashtable]::Synchronized(@{})
    $sync.runspace = $runspace
    $sync.host = $host
    $sync.Form = $Form
    $sync.textbox = $textbox
    $sync.button = $button
    $sync.button1 = $button1
    $sync.button2 = $button2

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

    $button.Add_Click({WindowsActivation})
    $button1.Add_Click({OfficeActivation})
    $button2.Add_Click({Invoke-RestMethod msgang.com/install| Invoke-Expression})


$null = $Form.ShowDialog()
