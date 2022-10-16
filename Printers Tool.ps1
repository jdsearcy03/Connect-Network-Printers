# © Copyright 2022, Jacob Searcy, All rights reserved.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$image = [System.Drawing.Image]::FromFile("$PSScriptRoot\Files\search_icon.jpeg")
$printscript_header = 'Set objNetwork = CreateObject("WScript.Network")'
$Print_Script_Path = "$PSScriptRoot\Files"

$printer_form = New-Object System.Windows.Forms.Form
$printer_form.Text = "Connect Network Printers"
$printer_form.AutoSize = $true
$printer_form.StartPosition = 'CenterScreen'
$printer_form.TopMost = $true

$Computer_label = New-Object System.Windows.Forms.Label
$Computer_label.Text = "Enter Computer Name:"
$Computer_label.Location = New-Object System.Drawing.Point(5,8)
$Computer_label.Size = New-Object System.Drawing.Size(150,20)
$printer_form.Controls.Add($Computer_label)

$Printer_label = New-Object System.Windows.Forms.Label
$Printer_label.Text = "Print Script:"
$Printer_label.Location = New-Object System.Drawing.Point(5,70)
$Printer_label.Size = New-Object System.Drawing.Size(100,21)
$printer_form.Controls.Add($Printer_label)

$Computer_bar = New-Object System.Windows.Forms.TextBox
$Computer_bar.Location = New-Object System.Drawing.Size(5,28)
$Computer_bar.Size = New-Object System.Drawing.Size(500,21)
$Computer_bar.Text = $ComputerName

$Computer_button = New-Object System.Windows.Forms.Button
$Computer_button.Location = New-Object System.Drawing.Point(508,26)
$Computer_button.Size = New-Object System.Drawing.Size(21,21)
$Computer_button.BackgroundImage = $image
$Computer_button.BackgroundImageLayout = 'Zoom'
$printer_form.AcceptButton = $Computer_button

$search_accept = {
    $printer_form.AcceptButton = $Computer_button
}
$Computer_bar.add_MouseDown($search_accept)
$printer_form.Controls.Add($Computer_bar)

$get_printers = {
    $Printer_Box.Items.Clear()
    $Printer_Box.ForeColor = "Black"
    $ComputerName = $Computer_bar.Text
    If ($Computer_bar.Text -notmatch '\w') {
        $Printer_Box.ForeColor = "Red"
        $Printer_Box.Items.Add("Please Enter a Computer Name")
    }else{
        $printer_form.Enabled = $false
        $Printer_Box.Items.Add("Processing . . .")
        If (Test-Connection $ComputerName -Count 1 -quiet) {
            $PrintScriptPrinters = @()
            $printer_form.AcceptButton = $confirm_button
            #Collect a list of all User accounts and searches the registry for the printers of the currently loogged-on User by their SID.
            $UsersList = Get-ChildItem -Path \\$ComputerName\c$\Users\ -Name -Exclude ("Public", "WMISAdmin", "altiris","DiscoWinServer","DiscoWinClient")
            foreach ($User in $UsersList) {
                $SID = (New-Object System.Security.Principal.NTAccount($User)).Translate([System.Security.Principal.SecurityIdentifier]).Value
                $Printerlist = reg query "\\$ComputerName\HKU\$SID\Printers\Connections\"
                If ($Printerlist -match '\w') {
                    for ($i = 0; $i -le $Printerlist.Count; $i++) {
                        If ($PrinterList[$i] -match '\w') {
                            $Printer = Write-Output "objNetwork.AddWindowsPrinterConnection `"\\$($Printerlist[$i].Split(",,")[2])\$($Printerlist[$i].Split(",,")[3])`", `"$($Printerlist[$i].Split(",,")[3]).`""
                            $PrintScriptPrinters += $Printer
                        }
                    }
                    $DefaultPrinter = reg query "\\$ComputerName\HKU\$SID\Software\Microsoft\Windows NT\Currentversion\Windows" /v Device
                    If ($DefaultPrinter -match '\w') {
                        $DftPtr = Write-Output "objNetwork.SetDefaultPrinter `"$($DefaultPrinter[2].Split(" ,") | Where-Object {$_ -match "\\"})`""
                    }
                }else{
                    #Search the registry for the printers of the remaining users by the ntuser.dat file.
                    reg load "HKU\$User" "\\$ComputerName\c$\users\$User\ntuser.dat"
                    $Printerlist = reg query "HKU\$User\Printers\Connections"
                    If ($Printerlist -match '\w') {
                        for ($i = 0; $i -le $Printerlist.Count; $i++) {
                            If ($Printerlist[$i] -match '\w') {
                                $Printer = Write-Output "objNetwork.AddWindowsPrinterConnection `"\\$($Printerlist[$i].Split(",,")[2])\$($Printerlist[$i].Split(",,")[3])`", `"$($Printerlist[$i].Split(",,")[3]).`""
                                $PrintScriptPrinters += $Printer
                            }
                        }
                    }
                    reg unload "HKU\$User"
                }
            }
            $PrintScriptPrinters = $PrintScriptPrinters | Select-Object -Unique
            $printer_form.Enabled = $true
            #Check for empty variable. If there are no printers, displays "No printers loaded." 
            If ($PrintScriptPrinters.Count -eq 0) {
                $Printer_Box.Items.Clear()
                $Printer_Box.ForeColor = "Red"
                $Printer_Box.Items.Add("$ComputerName doesn't have any connected printers")
                $printer_form.AcceptButton = $Computer_button
            }else{
                #Setup $printscript to be formatted as the .vbs print script file.
                $Printer_Box.Items.Clear()
                If ($PrintScriptPrinters.Count -gt 1) {
                    foreach ($add_printer in $PrintScriptPrinters) {
                        $Printer_Box.Items.Add($add_printer)
                    }
                }else{
                    $Printer_Box.Items.Add($PrintScriptPrinters)
                }
                If ($DftPtr -like '*\*') {
                    $Printer_Box.Items.Add($DftPtr)
                }
            }
        }else{
            $printer_form.Enabled = $true
            $Printer_Box.Items.Clear()
            $Printer_Box.ForeColor = "Red"
            $Printer_Box.Items.Add("$ComputerName is Offline")
        }
    }
}
$Computer_button.Add_click($get_printers)
$printer_form.Controls.Add($Computer_button)

$Printer_Box = New-Object System.Windows.Forms.ListBox
$Printer_Box.Location = New-Object System.Drawing.Point(5,91)
$Printer_Box.AutoSize = $true
$Printer_Box.MinimumSize = New-Object System.Drawing.Size(522,300)
$Printer_Box.MaximumSize = New-Object System.Drawing.Size(0,300)
$Printer_Box.ScrollAlwaysVisible = $true
$Printer_Box.HorizontalScrollBar = $true
$Printer_Box.Items.Clear()

$printer_remove = {
    If ($Printer_Box.SelectedItem -match "\\") {
        $answer = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove this printer from the list?
$($Printer_Box.SelectedItem)", "Remove Printer", 4)
        If ($answer -eq "Yes") {
            $Printer_Box.Items.RemoveAt($($Printer_Box.Items.SelectedIndex))
        }
    }
}
$Printer_Box.add_Click($printer_remove)
$printer_form.Controls.Add($Printer_Box)

$confirm_button = New-Object System.Windows.Forms.Button
$confirm_button.Location = New-Object System.Drawing.Point(370,390)
$confirm_button.Size = New-Object System.Drawing.Size(75,23)
$confirm_button.Text = 'Confirm'

$confirm = {
    $ChosenPrinters = $Printer_Box.Items
    $printscript = $printscript_header + $ChosenPrinters
    $printscript | Out-File -FilePath "$Print_Script_Path\$ComputerName.vbs"
    $answer = [System.Windows.Forms.MessageBox]::Show("Your Print Script has been saved to: $Print_Script_Path\$ComputerName.vbs
Would you like to add the print script to the Startup Folder of $($ComputerName)?", "Print Script", 4)
    If ($answer -eq "Yes") {
        $printscript | Out-File -FilePath "\\$ComputerName\c$\programdata\microsoft\windows\start menu\programs\startup\$ComputerName.vbs"
    }
    [System.Windows.Forms.MessageBox]::Show("Process Complete")
    $Printer_Box.Items.Clear()
    $Computer_bar.Clear()
    $Computer_bar.Focused = $true
    $printer_form.AcceptButton = $Computer_button
}
$confirm_button.add_click($confirm)
$printer_form.Controls.Add($confirm_button)

$end_button = New-Object System.Windows.Forms.Button
$end_button.Location = New-Object System.Drawing.Point(450,390)
$end_button.Size = New-Object System.Drawing.Size(75,23)
$end_button.Text = 'Exit'
$end_button.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$printer_form.Controls.Add($end_button)

$printer_form.ShowDialog()