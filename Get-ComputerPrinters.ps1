# © Copyright 2022, Jacob Searcy, All rights reserved.

function Get-ComputerPrinters {

    [CmdletBinding()]

    Param (
        [String]$ComputerName,
        [Switch]$Copy
    )

    Begin {
        $error.Clear()
    }

    Process {
        $ErrorActionPreference = 'silentlycontinue'
        $PrintScriptPrinters = @()
        Clear-Variable -Name Printer -ErrorAction SilentlyContinue
        Clear-Variable -Name DftPtr -ErrorAction SilentlyContinue
        #Check that the computer is online.
        If (Test-Connection $ComputerName -Count 1 -quiet) {
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
            #Check for empty variable. If there are no printers, displays "No printers loaded." 
            If ($PrintScriptPrinters -notmatch '\w') {
                Write-Host "No Printers Loaded" -ForegroundColor Red
            }else{
                #Setup $printscript to be formatted as the .vbs print script file.
                $printscript_header = 'Set objNetwork = CreateObject("WScript.Network")'
                If ($DftPtr -like '*\*') {
                    $printscript = Write-Output $printscript_header $PrintScriptPrinters $DftPtr
                }else{
                    $printscript = Write-Output $printscript_header $PrintScriptPrinters
                }
                Write-Host "Print Script Preview" -ForegroundColor Green
                $printscript
                #Save print script as .vbs in \\wnis01\ftp$
                $printscript | Out-File -FilePath "$PSScriptRoot\Files\$ComputerName.vbs"
                Write-Host "This print script has been saved to:" -ForegroundColor Cyan
                Write-Host "$PSScriptRoot\Files\$ComputerName.vbs"
                If ($Copy) {
                    $printscript | Out-File -FilePath "\\$ComputerName\c$\programdata\microsoft\windows\start menu\programs\startup\$ComputerName.vbs"
                }else{
                    #Ask if .vbs should be saved in the startup folder of the computer.
                    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
                    $Answer = [System.Windows.Forms.MessageBox]::Show("Copy Print Script to Startup Folder of $($ComputerName)?" , "Copy Print Script" , 4)
                    If ($Answer -eq 'No') {
                        Write-Host "Print script NOT copied" -ForegroundColor Green
                    }else{
                        $printscript | Out-File -FilePath "\\$ComputerName\c$\programdata\microsoft\windows\start menu\programs\startup\$ComputerName.vbs"
                        Write-Host "Print script copied" -ForegroundColor Green
                    }
                }
            }
        }else{
            #If computer is offline, display "Offline"
            Write-Host "$ComputerName is Offline" -ForegroundColor Red -BackgroundColor White
        }
    }

    End {}
}

$ComputerName = Read-Host "Enter Computer Name:"
Get-ComputerPrinters -ComputerName $ComputerName