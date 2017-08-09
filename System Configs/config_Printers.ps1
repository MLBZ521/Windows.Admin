<#

Script Name:  config_Printers.ps1
By:  Zack Thompson / Created:  8/10/2016
Version:  2.0 / Updated:  2/8/2017 / By:  ZT

Description:  This script installs or uninstalls printers from the designated print server.
    (This script was written so that it would work on Windows 7 systems that cannot utilize the
     new Add-Printer cmdlets available in Windows 8+ machines.)

Syntax:  To use this script, you would call it with the action and printer(s) you want to install.

Examples:

    1)  Install/Uninstall a single printer:
            config_Printers.ps1 -Action Install -Server myPrintServer -Printers HR_HP_LaserJet_Color_M451dn
            config_Printers.ps1 -Action Uninstall -Server myPrintServer -Printers Graphics_Konica_C458

    2)  Install/uninstall multiple printers that begin with a 'location' designator:
            config_Printers.ps1 -Action Install -Server myPrintServer -Printers HR_
            config_Printers.ps1 -Action Uninstall -Server myPrintServer -Printers Graphcis_

#>

# ============================================================
# Define Variables
# ============================================================

Param (
    [Parameter(Mandatory=$True,Position=0)][string]$Action,
    [Parameter(Mandatory=$True,Position=1)][string]$Server,
    [Parameter(Mandatory=$True,Position=2)][string]$Printers
)

# ============================================================
# Script Body
# ============================================================

$AvailablePrinters = @()
$AvailablePrinters += (net view \\$($Server))

ForEach ( $Printer in $AvailablePrinters ) {
    If ( $Printer -match $Printers ) {
        $PrintQueue = $Printer.Split('  ')[0]

        If ( $Action -eq "Install" ) {
            Write-Host "Installing print queue:  \\$($Server)\$($PrintQueue)"
            rundll32 printui.dll,PrintUIEntry /ga /n "\\$($Server)\$($PrintQueue)"
        }
        ElseIf ( $Action -eq "Uninstall" ) {
            Write-Host "Uninstalling print queue:  \\$($Server)\$($PrintQueue)"
            rundll32 printui.dll,PrintUIEntry /gd /n "\\$($Server)\$($PrintQueue)"
        }
    }
}

# eos
