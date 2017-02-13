<#

Script Name:  Printers.ps1
By:  Zack Thompson / Created:  8/10/2016
Version:  1.2 / Updated:  8/22/2016 / By:  ZT

Description:  This script installs (or uninstalls) printers from a print server.

Note:  This script *will* work on Windows 7 and prior OS.

#>

# ============================================================
# Define Variables
# ============================================================

$Location = "Name or Prefix of Printer"
$PrintServer = "PrintServer"

# ============================================================
# Script Body
# ============================================================

$PrintersToInstall = @()
$PrintersToInstall += (net view $PrintServer)

# To uninstall printers, change the /ga switch to /gd
ForEach ( $Printer in $PrintersToInstall ) {
    If ( $Printer -match $Location ) {
        $PrintQueue = $Printer.Split('  ')[0]
        Write-Host "Installing print queue:  \\$($PrintServer)\$($PrintQueue)"
        rundll32 printui.dll,PrintUIEntry /ga /n "\\$($PrintServer)\$($PrintQueue)"
    }
}

# eos