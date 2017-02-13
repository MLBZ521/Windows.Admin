<#

Script Name:  Printers.ps1
By:  Zack Thompson / Created:  8/10/2016
Version:  1.1 / Updated:  8/10/2016 / By:  ZT

Description:  This script installs printers from a print server.

Note:  This script will not work on Windows 7 or prior OS.

#>

# ============================================================
# Define Variables
# ============================================================

$Location = "Name or Prefix of Printer"
$PrintServer = "PrintServer"

# ============================================================
# Script Body
# ============================================================

$PrintersToInstall = Get-Printer -ComputerName $PrintServer | Where-Object { $_.Name -Match $Location -and $_.DeviceType -eq "Print" }

ForEach ($PrintQueue in $PrintersToInstall) {
    Write-Host "Installing print queue:  \\$($PrintServer)\$($PrintQueue.ShareName)"
    rundll32 printui.dll,PrintUIEntry /ga /n "\\$($PrintServer)\$($PrintQueue.ShareName)"
}

# eos