<#

Script Name:  Printers.ps1
By:  Zack Thompson / Created:  8/10/2016
Version:  1.0 / Updated:  8/10/2016 / By:  ZT

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

$PrintersToInstall = Get-Printer -ComputerName $PrintServer | Where-Object { $_.Name -Match "$($Location)" }

ForEach ($PrintQueue in $PrintersToInstall) {
    # Write-Host "\\asuprint1\$($PrintQueue.ShareName)"
    Add-Printer -ConnectionName "\\asuprint1\$($PrintQueue.ShareName)"
}

# eos