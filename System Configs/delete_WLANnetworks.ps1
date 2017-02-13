<#

Script Name:  delete_WLANnetworks.ps1
By:  Zack Thompson / Created:  8/25/2016
Version:  1.0 / Updated:  8/25/2016 / By:  ZT

Description:  This script deletes WiFi networks that match $Name.

Note:  ( *This should be configured as a user login script.* )

#>

# ============================================================
# Define Variables
# ============================================================

$Name = "SSID Name"

# ============================================================
# Script Body
# ============================================================

$GetSSIDs = (netsh WLan Show Profile) | Select-String "All User Profile"

ForEach ($SSID in $GetSSIDs) { 
    $SSIDname = $SSIDs.ToString().Split(':')[1]
    If ( $SSIDname -match $Name ) {
        Write-Host "The network $SSIDname will be deleted..."
        netsh WLan Delete Profile Name="$($Name)"
     }
 }

# eos