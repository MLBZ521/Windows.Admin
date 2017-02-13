<#

Script Name:  update_CiscoVPN-Profile.ps1
By:  Zack Thompson / Created:  1/4/2017
Version:  1.0 / Updated:  1/4/2017 / By:  ZT

Description:  This script modifies the content in the Cisco Preferences XML file.

Note:  ( *This should be configured as a user startup script.* )

#>

# ============================================================
# Define Variables
# ============================================================
$DefaultUser = $env:username
$DefaultHostName =  "sslvpn.asu.edu"
# $AutoConnectOnStart = "true"
$AutoUpdate = "true"
$XMLlocation = "C:\Users\$env:username\AppData\Local\Cisco\Cisco AnyConnect Secure Mobility Client\preferences.xml"


# ============================================================
# Script Body
# ============================================================

# Check to see if the XML exists; exit if not.
If (!(Test-Path -Path $XMLlocation)) {
    Exit
}

# Get the content of the XML file.
$CiscoXML = [xml](Get-Content $XMLlocation)
$Node = $CiscoXML.AnyConnectPreferences

# Set values
$Node.DefaultUser = $DefaultUser
$Node.DefaultHostName = $DefaultHostName


# Since these are nested elements and do not exist by default, they have to be assigned values differently.
If ( $Node.ControllablePreferences.AutoConnectOnStart -eq $null ) {
    $newNode_ACOS = $CiscoXML.CreateElement("AutoConnectOnStart")
    $newNode_ACOS.InnerXml = $AutoConnectOnStart
    $CiscoXML.SelectSingleNode("//ControllablePreferences").AppendChild($newNode_ACOS)
}
Else {
    $CiscoXML.SelectSingleNode("//AutoConnectOnStart").InnerXml = $AutoConnectOnStart
}

If ( $Node.ControllablePreferences.AutoUpdate -eq $null ) {
    $newNode_AU = $CiscoXML.CreateElement("AutoUpdate")
    $newNode_AU.InnerXml = $AutoUpdate
    $CiscoXML.SelectSingleNode("//ControllablePreferences").AppendChild($newNode_AU)
}
Else {
    $CiscoXML.SelectSingleNode("//AutoUpdate").InnerXml = $AutoUpdate
}

# Save the changes to the XML file.
$CiscoXML.Save($XMLlocation)

# Restart the Cisco VPN Client to pick up the changes from preferences.xml.
$CiscoPath = Get-Process -Name vpnui | Select-Object Path
Get-Process -Name vpnui | % { $_.CloseMainWindow() }
Start-Process $CiscoPath.Path -

# eos