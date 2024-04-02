<#

Script Name:  Update-CiscoVPNProfile.ps1
By:  Zack Thompson / Created:  1/4/2017
Version:  1.3.0 / Updated:  4/1/2024 / By:  ZT

Description:  This script modifies the content in the Cisco Preferences XML file.

Note:  The changes made in the script are not picked up until after the Cisco Client is restarted
	(if the Cisco service is enabled) as the service starts before the script is run.

Usage:  *This should be configured as a user login script.*

#>

# ============================================================
# Define Variables
# ============================================================
$DefaultUser = $env:username
$DefaultHostName = "sslvpn.my.org"
# $AutoConnectOnStart = "true"
$AutoUpdate = "true"
$Cisco_Root = "C:\Users\${DefaultUser}\AppData\Local\Cisco"
$AnyConnect_XMLlocation = "${Cisco_Root}\Cisco AnyConnect Secure Mobility Client\preferences.xml"
$SecureClient_XMLlocation = "${Cisco_Root}\Cisco Secure Client\VPN\preferences.xml"

# ============================================================
# Script Body
# ============================================================

# Check to see if the XML exists; exit if not.
# If (!(Test-Path -Path $AnyConnect_XMLlocation) -and !(Test-Path -Path $SecureClient_XMLlocation)) {
# 	Exit
# }

$AnyConnect_XMLlocation, $SecureClient_XMLlocation | ForEach-Object {

	If ( !( Test-Path -Path $_ ) ) {
		Continue
	}

	# Get the content of the XML file.
	$CiscoXML = [xml](Get-Content $_)
	$Node = $CiscoXML.AnyConnectPreferences

	# Set values
	$Node.DefaultUser = $DefaultUser
	$Node.DefaultHostName = $DefaultHostName

	# Since these are nested elements and do not exist by default,
		# they have to be assigned values differently.
	<#  Disabled the AutoConnectOnStart option since the
			Cisco VPN Service is set to Automatic by default.
		This lead to the following scenarios I decided to avoid:
			- At login the VPN client launches, asking for credentials,
				even if the user doesn't need/use VPN for their job function.
			- If we disable the Cisco VPN Service/client from launching at login,
				then users that need it (may) not be able to find it.

	If ( $Node.ControllablePreferences.AutoConnectOnStart -eq $null ) {
		$newNode_ACOS = $CiscoXML.CreateElement("AutoConnectOnStart")
		$newNode_ACOS.InnerXml = $AutoConnectOnStart
		$CiscoXML.SelectSingleNode("//ControllablePreferences").AppendChild($newNode_ACOS)
	}
	Else {
		$CiscoXML.SelectSingleNode("//AutoConnectOnStart").InnerXml = $AutoConnectOnStart
	}
	#>

	If ( $Node.ControllablePreferences.AutoUpdate -eq $null ) {
		$newNode_AU = $CiscoXML.CreateElement("AutoUpdate")
		$newNode_AU.InnerXml = $AutoUpdate
		$CiscoXML.SelectSingleNode("//ControllablePreferences").AppendChild($newNode_AU)
	}
	Else {
		$CiscoXML.SelectSingleNode("//AutoUpdate").InnerXml = $AutoUpdate
	}

	# Save the changes to the XML file.
	$CiscoXML.Save($_)

}

# This part is optional -- if you leave it, it causes AnyConnect to open on the
	# desktop at login -- this confuses users and has resulted in a couple tickets/questions.
<# Restart the Cisco VPN Client to pick up the changes from preferences.xml.
$CiscoPath = Get-Process -Name vpnui | Select-Object Path
Get-Process -Name vpnui | % { $_.CloseMainWindow() }
Start-Process $CiscoPath.Path -
#>

# eos