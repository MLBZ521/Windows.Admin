Windows Software Scripts
======

In this repository are various scripts that I have written to maintain Windows Software.  Most are written for use in a GPO.


#### install_DellBIOSConfig.ps1 ####

Description:  This script installs a customized BIOS Configuration file for Dell Systems.


#### install_DellDockFirmware.ps1 ####

Description:  This script checks hardware characteristics that make an assumption that a laptop is docked.  If it is, it will install the firmware update if applicable.  (On the Dell E-Port Advanced Docks using DisplayPort, there an issue with external monitors flickering and/or not display the video feed.  To apply this firmware, the laptop has to be docked, with power, and an external monitor connected.)


#### install_Drivers.ps1 ####

Description:  This script checks the model of the computer, checks driver versions and updates them if not the latest configured for it.


#### install_Updates.ps1 ####

Description:  This script installs .msu update files that are saved to a network share.  It checks the OS Version to see if it matches one that updates	are provided for, if so, it will install them if they are not already installed.

* Borrowed and modified from:  http://randygray.com/powershell-install-multiple-windows-updates-msu/


#### uninstall_Quicktime.ps1 ####

Description:  This script uninstalls Quicktime.


#### update_CiscoVPN-Profile.ps1 ####

Description:  This script modifies the content in the Cisco Preferences XML file.


#### XMind ####

Description:  A collection of scripts and files I used to deploy the XMind application.
