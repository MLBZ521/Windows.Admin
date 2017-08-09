Windows Desktop Scripts
======

In this repository are various scripts that I have written to configure the Windows Desktop OS.  Most are written for use in a GPO.


#### config_Printers.ps1 ####

Description:  This script installs or uninstalls printers from the designated print server.
* (This script was written so that it would work on Windows 7 systems that cannot utilize the new Add-Printer cmdlets available in Windows 8+ machines.)

Syntax:  To use this script, you would call it with the action and printer(s) you want to install.

Examples:

1.  Install/Uninstall a single printer:
```powershell
config_Printers.ps1 -Action Install -Server myPrintServer -Printers HR_HP_LaserJet_Color_M451dn
config_Printers.ps1 -Action Uninstall -Server myPrintServer -Printers Graphics_Konica_C458
```

2.  Install/uninstall multiple printers that begin with a 'location' designator:
```powershell
config_Printers.ps1 -Action Install -Server myPrintServer -Printers HR_
config_Printers.ps1 -Action Uninstall -Server myPrintServer -Printers Graphcis_
```

#### delete_WLANnetworks.ps1 ####

Description:  This script deletes WiFi networks that match $Name.


#### reset_WindowsUpdate.bat ####

Description:  This script resets the Windows Update Components.
* See:  https://support.microsoft.com/en-us/kb/971058