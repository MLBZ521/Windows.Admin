Windows SysAdmin Scripts
======

In this repository are various scripts that I have written to automate tasks from a system administration level.


#### compareFolders.ps1 ####

Description:  This script compares files in folders (and more).

* Borrowed and modified from:  helios456 @ http://serverfault.com/questions/532065/how-do-i-diff-two-folders-in-windows-powershell


#### delete_GPOsByGUID.ps1 ####

Description:  This script deletes Group Policy Objects based on their GUID and display information for each before deletion.

Syntax:  To use this script, you would call it with the domain and file you want to use.

Examples:
```Delete_GPOsByGUID.ps1 -Domain my.domain.com -File GPO-GUIDS.txt```


#### find_MissingEquipment.ps1 ####

Description:  This script checks the serial number of the local machine and if the machine's serial number matches one of the missing serial numbers in the array, it will send an email to the specified address.


#### move_UserProfile.ps1 ####

Description:  This script will move the specified users data from one PC to another PC.  The specified folders will be moved, if more folders are needed for your environment, they can be added in the $Profile array.


#### report_AcrobatLicenseKey.ps1 ####

Description:  This script pulls the encrypted Acrobat product key from a remote or local computer and then decrypts it; also allows for direct key input.  It logs all information as is goes.

* Function borrowed from:  https://gallery.technet.microsoft.com/scriptcenter/ConvertFrom-EncryptedAdobeK-1b1160e3


#### report_InactiveComputers.ps1 ####

Description:  This script finds inactive computer accounts in AD and pulls relevant information.  This information can then be displayed to screen, exported to a csv, or the accounts can be disabled.


#### report_Inventory.ps1 ####

Description:  This script pulls information for inventory purposes and then saves it to a csv file which can be imported into Excel.
