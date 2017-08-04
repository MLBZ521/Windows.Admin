Windows SysAdmin Scripts
======

Scripts in this folder are various scripts that I have written for use in a SysAdmin role.


#### delete_GPOsByGUID.ps1 ####

Description:  This script deletes Group Policy Objects based on their GUID and display information for each before deletion.

Syntax:  To use this script, you would call it with the domain and file you want to use.

Examples:  Delete_GPOsByGUID.ps1 -Domain my.domain.com -File GPO-GUIDS.txt



#### report_InactiveComputers.ps1 ####

Description:  This script finds inactive computer accounts in AD and pulls relevant information.  This information can then be displayed to screen, exported to a csv, or the accounts can be disabled.



#### report_Inventory.ps1 ####

Description:  This script pulls information for inventory purposes and then saves it to a csv file which can be imported into Excel.