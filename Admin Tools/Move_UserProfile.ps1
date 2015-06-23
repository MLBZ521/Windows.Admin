<#
Script Name:  Move_UserProfile.ps1
By:  Zack Thompson / Created:  2/3/2015
Version:  .1 / Updated:  2/3/2015 / By:  ZT

Description:  This script will move the specified users data from one PC
	to another PC.  The specificed folders will be moved, if more folders
	are needed for your environment, they can be added in the $Profile 
	variable.

#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**                  Profile Maintenance                  **"
Write-Host "**                                                       **"
Write-Host "***********************************************************"
Write-Host "***********************************************************"

# ============================================================
# Define Variables
# ============================================================
$user = Read-Host "Enter User Name"
$NewPC = Read-Host "Enter New Computer Name"
$OldPC = Read-Host "Enter Old Computer Name"
$OldLocation = "\\$($oldPC)\C$\$($user)\"
$NewLocation = "\\$($NewPC)\C$\$($user)\"
$Profile = "Downloads",
"Music",
"Pictures",
"Videos",
"AppData\Local\Mozilla",
"AppData\Local\Google",
"AppData\Roaming\Mozilla",
"\AppData\Local\Microsoft\Outlook\PST_Files\"

# ============================================================
# Function to prompt for action.
Function Prompt {
	$Caption = "Choose Action";
	$Message = "Do you want to copy $($user)'s files from $($OldPC) to $($NewPC)?"
	$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Yes";
	$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No","No";
	$Choices = [System.Management.Automation.Host.ChoiceDescription[]]($Yes,$No);
	$Answer = $Host.UI.PromptForChoice($Caption,$Message,$Choices,1)
}
# ============================================================
# Function that moves files.
Function MoveFiles {
	If ($answer -ne 1) {
		Write-Host "Moving $($user)'s data from $($OldPC) to $($NewPC)..."
		
		ForEach ($Folder in $Profile) {
			$OldFolder = $OldLocation + "\" + $Folder
			$NewFolder = $NewLocation + "\" + $Folder
			
			Copy-Item -ErrorAction silentlyContinue -recurse $OldFolder\* $NewFolder
		
	<#	If ($Folder -eq "AppData\Local\Mozilla") {
			rename-item $NewFolder "$($NewFolder).old"
		}
		If ($Folder -eq "AppData\Roaming\Mozilla") {
			rename-item $NewFolder "$($NewFolder).old"
		}
		If ($Folder -eq "AppData\Local\Google") {
			rename-item $NewFolder "$($NewFolder).old"
		} #>
		}
	}
	Else {
		Write-Host "You have selected to not move $($user)'s files."
		Exit
	}
}
# ============================================================
# Script Body
# ============================================================

# Call Profile Function.
Prompt

$OnNewPC = Test-Connection -Computername $NewPC -Quiet
$OnOldPC = Test-Connection -Computername $OldPC -Quiet
#Write-Host $OnNewPC
#Write-Host $OnOldPC

If ($OnNewPC -ne $True) {
	Write-Host "$($NewPC) is not on line or accessible from here."
}
If ($OnOldPC -ne $True) {
	Write-Host "$($OldPC) is not on line or accessible from here."
	Exit
}
# Call MoveFiles Function.
# MoveFiles
Write-Host "Function Move Files should have ran."