<#

Script Name:  uninstall_Quicktime.ps1
By:  Zack Thompson | Created:  4/20/2016
Version:  1.0 | Updated:  4/20/2016 | By:  ZT

Description:  This script uninstalls Quicktime.

Note:  ( *This should be configured as a computer startup script.* )

#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**                DO NOT CLOSE THIS WINDOW               **"
Write-Host "**                                                       **"
Write-Host "**                 Uninstalling Quicktime                **"
Write-Host "**                                                       **"
Write-Host "**  This window will automatically close once complete.  **"
Write-Host "**                                                       **"
Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host

# ============================================================
# Define variables
$SearchTerm = "QuickTime"

# ============================================================
# Script Body
# ============================================================

# Locations to check for installs
$FindSoftware = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | 
	Get-ItemProperty | 
	Where-Object {$_.DisplayName -match $SearchTerm } | 
	Select-Object -Property DisplayName, UninstallString

Write-Host $FindSoftware.DisplayName,  $FindSoftware.UninstallString

# Parse each key that matched the search term above.
ForEach ($SWItem in $FindSoftware) {
	Write-Host "Removing..."  $SWItem.DisplayName
	$Uninstall = $SWItem.UninstallString -replace "/I","/X"
	Start-Process "cmd.exe" -ArgumentList "/c $Uninstall /quiet /norestart"
}
Write-Host "$searchTerm has been uninstalled."
Write-Host "Script completed successfully!"

# eos