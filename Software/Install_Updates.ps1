<#
Script Name:  install_Updates.ps1
By:  Zack Thompson / Created:  3/9/2015
Version:  1.0 / Updated:  3/9/2015 / By:  ZT

Description:  This script installs .msu update files that are saved to a 
	network share.  It checks the OS Version to see if it matches one that 
	updates	are provided for, if so, it will install them if they are not 
	already	installed.

	(*This should be run as a computer shutdown script.*)

Borrowed and modified from:  http://randygray.com/powershell-install-multiple-windows-updates-msu/

#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**                DO NOT CLOSE THIS WINDOW               **"
Write-Host "**                                                       **"
Write-Host "**              Installing Microsoft Updates             **"
Write-Host "**                                                       **"
Write-Host "**  This window will automatically close once complete.  **"
Write-Host "**                                                       **"
Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host

# ============================================================
# Define Variables
# ============================================================
# Set $Path to repository directory
$Path = "\\server\Share\Microsoft Patches"

# ============================================================
# Function to Install Update.
Function Install-MSU($PatchesPath) {
	# Check patch repository for .msu's.
	$msus = Get-ChildItem -Path $PatchesPath *.msu -Recurse
	
	ForEach ($msu in $msus) {
        Write-Host $msu
		# Spilt file name to get KB Article Number.
		$Update = $msu.Name -Split '-'
		$KBArticle = $Update[1]

		# Check if update is already installed.
		$HotFix = Get-HotFix -id $KBArticle -ErrorAction 0
		Write-Host "Checking if $KBArticle is installed..."
                
		# If update is not installed, install it.
		If ($HotFix -eq $null) {
			Write-Host "Update is not installed; installing..."
			$Command = "$PatchesPath\$msu"
            $Switches = "/quiet /norestart"
			$Parameters = "`"$Command`" $Switches"
			$Install = [System.Diagnostics.Process]::Start("wusa",$Parameters)
			$Install.WaitForExit()
            Write-Host "$KBArticle has been installed."
            Write-Host
		}
		# run if update is installed
		Else {
			Write-Host "$KBArticle is already installed."
            Write-Host
		}
	}
}
# ============================================================
# Script Body
# ============================================================
# Get the Operating System and Architecture
$OS = Get-WmiObject -Query "Select Caption, OSArchitecture from win32_OperatingSystem"

# If Windows 7
If ($OS.Caption -Match 'Windows 7') {
	If ($OS.OSArchitecture -Match '64-bit') {
		$Folder = 'Windows 7\x64'
		$PatchesPath = "$Path\$Folder"
		# Call Install-MSU Function.
		Install-MSU($PatchesPath)
	}
	Else {
	$Folder = "Windows 7\x86"
	$PatchesPath = "$Path\$Folder"
	# Call Install-MSU Function.
	Install-MSU($PatchesPath)
	}
}
# If Server 2008 R2
ElseIf ($OS.Caption -Match 'Server 2008 R2') {
	$Folder = 'Server 2008 R2'
	$PatchesPath = "$Path\$Folder"
	# Call Install-MSU Function.
	Install-MSU($PatchesPath)
}
Else {
    Write-Host "This Operating System does not have any assigned updates at this time."
}