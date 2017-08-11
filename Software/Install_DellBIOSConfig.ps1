<#

Script Name:  Install_DellBIOSConfig.ps1
By:  Zack Thompson / Created:  6/11/2015
Version:  1.0 / Updated:  6/11/2015 / By:  ZT

Description:  This script installs a customized BIOS Configuration file
	for Dell Systems.

	(*This should be a computer shutdown script.*)

#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**                DO NOT CLOSE THIS WINDOW               **"
Write-Host "**                                                       **"
Write-Host "**               Updating BIOS Configuration             **"
Write-Host "**                                                       **"
Write-Host "**  This window will automatically close once complete.  **"
Write-Host "**                                                       **"
Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host

# ============================================================
# Check Processor Architecture (32bit vs 64bit).
Function ProcArch {
	$cpuArch="AMD64"
	If ($ENV:Processor_Architecture -eq $cpuArch) {
		$script:osArch="SysWow64"
	}
	Else {
	$script:osArch="System32"
	}
}
# ============================================================
# Check to see directory exists and create it if not.
Function DirCheck {
	If (!(Test-Path -Path $BIOS)) {
		New-Item -ItemType directory -Path $BIOS
	}
}
# ============================================================
# Check to see log-file exists.
Function LogCheck {
	If (!(Test-Path -Path $BIOS\$LogFile)) {
		$script:noLog="NotInstalled"
	}
}
# ============================================================
# Copy files to local disk and execute.
Function CopyExecute {
	Set-Location -Path $BIOS
	Copy-Item $serverLocation\* $BIOS
	Start-Process BIOSConfig.exe -Wait
}
# ============================================================
# Create log-file and append install date to file.
Function CreateLog {
	Set-Location -Path $BIOS
	$date = Get-Date
	Write-Output "BIOS Configured on:" $date | Out-File $LogFile
}
# ============================================================
# Script Body
# ============================================================

# Call ProcArch Function
ProcArch

# Define variables & locations
$serverLocation="\\server\IT\GPO Files\Dell\BIOS\StaffConfig"
$IT_Staging="C:\Windows\$osArch\IT_Staging"
$BIOS="$IT_Staging\BIOS"
$LogFile="Log_BIOSConfig.txt"

# Call DirCheck Function
DirCheck
# Call LogCheck Function
LogCheck

If ($noLog -eq "NotInstalled") {
	Write-Host "BIOS has not been configured on this PC..."
	Write-Host "Configuring BIOS..."

	# Call CopyExecute Function
	CopyExecute
	Write-Host "BIOS Configuration Complete."

	# Call CreateLog Function
	CreateLog
	Write-Host "Created log file."
}
Else {
	Write-Host "BIOS has already been configured on this PC."
}
Write-Host "Script completed successfully!"