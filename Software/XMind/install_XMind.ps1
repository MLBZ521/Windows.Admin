<#

Script Name:  install_XMind.ps1
By:  Zack Thompson / Created:  9/18/2014
Version:  2.2 / Updated:  4/28/2017 / By:  ZT

Description:  This script installs XMind. (*This has to be a computer login script.*)

#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**                    Installing XMind                   **"
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
	If (!(Test-Path -Path $localLocation)) {
		New-Item -ItemType directory -Path $localLocation
	}
}
# ============================================================
# Check to see log-file exists.
Function LogCheck {
	If (!(Test-Path -Path $localLocation\"Log_$SWUpdate.txt")) {
		$script:noLog="NotInstalled"
	}
}
# ============================================================
# Copy files to local disk and execute.
Function CopyExecute {
	Set-Location -Path $localLocation
	Copy-Item $serverLocation\* $localLocation
	Start-Process $SWinstaller -ArgumentList "/verysilent /norestart /loadinf=xmind_install-settings.inf" -Wait
}
# ============================================================
# If 32bit, delete Bonjour installer files to prevent XMind from asking end
# user to install on program launch, and make changes to registry.
Function arch32Do {
	Remove-Item "C:\Program Files\XMind\thirdparty\Bonjour.msi" -force
	Remove-Item "C:\Program Files\XMind\thirdparty\Bonjour64.msi" -force

    # I believe I did this so I could tell if *I* installed the software
    # (via this script) on the users machine or if the user installed it themselves.
	Set-Location -Path Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\XMind_is1
	Set-ItemProperty . DisplayName $SWversion
}
# ============================================================
# If 64bit, delete Bonjour installer files to prevent XMind from asking end
# user to install on program launch, and make changes to registry.
Function arch64Do {
	Remove-Item "C:\Program Files (x86)\XMind\thirdparty\Bonjour.msi" -force
	Remove-Item "C:\Program Files (x86)\XMind\thirdparty\Bonjour64.msi" -force

	# I believe I did this so I could tell if *I* installed the software
    # (via this script) on the users machine or if the user installed it themselves.
    Set-Location -Path Registry::HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\XMind_is1
	Set-ItemProperty . DisplayName $SWversion
}
# ============================================================
# Create log-file and append install date to file.
Function CreateLog {
	Set-Location -Path $localLocation
	$date = Get-Date
	Write-Output $SWversion "Installed on:" $date | Out-File "Log_$SWUpdate.txt"

	# This generic log-file is for the user preferences script.
	Write-Output "XMind is installed, copy user preferences on login." $date | Out-File Log_XMind.txt -append
}
# ============================================================

# ============================================================
# Script Body
# ============================================================

# Call ProcArch Function
ProcArch

# Define variables & locations
$Org="- Organization Name"  ### <-- Update this line to the name of your organization (optional) <-- ###
$SWUpdate="xmind-windows-3.5.3.201506180105"  ### <-- Update this line to the installer file name (minus the .exe) <-- ###
$SWversion="XMind 6 (v3.5.3.201506180105) $Org"  ### <-- Update this line to the latest version (used in above function) <-- ###

$SWname="XMind"
$serverLocation="\\fileServer\Software Installers\$SWname\$SWUpdate"
$IT_Staging="C:\Windows\$osArch\IT_Staging"
$localLocation="$IT_Staging\$SWname"
$SWinstaller=$SWUpdate + ".exe"

# Call DirCheck Function
DirCheck

# Call LogCheck Function
LogCheck

If ($noLog -eq "NotInstalled") {
    Write-Host "XMind or latest version is not installed on this PC"

	# Call CopyExecute Function
	CopyExecute

    Write-Host "Installing XMind..."
    Write-Host "Removing files and updating registry after installation..."

	If ($osArch -eq "System32") {
		# Call arch32Do Function
		arch32Do
	}
	Else {
		# Call arch64Do Function
		arch64Do
	}

	# Call CreateLog Function
	CreateLog

	Write-Host "Created log file."
}
Else {
    Write-Host "Latest version of XMind is currently installed on this PC."
}

Write-Host "Script completed successfully!"
# eos
