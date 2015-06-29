<#

Script Name:  install_XMind-prefs.ps1
By:  Zack Thompson / Created:  9/22/2014
Version:  1.1 / Updated:  6/29/2015 / By:  ZT

Description:  This script moves custom preferences files into place for 
	XMind users.  This script expects the files to have already been 
	moved by the XMind installer script.

(*This has to be a user login script.*)

#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**              Moving XMind Preference Files            **"
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
	If (!(Test-Path -Path $localLocation\Log_XMind.txt)) {
		$script:noLog="NotInstalled"
	}
}
# ============================================================
# Move preference files into place (this has to be done per user).
Function MovePrefs {
	Set-Location -Path $localLocation
	$prefDIR="$env:AppData\XMind\workspace-cathy\.metadata\.plugins\org.eclipse.core.runtime\.settings\"
		If (!(Test-Path -Path $prefDIR)) {
			New-Item -ItemType directory -Path $prefDIR
		}
		If (!(Test-Path -Path "$prefDIR\net.xmind.verify.prefs")) {
			Copy-Item net.xmind.verify.prefs $prefDIR
		}
		If (!(Test-Path -Path "$prefDIR\org.xmind.cathy.prefs")) {
			Copy-Item org.xmind.cathy.prefs $prefDIR
		}
}
# ============================================================

# ============================================================
# Script Body
# ============================================================

# Call ProcArch Function
ProcArch
Write-Host "OS Architecture is:  $osArch"

# Define variables & locations
$SWname="XMind"
$IT_Staging="C:\Windows\$osArch\IT_Staging"
$localLocation="$IT_Staging\$SWname"

# Call DirCheck Function
DirCheck
# Call LogCheck Function
LogCheck

If ($noLog -eq "NotInstalled") {
Write-Host "XMind is not Installed on this PC"
}
Else {
Write-Host "XMind is currently installed on this PC, beginning per user configurations..."
# Call MovePrefs Function
MovePrefs
Write-Host "Moved preference files into place for current user."
}

Write-Host "Script completed successfully!"
# eos