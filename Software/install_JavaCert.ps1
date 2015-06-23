<#

Script Name:  install_JavaCert.ps1
By:  Zack Thompson | Created:  11/23/2014
Version:  1.1 | Updated:  1/23/2015 | By:  ZT

Description:  This script installs a certificate into the default Java cacerts file.
	This certificate was used to sign the DeploymentRuleSet.jar package to whitelist
	Java applets that are accessed.  If this is not done, Firefox and Chrome
	will fail the authentication of the DeploymentRuleSet.jar package and it will
	only work on Internet Explorer.

Note:  (*This has to be a computer login script.*)

#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**                DO NOT CLOSE THIS WINDOW               **"
Write-Host "**                                                       **"
Write-Host "**		 Installing Certificate into Java Keystore       **"
Write-Host "**                                                       **"
Write-Host "**  This window will automatically close once complete.  **"
Write-Host "**                                                       **"
Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host

# ============================================================
# Check Processor Architecture (32bit vs 64bit).
Function ProcArch{
$cpuArch="AMD64"
if($ENV:Processor_Architecture -eq $cpuArch){
	$script:osArch="SysWow64"
	$script:ProgramFiles="Program Files (x86)"
	}
Else {
	$script:osArch="System32"
	$script:ProgramFiles="Program Files"
	}
}
# ============================================================
# Check to see directory exists and create it if not.
Function DirCheck {
if(!(Test-Path -Path $IT_Staging)){
    New-Item -ItemType directory -Path $IT_Staging
	}
}
# ============================================================
# Check to see log-file exists.
Function LogCheck{
if(!(Test-Path -Path $IT_Staging\"Log_$logName.txt")){
$script:noLog="NotInstalled"
	}
}
# ============================================================
# Create log-file and append install date to file.
Function CreateLog {
Set-Location -Path $IT_Staging
$date = Get-Date
Write-Output "Certificate installed on:" $date | Out-File "Log_$logName.txt"
}
# ============================================================

# ============================================================
# Script Body
# ============================================================

# Call ProcArch Function
ProcArch

# Define variables
$IT_Staging = "C:\Windows\$osArch\IT_Staging\Java\"
$keyTool = "C:\$ProgramFiles\Java\jre7\bin\keytool.exe"
$cacerts = "C:\$ProgramFiles\java\jre7\lib\security\cacerts"
$signingCert = "\\Share\Location\CodeSignCert.cer"
$alias = "fb9fcc11-bfe5-4613-88fc-460ea7e29230"
$logName = "Java7U75_CertInstall"

# Call DirCheck Function
DirCheck
# Call LogCheck Function
LogCheck

If ($noLog -eq "NotInstalled") {
	Write-Host "Certificate has not been install"
	Write-Host "Installing Certificate..."

	# Installs Certificate
	Start-Process $keytool -ArgumentList "-importcert -keystore `"$cacerts`" -storepass changeit -file `"$signingCert`" -alias `"$alias`" -noprompt" -NoNewWindow -Wait
		
	# Call CreateLog Function
	CreateLog
	Write-Host "Created log file."
	}
Else {
	Write-Host "Certificate has been installed on this PC."
	}

Write-Host "Script completed successfully!"

# eos