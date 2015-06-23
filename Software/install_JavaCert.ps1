<#

Script Name:  install_JavaCert.ps1
By:  Zack Thompson | Created:  11/23/2014
Version:  1.2 | Updated:  3/9/2015 | By:  ZT

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
# Create log-file and append install date to file.
Function CreateLog {
Set-Location -Path $IT_Staging
$date = Get-Date
Write-Output "Certificate installed on:" $date | Out-File $logName -append
}
# ============================================================
# Script Body
# ============================================================

# Call ProcArch Function
ProcArch

# Define variables
$IT_Staging = "C:\Windows\$osArch\IT_Staging\Java\"
$keyTool = "C:\$ProgramFiles\Java\jre8\bin\keytool.exe"
$cacerts = "C:\$ProgramFiles\java\jre8\lib\security\cacerts"
$signingCert = "\\Share\Location\CodeSignCert.cer"
$alias = "fb9fcc11-bfe5-4613-88fc-460ea7e29230"
$alias = "abcdefgh-ijkl-mnop-qrst-uvwxyz012345"
$logName = "Log_JavaCertInstall.txt"

# Call DirCheck Function
DirCheck

# Installs Certificate
Write-Host "Installing Certificate..."
Start-Process $keytool -ArgumentList "-importcert -keystore `"$cacerts`" -storepass changeit -file `"$signingCert`" -alias `"$alias`" -noprompt" -NoNewWindow -Wait
Write-Host "Certificate has been installed on this PC."		

# Call CreateLog Function
CreateLog
Write-Host "Appended log file."
Write-Host "Script completed successfully!"

# eos