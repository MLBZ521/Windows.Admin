<#

Script Name:  install_JavaCert.ps1
By:  Zack Thompson | Created:  11/23/2014
Version:  1.3 | Updated:  9/11/2015 | By:  ZT

Description:  This script installs a certificate into the default Java cacerts file.
	This certificate was used to sign the DeploymentRuleSet.jar package to whitelist
	Java applets that are pre-approved.  If this is not done, Firefox will fail the 
	authentication of the DeploymentRuleSet.jar package and it will	only work on 
	Internet Explorer.

Note:  (*This has to be a computer login script that runs AFTER a Java update if present.*)

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
If ($ENV:Processor_Architecture -eq $cpuArch) {
	$script:osArch="SysWow64"
	$script:ProgramFiles="Program Files (x86)"
	}
Else {
	$script:osArch="System32"
	$script:ProgramFiles="Program Files"
	}
}
# ============================================================
# Script Body
# ============================================================

# Call ProcArch Function
ProcArch

# Define variables
$keyTool = "C:\$ProgramFiles\Java\jre8\bin\keytool.exe"
$cacerts = "C:\$ProgramFiles\java\jre8\lib\security\cacerts"
$signingCert = "\\Share\Location\CodeSignCert.cer"
$alias = "abcd-1234-efgh-5678-ijkl-90mnopqrstuv"

# Setup the process startup info
$processCheck = New-Object System.Diagnostics.ProcessStartInfo
$processCheck.FileName = "$keytool"
$processCheck.Arguments = "-list -keystore `"$cacerts`" -storepass changeit -alias `"$alias`" -noprompt"
$processCheck.UseShellExecute = $false
$processCheck.CreateNoWindow = $true
$processCheck.RedirectStandardOutput = $true

# Create a process object using the startup info
$process = New-Object System.Diagnostics.Process
$process.StartInfo = $processCheck

# Start the process
Write-Host "Checking for certificate..."
$process.Start() | Out-Null

# Get output from Standard Output
$stdout = $process.StandardOutput.ReadToEnd()

If ($stdout -match "does not exist") {
	# Installs Certificate
	Write-Host "Certificate not found; installing certificate..."
	Start-Process $keytool -ArgumentList "-importcert -keystore `"$cacerts`" -storepass changeit -file `"$signingCert`" -alias `"$alias`" -noprompt" -NoNewWindow -Wait
	Write-Host "Certificate has been installed on this system."
    }
Else {
    Write-Host "Certificate has already been installed on this system."
	}

# eos