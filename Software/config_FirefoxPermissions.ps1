<#
Script Name:  install_FirefoxPermissions.ps1
By:  Zack Thompson / Created:  3/3/2015
Version:  1.2 / Updated:  3/4/2015 / By:  ZT

Description:  This script uses sqlite3.exe (CLI utility) to check and add
    exceptions or permissions to allow a specified website to run Java 
    and Flash plugins.
	(*This has to be applied as a USER script.*)
	
	***If any more permissions are ever needed to be added, this check 
	processes should be converted to a loop.***
#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**                DO NOT CLOSE THIS WINDOW               **"
Write-Host "**                                                       **"
Write-Host "**           Adding Plugin Permissions to Firefox        **"
Write-Host "**                                                       **"
Write-Host "**  This window will automatically close once complete.  **"
Write-Host "**                                                       **"
Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host

# ============================================================
# Check Processor Architecture (32bit vs 64bit).
$cpuArch="AMD64"
If($ENV:Processor_Architecture -eq $cpuArch){
	$script:osArch="SysWow64"
}
Else {
	$script:osArch="System32"
}
# ============================================================
# Define variables & locations
# ============================================================

# Get currently logged on user.
	$user = $env:username
# Define location to the new destination for PST Files.
	$FirefoxPath = "C:\Users\" + $user + "\AppData\Roaming\Mozilla\Firefox\Profiles\"
# Get the logged on users Firefox Profile; ErrorAction Stop script if profile doesn't exist.
	$FirefoxProfile = Get-ChildItem -Path $FirefoxPath -ErrorAction Stop | Select Name
# Set the local of the permissions.sqlite file.
	$PermFile = $FirefoxPath + $FirefoxProfile.Name + "\permissions.sqlite"
# Locations for Staging Folder and the sqlite3 CLI application
	$IT_Staging = "C:\Windows\$osArch\IT_Staging\"
	$Sqlite = "$($IT_Staging)sqlite3.exe"
# Define SQL Strings
	$JavaString = "select host from moz_hosts where host='web.site.com' AND type='plugin:java' AND permission='1' AND expireType='0';"
	$FlashString = "select host from moz_hosts where host='web.site.com' AND type='plugin:flash' AND permission='1' AND expireType='0';"
	$JavaAdd = "insert into moz_hosts(host, type, permission, expireType) values('web.site.com', 'plugin:java', 1, 0);"
	$FlashAdd = "insert into moz_hosts(host, type, permission, expireType) values('web.site.com', 'plugin:flash', 1, 0);"

# ============================================================
# Script Body
# ============================================================

# Java Check Process
# Setup the process startup info
	$JavaCheck = New-Object System.Diagnostics.ProcessStartInfo
	$JavaCheck.FileName = "$Sqlite"
	$JavaCheck.Arguments = "`"$PermFile`" `"$JavaString`""
	$JavaCheck.UseShellExecute = $false
	$JavaCheck.CreateNoWindow = $true
	$JavaCheck.RedirectStandardOutput = $true
	$JavaCheck.RedirectStandardError = $true
# Create a process object using the startup info
	$ProcessJava = New-Object System.Diagnostics.Process
	$ProcessJava.StartInfo = $JavaCheck
# Start the process
	$ProcessJava.Start() | Out-Null
# Get output from stdout and stderr
	$JavaStdOut = $ProcessJava.StandardOutput.ReadToEnd()
	$JavaStdErr = $ProcessJava.StandardError.ReadToEnd()

# Flash Check Process
# Setup the process startup info
	$FlashCheck = New-Object System.Diagnostics.ProcessStartInfo
	$FlashCheck.FileName = "$Sqlite"
	$FlashCheck.Arguments = "`"$PermFile`" `"$FlashString`""
	$FlashCheck.UseShellExecute = $false
	$FlashCheck.CreateNoWindow = $true
	$FlashCheck.RedirectStandardOutput = $true
	$FlashCheck.RedirectStandardError = $true
# Create a process object using the startup info
	$ProcessFlash = New-Object System.Diagnostics.Process
	$ProcessFlash.StartInfo = $FlashCheck
# Start the process
	$ProcessFlash.Start() | Out-Null
# Get output from stdout and stderr
	$FlashStdOut = $ProcessFlash.StandardOutput.ReadToEnd()
	$FlashStdErr = $ProcessFlash.StandardError.ReadToEnd()

# Check to see if Java is already added to this permissions.sqlite file, if not add it.
	If ($JavaStdOut.length -eq 0) {
		Start-Process $Sqlite -ArgumentList "`"$PermFile`" `"$JavaAdd`"" -NoNewWindow -Wait
	}

# Check to see if Flash is already added to this permissions.sqlite file, if not add it.
	If ($FlashStdOut.length -eq 0) {
		Start-Process $Sqlite -ArgumentList "`"$PermFile`" `"$FlashAdd`"" -NoNewWindow -Wait
	}