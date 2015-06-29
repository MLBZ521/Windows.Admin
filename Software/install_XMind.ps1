<#

Script Name:  Install_XMind.ps1
By:  Zack Thompson / Created:  9/18/2014
Version:  1.3 / Updated:  9/18/2014 / By:  ZT

Description:  This script installs XMind. (*This has to be a user login script.*)

#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**					  Installing XMind					 **"
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
}
Else {
$script:osArch="System32"
}
}
# ============================================================
# Check to see directory exists and create it if not.
Function DirCheck {
if(!(Test-Path -Path $localLocation)){
    New-Item -ItemType directory -Path $localLocation
}
}
# ============================================================
# Check to see log-file exists.
Function LogCheck{
if(!(Test-Path -Path $localLocation\Log_XMind.txt)){
$script:noLog="NotInstalled"
}
}
# ============================================================
# Copy files to local disk and execute.
Function CopyExecute {
cd $localLocation
Copy-Item $serverLocation\* $localLocation
Start-Process "xmind-windows-3.4.1.201401221918.exe" -ArgumentList "/norestart /loadinf=xmind_install-settings.inf" -Wait
}
# ============================================================
# If 32bit, delete Bonjour installer files to prevent XMind
# from asking end user to install on program launch, and make
# changes to registry.
Function arch32Do {
Remove-Item "C:\Program Files\XMind\thirdparty\Bonjour.msi" -force
Remove-Item "C:\Program Files\XMind\thirdparty\Bonjour64.msi" -force
Set-Location -Path Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\XMind_is1
Set-ItemProperty . DisplayName "XMind 2013 (v3.4.1) - $Org"
}
# ============================================================
# If 64bit, delete Bonjour installer files to prevent XMind
# from asking end user to install on program launch, and make
# changes to registry.
Function arch64Do {
Remove-Item "C:\Program Files (x86)\XMind\thirdparty\Bonjour.msi" -force
Remove-Item "C:\Program Files (x86)\XMind\thirdparty\Bonjour64.msi" -force
Set-Location -Path Registry::HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\XMind_is1
Set-ItemProperty . DisplayName "XMind 2013 (v3.4.1) - $Org"
}
# ============================================================
# Create log-file and append install date to file.
Function CreateLog {
Write-Output "XMind installed on: " | Out-File Log_XMind.txt
Get-Date | Out-File Log_XMind.txt -append	
}
# ============================================================
# Move preference files into place (this has to be done per user).
Function MovePrefs {
cd $localLocation
$prefDIR="$env:AppData\XMind\workspace-cathy\.metadata\.plugins\org.eclipse.core.runtime\.settings\"
if(!(Test-Path -Path $prefDIR)){
    New-Item -ItemType directory -Path $prefDIR
}
Copy-Item net.xmind.verify.prefs $prefDIR
Copy-Item org.xmind.cathy.prefs $prefDIR
}
# ============================================================

# ============================================================
# Script Body
# ============================================================

# Call ProcArch Function
ProcArch
Write-Host "OS Architecture is:  $osArch"

# Define variables & locations
$newDir="XMind"
$serverLocation="\\bbcfile\Installs\GPO Files\Software Installs\XMind"
$IT_Staging="C:\Windows\$osArch\IT_Staging"
$localLocation="$IT_Staging\$newDir"
$Org=""

# Call DirCheck Function
DirCheck
# Call LogCheck Function
LogCheck

If ($noLog -eq "NotInstalled") {
Write-Host "XMind is not Installed on this PC"
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
Write-Host "XMind is currently installed on this PC."
}

# Call MovePrefs Function
MovePrefs
Write-Host "Moved preference files into place for current user."

Write-Host "Script completed successfully!"
# eos