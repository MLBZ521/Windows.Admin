<# 
Script Name:  Restore_EagleSoft.ps1
By:  Zack Thompson / Created:  10/23/2015
Version:  1.2 / Updated:  10/28/2015 / By:  ZT

Description:  This script uses rsync to incrementally restore the Eaglesoft
	backup files from the external USB drive to the local Data directory.
	
#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**                DO NOT CLOSE THIS WINDOW               **"
Write-Host "**                                                       **"
Write-Host "**                Restoring Eaglesoft Data               **"
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

# Get the Drive Letter of the External Drive (it MUST be named "Backup"!)
$BackupDrive = Get-WmiObject win32_LogicalDisk | Where { $_.VolumeName -eq "Backup" } | Select-Object DeviceID

$IT_Staging = "C:\Windows\$osArch\IT_Staging\rsyncbin\"
$EaglesoftDataDir =  "C:\EagleSoft\Data"
$PattersonTechAid = "C:\EagleSoft\Shared Files\techaid.exe"
$PattersonServer = "C:\EagleSoft\Shared Files\PattersonServerStatus.exe"
$Eaglesoft = "C:\EagleSoft\Shared Files\Eaglesoft.exe"
$Rsync = "$($IT_Staging)rsync.exe"
$DataDir = "/cygdrive/c/Eaglesoft/Data"
$BackupDir = "/cygdrive/$($BackupDrive.DeviceID[0])/Eaglesoft/Data/"
$WinBackupDir = "$($BackupDrive.DeviceID)\Eaglesoft\Data"
$logfile = "$($BackupDrive.DeviceID)\Logs\RestoreLog.txt"

# ============================================================
# Script Body
# ============================================================

# Check to see if the BackupDir exists, if not, create it.
If (!(Test-Path -Path $WinBackupDir)) {
New-Item -ItemType directory -Path $WinBackupDir
}

If (!(Test-Path -Path "$($BackupDrive.DeviceID)\Logs")) {
New-Item -ItemType directory -Path "$($BackupDrive.DeviceID)\Logs"
}

# Check to see if the RsyncBin location exists before continuing.
If (Test-Path -Path $IT_Staging) {
	Write-Output "Script started on:  $(Get-Date -UFormat "%m-%d-%Y @ %r")" | Out-File ($logfile) -append

	Write-Output "Please note:  A full restore may take up to two hours."
	Write-Output "Incremental restores should take significantly less time."
	
	# Stop EagleSoft Services
	Write-Output "Stopping EagleSoft Services..."
	Start-Process $PattersonTechAid -ArgumentList "-stop" -NoNewWindow -Wait
	# Start-Process $PattersonServer -ArgumentList "-stop" -NoNewWindow -Wait	
	
	# Setup the process startup info
	$processRsync = New-Object System.Diagnostics.ProcessStartInfo
	$processRsync.FileName = "$Rsync"
	$processRsync.Arguments = "-rltgoDh --stats `"$BackupDir`" `"$DataDir`""
	$processRsync.UseShellExecute = $false
	$processRsync.CreateNoWindow = $true
	$processRsync.RedirectStandardOutput = $true
	$processRsync.RedirectStandardError = $true

	# Create a process object using the startup info
	$process = New-Object System.Diagnostics.Process
	$process.StartInfo = $processRsync

	Write-Output "Restoring data..."
	Write-Output "$( Get-Date -UFormat "%r |" ) Restore started..." | Out-File ($logfile) -append

	# Start the process
	$process.Start() | Out-Null

	# Get output from stdout and stderr
	$stdout = $process.StandardOutput.ReadToEnd()
	$stderr = $process.StandardError.ReadToEnd()

	Write-Output "$( Get-Date -UFormat "%r |" ) Restore process completed..." | Out-File ($logfile) -append
	
	If ($stdout.length -eq 0) {
		# If Stdout is empty, then an error occurred, so log this locally to log file.
		$LineFeed=[char]0x000A
		$Bang=[char]0x0021
		$ReplacedStderr = $stderr.Replace($LineFeed,$Bang)
		$separator = "!"
		$option = [System.StringSplitOptions]::None
		$ReplacedStderr.Split($separator,$option) | Out-File ($logfile) -append
		Write-Output "$( Get-Date -UFormat "%r |" ) Restore failed:" | Out-File ($logfile) -append
		Write-Output "Please try the restore again, if it continues to fail, please contact the IT Department."
		}
	Else {
		# Else, if Stdout is not empty, then backup was successful, log this locally to log file.
		$stdoutTrimmed=$stdout.TrimStart()
		$LineFeed=[char]0x000A
		$Bang=[char]0x0021
		$ReplacedStdout = $stdoutTrimmed.Replace($LineFeed,$Bang)
		$separator = "!"
		$option = [System.StringSplitOptions]::None
		$ReplacedStdout.Split($separator,$option) | Out-File ($logfile) -append
		Write-Output "$( Get-Date -UFormat "%r |" ) Restore was successful!" | Out-File ($logfile) -append
		
		Write-Output "$( Get-Date -UFormat "%r |" ) Updating file permissions..." | Out-File ($logfile) -append
		
		$ErrorActionPreference = "SilentlyContinue"
		$SetPermissions = New-Object System.Security.AccessControl.FileSystemAccessRule("officeserver\dental","FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
		
		$ACL = (Get-Item $EaglesoftDataDir).GetAccessControl("Access")
		$ACL.SetAccessRuleProtection($false,$false)
		$ACL.Access | ForEach-Object { $ACL.RemoveAccessRule($_) } | Out-Null
		$ACL.SetAccessRule($SetPermissions)
		Set-ACL $EaglesoftDataDir $ACL

		$Items = Get-ChildItem -Path $EaglesoftDataDir -Recurse -ErrorAction SilentlyContinue | Select-Object Fullname

		ForEach ($Item in $Items) {
			$ACL = (Get-Item $Item.Fullname).GetAccessControl("Access")
			$ACL.Access | %{$ACL.RemoveAccessRule($_)} | Out-Null
			$ACL.Access | %{$ACL.RemoveAccessRule($_)} | Out-Null			
			$ACL.SetAccessRuleProtection($false,$false)
			Set-ACL $Item.Fullname $ACL
		}

		$ErrorActionPreference = "Continue"
				
		Write-Output "$( Get-Date -UFormat "%r |" ) File permissions updated!" | Out-File ($logfile) -append
		
		# Start EagleSoft Services
		Write-Output "Starting EagleSoft Services..."
		Start-Process $PattersonServer -ArgumentList "-start" -NoNewWindow -Wait
		Start-Sleep -s 10

		# Start EagleSoft
		Write-Output "Starting EagleSoft..."
		Start-Process $EagleSoft

		Write-Output "Script completed successfully!"
		}
	Write-Output "Press any key to exit."
	$Pause = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")
}