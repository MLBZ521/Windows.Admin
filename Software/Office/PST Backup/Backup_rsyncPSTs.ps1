<# 
Script Name:  Backup_rsyncPSTs.ps1
By:  Zack Thompson / Created:  1/2/2015
Version:  1.3 / Updated:  2/23/2015 / By:  ZT

Description:  This script uses rsync to incrementally backup local PST 
	files. Backup destination is the specified in a variable with exact 
	location configured on the server.
	(*This should be run as a scheduled task in the user evironment.*)
#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**                DO NOT CLOSE THIS WINDOW               **"
Write-Host "**                                                       **"
Write-Host "**                    Backing up PSTs                    **"
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
$user = $env:username
$computer = $env:computername
$IT_Staging = "C:\Windows\$osArch\IT_Staging\rsyncbin\"
$Rsync = "$($IT_Staging)rsync.exe"
$LocalPSTLoc = "/cygdrive/c/Users/$user/AppData/Local/Microsoft/Outlook/PST_Files/"
$logfile = "C:\Users\$user\AppData\Local\Microsoft\Outlook\PST_Files\Log_rsyncPST.txt"
$PSTLocation = "C:\Users\$user\AppData\Local\Microsoft\Outlook\PST_Files\"

# ============================================================
# Script Body
# ============================================================

# Check and see if there are any PSTs in the users' directory.
$AnyPSTs = Get-ChildItem *.pst -Path "$PSTLocation" -Recurse -ErrorAction SilentlyContinue

If ($AnyPSTs -eq $null) {
	# If there are not any PSTs in the directory, then exit the script.
	Exit
}
# Check to see if the RsyncBin location has been created before continuing.
ElseIf(Test-Path -Path $IT_Staging){
Write-Output "Script started on:  $(Get-Date -UFormat "%m-%d-%Y @ %r")" | Out-File ($logfile) -append

# Setup the process startup info
$processRsync = New-Object System.Diagnostics.ProcessStartInfo
$processRsync.FileName = "$Rsync"
$processRsync.Arguments = "-r --numeric-ids --delete-after --stats `"$LocalPSTLoc`" 172.16.100.60::Backup_Destination/`"$user`""
$processRsync.UseShellExecute = $false
$processRsync.CreateNoWindow = $true
$processRsync.RedirectStandardOutput = $true
$processRsync.RedirectStandardError = $true

# Create a process object using the startup info
$process = New-Object System.Diagnostics.Process
$process.StartInfo = $processRsync

# Start the process
$process.Start() | Out-Null

# Get output from stdout and stderr
$stdout = $process.StandardOutput.ReadToEnd()
$stderr = $process.StandardError.ReadToEnd()

If ($stdout.length -eq 0) {
	# If Stdout is empty, then an error occurred, so log this locally to log file.
	$LineFeed=[char]0x000A
	$Bang=[char]0x0021
	$ReplacedStderr = $stderr.Replace($LineFeed,$Bang)
	$separator = "!"
	$option = [System.StringSplitOptions]::None
	$ReplacedStderr.Split($separator,$option) | Out-File ($logfile) -append
	Write-Output "$( Get-Date -UFormat "%r |" ) Backup failed:" | Out-File ($logfile) -append
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
	Write-Output "$( Get-Date -UFormat "%r |" ) Backup was successful!" | Out-File ($logfile) -append
}

If ($stderr.length -ne 0) {
	# If Stderr is not empty, then an error occurred, so email information to the Helpdesk.
    $Return=[char]0x000A
	$MessageBody = "Backup ran on:  $(Get-Date -UFormat "%m-%d-%Y @ %r") $($Return) $($Return) For User:  $($user) $($Return) On Computer:  $computer $($Return) $($Return) $($stderr)"
	$FromAddress = "PSTBackups@IT.org"
	$ToAddress = "Helpdesk@IT.org"
	$MessageSubject = "PST Backup Error"
	$SendingServer = "bbcmailbox"
	$SMTPMessage = New-Object System.Net.Mail.MailMessage $FromAddress, $ToAddress, $MessageSubject, $MessageBody
	$SMTPClient = New-Object System.Net.Mail.SMTPClient $SendingServer
	$SMTPClient.Send($SMTPMessage)
}
}