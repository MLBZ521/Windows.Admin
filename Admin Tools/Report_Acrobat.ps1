<#

Script Name:  Report_Acrobat.ps1
By:  Zack Thompson / Created:  5/15/2015
Version:  1.0 / Updated:  5/18/2015 / By:  ZT

Description:  This script pulls Computer Objects from AD and then checks
	for Acrobat installations and makes logs along the way.

#>

# ============================================================
# Define Variables
# ============================================================
$WorkDir = "c:\Scripts\SoftwareChecks\Adobe\Acrobat\"
$AcrobatLog = "Log_AdobeAcrobat.txt"
$NoAcrobatLog = "Log_NoAacrobat.txt"
$OfflineLog = "Log_Offline.txt"

# ============================================================
# Check to see directory exists and create it if not.
If (!(Test-Path -Path $WorkDir)) {
	New-Item -ItemType directory -Path $WorkDir
}

# Set location the script will use to output it's files.
Set-Location $WorkDir

# ============================================================
# Script Body
# ============================================================

# Pull all Computer Objects from AD, but ignore ones that that have "mac" in their names.
$list = Get-ADComputer -Filter * | Where {$_.name -notlike "*mac*"} | select name,operatingsystem

ForEach ($entry in $list) {
	$Computer = $entry.name
	$CheckOS = $entry.operatingsystem
	
	# Check if the OS is Windows or OS X
	If ($CheckOS -ne "Mac OS X") {
	
		# Check if computer is accessible, if not go to next computer in list.
		$Access = Test-Connection -ComputerName $Computer -Count 1
		If ($Access -eq $null) {
			"$Computer,Offline" | %{Write-Host $_; Out-File $OfflineLog -InputObject $_ -Append}
		}
		Else {
			# Pull the OS Version and Architecture (I'm only worried about Windows 7 machines).
			$OS = Get-WmiObject -ComputerName $Computer -Query "Select Caption, OSArchitecture from win32_OperatingSystem"
	
			#Change program files path depending on OS.
			If ($OS.Caption -Match 'Windows 7') {
				If ($OS.OSArchitecture -Match '64-bit') {
					$ProgramFiles = "Program Files (x86)"
				}
				ElseIf ($OS.OSArchitecture -Match '32-bit') {
				$ProgramFiles = "Program Files"
				}
			}

			# Check to see what version of Acrobat is installed.
			$GetAcrobatVersion = Get-ChildItem "\\$Computer\C$\$ProgramFiles\Adobe\Acrobat*" -Recurse | Where {$_.name -eq "Acrobat.exe"} | Select versioninfo
	
			# If Acrobat is found, write to log file.
			If ($GetAcrobatVersion -ne $null) {
				$Version = $GetAcrobatVersion.versioninfo.productversion
				"$Computer,$Version" | %{Write-Host $_; Out-File $AcrobatLog -InputObject $_ -Append}
			}
			#If Acrobat is not found, write to log file
			ElseIf ($GetAcrobatVersion -eq $null) {
				"$Computer,Acrobat not installed" | %{Write-Host $_; Out-File $NoAcrobatLog -InputObject $_ -Append}
			}
		}
	}
}

#eos