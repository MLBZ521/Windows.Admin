<#

Script Name:  Report_Acrobat.ps1
By:  Zack Thompson / Created:  5/15/2015
Version:  2.0 / Updated:  9/22/2015 / By:  ZT

Description:  This script pulls the encrypted Acrobat product key from a 
    remote computer and then decrypts it.  It logs all information as is goes.

#>

# ============================================================
# Define Variables
# ============================================================
$WorkDir = "c:\Scripts\SoftwareChecks\Adobe\Acrobat\"
$AcrobatLog = "Log_AdobeAcrobat.txt"
$NoAcrobatLog = "Log_NoAacrobat.txt"
$OfflineLog = "Log_Offline.txt"
# ============================================================
# Function borrowed from:  https://gallery.technet.microsoft.com/scriptcenter/ConvertFrom-EncryptedAdobeK-1b1160e3
Function ConvertFrom-EncryptedAdobeKey {
    [CmdletBinding()]
    Param(
        [Parameter(Position=0, Mandatory=$true)] 
        [string]
        [ValidateLength(24,24)]
        $EncryptedKey
    )

    $AdobeCipher = "0000000001", "5038647192", "1456053789", "2604371895",
        "4753896210", "8145962073", "0319728564", "7901235846",
        "7901235846", "0319728564", "8145962073", "4753896210",
        "2604371895", "1426053789", "5038647192", "3267408951",
        "5038647192", "2604371895", "8145962073", "7901235846",
        "3267408951", "1426053789", "4753896210", "0319728564"

    $counter = 0

    $DecryptedKey = ""

    While ($counter -ne 24) {
        $DecryptedKey += $AdobeCipher[$counter].substring($EncryptedKey.SubString($counter, 1), 1)
        $counter ++
    }

    $DecryptedKey
}
# ============================================================
# Function to prompt admin for action.
Function PromptAdmin {
	$Title = "Choose Action";
	$Message = "Do you want to pull computers from Active Directory or from a List?"
	$AD = New-Object System.Management.Automation.Host.ChoiceDescription "&AD","Pull computer names from Active Directory.";
	$File = New-Object System.Management.Automation.Host.ChoiceDescription "&File","Pull computer names from a txt file.";
    $SinglePC = New-Object System.Management.Automation.Host.ChoiceDescription "&Single PC","Enter a single PC name.";
	$Options = [System.Management.Automation.Host.ChoiceDescription[]]($AD,$File,$SinglePC);
	$script:Answer = $Host.UI.PromptForChoice($Title,$Message,$Options,0)
}
# ============================================================
# Script Body
# ============================================================

# Check to see directory exists and create it if not.
If (!(Test-Path -Path $WorkDir)) {
	New-Item -ItemType directory -Path $WorkDir
}

# Set location the script will use to output it's files.
Set-Location $WorkDir

# Call Profile Function.
PromptAdmin

Write-host $Answer

If ($answer -eq 0) {
    Write-Host "Pulling computer names from Active Directory..."
    $list = Get-ADComputer -Filter { OperatingSystem -like "*Windows 7*" } | ForEach-Object { $_.Name }
}
ElseIf ($Answer -eq 1) {
    $FileLocation = Read-Host "Please enter file location"
    Write-Host "Pulling computer names from a txt file..."
    $list = Get-Content $FileLocation
}
ElseIf ($Answer -eq 2) {
    $ComputerName = Read-Host "Please enter the computer name"
    $list = $ComputerName
    Write-Host "Checking computer..."
}

ForEach ($entry in $list) {
    $Computer = $null
    $Version = $null
    $Serial = $null
	$Computer = $entry

	# Check if computer is accessible, if not go to next computer in list.
    Write-Host "Testing access to machine..."
	$Access = Test-Connection -ComputerName $Computer -Count 1
	If ($Access -eq $null) {
		"$Computer,Offline" | %{Write-Host $_; Out-File $OfflineLog -InputObject $_ -Append}
	}
	Else {
		# Pull the OS Architecture.
        Write-Host "Pulling the OS Architecture..."
		$OS = Get-WmiObject -ComputerName $Computer -Query "Select OSArchitecture from win32_OperatingSystem"

		# Set variables depending on OS Architecture.
		If ($OS.OSArchitecture -Match '64-bit') {
			$ProgramFiles = "Program Files (x86)"
            $64BitReg = "\Wow6432Node"
		}
		ElseIf ($OS.OSArchitecture -Match '32-bit') {
			$ProgramFiles = "Program Files"
		}

    	# Check to see what version of Acrobat is installed.
        Write-Host "Checking to see what version of Acrobat is installed..."
		$GetAcrobatVersion = Get-ChildItem "\\$Computer\C$\$ProgramFiles\Adobe\Acrobat*" -Recurse | Where {$_.name -eq "Acrobat.exe"} | Select versioninfo
	
		# If Acrobat is found, check the version
		If ($GetAcrobatVersion -ne $null) {
			$Version = $GetAcrobatVersion.versioninfo.productversion
            Write-Host "Found version $Version"
            # Split the Main version number so we can attempt to find the serial
            $VersionNumber = $Version -Split "\."
            $VersionBranch = $VersionNumber[0]

            If ($VersionBranch -eq 9) {
                # Pull the encrypted serial number from remote computer registry

                # Orginal line, but only works locally
                # $Encrypted = Get-ChildItem -Path "HKLM:\SOFTWARE$($64BitReg)\Adobe\Adobe Acrobat\$($VersionBranch).0\Registration" | Get-ItemProperty | Select-Object -Property Serial

                # This version uses a third party module, and would work, but trying to do this with built-in commands
                # Import-Module PSRemoteRegistry
                # Get-GPRegistryValue -ComputerName $Computer -Key "SOFTWARE$($64BitReg)\Adobe\Adobe Acrobat\$($VersionBranch).0\Registration" -ValueName Serial
                # http://psremoteregistry.codeplex.com/

                $Encrypted = Invoke-Command -ComputerName $Computer -ScriptBlock { param($64bit,$vBranch) Get-ChildItem -Path "HKLM:\SOFTWARE$($64Bit)\Adobe\Adobe Acrobat\$($VBranch).0\" | Get-ItemProperty | Select-Object -Property Serial } -ArgumentList $64BitReg,$VersionBranch
                $Convert = $Encrypted.Serial | Out-String

            }
            Else {
                # Pull the encrypted serial number from remote computer XML file.
                $XMLlocation = "\\$Computer\C$\ProgramData\regid.1986-12.com.adobe\"
                $XMLFile = Get-ChildItem -Path $XMLlocation | ForEach-Object { $_.FullName }
                [xml]$AdobeXML = Get-Content $XMLFile
                $Convert = $AdobeXML.software_identification_tag.serial_number
            }

            $Serial = ConvertFrom-EncryptedAdobeKey ($Convert.Trim())

            # Write to log file.
    		"$Computer,$Version,$Serial" | %{Write-Host $_; Out-File $AcrobatLog -InputObject $_ -Append}
		}

    	# If Acrobat is not found, write to log file
		ElseIf ($GetAcrobatVersion -eq $null) {
    		"$Computer,Acrobat not installed" | %{Write-Host $_; Out-File $NoAcrobatLog -InputObject $_ -Append}
		}
	}
}

# eos