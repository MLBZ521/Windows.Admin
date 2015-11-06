<#

Script Name:  Install_Drivers.ps1
By:  Zack Thompson / Created:  7/16/2015
Version:  v2.0 / Updated:  8/20/2015 / By:  ZT

Description:  This script checks the model of the tower, checks driver
    versions and updates them if not the latest configured for it.

    (*This should be applied as a computer shutdown script.*)

#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**                    Checking Drivers                   **"
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
# This function checks the driver version, copies down the driver, and installs it.
Function DriverInstall {
    ForEach ($Device in $DriverVersion) {
		If ($Device.DriverVersion -ne $LatestVersion) {
			Write-Host "Installed version is not current; updating driver..."
			$serverLocation = "$Driver_Repo\" + $Model.Name + "\$Driver\$LatestVersion"
			$localLocation = "$IT_Staging\" + $Model.Name + "\$Driver\$LatestVersion"

			If ($osArch -eq "SysWow64") {
				$64bit = "64bit"
				If (!(Test-Path -Path "$localLocation\$64bit")) {
					New-Item -ItemType directory -Path "$localLocation\$64bit" | Out-Null
				}
				Copy-Item "$serverLocation\$64bit\*" "$localLocation\$64bit" -Recurse
				PnPUtil -i -a "$localLocation\$64bit\$Driverx64"
			}
			Else {
				$32bit = "32bit"
				If (!(Test-Path -Path "$localLocation\$32bit")) {
					New-Item -ItemType directory -Path "$localLocation\$32bit" | Out-Null
				}
				Copy-Item "$serverLocation\32bit\*" "$localLocation\$32bit" -Recurse
				PnPUtil -i -a "$localLocation\$32bit\$Driverx32"
			}
			Write-Host "Driver has been updated."
		}
		Else {
		Write-Host "Installed version is current."
		}
    }
}
# ============================================================
# Script Body
# ============================================================

# Call ProcArch Function
ProcArch

# ============================================================
# Define variables & locations
$Model = Get-WmiObject Win32_ComputerSystemProduct | Select-Object Name
$IT_Staging = "C:\Windows\$osArch\IT_Staging"
$Driver_Repo = "\\server\share\GPO Files\Driver_Repo"
# ============================================================

Write-Host "Computer Model: " $Model.Name

# OptiPlex 380
If ($Model.Name -eq "OptiPlex 380                 ") {
	# The value pulled in $Model is actually "OptiPlex 380                 " so this has to be trimmed to work in the flow of the script.
	$Model.Name = $Model.Name.Trim()

	# Update Graphics Driver
	Write-Host "Checking Graphics Driver version..."
	$DriverVersion = Get-WmiObject Win32_PnPSignedDriver | Select DeviceName, DriverVersion | Where {$_.DeviceName -like "*Intel*" -and $_.DeviceName -like "*G41*"}
	Write-Host "Installed version: " $DriverVersion[1].DriverVersion
	# There are two instances of this driver that are found, only want to worry about output for one of them.
	$Driver = "GFX"
	$LatestVersion = "8.15.10.2869"
	$Driverx64 = "d2869-64.inf"
	$Driverx32 = "kit49575.inf"
	
	# Call DriverInstall Function
	DriverInstall
	
    # Update NIC Driver
	Write-Host "Checking NIC Driver version..."
    $DriverVersion = Get-WmiObject Win32_PnPSignedDriver | Select DeviceName, DriverVersion | Where {$_.DeviceName -like "*Broadcom*" -and $_.DeviceName -like "*Ethernet*"}
    Write-Host "Installed version: " $DriverVersion.DriverVersion
	$Driver = "NIC"
	$LatestVersion = "15.6.0.10"
	$Driverx64 = "k57nd60a.inf"
	$Driverx32 = "k57nd60x.inf"
	
	# Call DriverInstall Function
	DriverInstall
	
}

# AMD Graphics Driver
If ($Model.Name -eq "OptiPlex 3020") {

# Need to check the Graphics Card in the 3010 before including this one
# If ($Model.Name -eq "OptiPlex 3020" -or $Model.Name -eq "OptiPlex 3010") {

	Write-Host "Checking if AMD Catalyst Software is installed..."

	# Check if AMD Catalyst Software is installed, and uninstall it (Prevent pop ups regarding driver updates).
    If (Test-Path "C:\Program Files\AMD\CIM\Bin64\ATISetup.exe") {
        Write-Host "CMD Catalyst is installed; removing software..."
        Start-Process "C:\Program Files\AMD\CIM\Bin64\ATISetup.exe" -ArgumentList "-Uninstall ALL"
		Wait-Process -name ATISetup
    }

	If (Test-Path "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{F2A7CE36-57BF-5C86-952D-90DBF3746D82}") {
		Start-Process "cmd.exe" -ArgumentList "/c msiexec /x{F2A7CE36-57BF-5C86-952D-90DBF3746D82} REBOOT=ReallySuppress"
		Wait-Process -name msiexec
	}

	If (Test-Path "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{426582A8-202F-D13C-8BD5-F00551BAFC93}") {
		Start-Process "cmd.exe" -ArgumentList "/c msiexec /X{426582A8-202F-D13C-8BD5-F00551BAFC93}"
		Wait-Process -name msiexec
	}

    # Update Graphics Driver
    Write-Host "Checking AMD Graphics Driver version..."
    $DriverVersion = Get-WmiObject Win32_PnPSignedDriver | Select DeviceName, DriverVersion | Where {$_.DeviceName -like "*AMD Radeon*" -or $_.DeviceName -like "*Standard VGA Graphics Adapter*"}
    Write-Host "Installed version: " $DriverVersion.DriverVersion
	$Driver = "AMD Graphics"
	$LatestVersion = "15.200.1046.0"
    $Driverx64 = "C7186188.inf"
	$Driverx32 = "CW186187.inf"
	
	# Call DriverInstall Function
	DriverInstall

}