<#

Script Name:  install_DellDockFirmware.ps1
By:  Zack Thompson / Created:  7/11/2016
Version:  1.0 / Updated:  7/11/2016 / By:  ZT

Description:  This script checks hardware characteristics that make an assumption 
	that a laptop is docked.  If it is, it will install the firmware update if applicable.
	
Note:  ( *This should be configured as a computer shutdown script.* )

#>

# ============================================================
# Define Variables

$LatestVersion = "3.10.001"
$Path = "HKLM:\SYSTEM\CurrentControlSet\Control\MST\Synaptics"
$Install = "C:\ProgramData\dell\drivers\Video_Firmware_FV9KP_WN32_3.10.1_A07\UpdateTool_x6_ver3_10_001.exe"

# ============================================================
# Script Body
# ============================================================

# Run the firmware update to get the current firmware version before installing.
Start-Process "cmd.exe" -ArgumentList "/c $Install -v -p2 -rp" -Wait -NoNewWindow

# Check if the Registry Path exists.
If (Test-Path $Path) {
	# Get the current version from the Registry Key Value.
	$CurrentVersion = Get-ItemProperty $Path | Select-Object "Orignial Version"

	# Compare version.
	If ($CurrentVersion.'Orignial version' -eq $LatestVersion) {
		# If machine is already up to date, then exit.
		Exit
	}
	Else {
		# Get the Power Status.
		$PowerStatus = (Get-WmiObject -Class BatteryStatus -Namespace root\wmi | Select-Object PowerOnline)

		# Check if the Laptop is plugged into AC power.
		If ($PowerStatus.PowerOnline -eq $True) {
			# Get the number of active monitors.
			$ActiveMonitors = Get-WmiObject -Namespace root\wmi -Class wmiMonitorID

			# Check if the number of monitors is more than one.
			If ($ActiveMonitors.Count -gt "1") {
				# If the laptop is connected to AC power and has more than one monitor, we will assume it is docked.

				# Install Firmware Update
				Start-Process "cmd.exe" -ArgumentList "/c $Install -p2 -rp" -Wait -NoNewWindow
			}
		}
	}
}

# eos