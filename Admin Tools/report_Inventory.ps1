<#

Script Name:  report_Inventory.ps1
By:  Zack Thompson / Created:  5/23/2016
Version:  1.0 / Updated:  5/23/2016 / By:  ZT

Description:  This script pulls information for inventory purposes and 
	then saves it to a csv file which can be imported into Excel.

#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**                       Inventory                       **"
Write-Host "**                                                       **"
Write-Host "***********************************************************"
Write-Host "***********************************************************"

# ============================================================
# Define Variables
$DriveLetter = Read-Host "Please enter the drive letter of your USB Drive"
$InventoryFile = "$($DriveLetter):\Inventory.csv"

# ============================================================
# Set variables
$Location = Read-Host "Please enter the location for this equipment"
$Desk = Read-Host "Please enter the desk number for this equipment"
$ComputerPCN = Read-Host "Please enter the Computer's Proptery Control Number"

$ComputerName = $ENV:ComputerName
$User = $env:USERNAME
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$FirstName = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity([System.DirectoryServices.AccountManagement.ContextType]::Domain,(whoami)).GivenName
$LastName = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity([System.DirectoryServices.AccountManagement.ContextType]::Domain,(whoami)).Surname
$Fullname = Get-WMIObject Win32_NetworkLoginProfile | Select Fullname
$ServiceTag = Get-WmiObject -Class Win32_BIOS | Select-Object SerialNumber
$SystemInfo = Get-WmiObject -Class Win32_ComputerSystem | Select-Object Manufacturer,Model
$ProcessorModel = Get-WmiObject -Class Win32_Processor | Select-Object Name
$SystemRAM = Get-WMIObject -Class Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum | Select-Object @{N="Total Physical Ram"; E={[Math]::Round(($_.Sum / 1GB),2)}}

# ============================================================
# Function for gathering monitor info. 
Function Get-MonitorInfo {
	$ActiveMonitors = Get-WmiObject -Namespace root\wmi -Class wmiMonitorID
	$MonitorInfo = @()

	ForEach ($Monitor in $ActiveMonitors) {
		If ($Monitor.UserFriendlyName -ne $null) {
			$Display = $null
			$Display = New-Object PSObject -Property @{
				UserFriendlyName = ($monitor.UserFriendlyName | % {[char]$_}) -join ''
				SerialNumberID = ($monitor.SerialNumberID | % {[char]$_}) -join ''
			}
			$MonitorInfo += $Display
		}
	}

	ForEach ($Info in $MonitorInfo) {
	Write-Output $Info.UserFriendlyName.Replace(" ",',')
	Write-Output $Info.SerialNumberID
	$MonitorPCN = Read-Host "Please enter the Monitor's Proptery Control Number"
	"$FirstName,$LastName,$User,Monitor,$Location,$Desk,$($Info.UserFriendlyName.Replace(" ",',')),$($Info.SerialNumberID),$MonitorPCN"  | Out-File $InventoryFile -Append
	}
}

# ============================================================
# Display to host console information gathered.
Write-Output "Computer Name is:  $ComputerName"
Write-Output "User's full name is:  $($FirstName) $($LastName)"
Write-Output "Username is:  $User"
Write-Output "Serial Number is:  $($ServiceTag.SerialNumber)"
Write-Output "OEM is:  $($SystemInfo.Manufacturer)"
Write-Output "Model is:  $($SystemInfo.Model)"
Write-Output "Procossor is:  $($ProcessorModel.Name)"
Write-Output "RAM is:  $($SystemRAM.'Total Physical Ram')GB"

# ============================================================
# Save information to the defined csv file above.
"$FirstName,$LastName,$User,Computer,$Location,$Desk,$($SystemInfo.Manufacturer),$($SystemInfo.Model),$($ServiceTag.SerialNumber),$ComputerPCN,,,$($SystemRAM.'Total Physical Ram')GB,$($ProcessorModel.Name),$ComputerName" | Out-File $InventoryFile -Append

# Call Get-MonitorInfo Function
Get-MonitorInfo

# eos