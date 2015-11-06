<#

Script Name:  Install_Drivers.ps1
By:  Zack Thompson / Created:  7/16/2015
Version:  v1.0 / Updated:  7/17/2015 / By:  ZT

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
# Script Body
# ============================================================

# Call ProcArch Function
ProcArch

# Define variables & locations
$Model = Get-WmiObject Win32_ComputerSystemProduct | Select-Object Name
$IT_Staging = "C:\Windows\$osArch\IT_Staging"
$Driver_Repo = "\\server\share\GPO Files\Driver_Repo"

Write-Host "Computer Model: " $Model.Name

# OptiPlex 380
If ($Model.Name -eq "OptiPlex 380                 ") {
    
    # The value pulled in $Model is actually "OptiPlex 380                 " so this has to be trimmed to work in the flow of the script.
    $Model.Name = $Model.Name.Trim()

    # Update Graphics Driver
    Write-Host "Checking Graphics Driver version..."
    $DriverVersion = Get-WmiObject Win32_PnPSignedDriver | Select DeviceName, DriverVersion | Where {$_.DeviceName -like "*Intel*" -and $_.DeviceName -like "*G41*"}
    $LatestVersion = "8.15.10.2869"
    Write-Host "Installed version: " $DriverVersion[1].DriverVersion
    # There are two instances of this driver that are found, only want to worry about output for one of them.
    
    If ($DriverVersion[1].DriverVersion -EQ $LatestVersion) {
        Write-Host "Installed version is not current; updating driver..."
        $Driver = "GFX"
        $serverLocation = "$Driver_Repo\" + $Model.Name + "\$Driver\$LatestVersion"
        $localLocation = "$IT_Staging\" + $Model.Name + "\$Driver\$LatestVersion"

        If ($osArch -eq "SysWow64") {
            Copy-Item "$serverLocation\64bit" -Recurse "$localLocation\64bit"
            $Driverx64 = "$localLocation\64bit\d2869-64.inf"
            PnPUtil -i -a $Driverx64
        }
        Else {
            Copy-Item "$serverLocation\32bit" -Recurse "$localLocation\32bit"
            $Driverx32 = "$localLocation\32bit\kit49575.inf"
            PnPUtil -i -a $Driverx32
        }
        Write-Host "Graphics Driver has been updated."
    }
    Else {
    Write-Host "Installed version is current."
    }
    }

    # Update NIC Driver
    Write-Host "Checking NIC Driver version..."
    $DriverVersion = Get-WmiObject Win32_PnPSignedDriver | Select DeviceName, DriverVersion | Where {$_.DeviceName -like "*Broadcom*" -and $_.DeviceName -like "*Ethernet*"}
    $LatestVersion = "15.6.0.10"
    Write-Host "Installed version: " $DriverVersion.DriverVersion

    If ($DriverVersion.DriverVersion -ne $LatestVersion) {
        Write-Host "Installed version is not current; updating driver..."
        $Driver = "NIC"
        $serverLocation = "$Driver_Repo\" + $Model.Name + "\$Driver\$LatestVersion"
        $localLocation = "$IT_Staging\" + $Model.Name + "\$Driver\$LatestVersion"

        If ($osArch -eq "SysWow64") {
            Copy-Item "$serverLocation\64bit" -Recurse "$localLocation\64bit"
            $Driverx64 = "$localLocation\k57nd60a.inf"
            PnPUtil -i -a $Driverx64
        }