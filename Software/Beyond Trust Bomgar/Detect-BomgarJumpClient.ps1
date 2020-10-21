<#

Script Name:  Detect-BomgarJumpClient.ps1
By:  Zack Thompson / Created:  5/1/2020
Version:  1.0.0 / Updated:  5/13/2020 / By:  ZT

Description:  Detects if the Bomgar Jump Client installed, running, and configured with the proper key.

#>

if ( Test-Path -Path "HKLM:\SOFTWARE\Bomgar\JumpClientUninstall" ) {

    # Get the Uninstall Command
    $UninstallCommand = Get-ItemProperty -Path "HKLM:\SOFTWARE\Bomgar\JumpClientUninstall" | Select-Object * -ExcludeProperty PS* | Select-Object -ExpandProperty * -ErrorAction SilentlyContinue

    if ( $UninstallCommand -ne $null ) {

        # Get the Install Directory
        $InstallDirectory = ( Get-Item $Uninstallcommand ).Directory.FullName

        # Get the Instance ID
        $InstanceID = ( Get-Item $InstallDirectory ).BaseName

        # Get the Service
        $BomgarService = Get-WmiObject win32_service | Where-Object { $_.PathName -match "${InstanceID}" } | Select-Object State

        # Get the Registry configuration
        $RegistryPath = "HKLM:\SOFTWARE\Company\Bomgar"
        $BomgarRegistryParameters = Get-ItemProperty -Path $RegistryPath

        # Get the Key
        $Key = $BomgarRegistryParameters | Select-Object -ExpandProperty JumpKey

        # Verify the Install Directory Exists and the Service is running and the instance matches the configured Key
        if ( ( Test-Path $InstallDirectory ) -eq $true -and $BomgarService.State -eq "Running" -and ( Test-Path "${InstallDirectory}\scc-${Key}.exe" ) -eq $true ) {
            $success=$True
        }
        else {
            $success=$False
            # Write-Error "ERROR:  Failed to locate the Bomgar Jump Client install path." -ErrorAction Stop
            # Write-Error "ERROR:  The Bomgar Jump Client service is not running." -ErrorAction Stop
            # Write-Error "ERROR:  Current Bomgar Jump Client instance does not match the configured instance" -ErrorAction Stop
        }

        # Report Success
        if ( $success -eq $true ) {
            Write-Host "Successfully installed the Bomar Jump Client"
        }

    }

}