<#

Script Name:  Install-BomgarJumpClient.ps1
By:  Zack Thompson / Created:  4/28/2020
Version:  1.0.1 / Updated:  5/13/2020 / By:  ZT

Description:  Installs a Bomgar Jump Client with the passed parameters

#>

param (
    [Parameter(Mandatory=$false, HelpMessage = "Action to perform.")][ValidateSet("Install", "Uninstall", "Repair", IgnoreCase = $true)][string]$Action = "Install"
)

<# If you want to use parameters, I'll leave this here...
param (
    [Parameter(Mandatory=$true, HelpMessage = "Sets the Jump Client Key.")][string]$Key,
    [Parameter(Mandatory=$true, HelpMessage = "Sets the Jump Client Group.  You must pass the Jump Group `"code_name`".")][string]$Group,
    [Parameter(Mandatory=$false, HelpMessage = "Sets the Jump Client Tag.")][string]$Tag,
    [Parameter(Mandatory=$false, HelpMessage = "Sets the Jump Client Comments.  default value:  <Manufacture>, <Model>, <Serial Number>")][string]$Comments,
    [Parameter(Mandatory=$false, HelpMessage = "Associates the Jump Client with the public portal which has the given hostname as a site address.  default value:  bomgar.company.org")][string]$Site = "bomgar.company.org",
    [Parameter(Mandatory=$false, HelpMessage = "Policy that controls the permission policy during a support session if the customer is present at the console.  You must pass the Policy's `"code_name`"")][string]$PolicyPresent,
    [Parameter(Mandatory=$false, HelpMessage = "Policy that controls the permission policy during a support session if the customer is not present at the console.  You must pass the Policy\'s `"code_name`".  default value:  Policy-Unattended-Jump")][string]$PolicyNotPresent = "Policy-Unattended-Jump"
)
#>

function Get-MSIInfo {
    param (
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo]$Path,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("ProductCode", "ProductVersion", "ProductName", "Manufacturer", "ProductLanguage", "FullVersion", "InstallPrerequisites")]
        [string]$Property
    )
    Process {
        try {
            # Read property from MSI database
            $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
            $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($Path.FullName, 0))
            $Query = "SELECT Value FROM Property WHERE Property = '$($Property)'"
            $View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
            $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
            $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
            $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
            # Commit database and close view
            $MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null)
            $View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null)
            $MSIDatabase = $null
            $View = $null
            # Return the value
            return $Value
        }
        catch {
            Write-Warning -Message $_.Exception.Message; break
        }
    }
    End {
        # Run garbage collection and release ComObject
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
        [System.GC]::Collect()
    }
}

function Install {

    # ============================================================
    # Read in Registry parameters
    # ============================================================

    $RegistryPath = "HKLM:\SOFTWARE\Company\Bomgar"
    $BomgarRegistryParameters = Get-ItemProperty -Path $RegistryPath

    $Key = $BomgarRegistryParameters | Select-Object -ExpandProperty JumpKey -ErrorAction SilentlyContinue
    $Group = $BomgarRegistryParameters | Select-Object -ExpandProperty JumpGroup -ErrorAction SilentlyContinue
    $Tag = $BomgarRegistryParameters | Select-Object -ExpandProperty JumpTag -ErrorAction SilentlyContinue
    $Comments = $BomgarRegistryParameters | Select-Object -ExpandProperty JumpComments -ErrorAction SilentlyContinue
    $Site = $BomgarRegistryParameters | Select-Object -ExpandProperty JumpSite -ErrorAction SilentlyContinue
    $PolicyPresent = $BomgarRegistryParameters | Select-Object -ExpandProperty JumpPolicyPresent -ErrorAction SilentlyContinue
    $PolicyNotPresent = $BomgarRegistryParameters | Select-Object -ExpandProperty JumpPolicyNotPresent -ErrorAction SilentlyContinue

    # ============================================================
    # Build install parameters
    # ============================================================

    # Set the Jump Client Key
    if ( $Key -ne $null ) {
        $JumpKey = "KEY_INFO=`"${Key}`""
    }
    else {
        Write-Error "ERROR:  A Jump Key was not provided." -ErrorAction Stop
    }

    # Set the Jump Client Group
    if ( $Group -ne $null ) {
        $JumpGroup = "jc_jump_group=`"jumpgroup:${Group}`""
    }
    else {
        $JumpGroup = ""
        # Write-Error "ERROR:  A Jump Group was not provided." -ErrorAction Stop
    }

    # Set the Jump Client Tag
    if ( $Tag -ne $null ) {
        $JumpTag = "jc_tag=`"${Tag}`""
    }
    else {
        $JumpTag = ""
    }

    # Set the Jump Client Comments
    if ( $Comments -eq $null ) {

        $ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object Model, Manufacturer
        $SerialNumber = Get-CimInstance -ClassName Win32_BIOS | Select-Object -ExpandProperty SerialNumber
        $JumpComments = "jc_comments=`"$( ${ComputerSystem}.Manufacturer ), $( ${ComputerSystem}.Model ), ${SerialNumber}`""

    }
    else {
        $JumpComments = "jc_comments=`"${Comments}`""
    }

    # Set the Jump Client Site
    if ( $Site -ne $null ) {
        $JumpSite = "jc_public_site_address=`"${Site}`""
    }
    else {
        $JumpSite = "jc_public_site_address=`"bomgar.company.org`""
    }

    # Set the Jump Client Console User Present Policy
    if ( $PolicyPresent -ne $null ) {
        $JumpPolicyPresent = "jc_session_policy_present=`"${PolicyPresent}`""
    }
    else {
        $JumpPolicyPresent = ""
    }

    Set the Jump Client Console User Not Present Policy
    if ( $PolicyNotPresent -ne $null ) {
        $JumpPolicyNotPresent = "jc_session_policy_not_present=`"${PolicyNotPresent}`""
    }
    else {
        $JumpPolicyNotPresent = "jc_session_policy_not_present=`"policy-unattended-jump`""
    }

    # Combine the Parameters into an array
    $Parameters = $JumpKey, $JumpGroup, $JumpSite, $JumpPolicyNotPresent, $JumpTag, $JumpComments, $JumpPolicyPresent

    # Combine the Parameters into a string
    $InstallParameters = ( $Parameters ).Where( { $_ } ) -join " "

    # ============================================================
    # Perform Install
    # ============================================================

    # Perform install
    $InstallProcess = Start-Process msiexec -ArgumentList "/i ${PSScriptRoot}\${MSI} ${InstallParameters} /quiet" -PassThru -Wait

    # Check the exit code
    if ( $InstallProcess.ExitCode -eq 0 ) {

        # Check if the install completed in less than five seconds; there's an issue that after an uninstall, it doesn't
        # successfully re-install the first time, but still exits with an exit code of 0.
        if ( $( $InstallProcess.ExitTime - $InstallProcess.StartTime ) -lt $( New-TimeSpan -Seconds 5 ) -and $InstallAttempt -eq 0 ) {

            $InstallAttempt = $InstallAttempt + 1

            # Run the install function again
            Install

        }

        # Add firewall rule to allow connections on port 5832 for Bomgar Passive Jump Client
        $GetRule = Get-NetFirewallRule -DisplayName "Bomgar Jump Client Passive" -ErrorAction SilentlyContinue

        if ( $GetRule -eq $null ) {

            New-NetFirewallRule -DisplayName "Bomgar Jump Client Passive" -Direction Inbound -Action Allow -Enabled True -Protocol tcp -LocalPort 5832

        }

        # ============================================================
        # Wait for the client to register with the appliance
        # ============================================================

        # Get current time
        $Start = Get-Date

        # Set the amount of time to wait
        $WaitFor = New-TimeSpan -Minutes 5

        do {

            # Get the Uninstall Command
            $KeyExists = Get-UninstallCode

            Start-Sleep -Seconds 1

            $TimedOut = $( ( Get-Date ) - $Start ) -gt $WaitFor

        } until ( $KeyExists -or $TimedOut )

        if ( $TimedOut -eq $true ) {

            Write-Error "`nERROR:  Doesn't appear the Bomgar Jump Client was able to register with the appliance.`nInstall Parameters:  ${InstallParameters}" -ErrorAction Stop

        }
        elseif ( $KeyExists ) {

            continue

        }

    }
    else {

        Write-Error "`nERROR:  Failed to install the Bomgar Jump Client.`nExit code:  $( ${InstallProcess}.ExitCode )`nInstall Parameters:  ${InstallParameters}" -ErrorAction Stop

    }

}

function Get-UninstallCode {

    # Get the MSI Product Code
    $ProductCode = Get-MSIInfo -Path ${PSScriptRoot}\${MSI} -Property "ProductCode"

    # Get the Registry Key Contents
    $RegistryPath = "HKLM:\SOFTWARE\Bomgar\JumpClientUninstall"
    $BomgarRegistryParameters = Get-ItemProperty -Path $RegistryPath

    # Get the Uninstall Command
    $UninstallCommand = $BomgarRegistryParameters | Select-Object -ExpandProperty "${ProductCode}".Trim()

    return $UninstallCommand

}

function Uninstall {

    $UninstallKey = Get-UninstallCode

    # Get the Install Directory
    $InstallDirectory = ( Get-Item $UninstallKey ).Directory.FullName

    # Perform uninstall
    $UninstallProcess = Start-Process "${InstallDirectory}\bomgar-scc.exe" -ArgumentList "-pinned win32uninstall silent" -PassThru -Wait

    # Check the exit code
    if ( $UninstallProcess.ExitCode -ne 0 ) {

        Write-Error "`nERROR:  Failed to install the Bomgar Jump Client.`nExit code:  $( ${UninstallProcess}.ExitCode )" -ErrorAction Stop

    }

}

# ============================================================
# Bits Staged...
# ============================================================

# Install attempt
$InstallAttempt=0

# Set the values based on the OS Architecture
if ( $ENV:Processor_Architecture -eq "AMD64" ) {

    $MSI = "bomgar-scc-win64.msi"

}
else {

    $MSI = "bomgar-scc-win32.msi"

}

switch ( "${Action}" ) {

    "Install" {
        Install
    }

    "Uninstall" {
        Uninstall
    }

    "Repair" {
        Uninstall
        Install
    }

}
