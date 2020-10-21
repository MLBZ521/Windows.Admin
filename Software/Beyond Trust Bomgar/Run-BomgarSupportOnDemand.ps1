<#

Script Name:  Invoke-BomgarSupportOnDemand.ps1
By:  Zack Thompson / Created:  5/11/2020
Version:  1.0.0 / Updated:  5/11/2020 / By:  ZT

Description:  Utilizing the Bomgar API, downloads a Bomgar Suppot Client and assigns it to the
              supplied team's queue based on the passed parameters.

Documentation:  https://www.beyondtrust.com/docs/remote-support/how-to/integrations/api/session-gen/index.htm

#>

# ============================================================
# Define Variables
# ============================================================

# Read in Registry parameters
# https://stackoverflow.com/a/19381092
$key = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
$subKey =  $key.OpenSubKey("SOFTWARE\Company\Bomgar")
$Group = $subKey.GetValue("IssueCodeName")
$Site = $subKey.GetValue("BomgarSiteURL")

# Set the OnDemand Support Group (Issue Code Name)
if ( $Group -eq $null ) {

    Write-Error "ERROR:  The IssueCodeName was not provided." -ErrorAction Stop

}

# Set the OnDemand Support Site URL
if ( $Site -eq $null ) {
    
    $Site = "bomgar.company.org"

}

# Set the value based on the OS Architecture
if ( $ENV:Processor_Architecture -eq "AMD64" ) {

    $platform = "winNT-64"

}
else {

    $platform = "winNT-32"

}

# Get the system details
$ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object Model, Manufacturer
$SerialNumber = Get-CimInstance -ClassName Win32_BIOS | Select-Object -ExpandProperty SerialNumber
$CustomerDetails = "$( ${ComputerSystem}.Manufacturer ), $( ${ComputerSystem}.Model ), ${SerialNumber}"

# Get Console User
$ConsoleUser = " ($env:USERNAME)"

# Get the Console Users' Full Name
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$DisplayName = [System.DirectoryServices.AccountManagement.UserPrincipal]::Current.DisplayName

# Set URL and clean it up
$URL = "https://${Site}/api/start_session?issue_menu=1&codeName=${Group}&platform=${platform}&customer.name=${DisplayName}${ConsoleUser}&customer.company=${CustomerDetails}"

# ============================================================
# Bits Staged...
# ============================================================

# Download the client
$response = Invoke-WebRequest -Method GET -Uri "${URL}" -UseBasicParsing
$filename = $response.Headers.'Content-Disposition' -replace '.*\bfilename=(.+)(?: |$)', '$1'
[IO.File]::WriteAllBytes("${env:TEMP}/${filename}", $response.Content)

# Run the executable
Start-Process "${env:TEMP}/${filename}" -Wait
