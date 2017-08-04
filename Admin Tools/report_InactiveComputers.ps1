<#

Script Name:  report_InactiveComputers.ps1
By:  Zack Thompson / Created:  11/15/2016
Version:  1.7 / Updated:  12/8/2016 / By:  ZT

Description:  This script finds inactive computer accounts in AD and pulls relevant information.
    This information can then be displayed to screen, exported to a csv, or the accounts can be disabled.

Notes:
*BitLocker Key retrieval process borrowed from:  https://ndswanson.wordpress.com/2014/10/20/get-bitlocker-recovery-from-active-directory-with-powershell/

#>

# ============================================================
# Define Variables
# ============================================================

$Name = Read-Host "Enter the OU Name to search"
$OUName = Get-ADOrganizationalUnit -Filter 'Name -like $Name' | Select-Object DistinguishedName
$OUDSName = $OUName.DistinguishedName
$TimeSpan = Read-Host "Enter the TimeSpan to search"

# ============================================================
# Functions
# ============================================================

# Function to prompt user for a choice.
Function Action {
	$Title = "Choose Action";
	$Message = "Select option to perform.  Enter ? for more information on the options."
	$Display = New-Object System.Management.Automation.Host.ChoiceDescription "&Display","This option will display all information to the screen.";
	$Export = New-Object System.Management.Automation.Host.ChoiceDescription "&Export","This option will export all information to a CSV file.";
    $Disable = New-Object System.Management.Automation.Host.ChoiceDescription "D&isable","This option will disable the inactive computer objects.";
	$Exit = New-Object System.Management.Automation.Host.ChoiceDescription "E&xit","This option will quit the script.";
    $Options = [System.Management.Automation.Host.ChoiceDescription[]]($Display,$Export,$Disable,$Exit);
	$script:Answer = $Host.UI.PromptForChoice($Title,$Message,$Options,0)
}

# Function to perform the AD Search
Function SearchAD {
    # Get all inactive computers
    $InActiveComputers = Search-ADAccount -AccountInactive -ComputersOnly -TimeSpan "$($TimeSpan)" -SearchBase "$($OUDSName)"

    Write-Host "There were $($InActiveComputers.count) computer accounts found in the search."

    ForEach ($Computer in $InActiveComputers) {

        # Get information for each computer account.
        $ComputerProperties = Get-ADComputer -Identity $Computer.Name -Properties * | Select-Object Name, Enabled, Created, LastLogonDate, Description, DistinguishedName

        # Get the latest BitLocker Recovery Key for the object.
       $BitLockerKey = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $ComputerProperties.DistinguishedName -Properties 'msFVE-RecoveryPassword' | Select-Object msFVE-RecoveryPassword | Select-Object -Last 1

        # Add all desired information into an array.
        $Properties = @{
            Name = $ComputerProperties.Name
            Enabled = $ComputerProperties.Enabled
            Created = $ComputerProperties.Created
            LastLogonDate = $ComputerProperties.LastLogonDate
            Description = $ComputerProperties.Description
            DistinguishedName = $ComputerProperties.DistinguishedName
           msFVERecoveryPassword = $BitLockerKey.'msFVE-RecoveryPassword'
        }

        # Convert array into an object and then at object to an array that will be exported.
        $script:ExportArray += New-Object PSObject -Property $Properties
    }
}

# Function to disable Inactive Computers
Function DisableComputers {
    # Get all inactive computers.
    $DisableObjects = Search-ADAccount -AccountInactive -ComputersOnly -TimeSpan "$($TimeSpan)" -SearchBase "$($OUDSName)" | Select-Object Name, DistinguishedName, Enabled | Where-Object {$_.Enabled -eq "True"}
    
    Write-Host "There are $($DisableObjects.count) computer accounts that need to be disabled."
    
    If ($DisableObjects.count -gt 0) {
        Write-host "Disabling inactive computers..."
        ForEach ($DisableObject in $DisableObjects) {
            Disable-ADAccount -Identity $DisableObject.DistinguishedName
            $Date = Get-Date -Format g
            Get-ADComputer -Identity $DisableObject.DistinguishedName -Properties Description | ForEach-Object { Set-ADComputer -Identity $_.DistinguishedName -Description ($_.Description + " *** Disabled by $($env:USERNAME) on $($Date)") }
        }
    }
    Else {
        Write-Host "All objects are already disabled."
    }
}

# ============================================================
# Script Body
# ============================================================

Do {

    # Function Action
    Action
    
    $ExportArray = @()
    
    If ($Answer -eq 0) {
        Write-host "Searching AD for inactive computers..."

        # Function SearchAD
        SearchAD

        # Display data to standard out.
        $ExportArray | Select-Object Name, Enabled, Created, LastLogonDate, Description, msFVERecoveryPassword, DistinguishedName | Format-Table -AutoSize
        $ExportArray = $null
    }
    ElseIf ($Answer -eq 1) {
        $OutFileLocation = Read-Host "Enter the the folder path to save a csv file"
        $OutFile = "$($OutFileLocation)\Report_$($Name)_InActiveComputers.csv"
    
        Write-host "Searching AD for inactive computers..."
        # Function SearchAD
        SearchAD

        # Export data to a file.
        $ExportArray | Select-Object Name, Enabled, Created, LastLogonDate, Description, msFVERecoveryPassword, DistinguishedName | Export-Csv -Path $OutFile -Append
        $ExportArray = $null
    }
    ElseIf ($Answer -eq 2) {
        Write-host "Searching AD for inactive computers..."
        # Function DisableComputers
        DisableComputers
        $ExportArray = $null
    }
    ElseIf ($Answer -eq 3) {
        Write-Host "Script Complete!"
    }
}

While ( $Answer -ne 3 )

# eos