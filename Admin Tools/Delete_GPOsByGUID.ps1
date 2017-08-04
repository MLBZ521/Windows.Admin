<#

Script Name:  Delete_GPOsByGUID.ps1
By:  Zack Thompson | Created:  10/19/2016
Version:  1.1 | Updated:  10/19/2016 | By:  ZT

Description:  This script deletes Group Policy Objects based on their 
    GUID and display information for each before deletion.

Syntax:  To use this script, you would call it with the domain and file you want to use.

Examples:  Delete_GPOsByGUID.ps1 -Domain my.domain.com -File GPO-GUIDS.txt

#>

# ============================================================
# Define variables

Param([string]$Domain,[string]$File)
$GPOGuids = Get-Content $File

# ============================================================
# Script Body
# ============================================================

ForEach ($GUID in $GPOGuids) {
    $GPOInfo = Get-GPO -Domain $Domain -Guid $GUID | Select-Object DisplayName, Owner, ID

    If ($GPOInfo -ne $null) {
        Write-Host "Deleting object $($GPOInfo.DisplayName) ($($GPOInfo.ID)) created by $($GPOInfo.Owner)."

        Remove-GPO -Domain $Domain -Guid $GUID -WhatIf
    }
    Else {
        Write-host "Unable to locate:  $($GUID)"
    }
}

# eos