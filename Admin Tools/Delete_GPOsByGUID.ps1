<#

Script Name:  Delete_GPOsByGUID.ps1
By:  Zack Thompson | Created:  10/19/2016
Version:  1.0 | Updated:  10/19/2016 | By:  ZT

Description:  This script deletes Group Policy Objects based on their 
    GUID and display information for each before deletation.

Note:  

#>

# ============================================================
# Define variables
$ADdomain = "AD.domain.name"

$GPOGuids = @()
$GPOGuids += "4f243dc8-4dcc-4629-9145-94b3123c6160"
$GPOGuids += "a96f9ec3-1d7e-44fb-b0f2-0084e426e5d2"
$GPOGuids += "fe37c125-ef80-437d-b16e-fdf763fc335c"
$GPOGuids += "c21076b6-b7ea-40a0-9615-58499f88e759"
$GPOGuids += "f7a372b3-189b-435e-a712-199703ad0b39"
$GPOGuids += "fb754087-47ef-4d4f-af49-0297a9e2b829"
$GPOGuids += "d69e771b-8da0-4c69-9774-3dd0f999a850"
$GPOGuids += "1777136f-069a-4aa8-a390-d4862925434b"
$GPOGuids += "d5552ed2-a627-4919-ab30-ebd6acee43f9"
$GPOGuids += "70b16935-f2da-41bf-90bb-bedc7952e309"
$GPOGuids += "1b190192-8e19-4ead-b285-20431d4003c2"
$GPOGuids += "ceaf8d61-596b-4720-82fe-e49b942b0f49"
$GPOGuids += "f5478959-1184-4ab0-bdd8-4e689b90c407"
$GPOGuids += "92824760-c951-4d16-8b0d-c2e4183c8c4e"
$GPOGuids += "f24b248c-f246-40b6-bb00-c3cb1075ceb1"
$GPOGuids += "7aad9c3b-b404-49db-8246-12af2f68e191"
$GPOGuids += "e265676b-e84e-45e1-b100-de9ee4efa1c9"
$GPOGuids += "e3fb4879-c9f5-4ef2-aaf9-ad9890575075"
$GPOGuids += "6adda39a-0e3c-4d16-bb03-dd4bfd84c4a2"
$GPOGuids += "0b242372-5cb0-4f9a-a0ad-317cbce5522f"
$GPOGuids += "07086c57-3858-43ef-ab55-8cb5b2ec6efc"
$GPOGuids += "0d1dcf67-cb7a-445b-8c22-7d04018cab37"
$GPOGuids += "e97c0c17-f155-43f0-a84d-cee51a124d3b"
$GPOGuids += "ab03ab7c-7426-4743-a0ac-17cbe667ae19"
$GPOGuids += "68217a7b-c3c9-4f17-9077-95569311d76d"
$GPOGuids += "2ad799db-8f01-401a-bbe3-602477b43e03"

# ============================================================
# Script Body
# ============================================================

ForEach ($GUID in $GPOGuids) {

    $GPOInfo = Get-GPO -Domain "$ADdomain" -Guid $GUID | Select-Object DisplayName, Owner, ID

    Write-Host "Deleteing object $($GPOInfo.DisplayName) ($($GPOInfo.ID)) created by $($GPOInfo.Owner)."

    Remove-GPO -Domain "$ADdomain" -Guid $GUID

}

# eos
