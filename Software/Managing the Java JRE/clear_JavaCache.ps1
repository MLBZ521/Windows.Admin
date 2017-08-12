 <#

Script Name:  clear_JavaCache.ps1
By:  Zack Thompson / Created:  3/26/2015
Version:  1.0 / Updated:  3/26/2015 / By:  ZT

Description:  This script clears the users' Java Cache.

Note:  (* This has to be a user login script. *)

#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**                DO NOT CLOSE THIS WINDOW               **"
Write-Host "**                                                       **"
Write-Host "**                  Clearing Java Cache                  **"
Write-Host "**                                                       **"
Write-Host "**  This window will automatically close once complete.  **"
Write-Host "**                                                       **"
Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host

Write-Host "Checking OS Architecture..."
$cpuArch="AMD64"
If ($ENV:Processor_Architecture -eq $cpuArch) {
	$osArch="Program Files (x86)"
}
Else {
	$osArch="Program Files"
}

$JavaWS = "C:\$($osArch)\Java\jre8\bin\javaws.exe"

If (Test-Path -Path $JavaWS) {
    Write-Host "Clearing Java cache..."
	Start-Process $JavaWS -ArgumentList "-clearcache" -Wait
    Write-Host "Clearing Java applications..."
	Start-Process $JavaWS -ArgumentList "-uninstall" -Wait
}

Write-Host "Script complete!"

# eos