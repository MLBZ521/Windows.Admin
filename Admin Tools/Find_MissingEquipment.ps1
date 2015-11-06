<#
Script Name:  MissingEquip.ps1
By:  Zack Thompson / Created:  2/25/2015
Version:  1.0 / Updated:  2/25/2015 / By:  ZT

Description:  This script checks the serial number of the local machine 
	and if the machine's serial number matches one of the missing serial
	numbers in the array, it will send an email to the specified address.

	(*This should be run as a user logon script.*)
	
#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**                DO NOT CLOSE THIS WINDOW               **"
Write-Host "**                                                       **"
Write-Host "**                Missing Equipment Search               **"
Write-Host "**                                                       **"
Write-Host "**  This window will automatically close once complete.  **"
Write-Host "**                                                       **"
Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host

# Enter Serial Numbers below that you are looking for.
$Looking = @()
$Looking += "123ABCD"
$Looking += "456EFGH"

$User = $env:username
$Computer = $env:computername
$SerialNumber = gwmi win32_bios | Select-Object SerialNumber

ForEach ($SerialNumber in $Looking) {
    If ($SerialNumber.SerialNumber -eq $SerialNumber) {
        Write-Host "You have a missing piece of equipment!"

        $Return=[char]0x000A
        $MessageBody = "Notification Time:  $(Get-Date -UFormat "%m-%d-%Y @ %r") $($Return) $($Return) Found Service Tag:  $($SerialNumber) $($Return) Logged on User:  $($User) $($Return) Computer Name:  $($Computer)"
        $FromAddress = "SerialNumberSearch@domain.org"
        $ToAddress = "Helpdesk@domain.org"
        $MessageSubject = "Found Missing Equipment!"
        $SendingServer = "ExchangeHostOrIP"
        $SMTPMessage = New-Object System.Net.Mail.MailMessage $FromAddress, $ToAddress, $MessageSubject, $MessageBody
        $SMTPClient = New-Object System.Net.Mail.SMTPClient $SendingServer
        $SMTPClient.Send($SMTPMessage)
    }
}