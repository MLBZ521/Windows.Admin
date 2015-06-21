<#
Script Name:  Move_PSTs.ps1
By:  Zack Thompson / Created:  12/29/2014
Version:  1.4 / Updated:  3/24/2015 / By:  ZT

Description:  This script first checks for PST files that are attached to 
	Outlook, moves the attached PSTs to the specified destination, and then
	reattaches the previously attached PST files to Outlook.  Then searches 
	the user's profile directories for PST files and moves them to the 
	specified destination.  It also accounts for duplicate PST file names.
	(*This has to be a user login script.*)
#>

Write-Host "***********************************************************"
Write-Host "***********************************************************"
Write-Host "**                                                       **"
Write-Host "**                DO NOT CLOSE THIS WINDOW               **"
Write-Host "**                                                       **"
Write-Host "**                    PST Maintenance                    **"
Write-Host "**                                                       **"
Write-Host "**  This window will automatically close once complete.  **"
Write-Host "**                                                       **"
Write-Host "***********************************************************"
Write-Host "***********************************************************"

# ============================================================
# This function is called when there are no PST files found attached to Outlook or when they are already in the correct location.
Function NoPSTFiles {
    Write-Output "$( Get-Date -UFormat "%r |" ) Either there were no attached PST files found for $($user) or attached PSTs were in the correct location." | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
	
	# Close Outlook & Release ComObject references that are open.
	$Outlook.Quit()
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($NameSpace)){}
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook)){}
	[System.GC]::Collect()
	Start-Sleep -s 6
	Stop-Process -name Outlook | Out-Null
	Wait-Process -name Outlook

	# Call ProfilePSTs Function.
	ProfilePSTs
}
# ============================================================
# Look for PST files that are not attached to Outlook.
Function ProfilePSTs {
	Write-Host "$( Get-Date -UFormat "%r |" ) Searching for PST Files in $($user)'s profile..."
    $userProfile = $env:userprofile
	$ProfilePSTs = Get-ChildItem *.pst -Path "$userProfile" -Recurse -ErrorAction SilentlyContinue | Select-Object DirectoryName,Name,BaseName,CreationTime
    
	# Check to see if there were any items found in the above directory search.
	If ($ProfilePSTs -ne $null) {
		Write-Output "$( Get-Date -UFormat "%r |" ) Found and moved the following PSTs in the user's Profile that were not attached to Outlook:" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
	    ForEach ($ProfilePST in $ProfilePSTs) { 
			Write-Output "$( Get-Date -UFormat "%r |" )  Original Location:  $($ProfilePST.DirectoryName)\$($ProfilePST.Name)" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
			If (!(Test-Path "$($Destination)$($ProfilePST.Name)")) {
				Move-Item "$($ProfilePST.DirectoryName)\$($ProfilePST.name)" -Destination $Destination -ErrorVariable ErrorMove1
				Write-Output "$( Get-Date -UFormat "%r |" )   $($ErrorMove1) Moved to:  $($Destination)$($ProfilePST.Name)" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
			}
			Else {
				$Created = $ProfilePST.CreationTime.toString('_dd-MM-yyyy')
				Move-Item "$($ProfilePST.DirectoryName)\$($ProfilePST.name)" -Destination "$($Destination)$($ProfilePST.BaseName)$($Created).pst" -ErrorVariable ErrorMove2
				Write-Output "$( Get-Date -UFormat "%r |" )   $($ErrorMove2) Moved to:  $($Destination)$($ProfilePST.BaseName)$($Created).pst" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
				}
		}
	}
	Else {
	    Write-Output "$( Get-Date -UFormat "%r |" ) There were no PST files found in $($user)'s profile." | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
	}
		Write-Output "$( Get-Date -UFormat "%r |" ) Script complete! $($Return)" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}

	# eos
}
# ============================================================
# Define Variables
# ============================================================

# Get currently logged on user.
$user = $env:username

# Define location to the new destination for PST Files
$Destination = "C:\Users\" + $user + "\AppData\Local\Microsoft\Outlook\PST_Files\"

# Set LogFile Location and name
$LogFile = $Destination + "Log_MovedPSTs.txt"

# Get date for Log File Entries.
$date = Get-Date -UFormat "%m-%d-%Y %r"

$Return = [char]0x000A

# ============================================================
# Script Body
# ============================================================

# Check to see if Mitel process is running and stop it if so.
$Check = Get-Process -Name Mitel.PIM.ServiceHost -ErrorAction SilentlyContinue
If ($Check -ne $null) {
	Stop-Process -Name Mitel.PIM.ServiceHost
}

# Check to see directory exists and create it if not.
if(!(Test-Path -Path $Destination)){
    New-Item -ItemType directory -Path $Destination
}

# Write to log that script started to process.
Write-Output "Script ran on $($date)" | Out-File $LogFile -append

# Open Outlook.
$Outlook = New-Object -ComObject Outlook.Application
$NameSpace = $Outlook.getNamespace("MAPI")

# Check to see if there is an Outlook Profile, if not, quit processing script as they *shouldn't* have PST files on this computer then.
$ProfileExists = $Outlook.Application.DefaultProfileName

If ($ProfileExists -eq $null) {
	Write-Output "$( Get-Date -UFormat "%r |" ) $($user) does not have an Outlook Profile on this PC at this time." | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
	Write-Output "$( Get-Date -UFormat "%r |" ) Script complete!" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}

	# Close Outlook & Release ComObject references that are open.
    $Outlook.Quit()
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($NameSpace)){}
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook)){}
	[System.GC]::Collect()
	Start-Sleep -s 6
	Stop-Process -name Outlook | Out-Null
    Wait-Process -name Outlook

	# eos
}
Else {
	Write-Output "$( Get-Date -UFormat "%r |" ) $($user)'s local default Outlook Profile is $($ProfileExists)" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}

	# Pull PST information from Outlook.
	$Item = $NameSpace.stores | where {$_.ExchangeStoreType -eq 3} | Select-Object DisplayName,FilePath

	# Check to see if there any attached PST files.
	If ($Item -ne $null) {
		
		# Define empty arrays.
		$arrayNames = @()
		$arrayPaths = @()
		$arrayNewPaths = @()
		
		# Insert PST info into arrays, but if PST is already in the correct location, do not add to array.
		ForEach ($pstFile in $item) {
			$Original = Get-ChildItem $pstFile.Filepath | Select-Object DirectoryName,Name
			If ("$($Original.DirectoryName)\" -ne $Destination) {
				$arrayNames += $pstFile.DisplayName
				$arrayPaths += $pstFile.FilePath
			}
		}
		
		# Check to see if there are any PSTs added to the array that are not in the correct location, if not skip to the next action.
		If ($arrayNames.count -ne 0) {
		Write-Output "$( Get-Date -UFormat "%r |" ) Found the following PSTs attached to Outlook:" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
		$PSTCount = $arrayNames.Count
		$countPSTs = 0
		While ($countPSTs -ne $PSTCount) {
			$aPath = $arrayPaths[$countPSTs]
			$aName = $arrayNames[$countPSTs]
			$OriginalPST = Get-ChildItem $aPath | Select-Object DirectoryName,Name
			Write-Output "$( Get-Date -UFormat "%r |" )  $($OriginalPST.DirectoryName)\$($OriginalPST.Name)" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
			$countPSTs = ($countPSTs + 1)
		}
		
		# Remove PST Files from Outlook.
		Write-Output "$( Get-Date -UFormat "%r |" ) Removing the following PSTs from Outlook:" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
		ForEach ($pstName in $arrayNames) {
			$pstDisplayName = "$pstName"
			$pst = $NameSpace.Stores | Where {$_.DisplayName -eq $pstName}
			$pstRoot = $pst.GetRootFolder()
			$pstRoot.Name = $pstDisplayName
			$pstFolder = $NameSpace.Folders.Item($pstDisplayName)
			
			$NameSpace.GetType().InvokeMember('RemoveStore',[System.Reflection.BindingFlags]::InvokeMethod,$null,$Namespace,($pstFolder))
			Write-Output "$( Get-Date -UFormat "%r |" )  $($pstName)" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
		}

		# Close Outlook & Release ComObject references that are open.
		$Outlook.Quit()
		while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($pstFolder)){}
		while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($pstRoot)){}
		while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($pst)){}
		while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($NameSpace)){}
		while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook)){}
		[System.GC]::Collect()
		Start-Sleep -s 15
		Stop-Process -name Outlook | Out-Null
		Wait-Process -name Outlook
  
		# Move PST Files to New Location.
		Write-Output "$( Get-Date -UFormat "%r |" ) Moved the following PSTs:" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
		$PSTNumber = $arrayPaths.Count
		$countPST = 0
		While ($countPST -ne $PSTNumber) {
			$path = $arrayPaths[$countPST]
			$name = $arrayNames[$countPST]
			$Original = Get-ChildItem $path | Select-Object DirectoryName,Name,Basename
        
			# Check to see if a PST file of the same name already exists
			If (!(Test-Path "$($Destination)$($Original.Name)")) {
				Move-Item -Path  $path -Destination $Destination -ErrorVariable ErrorMove3
				Write-Output "$( Get-Date -UFormat "%r |" )  $($ErrorMove3) $($Destination)$($Original.Name)" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
				$arrayNewPaths += $Destination + $Original.Name
			}
			Else {
				Move-Item -Path  $path -Destination "$($Destination)$($Original.BaseName)$($countPST).pst" -ErrorVariable ErrorMove4
				Write-Output "$( Get-Date -UFormat "%r |" )  $($ErrorMove4) $($Destination)$($Original.BaseName)$($countPST).pst" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
				$arrayNames[$countPST] = "$($Original.BaseName)$($countPST)"
				$arrayNewPaths += $Destination + $Original.BaseName + $countPST + ".pst"
			}
			$countPST = ($countPST + 1)
		}

		# Open Outlook again.
		$Outlook = New-Object -ComObject Outlook.Application
		$NameSpace = $Outlook.getNamespace("MAPI")

		# Add PST Files back to Outlook.
		Write-Output "$( Get-Date -UFormat "%r |" ) Attaching the following PSTs back to Outlook:" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
		ForEach ($pstItem in $arrayNewPaths) {
			# Check to see if the PST is in the new location before adding it back; this prevents creating a new, empty PST with the same name.
			If (Test-Path "$($pstItem)") {
				$NameSpace.AddStore($pstItem)
				Write-Output "$( Get-Date -UFormat "%r |" )  $($pstItem)" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
			}
			Else {
				Write-Output "$( Get-Date -UFormat "%r |" )  It appears the following PST may have been in use and unable to be moved:  $($pstItem)" | %{Write-Host $_; Out-File $LogFile -InputObject $_ -Append}
			}
		}

		# Close Outlook & Release ComObject references that are open.
		$Outlook.Quit()
		while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($NameSpace)){}
		while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook)){}
		[System.GC]::Collect()
		Start-Sleep -s 6
		Stop-Process -name Outlook | Out-Null
		Wait-Process -name Outlook

		# Call ProfilePSTs Function.
		ProfilePSTs

	}
	Else {
		# Call NoPSTFiles Function.
		NoPSTFiles
	}
	}
	Else {
		# Call NoPSTFiles Function.
		NoPSTFiles
	}
}