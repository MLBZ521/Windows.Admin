<#
Script Name:  MovePSTs.ps1
By:  Zack Thompson / Created:  12/29/2014
Version:  0.6 / Updated:  1/12/2015 / By:  ZT

Description:  This script checks for PST files that are attached to Outlook, 
	then searches the user's profile directories for PST files, and then 
	moves them to the specified destination and reattaches the previously 
	attached PST files.  It also accounts for duplicate names and changes 
	the file name of the PST to that of the name of the PST in Outlook.
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
# Define location to the new destination for PST Files.
# ============================================================
$Destination = "C:\Users\" + $user + "\AppData\Local\Microsoft\Outlook\PST_Files\"

# ============================================================
# Look for PST files that are not attached to Outlook.
Function ProfilePSTs {
    $userProfile = $env:userprofile
	$count = 1
    Write-Output "$( Get-Date -UFormat "%r |" )  Found and moved the following PSTs in the user's Profile that were not attached to Outlook:" | Out-File ($Destination + "Log_MovedPSTs.txt") -append
    dir *.pst -Path "$userProfile" -Recurse -ErrorAction SilentlyContinue | Where { $_.Fullname -ne $null } | 
        foreach-object { 
			$destPSTExists = "$($Destination)$($_.name)"
			If (!(Test-Path $destPSTExists)) {
				Move-Item $_.Fullname -Destination $Destination
				Write-Output "$( Get-Date -UFormat "%r |" )  $($_.Fullname)" | Out-File ($Destination + "Log_MovedPSTs.txt") -append
			}
			Else {
				$dupPSTName = $_.Fullname
				$work = $_.Basename
				Move-Item $dupPSTName -Destination "$($Destination)$($work)$($count).pst"
				Write-Output "$( Get-Date -UFormat "%r |" )  $($Destination)$($work)$($count).pst" | Out-File ($Destination + "Log_MovedPSTs.txt") -append
				$count = $count + 1
			}
		}

		Write-Output "$( Get-Date -UFormat "%r |" )  Script complete!" | Out-File ($Destination + "Log_MovedPSTs.txt") -append

	# eos
}
# ============================================================
# Script Body
# ============================================================

# Get currently logged on user.
$user = $env:username

# Get date for Log File Entries.
$date = Get-Date -UFormat "%m-%d-%Y %r"

# Check to see directory exists and create it if not.
if(!(Test-Path -Path $Destination)){
    New-Item -ItemType directory -Path $Destination
}

# Write to log that script started to process.
Write-Output "Script ran on $($date)" | Out-File ($Destination + "Log_MovedPSTs.txt") -append

# Open Outlook.
$Outlook = New-Object -comObject Outlook.Application
$NameSpace = $Outlook.getNamespace("MAPI")

# Check to see if there is an Outlook Profile, if not, quit processing script.
$ProfileExists = $Outlook.Application.DefaultProfileName

If ($ProfileExists -eq $null) {
	Write-Output "$( Get-Date -UFormat "%r |" )  $($user) does not have an Outlook Profile on this PC at this time." | Out-File ($Destination + "Log_MovedPSTs.txt") -append
	Write-Output "$( Get-Date -UFormat "%r |" )  Script complete!" | Out-File ($Destination + "Log_MovedPSTs.txt") -append

	# Close Outlook & Release ComObject references that are open.
    $Outlook.Quit()
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($NameSpace)){}
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook)){}
	[System.GC]::Collect()
    Wait-Process -name Outlook

# eos
}
Else {
Write-Output "$( Get-Date -UFormat "%r |" )  $($user)'s local default Outlook Profile is $($ProfileExists)" | Out-File ($Destination + "Log_MovedPSTs.txt") -append

# Pull PST information from Outlook.
$item = $NameSpace.stores | where {$_.ExchangeStoreType -eq 3} | select-object DisplayName,FilePath

# Define empty arrays.
$arrayNames = @()
$arrayPaths = @()

# Insert PST info into arrays + if PST is already in the correct location, do not add to array.
foreach ($pstFile in $item | Where { $_.FilePath -ne $null } ) {


# **************Check the following line, is "+ $pstfile.DisplayName + ".pst"" need?  Or use $pstFile.Basename or .Filename?
    If ($pstFile.FilePath -ne $Destination + $pstFile.DisplayName + ".pst") {
        Write-Output "$( Get-Date -UFormat "%r |" )  Found PST named: $($pstFile.DisplayName) at location: $($pstFile.FilePath)" | Out-File ($Destination + "Log_MovedPSTs.txt") -append
    	$arrayNames += $pstFile.DisplayName
    	$arrayPaths += $pstFile.FilePath
    }
}

# Check to see if there are zero attached PST files.
If ($arrayPaths.Count -lt 1) {
    Write-Output "$( Get-Date -UFormat "%r |" )  Either there were no attached PST files found for $($user) or attached PSTs were in the correct location." | Out-File ($Destination + "Log_MovedPSTs.txt") -append

    # Close Outlook & Release ComObject references that are open.
    $Outlook.Quit()
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($NameSpace)){}
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook)){}
	[System.GC]::Collect()
    Wait-Process -name Outlook
    
    # Call ProfilePSTs Function.
    ProfilePSTs
}
Else {
	# Remove PST Files from Outlook.
	Write-Output "$( Get-Date -UFormat "%r |" )  Removing the following PSTs from Outlook:" | Out-File ($Destination + "Log_MovedPSTs.txt") -append
	ForEach ($pstName in $arrayNames) {
		$pstDisplayName = "$pstName"
		$pst = $NameSpace.Stores | ? {$_.DisplayName -eq $pstName}
		$pstRoot = $pst.GetRootFolder()
		$pstRoot.Name = $pstDisplayName
		$pstFolder = $Namespace.Folders.Item($pstDisplayName)

		$NameSpace.GetType().InvokeMember('RemoveStore',[System.Reflection.BindingFlags]::InvokeMethod,$null,$Namespace,($pstFolder))
		Write-Output "$( Get-Date -UFormat "%r |" )  $($pstName)" | Out-File ($Destination + "Log_MovedPSTs.txt") -append
	}

	# Close Outlook & Release ComObject references that are open.
	$Outlook.Quit()
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($pstFolder)){}
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($pstRoot)){}
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($pst)){}
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($NameSpace)){}
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook)){}
	[System.GC]::Collect()
    Wait-Process -name Outlook
  
	# Move PST Files to New Location.
	Write-Output "$( Get-Date -UFormat "%r |" )  Moved the following PSTs:" | Out-File ($Destination + "Log_MovedPSTs.txt") -append
	$PSTNumber = $arrayPaths.Length
	$countPST = 0
	While ($countPST -ne $PSTNumber) {
		$path = $arrayPaths[$countPST]
		$name = $arrayNames[$countPST]
		
		# **************Check the following line, is "$($Destination)$($name).pst" needed?  If Outlook does know the internal Outlook name of the PST, then the rename isn't needed.
		# Change to $($Destintation)$($path.Filname) or $($path.Basename).pst
				$PSTExists = "$($Destination)$($name).pst"
		
		
		# Check to see if a PST file of the same name already exists
		If (!(Test-Path $PSTExists)) {
		
		
		# **************Check the following line, is "$($Destination)$($name).pst" needed?  If Outlook does know the internal Outlook name of the PST, then the rename isn't needed.
		# Change to just $Destintation then along with all the Logs lines that follow.
		
			Move-Item -Path  $path -Destination "$($Destination)$($name).pst" | out-null
			Write-Output "$( Get-Date -UFormat "%r |" )  $($Destination)$($name).pst" | Out-File ($Destination + "Log_MovedPSTs.txt") -append
		}
		Else {
			Move-Item -Path  $path -Destination "$($Destination)$($name)$($countPST).pst" | out-null
			Write-Output "$( Get-Date -UFormat "%r |" )  $($Destination)$($name)$($countPST).pst" | Out-File ($Destination + "Log_MovedPSTs.txt") -append
			$arrayNames[$countPST] = "$($name)$($countPST)"
		}	
    $countPST = ($countPST + 1)
	}

	# Open Outlook again.
	$Outlook = New-Object -comObject Outlook.Application
	$NameSpace = $Outlook.getNamespace("MAPI")

	# Add PST Files back to Outlook.
	Write-Output "$( Get-Date -UFormat "%r |" )  Attaching the following PSTs back to Outlook:" | Out-File ($Destination + "Log_MovedPSTs.txt") -append
	ForEach ($pstItem in $arrayNames) {
	
		# **************Check the above ForEach statement, could just need the $Destination $item.ArrayNames.FileName/Basename option
		
		$pstFileName = $pstItem + ".pst"
		$NameSpace.AddStore("$Destination$pstFileName")
		Write-Output "$( Get-Date -UFormat "%r |" )  $($Destination)$($pstFileName)" | Out-File ($Destination + "Log_MovedPSTs.txt") -append
	}

	# Close Outlook & Release ComObject references that are open.
	$Outlook.Quit()
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($NameSpace)){}
	while([System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook)){}
	[System.GC]::Collect()
	Wait-Process -name Outlook

	# Call ProfilePSTs Function.
	ProfilePSTs
}
}
