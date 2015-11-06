<#

Script Name:  CompareFolders.ps1
By:  Zack Thompson / Created:  7/16/2015
Version:  1.3 / Updated:  7/20/2015 / By:  ZT

Description:  This script compares files in folders (and more).

Borrowed and modified from:  helios456 @ http://serverfault.com/questions/532065/how-do-i-diff-two-folders-in-windows-powershell

#>

Write-Host ""
Write-Host "This script compares two folders and outputs the files that are either:"
Write-Host "	1) In Folder 1, but not in Folder 2"
Write-Host "	2) In both folders, but do not match."
Write-Host ""
Write-Host "The output is displayed to the screen as well as to a log file that can be provided."
Write-Host ""

$Folder1 = Read-Host "Enter Folder 1 (Old)"
$Folder2 = Read-Host "Enter Folder 2 (New)"
$OutFile = Read-Host "Enter File to log data"

Out-File $OutFile "Folder 1 Location:  $Folder1"
Out-File $OutFile "Folder 2 Location:  $Folder2"

# Get all files under $Folder1, filter out directories
$firstFolder = Get-ChildItem -Recurse $Folder1 | Where-Object { -not $_.PsIsContainer }

$failedCount = 0
$i = 0
$totalCount = $firstFolder.Count

$firstFolder | ForEach-Object {
    $i = $i + 1
    Write-Progress -Activity "Searching Files" -status "Searching File  $i of $totalCount" -percentComplete ($i / $firstFolder.Count * 100)
    
	# Check if the file, from $Folder1, exists with the same path under $Folder2
    If ( Test-Path ( $_.FullName.Replace($Folder1, $Folder2) ) ) {
        
		# Compare the contents of the two files...
        If ( Compare-Object (Get-Content $_.FullName) (Get-Content $_.FullName.Replace($Folder1, $Folder2) ) ) {
            
			# List the paths of the files containing diffs
            $fileSuffix = $_.FullName.TrimStart($Folder1)
            $failedCount = $failedCount + 1
            Write-Output "$fileSuffix is in each location, but does not match" | %{Write-Host $_; Out-File $OutFile -InputObject $_ -Append}
        }
    }
    else
    {
        $fileSuffix = $_.FullName.TrimStart($Folder1)
        $failedCount = $failedCount + 1
        Write-Output "$fileSuffix is only in folder 1" | %{Write-Host $_; Out-File $OutFile -InputObject $_ -Append}
    }
}

Write-Output "Comparing complete!" | %{Write-Host $_; Out-File $OutFile -InputObject $_ -Append}

<#  Commented this out as I do not need it at this time.

$secondFolder = Get-ChildItem -Recurse $Folder2 | Where-Object { -not $_.PsIsContainer }

$i = 0
$totalCount = $secondFolder.Count
$secondFolder | ForEach-Object {
    $i = $i + 1
    Write-Progress -Activity "Searching for files only on second folder" -status "Searching File  $i of $totalCount" -percentComplete ($i / $secondFolder.Count * 100)
    # Check if the file, from $Folder2, exists with the same path under $Folder1
    If (!(Test-Path($_.FullName.Replace($Folder2, $Folder1))))
    {
        $fileSuffix = $_.FullName.TrimStart($Folder2)
        $failedCount = $failedCount + 1
        Write-Host "$fileSuffix is only in folder 2"
    }
}
#>