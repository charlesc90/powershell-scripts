########################################################################################################
# Script: MoveFile-MatchName.ps1
# Description: This script moves files based on matching a filename
########################################################################################################

Function MoveFile-MatchName {

	# This is the source directory
	$sourceDir = "C:\test"
	# This is the destination directory
	$destinationDir = "C:\test2"

	# Gets the contents of the source directory recursively, including hidden files, excluding directories
	$fileList = Get-ChildItem -Path $sourceDir -File -Force -Recurse
	
	# loop through $fileList
	Foreach ($file in $fileList) {
		If($file.BaseName -Match 'test-01') {
			#$matches.values gives just the values that match based on -Match flag
			$fileName = $matches.Values
			# Store full path to file in $fileToCheck variable
			$fileToCheck ="$destinationDir\$file"
			#Check if File Exists and if it does, print error
			If (Test-Path $fileToCheck -Pathtype Leaf) {
				Write-Warning "File $file.Name already exists at $destinationDir"
			}
			#If file does not exist then move the file to the destination directory               
			Else {
				Move-Item -Path $($file.FullName) -Destination $destinationDir
			}
		}
		Else {
			$null
		}
	}
}
MoveFile-MatchName
