########################################################################################################
# Script: MoveFile-MatchName.ps1
# Author: Charles Cox
# Date: 26 December 2019
# Description: This script moves files based on matching a filename
########################################################################################################

# This is the source directory
$srcRoot = "C:\Users\Charles.Cox\Documents\powershell-scripts-master\powershell-scripts-master\test"
# This is the destination directory
$dstRoot = "C:\Users\Charles.Cox\Documents\powershell-scripts-master\powershell-scripts-master\test2"

# Get-ChildItem: Gets the contents of the source directory recursively
# -File: Get list of files
# -Force: Allows the cmdlet to get items that cannot otherwise not be accessed
$fileList = Get-ChildItem -Path $srcRoot -File -Force -Recurse

# loop through $fileList
Foreach ($file in $fileList) {
   # try to find matches
   # Powershell conditional operators
   # -Match
   # -Like
   # -Contains
   if($file.BaseName -Match 'test-01') {
		#$matches.values gives just the values that match based on -Match flag
        $fileName = $matches.Values
        # Store full path to file in $fileToCheck variable
        $fileToCheck ="$dstDir\$file"
        #Check if File Exists and if it does, print error
        if (Test-Path $fileToCheck -Pathtype Leaf) {
			Write-Warning "File $file.Name already exists at $dstDir"
        }
        #If file does not exist then move the file to the destination directory               
        else {
			Move-Item -Path $($file.FullName) -Destination $dstDir
        }
    }
   else {
        $null
   }
}
