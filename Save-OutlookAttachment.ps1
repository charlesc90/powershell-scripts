########################################################################################################
# Script: Save-OutlookAttachment.ps1
# Author: Charles Cox
# Date: 20 December 2019
# Description: This script saves outlook attachments of a specified file type to the given destination
########################################################################################################

Function Save-OutlookAttachment {

    # Destination filepath
    $path = "C:\Users\Charles.Cox\Desktop"

    # Add .NET core class for outlook
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" | Out-Null
    $olFolders = “Microsoft.Office.Interop.Outlook.olDefaultFolders” -as [type]
    # Define outlook COM object and namespace
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNameSpace("MAPI") 
    # Select the outlook folder to query. If you neet to select a different folder, change ::olFolder...
    $folder = $namespace.getDefaultFolder($olFolders::olFolderInbox) 
    # Loop through selected folder and return attachment fileName for the selected fileType
    $folder.Items | foreach {
        $SendName = $_.SenderName
        $_.attachments | ForEach-Object {
            $attachmentName = $_.fileName
            # Specify the fileType here
            $fileType = ('xlsx')
            # Save attachment if it is of the correct file type
            If ( $attachmentName.Contains($fileType) ) {
                 $_.saveasfile(( Join-Path $path $SendName"_"$attachmentName ))
            }
        }
     }
 }

Save-OutlookAttachment
