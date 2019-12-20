# This function retreives mail from the default outlook account inbox.
Function Save-OutlookAttachment {

    # Add .NET core class for outlook
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" | Out-Null
    $olFolders = “Microsoft.Office.Interop.Outlook.olDefaultFolders” -as [type]
    # Define outlook COM object and namespace
    $outlook = New-Object -Comobject outlook.application
    $namespace = $outlook.GetNameSpace("MAPI") 
    # Select the outlook folder to query. If you neet to select a different folder, change ::olFolder...
    $folder = $namespace.getDefaultFolder($olFolders::olFolderInbox)
    $path = "C:\Users\Charles.Cox\Desktop" 

    # Loop through selected folder and return attachment fileName for the selected fileType
    $folder.Items | foreach {
        $SendName = $_.SenderName
        $_.attachments | ForEach-Object {
            $attachmentName = $_.fileName
            $fileType = ('txt')
            $i += 1
            If ( $attachmentName.Contains($fileType) ) {
                 # $_.fileName
                 $_.saveasfile(( Join-Path $path $i"_"$attachmentName ))
            }
        }
     }
 }

Save-OutlookAttachment
