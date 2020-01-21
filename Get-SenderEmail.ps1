########################################################################################################
# Script: Get-SenderEmail.ps1
# Date: 20 December 2019
# Description: This script displays the unique email addresses of senders in a specified mailbox
########################################################################################################

Function Get-SenderEmail {
    
    # Specify mailbox here
    $Folder = "Inbox"

    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" | Out-Null
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNameSpace("MAPI")
    $emailAddress = $namespace.Folders.Item(1).Folders.Item($Folder).Items
    $emailAddress | Sort-Object SenderEmailAddress -Unique | Format-Table SenderEmailAddress

}
Get-SenderEmail
