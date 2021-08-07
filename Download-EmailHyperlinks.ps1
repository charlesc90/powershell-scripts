Function Get-Reports {


    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" | Out-Null
    $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
    
    # define Outlook ComObject amd namespace
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNameSpace("MAPI")

    # This Outlook mailbox EntryID is the EntryID for the specific mailbox you want to download .csv (or any) attachments from.
    $folder = $namespace.GetFolderFromID("000000008EB31A6D577EF84F83DEE6F4DA1FF0AD0100D3055F6334628E49B54C2CDB3FFCC16700000000192F0000")
   
    # Select the emails where the sender name is name@domain.com
    $report = $folder.Items | Where-Object { $_.SenderName -eq "name@domain.com" }
    
    # Loop through the emails
    For ($i = 0; $i -lt $report.Count; $i++) {

        # Generally, emails come with an HTML body and the link to download your file somewhere in the HTML body. This is where you use a regex to match and return the url.
        If ($report[$i].HTMLBody -match '..dtaxshares.dtaxreports.prodims.taxtool_activity............') {

            $urls = $matches[0]
            
        }
        # If you need to name the file you are downloading from something in the HTML body, this regex matches and returns that pattern.
        If ($report[$i].HTMLBody -match 'CAN</p></font></td><td><font face="ARIAL" size="-1">&nbsp;&nbsp;&nbsp;.................') {

            # The CAN is after the first 70 characters of HTML in the HTMLBody
            $cans = $matches[0].Substring(70)
            
        }
        
        # Download the files from the hyperlinks into the directory you define here
        Invoke-WebRequest -Uri file:$urls -OutFile "C:\Dallas\reports\$cans.pdf"
        Write-Host "Downloaded:" $cans "from:"$urls -ForegroundColor Yellow
        Add-Content -Value "$cans","$urls" -Path C:\Dallas\log.csv

    }
    
}
Get-Reports
