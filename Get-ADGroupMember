Import-Module ActiveDirectory
$Group = Read-Host "Please enter a group name"
If ($?){
    Write-Host "Finding group members..."
    Get-ADGroupMember -Identity $Group | Select name | Export-CSV C:\Users\Charles.Cox\desktop\group_members.csv
}
