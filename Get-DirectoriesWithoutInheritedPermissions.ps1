$path = "\\server\share"
Get-ChildItem $path -recurse | Select @{Name='Path';Expression={$_.FullName}},@{Name='InheritedCount';Expression={(Get-Acl $_.FullName | Select -ExpandProperty Access | Where { $_.IsInherited }).Count}} | Where { $_.InheritedCount -eq 0 } | Select Path | Get-ACL | Format-Table -AutoSize -wrap | Out-File C:\users\Charles.Cox\desktop\perms.txt
