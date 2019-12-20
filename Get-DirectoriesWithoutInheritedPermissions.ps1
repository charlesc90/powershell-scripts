$filepath = "\\server\share"
Get-ChildItem $filepath -recurse | Select @{Name='Path';Expression={$_.FullName}},@{Name='InheritedCount';Expression={(Get-Acl $_.FullName | Select -ExpandProperty Access | Where { $_.IsInherited }).Count}} | Where { $_.InheritedCount -eq 0 } | Select Path
