Write-Host "This script will give you the time on a remote computer."
$computer = Read-Host -Prompt "Please enter a computer name: " 
If ($?){
	Get-WmiObject -Class win32_localtime -ComputerName $computer
	}
	Format-Table
