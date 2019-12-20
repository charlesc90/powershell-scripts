########################################################################################################
# Script: Change-LocalAdminPassword.ps1
# Author: Charles Cox
# Date: 20 December 2019
# Description: This script reads a list of computers from a file and changes their local admin passwords
########################################################################################################

# Read list of computers from text file
$Computers = Get-Content -path C:\Users\Charles.Cox\Desktop\computers.txt
$Password = Read-Host "Enter the password" -AsSecureString
$Confirmpassword = Read-Host "Confirm the password" -AsSecureString
$Pwd1_text = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
$Pwd2_text = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($confirmpassword))

if($Pwd1_text -ne $Pwd2_text) { 
	Write-Host "The passwords do not match."
}

if($Pwd1_text -eq $Pwd2_text) {
	Foreach($computer in $computers) {
		$Computer = $Computer.toupper()
		$Isonline = "OFFLINE"
		$Status = "SUCCESS"
		Write-Verbose "Working on $Computer"
		if((Test-Connection -ComputerName $Computer -count 1 -ErrorAction 0)) {
			$Isonline = "ONLINE"
			Write-Verbose "`t$Computer is Online"
        }
		else {
			Write-Verbose "`t$Computer is OFFLINE"
		}
		try {
			$Account = [ADSI]("WinNT://$Computer/Administrator,user")
			$Account.psbase.invoke("setpassword",$Pwd1_text)
			Write-Verbose "`tPassword Change completed successfully"
        }
		catch{
			$Status = "FAILED"
			Write-Verbose "`tFailed to Change the administrator password. Error: $_"
        }
		$Obj = New-Object -TypeName PSObject -Property @{
			ComputerName = $Computer
			IsOnline = $Isonline
			PasswordChangeStatus = $Status
		}
		$Obj | Select ComputerName, IsOnline, PasswordChangeStatus 

		if($Status -eq "FAILED" -or $Isonline -eq "OFFLINE") {
			$Stream.writeline("$Computer `t $Isonline `t Status")
		}
	}
}
Write-Host "Press any key to continue ..."
