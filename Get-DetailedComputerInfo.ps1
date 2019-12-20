$a = Get-Content "C:\Users\Charles.Cox\Documents\computers.txt"

#To seach and OU first to get list of computers
#$a = Get-ADComputer -SearchBase 'OU=OU Name,dc=domain,dc=com' -Filter '*' | Select -Exp Name

foreach ($i in $a)
   {
        $servicetag = (Get-WmiObject win32_SystemEnclosure -computername $i).serialnumber 
        $computermodel = (Get-WmiObject -computer $i -Class:Win32_ComputerSystem).Model
        $osversion = (Get-WMIObject -computer $i win32_operatingsystem).caption
        $osarchitechture = (Get-WmiObject -computer $i Win32_OperatingSystem).OSArchitecture
        $osbuild = (Get-WmiObject -computer $i Win32_OperatingSystem).buildNumber
        $computerinfo = [pscustomobject][ordered] @{
            "Computer Name"  = $i
            "Computer Model" = $computermodel
            "Computer Service Tag" = $servicetag 
            "Computer OS" = $osversion
            "OS Architechture" = $osarchitechture
            "OS Build" = $osbuild
          
            }
       $computerinfo | Export-Csv -append "C:\Users\Charles.Cox\Documents\computerInfo.csv" -noType
    }
