<# 
.SYNOPSIS 
Change the faro licnese server information remotely

.DESCRIPTION
Meant to be run with screenconnect or a RMM to call the script on the remote computer. It changes the license server configuration file to the new server addresses.

.AUTHOR
    Josh Razalan
#>

$Logfile = "\\Path\to\computers.log"

try {
    $configfile = "C:\Program Files (x86)\Common Files\Aladdin Shared\HASP\hasplm.ini"
    $hasplm = $configfile
    $oldipaddress = get-childitem -path $configfile | select-string -pattern "serveraddr = ip address"
    $newipaddress = get-childitem -path $configfile | select-string -pattern "serveraddr = other ip address"

    if (($null -ne $oldipaddress) -and ($null -ne $newipaddress)){
        Exit-PSHostProcess
    }
    else{
        (Get-Content $hasplm -Raw) -replace 'serveraddr = ip address', "serveraddr = other ip address2`nserveraddr = other ip address" | Set-Content $hasplm
        (Get-Content $hasplm -Raw) -replace "serveraddr = computer", "serveraddr = other computer`nserveraddr = other computer" | Set-Content $hasplm
        
        Add-Content $Logfile -Value "$env:COMPUTERNAME - Success"
    }
}
catch {
    Add-Content $Logfile -Value "$env:COMPUTERNAME - Failure"
}