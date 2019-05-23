<# 
.SYNOPSIS 
Remove and readd printers

.DESCRIPTION
prompts user for their computer location and will then set the print server location based on that input. 
Then it will remove all printers that are from that server and re-add those printers from the array.

.AUTHOR
    Josh Razalan
#>

$location = Read-Host "Enter computer location"

#Hashtable for print servers
$prtsvr = @{
}

#Set print serverlocation 
$prtsvrlocation = $prtsvr.$location

#Get array of printers on that server
$printers = Get-Printer -name "\\$prtsvrlocation\*"

#remove then add printers
Foreach ($printer in $printers) {
    Remove-Printer -Name $printer.Name
    Add-Printer -ConnectionName $printer.Name
}