<# 
.SYNOPSIS 
Remove and re-add printers

.DESCRIPTION
prompts user for their computer location and will then set the print server location based on that input. 
Then it will remove all printers that are from that server and re-add those printers from the array.

.AUTHOR
    Josh Razalan
#>

$location = Read-Host "Enter computer location"

$prtsvr = @{    
    Colorado  = "print server name";
}
$prtsvrlocation = $prtsvr.$location
$printers = Get-Printer -name "\\$prtsvrlocation\*"
Foreach ($printer in $printers) {
    Remove-Printer -Name $printer.Name
    Add-Printer -ConnectionName $printer.Name
}