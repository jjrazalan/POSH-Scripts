#Replaces the bgt name in the mgrfx ou to mgrfx to reflect the ou
#Author: Josh Razalan
Get-ADGroup -filter 'Name -like "*BGT*"' -searchbase "OU=MGRFX,DC=,DC=" | Foreach-Object {
	Set-ADGroup -Identity $_ -SamAccountName ($_.SamAccountName.Replace("BGT","MGRFX"))
	Rename-ADObject -Identity $_ -NewName ($_.Name.Replace('BGT','MGRFX'))
	}