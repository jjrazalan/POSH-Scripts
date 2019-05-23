<# 
.SYNOPSIS 
Add permissions to someone's calendar

.DESCRIPTION
Connects to MS exchange easily and simply prompts for user who needs permission and whos calendar.

.AUTHOR
    Josh Razalan
#>

#Prompt for admin email and app password
Write-Host "Enter your O365 credentials with app password"
Start-Sleep -Seconds 2
$LiveCred = Get-Credential

#Connect to MS Exchange Online
Write-Host "Connecting to MS Exchange Online"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection 
Import-PSSession $Session 

#Prompt for username, calendar, and access rights
$calendar = Read-Host "Enter the username to change calendar permissions"
$username = Read-Host "Enter username of the user to be added to $calendar's calendar"
$ar = Read-Host "Enter $username access rights ie: Editor, Contributor, PublishingEditor"

#Set Permissions
add-mailboxfolderpermission -Identity "$calendar@engsys.com:\calendar" -user "$Username@engsys.com" -AccessRights $ar