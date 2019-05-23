#Hide Task Sequence window
$TSProgressUI = New-Object -COMObject Microsoft.SMS.TSProgressUI
$TSProgressUI.CloseProgressDialog()
 
#Prompt for input
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$username = [Microsoft.VisualBasic.Interaction]::InputBox("Enter a username to add startup setup script, else leave empty.", "Username prompt")
 
#Create schedule task for setup script
if ($username) {
    #add user to local admin group
    Add-LocalGroupMember -Group "Administrators" -Member "engsys.net\$username"
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-noexit -ExecutionPolicy Bypass -File \\engsys.net\Public\Global\IT\POSH\MDT\firstTimeSetup.ps1"
    $principal = New-ScheduledTaskPrincipal -UserId SYSTEM -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries
    $definition = New-ScheduledTask -Action $action -Principal $principal -Trigger $trigger -Settings $settings
    Register-ScheduledTask -TaskName "FirstTimeSetup" -InputObject $definition -User "ENGSYS\$username"
    Write-Output "Scheduled Task added."
}
if ($null -eq $username){}
Write-Output "Script completed."
Start-Sleep -Seconds 5