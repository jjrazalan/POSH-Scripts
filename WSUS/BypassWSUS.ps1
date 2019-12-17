<# 
.SYNOPSIS 
Bypass WSUS to install additonal features

.DESCRIPTION
Changes registry key to use the microsoft servers instead of onprem WSUS server. This allows the installation of .NET addons or
RSAT tools.

.AUTHOR
    Josh Razalan
#>
try{
    #Renames the windows update key to pointing to the wsus server
    Rename-Item -path hklm:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate -NewName 'WindowsUpdateOld'
    Write-Host "Computer is now connected to the Microsoft Servers"
} 
catch{
    Write-Host "Computer is already connected to the Microsoft Servers"
}
Restart-Service -Name wuauserv

Write-Host -NoNewLine 'Press any key to revert WSUS...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

#Rename the key back so that it uses the WSUS server again.
Rename-Item -path hklm:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdateOld -NewName 'WindowsUpdate'
Restart-Service -Name wuauserv
Write-Host "Computer is now connected to WSUS"