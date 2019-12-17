<# 
.SYNOPSIS 
Finishes up tasks on user's first time setup on a new machine

.DESCRIPTION
Called and ran after the inital MDT setup under the assigned user. 

.AUTHOR
    Josh Razalan
#>

function logontoOutlook {
	#Open outlook
	Invoke-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\outlook.lnk"
	Start-Sleep -Seconds 90
	#Create signature
	& "\\Path\to\sig.ps1"
}

function setupVPN{
	#CMD command to request a vpn certificate from Certificate authority
	$command = "certreq -enroll -kerberos -q VPNCertificate"
	Invoke-Expression $command 
	
	#Copy shortcut to desktop
    $aclink = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Cisco\Cisco AnyConnect Secure Mobility Client\Cisco AnyConnect Secure Mobility Client.lnk"
    Copy-Item -Path $aclink -Destination "C:\Users\$env:Username\desktop"

	#Copy anyconnect profile to local user
	$item = "\\Path\to\preferences.xml"
	$dest = "C:\Users\$env:USERNAME\AppData\Local\Cisco\Cisco AnyConnect Secure Mobility Client\preferences.xml"

	
	mkdir "C:\Users\$env:USERNAME\AppData\Local\Cisco\Cisco AnyConnect Secure Mobility Client\"
	Copy-Item $item -Destination $dest -Recurse -Force   

	$configFile = $dest
	(Get-Content $configFile -Raw) -replace 'Temp', "$env:USERNAME" | Set-Content $configFile
	Write-Output "Anyconnect Profile copied."
	Start-Sleep -Seconds 3
}
Write-Host "Please wait 3 minutes as first time startup background processes completes..."

Do {
	$vpnuser = Read-Host "Is $env:Username a vpn user? Enter y or n"
	if ($vpnuser -eq "y" -or $vpnuser -eq "n") {
		$valid = $true
	}
	if ($null -eq $vpnuser) {
		$valid = $false
	}
}
until ($valid) {} 

#Prompt for secure password and change to plain text
$setup = Read-Host "Set up $env:Username's Outlook? y or n"

invoke-expression -Command "\\Path\to\AddWindowsAssetTags.ps1"

#if user is VPN then copy profile and request cert
if($vpnuser -eq "y"){
	setupVPN
}
else {}

#Check if known password then run logon
if (!($setup -eq "n")) {
	logontoOutlook 
}
#Copy common O365 program shortcuts
$o365Programs = @("Excel", "Outlook", "PowerPoint", "Word")
foreach ($program in $o365Programs){
	Copy-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\$program.lnk" -Destination "C:\Users\$env:Username\desktop"
}

Unregister-ScheduledTask -TaskName "FirstTimeSetup" -Confirm:$false
Write-Host "Setup complete. Check adobe."