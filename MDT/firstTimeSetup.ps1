function logontoOutlook {
	param (
		[String]$pw
	)
	#Open outlook and allow user to log in
	Invoke-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\outlook.lnk"
	Start-Sleep -Seconds 90

	#Invoke signature scripts for users
	& ".\Signatures\sig.ps1"
}

function setupVPN{
	#Using certutil to request a vpn cert through the VPNCertificate template
	$command = "certreq -enroll -kerberos -q VPNCertificate"
	Invoke-Expression $command 

	#Copy anyconnect preferences over to user's appdata
	$item = ".\Cisco\AnyConnect\PC\AnyConnect Profile\preferences.xml"
	$dest = "C:\Users\$env:USERNAME\AppData\Local\Cisco\Cisco AnyConnect Secure Mobility Client\preferences.xml"

	
	mkdir "C:\Users\$env:USERNAME\AppData\Local\Cisco\Cisco AnyConnect Secure Mobility Client\"
	Copy-Item $item -Destination $dest -Recurse -Force   

	#Then change the temp username in preferences to the current logged on user
	$configFile = $dest
	(Get-Content $configFile -Raw) -replace 'Temp', "$env:USERNAME" | Set-Content $configFile
	Write-Output "Anyconnect Profile copied."
	Start-Sleep -Seconds 3
}
#Wait for background processes to finish before starting the script
Write-Host "Please wait 3 minutes as first time startup background processes completes..."

#Prompt for VPN usage
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
$pw = Read-Host -AsSecureString "Enter $env:Username's O365 password if known, else enter 'n'"
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pw)
$pw = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

#if user is VPN then copy profile and request cert
if($vpnuser -eq "y"){
	setupVPN
}
else {}

#Check if known password then run logon
if (!($pw -eq "n")) {
	#logontoOutlook $pw
}
#Copy common O365 program shortcuts
$o365Programs = @("Excel", "Outlook", "PowerPoint", "Word")
foreach ($program in $o365Programs){
	Copy-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\$program.lnk" -Destination "C:\Users\$env:Username\desktop"
}

#Delete scheduled task from MDT setup
Unregister-ScheduledTask -TaskName "FirstTimeSetup" -Confirm:$false
Write-Host "Setup complete. Check adobe."