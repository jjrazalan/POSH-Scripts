<# 
.SYNOPSIS 
Connect to a Team Site

.DESCRIPTION
Creates the request URL for a given user to automatically map the Team Site sharepoint to the user's onedrive. It will then go ahead and invoke the web request

.AUTHOR
    Josh Razalan
#>

#Team Site Variables
$siteID = "&siteId=%7B775367df%2D8a74%2D4698%2D94fb%2Dc6bdae4c227a%7D"

#Grab the sync information through a browser with network monitoring and after click the sync button the team site
$webID = " "
$webTitle = " "
$webTemplate = " "
$webLogoURL = " "
$webURL = " "
$onPrem = " "
$libraryType = " "
$listID = " "
$listTitle = " "

#Prompt for input
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$username = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter your email address:", "Email prompt")

#Make sure that the user enter a valid email address
try {
    $userName = $username.substring(0, $username.indexof('@'))
}
catch {
    Add-Type -AssemblyName PresentationFramework
    $status = [System.Windows.MessageBox]::Show('Run script again and enter a valid email address.', 'User Error', '0', 'Error')
    
    if ($status) {
        exit
    }
}

#Change '-' to %2D  
$userEmail = "userEmail=$userName%40 %2E "

$result = "odopen://sync?$userEmail$siteID$webID$webTitle$webTemplate$webLogoURL$webURL$onPrem$libraryType$listID$listTitle"
Write-Host `n"Starting Web Request: " `n`n$result`n

& start "$result"