<# 
.SYNOPSIS 
Connect to MN Team Site

.DESCRIPTION
Creates the request URL for a given user to automatically map the MN Team Site sharepoint to the user's onedrive. It will then go ahead and invoke the web request

.AUTHOR
    Josh Razalan
#>

#MN Team Site Variables
$siteID = ""
$webID = ""
$webTitle = "&webTitle="
$webTemplate = "&webTemplate=64"
$webLogoURL = "&webLogoUrl="
$webURL = "&webUrl="
$onPrem = "&onPrem=0"
$libraryType = "&libraryType=3"
$listID = "&listId="
$listTitle = "&listTitle=Documents&scope=OPENLIST"

#Prompt for input
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$username = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter your email address:", "Email prompt")

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
$userEmail = "userEmail=$userName%40engsys%2Ecom"

$result = "odopen://sync?$userEmail$siteID$webID$webTitle$webTemplate$webLogoURL$webURL$onPrem$libraryType$listID$listTitle"
Write-Host `n"Starting Web Request: " `n`n$result`n

& Start-Process "$result"