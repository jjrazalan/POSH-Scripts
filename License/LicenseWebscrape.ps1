<# 
.SYNOPSIS 
Create User viewable webpage from web scrapped data 

.DESCRIPTION
Invoke a web request and grabs the session data (if any) of faro licensing to create a dashimo html table.

.AUTHOR
    Josh Razalan
#>
#Get UNIX time
$epoch = (Get-Date -Date ((Get-Date).DateTime) -UFormat %s)

#Set webrequest URI and capture the request
$uri = "http://farolicense_computer:1947/_int_/tab_sessions.html?haspid=0&featureid=-1&vendorid=0&productid=0&filterfrom=1&filterto=20&timestamp=$epoch?"
$sessions = Invoke-RestMethod -Uri $Uri

#Select only json data
$sessionsData = $sessions | 
ForEach-Object { [regex]::Matches($_, '(\{[^\}]+\},)') } |
ForEach-Object { $_.Groups[1].value }

#Parameters to appened to sessionsData
$start = "{`"sessions`":[
     "
$lastString = " ]}"

#Begin building correct Json
$sessionsData = "$start $sessionsData"
$sessionsData = $sessionsData -replace ",$"
$sessionsData = "$sessionsData $lastString"
#Convert object from Json to PSObject
$sessionsObj = ConvertFrom-Json -InputObject $sessionsData
#Select only relevent information and set aliases 
$sessionsobj = $sessionsObj.sessions | Select-Object -Property @{N = 'Product Name'; E = { $_.productname } },
@{N = 'Feature'; E = { $_.fn } },
@{N = 'User'; E = { $_.usr } },
@{N = 'IP Address'; E = { $_.cli } },
@{N = 'Computer Name'; E = { $_.mch } },
@{N = 'Login Time'; E = { $_.lt } }

#Output HTML files
Dashboard -Name "Sessions on Faro License Server" -FilePath \\Path\to\sessions.html -AutoRefresh 60 { 
    Section -Name  "Active Sessions on Faro License Server" {
        Table -HideFooter -DataTable $sessionsObj 
    }
}