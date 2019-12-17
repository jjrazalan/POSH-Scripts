<# 
.SYNOPSIS 
Functions to create user in database

.DESCRIPTION
Called from the AddADUsers(GUI).ps1 to create the newly created active directory user in the databases as well. It takes the checkbox inputs and dropdowns and converts
it to the correct corresponding variables to create the user in the database. This script requires the NET connector to connect to the my sql database
found here: https://dev.mysql.com/downloads/connector/net/

.AUTHOR
    Josh Razalan
#>

function GUIConvert {
    #Convert AD Offices to correct database offices
    if ($Corporate) {
        $Office = 'Corporate'
    }
    #No Tampa in database, convert to Florida office
    if ($Office -eq 'Tampa') {
        $Office = 'Georgia'
    }
    #Convert Space in North Carolina to be able to be in the Office Hashtable
    if ($Office -eq 'North Carolina') {
        $Office = 'NCarolina'
    }

    #Hashtable for Office to line up ID# to office name
    $Offices = @{
        Illinois  = '1'
        Colorado  = '2'
        Missouri  = '3'
        Georgia   = '4'
        Corporate = '5'
        Florida   = '6'
        Houston   = '7'
        Michigan  = '8'
        SoCal     = '9'
        Nebraska  = '10'
        Dallas    = '11'
        NCarolina = '12'
        Iowa      = '13'
        Seattle   = '15'
        Minnesota = '16'
        Miami     = '17'
    }
    #Set office to corresponding ID#
    $script:Office = $Offices.$Office


    #Convert Strings to 1 readable line for hashtable key
    if ($PG -eq 'Auto & Marine') {
        $PG = 'Auto'
    }
    if ($PG -eq 'Bio & Safety') {
        $PG = 'Bio'
    }
    if ($PG -eq 'Civil, Structural, & Enviormental') {
        $PG = 'Civil'
    }
    if ($PG -eq 'Fire & Explosion') {
        $PG = 'Fire'
    }
    if ($PG -eq 'Laboratory & Industrial Services') {
        $PG = 'Lab'
    }

    #Hashtable for Practice Group to line up ID#
    $PracticeGroups = @{
        None          = '0'
        Auto          = '1'
        Aviation      = '2'
        Bio           = '3'
        Civil         = '4'
        Materials     = '5'
        Fire          = '6'
        Mechanical    = '7'
        Lab           = '9'
        Visualization = '10'
        Electrical    = '11'
        Rail          = '12'
        Admin         = '13'
        Technical     = '14'
    } 
    #Set Practice Group to corresponding ID#
    $script:PG = $PracticeGroups.$PG

    #Setup corresponding values to checkedboxes from GUI
    if ($Print) {
        $script:Print = '1'
    }
    else {
        $script:Print = '0'
    }

    if ($Admin) {
        $script:Admin = 'Y'
    }
    else {
        $script:Admin = 'N'
    }
    if ($ESISearch) {
        $script:ESISearch = '1'
    }
    else {
        $script:ESISearch = '0'
    }
    if ($Exempt) {
        $script:Exempt = 'Y'
    }
    else { 
        $script:Exempt = 'N'
    }
    if ($Classification) {
        $script:Classification = '1'
    }
    else {
        $script:Classification = '0'
    }

}

function CreateScheduledTask {
    $TempleteFile = '\\Path\to\NewDBTemplate.ps1'
    #replace temp variables in the new database script and create task with the user id
    (Get-Content $TempleteFile) | ForEach-Object {
        $_ -replace 'tempusername', "$Username" `
            -replace 'tempemail', "$Username@domain.com" `
            -replace 'tempid', $UID
    } | Set-Content "\\Path\to\NewDB$UID.ps1"
    Start-Process "\\Path\to\schtaskDB.bat" $UID
}

function CreateUserSQL { 

    #Set required parameters for database insert
    $IOBoard = 1
    $Password = $Desc.ToLower() + '1'
    GUIConvert $Office $PG $Print $Admin $ESISearch $Exempt $Classification $Corporate

    #Setup database server parameters
    $SQLServer = '' 
    $Database = ''
    $DBUser = ''
    $DBPass = ''

    #Create connection to database
    Add-Type -Path "C:\Program Files (x86)\MySQL\MySQL Connector Net 8.0.17\Assemblies\v4.5.2\MySQL.Data.dll"

    # This is essential to solve the missing Renci.SshNet.dll error
    Add-Type -Path 'C:\Program Files (x86)\MySQL\MySQL Connector Net 8.0.17\Assemblies\v4.5.2\Renci.SshNet.dll'

    $Connection = New-Object MySql.Data.MySqlClient.MySqlConnection
    $Connection.ConnectionString = "server='$SQLServer'; database='$Database';uid=$DBUser; pwd=$DBPass"
    $Connection.Open()
    $Command = New-Object MySql.Data.MySqlClient.MySqlCommand
    $Command.Connection = $Connection

    #Create SQL Command to send to the database
    $SQL = "INSERT INTO Employee (Initials, Password, Lname, Fname, Mname, Location, PG, Print, ShowIOBoard, Classification, Admin, ESISearchPermitted, Exempt)
            VALUES ('$Desc','$Password','$Lastname','$Firstname','$MI','$script:Office','$script:PG','$script:Print','$IOBoard','$script:Classification','$script:Admin','$script:ESISearch','$script:Exempt');
            SELECT LAST_INSERT_ID();"

    $Command.CommandText = $SQL
    $UID = $Command.ExecuteScalar() 

    #Close connection to reduce overhead
    $Connection.Close()

    #Call function to create the schedule task powershell script
    CreateScheduledTask $UID $Username
}