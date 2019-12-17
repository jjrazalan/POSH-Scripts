<# 
.SYNOPSIS 
Callable script that creates user in the database

.DESCRIPTION
This script is modified after running the sqlfunctions script to change the temp variables to the information from the sqlfunction variables. This allows a schedule task to be called after
the database sync task. It then modifies the new database with the required information for the user to function.

.AUTHOR
    Josh Razalan
#>
#Setup database server parameters
$SQLServer = '0.0.0.0' 
$Database = ''
$DBUser = ''
$DBPass = ''

#Declare variables from SQLFunctions.ps1
$NDBID = 'tempid'
$Uname = 'tempusername'
$uemail = 'tempemail'

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
$SQL = "UPDATE user_rcrd
        SET user_rcrd_ldapusername = '$Uname', user_rcrd_email = '$uemail'
        WHERE user_rcrd_origID = $NDBID;"

$Command.CommandText = $SQL
$Command.ExecuteNonQuery() 

#Close connection to reduce overhead
$Connection.Close()

Unregister-ScheduledTask -TaskName "NewDBUpdate$NDBID" -Confirm:$false