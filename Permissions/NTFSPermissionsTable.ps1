<# 
.SYNOPSIS 
Create report for folders with explicitly defined NTFS permissions

.DESCRIPTION
Recurssive searches the selected folder for any directories that contain explicitly defined permissions. Then saves the file as HTML with permissions for the parent folder, 
permissions for the explicitly defined folders, and list out all members of the explicitly defined folder permission groups.

.AUTHOR
    Nick Yakich & Josh Razalan
#>

#Function to draw folder browsing box
function Select-Folder {
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $browse = New-Object System.Windows.Forms.FolderBrowserDialog
    #Default browsing location 
    $browse.SelectedPath = " "
    #Hide New Folder button
    $browse.ShowNewFolderButton = $false
    $browse.Description = "Select folder to check for explicit permissions"

    $loop = $true
    while ($loop) {
        if ($browse.ShowDialog() -eq "OK") {
            #end loop after folder is selected
            $loop = $false
        }
        else {
            $res = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
            #exit if canceled
            if ($res -eq "Cancel") {
                exit
            }
        }
    }
    #return selected folder
    $browse.SelectedPath
    $browse.Dispose()
}

#Function to select output file
function Select-OutFile {
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.SaveFileDialog -Property @{
        #Set a file extension filter
        Filter = 'HTML Files (*.HTML)|*.HTML|All files (*.*)|*.*'
    }
    $FileBrowser.FileName = (Get-Item -path $folder | Select-Object Name).name
    [void]$FileBrowser.ShowDialog()

    #return outfile name if not canceled
    If ($FileBrowser.FileNames -like "*\*") {
        Return $FileBrowser.FileName
    }
    else {
        exit    
    }
}

#Function to translate SIDs to name
function Find-Name {
    $adSid = ($access.identityreference) 
    #Check Group first
    try {
        $objUser = (Get-ADGroup -Identity $adSid).Name
        $name = @{
            name  = $objUser
            group = $objUser
        }
    }
    catch { $objUser = $null }
    #Check User
    if (!($objUser)) {
        try {
            $objUser = (Get-ADUser -Identity $adSid).Name
            $name = @{
                name = $objUser
            }
        }
        catch { $objUser = $null }
    }
    #Else keep adSID account name 
    if (!($objUser)) {
        $objUser = $adSid
        $name = @{
            name = $objUser
        }
    }
    return $name
}

#create HTML table for folder permissions and gather groups
function New-PermissionTables {
    #Declare parameters
    $Script:permissionArray = @()
    $Script:groups = @()
    $Script:accessTable = New-Object System.Collections.ArrayList
    $count = 0
    foreach ($explicitFolder in $explicitResults) {
        $permissionFolder = (Get-Acl -Path $explicitFolder.FullName).Access | Select-Object identityreference, FilesystemRights, AccessControlType
        #Flush out arrays for loop
        $groupArray = @()
        $accessArray = @()
        foreach ($access in $permissionFolder) {
            $names = Find-Name $access
            $access.IdentityReference = $names.name
            $accessArray += $access
            $groupArray += $names.group
        }
        $accessTable.Add($accessArray) | Out-Null
        $folderName = $explicitFolder.PsChildName
        $accessTableName = '$accessTable[' + "$count" + ']'
        $Script:permissionArray += "Section -Name '$folderName' {
            Table -HideFooter -DataTable $accessTableName
        }
        " 
        $Script:groups += $groupArray
        $count++
    }
}

#Function to find and create tables for group members
function Find-GroupMembers {
    $Script:GroupHTML = @()
    $script:groupMembersArray = @()
    $Script:groupMembersTable = New-Object System.Collections.ArrayList
    $count = 0
    foreach ($group in $groups) {
        if ($group) {
            $members = Get-ADGroupMember -Identity $group | Select-Object name
            $groupName = $group 
            
            $groupMembersTable.Add($members) | Out-Null
            $groupMembersName = '$groupMembersTable[' + "$count" + ']'
            $Script:GroupHTML += "Section -Name '$groupName' {
                Table -HideFooter -DataTable $groupMembersName
            }
            "
            $count++     
        }        
    }
}


#Ask user for the directory to check permissions and get that folder's permission
$Folder = Select-Folder
$parentFolder = (Get-Acl -Path $Folder).Access | Select-Object identityreference, FilesystemRights, AccessControlType
foreach ($access in $parentFolder) {
    $names = Find-Name $access
    $access.identityreference = $names.name
}
#Create HTML table for parent folder
$parentFolder = [PSCustomobject] $parentFolder #| ConvertTo-Html -Title "$Folder" -PreContent "<h3>$Folder<h3>"  

#Recusivly search selected directory for folder containing explicit permissions and output to selected file. 	
Write-Host "Looking for explicit permissions..."
$explicitResults = Get-ChildItem $Folder -Recurse -Directory -ErrorAction SilentlyContinue |
Where-Object { (Get-Acl -Path $_.FullName -ErrorAction SilentlyContinue).Access | 
    Where-Object { $_.isInherited -eq $false } } 

#If permissions found then get those folders get access report
if ($explicitResults) {
    Write-Host "Explicit permissions found."
    $explicitResultsHTML = $explicitResults | Select-Object PsChildName #| ConvertTo-Html -Title "Explicit Folders" -PreContent "<h3>Explicit Folders<h3>"
    #Call function to create permissions table for each explicit folder.
    New-PermissionTables $explicitResults

    #Call function to create table for groups.
    $groups = $groups | Select-Object -Unique
    Find-GroupMembers $groups
}
else {
    Write-Warning "None found!"
}
#Ask user to select an output file
$OutFile = Select-OutFile	

#Create Table
$permissionArray = [Scriptblock]::Create($permissionArray)
$groupHTML = [Scriptblock]::Create($groupHTML)
Dashboard -Name "Explicit Permissions Report" -FilePath $OutFile {
    Tab -Name "$Folder" {
        Section -Name  "Explicit Permissions Report" {
            Table -HideFooter -DataTable $parentFolder
        }
        Section -Name "Explicit Folders" {
            Table -HideFooter -DataTable $explicitResultsHTML
        }
    }
    Tab -Name "Groups and members" {
        . $permissionArray
        . $GroupHTML
    }
}


#$Explicitresults | Add-Content -Path $OutFile
Invoke-Item $OutFile