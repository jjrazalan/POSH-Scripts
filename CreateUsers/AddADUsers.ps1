<#
	This script is used to create Active Directory users by pulling information from a CSV file and setting the correct attributes,
	including attributes for msdirsync.
    Author: Joshua Razalan
    #>

# Import active directory module for running AD cmdlets
Import-Module ActiveDirectory
  
#Store the data from ADUsers.csv in the $ADUsers variable
$ADUsers = Import-csv "\\Path\to\ADUsers.csv"
#Connect to Exchange Online
$UserCredential = Get-Credential -Message "Enter credentials for Exchange Online (App Password)."
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking
Add-MailboxPermission -identity "$username@domain.com" -user "Organization Management" -AccessRights FullAccess -inheritancetype all -automapping $false
Remove-PSSession $Session

#Loop through each row containing user details in the CSV file 
foreach ($User in $ADUsers) {
	
    $Username = $User.username
    $Desc = ''

    #Read user data from each field in each row and assign the data to a variable as below
    $Firstname = $User.firstname
    $MI = $User.mi
    $Lastname = $User.lastname
    $Password = $User.password
    $Office = $User.office
    $Department = $User.department
    $JobTitle = $User.jobtitle
    $License = $User.o365
    
    #Set desciption by user intials
    $tempDesc = $Firstname + ' ' + $MI + ' ' + $Lastname	
    $tempDesc.split(' ') | ForEach-Object { $Desc += $_[0] }
    $Desc = $Desc.ToUpper()

    #Set correct syntax for Display Names.
    if (!$MI) {
        $Name = $Firstname + ' ' + $Lastname
    }
    else { 
        $Name = $Firstname + ' ' + $MI + '. ' + $Lastname
    }   

    #O365 License
    Switch ($License) {
        "y" {
            $O365 = $true
            $License = 'E3'
        }
        "yes" {
            $O365 = $true
            $License = 'E3'
        }
        "Y" {
            $O365 = $true
            $License = 'E3'
        }
        "t" {
            $O365 = $false
            $License = 'E3'
        }
        "true" {
            $O365 = $false
            $License = 'E3'
        }
        "T" {
            $O365 = $false
            $License = 'E3'
        }
        "n" {
            $O365 = $false
        }
        "no" {
            $O365 = $false
        }
        "N" {
            $O365 = $false
        }
        default {
            $O365 = $false
        }
    }

    #Connect to Microsoft Online
    Get-MsolDomain -ErrorAction SilentlyContinue
    if ($?) {
    }
    else {
        Connect-MsolService
    }


    #Check if user exists then call methods to create user in AD in the various servers
    if (-Not (Get-ADUser -Filter 'SamAccountName -eq $Username' -Server "")) {
        try {
            . \\Path\to\domainFunctions.ps1
            createUserdomain $Firstname $MI $Lastname $Username $Password $Office $Department $JobTitle $Desc $Name 
            . \\Path\to\AddO365Users.ps1
            CreateO365User $Firstname $Lastname $Username $Password $Name $O365 $License
            Write-Host "$Username sucessfully created in domain"
        }
        catch {
            Write-Host "$Username creation unsucessful in domain"
        }
    }
    else {
        Write-Host "$Username already exist in domain"
    }

    if (-Not (Get-ADUser -Filter 'SamAccountName -eq $Username' -Server "other domain")) {
        try {
            . \\Path\to\otherdomainFunctions.ps1
            createUserotherdomain $Firstname $MI $Lastname $Username $Password $Office $Department $JobTitle $Desc $Name
            Write-Host "$Username sucessfully created in other domain"
        }
        catch {
            Write-Host "$Username creation unsucessful in other domain"
        }
    }
    else {
        Write-Host  "$Username already exist in other domain"
    }
}