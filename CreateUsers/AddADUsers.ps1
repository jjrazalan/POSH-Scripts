<#
	This script is used to create Active Directory users by pulling information from a CSV file and setting the correct attributes,
	including attributes for msdirsync.
    Author: Joshua Razalan
    #>

# Import active directory module for running AD cmdlets
Import-Module ActiveDirectory
  
#Store the data from ADUsers.csv in the $ADUsers variable
$ADUsers = Import-csv "P:\Global\IT\POSH\CreateUsers\ADUsers.csv"

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
    $tempDesc.split(' ') | ForEach-Object {$Desc += $_[0]}
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


    #Check if user exists then call methods to create user in AD in ENGSYS.NET and esi.local
    if (-Not (Get-ADUser -Filter 'SamAccountName -eq $Username' -Server "IL-DC01.ENGSYS.NET")) {
        try {
            . P:\Global\IT\POSH\CreateUsers\engsysFunctions.ps1
            createUserEngsys $Firstname $MI $Lastname $Username $Password $Office $Department $JobTitle $Desc $Name 
            . P:\Global\IT\POSH\CreateUsers\AddO365Users.ps1
            CreateO365User $Firstname $Lastname $Username $Password $Name $O365 $License
            Write-Host "$Username sucessfully created in engsys.net"
        }
        catch {
            Write-Host "$Username creation unsucessful in engsys.net"
        }
    }
    else {
        Write-Host "$Username already exist in Engsys"
    }
    
    if (-Not (Get-ADUser -Filter 'SamAccountName -eq $Username' -Server "IL-HADC.esi.local")) {
        try {
            . P:\Global\IT\POSH\CreateUsers\esilocalFunctions.ps1
            createUserEsilocal $Firstname $MI $Lastname $Username $Password $Office $Department $JobTitle $Desc $Name
            Write-Host "$Username sucessfully created in esi.local"
        }
        catch {
            Write-Host "$Username creation unsucessful in esi.local"
        }
    }
    else {
        Write-Host"$Username already exist in Esi.local"
    }
}