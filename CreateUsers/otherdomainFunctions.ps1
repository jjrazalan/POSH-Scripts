<# 
.SYNOPSIS 
Function to create user in other domain

.DESCRIPTION
Callable function with required parameters to create and assign a user licenses in the other domain. Includes setting the correct office information.

.AUTHOR
    Josh Razalan
#>
#Function to set various office information in other domain
function setOffice {
    param($Office)
    $User = "" | Select-Object -Property OU, Company, StreetAddress, City, State, PostalCode, HomePhone, Fax, Group1, Group2
    switch ($Office) {
        "Colorado" {
            $User.OU = "OU="
            $User.Company = ""
            $User.StreetAddress = ""
            $User.City = ""
            $User.State = "Colorado"
            $User.PostalCode = ""
            $User.HomePhone = "(000) 000-0000"
            $User.Fax = "(000) 000-0000"
            $User.Group1 = ""
            $User.Group2 = ""
            return $User

        }
    }
}
#Function to create users in esi.local
function createUserEsilocal([String]$Firstname, 
    [String]$MI,
    [String]$Lastname, 
    [String]$Username, 
    [String]$Password, 
    [String]$Office, 
    [String]$Department, 
    [String]$JobTitle, 
    [String]$Desc, 
    [String]$Name) {
        
    #Call method to set user office
    $User = setOffice $Office

    #Set user initials with X if description exists already
    if (Get-ADUser -Filter 'Description -like $Desc' -Server "IL-HADC.esi.local") {
        Write-Host "$Desc already exists defaulting to X."
        $Desc = ''
        $tempDesc = $Firstname + ' ' + 'X' + ' ' + $Lastname	
        $tempDesc.split(' ') | ForEach-Object { $Desc += $_[0] }
        $Desc = $Desc.ToUpper() 
        if (Get-ADUser -Filter 'Description -like $Desc' -Server "IL-HADC.esi.local") {
            Write-Host "$Desc already exists defaulting to Z."
            $Desc = ''
            $tempDesc = $Firstname + ' ' + 'Z' + ' ' + $Lastname	
            $tempDesc.split(' ') | ForEach-Object { $Desc += $_[0] }
            $Desc = $Desc.ToUpper()
        }
    }

    #Hashtable for new user
    $NUparams = @{
        Server            = "IL-HADC.esi.local" 
        SamAccountName    = $Username 
        UserPrincipalName = "$Username@engsys.net" 
        Name              = $Name 
        GivenName         = $Firstname 
        Initials          = $MI 
        Surname           = $Lastname 
        DisplayName       = $Name 
        Description       = $Desc
        EmailAddress      = "$Username@engsys.com" 
        Title             = $JobTitle 
        Department        = $Department 
        Company           = $User.Company 
        HomePage          = $User.Company 
        Office            = $Office 
        StreetAddress     = $User.StreetAddress 
        State             = $User.State 
        City              = $User.City 
        PostalCode        = $User.PostalCode 
        HomePhone         = $User.HomePhone 
        Fax               = $User.Fax 
        Enabled           = $True 
        Path              = $User.OU 
        AccountPassword   = (convertto-securestring $Password -AsPlainText -Force) 
    }
    
    
    New-ADUser @NUparams
		
    $Groups = @($User.Group1, $User.Group2)

    #Set group membership
    try {
        ForEach ($Group in $Groups) {
            Add-ADPrincipalGroupMembership -Identity $Username -MemberOf $Group
        }
    }
    catch {
        Write-Warning "Could not add $Username membership."
        Start-Sleep 1
    }
        
    Write-Output "$Username creation successful in esi.local."
    Start-Sleep 1
    Return
}	