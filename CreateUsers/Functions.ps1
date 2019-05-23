function setOffice {
    param($Office)
    $User = "" | Select-Object -Property OU, Company, StreetAddress, City, State, PostalCode, HomePhone, Fax, Group1, Group2
    switch ($Office) {
        "Colorado" {
            $User.OU = "OU=Domain Users,OU=Colorado,DC=,DC="
            $User.Company = ""
            $User.StreetAddress = ""
            $User.City = "Colorado"
            $User.State = "Colorado"
            $User.PostalCode = ""
            $User.HomePhone = ""
            $User.Fax = ""
            $User.Group1 = "CN=Colorado Users,OU=Security Groups,OU=Colorado,DC=,DC="
            $User.Group2 = "CN=Colorado Projects,OU=Security Groups,OU=Colorado,DC=,DC="
            return $User

        }
    }
}
#Function to create users
function createUser([String]$Firstname, 
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
    if (Get-ADUser -Filter 'Description -like $Desc') {
        Write-Host "$Desc already exists defaulting to X."
        $Desc = ''
        $tempDesc = $Firstname + ' ' + 'X' + ' ' + $Lastname	
        $tempDesc.split(' ') | ForEach-Object {$Desc += $_[0]}
        $Desc = $Desc.ToUpper() 
        #Set user initials with Z if description exists already
        if (Get-ADUser -Filter 'Description -like $Desc') {
            Write-Host "$Desc already exists defaulting to Z."
            $Desc = ''
            $tempDesc = $Firstname + ' ' + 'Z' + ' ' + $Lastname	
            $tempDesc.split(' ') | ForEach-Object {$Desc += $_[0]}
            $Desc = $Desc.ToUpper()
        }
    }

    #Hashtable for User params
    $NUparams = @{
        Server            = "" 
        SamAccountName    = $Username 
        UserPrincipalName = "$Username@" 
        Name              = $Name 
        GivenName         = $Firstname 
        Initials          = $MI 
        Surname           = $Lastname 
        DisplayName       = $Name 
        Description       = $Desc 
        EmailAddress      = "$Username@" 
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

    #Set Users
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
		
    Write-Output "$Username creation successful."
    Start-Sleep 1
    Return
}	