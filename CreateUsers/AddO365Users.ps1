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

function CreateO365User ([String]$Firstname, 
    [String]$Lastname, 
    [String]$Username, 
    [String]$Password, 
    [String]$Name,
    [Boolean]$O365,
    [String]$License,
    [String]$Office,
    [String]$Department,
    [String]$Jobtitle) {

    $User = setOffice $Office
    
    #Hashtable for O365
    $O365params = @{
        DisplayName         = $Name 
        FirstName           = $Firstname 
        LastName            = $Lastname 
        UserPrincipalName   = "$Username@" 
        Department          = $Department
        Title               = $Jobtitle
        Office              = $Office
        StreetAddress       = $User.StreetAddress
        City                = $User.City
        State               = $User.State
        PostalCode          = $User.PostalCode
        PhoneNumber         = $User.PhoneNumber
        Fax                 = $User.Fax
        ForceChangePassword = $false 
        Password            = $Password 
        UsageLocation       = "US"
    }

    $O365Acc = Get-Msoluser -UserPrincipalName "$Username@" -ErrorAction SilentlyContinue 
    if ($O365Acc) { }
    else {
        New-MsolUser @O365params
    }

    #Set licenses
    if ($O365 -eq $true) {
        switch ($License) {
            #Replace License with your O365 site
            "ProPlus" {
                try {
                    Set-MsolUserLicense -UserPrincipalName "$Username@" -AddLicenses "Licese:EXCHANGESTANDARD" 
                    Set-MsolUserLicense -UserPrincipalName "$Username@" -AddLicenses "Licese:OFFICESUBSCRIPTION" 
                }
                catch {
                    Write-Host "Out of licenses."
                }
            }
            "E3" {
                try {
                    Set-MsolUserLicense -UserPrincipalName "$Username@" -AddLicenses "Licese:ENTERPRISEPACK" 
                }
                catch {
                    Write-Host "Out of licenses."
                }
            }
            "E1" {
                try {
                    Set-MsolUserLicense -UserPrincipalName "$Username@" -AddLicenses "Licese:STANDARDPACK" 
                }
                catch {
                    Write-Host "Out of licenses."
                }
            }   
        }
    }
}