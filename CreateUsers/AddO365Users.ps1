<# 
.SYNOPSIS 
Function to create user in O365

.DESCRIPTION
Callable function with required parameters to create and assign a user licenses in the O365 tenant. Includes setting the correct office information.

.AUTHOR
    Josh Razalan
#>
#Function to set various office information in O365 
function setOffice {
    param($Office)
    $User = "" | Select-Object -Property PhoneNumber, StreetAddress, City, State, PostalCode, Fax
    #Switch to set the correct info based on office location.
    switch ($Office) {
        "Colorado" {
            $User.StreetAddress = ""
            $User.City = ""
            $User.State = "Colorado"
            $User.PostalCode = ""
            $User.PhoneNumber = "(000) 000-0000"
            $User.Fax = "(000) 000-0000"
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
    
    #Hashtable for O365 properties
    $O365params = @{
        DisplayName         = $Name 
        FirstName           = $Firstname 
        LastName            = $Lastname 
        UserPrincipalName   = "$Username@domain.com" 
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

    #Check to see if user already exists if not, continue
    $O365Acc = Get-Msoluser -UserPrincipalName "$Username@domain.com" -ErrorAction SilentlyContinue 
    if ($O365Acc) { }
    else {
        New-MsolUser @O365params
        
        <# Commented out as this is called in the main program script
        Add necessary permissions for admins
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
        Import-PSSession $Session -DisableNameChecking
        Add-MailboxPermission -identity "$username" -user "Organization Management" -AccessRights FullAccess -inheritancetype all -automapping $false
        Remove-PSSession $Session#>

        #Set licenses
        if ($O365 -eq $true) {
            switch ($License) {
                "ProPlus" {
                    try {
                        Set-MsolUserLicense -UserPrincipalName "$Username@engsys.com" -AddLicenses "O365Tenant:EXCHANGESTANDARD" 
                        Set-MsolUserLicense -UserPrincipalName "$Username@engsys.com" -AddLicenses "O365Tenant:OFFICESUBSCRIPTION" 
                    }
                    catch {
                        Write-Host "Out of licenses."
                    }
                }
                "E3" {
                    try {
                        Set-MsolUserLicense -UserPrincipalName "$Username@engsys.com" -AddLicenses "O365Tenant:ENTERPRISEPACK" 
                    }
                    catch {
                        Write-Host "Out of licenses."
                    }
                }
                "E1" {
                    try {
                        Set-MsolUserLicense -UserPrincipalName "$Username@engsys.com" -AddLicenses "O365Tenant:STANDARDPACK" 
                    }
                    catch {
                        Write-Host "Out of licenses."
                    }
                }   
            }
        }
    
    }
}