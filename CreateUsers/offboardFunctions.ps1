<# 
.SYNOPSIS 
Function to call for offboarding users

.DESCRIPTION
From the CreateUsers GUI form, call this function to disable and move account to terminated users OU in AD. Block signing, remove licenses, remove from distrbution lists, 
add manager full access to mailbox, and reset password in O365. Finally, revoke VPN cert if necessary. 

.AUTHOR
    Josh Razalan
#>

function terminateUser {
    Disable-ADAccount -Identity $offboardUser

    try {
        Disable-ADAccount -server "il-hadc.esi.local" -Identity $offboardUser

    }
    catch {
        
    }

    #Get User dn and convert it to parent ou to move user to term OU
    $user = Get-ADUser $offboardUser 
    $DN = $user.DistinguishedName
    $OU = $DN -replace "^(.*?)\,", ""
    $OU = $OU -replace "Domain Users", "Terminated Users"
    Try {
        Move-ADObject -Identity $DN -TargetPath $OU
    }
    Catch {
        Write-Warning "User not found in traditional OU"
    }

    Try {
        Move-ADObject -Server "otherdomain" -Identity $DN -TargetPath $OU
    }
    Catch {
        Write-Warning "User not found in traditional OU"
    }


    #Connect to O365 service to remove licenses, block sign in, reset password, and set manager permissions
    Get-MsolDomain -ErrorAction SilentlyContinue
    if ($?) { }
    else {
        Connect-MsolService
    }
    $upn = "$offboardUser@domain.com"
    $Managerupn = "$offboardManager@domain.com"
    $ar = "owner"
    $O365User = Get-Msoluser -UserPrincipalName $upn -ErrorAction SilentlyContinue

    #Remove all licenses
    $O365User.licenses.AccountSkuId | ForEach-Object {
        Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses $_
    }

    #Call function to add access rights to terminated user's mailbox
    . \\Path\to\MailboxPermissions.ps1
    addMailboxPermissions $upn $Managerupn $ar

    #Block sign in and set terminated password.
    Set-MsolUser -UserPrincipalName $upn -BlockCredential $true
    Set-MsolUserPassword -UserPrincipalName $upn -NewPassword "Termpassword!"


    #Generate CSV for list of issued certificates to grab serial numbers of terminated user
    $CSVCreateSB = {
        certutil -view -out "RequestID,SerialNumber,RequesterName,RequestType,NotAfter,CommonName,Certificate Template" csv > "C:\CertTemp\tempcerts.csv"
    }
    Invoke-Command -ComputerName "Certificate Authority Computer" -ScriptBlock $CSVCreateSB
    
    #Grab a list of issued certs for terminated user and revoke through certutil
    $CSV = Import-Csv -Path "\\Path\to\tempcerts.csv"
    $IssuedCerts = $CSV | Where-Object { $_."Requester Name" -like "*$offboardUser" -and $_."Certificate Template" -like "*VPN*" }
    
    foreach ($cert in $IssuedCerts) {
        $certserial = $cert."Serial Number"
        $SB = {
            certutil -revoke $using:certserial 5
        }
        Invoke-Command -ComputerName "Certificate Authority Computer" -ScriptBlock $SB
    }
    Remove-Item -Path "\\Path\to\tempcerts.csv" -Force
    
}