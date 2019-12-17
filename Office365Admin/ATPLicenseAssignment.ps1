<#get-msoluser -all |
    Where-Object { (($_.licenses).accountskuID -like "*ESIFL:ENTERPRISEPACK*") -or (($_.licenses).accountskuID -like "*ESIFL:EXCHANGESTANDARD*") } |
    Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses "ESIFL:ATP_ENTERPRISE"#>

$users = get-msoluser -all | Where-Object {(($_.licenses).accountskuID -like "*ESIFL:ENTERPRISEPACK*") -or (($_.licenses).accountskuID -like "*ESIFL:EXCHANGESTANDARD*")}
Write-Host "Assigning to " $users.count " users."
ForEach($user in $users){
    Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses "ESIFL:ATP_ENTERPRISE"
}
