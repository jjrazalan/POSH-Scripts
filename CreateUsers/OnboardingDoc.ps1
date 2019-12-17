function SendEmail {
    #Set variables for email
    $fromaddress = "donotreply@engsys.com"
    $toaddress = "daboyd@engsys.com"
    $subject = "$Username IT Onboarding"
    $body = ""
    $attachment = $Path + "\$Username IT Onboarding.pdf"
    $smtpserver = "esifl.mail.protection.outlook.com"

    $message = new-object System.Net.Mail.MailMessage 
    $message.From = $fromaddress 
    $message.To.Add($toaddress) 
    $message.IsBodyHtml = $True 
    $message.Subject = $Subject 
    $attach = new-object Net.Mail.Attachment($attachment) 
    $message.Attachments.Add($attach) 
    $message.body = $body 
    $smtp = new-object Net.Mail.SmtpClient($smtpserver) 
    $smtp.Send($message) 

}
function CreateOnboardingDoc {
    #Set correct variables
    $Initialslower = $Initials.ToLower()
    $dbpw = $Initialslower + "1"
    $email = $username + "@engsys.com"
    $Path = "P:\Global\IT\New Hires"

    #Create a copy of the template so it doesn't overwrite itself
    Copy-Item -Path "$Path\OnboardingTemplate.docx" -Destination "$Path\$username IT Onboarding.docx"

    #Open word object to find and replace
    $objWord = New-Object -comobject Word.Application  
    $objWord.Visible = $false

    #Open the onboarding template 
    $objDoc = $objWord.Documents.Open("$Path\$username IT Onboarding.docx")
    $objSelection = $objWord.Selection

    #Create CSV of the find and replace variables
    $FindText = @"
    FIND,REPLACE
    tempusername,$username
    tempemail,$email
    temppassword,$password
    tdb,$Initialslower
    tdbpw,$dbpw
"@
    #convert from csv and loop through each object to replace words
    $FindTextObjs = ConvertFrom-Csv $FindText
    foreach ($FindTextObj in $FindTextObjs) {
        $a = $objSelection.Find.Execute($($FindTextObj.FIND), $false, $true, $False, $False, $False, $true, 1, $False, $($FindTextObj.REPLACE))
    }
    
    #Save as a new pdf file
    $Savefile = $Path + "\$Username IT Onboarding.pdf"
    $objDoc.SaveAS([ref] $Savefile, [ref] 17)
    $objDoc.Close()

    #Delete temp template from the folder
    Remove-Item -Path "$Path\$username IT Onboarding.docx" -Force

    #Call function to email Debbie
    SendEmail $username $Path
    
}