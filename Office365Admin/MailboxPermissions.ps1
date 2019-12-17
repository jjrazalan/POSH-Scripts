function addMailboxPermissions {
    param (
        [String]$Mailbox,
        [String]$User,
        [String]$AR
    )
    
    #Array for mailbox folder exclusions.
    $exclusions = @("Root",
    "Recoverableitemsroot",
    "Audits",
    "CalendarLogging",
    "RecoverableItemsDeletions",
    "RecoverableItemspurges",
    "RecoverableItemsversions")

    #Get list of user folders to add permissions to, but exclude those in exclusions
    $Folders = Get-MailboxFolderStatistics $Mailbox | Where-Object { $_.FolderType -notin $exclusions }

    #Set top level permission for user
    Add-MailboxFolderPermission $Mailbox -User $User -AccessRights $AR

    #Recursively set permissions on mailbox folders
    Foreach ($Folder in $Folders) {
        $FolderPath = $Folder.FolderPath.Replace("/", "\").Replace([char]63743, "/") 
        $MailboxFolder = "$Mailbox`:$FolderPath"
        Add-MailboxFolderPermission "$MailboxFolder" -User $User -AccessRights $AR
    }
}

#Connect to exchange online service
$getChildItemSplat = @{
    Path        = "$Env:LOCALAPPDATA\Apps\2.0\*\CreateExoPSSession.ps1"
    Recurse     = $true
    ErrorAction = 'SilentlyContinue'
    Verbose     = $false
}

$MFAExchangeModule = ((Get-ChildItem @getChildItemSplat | Select-Object -ExpandProperty Target -First 1).Replace("CreateExoPSSession.ps1", ""))
. "$MFAExchangeModule\CreateExoPSSession.ps1"
Connect-EXOPSSession -UserPrincipalName "$env:username@engsys.com"