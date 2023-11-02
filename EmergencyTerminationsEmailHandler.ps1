$Regex = "EmployeeNumber\:.*(?<empnumber>100\d{6})"
$Regex2 = "Email\:\s*(?<email>[^\<\s]+)"
$Regex3 = "SAMAccountName\:\s*(?<SAMAccountName>[^\<\s]+)"
$InboxID = <Inbox ID of automation mailbox>

#Declare Microsoft Graph authentication information
$myTenantId = <Tenant ID>
$clientID = <Client ID>
$clientSecret = Get-Content <Location of client secret secure string> | ConvertTo-SecureString

#Get Access token
$myToken = Get-MsalToken -clientID $clientID -clientSecret $clientSecret -tenantID $myTenantId
$AccessToken = $MyToken.AccessToken | ConvertTo-SecureString -AsPlainText -Force

#Connect to MGGraph
Connect-MgGraph -AccessToken $AccessToken >> $Null


$Messages = Get-MGUserMailFolderMessage -UserID <Automation Email Address> -MailFolderId $InboxID

ForEach($Message in $Messages) {
    $Matches = $Null
    $Email = $Message.body.Content
    $Email -match $Regex >> $Null
    $EmployeeNumber = $Matches.empnumber
    $Email -match $Regex2 >> $Null
    $EmailAddress = $Matches.email
    $Email -match $Regex3 >> $Null
    $SAMAccountName = $Matches.SAMAccountName
    $Sender = $Message.sender.EmailAddress.address
    $ReplyAddresses  = @("$Sender",<Admin Distribution List>)
    $Recipient = $ReplyAddresses | % {@{emailAddress = @{ address = $_ }}}
    $MessageBody = $Message.Body.Content


    If($EmployeeNumber) {
        & 'E:\PowerShell\Scheduled Task\Emergency_Term_Auto.ps1' -EmployeeNum $EmployeeNumber -MessageID $Message.id -MailDelegate
        Remove-MGUserMessage -MessageId $Message.Id -UserID <Automation Email Address>
    }
    ElseIf($EmailAddress) {
        & 'E:\PowerShell\Scheduled Task\Emergency_Term_Auto.ps1' -Email $EmailAddress -MessageID $Message.id -MailDelegate
        Remove-MGUserMessage -MessageId $Message.Id -UserID <Automation Email Address>
    }
    ElseIf($SAMAccountName) {
        & 'E:\PowerShell\Scheduled Task\Emergency_Term_Auto.ps1' -SAMAccountName $SAMAccountName -MessageID $Message.id -MailDelegate
        Remove-MGUserMessage -MessageId $Message.Id -UserID <Automation Email Address>
    }
    Else {

        $Body = "<p>Hi Team,</p>"
        $Body += "<p></P>"
        $Body += "<p>We were unable to properly parse the user data from this message. Please review the formatting of the message and resubmit.</p>"
        $Body += "<p></P>"
        $Body += "<p>Thank you</p>"
        $Body += "<p></p>"
        $Body += "The UT Systems Administration Team"
        $Body += "<p> </p>"
        $Body += '<div><div style="border:none; border-top:solid #E1E1E1 1.0pt; padding:4.0pt 0in 0in 0in">'
        $Body += "<p> </p>"
        $Body += $MessageBody
        $Body += "</div>"

        $Params = @{
            "Message" = @{
                "Body" = @{
                    "Content" = $Body
                    "ContentType" = "HTML"
                }
                "ToRecipients" = $Recipient
                "Importance" = "High"
            }
        }

        Invoke-MgReplyAllUserMessage -UserID <Automation Email Address> -MessageID $Message.Id -BodyParameter $Params
        Remove-MGUserMessage -MessageId $Message.Id -UserID <Automation Email Address>
    }
}
