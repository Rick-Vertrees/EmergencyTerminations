[CmdletBinding(DefaultParameterSetName = 'EmployeeNumber')]
param(
    #EmployeeNumber
    [Parameter(Mandatory=$true,Position=0, ParameterSetName='EmployeeNumber')]
    [string]$EmployeeNum,

    #SAMAccountName
    [Parameter(Mandatory=$true,Position=0, ParameterSetName='SAMAccountName')]
    [string]$SAMAccountName,

    #UPN
    [Parameter(Mandatory=$true,Position=0, ParameterSetName='Email')]
    [string]$Email,

    #Mail Switch
    [Parameter(ParameterSetName = 'EmployeeNumber')]
    [Parameter(ParameterSetName = 'SAMAccountName')]
    [Parameter(ParameterSetName = 'Email')]
    [switch]$MailDelegate,

    #MessageID
    [Parameter(ParameterSetName = 'EmployeeNumber')]
    [Parameter(ParameterSetName = 'SAMAccountName')]
    [Parameter(ParameterSetName = 'Email')]
    $MessageID
)

function GeneratePassPhrase {
    $RandomPassPhrase = -join ((33..90) + (97..122) | Get-Random -Count 12 | % {[char]$_})
    return $RandomPassPhrase
}

If($EmployeeNum) {
    $ADUser = Get-ADUser -Filter {EmployeeNumber -eq $EmployeeNum} -Properties *
    $identifier = $EmployeeNumber
}

If($SAMAccountName) {
    $ADUser = Get-ADUser -Filter {SAMAccountName -eq $SAMAccountName} -Properties *
    $Identifier = $SAMAccountName
}

If($Email) {
    $ADUser = Get-ADUser -Filter {Mail -eq $Email} -Properties *
    $Identifier = $Email
}

$StrToday = Get-Date -Format yyyyMMdd
$UserEmail = $ADUser.mail
$UserUPN = $ADUser.UserPrincipalName
$ManagerDN = $ADUser.manager
$Manager = Get-ADUser -Filter {DistinguishedName -eq $ManagerDN} -Properties *
$ManagerUPN = $Manager.UserPrincipalName
[String]$StrToday = Get-Date -Format yyyyMMdd
$Domain = (Get-ADDomain).name
$nPass = GeneratePassPhrase
$sPass = ConvertTo-SecureString -String $nPass -AsPlainText -force
$DisplayName = $ADUser.DisplayName
$Regex = "^(?<URL>.*)/[^/]*$"
$Regex2 = "^CN=(?<Name>[^\,]*),(?<OU>.*)$"
$Regex3 = "^(?<SIDPrefix>.*)-(?<SIDPostfix>[^\-]*)$"
$InboxID = <Automation mailbox inbox ID>

#Declare Microsoft Graph authentication information
$myTenantId = <Tenant ID>
$clientID = <Client ID>
$clientSecret = Get-Content <Location of client secret secure string> | ConvertTo-SecureString

#Get Graph Access token
$myToken = Get-MsalToken -clientID $clientID -clientSecret $clientSecret -tenantID $myTenantId
$AccessToken = $MyToken.AccessToken | ConvertTo-SecureString -AsPlainText -Force
$Headers = @{Authorization = "Bearer $($myToken.AccessToken)"}

#Get Exchange Access token
$myTokenExchange = Get-MsalToken -clientID $clientID -clientSecret $clientSecret -tenantID $myTenantId -Scopes "https://outlook.office365.com/.default"

#Get Drive ID for Active Direcotry Engineers Sharepoint Site
$url = "https://graph.microsoft.com/v1.0/sites/<Tenant Sharepoint URL>:\sites\<Sharepoint Site Name>:\drive"
$Global:driveID = Invoke-RestMethod -Uri $url -Headers $Headers | Select-Object ID -ExpandProperty ID

#Connect to MGGraph and Exchange
Connect-MgGraph -AccessToken $AccessToken >> $Null
Connect-ExchangeOnline -AccessToken $myTokenExchange.AccessToken -Organization <Tenant onmicrosoft url>

$UserID = (Get-MgUser -Filter "UserPrincipalName eq '$UserUPN'").Id

$ADUser.DistinguishedName -Match $Regex2 > $Null
$OU = $Matches.OU
$LogText = ""
$Filename = $StrToday + "_" + $ADUser.GivenName + $ADUser.Surname + "_" + $Domain + ".txt"
$FileLocation = "E:\Reports\User_Disable\" + $FileName
$LogText += "=========================Terminated User Info========================= `r`n"
$LogText += "Name: " + $ADUser.Displayname + " `r`n"
$LogText += "UPN: " + $ADUser.UserPrincipalname + " `r`n"
$LogText += "SAMAccountName: " + $ADUser.SAMAccountName + " `r`n"
$LogText += "OU: " + $Matches.OU + " `r`n"
$LogText += "Manager: " + $ManagerDN + " `r`n"
$LogText += " `r`n"
$LogText += "=========================Removed Group Info========================= `r`n"

$GroupIds = ""
$Groups = $ADUser.MemberOf

ForEach($Group in $Groups) {
    If(!($Group -match $Regex4)) {
        $ADGroup = Get-ADGroup -Filter {distinguishedName -eq $Group} -Properties ObjectSid
        $ADGroup.ObjectSid -Match $Regex3 > $Null
        $GroupID = [string]$Matches.SIDPostfix
        $GroupIDs = $GroupIDs+$GroupID+","
        $LogText += "Group: " + $ADGroup.Name + " `r`n"
        $ADGroup | Remove-ADGroupMember -members $ADUser -Confirm:$false
    }
}

If($GroupIDs) {
    $GroupIDs = $GroupIDs.Substring(0,$GroupIds.Length-1)
    $ADUser | Set-ADUser -Replace @{ExtensionAttribute10= $GroupIDs}
}

If($ManagerDN) {
    $ADUser | Set-ADUser -Replace @{ExtensionAttribute9= $ManagerDN}
}

[String]$UserDescription = $ADUser.Description
$NewDescription = $StrToday +" | Terminated | " + $UserDescription
$ADUser | Set-ADUser -Description $NewDescription -Confirm:$False

$ADUser | Set-ADUser -Clear Manager
$ADUser | Set-ADUser -Add @{ExtensionAttribute11= $OU}
$ADUser | Set-ADUser -Replace @{ExtensionAttribute12= $StrToday}
$ADUser | Set-ADUser -replace @{msExchHideFromAddressLists= $True}
$ADUser | Set-ADAccountPassword -Reset -NewPassword $sPass
$ADUser | Disable-ADAccount
$ADUser | Move-ADObject -TargetPath <Terminated User OU DN>

$LogText | Out-File $FileLocation

$url2 = "https://graph.microsoft.com/v1.0/drives/$driveID/items/root:<Folder Path>\$($Filename):/content"
$upload = Invoke-RestMethod -Uri $url2 -Headers $Headers -Method Put -InFile $FileLocation -ContentType 'text/plain'

Remove-Item -Path $FileLocation -Force

#Remove Mobile Devices
#List Devices
Get-MobileDevice -Mailbox $UserEmail | select Identity, DeviceOS, DeviceType, Guid | FT
#remove Devices and terminate ActiveSync Permissions
Get-MobileDevice -mailbox $UserEmail | select Guid, Identity, DeviceID | %{Clear-MobileDevice -Identity $_.Identity -AccountOnly -NotificationEmailAddresses ba_ops@radpartners.com -confirm:$false}
#RevokeTokens
Revoke-MgUserSignInSession -UserId $UserID
 
#Delegate Mailbox
If($MailDelegate) {
    $CustodianEMail = $Manager.mail
    Add-mailboxpermission –Identity $UserUPN –User $CustodianEmail –accessrights Fullaccess, readpermission  –inheritancetype All  –Automapping:$True
}

#Add to restricted senders list
Add-DistributionGroupMember -Identity <DL address that restricts senders> -Member $UserEmail

$Message = $Messages = Get-MGUserMailFolderMessage -UserID <Automation Email address> -MessageID $MessageID -MailFolderId $InboxID

$Sender = $Message.sender.EmailAddress.address
$ReplyAddresses  = @("$Sender",<Admin Distribution list address>)
$Recipient = $ReplyAddresses | % {@{emailAddress = @{ address = $_ }}}
$MessageBody = $Message.Body.Content

$Body = "<p>Hi Team,</p>"
$Body += "<p></P>"
$Body += "<p>The user $DisplayName has been terminated using the identifier $Identifier.</p>"
$Body += "<p></P>"
$Body += "<p>Thank you</p>"
$Body += "<p></p>"
$Body += "The UT Systems Administration Team"
$Body += "<p></p>"
$Body += '<div><div style="border:none; border-top:solid #E1E1E1 1.0pt; padding:3.0pt 0in 0in 0in">'
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

Invoke-MgReplyAllUserMessage -UserID <Automation Email Address> -MessageID $MessageID -BodyParameter $Params
