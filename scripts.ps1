#connecting to exhcange online using -UserPrincipalName
Connect-EchangeOnline -UserPrincipalName #UserPrincipalName

#To disconnect
Remove-PSSession $Session

#Connecting to AzureAd
Connect-MsolService
Connect-AzureAD

#Command to change field in Exhcange Online
PS C:\> Import-Csv "csv-name.csv" | foreach { Set-User -Identity $_.user -Company $_.company} ($user_file | ForEach {Set-User $_.name -company $_.company})

#Address List
New-AddressList -Name <#Address List Name #> -ReceipientFilter {((<#filters here#>))}

#Dynamic Groups
#View Dynamic Distribution group members
$DDG = Get-DynamicDistributionGroup #Distribution group email address here
Get-Receipient -ResultSize Unlimited -ReceipientPriverFilter $DDG.ReceipientFilter | Select PrimarySmptAddress | export-csv "csv name"

#Edit Existing Dynamic Distirbution Group
Set-DynamicDistributionGroup -Identity "DDG Identity" -ConditionalCompany  "company" -IncludedRecipients MailboxUsers

#Distribution Group
Get-DistirbutionGroup #Type in address when prompted for identity
Get-DistirbutionGroup -Identity "identity here" | filters

#Get All users of a country
Get-User -Filter "countryoregion -eq 'contry here'"

#Gain admin Full Access
Get-Mailbox -ResultSize unlimited -Filter {(RecipientTypeDetails -eq 'UserMailbox') -and (Alias -ne 'Admin')} | Add-MailboxPermission -User "user here" -AccessRights fullaccess -InheritanceType all

#For Troubleshooting Mailbox and testing

Get-Mailbox -ResultSize limited -50 {(RecipientTypeDetails -eq 'UserMailbox') -and (Alias -ne 'Admin')} | Add-MailboxPermission -User "user here" -AccessRights fullaccess -InheritanceType all
Remove-MailboxPermission -Identity "identity here" -User "user here without quotes" -AccessRights FullAccess -InheritanceType All

Import-csv "csv location here" | foreach-object { Remove-MailboxPermission -Identity $_.email -User "user here without quotes" -AccessRights FullAccess -Confirm:$false }

$user_file = Import-CSV "csv location here without quotes"
>$user_file | ForEach {Set-user $_.name $_.company -Title}


#To add contacts from a file 
Import-Csv "csv location here" |% {New-MailContact -ExternalEmailAddress $_.ExternalEmailAddress -Name $_.Name -FirstName $_.FirstName -LastName $_.LastName}

#To hide contacts from the GAL
Get-MailContact | Set-MailContact -HiddenFromAddressListsEnabled $True

#Cut over migration
Get- MigrationBatch | Set-MigrationBatch  LargeItemLimit 100

#Set user license
Set-MsolUserLicense -userprincipalname $_.newupn -AddLicenses "license here"

#aPassword never expires
Import-Csv "csv location here" | foreach {Set-MsolUser -UserPrincipalName $_.NewUPN -PasswordNeverExpires $true}

#To change license in bulk. - First remove and then add new license.
Import-csv "csv location here" | foreach {Set-MsolUserLicense -UserPrincipalName $_.newupn -RemoveLicenses "license here"}
Import-Csv "csv location here" | foreach {Set-MsolUserLicense -userprincipalname $_.newupn -AddLicenses "license here" -RemoveLicenses "license here"}

#Import bulk users upload - Group accounts
Import-Csv "csv location here" | foreach{Add-DistributionGroupMember -Identity "identity here" -Member $_.alias}
Import-Csv "csv location here" | foreach{Add-DistributionGroupMember -Identity "identity here" -Member $_.userprincipalname}


#Import Bulk group accounts- Name,Type,Smtp,owner.
Import-CSV "csv location here" | foreach {New-DistributionGroup -Name $_.name -Type $_.Type -Primarysmtpaddress $_.primarysmtpaddress -Managedby $_.managedby}


#Delete Bulk Groups
Import-csv "csv location here" | foreach {Remove-DistributionGroup $_.RemoveDG}


#find Mailbox
Get-Mailbox -resultsize unlimited | where {$_.EmailAddresses -match "address here"} 
New-MsolUser -UserPrincipalName "user principal name here without quotes" -Password "Password here without quotes" -ForceChangePassword $false -DisplayName "display name here" 

#Get Email aliases
Get-Mailbox -identity "Identity here" | Select -Expand EmailAddresses Alias

#Get recipient where display is a mathc
Get-Recipient | where {$_.DisplayName -match "display name here"} | FL Name, DisplayName, EmailAddresses
Get-MSOLuser -User "user here without quotes" |Select -Expand Proxyaddresses 


#Get all users subscription type.

#NEW SCRIPT
Get-MSOLUser -all | % { $user=$_; $_.Licenses | Select {$user.Userprincipalname},AccountSKuid } | Export-CSV "csv location here" 
Get-MsolUser -all | Where-Object { $_.isLicensed -eq "TRUE" } | Select-Object UserPrincipalName, DisplayName, alias, Country, Department | Export-Csv "Csv location here without quotes"


#Export users via results
Get-Mailbox -ResultSize unlimited | Select -Expand EmailAddresses Alias | Export-Csv "CSV location here without quotes"
Get-mailbox -resultsize unlimited| Get-MailboxStatistics | select userprincipalname,displayname, lastlogontime | Export-Csv "Csv location here without quotes"

#Mailbox size
Get-mailbox -resultsize unlimited | get-mailboxstatistics | select Displayname,TotalItemSize | export-csv "Csv location here without quotes"

#User Export list
get-user -resultsize unlimited | select userprincipalname, company, CountryOrRegion | export-csv "csv location here"
get-user -resultsize unlimited | select userprincipalname,company,CountryOrRegion,title,department,office,phone,mobilephone | export-csv "csv location here"
Get-Recipient -ResultSize Unlimited | Export-Csv "Csv location here without quotes"
Get-Recipient -ResultSize Unlimited -Filter {(RecipientTypeDetails -eq "UserMailbox")} | Select-Object WindowsLiveID,Company,CountryOrRegion | Export-Csv -Path "Csv location here without quotes"

#User Export List - Lync
Get-Csonlineuser | Select-Object DisplayName, SipAddress | Export-Csv -Path "Csv location here without quotes"

#Bulk password change
Import-Csv "Csv location here without quotes"|%{Set-MsolUserPassword -userPrincipalName $_.UserPrincipalName -NewPassword $_.newpassword -ForceChangePassword $true}

#Use sample csv file otherwise it would not work.
Set-MsolUserPassword -UserPrincipalName "user principal name here" -ForceChangePassword $true
Import-Csv C:\password.csv|%{Set-MsolUserPassword  userPrincipalName $_.UserPrincipalName -NewPassword $_.newpassword -ForceChangePassword $false}
Import-Csv "Csv location here" | foreach { Set-MsolUserPassword -UserPrincipalName $_.UserPrincipalName -ForceChangePassword $true} > fcp_log.txt
Import-Csv "Csv location here" | foreach {Set-MsolUserPassword -UserPrincipalName $_.UserPrincipalName -NewPassword $_.Password} > fcp_log1.txt

#Bulk delete
Import-Csv "Csv location here" | ForEach-Object {Remove-MsolUser  UserPrincipalName $_.EmailAddress -force} - 30 days in dumpster

Import-Csv "Csv location here" | ForEach-Object {Remove-MsolUser  UserPrincipalName $_.EmailAddress -force -removefromrecyclebin} - force delete from dumpster

#Bulk Restore Deleted Accounts
Import-Csv "Csv location here"| ForEach-Object {Restore-MsolUser -UserPrincipalName $_.EmailAddress}


#To change from old upn to new upn
Import-Csv "C:\guy.csv" | foreach {Set-MsolUserPrincipalName -UserPrincipalName $_.OldUPN -NewUserPrincipalName $_.NewUPN} 


#change company
Import-Csv "Csv location here" | foreach {Set-User -Identity $_.NAME -Company $_.Company}
Import-Csv "Csv location here" | foreach {Set-User -Identity $_.NewUPN -Company $_.Company}
Import-Csv "Csv location here" | foreach {Set-User -Identity $_.NewUPN -title $_.title -department $_.department}

#Bulk add Custom Attributes.
Import-Csv "Csv location here" | foreach {Set-Mailbox -Identity $_.NewUPN -CustomAttribute1 $_.Customattribute1}

#Force deleted users
Set-ExecutionPolicy Unrestricted
Set-ExecutionPolicy RemoteSigned
Get-MsolUser -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin  Force




#SCRIPTS

#Adding Proxy(alias) addresses.
Set-executionpolicy remotesigned 

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection

Import-PSSession $Session 

Import-Csv C:\File.csv | ForEach-Object {
  $UserPrincipalName = $_.UserPrincipalName
  $EmailAddresses = $_.EmailAddresses
  Set-Mailbox -Identity $UserPrincipalName -EmailAddresses @{add= $EmailAddresses}
}


#Get 0365 user - OWA access only
Get-ConnectionByClientTypeDetailReport  | where {$_.UserName -eq "User"}

#Export 0365 OWA access ALL users
Get-O365ClientBrowserDetailReport | export-csv "c:\users.csv"

Search-Mailbox -SearchDumpsterOnly -Identity "identity here" -TargetMailbox "identity here" -TargetFolder RecoverData

#Recover items from the dumpster Chande Identity email , target mailbox and target folder to whatever you need.
$user_file = Import-CSV "Csv location here without quotes"
Import-CSV "Csv location here without quotes"| ForEach {Set-user $_.login_name -company $_.company}

get-user -identity "identity here" |  fl *company*

#Bulk Import External Contacts to Exchange.
Import-Csv "Csv location here without quotes"|%{New-MailContact -Name $_.Name -DisplayName $_.Name -ExternalEmailAddress $_.ExternalEmailAddress -FirstName $_.FirstName -LastName $_.LastName}

#Delete Emails by search Content Subject
New-ComplianceSearchAction -SearchName "Phish" -Purge -PurgeType SoftDelete
#check statues
Get-ComplianceSearchAction 

#Get status of password policy
Get-MsolPasswordPolicy  DomainName "domain name here"  

#to set password to never expire
Set-msoluser  -UserPrincipalName "principal name here" -PasswordNeverExpires $True 

#Password never expire in bulk
Import-Csv "Csv location here" | foreach {Set-MsolUser -UserPrincipalName $_.NewUPN -PasswordNeverExpires $true} 

#Bulk Disable user
Import-Csv "CSV location here"| ForEach {Set-MsolUser -UserPrincipalName $_.NewUPN -BlockCredential $true} 

#To determine if password is set to expire or not
Get-MGUser -UserId "user id here" | fl 

#To Determine if password is set to expire or not
Get-MsolUser -UserPrincipalName "user principal name here" | fl 

#Set all passwords to expire based on policy
Get-MsolUser -all | Set-MsolUser  PasswordNeverExpires $False

#Removing Forwards in bulk
Set-Mailbox  -Identity "mailbox identity here" -ForwardingsmtpAddress $null
Import-Csv "CSV location here" | foreach {Set-Mailbox  Identity $_.NewUPN -ForwardingsmtpAddress $null}

#Adding Forward in Bulk
Import-Csv "CSV location here" | foreach {Set-Mailbox  Identity $_.NewUPN -ForwardingsmtpAddress $_.Forward -DeliverToMailboxAndForward $true}

#Bulk Add Alias to one email Account only via csv

Import-Csv "CSV location here without quotes" | ForEach-Object {$name = $_.Name}
Import-Csv "CSV location here without quotes" | ForEach-Object {$proxy = $_.ProxyAddresses -split  ; }
Set-Mailbox -Identity $name -EmailAddresses @{add= $proxy}

Get-MobileDevice -ResultSize Unlimited | Select-Object @{Name='User';Expression={(Get-Mailbox -Identity $_.UserDisplayName) | Select-Object -expand WindowsEmailAddress}},DeviceID,DeviceType,DeviceUserAgent,DeviceModel | Export-Csv "CSV location here without quotes"


get-mailbox -resultsize unlimited| foreach {$user = Get-mailbox $_; get-activesyncdevicestatistics -mailbox $_.name -ea "SilentlyContinue" | where {$_.deviceid -ne $null} | select @{n="Name"; e={$user.Name}}, @{n="customattribute8"; e={$user.CustomAttribute8}}, DeviceID, Devicetype, LastSuccessSync} | export-csv "CSV location here without quotes"

Add-MailboxPermission -Identity "identity here" -User "user here" -AccessRights FullAccess -InheritanceType All -AutoMapping $true

Search-Mailbox "mailbox here" -EstimateResultOnly -SearchQuery {kind:email OR kind:meetings}

Search-Mailbox "mailbox here" -DeleteContent -SearchQuery {kind:email OR kind:meetings}

Get-MobileDevice |  select-Object @{Name='User';Expression={(Get-Mailbox -Identity $_.UserDisplayName) | export-csv "Csv location here"

#Disable users in bulk
Import-Csv "CSV location here without quotes" | ForEach {Set-MsolUser -UserPrincipalName $_ -BlockCredential $true}

#Set Automated response in bulk message with out csv.
Import-Csv "CSV location here without quotes" | ForEach {Set-MailboxAutoReplyConfiguration -Identity $_.NewUPN -ExternalAudience All -InternalMessage "Thank you" -ExternalMessage "Thank you" -AutoReplyState enabled}

#Set Auto response with message in csv
Import-Csv "CSV location here without quotes" | ForEach {Set-MailboxAutoReplyConfiguration -Identity $_.NewUPN -ExternalAudience All -InternalMessage $_.Internal -ExternalMessage $_.External -AutoReplyState enabled}

#Last user log on
#Open AD Powershell (MSonline)
Set-ExecutionPolicy RemoteSigned 
<#If prompted that cmdlets not recognized change directory to location of script\lastUserActivityTimeReport.ps1 - scriptname to run creds will be asked for to run report.#>

#Upgrade users in bulk to Teams

Import-Csv "CSV location here without quotes" | ForEach {Grant-CsTeamsUpgradePolicy -Identity $_.UserPrincipalName -PolicyName "UpgradeToTeams"}

Set-CSTeamsMeetingPolicy -Identity Global -AllowEngagementReport "Enabled" -Download participants

get-mailboxLocation  -user "user name here" | fl mailboxGuid,mailboxLocationType

Start-ManagedFolderAssistant -identity "identity here"

Get-Mailbox "mailbox here" | Select *quota

Set-Mailbox -Identity "identity here" -RetainDeletedItemsFor 0

Set-Mailbox -Identity - "identity here" -RetainDeletedItemsFor 0

Get-MailboxFolderStatistics -Identity "identity here"

#Increase Mailbox size
Set-Mailbox "mailbox here" -ProhibitSendQuota 99Gb -ProhibitSendReceiveQuota 99Gb -IssueWarningQuota 90Gb 

#Check if there are any holds on mailbox
Get-Mailbox "mailbox here" | FL DelayHoldApplied,DelayReleaseHoldApplied 

Get-Mailbox "mailbox here" | FL *HoldApplied*

Export-MailboxDiagnosticLogs -Identity "identity here" -ComponentName HoldTracking

Set-Mailbox "malbox here" -RemoveDelayHoldApplied

Set-Mailbox "mailbox here" -RemoveDelayReleaseHoldApplied


#Alias in bulk
Import-CSV "CSV location here"| ForEach {Set-Mailbox $_.Mailbox -EmailAddresses @{add=$_.NewEmailAddress}}

#Csv - Mailbox | Newemailaddress
Get-MailboxFolderStatistics "mailbox here" -FolderScope RecoverableItems | FL Name,FolderAndSubfolderSize,ItemsInFolderAndSubfolders


#MFA Status
Get-MsolUser -all | select DisplayName,UserPrincipalName,@{N="MFA Status"; E={ if( $_.StrongAuthenticationRequirements.State -ne $null){ $_.StrongAuthenticationRequirements.State} else { "Disabled"}}} | Export-csv C:\0365\mfa1.csv

Get-Mailbox "mailbox here" |select Name,ExchangeGuid

Get-MsolRole | %{$role = $_.name; Get-MsolRoleMember -RoleObjectId $_.objectid} | select @{Name="Role"; Expression = {$role}}, DisplayName, EmailAddress | export-csv "CSV location here"

#Get all domains
Get-MsolDomain | export-csv "csv location here"

Set-MailboxRegionalConfiguration -Identity "identity here" -TimeZone "Atlantic Standard Time"

Get-MailboxRegionalConfiguration -Identity "identity here"

$Users = Get-Mailbox -ResultSize unlimited -Filter {(RecipientTypeDetails -eq 'RoomMailbox')} $users | %{Set-MailboxRegionalConfiguration $_.Identity -TimeZone "Atlantic Standard Time" }

#Mailbox auto Expanding Archive
Enable-Mailbox "mail box here" -AutoExpandingArchive

#Check to see if mailbox archiving is enabled
Get-Mailbox "mailbox here" | FL AutoExpandingArchiveEnabled

