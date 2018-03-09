
## Connect and import exchange commands into PS ##

#$UserCredential = Get-Credential
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<server_IP or name>/PowerShell/ -Authentication Kerberos -Credential $UserCredential
#Import-PSSession $Session

## Delete emails from one mailbox
#Search-Mailbox -Identity albert -SearchQuery 'subject:Low memory AND received>07/18/2017' -DeleteContent -Force

## Delete emails from one mailbox with complete reporting
#Search-Mailbox -Identity brian -SearchQuery 'subject:VPN connection status changed AND received>07/15/2017' -DeleteContent -Force -LogLevel full -TargetMailbox armin@kaganonline.com -TargetFolder 'Search Query Logs'

## Delete email from multiple mailboxes
#foreach ($mailbox in (get-mailbox)) {Search-Mailbox -id $mailbox -SearchQuery '#Your query#' -DeleteContent -Force}

## Delete email from multiple mailboxes with full reporting
#foreach ($mailbox in (get-mailbox)) {Search-Mailbox -id $mailbox -SearchQuery '#Your query#' -DeleteContent -Force -LogLevel full -TargetMailbox armin@kaganonline.com -TargetFolder 'Search Query Logs'}

## Example
#foreach ($mailbox in (get-mailbox)) {Search-Mailbox -Identity $mailbox.Alias -SearchQuery 'subject:2016 Kagan Holiday Party AND received>08/01/2017 kind:meetings' -DeleteContent -Force -LogLevel full -TargetMailbox armin@kaganonline.com -TargetFolder 'Search Query Logs'}
