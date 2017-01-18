$ExportPathName = 'C:\temp\office365-spaceusage.csv'

function Invoke-Office365TenantLogon {
 #### Pop-up a dialog for username and request your password
 $cred = Get-Credential
 #### Import the Local Microsoft Online PowerShell Module Cmdlets and Connect to O365 Online
 Import-Module MSOnline
 Connect-MsolService -Credential $cred
 #### Establish an Remote PowerShell Session to Exchange Online
 $msoExchangeURL = “https://ps.outlook.com/powershell/”
 $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $msoExchangeURL -Credential $cred -Authentication Basic -AllowRedirection
 Import-PSSession $session -AllowClobber
 }

function Invoke-Office365TenantLogoff {
 #### Remove the Remote PowerShell Session to Exchange Online ----
 Get-PsSession | Remove-PsSession
 #Remove-PsSession $session
 }

function Active-MessageCopyForSendItemsOnSharedMailbox {
    #List SharedMailbox and show MessageCopyForSentAsEnabled and MessageCopyForSendOnBehalfEnabled status
    $SharedMailboxList = Get-Mailbox | Where-Object {($_.RecipientTypeDetails) -eq 'SharedMailbox'}
    
    #Active MessageCopyForSentAsEnabled and MessageCopyForSendOnBehalfEnabled for all SharedMailbox with this parameter defined as 'False'
    $SharedMailboxList <#| Where-Object {($_.MessageCopyForSentAsEnabled) -eq 'False'}#> | Set-Mailbox -MessageCopyForSentAsEnabled $True
    $SharedMailboxList <#| Where-Object {($_.MessageCopyForSendOnBehalfEnabled) -eq 'False'}#> | Set-Mailbox -MessageCopyForSendOnBehalfEnabled $True
}

    
#Connect to Office 365 Tenant
Invoke-Office365TenantLogon

#List all mail accounts
Get-Mailbox -ResultSize Unlimited |
Get-MailboxStatistics |
Select DisplayName, `

@{name="TotalItemSize (MB)"; expression={[math]::Round( `
($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)}}, `
ItemCount |
Sort "TotalItemSize (MB)" -Descending |
Export-CSV -Path $ExportPathName -Encoding UTF8 -NoType -Force

#Disconnect to Office 365 Tenant
Invoke-Office365TenantLogoff
