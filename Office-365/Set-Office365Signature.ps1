$SigFile = "SigTemplate.htm"
$temp = Get-Content -Path $Sigfile -ReadCount 0

#Connect to Office 365 Tenant
$Cred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Cred -Authentication Basic –AllowRedirection
Import-PSSession $Session

#Import the active directory module which is needed for Get-ADUser
Import-Module ActiveDirectory

#Import active AD users with defined e-mail address
$users = Get-ADUser -filter * -Properties * | Where-Object {($_.Enabled) -eq "True" -AND ($_.EmailAddress -ne $null)}

#Add signature template to all selected AD users, and active AutoAddSignature on new e-mail
foreach ($user in $users) {
    $Office365ID =  $user.EmailAddress.Split("@")[0]
    Set-MailboxMessageConfiguration -identity $Office365ID -SignatureHtml $temp -AutoAddSignature $True    
    }
