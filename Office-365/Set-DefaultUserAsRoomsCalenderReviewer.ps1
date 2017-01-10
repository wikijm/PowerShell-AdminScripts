# Work in progress for this script, be careful with it!

function Connect-EXOnline
{
	#Define URL to contact Office 365
	$Office365URL = "https://ps.outlook.com/powershell"
	
	#Imports the installed Azure Active Directory module.
	Import-Module MSOnline
	
	#Capture administrative credential for future connections.
	$Office365Credentials = Get-Credential -Message "Enter your Office 365 admin credentials"
	
	#Establishes Online Services connection to Office 365 Management Layer.
	Connect-MsolService -Credential $Office365Credentials
	
	#Creates an Exchange Online session using defined credential.
	$EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Office365URL -Credential $Office365Credentials -Authentication Basic -AllowRedirection -Name "Exchange Online"
	
	#This imports the Office 365 session into your active Shell.
	Import-PSSession $EXOSession
}

function Disconnect-EXOnline
{
	Remove-PSSession -Name "Exchange Online"
}

# Disclaimer
$Disclaimer = [Windows.Forms.MessageBox]::Show(
	"
This script aims to set Exchange/Office 365 users as Room mailbox calendar's reviewer.

									/!\ CAUTION /!\
If you are not sure of the type of action to carry on, or the impact on the mailing service, please stop this script now.
Do you wish to proceed?
", 'Set Exchange/Office 365 users as Room mailbox calendar's reviewer', 1, [Windows.Forms.MessageBoxIcon]::Question)
If ($Disclaimer -eq "OK")
{
	Write-Information 'Please wait, work in progress...'
}
Else
{
	Write-Error -Message 'Opération annulée'
	[Windows.Forms.MessageBox]::Show("Le script n'est pas en mesure de continuer. Opération stoppée.", 'Opération stoppée', 0, [Windows.Forms.MessageBoxIcon]::Error)
	Stop-TranscriptOnLog
	Exit
}



# Start a connection to Office 365
Connect-EXOnline

# List all RoomMailbox
$rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox

# Set Exchange/Office 365 users as Room mailbox calendar's reviewer
&rooms | %{Set-MailboxFolderPermission $_":\Calendar" -User Default -AccessRights Reviewer}

[Windows.Forms.MessageBox]::Show(
	"Action menée avec succès.
", 'Boîte partagée', 0, [Windows.Forms.MessageBoxIcon]::Information)


# Stop the connection to Office 365
Disconnect-EXOnline
