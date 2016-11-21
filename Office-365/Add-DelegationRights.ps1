# Work in progress for this script, be careful with it!

Add-Type -AssemblyName System.Windows.Forms

$PathDesktop = 'C:' + $env:HOMEPATH + '\Desktop'

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

function Select-FileDialog
{
	[CmdletBinding()]
	param ([string]$Title,
		[string]$Filter = 'All files *.*|*.*')
	Add-Type -AssemblyName System.Windows.Forms | Out-Null
	$fileDialogBox = New-Object -TypeName Windows.Forms.OpenFileDialog
	$fileDialogBox.ShowHelp = $false
	$fileDialogBox.initialDirectory = $PathDesktop
	$fileDialogBox.filter = $Filter
	$fileDialogBox.Title = $Title
	$Show = $fileDialogBox.ShowDialog()
	
	if ($Show -eq 'OK')
	{
		Return $fileDialogBox.FileName
	}
	Else
	{
		Write-Error -Message 'Opération annulée'
		[Windows.Forms.MessageBox]::Show("Le script n'est pas en mesure de continuer. Opération stoppée.", 'Opération stoppée', 0, [Windows.Forms.MessageBoxIcon]::Error)
		Stop-TranscriptOnLog
		Exit
	}
}

# Disclaimer
$Disclaimer = [Windows.Forms.MessageBox]::Show(
	"
Ce script a pour but de déléguer des droits sur une boîte partagée.
Pour cela, il injecte des données venant d'un fichier .csv directement sur Office 365.


									/!\ Attention /!\

Si vous n'êtes pas sûr des actions à mener, ou de l'impact sur la messagerie, quitter ce script dès à présent.

Souhaitez-vous continuer ?


", 'Boîte partagée', 1, [Windows.Forms.MessageBoxIcon]::Question)
If ($Disclaimer -eq "OK")
{
	Write-Information 'Patientez, traitement en cours ...'
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

# Import CSV file
[Windows.Forms.MessageBox]::Show(
	"
Sélectionner dans cette fenêtre le fichier contenant :
  - L'adresse e-mail du nom de l'utilisateur à attacher à une boîte partagée
  - Le nom de la boîte partagée concernée
  - Si la boîte partagée doit s'automonter (mapping)
  - Si l'utilisateur a tous les droits sur la boîte partagée
  - Si l'utilisateur est en mesure d'envoyer des messages 'de la part de'


Le fichier doit être de la forme suivante :

UserEmailAddress	SharedboxName	SendAs	Automapping
john.doe@dom.com	Production		Yes		No
jane.roe@dom.com	Production		No		Yes

", 'Shared mailbox', 0, [Windows.Forms.MessageBoxIcon]::Question)

# Import list of users and related sharedmailbox and rights
$CSVInputFile = Select-FileDialog -Title 'Select CSV file' -Filter 'Fichier CSV (*.csv) |*.csv'
$csvValues = Import-Csv $CSVInputFile -Delimiter ';'

# Set parameter for delegation with a loop
foreach ($line in $csvValues)
{
	$UserEmailAddress = $line.UserEmailAddress
	$SharedboxName = $line.SharedboxName
	$SendAs = $line.SendAs
	$Automapping = $line.Automapping
	switch ($SendAs)	{
		'Oui' { $SendAs = 'Yes' }
		'Non' { $SendAs = 'No' }
		default { $SendAs = $false }
	}
		switch ($Automapping) {
		'Oui' { $Automapping = $true }
		'Non' { $Automapping = $false }
		default { $Automapping = $true }
	}
	
	Write-Host $UserEmailAddress $SharedboxName $SendAs $Automapping
	
	#Adding users to the shared mailbox is a two-step process. First, we'll need to give the user access to the mailbox
	Add-MailboxPermission -Identity $SharedboxName -AccessRights 'FullAccess' -InheritanceType All -AutoMapping:$Automapping -User $UserEmailAddress
	
	if ($SendAs -eq 'Yes')
	{
		#Give the end user permission to send as the account
		Add-RecipientPermission -Identity $SharedboxName -AccessRights SendAs -Confirm:$false -Trustee $UserEmailAddress
	}
}


[Windows.Forms.MessageBox]::Show(
	"Action menée avec succès.
", 'Boîte partagée', 0, [Windows.Forms.MessageBoxIcon]::Information)


# Stop the connection to Office 365
Disconnect-EXOnline
