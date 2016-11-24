Add-Type -AssemblyName System.Windows.Forms

$PathDesktop = 'C:' + $env:HOMEPATH + '\Desktop'
$ContactOU = 'OU=Contacts,OU=Users,DC=domain,DC=dom'

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
Ce script a pour but de rajouter des contacts dans l'Active Directory.
Pour cela, il injecte des données venant d'un fichier .csv.


									/!\ Attention /!\

Si vous n'êtes pas sûr des actions à mener ou de l'impact, quitter ce script dès à présent.

Souhaitez-vous continuer ?


", "Active Directory - Ajout de contact", 1, [Windows.Forms.MessageBoxIcon]::Question)
If ($Disclaimer -eq "OK")
{
	Write-Host 'Patientez, traitement en cours ...'
}
Else
{
	Write-Error -Message 'Opération annulée'
	[Windows.Forms.MessageBox]::Show("Le script n'est pas en mesure de continuer. Opération stoppée.", 'Opération stoppée', 0, [Windows.Forms.MessageBoxIcon]::Error)
	Stop-TranscriptOnLog
	Exit
}

Import-Module ActiveDirectory

##region Import information coming from input .CSV file
# Import CSV file
[Windows.Forms.MessageBox]::Show(
	"
Sélectionner dans cette fenêtre le fichier contenant :
  - Le 'Prénom' du contact à créer
  - Le 'Nom' du contact à créer ()
  - L'adresse e-mail du contact à créer
  

Le fichier doit être de la forme suivante :

Prenom	NOM	Mail				Mobile		Titre		Service			Bureau	Societe		Adresse				Ville	CP		TelSociete	FaxSociete
John	Doe	john.doe@dom.com	0601234567	Admin Sys	IT				1		Microsoft	123 Rue des lilas	Paris	75001	0401234567	0401234444
Jane 	Roe jane.roe@dom.com	0701234567	DevOps		Developpment	2 		Apple 		456 Allée des roses Lyon	69001	0489012345	0489104444

", "Office 365 - Ajout de contact", 0, [Windows.Forms.MessageBoxIcon]::Question)

# Import list of users and related sharedmailbox and rights
$CSVInputFile = Select-FileDialog -Title 'Select CSV file' -Filter 'Fichier CSV (*.csv) |*.csv'
$csvValues = Import-Csv $CSVInputFile -Delimiter ';'
#endregion Import information coming from input .CSV file

#region Apply modification
foreach ($line in $csvValues)
{
	$UserFirstName = $line.'Prenom'
	$UserLASTNAME = $line.'Nom'.ToUpper()
	$UserName = $UserFirstName + ' ' + $UserLASTNAME
	$UserMailAddress = $line.Mail
	$UserMobile = $line.'Mobile'
	$UserTitle = $line.'Titre'
	$UserDepartment = $line.'Service'
	$UserCompanyOffice = $line.'Bureau'
	$UserCompanyName = $line.'Societe'
	$UserCompanyAddress = $line.'Adresse'
	$UserCompanyCity = $line.'Ville'
	$UserCompanyPostalCode = $line.CP
	$UserCompanyPhoneLine = $line.'TelSociete'
	$UserCompanyFax = $line.'FaxSociete'
	
	Write-Host 'Création du contact' $UserName $UserMailAddress $UserMobile	$UserTitle	$UserDepartment	$UserCompanyOffice $UserCompanyName	$UserCompanyAddress	$UserCompanyCity	$UserCompanyPostalCode	$UserCompanyPhoneLine	$UserCompanyFax
	New-ADObject -Type Contact -Name $UserName -Path $ContactOU -OtherAttributes @{ 'displayName' = $UserName; 'mail' = $UserMailAddress; 'proxyAddresses' = $UserMailAddress; 'givenName' = $UserFirstName; 'sn' = $UserLASTNAME; 'physicalDeliveryOfficeName' = $UserCompanyOffice; 'telephoneNumber' = $UserCompanyPhoneLine; 'mobile' = $UserMobile; 'facsimileTelephoneNumber' = $UserCompanyFax; 'streetAddress' = $UserCompanyAddress; 'l' = $UserCompanyCity; 'postalCode' = $UserCompanyPostalCode; 'c' = 'FR'; 'Title' = $UserTitle; 'department' = $UserDepartment; 'company' = $UserCompanyName }
}
#endregion Apply modification

#region Show MessageBox to confirm action
[Windows.Forms.MessageBox]::Show(
	"Action menée avec succès
", "ActiveDirectory - Ajout de contact", 0, [Windows.Forms.MessageBoxIcon]::Information)
#endregion Show MessageBox to confirm action
