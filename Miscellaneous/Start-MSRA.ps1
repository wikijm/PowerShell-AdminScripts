#requires -version 2
<#
.SYNOPSIS
  <None>

.DESCRIPTION
  <None>

.INPUTS
  <None>

.OUTPUTS
  Fill a log file similar to $ScriptDir\[SCRIPTNAME]_[YYYY_MM_DD]_[HHhMMmSSs].log
  
   
   
.NOTES
  Version:        1.2
  Author:         ALBERT Jean-Marc
  Creation Date:  22/02/2017 (DD/MM/YYYY)
  Purpose/Change: 1.0 - 2017.02.22 - ALBERT Jean-Marc - Initial script development
  		  1.1 - 2017.02.27 - ALBERT Jean-Marc - Add combobox for Computer and User, thanks to request on Active Directory objects
		  					Modify logged information thanks to modification somewhat above
							Move buttons
							Apply templace to this script
  		  1.2 - 2017.02.27 - ALBERT Jean-Marc - Verify log's first line
							Add #region to [Execution] region
                                                  
.SOURCES
  <None>
  
  
.EXAMPLE
  <None>

#>

#region ---------------------------------------------------------[Initialisations]--------------------------------------------------------
Set-StrictMode -version Latest

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

$launchDate = get-date -f "dd/MM/yyyy"
$launchHour = get-date -f "HH:mm:ss"

Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms

#endregion

#region ----------------------------------------------------------[Declarations]----------------------------------------------------------

$scriptName = [System.IO.Path]::GetFileName($scriptFile)
$scriptVersion = "1.2"

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$logFileName = "pmad.log"
$logPathName = "C:\temp\$logFileName"
$logFileFirstLine = "Date,Heure,PosteAdmin,Admin,Utilisateur,PosteUser,Ticket"

$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()

$LDAPUserArray = Get-ADuser -SearchBase "OU=Users,DC=contorso,DC=dom" -Filter * | Select Name,SamAccountName | Sort Name
$LDAPComputerArray = Get-ADComputer -SearchBase "OU=Computers,DC=contorso,DC=dom" -Filter * | Select Name | Sort Name

#endregion

#region -----------------------------------------------------------[Functions]------------------------------------------------------------
function Set-FirstLine {
    param (
        [string]$Path,
        [string]$Value
    )
    $oldcontent = Get-Content $Path
    Set-Content -Path $Path -Value $Value
    Add-Content -Path $Path -Value $oldcontent 
}
#endregion

#region -----------------------------------------------------------[Execution]----------------------------------------------------------

#region Verify log's first line
	$logFileFirstLineContent = Get-content $logPathName -First 1
	if($logFileFirstLineContent -ne $logFileFirstLine){
    	Set-FirstLine -Path $logPathName -Value $logFileFirstLine 
	}
#endregion

#region Generate & show form
$Form = New-Object system.Windows.Forms.Form
$Form.Text = "Contrôle à distance"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$Form.TopMost = $false
$Form.Width = 460
$Form.Height = 305

$labelDescription = New-Object system.windows.Forms.Label
$labelDescription.Text = "Merci de remplir les champs ci-dessous afin de prendre la main
sur un poste CONTORSO.

Tous sont obligatoires."
$labelDescription.AutoSize = $true
$labelDescription.Width = 25
$labelDescription.Height = 10
$labelDescription.location = new-object system.drawing.point(10,20)
$labelDescription.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($labelDescription)

$LabelProd = New-Object system.windows.Forms.Label
$LabelProd.Text = "Administrateur"
$LabelProd.AutoSize = $true
$LabelProd.Width = 25
$LabelProd.Height = 10
$LabelProd.location = new-object system.drawing.point(24,101)
$LabelProd.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($LabelProd)
$textBoxProd = New-Object system.windows.Forms.TextBox
$textBoxProd.Text = $env:USERNAME
$textBoxProd.Enabled = $false
$textBoxProd.Width = 100
$textBoxProd.Height = 20
$textBoxProd.location = new-object system.drawing.point(125,101)
$textBoxProd.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($textBoxProd)
$textBoxProdComputer = New-Object system.windows.Forms.TextBox
$textBoxProdComputer.Text = $env:COMPUTERNAME
$textBoxProdComputer.Enabled = $false
$textBoxProdComputer.Width = 100
$textBoxProdComputer.Height = 20
$textBoxProdComputer.location = new-object system.drawing.point(265,101)
$textBoxProdComputer.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($textBoxProdComputer)

$LabelUtilisateur = New-Object system.windows.Forms.Label
$LabelUtilisateur.Text = "Utilisateur"
$LabelUtilisateur.AutoSize = $true
$LabelUtilisateur.Width = 25
$LabelUtilisateur.Height = 10
$LabelUtilisateur.location = new-object system.drawing.point(24,151)
$LabelUtilisateur.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($LabelUtilisateur)
$comboBoxUtilisateur = New-Object system.windows.Forms.ComboBox
$comboBoxUtilisateur.Text = "xxxN"
$comboBoxUtilisateur.Width = 220
$comboBoxUtilisateur.Height = 20
$comboBoxUtilisateur.location = new-object system.drawing.point(125,151)
$comboBoxUtilisateur.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($comboBoxUtilisateur)
	foreach ($LDAPUser in $LDAPUserArray)	{
    $LDAPUserResult = "{0} ({1})" -f $LDAPUser.Name, $LDAPUser.SamAccountName
    [void] $comboBoxUtilisateur.Items.Add($LDAPUserResult)
    }

$LabelComputerName = New-Object system.windows.Forms.Label
$LabelComputerName.Text = "Poste"
$LabelComputerName.AutoSize = $true
$LabelComputerName.Width = 25
$LabelComputerName.Height = 10
$LabelComputerName.location = new-object system.drawing.point(23,188)
$LabelComputerName.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($LabelComputerName)
$comboBoxComputerName = New-Object system.windows.Forms.ComboBox
$comboBoxComputerName.Text = "Uxnnnnnnnnn"
$comboBoxComputerName.Width = 110
$comboBoxComputerName.Height = 20
$comboBoxComputerName.location = new-object system.drawing.point(125,189)
$comboBoxComputerName.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($comboBoxComputerName)
   	foreach ($LDAPComputer in $LDAPComputerArray)	{
    [void] $comboBoxComputerName.Items.Add($LDAPComputer.Name)
    }

$labelTicket = New-Object system.windows.Forms.Label
$labelTicket.Text = "Ticket"
$labelTicket.AutoSize = $true
$labelTicket.Width = 25
$labelTicket.Height = 10
$labelTicket.location = new-object system.drawing.point(22,230)
$labelTicket.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($labelTicket)
$textBoxTicket = New-Object system.windows.Forms.TextBox
$textBoxTicket.Text = "xxxxx"
$textBoxTicket.Width = 110
$textBoxTicket.Height = 20
$textBoxTicket.location = new-object system.drawing.point(125,230)
$textBoxTicket.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($textBoxTicket)

$buttonStartRC = New-Object system.windows.Forms.Button
$buttonStartRC.Text = "Lancer contrôle à distance"
$buttonStartRC.Width = 140
$buttonStartRC.Height = 50
$buttonStartRC.Add_MouseClick({
    $Result = "$env:COMPUTERNAME,$env:USERNAME,$($comboBoxUtilisateur.Text),$($textBoxComputerName.Text),$($textBoxTicket.Text)"
	$launchDate + ',' + $launchHour + ',' +$env:COMPUTERNAME + ',' + $env:USERNAME + ',' + $comboBoxUtilisateur.Text + ',' + $comboBoxComputerName.Text + ',' + $textBoxTicket.Text | Out-File -filepath $logPathName -Append -Encoding UTF8
    msra.exe /offerRA $comboBoxComputerName.Text
})
$buttonStartRC.location = new-object system.drawing.point(240,195)
$buttonStartRC.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($buttonStartRC)

$buttonLog = New-Object system.windows.Forms.Button
$buttonLog.Text = "Log"
$buttonLog.Width = 60
$buttonLog.Height = 30
$buttonLog.location = new-object system.drawing.point(385,205)
$buttonLog.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($buttonLog)
$buttonLog.Add_MouseClick({
     Invoke-Item $logPathName
})


[void]$Form.ShowDialog()
$Form.Dispose()
#endregion

#endregion

#endregion
