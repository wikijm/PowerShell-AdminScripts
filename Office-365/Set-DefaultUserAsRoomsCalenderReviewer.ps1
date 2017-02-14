#requires -version 2
<#
.SYNOPSIS
  <None>

.DESCRIPTION
  <None>

.INPUTS
  <None>

.OUTPUTS
  Create transcript log file similar to $ScriptDir\[SCRIPTNAME]_[YYYY_MM_DD]_[HHhMMmSSs].log
  Create a list of SUBJECT, similar to $ScriptDir\[SCRIPTNAME]_SUBJECT_[YYYY_MM_DD]_[HHhMMmSSs].csv
   
   
.NOTES
  Version:        1.2
  Author:         ALBERT Jean-Marc
  Creation Date:  10/01/2017 (DD/MM/YYYY)
  Current Date:	  14/02/2017 (DD/MM/YYYY)
  Purpose/Change: 1.0 - 2017.01.10 - ALBERT Jean-Marc - Initial script development
                  1.1 - 2017.01.10 - ALBERT Jean-Marc - Redesign script with PowerShell-ScriptTemplate.ps1
		  1.2 - 2017.02.14 - ALBERT Jean-Marc - Add "current date" and "region"
                
                    
                                                  
.SOURCES
  <None>
  
  
.EXAMPLE
  <None>

#>

#region ---------------------------------------------------------[Initialisations]--------------------------------------------------------
Set-StrictMode -version Latest

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$scriptFile = $MyInvocation.MyCommand.Definition
$launchDate = get-date -f "yyyy.MM.dd-HHhmmmss"
$logDirectoryPath = $scriptPath + "\" + $launchDate
#endregion

#region ----------------------------------------------------------[Declarations]----------------------------------------------------------
$scriptName = [System.IO.Path]::GetFileName($scriptFile)
$scriptVersion = "1.2"

if(!(Test-Path $logDirectoryPath)) {
    New-Item $logDirectoryPath -type directory | Out-Null
}

$logFileName = "Log_" + $launchDate + ".log"
$logPathName = "$logDirectoryPath\$logFileName"

$global:streamWriter = New-Object System.IO.StreamWriter $logPathName
#endregion

#region -----------------------------------------------------------[functions]------------------------------------------------------------
function Start-Log {    
    [CmdletBinding()]  
    Param ([Parameter(Mandatory=$true)][string]$scriptName, [Parameter(Mandatory=$true)][string]$scriptVersion, 
        [Parameter(Mandatory=$true)][string]$streamWriter)
    Process{                  
        $global:streamWriter.WriteLine("================================================================================================")
        $global:streamWriter.WriteLine("[$ScriptName] version [$ScriptVersion] started at $([DateTime]::Now)")
        $global:streamWriter.WriteLine("================================================================================================`n")       
    }
}
 
function Write-Log {
    [CmdletBinding()]  
    Param ([Parameter(Mandatory=$true)][string]$streamWriter, [Parameter(Mandatory=$true)][string]$infoToLog)  
    Process{    
        $InfoMessage = "$([DateTime]::Now) [INFO] $infoToLog"
        $global:streamWriter.WriteLine($InfoMessage)
        Write-Host $InfoMessage -ForegroundColor Cyan
    }
}
 
function Write-Error {
    [CmdletBinding()]  
    Param ([Parameter(Mandatory=$true)][string]$streamWriter, [Parameter(Mandatory=$true)][string]$errorCaught, [Parameter(Mandatory=$true)][boolean]$forceExit)  
    Process{
        $ErrorMessage = "$([DateTime]::Now) [ERROR] $errorCaught"
        $global:streamWriter.WriteLine($ErrorMessage)
        Write-Host $ErrorMessage -ForegroundColor Red      
        if ($forceExit -eq $true){
            End-Log -streamWriter $global:streamWriter
            break;
        }
    }
}
 
function End-Log { 
    [CmdletBinding()]  
    Param ([Parameter(Mandatory=$true)][string]$streamWriter)  
    Process{    
        $global:streamWriter.WriteLine("`n================================================================================================")
        $global:streamWriter.WriteLine("Script ended at $([DateTime]::Now)")
        $global:streamWriter.WriteLine("================================================================================================")
  
        $global:streamWriter.Close()   
    }
}

function Connect-EXOnline {
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

function Disconnect-EXOnline {
	Remove-PSSession -Name "Exchange Online"
}
#endregion

#region ----------------------------------------------------------[Execution]----------------------------------------------------------
Start-Log -scriptName $scriptName -scriptVersion $scriptVersion -streamWriter $global:streamWriter
cls
Write-Host "================================================================================================"
Write-Host "[$ScriptName] version [$ScriptVersion] started at $([DateTime]::Now)"
Write-Host "================================================================================================"

# Disclaimer
$Disclaimer = [Windows.Forms.MessageBox]::Show(
	"
Ce script permet à tous les utilisateurs Exchange/Office 365 de devenir relecteur sur les calendriers de salle de réunion et équipements.

                                 /!\ Attention /!\
Si vous n'êtes pas sûr des actions à mener ou de l'impact, quitter ce script dès à présent.

Souhaitez-vous continuer ?

", 'Modification de droits sur les calendriers de salle et équipement', 1, [Windows.Forms.MessageBoxIcon]::Question)
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

# List all RoomMailbox and set Exchange/Office 365 users as calendar reviewer
$rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox
$rooms | ForEach-Object {Write-Log -streamWriter $global:streamWriter -infoToLog "Room: $_"}
$rooms | %{Set-MailboxFolderPermission $_":\Calendrier" -User Default -AccessRights Reviewer} | ForEach-Object {Write-Log -streamWriter $global:streamWriter -infoToLog $_} #Replace :\Calendrier with :\Calendar for English context

# List all EquipmentMailbox and set Exchange/Office 365 users as calendar reviewer
$equipments = Get-Mailbox -RecipientTypeDetails EquipmentMailbox
$equipments | ForEach-Object {Write-Log -streamWriter $global:streamWriter -infoToLog "Equipment: $_"} 
$equipments | %{Set-MailboxFolderPermission $_":\Calendrier" -User Default -AccessRights Reviewer} | ForEach-Object {Write-Log -streamWriter $global:streamWriter -infoToLog $_} #Replace :\Calendrier with :\Calendar for English context

Write-Log -streamWriter $global:streamWriter -infoToLog $rooms

[Windows.Forms.MessageBox]::Show(
	"Action menée avec succès.
", 'Modification de droits sur les calendriers de salle et équipement', 0, [Windows.Forms.MessageBoxIcon]::Information)


# Stop the connection to Office 365
Disconnect-EXOnline


#Writing informations in the log file
Write-Progress -Activity "Write informations in the log file" -status "Running..." -id 1
End-Log -streamWriter $global:streamWriter
notepad $logPathName
cls
#endregion
