#requires -version 2
<#
.SYNOPSIS
  Add writes on shared mailbox to Office 35 users, with or without automapping

.DESCRIPTION
  Add writes on shared mailbox to Office 35 users, with or without automapping

.INPUTS
  .csv file as defined below

.OUTPUTS
  Create transcript log file similar to $ScriptDir\[SCRIPTNAME]_[YYYY_MM_DD]_[HHhMMmSSs].log
  Create a list of SUBJECT, similar to $ScriptDir\[SCRIPTNAME]_SUBJECT_[YYYY_MM_DD]_[HHhMMmSSs].csv
   
   
.NOTES
  Version:        0.2
  Author:         ALBERT Jean-Marc
  Creation Date:  14/09/2017 (DD/MM/YYYY)
  Purpose/Change: 0.1 - 2017.09.14 - ALBERT Jean-Marc - Initial script development
                  0.2 - 2017.09.27 - ALBERT Jean-Marc - Verify if started as admin
                    
                                                  
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
$buffer = "$scriptPath\bufferCommand.txt"
$fullScriptPath = (Resolve-Path -Path $buffer).Path

Add-Type -AssemblyName System.Windows.Forms
#endregion

#region ----------------------------------------------------------[Declarations]----------------------------------------------------------

$scriptName = [System.IO.Path]::GetFileName($scriptFile)
$scriptVersion = "0.2"

if(!(Test-Path $logDirectoryPath)) {
    New-Item $logDirectoryPath -type directory | Out-Null
}

$logFileName = "Log_" + $launchDate + ".log"
$logPathName = "$logDirectoryPath\$logFileName"

$global:streamWriter = New-Object System.IO.StreamWriter $logPathName

#endregion

#region -----------------------------------------------------------[Functions]------------------------------------------------------------

Function Test-IsAdmin { 
<#     
.SYNOPSIS     
   Function used to detect if current user is an Administrator.  
     
.DESCRIPTION   
   Function used to detect if current user is an Administrator. Presents a menu if not an Administrator  
      
.NOTES     
    Name: Test-IsAdmin  
    Author: Boe Prox   
    DateCreated: 30April2011    
      
.EXAMPLE     
    Test-IsAdmin  
      
   
Description   
-----------       
Command will check the current user to see if an Administrator. If not, a menu is presented to the user to either  
continue as the current user context or enter alternate credentials to use. If alternate credentials are used, then  
the [System.Management.Automation.PSCredential] object is returned by the function.  
#>  
    [cmdletbinding()]  
    Param()  
      
    Write-Verbose "Checking to see if current user context is Administrator"  
    If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))  
    {  
        Write-Warning "You are not currently running this under an Administrator account! `nThere is potential that this command could fail if not running under an Administrator account."  
        Write-Verbose "Presenting option for user to pick whether to continue as current user or use alternate credentials"  
        #Determine Values for Choice  
        $choice = [System.Management.Automation.Host.ChoiceDescription[]] @("Use &Alternate Credentials","&Continue with current Credentials")  
  
        #Determine Default Selection  
        [int]$default = 0  
  
        #Present choice option to user  
        $userchoice = $host.ui.PromptforChoice("Warning","Please select to use Alternate Credentials or current credentials to run command",$choice,$default)  
  
        Write-Debug "Selection: $userchoice"  
  
        #Determine action to take  
        Switch ($Userchoice)  
        {  
            0  
            {  
                #Prompt for alternate credentials  
                Write-Verbose "Prompting for Alternate Credentials"  
                $Credential = Get-Credential  
                Write-Output $Credential      
            }  
            1  
            {  
                #Continue using current credentials  
                Write-Verbose "Using current credentials"  
                Write-Output "CurrentUser"  
            }  
        }          
          
    }  
    Else   
    {  
        Write-Verbose "Passed Administrator check"  
    }  
}

Function Hide-PowershellConsole() {
    # Hide the powershell console window without hiding the other child windows that it spawns. (I.E. hide the powershell window, but not the Out-Gridview window)
    # https://community.spiceworks.com/topic/1710213-hide-a-powershell-console-window-when-running-a-script
    $Script:showWindowAsync = Add-Type -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru

    $null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 2)
}

Function Start-Log {    
    [CmdletBinding()]  
    Param ([Parameter(Mandatory=$true)][string]$scriptName, [Parameter(Mandatory=$true)][string]$scriptVersion, 
        [Parameter(Mandatory=$true)][string]$streamWriter)
    Process{                  
        $global:streamWriter.WriteLine("================================================================================================")
        $global:streamWriter.WriteLine("[$ScriptName] version [$ScriptVersion] started at $([DateTime]::Now)")
        $global:streamWriter.WriteLine("================================================================================================`n")       
    }
}
 
Function Write-Log {
    [CmdletBinding()]  
    Param ([Parameter(Mandatory=$true)][string]$streamWriter, [Parameter(Mandatory=$true)][string]$infoToLog)  
    Process{    
        $InfoMessage = "$([DateTime]::Now) [INFO] $infoToLog"
        $global:streamWriter.WriteLine($InfoMessage)
        Write-Host $InfoMessage -ForegroundColor Cyan
    }
}
 
Function Write-Error {
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
 
Function End-Log { 
    [CmdletBinding()]  
    Param ([Parameter(Mandatory=$true)][string]$streamWriter)  
    Process{    
        $global:streamWriter.WriteLine("`n================================================================================================")
        $global:streamWriter.WriteLine("Script ended at $([DateTime]::Now)")
        $global:streamWriter.WriteLine("================================================================================================")
  
        $global:streamWriter.Close()   
    }
}

Function Invoke-Office365TenantLogon {
    #### Pop-up a dialog for username and request your password
    $cred = Get-Credential
    #### Import the Local Microsoft Online PowerShell Module Cmdlets and Connect to O365 Online
    Import-Module MSOnline
    Connect-MsolService -Credential $cred
    #### Establish an Remote PowerShell Session to Exchange Online
    $msoExchangeURL = “https://ps.outlook.com/powershell/”
    $sessionOption = New-PSSessionOption -SkipRevocationCheck #Avoid Certificate error (https://support.microsoft.com/fr-fr/help/2792168/-ssl-certificate-could-not-be-checked-for-revocation-error-when-you-co)
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $msoExchangeURL -Credential $cred -Authentication Basic -AllowRedirection -SessionOption $sessionOption

    if (!$session) {
	    Write-Error -Message 'Opération annulée'
		[Windows.Forms.MessageBox]::Show("Le script n'est pas en mesure de continuer, sans doute à cause de mauvaises informations d'identification. Opération stoppée.", 'Opération stoppée', 0, [Windows.Forms.MessageBoxIcon]::Error)
		Stop-TranscriptOnLog
		Exit
	}
	Else {
        Import-PSSession $session -AllowClobber
	}

 }

Function Invoke-Office365TenantLogoff {
    #### Remove the Remote PowerShell Session to Exchange Online ----
    Get-PsSession | Remove-PsSession
    #Remove-PsSession $session
}

Function Select-FileDialog {
	[CmdletBinding()]
	param ([string]$Title,
		[string]$Filter = 'All files *.*|*.*')
	Add-Type -AssemblyName System.Windows.Forms | Out-Null
	$fileDialogBox = New-Object -TypeName Windows.Forms.OpenFileDialog
	$fileDialogBox.ShowHelp = $false
	$fileDialogBox.initialDirectory = $scriptPath
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

Function Stop-Script () {   
    Begin{
        Write-Log -streamWriter $global:streamWriter -infoToLog '--- Script terminating ---'
    }
    Process{        
        'Script terminating...' 
        Write-Verbose -Message '================================================================================================'
        End-Log -streamWriter $global:streamWriter       
        Exit
    }
}
#endregion

#region ----------------------------------------------------------[Execution]----------------------------------------------------------

# Check if Admin
$admincheck = Test-IsAdmin
    If ($admincheck -is [System.Management.Automation.PSCredential]) {
        Write-Error -Message 'Opération annulée'
        Start-Process -FilePath PowerShell.exe -Credential $admincheck -ArgumentList $myinvocation.mycommand.definition
        Break
    }

# Start log creation and completion
Start-Log -scriptName $scriptName -scriptVersion $scriptVersion -streamWriter $global:streamWriter

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
Invoke-Office365TenantLogon

# Import CSV file
[Windows.Forms.MessageBox]::Show(
	"
Sélectionner dans cette fenêtre le fichier contenant :
  - L'adresse e-mail du nom de l'utilisateur à attacher à une boîte partagée
  - Le nom de la boîte partagée concernée
  - Si la boîte partagée doit s'automonter (mapping)

Le fichier doit être de la forme suivante :

UserEmailAddress	SharedboxName	SendAs	Automapping
john.doe@dom.com	Production		Yes		No
jane.roe@dom.com	Production		No		Yes

", 'Shared mailbox', 0, [Windows.Forms.MessageBoxIcon]::Question)

# Import list of users and related sharedmailbox and rights
$CSVInputFile = Select-FileDialog -Title 'Select CSV file' -Filter 'Fichier CSV (*.csv) |*.csv'
$csvValues = Import-Csv $CSVInputFile -Delimiter ';'
Write-Log -streamWriter $global:streamWriter -infoToLog Fichier source sélectionné : $CSVInputFile

# Set parameter for delegation with a loop
foreach ($line in $csvValues)
{
	$UserEmailAddress = $line.UserEmailAddress
	$SharedboxName = $line.SharedboxName
	$SendAs = $line.SendAs
	$Automapping = $line.Automapping
	switch ($SendAs)
	{
		'Oui' { $SendAs = 'Yes' }
		'Non' { $SendAs = 'No' }
		default { $SendAs = $false }
	}
	switch ($Automapping)
	{
		'Oui' { $Automapping = $true }
		'Non' { $Automapping = $false }
		default { $Automapping = $true }
	}
	
	Write-Host $UserEmailAddress $SharedboxName $SendAs $Automapping
    Write-Log -streamWriter $global:streamWriter -infoToLog $UserEmailAddress $SharedboxName $SendAs $Automapping
    	
	#Adding users to the shared mailbox is a two-step process. First, we'll need to give the user access to the mailbox
	Add-MailboxPermission -Identity $SharedboxName -AccessRights 'FullAccess' -InheritanceType All -AutoMapping:$Automapping -User $UserEmailAddress
	
	if ($SendAs -eq 'Yes')
	{
		#Give the end user permission to send as the account
		Add-RecipientPermission -Identity $SharedboxName -AccessRights SendAs -Confirm:$false -Trustee $UserEmailAddress
	}
}


# Success message
[Windows.Forms.MessageBox]::Show(
	"Action menée avec succès.
", 'Boîte partagée', 0, [Windows.Forms.MessageBoxIcon]::Information)
Write-Log -streamWriter $global:streamWriter -infoToLog 'Action menée avec succès.'


# Stop the connection to Office 365
Invoke-Office365TenantLogoff
Write-Log -streamWriter $global:streamWriter -infoToLog "Déconnexion de l'instance Office 365 Tenant"

Stop-Script
#endregion
