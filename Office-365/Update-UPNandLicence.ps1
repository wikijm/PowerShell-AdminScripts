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
  Version:        0.1
  Author:         ALBERT Jean-Marc
  Creation Date:  08/02/2017 (DD/MM/YYYY)
  Purpose/Change: 1.0 - 2017.02.08 - ALBERT Jean-Marc - Initial script development
                    
                                                  
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
#endregion

#region ----------------------------------------------------------[Declarations]----------------------------------------------------------
$scriptName = [System.IO.Path]::GetFileName($scriptFile)
$scriptVersion = "0.1"

if(!(Test-Path $logDirectoryPath)) {
    New-Item $logDirectoryPath -type directory | Out-Null
}

$logFileName = "Log_" + $launchDate + ".log"
$logPathName = "$logDirectoryPath\$logFileName"
$global:streamWriter = New-Object System.IO.StreamWriter $logPathName
#endregion

#region -----------------------------------------------------------[Functions]------------------------------------------------------------

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

Function Select-FileDialog {
    param([string]$Title,[string]$Filter="All files *.*|*.*")
	[System.Reflection.Assembly]::LoadWithPartialName( 'System.Windows.Forms' ) | Out-Null
	$fileDialogBox = New-Object Windows.Forms.OpenFileDialog
	$fileDialogBox.ShowHelp = $false
	$fileDialogBox.initialDirectory = $ScriptDir
	$fileDialogBox.filter = $Filter
    $fileDialogBox.Title = $Title
	$Show = $fileDialogBox.ShowDialog( )

        If ($Show -eq "OK")
            {
                Return $fileDialogBox.FileName
            }
        Else
            {
                Write-Error "Canceled operation"
		          [System.Windows.Forms.MessageBox]::Show("Script is not able to continue. Operation stopped." , "Operation canceled" , 0, [Windows.Forms.MessageBoxIcon]::Error)
                Stop-TranscriptOnLog
		        Exit
            }
}

Function Connect-Office365_PSSession {
<#
	.SYNOPSIS
		Connect to Office 365 Cloud Services into active PowerShell shell
	
	.PARAMETER Office365_URL
		URL to Office 365 Cloud Services for PowerShell
	
	.PARAMETER Office365_Credentials
		Capture administrative credential for future connections
	
#>
	[CmdletBinding()]
	param
	(
		[string]$Office365_URL = "https://ps.outlook.com/powershell",
			$Office365_Credentials
	)
	
	#Imports the installed Azure Active Directory module.
	Import-Module MSOnline
	
	#Capture administrative credential for future connections.
	$Office365_Credentials = Get-Credential -Message "Enter your Office 365 admin credentials"
	
	#Establishes Online Services connection to Office 365 Management Layer.
	Connect-MsolService -Credential $Office365_Credentials
	
	#Creates an Exchange Online session using defined credential.
	$EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Office365_URL -Credential $Office365_Credentials -Authentication Basic -AllowRedirection -Name "Exchange Online"
	
	#This imports the Office 365 session into your active Shell.
	Import-PSSession $EXOSession
}

Function Disconnect-Office365_PSSession {
<#
	.SYNOPSIS
		Disconnect PowerShell session to Office 365 Cloud Services
	
	.PARAMETER PSSession_Name
		Default name is "Exchange Online"
#>
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$PSSession_Name
	)
	
	Remove-PSSession -Name $PSSession_Name
}

#endregion

#region ----------------------------------------------------------[Execution]----------------------------------------------------------

Start-Log -scriptName $scriptName -scriptVersion $scriptVersion -streamWriter $global:streamWriter
cls
Write-Host "================================================================================================"
# Prerequisites
#N/A
Write-Host "================================================================================================"

# Connect to Office 365 Cloud Services into active PowerShell shell
Connect-Office365_PSSession

# Import CSV input file (user list) 
[System.Windows.Forms.MessageBox]::Show(
"
Select on this window the CSV file who contains user list.
Its content must be similar to:

EmailAddress,UserName,Password,o365OldUPN,UsageLocation,o365License,o365Password       (Required line)
firstname.lastname@domain.com,firstname.lastname@domain.com,StrongPassword,firstname.lastname@domain.onmicrosoft.com,FR,reseller-account:EXCHANGESTANDARD, 
", "User list", 0, [Windows.Forms.MessageBoxIcon]::Question)

$CSVInputFile = Select-FileDialog -Title "Select CSV file" -Filter "CSV File (*.csv) |*.csv"

# Import list of mailbox who need to create online archive
$csvValues = Import-Csv $CSVInputFile -Delimiter ','


# Update UPN, user license and password
ForEach ($usr in $csvValues)
{
	$Lic = $usr.o365License
	$OldUPN = $usr.o365OldUPN
	$NewUPN = $usr.EmailAddress
	# $NewPass = $usr.o365Password
	
	$msoluser = Get-MsolUser -UserPrincipalName $OldUPN 2>$null | Select-Object UserPrincipalName, UsageLocation
	
	if ($msoluser)
	{
		Write-Log -streamWriter $global:streamWriter -infoToLog 'Update $OldUPN to $NewUPN and Licence $Lic'
		Set-MsolUserPrincipalName -UserPrincipalName $OldUPN -NewUserPrincipalName $NewUPN
		Set-Msoluser -UserPrincipalName $NewUPN -UsageLocation $UsageLocation
		Set-MsolUserLicense -UserPrincipalName $NewUPN -AddLicenses $Lic
	}
	else
	{
		Write-Error -Message '$OldUPN does not exists on the Tenant'
		Write-Log -streamWriter $global:streamWriter -infoToLog 'WARNING: $OldUPN does not exists on the Tenant'
	}
}

End-Log -streamWriter $global:streamWriter
notepad $logPathName
cls
#endregion
