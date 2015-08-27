#requires -version 2
<#
.SYNOPSIS  
    This script uses the Outlook COM object to display the data stores in the current profile.
.DESCRIPTION  
    This script creates an Outlook object, displays user information, and the stores currently attached to the profile.
.INPUTS
  <None>
.OUTPUTS
    Results on a console. See .EXAMPLE to have an overview.
.NOTES
  Version:        1.0
  Author:         ALBERT Jean-Marc
  Creation Date:  29/06/2015
  Purpose/Change: 1.0 - 2015.06.29 - ALBERT Jean-Marc - Initial script development
.EXAMPLE  
    PSH [C:\foo]: .\List-AllMailboxAndPST.ps1'  
    Current profile has the following configured accounts:  
  
    Account Type           					User Name				SMTP Address  
    ------------           					---------        		------------  
    Jean-Marc.Albert-EXT@domain.com		Jean-Marc.ALBERT-EXT    Jean-Marc.Albert-EXT@domain.com  
      
    Exchange Offile Folder Store:  
    C:\Users\9999912\AppData\Local\Microsoft\Outlook\Jean-Marc.Albert-EXT@domain.com.ost  
     
    PST Files  
    Display Name    File Path   
    ------------    ---------  
    Archive Folders C:\Users\jean-marc.albert\AppData\Local\Microsoft\Outlook\archive.pst  
#>
#---------------------------------------------------------[Initialisations]--------------------------------------------------------
#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#----------------------------------------------------------[Declarations]----------------------------------------------------------
#Script Version
$sScriptVersion = "1.0"
#Write script directory path on "ScriptDir" variable
$ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
#Log file creation, similar to $ScriptDir\[SCRIPTNAME]_[YYYY_MM_DD].log
$SystemTime = Get-Date -uformat %Hh%Mm%Ss
$SystemDate = Get-Date -uformat %Y.%m.%d
$ScriptLogFile = "$ScriptDir\$([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))" + "_" + $SystemDate + "_" + $SystemTime + ".log"

#-----------------------------------------------------------[Functions]------------------------------------------------------------
function Stop-TranscriptOnLog
	{
	Stop-Transcript
	# Add EOL required for Notepad.exe application usage
	[string]::Join("`r`n", (Get-Content $ScriptLogFile)) | Out-File $ScriptLogFile
	}

function Pause ($Message = "Press any key to continue . . . ")
	{
	if ((Test-Path variable:psISE) -and $psISE)
		{
			$Shell = New-Object -ComObject "WScript.Shell"
			$Button = $Shell.Popup("Click OK to continue.", 0, "Script Paused", 0)
		}
	else
		{
			Write-Host -NoNewline $Message
			[void][System.Console]::ReadKey($true)
			Write-Host
		}
	}

#------------------------------------------------------------[Actions]-------------------------------------------------------------
# Start of log completion
Start-Transcript $ScriptLogFile | Out-Null

# Create Outlook object
$Outlook = New-Object -ComObject Outlook.Application  
$stores = $Outlook.Session.Stores
$accounts = $outlook.session.accounts

#$OutputData = @(
# Basic information  
"Current profile has the following configured accounts:"  
$dn = @{label = "Account Type"; expression={$_.displayname}}  
$un = @{label = "User Name"; expression = {$_.username}}  
$sm = @{label = "SMTP Address"; expression = {$_.smtpaddress}}  
$accounts | Format-Table -AutoSize $dn,$un,$sm  
# Check number of stores > 0  
if ($stores.Count -le 0) {"No stores found"; return}  
  
# Outlook Off-Line folder store  
$ost = $stores | where{$_.filepath -match ".ost$"}  
if (!$ost)  
{  
   "No Outlook Offline Folder store found"  
}  
else  
{  
   "Exchange Offile Folder Store:"  
   $ost | ft filepath -HideTableHeaders  
  }  
  
# PST Files  
$pst = $stores | where {$_.filepath -match ".pst$"}  
if (!$pst)  
  	{  
    "No PST files found"  
  	}  
else  
  	{  
    "PST Files"  
    $dn = @{label = "Display Name"; expression={$_.displayname}}  
    $fn = @{label = "File Path"; expression={$_.filepath}}  
    $pst | ft $dn,$fn
	}
#)

   #[System.Windows.Forms.MessageBox]::Show("$OutputData" , "Information" , 0, [Windows.Forms.MessageBoxIcon]::Information)

# Pause the script to see results on console
Pause

# End Script