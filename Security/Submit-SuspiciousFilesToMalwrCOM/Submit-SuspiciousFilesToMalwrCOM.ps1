#requires -version 2
<#
.SYNOPSIS
  Submit all files of a specific folder to Malwr.com webservice

.DESCRIPTION
   Submit all files of a specific folder to Malwr.com webservice

.INPUTS
  Personal API Key
  Folder path


.OUTPUTS
   Create transcript log file similar to $ScriptDir\[SCRIPTNAME]_[YYYY_MM_DD]_[HHhMMmSSs].log
   Generate results on console, in the meantime visible on log file   
   
.NOTES
  Version:        1.0
  Author:         ALBERT Jean-Marc
  Creation Date:  10/03/2016
  Purpose/Change: 1.0 - 2016.03.10 - ALBERT Jean-Marc - Initial script development
  
  
.SOURCES
  https://malwr.com/account/profile/ - API tab on a personal account profile

  
.EXAMPLE
  <None>
#>
 

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Get input strings
param(
[string] $computername = "$ENV:COMPUTERNAME",
[string] $reportfile = "$ScriptDir\$([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))" + "_" + $SystemDate + "_" + $SystemTime + ".html",
[string] $APIkey = "personal_api_key",
[string] $IsFileShared = "no",
[string] $SuspiciousFileContainer = "C:\suspiciousfilecontainer\"
)



#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = "1.0"

#Write script directory path on "ScriptDir" variable
$ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$Script
$SystemTime = Get-Date -uformat %Hh%Mm%Ss
$SystemDate = Get-Date -uformat %Y.%m.%d



#-----------------------------------------------------------[Functions]------------------------------------------------------------
 
#Send mail function
Function send_mail([string]$message,[string]$subject) {
$emailFrom = "name@domain.com"
$emailCC = "name@domain.com"
$emailTo = "name@domain.com"
$smtpServer = "smtp.domain.com"
Send-MailMessage -SmtpServer $smtpServer -To $emailTo -Cc $emailCC -From $emailFrom -Subject $subject -Body $message -BodyAsHtml -Priority High
}



#------------------------------------------------------------[Actions]------------------------------------------------------------- 

#Start stopwatch
$totalTime = New-Object -TypeName System.Diagnostics.Stopwatch
$totalTime.Start()
 
#Credits
Write-Host
Write-Host "Malwr.com - Suspicious file submission " -ForegroundColor "Yellow"
Write-Host

# List all files and import them
$fileEntries = [IO.Directory]::GetFiles("$SuspiciousFileContainer"); 
foreach($fileName in $fileEntries) 
{ 
    #Submit all files (located on the selected directory) to Malwr.com threat analysis webservice
    curl -F api_key=$APIkey -F shared=$IsFileShared -F file="$fileName" https://malwr.com/api/analysis/add/
    
} 


#Stop stopwatch
$totalTime.Stop()
$ts = $totalTime.Elapsed
$totalTime = [system.String]::Format("{0:00}:{1:00}:{2:00}",$ts.Hours, $ts.Minutes, $ts.Seconds)
Write-Host "Process total time: $totalTime" -ForegroundColor Yellow
Write-Host

Invoke-Item $reportfile