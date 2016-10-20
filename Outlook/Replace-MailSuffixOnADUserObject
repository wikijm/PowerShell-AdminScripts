#requires -version 2
<#
.SYNOPSIS
  Replace @olddomain.fr to @newdomain.com on AD User e-mail address

.DESCRIPTION
  List AD users with old suffix on e-mail address, export a list on CSV then replace old suffix by new one.
  Be aware to modify  $oldSuffix and $newSuffix (lines 128 & 129) before usage.


.INPUTS
  <None>

.OUTPUTS
  Create transcript log file similar to $ScriptDir\[SCRIPTNAME]_[YYYY_MM_DD]_[HHhMMmSSs].log
  Create a list of AD user objects, similar to $ScriptDir\[SCRIPTNAME]_SUBJECT_[YYYY_MM_DD]_[HHhMMmSSs].csv
   
   
.NOTES
  Version:        1.1
  Author:         ALBERT Jean-Marc
  Creation Date:  19/10/2016 (DD/MM/YYYY)
  Purpose/Change: 1.0 - 2016.10.19 - ALBERT Jean-Marc - Initial script development
                  1.1 - 2016.10.20 - ALBERT Jean-Marc - Replace hard-coded value for old suffix research on Get-AdUser selection with $oldsuffix
                    
                                                  
.SOURCES
  N/A
  
  
.EXAMPLE
  <None>

#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
Set-StrictMode -version Latest

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$scriptFile = $MyInvocation.MyCommand.Definition
$launchDate = get-date -f "yyyy.MM.dd-HHmmmss"
$logDirectoryPath = $scriptPath + "\" + $launchDate
$buffer = "$scriptPath\bufferCommand.txt"
$fullScriptPath = (Resolve-Path -Path $buffer).Path



#----------------------------------------------------------[Declarations]----------------------------------------------------------

$scriptName = [System.IO.Path]::GetFileName($scriptFile)
$scriptVersion = "1.0"

if(!(Test-Path $logDirectoryPath)) {
    New-Item $logDirectoryPath -type directory | Out-Null
}

$logFileName = "Log_" + $launchDate + ".log"
$logPathName = "$logDirectoryPath\$logFileName"
$ExportFileName = "Export_" + $launchDate + ".csv"
$ExportPathName = "$logDirectoryPath\$ExportFileName"

$global:streamWriter = New-Object System.IO.StreamWriter $logPathName

$PrerequisitesModules = "ActiveDirectory"
$PrerequisitesModulesError = "$Module does not exist, please install it then relaunch the script."


#-----------------------------------------------------------[Functions]------------------------------------------------------------

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


#----------------------------------------------------------[Execution]----------------------------------------------------------

Start-Log -scriptName $scriptName -scriptVersion $scriptVersion -streamWriter $global:streamWriter
Write-Host "================================================================================================"
Write-Host "List AD user account without a recent password (unmodified since $NumberOfDays days)"
Write-Host "================================================================================================"


$oldSuffix = "@olddomain.fr"
$newSuffix = "@newdomain.com"

$ADUserSelection = Get-AdUser -Filter * -Property GivenName,sn,mail,lastlogondate,samAccountName | Where-Object {($_.Enabled) -eq "True" -AND ($_.mail) -match $oldSuffix} | select GivenName,sn,mail,lastlogondate,samAccountName
$ADUserSelection | Export-CSV -Path $ExportPathName -Encoding UTF8 -NoType -Force
$ADUserSelection | ForEach-Object {
                        $newUpn = $_.mail.Replace($oldSuffix,$newSuffix)
                        $newUpn
                        Set-ADUser -identity $_.samaccountname -EmailAddress $newUpn
                        }

#Writing informations in the log file
Write-Progress -Activity "Write informations in the log file" -status "Running..." -id 1
Write-Log -streamWriter $global:streamWriter -infoToLog "If you don't see any lines between start & end of logfile, it means that everything's OK"
End-Log -streamWriter $global:streamWriter
notepad $logPathName
#cls
