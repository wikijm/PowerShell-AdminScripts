#requires -version 2
<#
    .SYNOPSIS
    List all AD Objects on a selected OU, and export results into a CSV

    .DESCRIPTION
    List all AD objects on a graphically selected OU, and export results into a CSV (CommonName and Description)

    .INPUTS
    User Interface

    .OUTPUTS
    Create transcript log file similar to $ScriptDir\[SCRIPTNAME]_[YYYY_MM_DD]_[HHhMMmSSs].log
    Create a list of AD Objects, similar to $ScriptDir\[SCRIPTNAME]_FoundADObjects_[YYYY_MM_DD]_[HHhMMmSSs].csv
   
   
    .NOTES
    Version:        1.1
    Author:         ALBERT Jean-Marc
    Creation Date:  15/07/2015
    Purpose/Change: 1.0 - 2015.07.15 - ALBERT Jean-Marc - Initial script development
                  1.1 - 2015.07.16 - ALBERT Jean-Marc - Optimize code (minification)
                  1.2 - 2015.07.17 - ALBERT Jean-Marc - Add sAMAccountName and homedirectory to "Get-ADObject parameters"
                  1.3 - 2015.09.12 - ALBERT Jean-Marc - Use functions defined on external .ps1 files
                  1.4 - 2016.10.31 - ALBERT Jean-Marc - Add missing .NET Assembly and replace $global:streamWriter per $script:streamWriter
                  
                                                  
    .SOURCES
  
  
    .EXAMPLE
    <None>
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Set-StrictMode -version Latest

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$scriptFile = $MyInvocation.MyCommand.Definition
$launchDate = get-date -f "yyyy.MM.dd-HHhmmmss"
$logDirectoryPath = $scriptPath + "\" + $launchDate
$buffer = "$scriptPath\bufferCommand.txt"
$fullScriptPath = (Resolve-Path -Path $buffer).Path

$loggingFunctions = "$scriptPath\logging\Logging.ps1"
$utilsFunctions = "$scriptPath\utilities\Utils.ps1"
$addsFunctions = "$scriptPath\utilities\ADDS.ps1"


#----------------------------------------------------------[Declarations]----------------------------------------------------------

$scriptName = [System.IO.Path]::GetFileName($scriptFile)
$scriptVersion = "1.4"

if(!(Test-Path $logDirectoryPath)) {
    New-Item $logDirectoryPath -type directory | Out-Null
}

$logFileName = "Log_" + $launchDate + ".log"
$logPathName = "$logDirectoryPath\$logFileName"

$script:streamWriter = New-Object System.IO.StreamWriter $logPathName

#Define CSV file export of AD Group objects
$CSVExportADObjects = "$ScriptPath\$scriptFile" + "_" + "FoundADObjects" + "_" + $launchDate + ".csv"


#-----------------------------------------------------------[Functions]------------------------------------------------------------

. $loggingFunctions
. $utilsFunctions
. $addsFunctions


#----------------------------------------------------------[Execution]----------------------------------------------------------

Start-Log -scriptName $scriptName -scriptVersion $scriptVersion -streamWriter $script:streamWriter
Clear-Host
Write-Host "================================================================================================"

# Prerequisites
if($adminFlag -eq $false){
    Write-Host "You have to launch this script with " -nonewline; Write-Host "local Administrator rights!" -f Red    
    $scriptPath = Split-Path $MyInvocation.InvocationName    
    $RWMC = $scriptPath + "\$scriptName.ps1"
    $ArgumentList = 'Start-Process -FilePath powershell.exe -ArgumentList \"-ExecutionPolicy Bypass -File "{0}"\" -Verb Runas' -f $RWMC;
    Start-Process -FilePath powershell.exe -ArgumentList $ArgumentList -Wait -NoNewWindow;    
    Stop-Script
}

Write-Host "================================================================================================"

 # Execute action with a progressbar
 Write-Progress -Activity "Starting GUI to select OU" -status "Running..." -id 1 
 # Call OnApplicationLoad to initialize
   If((OnApplicationLoad) -eq $true) {
       #Call the form
       Call-AD_OU_select_pff | Out-Null
       #Perform cleanup
       OnApplicationExit
                                     }

 # Show selected OU on a messagebox
   [System.Windows.Forms.MessageBox]::Show("Selected OU: $objSelectedOU" , "Information" , 0, [Windows.Forms.MessageBoxIcon]::Information)


 # Writing informations in the log file
 Write-Progress -Activity "Write informations in the log file" -status "Running..." -id 1
 
 End-Log -streamWriter $script:streamWriter
 Stop-ScriptMessageBox
 notepad $logPathName
 Clear-Host
