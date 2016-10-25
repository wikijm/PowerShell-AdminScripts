#requires -version 3
<#
.SYNOPSIS
  List AD Computers that have a pending reboot after Windows Update KB install
  
.DESCRIPTION
  List AD Computers that have a pending reboot after Windows Update KB install, thanks to specific values on two registry key, then export it to .CSV file
  
.INPUTS
  <None>
.OUTPUTS
  Create transcript log file similar to $ScriptDir\[SCRIPTNAME]_[YYYY_MM_DD]_[HHhMMmSSs].log
  Create a list of AD user objects, similar to $ScriptDir\[SCRIPTNAME]_SUBJECT_[YYYY_MM_DD]_[HHhMMmSSs].csv
   
   
.NOTES
  Version:        1.0
  Author:         ALBERT Jean-Marc
  Creation Date:  25/10/2016 (DD/MM/YYYY)
  Purpose/Change: 1.0 - 2016.10.25 - ALBERT Jean-Marc - Initial script development
                                      
                                                  
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

Function Get-PendingReboot($computer = '.') {
    $hkey       = 'LocalMachine';
    $path_server    = 'SOFTWARE\Microsoft\ServerManager';
    $path_control   = 'SYSTEM\CurrentControlSet\Control';
    $path_session   = join-path $path_control 'Session Manager';
    $path_name  = join-path $path_control 'ComputerName';
    $path_name_old  = join-path $path_name 'ActiveComputerName';
    $path_name_new  = join-path $path_name 'ComputerName';
    $path_wsus = 'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired';
    $pending_rename = 'PendingFileRenameOperations';
    $pending_rename_2   = 'PendingFileRenameOperations2';
    $attempts   = 'CurrentRebootAttempts';
    $computer_name  = 'ComputerName';
 
    $num_attempts   = 0;
    $name_old   = $null;
    $name_new   = $null;
 
    $reg= [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($hkey, $computer);
 
    $key_session    = $reg.OpenSubKey($path_session);
    if ($key_session -ne $null) {
        $session_values = @($key_session.GetValueNames());
        $key_session.Close() | out-null;
    }
 
    $key_server = $reg.OpenSubKey($path_server);
    if ($key_server -ne $null) {
        $num_attempts = $key_server.GetValue($attempts);
        $key_server.Close() | out-null;
    }
 
    $key_name_old   = $reg.OpenSubKey($path_name_old);
    if ($key_name_old -ne $null) {
        $name_old = $key_name_old.GetValue($computer_name);
        $key_name_old.Close() | out-null;
 
        $key_name_new   = $reg.OpenSubKey($path_name_new);
        if ($key_name_new -ne $null) {
            $name_new = $key_name_new.GetValue($computer_name);
            $key_name_new.Close() | out-null;
        }
    }
     
        $key_wsus = $reg.OpenSubKey($path_wsus);
        if ($key_wsus -ne $null) {
        $wsus_values = @($key_wsus.GetValueNames());
        if ($wsus_values) {
        $wsus_rbpending = $true
        } else {
        $wsus_rbpending = $false
        }
        $key_wsus.Close() | out-null;
        }
 
    $reg.Close() | out-null;
    #modified return section:   
        if ( `
        (($session_values -contains $pending_rename) -or ($session_values -contains $pending_rename_2)) `
        -or (($num_attempts -gt 0) -or ($name_old -ne $name_new)) `
        -or ($wsus_rbpending)) {
        return $true;
        }
        else {
        return $false;
        }
}


#----------------------------------------------------------[Execution]----------------------------------------------------------

Start-Log -scriptName $scriptName -scriptVersion $scriptVersion -streamWriter $global:streamWriter
Write-Host "================================================================================================"
Write-Host "List AD computers with a pending reboot (after Windows Update KB install)"
Write-Host "================================================================================================"


Import-Module ActiveDirectory

Write-Progress -Activity "List all AD Computer on Windows 7" -status "Running..." -id 1
$Computers = Get-ADComputer -Filter * -Property OperatingSystem | Where-Object {$_.OperatingSystem -match 'Windows 7'} | Select Name

Write-Progress -Activity "Check on Windows 7 AD computers if there is a pending reboot " -status "Running..." -id 1
ForEach ($Computer in $Computers) {
    $ComputerName = $Computer.Name
    $IsRebootPending = $(Get-PendingReboot $ComputerName)
    $ComputerName + ';' + $IsRebootPending | Out-File -Append $ExportPathName
}

#Writing informations in the log file
Write-Progress -Activity "Write informations in the log file" -status "Running..." -id 1
Write-Log -streamWriter $global:streamWriter -infoToLog "If you don't see any lines between start & end of logfile, it means that everything's OK"
End-Log -streamWriter $global:streamWriter
notepad $logPathName
#cls
