function OnApplicationLoad {
	return $true #return true for success or false for failure
                           }
function OnApplicationExit {
	$script:ExitCode = 0 #Set the exit code for the Packager
                           }

function Set-RegistryKey($computername, $parentKey, $nameRegistryKey, $valueRegistryKey) {
    try{    
        $remoteBaseKeyObject = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$computername)     
        $regKey = $remoteBaseKeyObject.OpenSubKey($parentKey,$true)
        $regKey.Setvalue("$nameRegistryKey", "$valueRegistryKey", [Microsoft.Win32.RegistryValueKind]::DWORD) 
        $remoteBaseKeyObject.close()
    }
    catch {
        $_.Exception
    }
}

function Disable-UAC($computername) {
    $parentKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System"
    $nameRegistryKey = "LocalAccountTokenFilterPolicy"
    $valueRegistryKey = "1"

    $objReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computername)
    $objRegKey= $objReg.OpenSubKey($parentKey)
    $test = $objRegkey.GetValue($nameRegistryKey)
    if($test -eq $null){    
        Set-RegistryKey $computername $parentKey $nameRegistryKey $valueRegistryKey     
        Write-Host "Registry key setted, you have to reboot the remote computer" -foregroundcolor "magenta"
        Stop-Script
    }
    else {
        if($test -ne 1){
            Set-RegistryKey $computername $parentKey $nameRegistryKey $valueRegistryKey     
            Write-Host "Registry key setted, you have to reboot the remote computer" -foregroundcolor "magenta"
            Stop-Script
        }
    }
}

function CreateDirectoryIfNeeded ( [string] $directory ) {
	if (!(Test-Path -Path $directory -type "Container")) {
		New-Item -type directory -Path $directory > $null
	}
}

function Run-WmiRemoteProcess {
    Param(
        [string]$computername=$env:COMPUTERNAME,
        [string]$cmd=$(Throw "You must enter the full path to the command which will create the process."),
        [int]$timeout = 0
    )
 
    Write-Host "Process to create on $computername is $cmd"
    [wmiclass]$wmi="\\$computername\root\cimv2:win32_process"
    # Exit if the object didn't get created
    if (!$wmi) {return}
 
    try{
    $remote=$wmi.Create($cmd)
    }
    catch{
        $_.Exception
    }
    $test =$remote.returnvalue
    if ($remote.returnvalue -eq 0) {
        Write-Host ("Successfully launched $cmd on $computername with a process id of " + $remote.processid)
    } else {
        Write-Host ("Failed to launch $cmd on $computername. ReturnValue is " + $remote.ReturnValue)
    }    
    return
}

function Stop-Script () {   
    Begin{
        Write-Log -streamWriter $global:streamWriter -infoToLog "--- Script terminating ---"
    }
    Process{        
        "Script terminating..." 
        Write-Host "================================================================================================"
        End-Log -streamWriter $global:streamWriter       
        Exit
    }
}

Function Stop-ScriptMessageBox () {
 # MessageBox who inform of the end of the process
   [System.Windows.Forms.MessageBox]::Show(
"Process done.
The log file will be opened when click on 'OK' button.
Please, check the log file for further informations.
" , "End of process" , 0, [Windows.Forms.MessageBoxIcon]::Information)
}

function Test-InternetConnection {
    if(![Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet){
        Write-Host "The script need an Internet Connection to run" -f Red    
        Stop-Script
    }
}

function Test-LocalAdminRights {
    $myComputer = Get-WMIObject Win32_ComputerSystem | Select-Object -ExpandProperty name
    $myUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $amIAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent())
    $adminFlag = $amIAdmin.IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    if($adminFlag -eq $true){
        $adminMessage = " with administrator rights on " 
    }
    else {
        $adminMessage = " without administrator rights on "
    }

    Write-Host "RWMC runs with user " -nonewline; Write-Host $myUser.Name -f Red -nonewline; Write-Host $adminMessage -nonewline; Write-Host $myComputer -f Red -nonewline; Write-Host " computer"
    return $adminFlag
}