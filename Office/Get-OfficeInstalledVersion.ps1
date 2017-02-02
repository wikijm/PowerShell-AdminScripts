$version = 0
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computer)
$reg.OpenSubKey('software\Microsoft\Office').GetSubKeyNames() |% {
    if ($_ -match '(\d+)\.') {
        if ([int]$matches[1] -gt $version) {
            $version = $matches[1] 
        }
    }   
}

if ($version) {
    "$computer : found $version" 
}

else {
    "Could not find version for $computer" 
}
