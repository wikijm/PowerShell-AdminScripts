$DomoticzApiProtocol = 'http'
$DomoticzApiIP = '192.168.0.xx'
$DomoticzApiPort = '8080'
$DomoticzDeviceID = 'x'

function Get-LightSwitchStatus {
    Param([Parameter(Mandatory=$true)][string]$DomoticzDeviceID)
    $DomoticzRequestURL = $DomoticzApiProtocol + '://' + $DomoticzApiIP + ':' + $DomoticzApiPort + '/json.htm?type=devices&rid=' + $DomoticzDeviceID
    $DomoticzRequestResult = Invoke-WebRequest $DomoticzRequestURL | ConvertFrom-Json | Select -expand result

    Write-Host $DomoticzRequestResult.HardwareName
    Write-Host $DomoticzRequestResult.Type
    Write-Host $DomoticzRequestResult.Status
    Write-Host $DomoticzRequestResult.LastUpdate
}

Get-LightSwitchStatus $DomoticzDeviceID
