$DomoticzApiProtocol = 'http'
$DomoticzApiIP = '192.168.0.XX'
$DomoticzApiPort = '8080'
$DomoticzDeviceID = 'X'

function Get-LightSwitchStatus {
    Param([Parameter(Mandatory=$true)][string]$DomoticzDeviceID)
    $DomoticzRequestURL = $DomoticzApiProtocol + '://' + $DomoticzApiIP + ':' + $DomoticzApiPort + '/json.htm?type=devices&rid=' + $DomoticzDeviceID
    $DomoticzRequestResult = Invoke-WebRequest $DomoticzRequestURL | ConvertFrom-Json

Write-Host $DomoticzRequestResult.result.HardwareName
Write-Host $DomoticzRequestResult.result.Type
Write-Host $DomoticzRequestResult.result.Status
Write-Host $DomoticzRequestResult.result.LastUpdate
}

Get-LightSwitchStatus X
