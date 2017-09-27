<#
#USAGE
PS C:\temp\> ./Get-MicrosoftLicenceStatus.ps1 

# RESULT
  Name                                            ApplicationId                        LicenseStatus
  ----                                            -------------                        -------------
  Windows(R), Professional edition                55c92734-d682-4d71-983e-d6ec3f16059f Licensed     
  Office 16, Office16ProPlusVL_KMS_Client edition 0ff1ce15-a989-479d-af46-f275c6370663 Licensed     
#>

$lstat = DATA {
    ConvertFrom-StringData -StringData @'
    0 = Unlicensed
    1 = Licensed
    2 = OOB Grace
    3 = OOT Grace
    4 = Non-Genuine Grace
    5 = Notification
    6 = Extended Grace
'@}

function Get-MicrosoftLicenseStatus {
    param (
        [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$computername="$env:COMPUTERNAME"
    )

    PROCESS {
        Get-WmiObject SoftwareLicensingProduct -ComputerName $computername |
        where {$_.PartialProductKey} | select Name, ApplicationId,@{N="LicenseStatus"; E={$lstat["$($_.LicenseStatus)"]} }
    }
}

Get-MicrosoftLicenseStatus
