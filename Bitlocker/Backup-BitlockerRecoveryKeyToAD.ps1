$BitLockerDrives = Get-BitLockerVolume  {$_.KeyProtector  {$_.KeyProtectorType -eq 'RecoveryPassword'}}
Foreach ($BitLockerDrive in $BitLockerDrives) {
    $BitlockerDriveKey = $BitLockerDrive  select -exp KeyProtector  {$_.KeyProtectorType -eq 'RecoveryPassword'}
    Backup-BitLockerKeyProtector $BitlockerDrive.MountPoint $BitlockerDriveKey.KeyProtectorId
    Write-Host -ForegroundColor Green Backing up drive $BitlockerDrive ($($BitLockerDrive.VolumeType)), key $($BitlockerDriveKey.KeyProtectorId), password $($BitlockerDriveKey.RecoveryPassword)
}
