$GPOSettingsFolderPath = "D:\Scripts\AD-ExportGPOSettings\"
$GPOSettingsResultFolderPath = $GPOSettingsFolderPath + "\Results\"

cd $GPOSettingsFolderPath

Get-GPO -All | % {$_.GenerateReport('html') | Out-File "$($_.DisplayName).htm"}
Get-GPO -All | % {$_.GenerateReport('xml') | Out-File "$($_.DisplayName).xml"}

Get-ChildItem *.xml | ForEach-Object { $NewName = $_.BaseName + "_" + $_.LastWriteTime.toString("yyyy.MM.dd.HH.mm") + ".xml" ; Rename-Item -Path $_.FullName -newname $NewName }
Get-ChildItem *.htm | ForEach-Object { $NewName = $_.BaseName + "_" + $_.LastWriteTime.toString("yyyy.MM.dd.HH.mm") + ".htm" ; Rename-Item -Path $_.FullName -newname $NewName }

If (!(Test-Path -Path $GPOSettingsResultFolderPath)) {
    New-Item -ItemType "Directory" -Path $GPOSettingsResultFolderPath -Force
}

Get-ChildItem *.xml -Recurse | Move-Item -Destination $GPOSettingsResultFolderPath
Get-ChildItem *.htm -Recurse | Move-Item -Destination $GPOSettingsResultFolderPath
