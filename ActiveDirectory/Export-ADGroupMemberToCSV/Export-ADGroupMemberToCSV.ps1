$GroupName = ""
Get-ADGroupMember -Identity $GroupName | Select name,objectClass,SamAccountName | Export-CSV C:\exportfile.csv -NoType -Force


