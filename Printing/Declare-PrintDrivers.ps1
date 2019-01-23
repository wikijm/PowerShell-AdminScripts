#Define drivers that we want to install
$DriverBase = @'
Xerox GPD PS V5.645.5.0
Xerox GPD PS V5.645.5.0
Xerox GPD PCL6 V5.645.5.0
Xerox GPD PCL6 V5.645.5.0
Xerox GPD PCL V5.645.5.0
Xerox GPD PCL V5.645.5.0
Xerox Global Print Driver V4 PS
Xerox Global Print Driver V4 PS
Xerox Global Print Driver V4 PCL6
Xerox Global Print Driver V4 PCL6
Xerox Global Print Driver PS
Xerox Global Print Driver PS
Xerox Global Print Driver PCL6
Xerox Global Print Driver PCL6
Xerox Global Print Driver PCL
Xerox Global Print Driver PCL
'@


#Installation of printers in Windows
$DriversSourcePath = "\\NETWORKSHARE\C$\Drivers\Printers\"
Get-ChildItem -Filter "*.inf" -Path $DriversSourcePath -Recurse | ForEach-Object { PnPutil.exe -i -a $_.FullName }

#Inject drivers (x86 and x64) to the spool service
foreach ($DriverName in $DriverBase) {
    Add-PrinterDriver -Name $DriverName -PrinterEnvironment "Windows NT x86"
    Add-PrinterDriver -Name $DriverName -PrinterEnvironment "Windows x64"
}
