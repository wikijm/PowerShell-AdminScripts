#Load the required assemblies
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
#Remove any registered events related to notifications
Remove-Event BalloonClicked_event -ea SilentlyContinue
Unregister-Event -SourceIdentifier BalloonClicked_event -ea silentlycontinue
Remove-Event BalloonClosed_event -ea SilentlyContinue
Unregister-Event -SourceIdentifier BalloonClosed_event -ea silentlycontinue
Remove-Event Disposed -ea SilentlyContinue
Unregister-Event -SourceIdentifier Disposed -ea silentlycontinue

#Get domain
$MachineDomain = (Get-WmiObject Win32_ComputerSystem).Domain


#Create the notification object
$notification = New-Object System.Windows.Forms.NotifyIcon 
#Define various parts of the notification
$notification.Icon = [System.Drawing.SystemIcons]::Error
$notification.BalloonTipTitle = "Connection error on $MachineDomain"
$notification.BalloonTipIcon = "Error"
$title = "Credentials have expired, please reauthenticate by clicking on this message"
$notification.BalloonTipText = $title

#Make balloon tip visible when called
$notification.Visible = $True

## Register a click event with action to take based on event
#Balloon message clicked
register-objectevent $notification BalloonTipClicked BalloonClicked_event -Action {
    Start-Process 'iexplore.exe' -ArgumentList 'http://www.google.com' -WindowStyle Maximized -Verb Open
    #Get rid of the icon after action is taken
    $notification.Dispose()
    } | Out-Null

#Balloon message closed
register-objectevent $notification BalloonTipClosed BalloonClosed_event -Action {$notification.Dispose()} | Out-Null

#Call the balloon notification
$notification.ShowBalloonTip(1000)

#Add a Sleep to avoid end of PowerShell process before user click on balloon
Sleep(1)
