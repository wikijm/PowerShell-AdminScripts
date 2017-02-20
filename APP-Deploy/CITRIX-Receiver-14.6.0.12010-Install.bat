REM                                                        
REM  DATE:		Feb 20, 2017
REM  VERSION:	0.4
REM  INPUTS:	(1) Current version of Citrix package                
REM		        (2) Current Version of Citrix WebHelper module
REM		        (3) Package Location/Deployment Directory     
REM      		  (4) Script Logging Directory                  
REM		        (5) Package Installer Command Line Options
REM
REM  CHANGELOG: v0.1 - 25.01.2017 - Initial write
REM		v0.2 - 10.02.2017 - Add ',DesktopViewer' to (4)
REM		v0.3 - 16.02.2017 - Add a verification on WebHelper version (new (2) parameter)
REM				                  Redefine parameters
REM							               Add version of Citrix package and Citrix DesktopViewer module version on log
REM		v0.4 - 20.02.2017 - Add few steps to verify Citrix presence, then DesktopViewer version
REM							               If Citrix present, then check DesktopViewer version
REM							               If DesktopViewer not present/ with bad version, then uninstall and reinstall Citrix with DesktopViewer
REM
REM                                                        
REM  (1) Current version of Citrix package                        
Set DesiredVersion=14.6.0.12010
REM
REM  (2) Current version of Citrix WebHelper module
Set DesiredDesktopViewerVersion=14.6.0.12010
REM
REM  (3) Package Location/Deployment Directory             
Set DeployDirectory=\\contorso.dom\SYSVOL\contorso.dom\scripts\Citrix\
REM                                                        
REM  (4) SCRIPT LOGGING DIRECTORY                          
Set logshare=\\contorso.dom\SYSVOL\contorso.dom\scripts\Citrix\logs\
REM                                                    
REM  (5) PACKAGE INSTALLER COMMAND LINE OPTIONS            
Set CommandLineOptions=/noreboot /silent /EnableTracing=true /rcu /EnableCEIP=false /ALLOWADDSTORE=N ADDLOCAL=ICA_Client,WebHelper,DesktopViewer
REM
REM Start
echo %date% %time% %ComputerName% [INFO] The %0 script is running >> %logshare%global.log
REM Check if the machine is 64bit
IF NOT "%ProgramFiles(x86)%"=="" SET WOW6432NODE=WOW6432NODE\
REM 
REM Check if Citrix is installed
REM
IF EXIST "C:\Program Files (x86)\Citrix\ICA Client\redirector.exe" (goto CheckPackagePresence) else (goto CitrixNotFound)
REM Check if the Desired plug-in is installed
REM
:CheckPackagePresence
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\%WOW6432NODE%Citrix\PluginPackages\XenAppSuite\DesktopViewer" | findstr %DesiredDesktopViewerVersion%
if %errorlevel%==1 (goto PackageNotFound) else (goto End)
REM
:CitrixNotFound
echo %date% %time% %ComputerName% [WARNING] Citrix not detected, Begin Deployment >> %logshare%global.log
start /wait %DeployDirectory%\CitrixReceiver%DesiredVersion%.exe DONOTSTARTCC=1 %CommandLineOptions%
if %errorlevel% neq 0 (goto BadInstall) else (goto End)

:PackageNotFound
echo %date% %time% %ComputerName% [WARNING] DesktopViewer package not detected, Begin Deployment >> %logshare%global.log
start /wait %DeployDirectory%\CitrixReceiver%DesiredVersion%.exe /uninstall /noreboot /silent
start /wait %DeployDirectory%\CitrixReceiver%DesiredVersion%.exe DONOTSTARTCC=1 %CommandLineOptions%
if %errorlevel% neq 0 (goto BadInstall) else (goto End)

:BadInstall
echo %date% %time% %ComputerName% [ERROR] %DesiredVersion% %DesiredDesktopViewerVersion% Deployment ended with error code %errorlevel%. >>%logshare%global.log

:End
echo %date% %time% %ComputerName% [INFO] %DesiredVersion% %DesiredDesktopViewerVersion% Deployment ended with error code %errorlevel%. >>%logshare%global.log
exit 0
