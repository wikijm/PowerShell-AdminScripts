REM                                       
REM  DATE:		Feb 10, 2017
REM  VERSION:	0.1
REM  INPUTS:	(1) Current version of package                
REM		(2) Package Location/Deployment Directory     
REM		(3) Script Logging Directory                  
REM		(4) Package Installer Command Line Options
REM
REM  USAGE: Launch manually or use on 'Computer Configuration\Windows Settings\Scripts\Startup' GPO
REM
REM  CHANGELOG: v0.1 - 10.02.2017 - Initial write
REM
REM                                                        
REM  (1) Current version of package                        
Set DesiredVersion=14.6.0.12010
REM                                                        
REM  (2) Package Location/Deployment Directory             
Set DeployDirectory=\\contorso.dom\apps\Citrix\
REM                                                        
REM  (3) SCRIPT LOGGING DIRECTORY                          
Set logshare=\\contorso.dom\apps\Citrix\Receiver\logs\
REM                                                        
REM  (4) PACKAGE INSTALLER COMMAND LINE OPTIONS            
Set CommandLineOptions=/noreboot /silent /EnableTracing=true /rcu /EnableCEIP=false /ALLOWADDSTORE=N ADDLOCAL=ICA_Client,WebHelper,DesktopViewer

REM Start
echo %date% %time% %ComputerName% [INFO] The %0 script is running >> %logshare%global.log
REM Check if the machine is 64bit
IF NOT "%ProgramFiles(x86)%"=="" SET WOW6432NODE=WOW6432NODE\
 
REM Check if the Desired plug-in is installed
REM
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\%WOW6432NODE%Citrix\PluginPackages\XenAppSuite\ICA_Client" | findstr %DesiredVersion%
if %errorlevel%==1 (goto NotFound) else (goto End)
REM
REM If 1 was returned, the registry query found the Desired Version is not installed.
REM
REM *
REM Deployment begins here
REM *
REM
:NotFound
echo %date% %time% %ComputerName% [WARNING] Package not detected, Begin Deployment >> %logshare%global.log
start /wait %DeployDirectory%\CitrixReceiver%DesiredVersion%.exe DONOTSTARTCC=1 %CommandLineOptions%
if %errorlevel% neq 0 (goto BadInstall) else (goto End)

:BadInstall
echo %date% %time% %ComputerName% [ERROR] Deployment ended with error code %errorlevel%. >>%logshare%global.log

:End
echo %date% %time% %ComputerName% [INFO] Deployment ended with error code %errorlevel%. >>%logshare%global.log
exit 0
