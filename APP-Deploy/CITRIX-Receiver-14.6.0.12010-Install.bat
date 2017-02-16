REM                                                        
REM  DATE:		Feb 16, 2017
REM  VERSION:	0.3
REM  INPUTS:	(1) Current version of Citrix package                
REM		        (2) Current Version of Citrix WebHelper module
REM		        (3) Package Location/Deployment Directory     
REM      		  (4) Script Logging Directory                  
REM		        (5) Package Installer Command Line Options
REM
REM  CHANGELOG: v0.1 - 25.01.2017 - Initial write
REM		           v0.2 - 10.02.2017 - Add ',DesktopViewer' to (4)
REM	           	v0.3 - 16.02.2017 - Add a verification on WebHelper version (new (2) parameter)
REM				                             Redefine parameters
REM						                          	Add version of Citrix package and Citrix DesktopViewer module version on log
REM
REM                                                        
REM  (1) Current version of Citrix package                        
Set DesiredVersion=14.6.0.12010
REM
REM  (2) Current version of Citrix WebHelper module
Set DesiredDesktopViewerVersion=14.6.0.12010
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
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\%WOW6432NODE%Citrix\PluginPackages\XenAppSuite\DesktopViewer" | findstr %DesiredWebHelperVersion%
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
echo %date% %time% %ComputerName% [ERROR] %DesiredVersion% %DesiredDesktopViewerVersion% Deployment ended with error code %errorlevel%. >>%logshare%global.log

:End
echo %date% %time% %ComputerName% [INFO] %DesiredVersion% %DesiredDesktopViewerVersion% Deployment ended with error code %errorlevel%. >>%logshare%global.log
exit 0
