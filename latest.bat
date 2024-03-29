:: Made by Ali
:: mobielstraat.nl
:: Main repository at https://github.com/AliBello/MSO-For-Ayasofya-Arnhem
:: Office Downloader, installer, and activator made for Ayasofya Arnhem

:start
@echo off
cls
set errorlevel=0
set gotoerror=0
set version=2.2.5
set rawbase=https://raw.githubusercontent.com/AliBello/MSO-For-Ayasofya-Arnhem/main
set base=https://github.com/AliBello/MSO-For-Ayasofya-Arnhem

:: Options
:: Text options
set kms=server.mc.mobielstraat.nl
:: Toggle options, valid options are "y" and "n"
set interactive=y
set debug=n
set updateprompt=y

CLS

:initadmin
setlocal DisableDelayedExpansion
set "batchPath=%~0"
for %%k in (%0) do set batchName=%%~nk
set "vbsGetPrivileges=%temp%\OEgetPriv_%batchName%.vbs"
setlocal EnableDelayedExpansion

:checkprivileges
NET FILE 1>NUL 2>NUL
if '%errorlevel%' == '0' ( goto gotprivileges ) else ( goto getprivileges )

:getprivileges
if '%1'=='ELEV' (echo ELEV & shift /1 & goto gotprivileges)

ECHO Set UAC = CreateObject^("Shell.Application"^) > "%vbsGetPrivileges%"
ECHO args = "ELEV " >> "%vbsGetPrivileges%"
ECHO For Each strArg in WScript.Arguments >> "%vbsGetPrivileges%"
ECHO args = args ^& strArg ^& " "  >> "%vbsGetPrivileges%"
ECHO Next >> "%vbsGetPrivileges%"
ECHO UAC.ShellExecute "!batchPath!", args, "", "runas", 1 >> "%vbsGetPrivileges%"
"%SystemRoot%\System32\WScript.exe" "%vbsGetPrivileges%" %*
echo ==== ERROR ====
echo.
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'
echo.
echo Press any key to exit...
pause >nul
exit /B 1

:gotprivileges
setlocal & pushd .
cd /d %~dp0
if '%1'=='ELEV' (del "%vbsGetPrivileges%" 1>nul 2>nul  &  shift /1)

:versioncheck
curl %rawbase%/dependencies/latestversion -o %temp%/officebatchversion.txt >nul
>nul find "%version%" %temp%\officebatchversion.txt && (
  goto confirm
) || (
  goto update
)

:update
if %updateprompt% == n goto confirm
cls
if /I %version% == debug (
echo  ---------------------------------
echo ^| -----  Challenge complete^!       ^|
echo ^| ^\---^/                           ^|
echo ^|  ^\-^/   How did we get here?     ^|
echo  ---------------------------------
echo.
echo Looks like you're using a debug version of the software,
echo These versions can be unstable and are intended for developement and debugging purposes.
echo Usage of this version is not reccomended.
echo.
 )
echo Online Office Installer For Ayasofya
echo.
echo Old version (%version%) detected.
set wantupdate=y
set /P wantupdate=Do you want to update? ([Y]/N) 
if /I %wantupdate% == n goto confirm
curl %rawbase%/latest.bat -s -o %temp%/latestofficeinstaller.bat >nul
if %errorlevel% == 1 echo Download failed, please update manually.
if %errorlevel% == 1 echo Download is at %base%
if %errorlevel% == 1 echo Press any key to exit...
if %errorlevel% == 1 pause >nul
if %errorlevel% == 1 exit /B 1
start %temp%\latestofficeinstaller.bat && exit

:confirm
cls
set errorlevel=0
echo This will install Office 2021 on this computer.
echo Warning: If you install office using this script, it will mark Ayasofya Arnhem as the orginization for office.
set proceed=
if /I "%interactive%"=="n" goto cpucheck
set /P proceed="Proceed? (Y/N) "
if /I "%proceed%"=="y" goto cpucheck
if /I "%proceed%"=="n" exit /B 0
if /I "%proceed%"=="debug" (
  goto debugmenu
  set debugconfirm=Y
)
cls
echo Online Office Installer For Ayasofya
echo.
echo ==== ERROR ====
echo.
echo Please choose an option
echo.
goto confirm

:cpucheck
if /I "%debug%"=="y" goto debugmenu
if %PROCESSOR_ARCHITECTURE% == x86 goto installx86
if %PROCESSOR_ARCHITECTURE% == AMD64 goto installx64
cls
echo Online Office Installer For Ayasofya
echo.
echo ==== ERROR ====
echo.
echo Your CPU architecture is unsupported.
echo.
echo Press any key to exit...
cd %temp% >nul
del .\officebatchversion.txt >nul
del .\latestofficeinstaller.bat >nul
pause >nul
exit /B 1

:installx64
cls
echo Online Office Installer For Ayasofya
echo Installing...
C:
cd %temp%
mkdir office-setup >nul
cd ./office-setup
curl %rawbase%/dependencies/installer.exe -s -o "./installer.exe" >nul
curl %rawbase%/dependencies/settingsx64.xml -s -o "./settingsx64.xml" >nul
if %errorlevel% == 1 echo Download failed, please update.
if %errorlevel% == 1 echo Press any key to exit...
if %errorlevel% == 1 pause >nul
if %errorlevel% == 1 exit /B 1
START /W installer.exe /configure settingsx64.xml >nul
if %errorlevel% == 1 exit /B 1
goto activatex64

:installx86
cls
echo Online Office Installer For Ayasofya
echo Installing...
C:
cd %temp%
mkdir office-setup >nul
cd ./office-setup
curl %rawbase%/dependencies/installer.exe -s -o "./installer.exe" >nul
curl %rawbase%/dependencies/settingsx86.xml -s -o "./settingsx86.xml" >nul
if %errorlevel% == 1 echo Download failed, please update.
if %errorlevel% == 1 echo Press any key to exit...
if %errorlevel% == 1 pause >nul
if %errorlevel% == 1 exit /B 1
START /W installer.exe /configure settingsx86.xml >nul
if %errorlevel% == 1 exit /B 1
goto activatex86

:activatex64
cls
echo Online Office Installer For Ayasofya
echo Activating...
cd \Program Files\Microsoft Office\Office16 >nul
cscript ospp.vbs /sethst:server.mc.mobielstraat.nl >nul
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul
cscript ospp.vbs /act >nul
if %errorlevel% == 1 exit /B 1
goto cleanup

:activatex86
cls
echo Online Office Installer For Ayasofya
echo Activating...
cd \Program Files (x86)\Microsoft Office\Office16 >nul
cscript ospp.vbs /sethst:server.mc.mobielstraat.nl >nul
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul
cscript ospp.vbs /act >nul
if %errorlevel% == 1 exit /B 1
goto cleanup

:cleanup
cls
echo Online Office Installer For Ayasofya
echo Cleaning up...
cd %temp%
del .\officebatchversion.txt >nul
del office-setup\installer.exe >nul
del office-setup\settingsx86 >nul
del office-setup\settingsx64 >nul
rmdir office-setup >nul
set interactive=
set debug=
set updateprompt=

:exit
cls
echo Online Office Installer For Ayasofya
echo Done
if /I %interactive% == y echo Press any key to exit...
if /I %interactive% == y pause >nul
exit /B 0

:debugmenu
:maindebug
cls
echo Welcome to the debug menu.
echo These are menus you can go to.
echo.
echo ---------------------------------------------------
echo ^| 2. goto              ^| Created by Ali Bal       ^|
echo ^| 3. set               ^| mobielstraat.nl          ^|
echo ^|                      ^|                          ^|
echo ^|                      ^| Made for Ayasofya Arnhem ^|
echo ^| 1. exit              ^| arnhemayasofya.nl        ^|
echo ---------------------------------------------------
echo.
choice /n /m "Please select. " /C "123"
if ERRORLEVEL 3 goto setmenu
if ERRORLEVEL 2 goto gotomenu
if ERRORLEVEL 1 exit /B 0
set errorlevel
echo You discovered a bug!
echo Please create an issue on github and tell what happened, and how to reproduce it
echo Repository: %base%
pause
exit /B 1

:gotomenu
cls
if %gotoerror% == 1 echo ==== ERROR ====
if %gotoerror% == 1 echo.
if %gotoerror% == 1 echo Option not vaild.
if %gotoerror% == 1 echo.
set gotoerror=0
echo  --------------------------------
echo ^| Goto options:                  ^|
echo ^| start                          ^|
echo ^| initadmin                      ^|
echo ^| checkprivileges                ^|
echo ^| getprivileges                  ^|
echo ^| gotprivileges                  ^|
echo ^| versioncheck                   ^|
echo ^| update                         ^|
echo ^| confirm                        ^|
echo ^| cpucheck                       ^|
echo ^| installx64                     ^|
echo ^| activatex64                    ^|
echo ^| installx86                     ^|
echo ^| activatex86                    ^|
echo ^| cleanup                        ^|
echo ^| maindebug                      ^|
echo ^| gotomenu ^(this menu^)           ^|
echo ^| setmenu                        ^|
echo ^| exit                           ^|
echo  --------------------------------
echo.
set /P goto=Goto where? 
goto %goto%
set gotoerror=1
goto gotomenu
echo You discovered a bug!
echo Please create an issue on github and tell what happened, and how to reproduce it
echo Repository: %base%
pause
exit /B 1

:setmenu
echo  --------------------------------
echo ^| set options:                   ^|
echo ^| errorlevel                     ^|
echo ^| gotoerror                      ^|
echo ^| version                        ^|
echo ^| rawbase                        ^|
echo ^| base                           ^|
echo ^| kms                            ^|
echo ^| interactive                    ^|
echo ^| debug                          ^|
echo ^| updateprompt                   ^|
echo ^| wantupdate                     ^|
echo ^| proceed                        ^|
echo ^| debugconfirm                   ^|
echo ^| PROCESSOR_ARCHITECTURE         ^|
echo ^| setquery                       ^|
echo ^|                                ^|
echo ^| exit (will go to debug menu)   ^|
echo ^|                                ^|
echo ^| You can also enter set options ^|
echo ^| that aren't in this list.      ^|
echo  --------------------------------
echo.
set /P setquery=What set do you want to query? 
if "%setquery%" == "exit" goto maindebug
set setquery
goto setmenu
echo You discovered a bug!
echo Please create an issue on github and tell what happened, and how to reproduce it
echo Repository: %base%
pause
exit /B 1