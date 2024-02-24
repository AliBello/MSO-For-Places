:: Made by Ali
:: mobielstraat.nl
:: Main repository at https://github.com/AliBello/MSO-For-Ayasofya-Arnhem
:: Office Downloader, installer, and activator made for Ayasofya Arnhem
:: N = Non-interactive
:: I = Interactive

:start
@echo off
cls
set errorlevel=0
set gotoerror=0
set version=2.2.3
set base=https://raw.githubusercontent.com/AliBello/MSO-For-Ayasofya-Arnhem/main/

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
curl %base%latestversion -o %temp%/officebatchversion.txt >nul
>nul find "%version%" %temp%\officebatchversion.txt && (
  goto confirm
) || (
  goto update
)

:update
cls
if %updateprompt% == n goto confirm
echo Old version (%version%) detected.
set wantupdate=y
set /P wantupdate=Do you want to update? ([Y]/N) 
if /I %wantupdate% == n goto confirm
curl %base%latest.bat -o %temp%/latestofficeinstaller.bat >nul
start %temp%\latestofficeinstaller.bat && exit

:confirm
cls
set errorlevel=0
echo This will install Office 2021 on this computer.
echo Warning: If you install office via this batch file, it will mark Ayasofya Arnhem as the orginization for office.
set proceed=y
if /I "%interactive%"=="n" goto cpucheck
set /P proceed="Proceed? (Y/N) "
if /I "%debug%"=="y" goto debugmenu
if /I "%interactive%"=="n" goto cpucheck
if /I "%proceed%"=="y" goto cpucheck
if /I "%proceed%"=="n" exit /B 0
if /I "%proceed%"=="debug" goto debugmenu
cls
echo ==== ERROR ====
echo.
echo Please choose an option
echo.
goto confirm

:cpucheck
if %PROCESSOR_ARCHITECTURE% == x86 goto installx86
if %PROCESSOR_ARCHITECTURE% == AMD64 goto installx64
cls
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
curl %base%installer.exe -o "./installer.exe" >nul
curl %base%/settingsx64.xml -o "./settingsx64.xml" >nul
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
curl %base%installer.exe -o "./installer.exe" >nul
curl %base%settingsx86.xml -o "./settingsx86.xml" >nul
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

:exit
cls
echo Online Office Installer For Ayasofya
echo Done!
if /I %interactive% == y echo Press any key to exit...
if /I %interactive% == y pause >nul
exit /B 0

:debugmenu
cls
if %gotoerror% == 1 echo ==== ERROR ====
if %gotoerror% == 1 echo.
if %gotoerror% == 1 echo Option not vaild.
if %gotoerror% == 1 echo.
echo ----------------------------------
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
echo ^| exit                           ^|
echo ----------------------------------
echo.
set /P goto=Goto where? 
goto %goto%
set gotoerror=1
goto debugmenu
