:: Made by Ali
:: mobielstraat.nl
:: Version 2.1N
:: N = Non-interactive
:: I = Interactive

:start
@echo off
set errorlevel=0
set version=2.1N

fsutil dirty query %systemdrive%  >nul 2>&1 || (
echo ==== ERROR ====
echo.
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'
echo.
exit /B 1
)

set debug=n
if /I %debug% == y goto debugmenu

:versioncheck
curl "https://raw.githubusercontent.com/AliBello/latestofficeinstallercheck/main/latestversion?token=GHSAT0AAAAAACNQIEYEU7JW5RKWXMJPGWIEZON7QUA" -o %temp%/officebatchversion.txt
>nul find %version% %temp%/officebatchversion.txt && (
  goto cpucheck
) || (
  goto update
)

:update
cls
echo Old version (%version%) detected.
set /D wantupdate=Do you want to update? ([Y]/N) 
if /I %wantupdate% == n goto cpucheck
curl "https://raw.githubusercontent.com/AliBello/latestofficeinstallercheck/main/latestN.bat?token=GHSAT0AAAAAACNQIEYF6GKD4D2PHPFFIFGWZON627A" -o %temp%/latestN.bat
start %temp%/latestN.bat && exit

:cpucheck
if %PROCESSOR_ARCHITECTURE% == x86 goto installx86
if %PROCESSOR_ARCHITECTURE% == x64 goto installx64
cls
echo ==== ERROR ====
echo.
echo Your CPU architecture is unsupported.
echo.
exit /B 1

:installx64
cls
echo Online Office Installer For Ayasofya
echo Installing...
C:
cd %temp%
mkdir office-setup >nul
cd ./office-setup
curl https://files.mobielstraat.nl/api/public/dl/I0m2eh0U/Downloads/office/installer.exe -o "./installer.exe" >nul
curl https://files.mobielstraat.nl/api/public/dl/48w9Rpak/Downloads/office/settingsx64.xml -o "./settingsx64.xml" >nul
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
curl https://files.mobielstraat.nl/api/public/dl/I0m2eh0U/Downloads/office/installer.exe -o "./installer.exe" >nul
curl https://files.mobielstraat.nl/api/public/dl/Gkus_YbC/Downloads/office/settingsx86.xml -o "./settingsx86.xml" >nul
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
del officebatchversion.txt >nul
del %temp%/latestN.bat >nul
del office-setup/* >nul
y
rmdir office-setup >nul

:exit
cls
echo Online Office Installer For Ayasofya
echo Done!
taskkill /F /IM installer.exe /T
exit /B 0

:debugmenu
cls
echo ----------------------------------
echo | Goto options:                  |
echo | start                          |
echo | installx64                     |
echo | activatex64                    |
echo | installx86                     |
echo | activatex86                    |
echo | cleanup                        |
echo | exit                           |
echo ----------------------------------
echo.
set /P goto=Goto where? 
goto %goto%
