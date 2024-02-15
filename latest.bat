:: Made by Ali
:: mobielstraat.nl
:: Version 2.2
:: N = Non-interactive
:: I = Interactive

:start
@echo on
set errorlevel=0
set version=2.2

set interactive=y
set debug=n

fsutil dirty query %systemdrive%  >nul 2>&1 || (
echo ==== ERROR ====
echo.
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'
echo.
echo Press any key to exit...
pause >nul
exit /B 1
)

:versioncheck
curl "https://raw.githubusercontent.com/AliBello/MSO-For-Ayasofya-Arnhem/main/latestversion" -o %temp%/officebatchversion.txt
 find "%version%" %temp%\officebatchversion.txt && (
  goto confirm
) || (
  goto update
)

:update
cls
echo Old version (%version%) detected.
set /P wantupdate=Do you want to update? ([Y]/N) 
if /I %wantupdate% == n goto confirm
curl "https://raw.githubusercontent.com/AliBello/MSO-For-Ayasofya-Arnhem/main/latest.bat" -o %temp%/latestofficeinstaller.bat
start %temp%\latestofficeinstaller.bat && exit

:confirm
cls
set errorlevel=0
echo This will install Office 2021 on this computer.
echo Warning: If you install office via this batch file, it will mark Ayasofya Arnhem as the orginization for office.
set cancel=null
if /I %interactive% == y (
set /P cancel="Proceed? (Y/N) "
 )
if /I %debug% == y goto debugmenu
if /I %interactive% == n goto cpucheck
if /I %cancel% == y goto cpucheck
if /I %cancel% == n exit /B 0
if /I %cancel% == debug goto debugmenu
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
curl https://files.mobielstraat.nl/api/public/dl/I0m2eh0U/Downloads/office/installer.exe -o "./installer.exe" 
curl https://files.mobielstraat.nl/api/public/dl/48w9Rpak/Downloads/office/settingsx64.xml -o "./settingsx64.xml" 
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
curl https://files.mobielstraat.nl/api/public/dl/I0m2eh0U/Downloads/office/installer.exe -o "./installer.exe" 
curl https://files.mobielstraat.nl/api/public/dl/Gkus_YbC/Downloads/office/settingsx86.xml -o "./settingsx86.xml" 
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
del .\latestofficeinstaller.bat >nul
del office-setup\* >nul
y
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
echo ----------------------------------
echo | Goto options:                  |
echo | start                          |
echo | versioncheck                   |
echo | update                         |
echo | confirm                        |
echo | cpucheck                       |
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
