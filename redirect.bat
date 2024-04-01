:: Made by Ali
:: mobielstraat.nl
:: Main repository at https://github.com/AliBello/MSO-For-Places
:: Redirector

set version=1
set redirected=1
set settings=https://raw.githubusercontent.com/AliBello/MSO-For-Places/Expo
set rawbase=https://raw.githubusercontent.com/AliBello/MSO-For-Places/main

:: Valid options are Y and N
set interactive=y
set debug=n
set updateprompt=y

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
echo Online Office Redirector
echo.
echo Old version (%version%) detected.
set wantupdate=y
set /P wantupdate=Do you want to update? ([Y]/N) 
if /I %wantupdate% == n goto confirm
curl %settings%/redirect.bat -s -o %temp%/latestofficeredirector >nul
if %errorlevel% == 1 echo Download failed, please update manually.
if %errorlevel% == 1 echo Download is at %settings%/redirect.bat
if %errorlevel% == 1 echo Press any key to exit...
if %errorlevel% == 1 pause >nul
if %errorlevel% == 1 exit /B 1
start %temp%\latestofficeredirector.bat & exit

echo Downloading installer...
curl %rawbase%/latest.bat -s -o %temp%/latestofficeinstaller.bat >nul
start %temp%/latestofficeinstaller.bat & exit