@echo off
setlocal
title Install Slide Tools

rem Installs the Slide Tools add-in for PowerPoint for the current Windows user.
rem No admin rights needed: it saves the add-in's install file in your user folder
rem and adds one per-user PowerPoint setting that points to it.

set "ST_DIR=%USERPROFILE%\Slidetools"
set "ST_FILE=%USERPROFILE%\Slidetools\slidetools-manifest.xml"
set "ST_URL=https://chadisfred.github.io/Slidetools/slidetools-manifest.xml"
set "ST_KEY=HKCU\Software\Microsoft\Office\16.0\WEF\Developer"

echo.
echo  Slide Tools installer
echo  ---------------------
echo.

tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if not errorlevel 1 (
  echo  PowerPoint is open. Save your work, close PowerPoint, then press any key.
  pause >nul
)

if not exist "%ST_DIR%" mkdir "%ST_DIR%"
if exist "%ST_FILE%" del /f /q "%ST_FILE%"

echo  Downloading the install file...
where curl.exe >nul 2>nul
if not errorlevel 1 curl.exe -fsSL -o "%ST_FILE%" "%ST_URL%"
if not exist "%ST_FILE%" powershell -NoProfile -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -UseBasicParsing -Uri $env:ST_URL -OutFile $env:ST_FILE"

if not exist "%ST_FILE%" (
  echo.
  echo  The install file couldn't be downloaded. Check your internet connection,
  echo  or follow the manual steps in the install guide.
  echo.
  pause
  exit /b 1
)

findstr /C:"OfficeApp" "%ST_FILE%" >nul
if errorlevel 1 (
  del /f /q "%ST_FILE%"
  echo.
  echo  The download didn't contain the add-in's install file. A network filter
  echo  may have blocked it. Follow the manual steps in the install guide.
  echo.
  pause
  exit /b 1
)

reg add "%ST_KEY%" /v SlideTools /t REG_SZ /d "%ST_FILE%" /f >nul
if errorlevel 1 (
  echo.
  echo  PowerPoint's add-in setting couldn't be saved. Your organisation may block it.
  echo  Ask IT to deploy Slide Tools from the Microsoft 365 admin center instead.
  echo.
  pause
  exit /b 1
)

echo.
echo  Slide Tools is installed.
echo.
echo  Open PowerPoint and look for the Slide Tools Panel button on the Home tab,
echo  in the Slide Tools group. If it isn't there, go to Home, Add-ins,
echo  My Add-ins.
echo.
echo  Install file saved in: %ST_DIR%
echo  Leave that folder where it is. PowerPoint reads the file from there.
echo.
pause
