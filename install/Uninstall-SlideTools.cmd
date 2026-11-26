@echo off
setlocal
title Uninstall Slide Tools

rem Removes the Slide Tools add-in for the current Windows user.

set "ST_DIR=%USERPROFILE%\Slidetools"
set "ST_KEY=HKCU\Software\Microsoft\Office\16.0\WEF\Developer"

echo.
echo  Slide Tools uninstaller
echo  -----------------------
echo.

tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if not errorlevel 1 (
  echo  PowerPoint is open. Save your work, close PowerPoint, then press any key.
  pause >nul
)

reg delete "%ST_KEY%" /v SlideTools /f >nul 2>nul
if exist "%ST_DIR%" rmdir /s /q "%ST_DIR%"

tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if errorlevel 1 (
  if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef"
)

echo  Slide Tools has been removed. The Slide Tools Panel button disappears the next
echo  time PowerPoint starts.
echo.
pause
