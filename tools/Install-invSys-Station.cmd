@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-invSys-Station.ps1" %*
set "invsys_exit=%ERRORLEVEL%"
if not "%invsys_exit%"=="0" pause
exit /b %invsys_exit%
