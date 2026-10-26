@echo off
setlocal
start "" powershell.exe -NoLogo -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0amtrak_fmd_lap_v1.16.2.ps1"
exit /b
