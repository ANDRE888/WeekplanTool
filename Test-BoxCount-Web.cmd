@echo off
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test-BoxCount.ps1" -Web %*
echo.
pause
