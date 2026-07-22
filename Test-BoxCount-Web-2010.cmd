@echo off
chcp 65001 >nul
rem TEST-launcher met vaste tijd (alsof het nu 21/07/2026 20:10 is) -> ploeg 13:00-21:00
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test-BoxCount.ps1" -Web -Port 8772 -Now "2026-07-21 20:10" %*
echo.
pause
