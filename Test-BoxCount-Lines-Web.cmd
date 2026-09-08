@echo off
chcp 65001 >nul
rem ============================================================
rem  Dashboard LIJN 9 + LIJN 11 - live, op de gegevens uit config.txt / config-L11.txt.
rem  Een server, een pagina: wisselen tussen de lijnen doe je met de knoppen linksboven.
rem  Poort 8771. Open eventueel meteen op lijn 11 met:  %~nx0 -Line 11
rem ============================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test-BoxCount-Lines.ps1" %*
echo.
pause
