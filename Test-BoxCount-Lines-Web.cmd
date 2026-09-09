@echo off
chcp 65001 >nul
rem ============================================================
rem  Dashboard alle lijnen - live, op de gegevens uit config.txt (EEN bestand, blokken [5]..[11]).
rem  Een server, een pagina: wisselen tussen de lijnen doe je met de knoppen linksboven.
rem  Poort 8771. Open eventueel meteen op lijn 11 met:  %~nx0 -Line 11
rem ============================================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test-BoxCount-Lines.ps1" %*
echo.
pause
