@echo off
chcp 65001 >nul
rem ============================================================
rem  TEST-launcher LIJN 9 + LIJN 11 - draait op de data in de map 'test'.
rem
rem  Pakt automatisch het NIEUWSTE Data_boxprintingbin3v7_*.xlsb uit die map en zet "Nu"
rem  op het tijdstempel in de bestandsnaam (= vlak na de laatste doos, ruim voor het
rem  ploegeinde), zodat prognose en grafieken echt iets te tonen hebben. Nieuwe snapshot
rem  in 'test' zetten volstaat - hier hoeft niets aangepast te worden.
rem
rem  Het weekplan kiest het script zelf uit dezelfde map (-PlanFolder).
rem  Wisselen tussen lijn 9 en 11 doe je op de pagina zelf. Poort 8790, zodat dit naast
rem  het live dashboard (8771) kan draaien.
rem ============================================================
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $d=Join-Path '%~dp0' 'test'; $b=Get-ChildItem -LiteralPath $d -Filter 'Data_boxprintingbin3v7*.xlsb' -File | Sort-Object Name -Descending | Select-Object -First 1; if(-not $b){ Write-Host ('GEEN Data_boxprintingbin3v7*.xlsb gevonden in ' + $d) -ForegroundColor Red; exit 1 }; $p=Get-ChildItem -LiteralPath $d -Filter 'daily shift NDwk*.xls*' -File | Sort-Object Name -Descending | Select-Object -First 1; $now = if($b.Name -match '_(\d{4}-\d{2}-\d{2})_(\d{2})-(\d{2})-(\d{2})'){ '{0} {1}:{2}' -f $Matches[1],$Matches[2],$Matches[3] } else { $b.LastWriteTime.ToString('yyyy-MM-dd HH:mm') }; $a=@{ Port=8790; Now=$now; BoxPrintingFile=$b.FullName; PlanFolder=$d }; Write-Host ('Snapshot : ' + $b.Name) -ForegroundColor Cyan; Write-Host ('Weekplan : ' + $(if($p){$p.Name}else{'(geen in test - target uit config)'})) -ForegroundColor Cyan; Write-Host ('Nu       : ' + $now) -ForegroundColor Cyan; Write-Host ''; & (Join-Path '%~dp0' 'Test-BoxCount-Lines.ps1') @a" %*
echo.
pause
