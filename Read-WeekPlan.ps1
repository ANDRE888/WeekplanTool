#requires -version 5
<#
    Read-WeekPlan.ps1
    Toont voor de lijn (werkblad L9RCDB in Data_boxprintingbin3v7.xlsb) de productie van de
    HUIDIGE PLOEG tegenover het ploegplan, plus de weekstand.

    Bestand tegen PRODUCTWISSEL midden in de ploeg: ALLE SKU's die deze ploeg gedraaid hebben
    worden getoond (elk met gemaakt / plan / nog te doen). De SKU die NU draait (laatste uur)
    krijgt << nu en de weekstand.

    SKU('s) worden automatisch bepaald uit L9RCDB; is er nog niets gemaakt, dan wordt de SKU
    volgens het weekplan (dag + ploeg) genomen. Met -Sku kan op één code worden gefocust.

    Het weekplan ("daily shift dpp") bevat 3 kolommen per dag = ploeg 1/2/3.
    Productiedag start om 05:00 (dag D = 05:00 D -> 05:00 D+1).
    Ploegen: 1 = 05:00-13:00, 2 = 13:00-21:00, 3 = 21:00-05:00. Koppeling op dag van de maand.

    Paden naar de databestanden staan in config.txt (naast dit script); bestaat het niet, dan
    wordt het aangemaakt. Voorrang: parameter > config.txt > standaard.

    MONITOR (standaard): blijft draaien, herlaadt zodra de opslagdatum van het box-printing
    bestand verandert (elke -IntervalSeconds). -Once = één keer. Ctrl+C = stop.

    Parameters: -Once -IntervalSeconds N -Sku CODE -RcdbSheet L6RCDB -Week NN
                -Now "2026-04-28 14:30" -PlanFile "..." -BoxPrintingFile "..."
    Vereist Microsoft Excel (via COM).
#>
param(
    [string]  $Sku,
    [int]     $Week,
    [string]  $PlanFile,
    [string]  $BoxPrintingFile,
    [string]  $SheetName,
    [string]  $RcdbSheet,
    [datetime]$Now,
    [switch]  $Once,
    [int]     $IntervalSeconds = 60
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$nl = [System.Globalization.CultureInfo]::GetCultureInfo('nl-BE')

# ============================ CONFIG ============================
# Paden naar de databestanden staan in config.txt (naast dit script).
# Bestaat config.txt niet, dan wordt het automatisch aangemaakt.
# Voorrang: parameter  >  config.txt  >  ingebouwde standaard.
$here = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if ([string]::IsNullOrWhiteSpace($SheetName)) { $SheetName = "daily shift dpp" }
$PlanPattern = "daily shift NDwk*.xls*"
$ConfigFile  = Join-Path $here "config.txt"

$script:HasNow      = $PSBoundParameters.ContainsKey('Now')
$script:HasWeek     = $PSBoundParameters.ContainsKey('Week')
$script:HasPlanFile = -not [string]::IsNullOrWhiteSpace($PlanFile)
$script:HasSku      = -not [string]::IsNullOrWhiteSpace($Sku)
$script:HasBox      = -not [string]::IsNullOrWhiteSpace($BoxPrintingFile)
$script:HasRcdb     = -not [string]::IsNullOrWhiteSpace($RcdbSheet)
# ===============================================================


function Convert-Serial([object]$v) {
    if ($v -is [double] -and $v -gt 40000 -and $v -lt 60000) { return [DateTime]::FromOADate($v) }
    return $null
}
function Get-PlanCandidates([string]$folder, [string]$pattern) {
    Get-ChildItem -LiteralPath $folder -Filter $pattern -File -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.Name -match 'NDwk0*(\d+)') { [pscustomobject]@{ File = $_.FullName; Week = [int]$Matches[1]; Modified = $_.LastWriteTime } }
    }
}
function Read-ConfigFile([string]$path) {
    $cfg = @{}
    if (Test-Path -LiteralPath $path) {
        foreach ($line in (Get-Content -LiteralPath $path)) {
            $t = $line.Trim()
            if ($t -eq '' -or $t.StartsWith('#')) { continue }
            $i = $t.IndexOf('=')
            if ($i -gt 0) {
                $k = $t.Substring(0, $i).Trim()
                $v = $t.Substring($i + 1).Trim().Trim('"').Trim("'")
                if ($k) { $cfg[$k] = $v }
            }
        }
    }
    return $cfg
}
function New-ConfigFile([string]$path, [string]$box, [string]$planFolder, [string]$rcdb) {
    $tmpl = @"
# ============================================================
#  Read-WeekPlan - configuratie / конфигурация
#  Пути к файлам с данными. Строки с # - комментарии.
#  Формат: ключ = значение
# ============================================================

# Путь к файлу сбора коробок (фактическая продукция):
BoxPrintingFile = $box

# Папка с файлами недельного плана (daily shift NDwkNN-YYYY.xls):
PlanFolder = $planFolder

# Лист линии в файле сбора коробок (L9RCDB, L6RCDB, ...):
RcdbSheet = $rcdb
"@
    [System.IO.File]::WriteAllText($path, $tmpl, (New-Object System.Text.UTF8Encoding($true)))
}
function Format-Sap([object]$v) {
    if ($null -eq $v) { return "" }
    $s = ([string]$v).Trim()
    if ($s.EndsWith('.0')) { $s = $s.Substring(0, $s.Length - 2) }
    return $s
}
function Is-Sku([string]$s) { return ($s.Length -eq 9 -and $s.StartsWith('3400')) }
function NF($x) { return ([double]$x).ToString('n0', $script:nl) }
function PF($x) { return ([double]$x).ToString('n1', $script:nl) }
function HourOrder([int]$h) { if ($h -ge 5) { return $h - 5 } else { return $h + 19 } }   # 5..23,0..4

function Resolve-PlanFile {
    if ($script:HasPlanFile) {
        if (-not (Test-Path -LiteralPath $PlanFile)) { throw "Opgegeven planbestand niet gevonden: $PlanFile" }
        return $PlanFile
    }
    $cands = @(Get-PlanCandidates $script:PlanFolder $PlanPattern)
    if ($cands.Count -eq 0) { throw "Geen planbestand gevonden ($PlanPattern in $script:PlanFolder)." }
    if ($script:HasWeek) {
        $pick = $cands | Where-Object { $_.Week -eq $Week } | Sort-Object Modified -Descending | Select-Object -First 1
        if (-not $pick) { throw "Planbestand voor week $Week niet gevonden. Beschikbaar: $(( $cands | Sort-Object Week | Select-Object -Expand Week ) -join ', ')" }
        return $pick.File
    }
    return ($cands | Sort-Object Week, Modified -Descending | Select-Object -First 1).File
}

function Get-Production($excel, [string]$boxFile, [string]$rcdbSheet) {
    # $prod[dag][uur][sku] = som aantal
    $wb2 = $excel.Workbooks.Open($boxFile, 0, $true)
    try {
        $ws2 = $null
        foreach ($s in $wb2.Worksheets) { if ($s.Name -eq $rcdbSheet) { $ws2 = $s; break } }
        if ($null -eq $ws2) { foreach ($s in $wb2.Worksheets) { if ($s.Name -like "*$rcdbSheet*") { $ws2 = $s; break } } }
        if ($null -eq $ws2) { throw "Werkblad '$rcdbSheet' niet gevonden in $boxFile" }
        $rng = $ws2.Range("A1:FN320").Value2
        $cmax = $rng.GetUpperBound(1); $rmax = $rng.GetUpperBound(0)
        $dayOfCol = @{}
        for ($c = 1; $c -le $cmax; $c++) {
            if ((($c - 1) % 5) -eq 0) { $dv = $rng.GetValue(2, $c); if ($dv -is [double]) { $dayOfCol[$c] = [int]$dv } }
        }
        $prod = @{}
        foreach ($c in $dayOfCol.Keys) {
            $day = $dayOfCol[$c]
            for ($r = 4; $r -le $rmax; $r++) {
                $sap = Format-Sap ($rng.GetValue($r, $c + 1))
                if (Is-Sku $sap) {
                    $a = $rng.GetValue($r, $c + 4); $u = $rng.GetValue($r, $c)
                    if ($a -is [double] -and $u -is [double]) {
                        $hr = [int]$u
                        if ($hr -ge 0 -and $hr -le 23) {
                            if (-not $prod.ContainsKey($day)) { $prod[$day] = @{} }
                            if (-not $prod[$day].ContainsKey($hr)) { $prod[$day][$hr] = @{} }
                            if ($prod[$day][$hr].ContainsKey($sap)) { $prod[$day][$hr][$sap] += $a } else { $prod[$day][$hr][$sap] = [double]$a }
                        }
                    }
                }
            }
        }
        return $prod
    }
    finally { if ($wb2) { $wb2.Close($false); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb2) } }
}
function Detect-CurrentSku($prod, [int]$day, [int]$maxOrder) {
    if (-not $prod -or -not $prod.ContainsKey($day)) { return $null }
    $bestOrder = -1; $bestHour = $null
    foreach ($h in $prod[$day].Keys) { $o = HourOrder $h; if ($o -le $maxOrder -and $o -gt $bestOrder) { $bestOrder = $o; $bestHour = $h } }
    if ($null -eq $bestHour) { return $null }
    $bestSku = $null; $bestVal = -1
    foreach ($s in $prod[$day][$bestHour].Keys) { if ($prod[$day][$bestHour][$s] -gt $bestVal) { $bestVal = $prod[$day][$bestHour][$s]; $bestSku = $s } }
    return $bestSku
}
function Get-ShiftSkus($prod, [int]$day, [int[]]$hours, [int]$maxOrder) {
    # alle SKU's met productie in deze ploeg (t/m nu): sku -> som aantal
    $m = @{}
    if (-not $prod -or -not $prod.ContainsKey($day)) { return $m }
    foreach ($h in $hours) {
        if ((HourOrder $h) -le $maxOrder -and $prod[$day].ContainsKey($h)) {
            foreach ($s in $prod[$day][$h].Keys) {
                if ($m.ContainsKey($s)) { $m[$s] += $prod[$day][$h][$s] } else { $m[$s] = [double]$prod[$day][$h][$s] }
            }
        }
    }
    return $m
}
function Shift-Made($prod, [int]$day, [int[]]$hours, [string]$sku, [int]$maxOrder = 24) {
    if (-not $prod -or -not $prod.ContainsKey($day)) { return 0 }
    $t = 0; foreach ($h in $hours) { if ((HourOrder $h) -le $maxOrder -and $prod[$day].ContainsKey($h) -and $prod[$day][$h].ContainsKey($sku)) { $t += $prod[$day][$h][$sku] } }; return $t
}
function Day-Made($prod, [int]$day, [string]$sku, [int]$maxOrder = 24) {
    if (-not $prod -or -not $prod.ContainsKey($day)) { return 0 }
    $t = 0; foreach ($h in $prod[$day].Keys) { if ((HourOrder $h) -le $maxOrder -and $prod[$day][$h].ContainsKey($sku)) { $t += $prod[$day][$h][$sku] } }; return $t
}
function Get-PlanColumns($vals, [int]$colMax) {
    $cols = @(); $prevDate = $null; $shift = 0
    for ($c = 4; $c -le $colMax; $c++) {
        $d = Convert-Serial ($vals.GetValue(2, $c))
        if ($null -ne $d) {
            $key = $d.ToString('yyyy-MM-dd')
            if ($key -ne $prevDate) { $shift = 0; $prevDate = $key } else { $shift++ }
            $cols += [pscustomobject]@{ Col = $c; Date = $key; Shift = $shift }
        }
    }
    return $cols
}
function Get-PlanRowForSku($vals, [int]$rowMax, [string]$sku) {
    $n = [double]$sku
    for ($r = 2; $r -le $rowMax; $r++) {
        $b = $vals.GetValue($r, 2); $a = [string]($vals.GetValue($r, 1))
        if (($b -is [double] -and [math]::Abs($b - $n) -lt 0.5) -or ($a -like "$sku*")) { return $r }
    }
    return 0
}
function Get-SkuFromPlan($vals, [int]$rowMax, $planCols, [string]$dateKey, [int]$shiftNo) {
    $col = $planCols | Where-Object { $_.Date -eq $dateKey -and $_.Shift -eq ($shiftNo - 1) } | Select-Object -First 1
    if (-not $col) { return $null }
    $bestSku = $null; $bestVal = 0
    for ($r = 2; $r -le $rowMax; $r++) {
        $v = $vals.GetValue($r, $col.Col)
        if ($v -is [double] -and $v -gt $bestVal) {
            $sap = Format-Sap ($vals.GetValue($r, 2))
            if (Is-Sku $sap) { $bestVal = $v; $bestSku = $sap }
        }
    }
    return $bestSku
}
function Get-PlanShift($vals, [int]$rowMax, $planCols, [string]$sku, [string]$dateKey, [int]$shiftNo) {
    $row = Get-PlanRowForSku $vals $rowMax $sku
    $desc = if ($row -gt 0) { [string]($vals.GetValue($row, 3)) } else { "" }
    $plan = 0.0
    if ($row -gt 0) {
        $col = $planCols | Where-Object { $_.Date -eq $dateKey -and $_.Shift -eq ($shiftNo - 1) } | Select-Object -First 1
        if ($col) { $v = $vals.GetValue($row, $col.Col); if ($v -is [double]) { $plan = [double]$v } }
    }
    return [pscustomobject]@{ Row = $row; Desc = $desc; Plan = $plan }
}

function Invoke-Report {
    $nowDt = if ($script:HasNow) { $Now } else { Get-Date }
    $hNow  = $nowDt.Hour
    if     ($hNow -ge 5  -and $hNow -lt 13) { $shiftNo = 1; $shiftHours = 5..12;                  $shiftLbl = "ploeg 1 (05:00-13:00)" }
    elseif ($hNow -ge 13 -and $hNow -lt 21) { $shiftNo = 2; $shiftHours = 13..20;                 $shiftLbl = "ploeg 2 (13:00-21:00)" }
    else                                    { $shiftNo = 3; $shiftHours = @(21,22,23,0,1,2,3,4);  $shiftLbl = "ploeg 3 (21:00-05:00)" }
    $prodDate = if ($hNow -lt 5) { $nowDt.Date.AddDays(-1) } else { $nowDt.Date }
    $prodKey  = $prodDate.ToString('yyyy-MM-dd')
    $prodDom  = $prodDate.Day
    $nowOrder = HourOrder $hNow

    $excel = $null; $wb = $null
    try {
        $planFilePath = Resolve-PlanFile
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AskToUpdateLinks = $false; $excel.ScreenUpdating = $false

        $wb = $excel.Workbooks.Open($planFilePath, 0, $true)
        $ws = $null
        foreach ($s in $wb.Worksheets) { if ($s.Name -eq $SheetName) { $ws = $s; break } }
        if ($null -eq $ws) { foreach ($s in $wb.Worksheets) { if ($s.Name -like "*dpp*") { $ws = $s; break } } }
        if ($null -eq $ws) { throw "Werkblad '$SheetName' niet gevonden in het planbestand." }

        $vals   = $ws.Range("A1:AD400").Value2
        $rowMax = $vals.GetUpperBound(0); $colMax = $vals.GetUpperBound(1)
        $weekNo = $null; $startSerial = $null; $stopSerial = $null
        for ($c = 1; $c -le $colMax; $c++) {
            $h = [string]($vals.GetValue(1, $c))
            if ($h -eq 'WEEK'  -and $c -lt $colMax) { $weekNo      = $vals.GetValue(1, $c + 1) }
            if ($h -eq 'start' -and $c -lt $colMax) { $startSerial = $vals.GetValue(1, $c + 1) }
            if ($h -eq 'stop'  -and $c -lt $colMax) { $stopSerial  = $vals.GetValue(1, $c + 1) }
        }
        $planCols = Get-PlanColumns $vals $colMax

        $prod = $null
        if (Test-Path -LiteralPath $BoxPrintingFile) {
            try { $prod = Get-Production $excel $BoxPrintingFile $RcdbSheet }
            catch { Write-Host "WAARSCHUWING: productie niet gelezen: $($_.Exception.Message)" -ForegroundColor Yellow }
        }
        else { Write-Host "WAARSCHUWING: box-printing bestand niet gevonden:`n  $BoxPrintingFile" -ForegroundColor Yellow }

        # koptekst
        Write-Host ""
        Write-Host ("=========== LIJN {0} - HUIDIGE PLOEG ===========" -f $RcdbSheet) -ForegroundColor Cyan
        Write-Host ("Tijd : {0}   {1}   productiedag {2}" -f $nowDt.ToString('dd/MM/yyyy HH:mm'), $shiftLbl, $prodDate.ToString('ddd dd/MM', $nl))
        $period = ""; $sd = Convert-Serial $startSerial; $ed = Convert-Serial $stopSerial
        if ($sd -and $ed) { $period = "{0} - {1}" -f $sd.ToString('dd/MM'), $ed.ToString('dd/MM') }
        if ($weekNo -ne $null) { Write-Host ("Week : {0} ({1})   Plan: {2}" -f [int]$weekNo, $period, (Split-Path $planFilePath -Leaf)) }
        Write-Host "================================================" -ForegroundColor Cyan

        # SKU('s) van de huidige ploeg bepalen
        if ($script:HasSku) {
            $focus = Format-Sap $Sku
            $shiftSkus = @{ $focus = (Shift-Made $prod $prodDom $shiftHours $focus $nowOrder) }
            $currentSku = $focus; $bron = "opgegeven (-Sku)"
        }
        else {
            $shiftSkus  = Get-ShiftSkus $prod $prodDom $shiftHours $nowOrder
            $currentSku = Detect-CurrentSku $prod $prodDom $nowOrder; $bron = "actief op lijn"
            if (-not $currentSku) { $currentSku = Get-SkuFromPlan $vals $rowMax $planCols $prodKey $shiftNo; $bron = "volgens plan (dag/ploeg)" }
            if ($currentSku -and -not $shiftSkus.ContainsKey($currentSku)) { $shiftSkus[$currentSku] = 0.0 }
        }

        if (-not $shiftSkus -or $shiftSkus.Count -eq 0) {
            Write-Host "Geen SKU: niets gemaakt deze ploeg en geen plan voor deze dag/ploeg." -ForegroundColor Yellow
            return
        }

        # --- ploeg-tabel (per SKU; robuust bij productwissel) ---
        Write-Host ("DEZE PLOEG ({0}) - SKU : gemaakt / plan / nog te doen:" -f $shiftLbl) -ForegroundColor White
        $tm = 0.0; $tp = 0.0; $tr = 0.0
        foreach ($s in ($shiftSkus.Keys | Sort-Object { $shiftSkus[$_] } -Descending)) {
            $made = [double]$shiftSkus[$s]
            $info = Get-PlanShift $vals $rowMax $planCols $s $prodKey $shiftNo
            $plan = [double]$info.Plan; $remain = $plan - $made
            $desc = if ($info.Desc) { $info.Desc } else { "(niet in plan)" }
            if ($desc.Length -gt 22) { $desc = $desc.Substring(0, 22) }
            $pctTxt = if ($plan -gt 0) { "{0} %" -f (PF (100.0 * $made / $plan)) } else { "-" }
            $mark = if ($s -eq $currentSku) { "  << nu" } else { "" }
            $col  = if ($s -eq $currentSku) { 'Green' } else { 'Gray' }
            Write-Host ("  {0,-9} {1,-22} {2,7} / {3,7} / {4,7}   {5}{6}" -f $s, $desc, (NF $made), (NF $plan), (NF $remain), $pctTxt, $mark) -ForegroundColor $col
            $tm += $made; $tp += $plan; $tr += $remain
        }
        if ($shiftSkus.Count -gt 1) {
            Write-Host ("  {0}" -f ('-' * 60)) -ForegroundColor DarkGray
            Write-Host ("  {0,-32} {1,7} / {2,7} / {3,7}" -f "Totaal", (NF $tm), (NF $tp), (NF $tr)) -ForegroundColor DarkGray
        }

        # --- weekstand voor de SKU die NU draait ---
        $crow  = Get-PlanRowForSku $vals $rowMax $currentSku
        $cdesc = if ($crow -gt 0) { [string]($vals.GetValue($crow, 3)) } else { "(niet in weekplan)" }
        Write-Host ""
        if ($crow -gt 0) {
            $planByDate = [ordered]@{}
            foreach ($pc in $planCols) {
                if (-not $planByDate.Contains($pc.Date)) { $planByDate[$pc.Date] = @(0.0, 0.0, 0.0) }
                $v = $vals.GetValue($crow, $pc.Col); if ($v -isnot [double]) { $v = 0 }
                ($planByDate[$pc.Date])[$pc.Shift] = [double]$v
            }
            $planWeek = 0.0; $madeWeek = 0.0
            foreach ($key in $planByDate.Keys) {
                $arr = $planByDate[$key]; $planWeek += ($arr[0] + $arr[1] + $arr[2])
                $dt = [DateTime]::ParseExact($key, 'yyyy-MM-dd', $null)
                if ($dt -gt $prodDate) { continue }
                $mo = if ($dt -eq $prodDate) { $nowOrder } else { 24 }
                $madeWeek += Day-Made $prod $dt.Day $currentSku $mo
            }
            $remain  = $planWeek - $madeWeek
            $pctMade = if ($planWeek -gt 0) { 100.0 * $madeWeek / $planWeek } else { 0 }
            $pctRem  = if ($planWeek -gt 0) { 100.0 * $remain   / $planWeek } else { 0 }
            Write-Host ("WEEK {0} - SKU die nu draait: {1} [{2}]" -f [int]$weekNo, $currentSku, $bron) -ForegroundColor White
            Write-Host ("  {0}" -f $cdesc) -ForegroundColor DarkGray
            Write-Host ("  Weekplan      : {0,8} dozen" -f (NF $planWeek))
            Write-Host ("  Reeds gemaakt : {0,8} dozen   ({1} %)" -f (NF $madeWeek), (PF $pctMade)) -ForegroundColor Green
            $wcol = if ($remain -gt 0) { 'Yellow' } else { 'Green' }
            Write-Host ("  Nog te maken  : {0,8} dozen   ({1} %)" -f (NF $remain), (PF $pctRem)) -ForegroundColor $wcol
        }
        else {
            Write-Host ("SKU die nu draait ({0}, {1}) staat niet in het weekplan." -f $currentSku, $bron) -ForegroundColor DarkGray
        }
    }
    catch { Write-Host "FOUT: $($_.Exception.Message)" -ForegroundColor Red }
    finally {
        if ($wb)    { $wb.Close($false) }
        if ($excel) { $excel.Quit() }
        if ($wb)    { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
        if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# ---- configuratie laden (config.txt naast het script) ----
if (-not (Test-Path -LiteralPath $ConfigFile)) {
    New-ConfigFile $ConfigFile (Join-Path $here "Data_boxprintingbin3v7.xlsb") $here "L9RCDB"
    Write-Host "config.txt is aangemaakt in de scriptmap - controleer/wijzig de paden indien nodig:" -ForegroundColor Yellow
    Write-Host ("  $ConfigFile") -ForegroundColor Yellow
    Write-Host ""
}
$cfg = Read-ConfigFile $ConfigFile
if (-not $script:HasBox) {
    if ($cfg.ContainsKey('BoxPrintingFile') -and $cfg['BoxPrintingFile']) { $BoxPrintingFile = $cfg['BoxPrintingFile'] }
    else { $BoxPrintingFile = Join-Path $here "Data_boxprintingbin3v7.xlsb" }
}
if (-not $script:HasRcdb) {
    if ($cfg.ContainsKey('RcdbSheet') -and $cfg['RcdbSheet']) { $RcdbSheet = $cfg['RcdbSheet'] }
    else { $RcdbSheet = "L9RCDB" }
}
$script:PlanFolder = if ($cfg.ContainsKey('PlanFolder') -and $cfg['PlanFolder']) { $cfg['PlanFolder'] } else { $here }

# ============================ UITVOEREN ============================
if ($Once) {
    Invoke-Report
    Write-Host ""
}
else {
    Write-Host ("MONITOR actief - controle elke {0}s op opslagdatum van:" -f $IntervalSeconds) -ForegroundColor DarkCyan
    Write-Host ("  {0}" -f $BoxPrintingFile) -ForegroundColor DarkCyan
    Write-Host  "  (Ctrl+C om te stoppen)" -ForegroundColor DarkCyan
    $lastStamp = $null
    while ($true) {
        $stamp = $null
        if (Test-Path -LiteralPath $BoxPrintingFile) { $stamp = (Get-Item -LiteralPath $BoxPrintingFile).LastWriteTime }
        if ($stamp -ne $lastStamp) {
            try { Clear-Host } catch {}
            $when = if ($stamp) { $stamp.ToString('dd/MM/yyyy HH:mm:ss') } else { 'bestand (nog) niet gevonden' }
            Write-Host ("[{0}] Opslagdatum gewijzigd -> {1} : gegevens herladen..." -f (Get-Date -Format 'HH:mm:ss'), $when) -ForegroundColor DarkCyan
            if ($stamp) { Invoke-Report }
            $lastStamp = $stamp
            Write-Host ""
            Write-Host ("Volgende controle elke {0}s. Laatst geladen: {1}. (Ctrl+C om te stoppen)" -f $IntervalSeconds, (Get-Date -Format 'HH:mm:ss')) -ForegroundColor DarkGray
        }
        Start-Sleep -Seconds $IntervalSeconds
    }
}
