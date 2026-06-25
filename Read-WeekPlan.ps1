#requires -version 5
<#
    Read-WeekPlan.ps1
    Toont voor de lijn (werkblad L9RCDB in Data_boxprintingbin3v7.xlsb) de productie van de
    HUIDIGE PLOEG tegenover het ploegplan, plus de weekstand. Robuust bij PRODUCTWISSEL:
    alle SKU's van de ploeg worden getoond; de SKU die NU draait krijgt << nu + weekstand.

    WEERGAVE:
      - Console (standaard): monitor die elke minuut de opslagdatum van het box-printing
        bestand controleert en bij wijziging herlaadt. -Once = één keer.
      - Web (-Web): start een lokale webserver (alleen .NET TcpListener, geen extra software),
        opent de browser en toont dezelfde gegevens als HTML-dashboard (auto-ververst).

    SKU automatisch uit L9RCDB (laatste uur); niets gemaakt -> SKU volgens weekplan (dag+ploeg);
    -Sku = handmatige focus. Plan "daily shift dpp": 3 kolommen per dag = ploeg 1/2/3.
    Productiedag start om 05:00. Ploegen 05-13 / 13-21 / 21-05. Ploegnamen: nacht (21-05) = Z;
    namiddag (13-21) = Y bij even weeknr, anders X; ochtend (05-13) = omgekeerd. Koppeling op dag v/d maand.
    Paden in config.txt (naast script; wordt aangemaakt indien afwezig). Vereist Microsoft Excel (COM).

    Parameters: -Web -Port 8770 -NoBrowser -Once -IntervalSeconds N -Sku CODE -RcdbSheet L6RCDB
                -Week NN -Now "2026-04-28 14:30" -PlanFile "..." -BoxPrintingFile "..."
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
    [switch]  $Web,
    [int]     $Port = 8770,
    [switch]  $NoBrowser,
    [int]     $IntervalSeconds = 60
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$nl = [System.Globalization.CultureInfo]::GetCultureInfo('nl-BE')

# ============================ CONFIG ============================
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
function PF2($x) { return ([double]$x).ToString('n2', $script:nl) }
function HtmlEnc([string]$s) {
    if ($null -eq $s) { return "" }
    return ($s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
}
function HourOrder([int]$h) { if ($h -ge 5) { return $h - 5 } else { return $h + 19 } }   # 5..23,0..4
function HourToDateTime([datetime]$prodDate, [int]$h) {
    # productie-uur -> echte datum/tijd (uur 0..4 hoort bij de volgende kalenderdag)
    if ($h -ge 5) { return $prodDate.Date.AddHours($h) } else { return $prodDate.Date.AddDays(1).AddHours($h) }
}

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
    $wb2 = $excel.Workbooks.Open($boxFile, 0, $true)
    try {
        $ws2 = $null
        foreach ($s in $wb2.Worksheets) { if ($s.Name -eq $rcdbSheet) { $ws2 = $s; break } }
        if ($null -eq $ws2) { foreach ($s in $wb2.Worksheets) { if ($s.Name -like "*$rcdbSheet*") { $ws2 = $s; break } } }
        if ($null -eq $ws2) { throw "Werkblad '$rcdbSheet' niet gevonden in $boxFile" }
        $rng = $ws2.Range("A1:FN600").Value2   # ruime marge (blad nu ~293 rijen x 165 kol)
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
function Get-LineSkus($prod) {
    # alle SKU's die OOIT op deze lijn voorkomen (hele buffer) -> set om plan op lijn te filteren
    $set = @{}
    if ($prod) {
        foreach ($day in $prod.Keys) {
            foreach ($h in $prod[$day].Keys) {
                foreach ($s in $prod[$day][$h].Keys) { $set[$s] = $true }
            }
        }
    }
    return $set
}
function Get-SkuFirstHour($prod, [int]$day, [int[]]$hours, [string]$sku, [int]$maxOrder) {
    # eerste uur (in productie-volgorde) waarop de SKU deze ploeg draaide
    if (-not $prod -or -not $prod.ContainsKey($day)) { return $null }
    $bestOrder = 999; $bestHour = $null
    foreach ($h in $hours) {
        if ((HourOrder $h) -le $maxOrder -and $prod[$day].ContainsKey($h) -and $prod[$day][$h].ContainsKey($sku)) {
            $o = HourOrder $h
            if ($o -lt $bestOrder) { $bestOrder = $o; $bestHour = $h }
        }
    }
    return $bestHour
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
function Get-SkuFromPlan($vals, [int]$rowMax, $planCols, [string]$dateKey, [int]$shiftNo, $lineSkus) {
    # SKU met het hoogste plan voor deze dag+ploeg, MAAR alleen SKU's die op deze lijn draaien
    # (plan is fabriekbreed; $lineSkus = SKU's uit de L9RCDB-historie). Geen historie -> $null.
    $col = $planCols | Where-Object { $_.Date -eq $dateKey -and $_.Shift -eq ($shiftNo - 1) } | Select-Object -First 1
    if (-not $col) { return $null }
    $bestSku = $null; $bestVal = 0
    for ($r = 2; $r -le $rowMax; $r++) {
        $v = $vals.GetValue($r, $col.Col)
        if ($v -is [double] -and $v -gt $bestVal) {
            $sap = Format-Sap ($vals.GetValue($r, 2))
            if ((Is-Sku $sap) -and $lineSkus -and $lineSkus.ContainsKey($sap)) { $bestVal = $v; $bestSku = $sap }
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

# ---- berekent het rapport en geeft een data-object terug (geen uitvoer) ----
function Get-ReportData {
    $nowDt = if ($script:HasNow) { $Now } else { Get-Date }
    $hNow  = $nowDt.Hour
    if     ($hNow -ge 5  -and $hNow -lt 13) { $shiftNo = 1; $shiftHours = 5..12;                  $shiftTime = "05:00-13:00" }
    elseif ($hNow -ge 13 -and $hNow -lt 21) { $shiftNo = 2; $shiftHours = 13..20;                 $shiftTime = "13:00-21:00" }
    else                                    { $shiftNo = 3; $shiftHours = @(21,22,23,0,1,2,3,4);  $shiftTime = "21:00-05:00" }
    $prodDate = if ($hNow -lt 5) { $nowDt.Date.AddDays(-1) } else { $nowDt.Date }
    $prodKey  = $prodDate.ToString('yyyy-MM-dd')
    $prodDom  = $prodDate.Day
    $nowOrder = HourOrder $hNow

    $d = [ordered]@{
        Ok = $true; Error = $null; Warning = $null
        Line = $RcdbSheet; NowText = $nowDt.ToString('dd/MM/yyyy HH:mm')
        ShiftLabel = "ploeg ($shiftTime)"; ShiftCode = "?"
        ProdDateText = $prodDate.ToString('ddd dd/MM', $nl); WeekNo = $null; Period = ""; PlanFile = ""
        NoSku = $false; ShiftRows = @(); HasTotal = $false; TotalMade = 0; TotalPlan = 0; TotalRemain = 0
        CurrentSku = $null; Bron = ""; WeekInPlan = $false; WeekDesc = ""
        WeekPlan = 0; WeekMade = 0; WeekRemain = 0; PctMade = 0; PctRem = 0
        FileTimeText = "-"; HasForecast = $false; FcMade = 0; FcElapsedMin = 0
        FcPerMin = 0; FcPerHour = 0; FcProjection = 0; FcSkuStartText = ""; FcRemainMin = 0
    }

    $excel = $null; $wb = $null
    try {
        $planFilePath = Resolve-PlanFile
        $d.PlanFile = Split-Path $planFilePath -Leaf
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
        if ($weekNo -ne $null) { $d.WeekNo = [int]$weekNo }
        # ploegnaam X/Y/Z: nacht (21-05) = Z; namiddag (13-21) = Y bij even week / X bij oneven;
        # ochtend (05-13) = omgekeerd (X bij even, Y bij oneven). Even = weeknr deelbaar door 2.
        if ($shiftNo -eq 3) { $shiftCode = 'Z' }
        else {
            $even = ($d.WeekNo -ne $null) -and (($d.WeekNo % 2) -eq 0)
            if ($shiftNo -eq 2) { $shiftCode = if ($even) { 'Y' } else { 'X' } }
            else                { $shiftCode = if ($even) { 'X' } else { 'Y' } }
        }
        $d.ShiftCode  = $shiftCode
        $d.ShiftLabel = "ploeg $shiftCode ($shiftTime)"
        $sd = Convert-Serial $startSerial; $ed = Convert-Serial $stopSerial
        if ($sd -and $ed) { $d.Period = "{0} - {1}" -f $sd.ToString('dd/MM'), $ed.ToString('dd/MM') }
        $planCols = Get-PlanColumns $vals $colMax

        $prod = $null; $fileTime = $null
        if (Test-Path -LiteralPath $BoxPrintingFile) {
            $fileTime = (Get-Item -LiteralPath $BoxPrintingFile).LastWriteTime
            $d.FileTimeText = $fileTime.ToString('dd/MM/yyyy HH:mm')
            try { $prod = Get-Production $excel $BoxPrintingFile $RcdbSheet }
            catch { $d.Warning = "productie niet gelezen: $($_.Exception.Message)" }
        }
        else { $d.Warning = "box-printing bestand niet gevonden: $BoxPrintingFile" }

        if ($script:HasSku) {
            $focus = Format-Sap $Sku
            $shiftSkus = @{ $focus = (Shift-Made $prod $prodDom $shiftHours $focus $nowOrder) }
            $currentSku = $focus; $d.Bron = "opgegeven (-Sku)"
        }
        else {
            $shiftSkus  = Get-ShiftSkus $prod $prodDom $shiftHours $nowOrder
            $currentSku = Detect-CurrentSku $prod $prodDom $nowOrder; $d.Bron = "actief op lijn"
            if (-not $currentSku) {
                $lineSkus = Get-LineSkus $prod   # alleen SKU's die op deze lijn voorkomen
                $currentSku = Get-SkuFromPlan $vals $rowMax $planCols $prodKey $shiftNo $lineSkus
                $d.Bron = "volgens plan (dag/ploeg)"
            }
            if ($currentSku -and -not $shiftSkus.ContainsKey($currentSku)) { $shiftSkus[$currentSku] = 0.0 }
        }
        $d.CurrentSku = $currentSku

        if (-not $shiftSkus -or $shiftSkus.Count -eq 0) { $d.NoSku = $true; return [pscustomobject]$d }

        $tm = 0.0; $tp = 0.0; $tr = 0.0
        foreach ($s in ($shiftSkus.Keys | Sort-Object { $shiftSkus[$_] } -Descending)) {
            $made = [double]$shiftSkus[$s]
            $info = Get-PlanShift $vals $rowMax $planCols $s $prodKey $shiftNo
            $plan = [double]$info.Plan; $remain = $plan - $made
            $desc = if ($info.Desc) { $info.Desc } else { "(niet in plan)" }
            $pct  = if ($plan -gt 0) { 100.0 * $made / $plan } else { $null }
            $d.ShiftRows += [pscustomobject]@{ Sku = $s; Desc = $desc; Made = $made; Plan = $plan; Remain = $remain; Pct = $pct; Current = ($s -eq $currentSku) }
            $tm += $made; $tp += $plan; $tr += $remain
        }
        if ($d.ShiftRows.Count -gt 1) { $d.HasTotal = $true; $d.TotalMade = $tm; $d.TotalPlan = $tp; $d.TotalRemain = $tr }

        # --- prognose einde ploeg: tempo SINDS de START van de huidige SKU ---
        # 'nu' = opslagtijd van het box-printing bestand. Tempo = gemaakt / (nu - SKU-start).
        # Verwacht einde ploeg = gemaakt + tempo * resterende ploegtijd (ploeg = 8u, dus < 8u sinds start).
        $shiftStartHour = switch ($shiftNo) { 1 { 5 } 2 { 13 } 3 { 21 } }
        $shiftStart = $prodDate.Date.AddHours($shiftStartHour)
        $shiftEnd   = $shiftStart.AddHours(8)
        $curMade = if ($currentSku -and $shiftSkus.ContainsKey($currentSku)) { [double]$shiftSkus[$currentSku] } else { 0 }
        if ($fileTime -and $curMade -gt 0) {
            $firstHour = Get-SkuFirstHour $prod $prodDom $shiftHours $currentSku $nowOrder
            if ($null -ne $firstHour) {
                $skuStart = HourToDateTime $prodDate $firstHour
                $elapsed = ($fileTime - $skuStart).TotalMinutes
                if ($elapsed -ge 1) {
                    $remaining = ($shiftEnd - $fileTime).TotalMinutes
                    if ($remaining -lt 0) { $remaining = 0 }
                    $perMin = $curMade / $elapsed
                    $d.HasForecast    = $true
                    $d.FcMade         = $curMade
                    $d.FcElapsedMin   = [Math]::Round($elapsed)
                    $d.FcPerMin       = $perMin
                    $d.FcPerHour      = $perMin * 60
                    $d.FcProjection   = $curMade + $perMin * $remaining
                    $d.FcSkuStartText = $skuStart.ToString('HH:mm')
                    $d.FcRemainMin    = [Math]::Round($remaining)
                }
            }
        }

        $crow = Get-PlanRowForSku $vals $rowMax $currentSku
        $d.WeekDesc = if ($crow -gt 0) { [string]($vals.GetValue($crow, 3)) } else { "" }
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
            $d.WeekInPlan = $true
            $d.WeekPlan = $planWeek; $d.WeekMade = $madeWeek; $d.WeekRemain = $planWeek - $madeWeek
            $d.PctMade = if ($planWeek -gt 0) { 100.0 * $madeWeek / $planWeek } else { 0 }
            $d.PctRem  = if ($planWeek -gt 0) { 100.0 * $d.WeekRemain / $planWeek } else { 0 }
        }
    }
    catch { $d.Ok = $false; $d.Error = $_.Exception.Message }
    finally {
        if ($wb)    { $wb.Close($false) }
        if ($excel) { $excel.Quit() }
        if ($wb)    { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
        if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    return [pscustomobject]$d
}

# ---- console-weergave ----
function Render-Console($d) {
    if ($d.Warning) { Write-Host "WAARSCHUWING: $($d.Warning)" -ForegroundColor Yellow }
    Write-Host ""
    Write-Host ("=========== LIJN {0} - HUIDIGE PLOEG ===========" -f $d.Line) -ForegroundColor Cyan
    Write-Host ("Tijd : {0}   {1}   productiedag {2}" -f $d.NowText, $d.ShiftLabel, $d.ProdDateText)
    if ($d.WeekNo -ne $null) { Write-Host ("Week : {0} ({1})   Plan: {2}" -f $d.WeekNo, $d.Period, $d.PlanFile) }
    Write-Host "================================================" -ForegroundColor Cyan
    if ($d.Error) { Write-Host "FOUT: $($d.Error)" -ForegroundColor Red; return }
    if ($d.NoSku) { Write-Host "Geen SKU: niets gemaakt deze ploeg en geen plan voor deze dag/ploeg." -ForegroundColor Yellow; return }

    Write-Host ("DEZE PLOEG ({0}) - SKU : gemaakt / plan / nog te doen:" -f $d.ShiftLabel) -ForegroundColor White
    foreach ($r in $d.ShiftRows) {
        $desc = if ($r.Desc.Length -gt 22) { $r.Desc.Substring(0, 22) } else { $r.Desc }
        $pctTxt = if ($r.Pct -ne $null) { "{0} %" -f (PF $r.Pct) } else { "-" }
        $mark = if ($r.Current) { "  << nu" } else { "" }
        $col  = if ($r.Current) { 'Green' } else { 'Gray' }
        Write-Host ("  {0,-9} {1,-22} {2,7} / {3,7} / {4,7}   {5}{6}" -f $r.Sku, $desc, (NF $r.Made), (NF $r.Plan), (NF $r.Remain), $pctTxt, $mark) -ForegroundColor $col
    }
    if ($d.HasTotal) {
        Write-Host ("  {0}" -f ('-' * 60)) -ForegroundColor DarkGray
        Write-Host ("  {0,-32} {1,7} / {2,7} / {3,7}" -f "Totaal", (NF $d.TotalMade), (NF $d.TotalPlan), (NF $d.TotalRemain)) -ForegroundColor DarkGray
    }
    if ($d.HasForecast) {
        Write-Host ""
        Write-Host ("PROGNOSE einde ploeg (SKU {0}, o.b.v. opslagtijd {1}):" -f $d.CurrentSku, $d.FileTimeText) -ForegroundColor White
        Write-Host ("  Tempo                : {0} dozen/min" -f (PF2 $d.FcPerMin))
        Write-Host ("  Per uur              : {0} dozen/uur" -f (NF $d.FcPerHour))
        Write-Host ("  Verwacht einde ploeg : {0} dozen" -f (NF $d.FcProjection)) -ForegroundColor Cyan
        Write-Host ("  ({0} dozen sinds SKU-start {1} in {2} min; nog {3} min tot ploegeinde)" -f (NF $d.FcMade), $d.FcSkuStartText, (NF $d.FcElapsedMin), (NF $d.FcRemainMin)) -ForegroundColor DarkGray
    }
    Write-Host ""
    if ($d.WeekInPlan) {
        Write-Host ("WEEK {0} - SKU die nu draait: {1} [{2}]" -f $d.WeekNo, $d.CurrentSku, $d.Bron) -ForegroundColor White
        Write-Host ("  {0}" -f $d.WeekDesc) -ForegroundColor DarkGray
        Write-Host ("  Weekplan      : {0,8} dozen" -f (NF $d.WeekPlan))
        Write-Host ("  Reeds gemaakt : {0,8} dozen   ({1} %)" -f (NF $d.WeekMade), (PF $d.PctMade)) -ForegroundColor Green
        $wcol = if ($d.WeekRemain -gt 0) { 'Yellow' } else { 'Green' }
        Write-Host ("  Nog te maken  : {0,8} dozen   ({1} %)" -f (NF $d.WeekRemain), (PF $d.PctRem)) -ForegroundColor $wcol
    }
    else {
        Write-Host ("SKU die nu draait ({0}, {1}) staat niet in het weekplan." -f $d.CurrentSku, $d.Bron) -ForegroundColor DarkGray
    }
}

# ---- HTML-weergave (dashboard) ----
function Render-Html($d, $stamp) {
    $refresh = $IntervalSeconds
    $load = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')
    $stampTxt = if ($stamp) { $stamp.ToString('dd/MM/yyyy HH:mm:ss') } else { '-' }

    $meta = "$($d.NowText) &middot; $(HtmlEnc $d.ShiftLabel) &middot; productiedag $(HtmlEnc $d.ProdDateText)"
    if ($d.WeekNo -ne $null) { $meta += " &middot; Week $($d.WeekNo) ($(HtmlEnc $d.Period)) &middot; $(HtmlEnc $d.PlanFile)" }
    $warnHtml = if ($d.Warning) { "<div class='warn'>$(HtmlEnc $d.Warning)</div>" } else { "" }

    if ($d.Error) {
        $bodyHtml = "<div class='err'>FOUT: $(HtmlEnc $d.Error)</div>"
    }
    elseif ($d.NoSku) {
        $bodyHtml = "<div class='warn'>Geen SKU: niets gemaakt deze ploeg en geen plan voor deze dag/ploeg.</div>"
    }
    else {
        $rows = ""
        foreach ($r in $d.ShiftRows) {
            $cls = if ($r.Current) { "cur" } else { "" }
            $remCls = if ($r.Remain -gt 0) { "behind" } else { "done" }
            $pct = if ($r.Pct -ne $null) { (PF $r.Pct) + ' %' } else { '-' }
            $mark = if ($r.Current) { " <span class='nu'>nu</span>" } else { "" }
            $rows += "<tr class='$cls'><td class='sku'>$($r.Sku)$mark</td><td>$(HtmlEnc $r.Desc)</td><td class='num'>$(NF $r.Made)</td><td class='num'>$(NF $r.Plan)</td><td class='num $remCls'>$(NF $r.Remain)</td><td class='num'>$pct</td></tr>"
        }
        $totalRow = if ($d.HasTotal) { "<tr class='tot'><td colspan='2'>Totaal</td><td class='num'>$(NF $d.TotalMade)</td><td class='num'>$(NF $d.TotalPlan)</td><td class='num'>$(NF $d.TotalRemain)</td><td></td></tr>" } else { "" }
        $table = "<h2 class='sec'>Deze ploeg &mdash; gemaakt / plan / nog te doen</h2><table class='shift'><thead><tr><th>SKU</th><th>Product</th><th>Gemaakt</th><th>Plan</th><th>Nog te doen</th><th>%</th></tr></thead><tbody>$rows$totalRow</tbody></table>"

        if ($d.WeekInPlan) {
            $remCls = if ($d.WeekRemain -gt 0) { "behind" } else { "done" }
            $week = "<div class='week'><h2 class='sec'>Week $($d.WeekNo) &mdash; nu: $($d.CurrentSku) <span class='bron'>[$(HtmlEnc $d.Bron)]</span></h2><div class='wdesc'>$(HtmlEnc $d.WeekDesc)</div><div class='cards'><div class='card'><div class='lbl'>Weekplan</div><div class='val'>$(NF $d.WeekPlan)</div></div><div class='card'><div class='lbl'>Reeds gemaakt</div><div class='val done'>$(NF $d.WeekMade)</div><div class='sub'>$(PF $d.PctMade) %</div></div><div class='card'><div class='lbl'>Nog te maken</div><div class='val $remCls'>$(NF $d.WeekRemain)</div><div class='sub'>$(PF $d.PctRem) %</div></div></div></div>"
        }
        else {
            $week = "<div class='warn'>SKU $($d.CurrentSku) [$(HtmlEnc $d.Bron)] staat niet in het weekplan.</div>"
        }
        $fc = ""
        if ($d.HasForecast) {
            $fc = "<h2 class='sec'>Prognose einde ploeg &mdash; SKU $($d.CurrentSku) <span class='bron'>(o.b.v. opslagtijd $($d.FileTimeText))</span></h2>" +
                  "<div class='cards'>" +
                  "<div class='card'><div class='lbl'>Tempo</div><div class='val'>$(PF2 $d.FcPerMin)</div><div class='sub'>dozen / min</div></div>" +
                  "<div class='card'><div class='lbl'>Per uur</div><div class='val'>$(NF $d.FcPerHour)</div><div class='sub'>dozen / uur</div></div>" +
                  "<div class='card'><div class='lbl'>Verwacht einde ploeg</div><div class='val done'>$(NF $d.FcProjection)</div><div class='sub'>$(NF $d.FcMade) sinds $($d.FcSkuStartText) &middot; nog $(NF $d.FcRemainMin) min</div></div>" +
                  "</div>"
        }
        $bodyHtml = "$table$fc$week"
    }

    $css = "*{box-sizing:border-box}body{margin:0;background:#0f172a;color:#e2e8f0;font-family:Segoe UI,system-ui,Arial,sans-serif}" +
           ".wrap{max-width:920px;margin:0 auto;padding:24px}h1{font-size:22px;margin:0 0 4px}.meta{color:#94a3b8;font-size:13px;margin-bottom:16px}" +
           ".sec{font-size:14px;color:#cbd5e1;margin:18px 0 8px;text-transform:uppercase;letter-spacing:.03em}" +
           "table.shift{width:100%;border-collapse:collapse;background:#1e293b;border-radius:10px;overflow:hidden}" +
           ".shift th{text-align:left;font-size:11px;text-transform:uppercase;color:#94a3b8;padding:10px 12px;background:#172033}" +
           ".shift td{padding:11px 12px;border-top:1px solid #334155;font-size:15px}.shift td.num{text-align:right;font-variant-numeric:tabular-nums}" +
           ".shift tr.cur td{background:#14321f}.shift td.sku{font-weight:600}" +
           ".nu{background:#22c55e;color:#06210f;font-size:11px;padding:1px 7px;border-radius:8px;margin-left:4px}" +
           ".behind{color:#fbbf24;font-weight:600}.done{color:#34d399;font-weight:600}" +
           ".tot td{font-weight:700;border-top:2px solid #475569;color:#cbd5e1}" +
           ".week{margin-top:8px}.bron{color:#94a3b8;font-weight:400;font-size:13px;text-transform:none;letter-spacing:0}" +
           ".wdesc{color:#94a3b8;font-size:13px;margin-bottom:10px}.cards{display:flex;gap:12px;flex-wrap:wrap}" +
           ".card{flex:1;min-width:160px;background:#1e293b;border-radius:10px;padding:14px 16px}" +
           ".card .lbl{font-size:11px;color:#94a3b8;text-transform:uppercase}.card .val{font-size:30px;font-weight:700;margin-top:4px;font-variant-numeric:tabular-nums}" +
           ".card .sub{font-size:13px;color:#94a3b8;margin-top:2px}" +
           ".warn{background:#3a2e10;color:#fcd34d;padding:10px 12px;border-radius:8px;margin:10px 0}" +
           ".err{background:#3a1212;color:#fca5a5;padding:10px 12px;border-radius:8px;margin:10px 0}" +
           ".foot{color:#64748b;font-size:12px;margin-top:22px}"

    return "<!doctype html><html lang='nl'><head><meta charset='utf-8'><meta http-equiv='refresh' content='$refresh'>" +
           "<meta name='viewport' content='width=device-width, initial-scale=1'><title>WeekPlan $(HtmlEnc $d.Line)</title><style>$css</style></head>" +
           "<body><div class='wrap'><h1>Lijn $(HtmlEnc $d.Line) &mdash; huidige ploeg</h1><div class='meta'>$meta</div>$warnHtml$bodyHtml" +
           "<div class='foot'>Laatst geladen: $load &middot; bestand opgeslagen: $stampTxt &middot; ververst elke ${refresh}s</div></div></body></html>"
}

function Invoke-Report { Render-Console (Get-ReportData) }

function Start-WebServer([int]$port) {
    $listener = New-Object System.Net.Sockets.TcpListener ([System.Net.IPAddress]::Loopback, $port)
    try { $listener.Start() }
    catch { Write-Host "FOUT: kan poort $port niet openen: $($_.Exception.Message)" -ForegroundColor Red; return }
    $url = "http://127.0.0.1:$port/"
    Write-Host "Webserver actief op $url" -ForegroundColor Green
    Write-Host "(Ctrl+C om te stoppen)" -ForegroundColor DarkGray
    if (-not $NoBrowser) { try { Start-Process $url } catch {} }

    $cacheHtml = $null; $cacheStamp = $null; $cacheTime = [datetime]::MinValue
    try {
        while ($true) {
            $client = $listener.AcceptTcpClient()
            try {
                $client.ReceiveTimeout = 1500
                $stream = $client.GetStream()
                $buf = New-Object byte[] 4096
                try { [void]$stream.Read($buf, 0, $buf.Length) } catch {}   # request inlezen (best effort)
                $stamp = if (Test-Path -LiteralPath $BoxPrintingFile) { (Get-Item -LiteralPath $BoxPrintingFile).LastWriteTime } else { $null }
                $age = ([datetime]::Now - $cacheTime).TotalSeconds
                if ($null -eq $cacheHtml -or $stamp -ne $cacheStamp -or $age -gt 15) {
                    $cacheHtml = Render-Html (Get-ReportData) $stamp
                    $cacheStamp = $stamp; $cacheTime = [datetime]::Now
                }
                $body = [System.Text.Encoding]::UTF8.GetBytes($cacheHtml)
                $head = "HTTP/1.1 200 OK`r`nContent-Type: text/html; charset=utf-8`r`nContent-Length: $($body.Length)`r`nCache-Control: no-cache`r`nConnection: close`r`n`r`n"
                $hb = [System.Text.Encoding]::ASCII.GetBytes($head)
                $stream.Write($hb, 0, $hb.Length); $stream.Write($body, 0, $body.Length); $stream.Flush()
            }
            catch {}
            finally { $client.Close() }
        }
    }
    finally { $listener.Stop() }
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
if ($Web) {
    Start-WebServer $Port
}
elseif ($Once) {
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
