#requires -version 5
<#
    Test-BoxCount.ps1  (TEST-kopie, afgeleid van Read-WeekPlan.ps1 / Read-WeekPlan-Web.cmd)

    Nieuwe logica: leest de GECACHETE doos-registraties op werkblad Boxruw9 in
    Data_boxprintingbin3v7*.xlsb. Data begint in kolom B vanaf rij 11; er wordt
    van rij 11 naar beneden geparseerd tot de EERSTE lege cel in kolom B.
    Elke rij = 1 doos. Uit de etiket-string (kolom B) wordt het producttype
    (tag 14, bv. 340056956) gehaald; kolom A = tijdstip (Excel-serieel).

    Toont per producttype het AANTAL dozen van de HUIDIGE PLOEG-interval
    (05:00-13:00 / 13:00-21:00 / 21:00-05:00), PLUS een staafgrafiek met dozen
    per minuut en een horizontale target-lijn (ShiftTarget dozen per ploeg,
    omgerekend naar dozen/min = ShiftTarget / 480). Ververst elke minuut.

    PROGNOSE EINDE PLOEG (zoals in Read-WeekPlan.ps1, maar op doosniveau):
      * HOOFDPRODUCT = het producttype van de LAATSTE RIJ (de laatst geregistreerde
        doos), niet het product met de meeste dozen.
      * Tempo = dozen van de huidige RUN van dat hoofdproduct (aaneengesloten
        laatste blok in de rijen) / minuten sinds de eerste doos van die run.
      * 'Nu' = min(klok, opslagtijd van het bestand): verder dan het bestand
        reikt de data niet. Resterend = ploegeinde - nu.
      * Verwacht einde ploeg = gemaakt + tempo * resterende minuten.
        Tweede schatting op tempo van de laatste -RecentMinutes minuten.
      * Extra: nodig tempo om target te halen + verwacht tijdstip target.
      * Cumulatieve grafiek: gemaakt tot nu + gestippelde prognose tot ploegeinde,
        met de target-lijn als schuine referentie.

    STRIKTE REGELS die hier zijn afgedwongen:
      * Alle bestanden worden ALLEEN-LEZEN geopend (ReadOnly:=True + assert).
      * Er blijven geen processen hangen: Excel wordt afgesloten en, als de
        ActiveFactory COM-invoegtoepassing het proces vasthoudt, hard gekilld
        via HWND->PID (alleen ons eigen exemplaar).
      * Alleen standaard onderdelen (Excel COM + .NET TcpListener + inline SVG).
        Niets downloaden.
      * Cache wordt NIET herberekend: EnableEvents=False (dooft Workbook_Open dat
        anders xlAutomatic forceert), AutomationSecurity=ForceDisable, Calculation=Manual.

    TARGET KOMT UIT HET WEEKPLAN (daily shift NDwk<week>-<jaar>.xls, blad 'daily shift dpp'):
      * Kolom B = SAP-code, kolom C = omschrijving, daarna 3 gedateerde kolommen per dag
        = de taken van ploeg 1 / 2 / 3.
      * Target van deze ploeg = SOM van het plan van de producten die deze ploeg ECHT
        gelopen hebben (plan is fabriekbreed, dus filteren op wat de lijn draait).
      * Nog niets gemaakt -> SKU met het hoogste plan voor deze dag+ploeg, maar alleen
        uit de producten die in dit Boxruw-blad voorkomen.
      * Dag niet in het plan (oud bestand) of geen plan gevonden -> terugval op
        -ShiftTarget / config, met een waarschuwing. -ShiftTarget expliciet meegeven wint altijd.

    Parameters: -Web -Port 8771 -NoBrowser -Once -IntervalSeconds N
                -Now "2026-07-21 14:00"  (testen met de snapshot-datum)
                -BoxPrintingFile "..."   -BoxSheet Boxruw9   -ShiftTarget 1234
                -RecentMinutes 30        (venster voor het 'recent tempo')
                -StopMinutes 2           (gat dat als stilstand telt)
                -PlanFile "..."          -NoPlan (target niet uit het weekplan halen)
#>
param(
    [string]  $BoxPrintingFile,
    [string]  $BoxSheet,
    [datetime]$Now,
    [switch]  $Once,
    [switch]  $Web,
    [int]     $Port = 8771,
    [switch]  $NoBrowser,
    [int]     $IntervalSeconds = 60,
    [int]     $ShiftTarget = 1234,
    [int]     $RecentMinutes = 30,
    [double]  $StopMinutes = 2,
    [string]  $PlanFile,
    [switch]  $NoPlan
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$nl = [System.Globalization.CultureInfo]::GetCultureInfo('nl-BE')

# ============================ CONFIG ============================
$here = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ConfigFile = Join-Path $here "config.txt"
$script:HasNow    = $PSBoundParameters.ContainsKey('Now')
$script:HasBox    = -not [string]::IsNullOrWhiteSpace($BoxPrintingFile)
$script:HasTarget = $PSBoundParameters.ContainsKey('ShiftTarget')
$script:HasRecentMin = $PSBoundParameters.ContainsKey('RecentMinutes')
$script:HasStopMin   = $PSBoundParameters.ContainsKey('StopMinutes')
$script:HasPlanFile  = -not [string]::IsNullOrWhiteSpace($PlanFile)
if ([string]::IsNullOrWhiteSpace($BoxSheet)) { $BoxSheet = "Boxruw9" }

# Excel-constanten
$xlManual        = -4135
$xlUp            = -4162
$msoForceDisable = 3
$xlMaxRows       = 1048576

# HWND -> PID (om precies ons Excel-exemplaar af te sluiten indien Quit het niet doet)
if (-not ([System.Management.Automation.PSTypeName]'Win32Hwnd').Type) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Win32Hwnd {
  [DllImport("user32.dll", SetLastError=true)]
  public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int pid);
}
"@
}
# ===============================================================

function Rel($o) { if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }
function NF($x)  { return ([double]$x).ToString('n0', $script:nl) }
function PF($x)  { return ([double]$x).ToString('n1', $script:nl) }
function PF2($x) { return ([double]$x).ToString('n2', $script:nl) }
# SVG-coordinaat ALTIJD met punt (InvariantCulture), anders breekt nl-BE komma de SVG
function SvgN($x) { return ([Math]::Round([double]$x, 2)).ToString([System.Globalization.CultureInfo]::InvariantCulture) }
function HtmlEnc([string]$s) {
    if ($null -eq $s) { return "" }
    return ($s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
}
# waarschuwingen stapelen in plaats van overschrijven
function Add-Warn($d, [string]$msg) {
    if ([string]::IsNullOrWhiteSpace($msg)) { return }
    if ([string]::IsNullOrWhiteSpace($d.Warning)) { $d.Warning = $msg } else { $d.Warning = "$($d.Warning) | $msg" }
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

# Producttype = waarde van tag 14 in de etiket-string (tussen |14| en de volgende |)
function Get-ProductType([string]$label) {
    if ([string]::IsNullOrEmpty($label)) { return $null }
    $m = [regex]::Match($label, '\|14\|([^|]*)')
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
    return $null
}
# Doosteller = laatste veld na de laatste | (controltekens verwijderd)
function Get-Counter([string]$label) {
    if ([string]::IsNullOrEmpty($label)) { return '' }
    $clean = $label.Replace([string][char]1, '').Replace([string][char]26, '')
    $parts = $clean -split '\|'
    if ($parts.Length -gt 0) { return ($parts[$parts.Length - 1]).Trim() }
    return ''
}

# Huidige ploeg-interval op basis van 'nu'. Ploegen: 05-13 / 13-21 / 21-05.
function Get-ShiftWindow([datetime]$now) {
    $h = $now.Hour; $today = $now.Date
    if     ($h -ge 5  -and $h -lt 13) { $s = $today.AddHours(5);  $e = $today.AddHours(13);           $lab = '05:00-13:00'; $c = '1' }
    elseif ($h -ge 13 -and $h -lt 21) { $s = $today.AddHours(13); $e = $today.AddHours(21);           $lab = '13:00-21:00'; $c = '2' }
    elseif ($h -ge 21)                { $s = $today.AddHours(21); $e = $today.AddDays(1).AddHours(5);  $lab = '21:00-05:00'; $c = '3' }
    else                              { $s = $today.AddDays(-1).AddHours(21); $e = $today.AddHours(5); $lab = '21:00-05:00'; $c = '3' }
    return [pscustomobject]@{ Start = $s; End = $e; Label = $lab; Code = $c }
}

# Kies het box-printing bestand: config/param indien aanwezig, anders nieuwste snapshot in de map.
function Resolve-BoxFile {
    if ($BoxPrintingFile -and (Test-Path -LiteralPath $BoxPrintingFile)) {
        return (Resolve-Path -LiteralPath $BoxPrintingFile).Path
    }
    $cand = Get-ChildItem -LiteralPath $script:BoxFolder -Filter 'Data_boxprinting*.xls*' -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($cand) { return $cand.FullName }
    throw "Geen box-printing bestand gevonden (config-pad bestaat niet en geen Data_boxprinting*.xls* in $($script:BoxFolder))."
}

# ---------- weekplan (daily shift dpp) : target per ploeg ----------
function Format-Sap([object]$v) {
    if ($null -eq $v) { return "" }
    $s = ([string]$v).Trim()
    if ($s.EndsWith('.0')) { $s = $s.Substring(0, $s.Length - 2) }
    return $s
}
function Is-Sku([string]$s) { return ($s.Length -eq 9 -and $s.StartsWith('3400')) }
function Convert-Serial([object]$v) {
    if ($v -is [double] -and $v -gt 40000 -and $v -lt 60000) { return [DateTime]::FromOADate($v) }
    return $null
}
# Kandidaat-planbestanden, nieuwste weeknummer eerst.
function Get-PlanCandidates([string]$folder) {
    if ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder)) { return @() }
    $list = Get-ChildItem -LiteralPath $folder -Filter 'daily shift NDwk*.xls*' -File -ErrorAction SilentlyContinue |
            ForEach-Object {
                $w = 0; if ($_.Name -match 'NDwk0*(\d+)') { $w = [int]$Matches[1] }
                [pscustomobject]@{ File = $_.FullName; Week = $w; Modified = $_.LastWriteTime }
            }
    return @($list | Sort-Object Week, Modified -Descending)
}

# Leest het plan voor productiedag + ploeg. $skus = producten die deze ploeg draaiden,
# $lineSkus = alle producten die in dit Boxruw-blad voorkomen (filter, plan is fabriekbreed).
function Read-PlanTargets($excel, [string]$planPath, [datetime]$prodDate, [int]$shiftNo, $skus, $lineSkus) {
    $res = [ordered]@{
        Ok = $false; File = (Split-Path $planPath -Leaf); WeekNo = $null; Period = ''
        Covered = $false; Total = 0.0; PerSku = @{}; Desc = @{}; FallbackSku = $null; Error = $null
    }
    $wb = $null; $ws = $null; $sheets = $null
    try {
        $wb = $excel.Workbooks.Open($planPath, 0, $true)          # UpdateLinks=0, ReadOnly=True
        if (-not $wb.ReadOnly) { throw "Planbestand niet in alleen-lezen geopend - afgebroken." }
        $sheets = $wb.Worksheets
        foreach ($s in $sheets) { if ($s.Name -eq 'daily shift dpp') { $ws = $s; break } }
        if ($null -eq $ws) { foreach ($s in $sheets) { if ($s.Name -like '*dpp*') { $ws = $s; break } } }
        if ($null -eq $ws) { throw "Werkblad 'daily shift dpp' niet gevonden in $($res.File)." }

        $vals   = $ws.Range("A1:AD400").Value2
        $rowMax = $vals.GetUpperBound(0); $colMax = $vals.GetUpperBound(1)

        # rij 1: WEEK / start / stop
        $startD = $null; $stopD = $null
        for ($c = 1; $c -le $colMax; $c++) {
            $h = [string]($vals.GetValue(1, $c))
            if ($h -eq 'WEEK'  -and $c -lt $colMax) { $w = $vals.GetValue(1, $c + 1); if ($w -is [double]) { $res.WeekNo = [int]$w } }
            if ($h -eq 'start' -and $c -lt $colMax) { $startD = Convert-Serial ($vals.GetValue(1, $c + 1)) }
            if ($h -eq 'stop'  -and $c -lt $colMax) { $stopD  = Convert-Serial ($vals.GetValue(1, $c + 1)) }
        }
        if ($startD -and $stopD) { $res.Period = '{0} - {1}' -f $startD.ToString('dd/MM'), $stopD.ToString('dd/MM') }

        # rij 2: gedateerde kolommen, 3 per dag = ploeg 1/2/3
        $key = $prodDate.ToString('yyyy-MM-dd'); $col = 0; $prevKey = $null; $shift = 0
        for ($c = 4; $c -le $colMax; $c++) {
            $dt = Convert-Serial ($vals.GetValue(2, $c))
            if ($null -eq $dt) { continue }
            $k = $dt.ToString('yyyy-MM-dd')
            if ($k -ne $prevKey) { $shift = 0; $prevKey = $k } else { $shift++ }
            if ($k -eq $key -and $shift -eq ($shiftNo - 1)) { $col = $c; break }
        }
        if ($col -eq 0) { $res.Ok = $true; return [pscustomobject]$res }   # dag/ploeg staat niet in dit bestand
        $res.Covered = $true

        $want = @{}; foreach ($s in $skus) { if ($s) { $want[$s] = $true } }
        $best = $null; $bestVal = 0.0
        for ($r = 3; $r -le $rowMax; $r++) {
            $sap = Format-Sap ($vals.GetValue($r, 2))
            if (-not (Is-Sku $sap)) { continue }
            if ($want.ContainsKey($sap) -and -not $res.Desc.ContainsKey($sap)) {
                $res.Desc[$sap] = ([string]($vals.GetValue($r, 3))).Trim()
            }
            $v = $vals.GetValue($r, $col)
            if ($v -isnot [double]) { continue }
            # dezelfde SAP-code kan meerdere planregels hebben -> optellen
            if ($want.ContainsKey($sap)) {
                if ($res.PerSku.ContainsKey($sap)) { $res.PerSku[$sap] += [double]$v } else { $res.PerSku[$sap] = [double]$v }
            }
            if ($lineSkus -and $lineSkus.ContainsKey($sap) -and $v -gt $bestVal) { $bestVal = [double]$v; $best = $sap }
        }
        $tot = 0.0; foreach ($k2 in $res.PerSku.Keys) { $tot += [double]$res.PerSku[$k2] }
        # nog niets gemaakt deze ploeg -> pak het geplande product van de lijn
        if ($want.Count -eq 0 -and $best -and $bestVal -gt 0) {
            $res.FallbackSku = $best; $res.PerSku[$best] = $bestVal; $tot = $bestVal
        }
        $res.Total = $tot
        $res.Ok = $true
    }
    catch { $res.Error = $_.Exception.Message }
    finally {
        Rel $ws; Rel $sheets
        if ($wb) { try { $wb.Close($false) } catch {}; Rel $wb }
    }
    return [pscustomobject]$res
}

# ---- leest Boxruw9 uit de cache en telt dozen per producttype + per minuut voor de huidige ploeg ----
function Get-BoxData {
    $nowDt = if ($script:HasNow) { $Now } else { Get-Date }
    $win = Get-ShiftWindow $nowDt
    $shiftMin = [int][math]::Round(($win.End - $win.Start).TotalMinutes)
    if ($shiftMin -le 0) { $shiftMin = 480 }
    $minutes = New-Object 'int[]' $shiftMin
    $targetPerMin = if ($ShiftTarget -gt 0) { [double]$ShiftTarget / $shiftMin } else { 0 }

    $d = [ordered]@{
        Ok = $true; Error = $null; Warning = $null
        NowText = $nowDt.ToString('dd/MM/yyyy HH:mm')
        ShiftLabel = "ploeg $($win.Code) ($($win.Label))"
        WindowText = ('{0} -> {1}' -f $win.Start.ToString('dd/MM HH:mm'), $win.End.ToString('dd/MM HH:mm'))
        Sheet = $BoxSheet; BoxFile = ''; FileTimeText = '-'
        TargetSource = 'parameter/config'; PlanFileName = ''; PlanWeek = $null
        PlanPeriod = ''; PlanCovered = $false; PlanSku = $null; ShiftNo = 0
        Rows = @(); Total = 0; ParsedRows = 0; LastText = ''; StartRowSkipped = $false
        ShiftStart = $win.Start; ShiftEnd = $win.End; ShiftMin = $shiftMin; Minutes = $minutes; MaxPerMin = 0
        Target = $ShiftTarget; TargetPerMin = $targetPerMin; Pct = 0
        # --- prognose einde ploeg ---
        HasForecast = $false; FcWeak = $false
        MainProduct = $null; MainCount = 0        # hoofdproduct = producttype van de LAATSTE rij
        RunCount = 0; RunStartText = ''           # huidige aaneengesloten run van dat product
        RefNowText = '-'; ElapsedMin = 0; RemainMin = 0; NowOffsetMin = 0
        PerMin = 0; PerHour = 0
        RecentWin = 0; RecentCount = 0; HasRecent = $false; RecentPerMin = 0; ProjRecent = 0
        ProjTotal = 0; ProjMain = 0; ProjPct = 0; ProjDiff = 0
        NeedPerMin = 0; EtaText = ''
        # --- stilstand + achterstand ---
        HasStops = $false; Stops = @(); StopCount = 0; StopMin = 0
        LongestMin = 0; LongestText = ''; NowStill = $false; StillMin = 0
        ElapsedShiftMin = 0; RunMin = 0; AvailPct = 0; NetPerMin = 0; LostBoxes = 0
        BehindNow = 0; BehindEnd = 0; StopLimit = $StopMinutes
    }

    try { $boxFile = Resolve-BoxFile } catch { $d.Ok = $false; $d.Error = $_.Exception.Message; return [pscustomobject]$d }
    $d.BoxFile = Split-Path $boxFile -Leaf
    if (Test-Path -LiteralPath $boxFile) { $d.FileTimeText = (Get-Item -LiteralPath $boxFile).LastWriteTime.ToString('dd/MM/yyyy HH:mm') }

    $excel = $null; $wbs = $null; $blank = $null; $wb = $null; $sheets = $null; $ws = $null; $xlPid = 0
    try {
        $excel = New-Object -ComObject Excel.Application
        try { $hwnd = [IntPtr]$excel.Hwnd; [void][Win32Hwnd]::GetWindowThreadProcessId($hwnd, [ref]$xlPid) } catch {}
        $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.ScreenUpdating = $false; $excel.AskToUpdateLinks = $false

        # manuele berekening ZETTEN via een lege eerste werkmap (modus komt van 1e werkmap)
        $wbs = $excel.Workbooks
        $blank = $wbs.Add()
        $excel.Calculation = $xlManual
        $excel.CalculateBeforeSave = $false
        $excel.EnableEvents = $false               # dooft Workbook_Open (dat anders xlAutomatic forceert)
        $excel.AutomationSecurity = $msoForceDisable

        $wb = $wbs.Open($boxFile, 0, $true)         # UpdateLinks=0, ReadOnly=True
        if (-not $wb.ReadOnly) { throw "Bestand niet in alleen-lezen geopend - afgebroken (strikte regel)." }

        $sheets = $wb.Worksheets
        foreach ($s in $sheets) { if ($s.Name -eq $BoxSheet) { $ws = $s; break } }
        if ($null -eq $ws) { throw "Werkblad '$BoxSheet' niet gevonden in $($d.BoxFile)." }

        # laatste gevulde rij in kolom B (om het leesbereik te begrenzen)
        $allCells = $ws.Cells
        $anchor   = $allCells.Item($xlMaxRows, 2)
        $lastCell = $anchor.End($xlUp)
        $lastRow  = [int]$lastCell.Row
        Rel $lastCell; Rel $anchor; Rel $allCells

        if ($lastRow -lt 11) { $d.Warning = "Geen data vanaf rij 11 in $BoxSheet."; return [pscustomobject]$d }

        # A=tijd, B=etiket, C=lijn ; alles in EEN marshaling-call (leest cache, geen herberekening)
        $rng  = $ws.Range("A11:C$lastRow").Value2
        $rmax = $rng.GetUpperBound(0)

        $cutoff = if ($script:HasNow) { $nowDt } else { [datetime]::MaxValue }

        # rijen van de ploeg IN VOLGORDE bijhouden: het hoofdproduct is de LAATSTE rij,
        # en voor het tempo is de huidige aaneengesloten run van dat product nodig.
        $sTs   = New-Object 'System.Collections.Generic.List[datetime]'
        $sProd = New-Object 'System.Collections.Generic.List[string]'

        $counts = @{}; $allProds = @{}; $parsed = 0; $lastTs = $null; $lastProd = $null; $lastCtr = $null
        for ($i = 1; $i -le $rmax; $i++) {
            $b = $rng.GetValue($i, 2)
            if ($null -eq $b -or ([string]$b).Trim() -eq '') { break }   # stop bij EERSTE lege B
            $parsed++
            $label = [string]$b
            $prod  = Get-ProductType $label
            if ([string]::IsNullOrWhiteSpace($prod)) { $prod = '(onbekend)' }

            $tsRaw = $rng.GetValue($i, 1); $ts = $null
            if ($tsRaw -is [double]) { try { $ts = [DateTime]::FromOADate([double]$tsRaw) } catch {} }

            # De EERSTE rij is de startwaarde van de DELTA-ophaling (stand op het begin van het
            # ophaalvenster, exact op het hele uur) - die doos is EERDER geprint, dus niet meetellen.
            if ($i -eq 1 -and $null -ne $ts -and $ts.Minute -eq 0 -and $ts.Second -eq 0 -and $ts.Millisecond -eq 0) {
                $d.StartRowSkipped = $true
                continue
            }

            # alle producten van het blad (= wat op deze lijn draait) - filter voor het fabriekbrede plan
            if (-not $allProds.ContainsKey($prod)) { $allProds[$prod] = $true }

            # tellen alleen als het tijdstip in de HUIDIGE ploeg-interval valt
            # (met -Now ook niet verder tellen dan dat gesimuleerde moment: tijdmachine)
            if ($null -ne $ts -and $ts -ge $win.Start -and $ts -lt $win.End -and $ts -le $cutoff) {
                if ($counts.ContainsKey($prod)) { $counts[$prod]++ } else { $counts[$prod] = 1 }
                $off = [int][math]::Floor(($ts - $win.Start).TotalMinutes)
                if ($off -ge 0 -and $off -lt $shiftMin) { $minutes[$off]++ }
                $sTs.Add($ts); $sProd.Add($prod)
                if ($null -eq $lastTs -or $ts -gt $lastTs) { $lastTs = $ts; $lastProd = $prod; $lastCtr = Get-Counter $label }
            }
        }
        $d.ParsedRows = $parsed

        # hoofdproduct = producttype van de LAATSTE rij van de ploeg (wat er NU loopt)
        $mainProd = $null
        if ($sProd.Count -gt 0) { $mainProd = $sProd[$sProd.Count - 1] }
        $d.MainProduct = $mainProd

        $tot = 0; foreach ($p in $counts.Keys) { $tot += [int]$counts[$p] }
        $d.Total = $tot
        $mx = 0; foreach ($v in $minutes) { if ($v -gt $mx) { $mx = $v } }
        $d.MaxPerMin = $mx
        if ($lastTs) { $d.LastText = ('{0} - {1} (doos {2})' -f $lastTs.ToString('HH:mm'), $lastProd, $lastCtr) }

        # ---------------- TARGET UIT HET WEEKPLAN ----------------
        $shiftNo = [int]$win.Code; $prodDate = $win.Start.Date
        $d.ShiftNo = $shiftNo
        $effTarget = [double]$ShiftTarget
        $planPerSku = @{}; $planDesc = @{}
        if (-not $NoPlan -and -not $script:HasTarget) {
            $cands = @()
            if ($script:HasPlanFile) {
                if (Test-Path -LiteralPath $PlanFile) { $cands = @([pscustomobject]@{ File = (Resolve-Path -LiteralPath $PlanFile).Path }) }
                else { Add-Warn $d "opgegeven planbestand niet gevonden: $PlanFile" }
            }
            else { $cands = @(Get-PlanCandidates $script:PlanFolder | Select-Object -First 3) }

            $pt = $null
            foreach ($c in $cands) {
                $try = Read-PlanTargets $excel $c.File $prodDate $shiftNo (@($counts.Keys)) $allProds
                if ($null -eq $pt -or $try.Covered) { $pt = $try }
                if ($try.Covered) { break }
            }
            if ($null -eq $pt) {
                if ($cands.Count -eq 0) { Add-Warn $d "geen planbestand (daily shift NDwk*.xls) in $($script:PlanFolder) - target uit config." }
            }
            elseif ($pt.Error)   { Add-Warn $d "weekplan niet gelezen: $($pt.Error)" }
            else {
                $d.PlanFileName = $pt.File; $d.PlanWeek = $pt.WeekNo; $d.PlanPeriod = $pt.Period; $d.PlanCovered = $pt.Covered
                if (-not $pt.Covered) {
                    Add-Warn $d ("productiedag {0} (ploeg {1}) staat niet in {2} (week {3}, {4}) - target uit config." -f $prodDate.ToString('dd/MM'), $shiftNo, $pt.File, $pt.WeekNo, $pt.Period)
                }
                elseif ($pt.Total -le 0) {
                    Add-Warn $d ("geen plan voor deze dag/ploeg in {0} - target uit config." -f $pt.File)
                }
                else {
                    $effTarget  = [double]$pt.Total
                    $planPerSku = $pt.PerSku; $planDesc = $pt.Desc; $d.PlanSku = $pt.FallbackSku
                    $d.TargetSource = if ($pt.FallbackSku) {
                        "plan $($pt.File) - nog niets gemaakt, verwacht $($pt.FallbackSku)"
                    } else {
                        "plan $($pt.File) (week $($pt.WeekNo)), $($prodDate.ToString('ddd dd/MM', $script:nl)) ploeg $shiftNo"
                    }
                }
            }
        }
        elseif ($script:HasTarget) { $d.TargetSource = 'parameter -ShiftTarget' }
        elseif ($NoPlan)           { $d.TargetSource = 'config (-NoPlan)' }

        $d.Target = $effTarget
        $targetPerMin = if ($effTarget -gt 0) { $effTarget / $shiftMin } else { 0 }
        $d.TargetPerMin = $targetPerMin
        $d.Pct = if ($effTarget -gt 0) { 100.0 * $tot / $effTarget } else { 0 }

        $rowsOut = @()
        foreach ($p in ($counts.Keys | Sort-Object { $counts[$_] } -Descending)) {
            $pl = 0.0; if ($planPerSku.ContainsKey($p)) { $pl = [double]$planPerSku[$p] }
            $ds = '';  if ($planDesc.ContainsKey($p))   { $ds = [string]$planDesc[$p] }
            $rowsOut += [pscustomobject]@{ Product = $p; Count = [int]$counts[$p]; IsMain = ($p -eq $mainProd); Plan = $pl; Desc = $ds }
        }
        $d.Rows = $rowsOut

        # ---------------- PROGNOSE EINDE PLOEG ----------------
        if ($mainProd -and $sTs.Count -gt 0) {
            # 'nu' voor de data = niet later dan de opslagtijd van het bestand en niet na ploegeinde
            $refNow = $nowDt
            if (Test-Path -LiteralPath $boxFile) {
                $ft = (Get-Item -LiteralPath $boxFile).LastWriteTime
                if ($ft -lt $refNow) { $refNow = $ft }
            }
            if ($refNow -gt $win.End)   { $refNow = $win.End }
            if ($refNow -lt $lastTs)    { $refNow = $lastTs }   # data kan niet ouder zijn dan de laatste doos

            # ---- STILSTAND: gaten tussen opeenvolgende dozen groter dan -StopMinutes ----
            # ook het gat ploegstart -> eerste doos (opstart) en laatste doos -> nu (staat de lijn NU stil).
            $stopList = @(); $prevTs = $win.Start
            for ($k = 0; $k -lt $sTs.Count; $k++) {
                $gap = ($sTs[$k] - $prevTs).TotalMinutes
                if ($gap -gt $StopMinutes) {
                    $kind = if ($k -eq 0) { 'opstart' } else { 'stop' }
                    $stopList += [pscustomobject]@{ From = $prevTs; To = $sTs[$k]; Min = $gap; Kind = $kind }
                }
                $prevTs = $sTs[$k]
            }
            $tail = ($refNow - $prevTs).TotalMinutes
            if ($tail -gt $StopMinutes) {
                $stopList += [pscustomobject]@{ From = $prevTs; To = $refNow; Min = $tail; Kind = 'nu' }
                $d.NowStill = $true; $d.StillMin = [Math]::Round($tail)
            }

            $elapsedShift = ($refNow - $win.Start).TotalMinutes
            $stopSum = 0.0; foreach ($s in $stopList) { $stopSum += $s.Min }
            $runMin = $elapsedShift - $stopSum; if ($runMin -lt 0) { $runMin = 0 }
            $d.HasStops        = $true
            $d.Stops           = $stopList
            $d.StopCount       = $stopList.Count
            $d.StopMin         = $stopSum
            $d.ElapsedShiftMin = $elapsedShift
            $d.RunMin          = $runMin
            $d.AvailPct        = if ($elapsedShift -gt 0) { 100.0 * $runMin / $elapsedShift } else { 0 }
            $d.NetPerMin       = if ($runMin -gt 0) { $tot / $runMin } else { 0 }
            $d.LostBoxes       = $stopSum * $d.NetPerMin
            $d.BehindNow       = $tot - $targetPerMin * $elapsedShift
            if ($stopList.Count -gt 0) {
                $lg = $stopList | Sort-Object Min -Descending | Select-Object -First 1
                $d.LongestMin  = $lg.Min
                $d.LongestText = ('{0}-{1}' -f $lg.From.ToString('HH:mm'), $lg.To.ToString('HH:mm'))
            }

            # huidige RUN = laatste aaneengesloten blok rijen met het hoofdproduct
            $idx = $sProd.Count - 1
            while ($idx -gt 0 -and $sProd[$idx - 1] -eq $mainProd) { $idx-- }
            $runStart = $sTs[$idx]
            $runCount = $sProd.Count - $idx
            $mainCount = 0; if ($counts.ContainsKey($mainProd)) { $mainCount = [int]$counts[$mainProd] }

            $elapsed = ($refNow - $runStart).TotalMinutes
            $remain  = ($win.End - $refNow).TotalMinutes
            if ($remain -lt 0) { $remain = 0 }

            if ($elapsed -ge 1) {
                $perMin = $runCount / $elapsed
                $d.HasForecast  = $true
                $d.FcWeak       = ($elapsed -lt 5 -or $runCount -lt 5)   # te korte run -> voorlopig cijfer
                $d.MainCount    = $mainCount
                $d.RunCount     = $runCount
                $d.RunStartText = $runStart.ToString('HH:mm')
                $d.RefNowText   = $refNow.ToString('dd/MM HH:mm')
                $d.ElapsedMin   = [Math]::Round($elapsed)
                $d.RemainMin    = [Math]::Round($remain)
                $d.NowOffsetMin = ($refNow - $win.Start).TotalMinutes
                $d.PerMin       = $perMin
                $d.PerHour      = $perMin * 60
                $d.ProjTotal    = $tot + $perMin * $remain
                $d.ProjMain     = $mainCount + $perMin * $remain
                $d.BehindEnd    = $d.ProjTotal - $effTarget

                # tweede schatting: tempo van de laatste RecentMinutes minuten (hele lijn)
                $recWin = [Math]::Min([double]$RecentMinutes, ($refNow - $win.Start).TotalMinutes)
                if ($recWin -ge 1) {
                    $recFrom = $refNow.AddMinutes(-$recWin); $rc = 0
                    for ($k = $sTs.Count - 1; $k -ge 0; $k--) {
                        if ($sTs[$k] -le $recFrom) { break }
                        $rc++
                    }
                    $d.HasRecent    = $true
                    $d.RecentWin    = [Math]::Round($recWin)
                    $d.RecentCount  = $rc
                    $d.RecentPerMin = $rc / $recWin
                    $d.ProjRecent   = $tot + ($rc / $recWin) * $remain
                }

                if ($effTarget -gt 0) {
                    $d.ProjPct  = 100.0 * $d.ProjTotal / $effTarget
                    $d.ProjDiff = $d.ProjTotal - $effTarget
                    $todo = $effTarget - $tot
                    if ($todo -le 0)      { $d.EtaText = 'target al gehaald' }
                    elseif ($remain -le 0){ $d.EtaText = 'ploeg voorbij' }
                    else {
                        $d.NeedPerMin = $todo / $remain
                        if ($perMin -gt 0) {
                            $eta = $refNow.AddMinutes($todo / $perMin)
                            $d.EtaText = $eta.ToString('HH:mm')
                            if ($eta -gt $win.End) { $d.EtaText += ' (na ploegeinde)' }
                        }
                        else { $d.EtaText = 'niet haalbaar (tempo 0)' }
                    }
                }
            }
        }
    }
    catch { $d.Ok = $false; $d.Error = $_.Exception.Message }
    finally {
        Rel $ws; Rel $sheets
        if ($wb)    { try { $wb.Close($false) }    catch {}; Rel $wb }
        if ($blank) { try { $blank.Close($false) } catch {}; Rel $blank }
        Rel $wbs
        if ($excel) { try { $excel.Quit() } catch {}; Rel $excel }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        # vangnet: als het proces nog leeft (COM-invoegtoepassing houdt het vast), hard afsluiten - alleen ONS PID
        if ($xlPid -gt 0) {
            $p = Get-Process -Id $xlPid -ErrorAction SilentlyContinue
            if ($p) {
                Start-Sleep -Milliseconds 400
                $p = Get-Process -Id $xlPid -ErrorAction SilentlyContinue
                if ($p) { try { Stop-Process -Id $xlPid -Force } catch {} }
            }
        }
    }
    return [pscustomobject]$d
}

# ---- SVG staafgrafiek: dozen per minuut + horizontale target-lijn (server-side, geen externe libs) ----
function New-MinuteChartSvg($d) {
    $W = 1040; $L = 46; $R = 14; $T = 14; $B = 30; $H = 280
    $plotW = $W - $L - $R; $plotH = $H - $T - $B
    $n = [int]$d.ShiftMin; if ($n -le 0) { $n = 480 }
    $mins = $d.Minutes
    $tgt  = [double]$d.TargetPerMin
    $maxV = [double]$d.MaxPerMin
    $yTop = [Math]::Max($maxV, $tgt); if ($yTop -le 0) { $yTop = 1 }
    $yTop = $yTop * 1.15
    $yTopLabel = [int][Math]::Ceiling($yTop)
    $baseY = $T + $plotH

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<svg viewBox='0 0 $W $H' width='100%' preserveAspectRatio='xMidYMid meet' xmlns='http://www.w3.org/2000/svg' font-family='Segoe UI,system-ui,Arial,sans-serif'>")
    [void]$sb.Append("<rect x='$L' y='$T' width='$plotW' height='$plotH' fill='#0b1220' stroke='#334155' rx='6'/>")

    # het deel van de ploeg dat nog MOET komen (na 'nu') lichtjes arceren
    if ($d.HasForecast -and [double]$d.NowOffsetMin -lt $n) {
        $xNow = $L + ([double]$d.NowOffsetMin / $n) * $plotW
        $wFut = ($L + $plotW) - $xNow
        [void]$sb.Append("<rect x='$(SvgN $xNow)' y='$T' width='$(SvgN $wFut)' height='$plotH' fill='#1e293b' opacity='0.55'/>")
        [void]$sb.Append("<line x1='$(SvgN $xNow)' y1='$T' x2='$(SvgN $xNow)' y2='$baseY' stroke='#64748b' stroke-width='1' stroke-dasharray='3 3'/>")
    }

    # staven (dozen per minuut)
    $barW = [Math]::Max(1.2, ($plotW / $n) * 0.85)
    $bws  = SvgN $barW
    for ($i = 0; $i -lt $n; $i++) {
        $v = $mins[$i]; if ($v -le 0) { continue }
        $x = $L + ($i / $n) * $plotW
        $h = ($v / $yTop) * $plotH
        $y = $baseY - $h
        [void]$sb.Append("<rect x='$(SvgN $x)' y='$(SvgN $y)' width='$bws' height='$(SvgN $h)' fill='#38bdf8'/>")
    }

    # horizontale target-lijn
    $ty = $baseY - ($tgt / $yTop) * $plotH
    [void]$sb.Append("<line x1='$L' y1='$(SvgN $ty)' x2='$($L + $plotW)' y2='$(SvgN $ty)' stroke='#fbbf24' stroke-width='2' stroke-dasharray='6 4'/>")
    $tgtTxt = "target $(NF $d.Target)/ploeg = $(PF2 $tgt)/min"
    [void]$sb.Append("<text x='$($L + $plotW - 6)' y='$(SvgN ($ty - 6))' fill='#fbbf24' font-size='12' text-anchor='end'>$(HtmlEnc $tgtTxt)</text>")

    # gemiddeld tempo van de lopende run (= de lijn waarmee de prognose rekent)
    if ($d.HasForecast -and [double]$d.PerMin -gt 0) {
        $py = $baseY - ([double]$d.PerMin / $yTop) * $plotH
        [void]$sb.Append("<line x1='$L' y1='$(SvgN $py)' x2='$($L + $plotW)' y2='$(SvgN $py)' stroke='#34d399' stroke-width='2'/>")
        $pTxt = "tempo $(PF2 $d.PerMin)/min"
        $pLab = if ([Math]::Abs($py - $ty) -lt 14) { $py + 14 } else { $py - 6 }
        [void]$sb.Append("<text x='$($L + 6)' y='$(SvgN $pLab)' fill='#34d399' font-size='12'>$(HtmlEnc $pTxt)</text>")
    }

    # y-as labels (0 en max)
    [void]$sb.Append("<text x='$($L - 6)' y='$baseY' fill='#94a3b8' font-size='11' text-anchor='end'>0</text>")
    [void]$sb.Append("<text x='$($L - 6)' y='$($T + 10)' fill='#94a3b8' font-size='11' text-anchor='end'>$yTopLabel</text>")

    # x-as uur-ticks
    for ($m = 0; $m -le $n; $m += 60) {
        $x = $L + ($m / $n) * $plotW
        $lab = $d.ShiftStart.AddMinutes($m).ToString('HH:mm')
        [void]$sb.Append("<line x1='$(SvgN $x)' y1='$baseY' x2='$(SvgN $x)' y2='$($baseY + 4)' stroke='#475569'/>")
        [void]$sb.Append("<text x='$(SvgN $x)' y='$($baseY + 16)' fill='#94a3b8' font-size='11' text-anchor='middle'>$lab</text>")
    }
    [void]$sb.Append("</svg>")
    return $sb.ToString()
}

# ---- SVG: cumulatief gemaakt + gestippelde prognose tot ploegeinde + schuine target-lijn ----
function New-ForecastChartSvg($d) {
    $W = 1040; $L = 52; $R = 14; $T = 14; $B = 30; $H = 240
    $plotW = $W - $L - $R; $plotH = $H - $T - $B
    $n = [int]$d.ShiftMin; if ($n -le 0) { $n = 480 }
    $mins = $d.Minutes
    $baseY = $T + $plotH

    $nowOff = [double]$d.NowOffsetMin
    if ($nowOff -lt 0) { $nowOff = 0 }; if ($nowOff -gt $n) { $nowOff = $n }
    $proj = [double]$d.ProjTotal
    $yTop = [Math]::Max([Math]::Max([double]$d.Target, $proj), [double]$d.Total)
    if ($yTop -le 0) { $yTop = 1 }
    $yTop = $yTop * 1.12

    # let op: PowerShell-variabelen zijn hoofdletter-ONgevoelig -> geen $X/$Y (botst met $x/$y)
    $MapX = { param($m) $L + ($m / $n) * $plotW }
    $MapY = { param($v) $baseY - ($v / $yTop) * $plotH }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<svg viewBox='0 0 $W $H' width='100%' preserveAspectRatio='xMidYMid meet' xmlns='http://www.w3.org/2000/svg' font-family='Segoe UI,system-ui,Arial,sans-serif'>")
    [void]$sb.Append("<rect x='$L' y='$T' width='$plotW' height='$plotH' fill='#0b1220' stroke='#334155' rx='6'/>")

    # toekomst-zone
    $xNow = & $MapX $nowOff
    if ($nowOff -lt $n) {
        [void]$sb.Append("<rect x='$(SvgN $xNow)' y='$T' width='$(SvgN (($L + $plotW) - $xNow))' height='$plotH' fill='#1e293b' opacity='0.55'/>")
    }

    # schuine target-lijn (0 -> target over de hele ploeg)
    [void]$sb.Append("<line x1='$L' y1='$baseY' x2='$($L + $plotW)' y2='$(SvgN (& $MapY ([double]$d.Target)))' stroke='#fbbf24' stroke-width='2' stroke-dasharray='6 4'/>")
    # label op ~70% van de schuine lijn: bij het rechteruiteinde botst het met 'prognose'
    [void]$sb.Append("<text x='$(SvgN (& $MapX ($n * 0.7)))' y='$(SvgN ((& $MapY (0.7 * [double]$d.Target)) - 7))' fill='#fbbf24' font-size='12' text-anchor='end'>target $(NF $d.Target)</text>")

    # cumulatief gemaakt (tot 'nu')
    $pts = New-Object System.Text.StringBuilder
    [void]$pts.Append("$(SvgN $L),$(SvgN $baseY) ")
    $cum = 0; $upTo = [int][Math]::Floor($nowOff)
    if ($upTo -gt $n) { $upTo = $n }
    for ($i = 0; $i -lt $upTo; $i++) {
        $cum += $mins[$i]
        if ($mins[$i] -gt 0 -or ($i % 5) -eq 0) {
            [void]$pts.Append("$(SvgN (& $MapX ($i + 1))),$(SvgN (& $MapY $cum)) ")
        }
    }
    [void]$pts.Append("$(SvgN $xNow),$(SvgN (& $MapY ([double]$d.Total)))")
    [void]$sb.Append("<polyline points='$($pts.ToString())' fill='none' stroke='#38bdf8' stroke-width='2.5' stroke-linejoin='round'/>")

    # prognose: van (nu, gemaakt) naar (ploegeinde, verwacht)
    $yNow = & $MapY ([double]$d.Total); $yEnd = & $MapY $proj; $xEnd = $L + $plotW
    [void]$sb.Append("<line x1='$(SvgN $xNow)' y1='$(SvgN $yNow)' x2='$(SvgN $xEnd)' y2='$(SvgN $yEnd)' stroke='#34d399' stroke-width='2.5' stroke-dasharray='7 5'/>")
    [void]$sb.Append("<circle cx='$(SvgN $xNow)' cy='$(SvgN $yNow)' r='4' fill='#38bdf8'/>")
    [void]$sb.Append("<circle cx='$(SvgN $xEnd)' cy='$(SvgN $yEnd)' r='4' fill='#34d399'/>")
    $lblY = if ($yEnd -lt ($T + 18)) { $yEnd + 16 } else { $yEnd - 8 }
    [void]$sb.Append("<text x='$($xEnd - 6)' y='$(SvgN $lblY)' fill='#34d399' font-size='13' font-weight='600' text-anchor='end'>prognose $(NF $proj)</text>")

    # tweede schatting (recent tempo) als dunne stippellijn
    if ($d.HasRecent -and [double]$d.ProjRecent -gt 0) {
        $yEnd2 = & $MapY ([double]$d.ProjRecent)
        [void]$sb.Append("<line x1='$(SvgN $xNow)' y1='$(SvgN $yNow)' x2='$(SvgN $xEnd)' y2='$(SvgN $yEnd2)' stroke='#a78bfa' stroke-width='1.5' stroke-dasharray='3 4'/>")
    }

    # 'nu'-markering
    [void]$sb.Append("<line x1='$(SvgN $xNow)' y1='$T' x2='$(SvgN $xNow)' y2='$baseY' stroke='#64748b' stroke-width='1' stroke-dasharray='3 3'/>")

    # assen
    [void]$sb.Append("<text x='$($L - 6)' y='$baseY' fill='#94a3b8' font-size='11' text-anchor='end'>0</text>")
    [void]$sb.Append("<text x='$($L - 6)' y='$($T + 10)' fill='#94a3b8' font-size='11' text-anchor='end'>$([int][Math]::Ceiling($yTop))</text>")
    for ($m = 0; $m -le $n; $m += 60) {
        $x = & $MapX $m
        $lab = $d.ShiftStart.AddMinutes($m).ToString('HH:mm')
        [void]$sb.Append("<line x1='$(SvgN $x)' y1='$baseY' x2='$(SvgN $x)' y2='$($baseY + 4)' stroke='#475569'/>")
        [void]$sb.Append("<text x='$(SvgN $x)' y='$($baseY + 16)' fill='#94a3b8' font-size='11' text-anchor='middle'>$lab</text>")
    }
    [void]$sb.Append("</svg>")
    return $sb.ToString()
}

# ---- SVG: strook 'loopt / staat stil' over de ploeg + uur-ticks ----
function New-StopStripSvg($d) {
    $W = 1040; $L = 46; $R = 14; $T = 8; $H = 56; $barH = 26
    $plotW = $W - $L - $R
    $n = [int]$d.ShiftMin; if ($n -le 0) { $n = 480 }
    $baseY = $T + $barH

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<svg viewBox='0 0 $W $H' width='100%' preserveAspectRatio='xMidYMid meet' xmlns='http://www.w3.org/2000/svg' font-family='Segoe UI,system-ui,Arial,sans-serif'>")
    # hele ploeg = nog te gaan (grijs), verstreken deel = draait (groen)
    [void]$sb.Append("<rect x='$L' y='$T' width='$plotW' height='$barH' fill='#1e293b' stroke='#334155' rx='4'/>")
    $nowOff = [double]$d.NowOffsetMin; if ($nowOff -lt 0) { $nowOff = 0 }; if ($nowOff -gt $n) { $nowOff = $n }
    $wNow = ($nowOff / $n) * $plotW
    [void]$sb.Append("<rect x='$L' y='$T' width='$(SvgN $wNow)' height='$barH' fill='#10b981' rx='4'/>")

    # stilstanden er rood overheen
    foreach ($s in $d.Stops) {
        $a = ($s.From - $d.ShiftStart).TotalMinutes; if ($a -lt 0) { $a = 0 }
        $b = ($s.To   - $d.ShiftStart).TotalMinutes; if ($b -gt $n) { $b = $n }
        if ($b -le $a) { continue }
        $x = $L + ($a / $n) * $plotW
        $w = [Math]::Max(2.0, (($b - $a) / $n) * $plotW)
        $fill = if ($s.Kind -eq 'nu') { '#f97316' } else { '#ef4444' }
        [void]$sb.Append("<rect x='$(SvgN $x)' y='$T' width='$(SvgN $w)' height='$barH' fill='$fill'/>")
        # alleen de lange stops krijgen een tijdlabel (anders wordt het een kluwen)
        if ($s.Min -ge 10) {
            [void]$sb.Append("<text x='$(SvgN ($x + $w / 2))' y='$($T + 17)' fill='#450a0a' font-size='11' font-weight='600' text-anchor='middle'>$([int][Math]::Round($s.Min)) min</text>")
        }
    }

    # uur-ticks
    for ($m = 0; $m -le $n; $m += 60) {
        $x = $L + ($m / $n) * $plotW
        $lab = $d.ShiftStart.AddMinutes($m).ToString('HH:mm')
        [void]$sb.Append("<line x1='$(SvgN $x)' y1='$baseY' x2='$(SvgN $x)' y2='$($baseY + 4)' stroke='#475569'/>")
        [void]$sb.Append("<text x='$(SvgN $x)' y='$($baseY + 16)' fill='#94a3b8' font-size='11' text-anchor='middle'>$lab</text>")
    }
    [void]$sb.Append("<text x='$($L - 6)' y='$($T + 17)' fill='#94a3b8' font-size='11' text-anchor='end'>lijn</text>")
    [void]$sb.Append("</svg>")
    return $sb.ToString()
}

# ---- SVG: achterstand/voorsprong t.o.v. het target-tempo (0 = precies op schema) ----
function New-BehindChartSvg($d) {
    $W = 1040; $L = 52; $R = 14; $T = 14; $B = 30; $H = 200
    $plotW = $W - $L - $R; $plotH = $H - $T - $B
    $n = [int]$d.ShiftMin; if ($n -le 0) { $n = 480 }
    $mins = $d.Minutes
    $pace = [double]$d.TargetPerMin
    $nowOff = [double]$d.NowOffsetMin; if ($nowOff -lt 0) { $nowOff = 0 }; if ($nowOff -gt $n) { $nowOff = $n }
    $upTo = [int][Math]::Floor($nowOff); if ($upTo -gt $n) { $upTo = $n }

    # verloop van (gemaakt - target-tempo * minuten)
    $vals = New-Object 'System.Collections.Generic.List[double]'
    $cum = 0.0
    for ($i = 0; $i -lt $upTo; $i++) { $cum += $mins[$i]; $vals.Add($cum - $pace * ($i + 1)) }
    $lo = 0.0; $hi = 0.0
    foreach ($v in $vals) { if ($v -lt $lo) { $lo = $v }; if ($v -gt $hi) { $hi = $v } }
    $end = [double]$d.BehindEnd
    if ($end -lt $lo) { $lo = $end }; if ($end -gt $hi) { $hi = $end }
    $span = [Math]::Max([Math]::Abs($lo), [Math]::Abs($hi)); if ($span -le 0) { $span = 10 }
    $span = $span * 1.2

    $MapX = { param($m) $L + ($m / $n) * $plotW }
    $MapY = { param($v) $T + $plotH / 2 - ($v / $span) * ($plotH / 2) }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<svg viewBox='0 0 $W $H' width='100%' preserveAspectRatio='xMidYMid meet' xmlns='http://www.w3.org/2000/svg' font-family='Segoe UI,system-ui,Arial,sans-serif'>")
    [void]$sb.Append("<rect x='$L' y='$T' width='$plotW' height='$plotH' fill='#0b1220' stroke='#334155' rx='6'/>")
    if ($nowOff -lt $n) {
        $xn = & $MapX $nowOff
        [void]$sb.Append("<rect x='$(SvgN $xn)' y='$T' width='$(SvgN (($L + $plotW) - $xn))' height='$plotH' fill='#1e293b' opacity='0.55'/>")
    }
    # nullijn = precies op target-tempo
    $y0 = & $MapY 0
    [void]$sb.Append("<line x1='$L' y1='$(SvgN $y0)' x2='$($L + $plotW)' y2='$(SvgN $y0)' stroke='#fbbf24' stroke-width='2' stroke-dasharray='6 4'/>")
    [void]$sb.Append("<text x='$($L + 6)' y='$(SvgN ($y0 - 6))' fill='#fbbf24' font-size='12'>op target-tempo</text>")

    # de lijn zelf (kleur = staan we nu voor of achter)
    $behind = [double]$d.BehindNow
    $col = if ($behind -ge 0) { '#34d399' } else { '#f87171' }
    if ($vals.Count -gt 1) {
        $pts = New-Object System.Text.StringBuilder
        for ($i = 0; $i -lt $vals.Count; $i++) {
            if (($i % 2) -eq 0 -or $i -eq ($vals.Count - 1)) {
                [void]$pts.Append("$(SvgN (& $MapX ($i + 1))),$(SvgN (& $MapY $vals[$i])) ")
            }
        }
        [void]$sb.Append("<polyline points='$($pts.ToString().Trim())' fill='none' stroke='$col' stroke-width='2.5' stroke-linejoin='round'/>")
    }

    # doortrekken naar ploegeinde volgens de prognose
    $xNow = & $MapX $nowOff; $yNow = & $MapY $behind
    if ($d.HasForecast) {
        $xEnd = $L + $plotW; $yEnd = & $MapY $end
        [void]$sb.Append("<line x1='$(SvgN $xNow)' y1='$(SvgN $yNow)' x2='$(SvgN $xEnd)' y2='$(SvgN $yEnd)' stroke='#34d399' stroke-width='2' stroke-dasharray='7 5'/>")
        [void]$sb.Append("<circle cx='$(SvgN $xEnd)' cy='$(SvgN $yEnd)' r='4' fill='#34d399'/>")
        $endTxt = if ($end -ge 0) { "einde ploeg +$(NF $end)" } else { "einde ploeg $(NF $end)" }
        $ly = if ($yEnd -lt ($T + 18)) { $yEnd + 16 } else { $yEnd - 8 }
        [void]$sb.Append("<text x='$($L + $plotW - 6)' y='$(SvgN $ly)' fill='#34d399' font-size='13' font-weight='600' text-anchor='end'>$(HtmlEnc $endTxt)</text>")
    }
    [void]$sb.Append("<circle cx='$(SvgN $xNow)' cy='$(SvgN $yNow)' r='4' fill='$col'/>")
    $nowTxt = if ($behind -ge 0) { "+$(NF $behind)" } else { (NF $behind) }
    $nly = if ($yNow -lt ($T + 18)) { $yNow + 16 } else { $yNow - 8 }
    [void]$sb.Append("<text x='$(SvgN ($xNow - 6))' y='$(SvgN $nly)' fill='$col' font-size='13' font-weight='600' text-anchor='end'>$(HtmlEnc $nowTxt)</text>")

    # assen
    [void]$sb.Append("<text x='$($L - 6)' y='$($T + 10)' fill='#94a3b8' font-size='11' text-anchor='end'>+$([int][Math]::Ceiling($span))</text>")
    [void]$sb.Append("<text x='$($L - 6)' y='$($T + $plotH)' fill='#94a3b8' font-size='11' text-anchor='end'>-$([int][Math]::Ceiling($span))</text>")
    for ($m = 0; $m -le $n; $m += 60) {
        $x = & $MapX $m
        $lab = $d.ShiftStart.AddMinutes($m).ToString('HH:mm')
        [void]$sb.Append("<line x1='$(SvgN $x)' y1='$($T + $plotH)' x2='$(SvgN $x)' y2='$($T + $plotH + 4)' stroke='#475569'/>")
        [void]$sb.Append("<text x='$(SvgN $x)' y='$($T + $plotH + 16)' fill='#94a3b8' font-size='11' text-anchor='middle'>$lab</text>")
    }
    [void]$sb.Append("</svg>")
    return $sb.ToString()
}

# ---- console-weergave ----
function Render-Console($d) {
    if ($d.Warning) { Write-Host "WAARSCHUWING: $($d.Warning)" -ForegroundColor Yellow }
    Write-Host ""
    Write-Host ("=========== DOZEN PER PRODUCTTYPE - {0} ===========" -f $d.ShiftLabel) -ForegroundColor Cyan
    Write-Host ("Tijd : {0}    interval {1}" -f $d.NowText, $d.WindowText)
    Write-Host ("Blad : {0}    bestand: {1} (opgeslagen {2})" -f $d.Sheet, $d.BoxFile, $d.FileTimeText)
    Write-Host "=====================================================" -ForegroundColor Cyan
    if ($d.Error) { Write-Host "FOUT: $($d.Error)" -ForegroundColor Red; return }
    if ($d.Rows.Count -eq 0) {
        Write-Host "Geen dozen in deze ploeg-interval (geparseerde rijen: $($d.ParsedRows))." -ForegroundColor Yellow
        Write-Host ("Target ploeg : {0}  ({1})" -f (NF $d.Target), $d.TargetSource) -ForegroundColor DarkGray
        return
    }

    Write-Host ("{0,-11} {1,8} {2,8} {3,8}  {4}" -f "Producttype", "Dozen", "Plan", "Rest", "Omschrijving") -ForegroundColor White
    foreach ($r in $d.Rows) {
        $mark = if ($r.IsMain) { " << nu" } else { "" }
        $col  = if ($r.IsMain) { 'Green' } else { 'Gray' }
        $planTxt = if ($r.Plan -gt 0) { NF $r.Plan } else { "-" }
        $restTxt = if ($r.Plan -gt 0) { NF ($r.Plan - $r.Count) } else { "-" }
        $desc = if ($r.Desc.Length -gt 26) { $r.Desc.Substring(0, 26) } else { $r.Desc }
        Write-Host ("{0,-11} {1,8} {2,8} {3,8}  {4}{5}" -f $r.Product, (NF $r.Count), $planTxt, $restTxt, $desc, $mark) -ForegroundColor $col
    }
    Write-Host ("{0}" -f ('-' * 40)) -ForegroundColor DarkGray
    Write-Host ("{0,-11} {1,8} {2,8} {3,8}" -f "Totaal", (NF $d.Total), (NF $d.Target), (NF ($d.Target - $d.Total))) -ForegroundColor Green
    Write-Host ("Target ploeg : {0}  (gemaakt {1} %, tempo-lijn {2} dozen/min)" -f (NF $d.Target), (PF $d.Pct), (PF2 $d.TargetPerMin)) -ForegroundColor DarkGray
    Write-Host ("Target bron  : {0}" -f $d.TargetSource) -ForegroundColor DarkGray
    if ($d.LastText) { Write-Host ("Laatste doos : {0}" -f $d.LastText) -ForegroundColor DarkGray }
    $skipTxt = if ($d.StartRowSkipped) { " (rij 11 = startwaarde ophaalvenster, niet geteld)" } else { "" }
    Write-Host ("Piek dozen/min : {0}    geparseerde rijen: {1}{2}" -f (NF $d.MaxPerMin), $d.ParsedRows, $skipTxt) -ForegroundColor DarkGray

    if ($d.HasForecast) {
        Write-Host ""
        Write-Host ("PROGNOSE EINDE PLOEG (hoofdproduct {0} = laatste rij, data t/m {1}):" -f $d.MainProduct, $d.RefNowText) -ForegroundColor White
        Write-Host ("  Tempo                : {0} dozen/min  ({1} dozen/uur)" -f (PF2 $d.PerMin), (NF $d.PerHour))
        if ($d.HasRecent) {
            Write-Host ("  Tempo laatste {0,-3} min: {1} dozen/min  ({2} dozen)" -f (NF $d.RecentWin), (PF2 $d.RecentPerMin), (NF $d.RecentCount)) -ForegroundColor DarkGray
        }
        Write-Host ("  Verwacht einde ploeg : {0} dozen   ({1} % van target, {2}{3})" -f (NF $d.ProjTotal), (PF $d.ProjPct), $(if ($d.ProjDiff -ge 0) { '+' } else { '' }), (NF $d.ProjDiff)) -ForegroundColor $(if ($d.ProjDiff -ge 0) { 'Cyan' } else { 'Yellow' })
        if ($d.HasRecent) {
            Write-Host ("  Op recent tempo      : {0} dozen" -f (NF $d.ProjRecent)) -ForegroundColor DarkGray
        }
        if ($d.Rows.Count -gt 1) {
            Write-Host ("  Waarvan {0,-9}    : {1} dozen (nu {2})" -f $d.MainProduct, (NF $d.ProjMain), (NF $d.MainCount)) -ForegroundColor DarkGray
        }
        if ($d.NeedPerMin -gt 0) {
            Write-Host ("  Nodig voor target    : {0} dozen/min ({1} dozen in {2} min)" -f (PF2 $d.NeedPerMin), (NF ($d.Target - $d.Total)), (NF $d.RemainMin)) -ForegroundColor DarkGray
        }
        if ($d.EtaText) { Write-Host ("  Target bereikt om    : {0}" -f $d.EtaText) -ForegroundColor DarkGray }
        Write-Host ("  ({0} dozen in de lopende run sinds {1}, {2} min; nog {3} min tot ploegeinde)" -f (NF $d.RunCount), $d.RunStartText, (NF $d.ElapsedMin), (NF $d.RemainMin)) -ForegroundColor DarkGray
        if ($d.FcWeak) { Write-Host "  LET OP: korte run - prognose is nog voorlopig." -ForegroundColor Yellow }
    }

    if ($d.HasStops) {
        Write-Host ""
        Write-Host ("STILSTAND EN ACHTERSTAND (gat > {0} min telt als stilstand):" -f (PF $d.StopLimit)) -ForegroundColor White
        Write-Host ("  Stilstand            : {0} min in {1} stops" -f (NF $d.StopMin), $d.StopCount) -ForegroundColor $(if ($d.StopMin -gt 0) { 'Yellow' } else { 'Gray' })
        if ($d.LongestText) { Write-Host ("  Langste stop         : {0} min ({1})" -f (PF $d.LongestMin), $d.LongestText) -ForegroundColor DarkGray }
        Write-Host ("  Draaitijd            : {0} van {1} min ({2} %)" -f (NF $d.RunMin), (NF $d.ElapsedShiftMin), (PF $d.AvailPct)) -ForegroundColor DarkGray
        Write-Host ("  Netto tempo (draait) : {0} dozen/min" -f (PF2 $d.NetPerMin)) -ForegroundColor DarkGray
        Write-Host ("  Verlies stilstand    : ~{0} dozen" -f (NF $d.LostBoxes)) -ForegroundColor DarkGray
        $bn = [double]$d.BehindNow
        Write-Host ("  Achterstand nu       : {0}{1} dozen t.o.v. target-tempo" -f $(if ($bn -ge 0) { '+' } else { '' }), (NF $bn)) -ForegroundColor $(if ($bn -ge 0) { 'Green' } else { 'Yellow' })
        if ($d.NowStill) { Write-Host ("  LIJN STAAT NU STIL   : sinds {0} min geen doos" -f (NF $d.StillMin)) -ForegroundColor Red }
        $top = $d.Stops | Sort-Object Min -Descending | Select-Object -First 5
        foreach ($s in $top) {
            Write-Host ("    {0}-{1}  {2,6} min  {3}" -f $s.From.ToString('HH:mm'), $s.To.ToString('HH:mm'), (PF $s.Min), $s.Kind) -ForegroundColor DarkGray
        }
    }
}

# ---- HTML-dashboard ----
function Render-Html($d) {
    $refresh = $IntervalSeconds
    $load = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')
    $meta = "$($d.NowText) &middot; $(HtmlEnc $d.ShiftLabel) &middot; interval $(HtmlEnc $d.WindowText)"
    $meta += " &middot; $(HtmlEnc $d.Sheet) &middot; $(HtmlEnc $d.BoxFile)"
    $warnHtml = if ($d.Warning) { "<div class='warn'>$(HtmlEnc $d.Warning)</div>" } else { "" }

    if ($d.Error) {
        $bodyHtml = "<div class='err'>FOUT: $(HtmlEnc $d.Error)</div>"
    }
    else {
        $pct = PF $d.Pct
        $cards = "<div class='cards'>" +
            "<div class='card'><div class='lbl'>Gemaakt (ploeg)</div><div class='val done'>$(NF $d.Total)</div><div class='sub'>opgeslagen $(HtmlEnc $d.FileTimeText)</div></div>" +
            "<div class='card'><div class='lbl'>Target ploeg</div><div class='val'>$(NF $d.Target)</div><div class='sub'>$pct % &middot; tempo $(PF2 $d.TargetPerMin) dozen/min<br><span class='bron'>$(HtmlEnc $d.TargetSource)</span></div></div>"
        if ($d.HasForecast) {
            $projCls = if ($d.ProjDiff -ge 0) { "done" } else { "behind" }
            $diffTxt = if ($d.ProjDiff -ge 0) { "+$(NF $d.ProjDiff)" } else { (NF $d.ProjDiff) }
            $cards += "<div class='card hi'><div class='lbl'>Verwacht einde ploeg</div><div class='val $projCls'>$(NF $d.ProjTotal)</div>" +
                      "<div class='sub'>$(PF $d.ProjPct) % van target &middot; $diffTxt</div></div>" +
                      "<div class='card'><div class='lbl'>Tempo nu</div><div class='val'>$(PF2 $d.PerMin)</div><div class='sub'>dozen/min &middot; $(NF $d.PerHour) per uur</div></div>"
        }
        $cards += "</div>"
        $chart = "<h2 class='sec'>Dozen per minuut</h2><div class='chart'>$(New-MinuteChartSvg $d)</div>"

        $fc = ""
        if ($d.HasForecast) {
            $weak = if ($d.FcWeak) { "<div class='warn'>Korte run ($(NF $d.RunCount) dozen in $(NF $d.ElapsedMin) min) &mdash; prognose is nog voorlopig.</div>" } else { "" }
            $recHtml = ""
            if ($d.HasRecent) {
                $recHtml = "<div class='card'><div class='lbl'>Laatste $(NF $d.RecentWin) min</div><div class='val'>$(PF2 $d.RecentPerMin)</div>" +
                           "<div class='sub'>dozen/min &middot; prognose $(NF $d.ProjRecent)</div></div>"
            }
            $needHtml = ""
            if ($d.NeedPerMin -gt 0) {
                $needCls = if ($d.NeedPerMin -gt $d.PerMin) { "behind" } else { "done" }
                $needHtml = "<div class='card'><div class='lbl'>Nodig voor target</div><div class='val $needCls'>$(PF2 $d.NeedPerMin)</div>" +
                            "<div class='sub'>dozen/min &middot; nog $(NF ($d.Target - $d.Total)) in $(NF $d.RemainMin) min</div></div>"
            }
            $etaHtml = ""
            if ($d.EtaText) {
                $etaHtml = "<div class='card'><div class='lbl'>Target bereikt om</div><div class='val small'>$(HtmlEnc $d.EtaText)</div><div class='sub'>bij huidig tempo</div></div>"
            }
            $fc = "<h2 class='sec'>Prognose einde ploeg &mdash; hoofdproduct $(HtmlEnc $d.MainProduct) " +
                  "<span class='bron'>(laatste rij &middot; data t/m $(HtmlEnc $d.RefNowText))</span></h2>$weak" +
                  "<div class='chart'>$(New-ForecastChartSvg $d)</div>" +
                  "<div class='cards'>$recHtml$needHtml$etaHtml</div>" +
                  "<div class='wdesc'>Run van $(HtmlEnc $d.MainProduct): $(NF $d.RunCount) dozen sinds $(HtmlEnc $d.RunStartText) ($(NF $d.ElapsedMin) min) " +
                  "&middot; nog $(NF $d.RemainMin) min tot ploegeinde &middot; verwacht voor dit product $(NF $d.ProjMain) dozen</div>"
        }

        $st = ""
        if ($d.HasStops) {
            $stillCls = if ($d.StopMin -gt 0) { "behind" } else { "done" }
            $availCls = if ($d.AvailPct -ge 95) { "done" } else { "behind" }
            $bn = [double]$d.BehindNow
            $bnCls = if ($bn -ge 0) { "done" } else { "behind" }
            $bnTxt = if ($bn -ge 0) { "+$(NF $bn)" } else { (NF $bn) }
            $nowStillHtml = if ($d.NowStill) { "<div class='err'>Lijn staat nu stil: al $(NF $d.StillMin) min geen doos (laatste $(HtmlEnc $d.LastText)).</div>" } else { "" }
            $stopRows = ""
            foreach ($s in ($d.Stops | Sort-Object Min -Descending | Select-Object -First 6)) {
                $kindTxt = switch ($s.Kind) { 'opstart' { 'opstart ploeg' } 'nu' { 'staat nu stil' } default { 'stop' } }
                $stopRows += "<tr><td class='sku'>$($s.From.ToString('HH:mm')) &ndash; $($s.To.ToString('HH:mm'))</td><td class='num'>$(PF $s.Min) min</td><td>$kindTxt</td></tr>"
            }
            $st = "<h2 class='sec'>Stilstand <span class='bron'>(gat &gt; $(PF $d.StopLimit) min telt als stilstand)</span></h2>$nowStillHtml" +
                  "<div class='chart'>$(New-StopStripSvg $d)</div>" +
                  "<div class='cards'>" +
                  "<div class='card'><div class='lbl'>Stilstand</div><div class='val $stillCls'>$(NF $d.StopMin)</div><div class='sub'>min in $($d.StopCount) stops &middot; langste $(PF $d.LongestMin) min $(HtmlEnc $d.LongestText)</div></div>" +
                  "<div class='card'><div class='lbl'>Draaitijd</div><div class='val $availCls'>$(PF $d.AvailPct)</div><div class='sub'>% &middot; $(NF $d.RunMin) van $(NF $d.ElapsedShiftMin) min</div></div>" +
                  "<div class='card'><div class='lbl'>Netto tempo</div><div class='val'>$(PF2 $d.NetPerMin)</div><div class='sub'>dozen/min als de lijn draait</div></div>" +
                  "<div class='card'><div class='lbl'>Verlies stilstand</div><div class='val behind'>$(NF $d.LostBoxes)</div><div class='sub'>dozen bij netto tempo</div></div>" +
                  "</div>" +
                  "<h2 class='sec'>Achterstand t.o.v. target-tempo</h2>" +
                  "<div class='chart'>$(New-BehindChartSvg $d)</div>" +
                  "<div class='cards'><div class='card hi'><div class='lbl'>Achterstand nu</div><div class='val $bnCls'>$bnTxt</div>" +
                  "<div class='sub'>dozen t.o.v. $(PF2 $d.TargetPerMin)/min</div></div></div>"
            if ($stopRows) {
                $st += "<table class='shift'><thead><tr><th>Stilstand</th><th>Duur</th><th>Soort</th></tr></thead><tbody>$stopRows</tbody></table>"
            }
        }

        if ($d.Rows.Count -gt 0) {
            $rows = ""
            foreach ($r in $d.Rows) {
                $cls  = if ($r.IsMain) { "cur" } else { "" }
                $mark = if ($r.IsMain) { " <span class='nu'>nu</span>" } else { "" }
                $prog = if ($r.IsMain -and $d.HasForecast) { NF $d.ProjMain } else { "&mdash;" }
                $planTxt = if ($r.Plan -gt 0) { NF $r.Plan } else { "&mdash;" }
                $restVal = $r.Plan - $r.Count
                $restTxt = if ($r.Plan -gt 0) { "<span class='$(if ($restVal -gt 0) { 'behind' } else { 'done' })'>$(NF $restVal)</span>" } else { "&mdash;" }
                $rows += "<tr class='$cls'><td class='sku'>$(HtmlEnc $r.Product)$mark</td><td>$(HtmlEnc $r.Desc)</td><td class='num'>$(NF $r.Count)</td><td class='num'>$planTxt</td><td class='num'>$restTxt</td><td class='num'>$prog</td></tr>"
            }
            $totProj = if ($d.HasForecast) { NF $d.ProjTotal } else { "&mdash;" }
            $totRest = $d.Target - $d.Total
            $totalRow = "<tr class='tot'><td colspan='2'>Totaal</td><td class='num'>$(NF $d.Total)</td><td class='num'>$(NF $d.Target)</td><td class='num'>$(NF $totRest)</td><td class='num'>$totProj</td></tr>"
            $table = "<h2 class='sec'>Dozen per producttype <span class='bron'>(plan: $(HtmlEnc $d.TargetSource))</span></h2>" +
                     "<table class='shift'><thead><tr><th>Producttype</th><th>Product</th><th>Dozen nu</th><th>Plan ploeg</th><th>Nog te doen</th><th>Verwacht einde ploeg</th></tr></thead>" +
                     "<tbody>$rows$totalRow</tbody></table>"
            $skipTxt  = if ($d.StartRowSkipped) { " &middot; rij 11 = startwaarde ophaalvenster (niet geteld)" } else { "" }
            $lastHtml = if ($d.LastText) { "<div class='wdesc'>Laatste doos: $(HtmlEnc $d.LastText) &middot; geparseerde rijen: $($d.ParsedRows)$skipTxt</div>" } else { "" }
        }
        else {
            $table = "<div class='warn'>Geen dozen in deze ploeg-interval. Geparseerde rijen: $($d.ParsedRows). " +
                     "Target ploeg $(NF $d.Target) &middot; $(HtmlEnc $d.TargetSource).</div>"
            $lastHtml = ""
        }
        $bodyHtml = "$cards$chart$fc$st$table$lastHtml"
    }

    $css = "*{box-sizing:border-box}body{margin:0;background:#0f172a;color:#e2e8f0;font-family:Segoe UI,system-ui,Arial,sans-serif}" +
           ".wrap{max-width:960px;margin:0 auto;padding:24px}h1{font-size:22px;margin:0 0 4px}.meta{color:#94a3b8;font-size:13px;margin-bottom:16px}" +
           ".sec{font-size:14px;color:#cbd5e1;margin:18px 0 8px;text-transform:uppercase;letter-spacing:.03em}" +
           "table.shift{width:100%;border-collapse:collapse;background:#1e293b;border-radius:10px;overflow:hidden}" +
           ".shift th{text-align:left;font-size:11px;text-transform:uppercase;color:#94a3b8;padding:10px 12px;background:#172033}" +
           ".shift td{padding:11px 12px;border-top:1px solid #334155;font-size:16px}.shift td.num{text-align:right;font-variant-numeric:tabular-nums}" +
           ".shift td.sku{font-weight:600}.done{color:#34d399;font-weight:600}.behind{color:#fbbf24;font-weight:600}" +
           ".shift tr.cur td{background:#14321f}" +
           ".nu{background:#22c55e;color:#06210f;font-size:11px;padding:1px 7px;border-radius:8px;margin-left:4px;vertical-align:middle}" +
           ".tot td{font-weight:700;border-top:2px solid #475569;color:#cbd5e1}" +
           ".bron{color:#94a3b8;font-weight:400;font-size:13px;text-transform:none;letter-spacing:0}" +
           ".wdesc{color:#94a3b8;font-size:13px;margin-top:10px}.cards{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:8px}" +
           ".card{flex:1;min-width:200px;background:#1e293b;border-radius:10px;padding:14px 16px}" +
           ".card.hi{outline:1px solid #334155;background:#172033}" +
           ".card .lbl{font-size:11px;color:#94a3b8;text-transform:uppercase}.card .val{font-size:34px;font-weight:700;margin-top:4px;font-variant-numeric:tabular-nums}" +
           ".card .val.small{font-size:26px}.card .sub{font-size:13px;color:#94a3b8;margin-top:2px}" +
           ".chart{background:#0b1220;border-radius:10px;padding:8px 8px 2px;overflow-x:auto}.chart svg{display:block}" +
           ".warn{background:#3a2e10;color:#fcd34d;padding:10px 12px;border-radius:8px;margin:10px 0}" +
           ".err{background:#3a1212;color:#fca5a5;padding:10px 12px;border-radius:8px;margin:10px 0}" +
           ".foot{color:#64748b;font-size:12px;margin-top:22px}"

    return "<!doctype html><html lang='nl'><head><meta charset='utf-8'><meta http-equiv='refresh' content='$refresh'>" +
           "<meta name='viewport' content='width=device-width, initial-scale=1'><title>BoxCount $(HtmlEnc $d.Sheet)</title><style>$css</style></head>" +
           "<body><div class='wrap'><h1>Dozen per producttype &mdash; huidige ploeg</h1><div class='meta'>$meta</div>$warnHtml$bodyHtml" +
           "<div class='foot'>Laatst geladen: $load &middot; ververst elke ${refresh}s</div></div></body></html>"
}

function Start-WebServer([int]$port) {
    $listener = New-Object System.Net.Sockets.TcpListener ([System.Net.IPAddress]::Loopback, $port)
    try { $listener.Start() }
    catch { Write-Host "FOUT: kan poort $port niet openen: $($_.Exception.Message)" -ForegroundColor Red; return }
    $url = "http://127.0.0.1:$port/"
    Write-Host "Webserver actief op $url" -ForegroundColor Green
    Write-Host "(Ctrl+C om te stoppen)" -ForegroundColor DarkGray
    if (-not $NoBrowser) { try { Start-Process $url } catch {} }

    $cacheHtml = $null; $cacheTime = [datetime]::MinValue
    try {
        while ($true) {
            $client = $listener.AcceptTcpClient()
            try {
                $client.ReceiveTimeout = 1500
                $stream = $client.GetStream()
                $buf = New-Object byte[] 4096
                try { [void]$stream.Read($buf, 0, $buf.Length) } catch {}
                $age = ([datetime]::Now - $cacheTime).TotalSeconds
                if ($null -eq $cacheHtml -or $age -ge $IntervalSeconds) {   # herlezen hooguit ~1x per minuut
                    $cacheHtml = Render-Html (Get-BoxData)
                    $cacheTime = [datetime]::Now
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

# ---- configuratie laden (config.txt naast het script; wordt NIET gewijzigd) ----
$cfg = Read-ConfigFile $ConfigFile
if (-not $script:HasBox) {
    if ($cfg.ContainsKey('BoxPrintingFile') -and $cfg['BoxPrintingFile']) { $BoxPrintingFile = $cfg['BoxPrintingFile'] }
    else { $BoxPrintingFile = Join-Path $here "Data_boxprintingbin3v7.xlsb" }
}
if ($cfg.ContainsKey('BoxSheet') -and $cfg['BoxSheet']) { $BoxSheet = $cfg['BoxSheet'] }
if (-not $script:HasTarget -and $cfg.ContainsKey('ShiftTarget') -and $cfg['ShiftTarget']) { [int]$ShiftTarget = [int]$cfg['ShiftTarget'] }
if (-not $script:HasRecentMin -and $cfg.ContainsKey('RecentMinutes') -and $cfg['RecentMinutes']) { [int]$RecentMinutes = [int]$cfg['RecentMinutes'] }
if ($RecentMinutes -lt 1) { $RecentMinutes = 30 }
if (-not $script:HasStopMin -and $cfg.ContainsKey('StopMinutes') -and $cfg['StopMinutes']) { [double]$StopMinutes = [double]$cfg['StopMinutes'] }
if ($StopMinutes -le 0) { $StopMinutes = 2 }
$script:BoxFolder = if ($BoxPrintingFile) { Split-Path -Parent $BoxPrintingFile } else { $here }
if ([string]::IsNullOrWhiteSpace($script:BoxFolder)) { $script:BoxFolder = $here }
# weekplan: -PlanFile > config PlanFile > nieuwste 'daily shift NDwk*.xls*' in PlanFolder (config) / scriptmap
if (-not $script:HasPlanFile -and $cfg.ContainsKey('PlanFile') -and $cfg['PlanFile']) {
    $PlanFile = $cfg['PlanFile']; $script:HasPlanFile = $true
}
$script:PlanFolder = if ($cfg.ContainsKey('PlanFolder') -and $cfg['PlanFolder']) { $cfg['PlanFolder'] } else { $here }
if ([string]::IsNullOrWhiteSpace($script:PlanFolder)) { $script:PlanFolder = $here }

# ============================ UITVOEREN ============================
if ($Web) {
    Start-WebServer $Port
}
elseif ($Once) {
    Render-Console (Get-BoxData)
    Write-Host ""
}
else {
    Write-Host ("MONITOR actief - ververst elke {0}s (blad {1})." -f $IntervalSeconds, $BoxSheet) -ForegroundColor DarkCyan
    Write-Host  "  (Ctrl+C om te stoppen)" -ForegroundColor DarkCyan
    while ($true) {
        try { Clear-Host } catch {}
        Write-Host ("[{0}] herladen..." -f (Get-Date -Format 'HH:mm:ss')) -ForegroundColor DarkCyan
        Render-Console (Get-BoxData)
        Write-Host ""
        Write-Host ("Volgende ververs over {0}s. (Ctrl+C om te stoppen)" -f $IntervalSeconds) -ForegroundColor DarkGray
        Start-Sleep -Seconds $IntervalSeconds
    }
}
