#requires -version 5
<#
    Test-BoxCount-Lines.ps1  -  EEN dashboard voor LIJN 9 EN LIJN 11.

    Dit bestand vervangt het vroegere paar Test-BoxCount.ps1 (lijn 9) en
    Test-BoxCount-L11.ps1 (lijn 11). Een server, een pagina, en linksboven knoppen om
    tussen de lijnen te wisselen (?line=9 / ?line=11, net als ?lang=nl|fr|en|ru).

    WAT GEDEELD IS: het inlezen van de .xlsb-cache, het weekplan (blad 'daily shift dpp'),
    de historie uit de *RCDB-bladen, de vertalingen, de opmaak en de webserver.

    WAT PER LIJN VERSCHILT staat in $script:LineDefs (het lijnregister, zie verderop):
      * LIJN 9  - EEN verpakkingsmachine, dus EEN smaak tegelijk. Hoofdproduct = de smaak
        van de laatste doos; prognose = extrapolatie van de lopende run. Een doosblad
        (Boxruw9) en een historieblad (L9RCDB). Motor: Get-BoxData9 / weergave Render-Html9.
      * LIJN 11 - ACHT verpakkingsmachines naast elkaar (etiket-tag 8 = 1101..1108), elk met
        een eigen smaak, tempo en stilstand. 'Laatste doos' zegt hier niets, dus alles wordt
        PER MACHINE gerekend en de ploegprognose is de SOM van de machines. TWEE doosbladen
        (Boxruw111 + Boxruw112) en TWEE historiebladen (L11P1RCDB + L11P2RCDB); die worden
        OPGETELD, precies zoals SAPSTATus het doet (de som klopt 1:1 met blad 'W <week>').
        Motor: Get-BoxData11 / weergave Render-Html11.

    CONFIGURATIE: EEN bestand voor alle lijnen - config.txt naast dit script. Bovenaan het
    algemene blok (bronbestand, planmap, standaardinstellingen); daaronder mag per lijn een blok
    '[5]' ... '[11]' staan dat alleen zijn eigen lijn overschrijft. Wat op de opdrachtregel staat
    wint altijd en geldt voor ALLE lijnen (ze lezen immers hetzelfde bronbestand).

    CACHE: per lijn apart en pas bij het eerste bezoek gevuld - een lijn die niemand opent
    kost dus ook geen Excel-lees. Wijzigt het bronbestand (of begint een nieuwe ploeg), dan
    vervalt de cache van BEIDE lijnen.

    STRIKTE REGELS (ongewijzigd overgenomen): alles wordt ALLEEN-LEZEN geopend, de cache
    wordt NIET herberekend (EnableEvents=False dooft Workbook_Open dat anders xlAutomatic
    forceert), er blijven geen Excel-processen hangen, en er wordt niets gedownload.

    Parameters: -Line 9|11          (lijn waarmee de pagina OPENT; wisselen kan in de pagina)
                -Port 8771 -NoBrowser -IntervalSeconds N
                -Now "2026-09-08 12:46"  (testen met de snapshot-datum)
                -BoxPrintingFile "..."   -BoxSheet "..."   -ShiftTarget 1234
                -RecentMinutes 30        (venster voor het 'recent tempo')
                -StopMinutes 2           (gat dat als stilstand telt)
                -PlanFile "..."          -PlanFolder "..."   -NoPlan
                -RcdbSheet "..."         -HistoryDays 0 (week) / N dagen / -1 = geen historie
    LET OP: -BoxSheet / -RcdbSheet / -ShiftTarget gelden dan voor BEIDE lijnen; laat ze weg
    als je gewoon het lijnregister en de configuratiebestanden wilt laten beslissen.
#>
param(
    [string]  $BoxPrintingFile,
    [string]  $BoxSheet,
    [datetime]$Now,
    [int]     $Port = 8771,
    [switch]  $NoBrowser,
    [int]     $IntervalSeconds = 60,
    [int]     $ShiftTarget = 1234,
    [int]     $RecentMinutes = 30,
    [double]  $StopMinutes = 2,
    [string]  $PlanFile,
    [string]  $PlanFolder,
    [switch]  $NoPlan,
    [string]  $RcdbSheet,
    [string]  $Line = '9',          # lijn waarmee de pagina opent ('9' of '11')
    [int]     $HistoryDays = 0      # 0 = lopende productieweek (start zondag), >0 = zoveel dagen, <0 = uit
)

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$nl = [System.Globalization.CultureInfo]::GetCultureInfo('nl-BE')

# ============================ CONFIG ============================
$here = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ConfigFile = Join-Path $here "config.txt"       # EEN configuratiebestand voor ALLE lijnen
$script:HasNow    = $PSBoundParameters.ContainsKey('Now')
$script:HasBox    = -not [string]::IsNullOrWhiteSpace($BoxPrintingFile)
$script:HasBoxSheet  = -not [string]::IsNullOrWhiteSpace($BoxSheet)
$script:HasTarget = $PSBoundParameters.ContainsKey('ShiftTarget')
$script:HasRecentMin = $PSBoundParameters.ContainsKey('RecentMinutes')
$script:HasStopMin   = $PSBoundParameters.ContainsKey('StopMinutes')
$script:HasPlanFile  = -not [string]::IsNullOrWhiteSpace($PlanFile)
$script:HasPlanDir   = -not [string]::IsNullOrWhiteSpace($PlanFolder)
$script:HasRcdb      = -not [string]::IsNullOrWhiteSpace($RcdbSheet)
$script:HasHistDays  = $PSBoundParameters.ContainsKey('HistoryDays')
# (het standaardblad per lijn staat in $script:LineDefs, zie het lijnregister)

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
# duur (in minuten) -> leesbaar "3 min 56 sec" / "45 sec" / "3 min" (eenheden taalafhankelijk)
function Format-Dur([double]$minVal) {
    $totalSec = [int][Math]::Round($minVal * 60); if ($totalSec -lt 0) { $totalSec = 0 }
    $m = [int][Math]::Floor($totalSec / 60); $s = $totalSec % 60
    if ($m -le 0) { return ('{0} {1}' -f $s, (T 'svg_sec')) }
    if ($s -eq 0) { return ('{0} {1}' -f $m, (T 'svg_min')) }
    return ('{0} {1} {2} {3}' -f $m, (T 'svg_min'), $s, (T 'svg_sec'))
}
function HtmlEnc([string]$s) {
    if ($null -eq $s) { return "" }
    # OOK de rechte apostrof: tooltips staan in title='...' met ENKELE aanhalingstekens, en het
    # Franse blok zit er vol mee (l'objectif, a l'arret, d'equipe). Zonder deze regel brak zo'n
    # tekst het attribuut en verdween de halve tooltip uit de HTML.
    return ($s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&#39;')
}
# ============================ I18N (NL/FR/EN/RU) ============================
$script:Langs = @('nl','fr','en','ru')
$script:DefaultLang = 'nl'
$script:Lang = 'nl'
$script:CultMap = @{ 'nl' = 'nl-BE'; 'fr' = 'fr-FR'; 'en' = 'en-GB'; 'ru' = 'ru-RU' }
$script:I18N = @{
'nl' = @{
  'interval' = "interval"
  'shift_word' = "ploeg"
  'foot_loaded' = "Laatst geladen: {0} &middot; ververst elke {1}s"
  'err_prefix' = "FOUT"
  'card_made' = "Geproduceerd (ploeg)"
  'card_made_sub' = "opgeslagen {0}"
  'card_made_pct' = "{0}&nbsp;% van het plan"
  'card_target' = "Target ploeg"
  'card_target_sub' = "{0}&nbsp;% &middot; tempo {1} dozen/min"
  'card_expected_end' = "Verwachte eindstand"
  'card_expected_end_sub' = "{0}&nbsp;% van target &middot; {1}"
  'card_tempo_now' = "Tempo nu"
  'card_tempo_now_sub' = "dozen/min &middot; {0} per uur"
  'sec_per_minute' = "Dozen per minuut"
  'card_last_win' = "Laatste {0} min"
  'card_eta' = "Target bereikt om"
  'card_eta_sub' = "bij huidig tempo"
  'eta_done' = "target al gehaald"
  'eta_over' = "ploeg voorbij"
  'eta_impossible' = "niet haalbaar (tempo 0)"
  'eta_after' = " (na ploegeinde)"
  'stops_word' = "Stilstand"
  'stops_bron' = "(gat &gt; {0} min telt als stilstand)"
  'card_runtime' = "Draaitijd"
  'card_net' = "Netto tempo"
  'sec_behind' = "Achterstand t.o.v. target-tempo"
  'th_duration' = "Duur"
  'th_kind' = "Soort"
  'kind_startup' = "opstart ploeg"
  'kind_nowstill' = "staat nu stil"
  'kind_stop' = "stop"
  'kind_longest' = "langste"
  'kind_nowmark' = "nu"
  'kind_notstarted' = "nog niet gestart"
  'kind_weekdone' = "weekplan af"
  'weekdone_note' = "{0}: het weekplan was bij ploegstart al rond ({1} van {2}). Die smaak hoeft deze ploeg niet meer gedraaid te worden; de {3} dozen die er in het dagplan nog voor staan, worden in een andere smaak gemaakt. De ploegtarget van {4} dozen blijft ongewijzigd."
  'tt_tempo' = "{0} dozen in {1} min"
  'per_hour_short' = "/u"
  'sec_all_products' = "Producten &mdash; week en ploeg per regel"
  'grp_week' = "Deze week"
  'grp_shift' = "Deze ploeg"
  'th_plan' = "Plan"
  'th_prognose' = "Prognose ploeg"
  'th_producttype' = "Producttype"
  'th_product' = "Product"
  'th_plan_shift' = "Plan ploeg"
  'th_boxes_now' = "Dozen nu"
  'th_expected_end' = "Verwachte eindstand"
  'th_progress' = "Uitvoering"
  'sec_history' = "Historie"
  'hist_bron' = "{0} &ndash; {1} &middot; blad {2}"
  'hist_week' = "week {0}"
  'hist_more' = "Nog 7 dagen tonen"
  'th_day' = "Dag"
  'th_shift' = "Ploeg"
  'th_boxes' = "Dozen"
  'th_products' = "Producten"
  'week_bron' = "week {0} &middot; {1} &ndash; {2} &middot; dagplan blijft leidend"
  'th_made_week' = "Gemaakt"
  'th_rest_week' = "Restant"
  'week_note' = "Plan uit het weekplan &middot; lopende ploeg uit {0} &middot; eerdere dagen uit {1} (ontbreken daar uren, dan staat 'gemaakt' te laag)."
  'warn_weekplan' = "weekplan per smaak niet berekend: {0}"
  'warn_no_rcdb' = "blad {0} niet gevonden - geen historie."
  'warn_hist_read' = "historie niet gelezen: {0}"
  'total' = "Totaal"
  'last_box' = "Laatste doos: {0} &middot; geparseerde rijen: {1}{2}"
  'last_box_txt' = "{0} - {1} (doos {2})"
  'skip_txt' = " &middot; rij 11 = startwaarde ophaalvenster (niet geteld)"
  'skip_blank' = " &middot; {0} rij(en) zonder etiket overgeslagen"
  'empty_state' = "Geen dozen in deze ploeg-interval. Geparseerde rijen: {0}. Target ploeg {1} &middot; {2}."
  'ts_plan' = "plan {0} (week {1}), {2} ploeg {3}"
  'ts_planfallback' = "plan {0} &mdash; nog niets gemaakt, verwacht {1}"
  'ts_param' = "parameter -ShiftTarget"
  'ts_noplan' = "config (-NoPlan)"
  'ts_unknown' = "parameter/config"
  'warn_planfile_missing' = "opgegeven planbestand niet gevonden: {0}"
  'warn_no_planfile' = "geen planbestand (daily shift NDwk*.xls) in {0} - target uit config."
  'warn_plan_read' = "weekplan niet gelezen: {0}"
  'warn_day_not_in_plan' = "productiedag {0} (ploeg {1}) staat niet in {2} (week {3}, {4}) - target uit config."
  'warn_no_plan_dayshift' = "geen plan voor deze dag/ploeg in {0} - target uit config."
  'warn_no_data_row11' = "Geen data vanaf rij 11 in {0}."
  'svg_target_line' = "target {0}/ploeg = {1}/min"
  'svg_tempo_line' = "tempo {0}/min"
  'svg_clip_note' = "piek {0}/min - buiten schaal: {1}"
  'svg_line' = "lijn"
  'svg_min' = "min"
  'svg_on_target' = "op target-tempo"
  'svg_shift_end_word' = "einde ploeg"
  'stops_show_all' = "Alle {0} stops tonen"
  'stops_collapse' = "Inklappen"
  # --- lijn 11: machines (tag 8) ---
  'sec_machines' = "Machines"
  'mach_bron' = "({0} machines &middot; {1} draaien, {2} staan stil &middot; machine = tag 8 op het etiket)"
  'th_machine' = "Machine"
  'th_flavour' = "Smaak"
  'th_tempo' = "Tempo"
  'th_runtime_col' = "Draaitijd"
  'th_stop_col' = "Stilstand"
  'th_status' = "Status"
  'st_running' = "draait"
  'st_still' = "stil {0} min"
  'weak_run_l11' = "Pas {0} dozen in {1} min &mdash; prognose is nog voorlopig."
  'tt_tempo_mach' = "{0} dozen &middot; {1} machine(s): {2}"
  'card_stops_sub_l11' = "machine-min in {0} stops &middot; nu stil: {1} van {2}"
  'card_runtime_sub_l11' = "% &middot; {0} van {1} machine-min"
  'card_net_sub_l11' = "dozen/min terwijl de lijn produceert"
  'card_behind' = "Voor/achter op target"
  'card_behind_sub' = "dozen t.o.v. een vlak targettempo"
  'st_lag' = "data loopt {0} min achter"
  'tt_machine' = "eerste doos {0} - laatste {1} - netto {2}/min - langste stop {3}"
  'nowstill_l11' = "De HELE lijn staat stil: al {0} min geen enkele doos (laatste {1})."
  'stops_head_line' = "Hele lijn stil &mdash; alle {0} onderbrekingen, chronologisch"
  'warn_no_boxsheet' = "blad {0} niet gevonden in {1} - dat deel van de lijn ontbreekt."
  'svg_sec' = "sec"
  'svg_pct_run' = "draait"
  'svg_pct_stop' = "stil"
  'tt_strip_pct' = "{0}: draait {1} min, stil {2} min ({3} - {4})"
  'tt_next_at' = "volgens het weekrooster gepland vanaf {0}"
  'nxt_from' = "vanaf {0}"
  'tt_next_from' = "loopt volgens het weekrooster al vanaf {0}; het ploegplan is nog niet gehaald"
  'order_note' = "Volgorde: eerst wat nu draait, dan wat volgens het weekrooster nog komt (op gepland begin), dan wat deze week nog niet gemaakt is, en onderaan wat deze week al gemaakt is."

  # ---- sleutels die alleen lijn 9 gebruikt + de per-lijn titel ----
  'h1_l9' = "Dozen per producttype &mdash; huidige ploeg"
  'line_word' = "Lijn {0}"
  'h1_machines' = "Lijn {0} &mdash; dozen per smaak en machine, huidige ploeg"
  'sec_forecast' = "Prognose einde ploeg &mdash; hoofdproduct {0} "
  'forecast_bron' = "(laatste rij &middot; data t/m {0})"
  'weak_run' = "Korte run ({0} dozen in {1} min) &mdash; prognose is nog voorlopig."
  'wdesc_run' = "Run van {0}: {1} dozen sinds {2} ({3} min) &middot; nog {4} min tot ploegeinde &middot; verwacht voor dit product {5} dozen"
  'nowstill' = "Lijn staat nu stil: al {0} min geen doos (laatste {1})."
  'card_stops_sub' = "min in {0} stops &middot; langste {1} {2}"
  'card_runtime_sub' = "% &middot; {0} van {1} min"
  'card_net_sub' = "dozen/min als de lijn draait"
  'card_loss' = "Verlies stilstand"
  'card_loss_sub' = "dozen bij netto tempo"
  'stops_head_all' = "Stilstand &mdash; alle {0} stops, chronologisch"
  'sec_products' = "Dozen per producttype"
  'plan_bron' = "(plan: {0})"
  'th_todo' = "Nog te doen"
  'sec_weekplan' = "Weekplan per smaak"
  'th_plan_week' = "Plan week"
  'svg_target_short' = "target {0}"
  'svg_forecast' = "prognose {0}"
}
'fr' = @{
  'interval' = "intervalle"
  'shift_word' = "équipe"
  'foot_loaded' = "Dernier chargement : {0} &middot; actualisé toutes les {1}s"
  'err_prefix' = "ERREUR"
  'card_made' = "Produit (équipe)"
  'card_made_sub' = "enregistré {0}"
  'card_made_pct' = "{0}&nbsp;% du plan"
  'card_target' = "Objectif équipe"
  'card_target_sub' = "{0}&nbsp;% &middot; cadence {1} boîtes/min"
  'card_expected_end' = "Résultat final prévu"
  'card_expected_end_sub' = "{0}&nbsp;% de l'objectif &middot; {1}"
  'card_tempo_now' = "Cadence actuelle"
  'card_tempo_now_sub' = "boîtes/min &middot; {0} par heure"
  'sec_per_minute' = "Boîtes par minute"
  'card_last_win' = "Dernières {0} min"
  'card_eta' = "Objectif atteint à"
  'card_eta_sub' = "à la cadence actuelle"
  'eta_done' = "objectif déjà atteint"
  'eta_over' = "équipe terminée"
  'eta_impossible' = "impossible (cadence 0)"
  'eta_after' = " (après la fin d'équipe)"
  'stops_word' = "Arrêts"
  'stops_bron' = "(écart &gt; {0} min compte comme arrêt)"
  'card_runtime' = "Temps de marche"
  'card_net' = "Cadence nette"
  'sec_behind' = "Retard p/r à la cadence cible"
  'th_duration' = "Durée"
  'th_kind' = "Type"
  'kind_startup' = "démarrage équipe"
  'kind_nowstill' = "à l'arrêt"
  'kind_stop' = "arrêt"
  'kind_longest' = "le plus long"
  'kind_nowmark' = "en cours"
  'kind_notstarted' = "pas encore démarré"
  'kind_weekdone' = "plan hebdo fait"
  'weekdone_note' = "{0} : le plan hebdo était déjà atteint au début de l’équipe ({1} sur {2}). Cette saveur ne doit plus tourner ; les {3} boîtes encore prévues au plan du jour seront faites dans une autre saveur. L’objectif d’équipe de {4} boîtes reste inchangé."
  'tt_tempo' = "{0} boîtes en {1} min"
  'per_hour_short' = "/h"
  'sec_all_products' = "Produits &mdash; semaine et équipe par ligne"
  'grp_week' = "Cette semaine"
  'grp_shift' = "Cette équipe"
  'th_plan' = "Plan"
  'th_prognose' = "Prévision équipe"
  'th_producttype' = "Type de produit"
  'th_product' = "Produit"
  'th_plan_shift' = "Plan équipe"
  'th_boxes_now' = "Boîtes"
  'th_expected_end' = "Résultat final prévu"
  'th_progress' = "Avancement"
  'sec_history' = "Historique"
  'hist_bron' = "{0} &ndash; {1} &middot; feuille {2}"
  'hist_week' = "semaine {0}"
  'hist_more' = "Afficher 7 jours de plus"
  'th_day' = "Jour"
  'th_shift' = "&Eacute;quipe"
  'th_boxes' = "Bo&icirc;tes"
  'th_products' = "Produits"
  'week_bron' = "semaine {0} &middot; {1} &ndash; {2} &middot; le plan du jour reste prioritaire"
  'th_made_week' = "Réalisé"
  'th_rest_week' = "Restant"
  'week_note' = "Plan issu du plan hebdo &middot; équipe en cours depuis {0} &middot; jours précédents depuis {1} (heures manquantes → 'réalisé' sous-évalué)."
  'warn_weekplan' = "plan hebdo par saveur non calculé : {0}"
  'warn_no_rcdb' = "feuille {0} introuvable &mdash; pas d'historique."
  'warn_hist_read' = "historique non lu : {0}"
  'total' = "Total"
  'last_box' = "Dernière boîte : {0} &middot; lignes analysées : {1}{2}"
  'last_box_txt' = "{0} - {1} (boîte {2})"
  'skip_txt' = " &middot; ligne 11 = valeur initiale de la fenêtre (non comptée)"
  'skip_blank' = " &middot; {0} ligne(s) sans étiquette ignorée(s)"
  'empty_state' = "Aucune boîte dans cet intervalle d'équipe. Lignes analysées : {0}. Objectif équipe {1} &middot; {2}."
  'ts_plan' = "plan {0} (semaine {1}), {2} équipe {3}"
  'ts_planfallback' = "plan {0} &mdash; rien encore produit, attendu {1}"
  'ts_param' = "paramètre -ShiftTarget"
  'ts_noplan' = "config (-NoPlan)"
  'ts_unknown' = "paramètre/config"
  'warn_planfile_missing' = "fichier de plan indiqué introuvable : {0}"
  'warn_no_planfile' = "aucun fichier de plan (daily shift NDwk*.xls) dans {0} &mdash; objectif depuis la config."
  'warn_plan_read' = "plan hebdomadaire non lu : {0}"
  'warn_day_not_in_plan' = "jour de production {0} (équipe {1}) absent de {2} (semaine {3}, {4}) &mdash; objectif depuis la config."
  'warn_no_plan_dayshift' = "aucun plan pour ce jour/équipe dans {0} &mdash; objectif depuis la config."
  'warn_no_data_row11' = "Aucune donnée à partir de la ligne 11 dans {0}."
  'svg_target_line' = "objectif {0}/équipe = {1}/min"
  'svg_tempo_line' = "cadence {0}/min"
  'svg_clip_note' = "pic {0}/min - hors échelle: {1}"
  'svg_line' = "ligne"
  'svg_min' = "min"
  'svg_on_target' = "à la cadence cible"
  'svg_shift_end_word' = "fin équipe"
  'stops_show_all' = "Afficher les {0} arrêts"
  'stops_collapse' = "Réduire"
  # --- ligne 11 : machines (tag 8) ---
  'sec_machines' = "Machines"
  'mach_bron' = "({0} machines &middot; {1} en marche, {2} à l'arrêt &middot; machine = tag 8 de l'étiquette)"
  'th_machine' = "Machine"
  'th_flavour' = "Goût"
  'th_tempo' = "Cadence"
  'th_runtime_col' = "Tps march."
  'th_stop_col' = "Arrêt"
  'th_status' = "État"
  'st_running' = "en marche"
  'st_still' = "arrêt {0} min"
  'weak_run_l11' = "Seulement {0} boîtes en {1} min &mdash; prévision encore provisoire."
  'tt_tempo_mach' = "{0} boîtes &middot; {1} machine(s) : {2}"
  'card_stops_sub_l11' = "machine-min en {0} arrêts &middot; à l'arrêt : {1} sur {2}"
  'card_runtime_sub_l11' = "% &middot; {0} sur {1} machine-min"
  'card_net_sub_l11' = "boîtes/min pendant que la ligne produit"
  'card_behind' = "Avance/retard sur l'objectif"
  'card_behind_sub' = "boîtes par rapport à une cadence objectif régulière"
  'st_lag' = "données en retard de {0} min"
  'tt_machine' = "première boîte {0} - dernière {1} - net {2}/min - arrêt le plus long {3}"
  'nowstill_l11' = "TOUTE la ligne est à l'arrêt : {0} min sans la moindre boîte (dernière {1})."
  'stops_head_line' = "Ligne entière à l'arrêt &mdash; les {0} interruptions, chronologique"
  'warn_no_boxsheet' = "feuille {0} introuvable dans {1} &mdash; cette partie de la ligne manque."
  'svg_sec' = "s"
  'svg_pct_run' = "marche"
  'svg_pct_stop' = "arrêt"
  'tt_strip_pct' = "{0} : en marche {1} min, à l'arrêt {2} min ({3} - {4})"
  'tt_next_at' = "prévu au planning de la semaine à partir de {0}"
  'nxt_from' = "depuis {0}"
  'tt_next_from' = "prévu au planning de la semaine depuis {0} ; le plan de l'équipe n'est pas encore atteint"
  'order_note' = "Ordre : d'abord ce qui tourne maintenant, puis ce qui vient selon le planning de la semaine (par début prévu), puis ce qui n'a pas encore été produit cette semaine, et en bas ce qui est déjà fait cette semaine."

  # ---- sleutels die alleen lijn 9 gebruikt + de per-lijn titel ----
  'h1_l9' = "Boîtes par type de produit &mdash; équipe en cours"
  'line_word' = "Ligne {0}"
  'h1_machines' = "Ligne {0} &mdash; boîtes par goût et machine, équipe en cours"
  'sec_forecast' = "Prévision fin d'équipe &mdash; produit principal {0} "
  'forecast_bron' = "(dernière ligne &middot; données jusqu'à {0})"
  'weak_run' = "Série courte ({0} boîtes en {1} min) &mdash; prévision encore provisoire."
  'wdesc_run' = "Série de {0} : {1} boîtes depuis {2} ({3} min) &middot; encore {4} min avant la fin d'équipe &middot; prévu pour ce produit {5} boîtes"
  'nowstill' = "La ligne est à l'arrêt : déjà {0} min sans boîte (dernière {1})."
  'card_stops_sub' = "min en {0} arrêts &middot; le plus long {1} {2}"
  'card_runtime_sub' = "% &middot; {0} sur {1} min"
  'card_net_sub' = "boîtes/min quand la ligne tourne"
  'card_loss' = "Perte (arrêts)"
  'card_loss_sub' = "boîtes à cadence nette"
  'stops_head_all' = "Arrêts &mdash; les {0} arrêts, chronologique"
  'sec_products' = "Boîtes par type de produit"
  'plan_bron' = "(plan : {0})"
  'th_todo' = "Reste à faire"
  'sec_weekplan' = "Plan hebdo par saveur"
  'th_plan_week' = "Plan semaine"
  'svg_target_short' = "objectif {0}"
  'svg_forecast' = "prévision {0}"
}
'en' = @{
  'interval' = "interval"
  'shift_word' = "shift"
  'foot_loaded' = "Last loaded: {0} &middot; refreshes every {1}s"
  'err_prefix' = "ERROR"
  'card_made' = "Produced (shift)"
  'card_made_sub' = "saved {0}"
  'card_made_pct' = "{0}&nbsp;% of plan"
  'card_target' = "Target (shift)"
  'card_target_sub' = "{0}&nbsp;% &middot; rate {1} boxes/min"
  'card_expected_end' = "Expected final total"
  'card_expected_end_sub' = "{0}&nbsp;% of target &middot; {1}"
  'card_tempo_now' = "Current rate"
  'card_tempo_now_sub' = "boxes/min &middot; {0} per hour"
  'sec_per_minute' = "Boxes per minute"
  'card_last_win' = "Last {0} min"
  'card_eta' = "Target reached at"
  'card_eta_sub' = "at current rate"
  'eta_done' = "target already reached"
  'eta_over' = "shift over"
  'eta_impossible' = "not reachable (rate 0)"
  'eta_after' = " (after shift end)"
  'stops_word' = "Downtime"
  'stops_bron' = "(gap &gt; {0} min counts as downtime)"
  'card_runtime' = "Running time"
  'card_net' = "Net rate"
  'sec_behind' = "Behind vs target rate"
  'th_duration' = "Duration"
  'th_kind' = "Kind"
  'kind_startup' = "shift start-up"
  'kind_nowstill' = "stopped now"
  'kind_stop' = "stop"
  'kind_longest' = "longest"
  'kind_nowmark' = "now"
  'kind_notstarted' = "not started yet"
  'kind_weekdone' = "week plan done"
  'weekdone_note' = "{0}: the week plan was already met at the start of the shift ({1} of {2}). That flavour no longer needs to run; the {3} boxes still on the day plan for it will be made in another flavour. The shift target of {4} boxes stays unchanged."
  'tt_tempo' = "{0} boxes in {1} min"
  'per_hour_short' = "/h"
  'sec_all_products' = "Products &mdash; week and shift on one row"
  'grp_week' = "This week"
  'grp_shift' = "This shift"
  'th_plan' = "Plan"
  'th_prognose' = "Shift forecast"
  'th_producttype' = "Product type"
  'th_product' = "Product"
  'th_plan_shift' = "Shift plan"
  'th_boxes_now' = "Boxes now"
  'th_expected_end' = "Expected final total"
  'th_progress' = "Progress"
  'sec_history' = "History"
  'hist_bron' = "{0} &ndash; {1} &middot; sheet {2}"
  'hist_week' = "week {0}"
  'hist_more' = "Show 7 more days"
  'th_day' = "Day"
  'th_shift' = "Shift"
  'th_boxes' = "Boxes"
  'th_products' = "Products"
  'week_bron' = "week {0} &middot; {1} &ndash; {2} &middot; the daily plan stays leading"
  'th_made_week' = "Made"
  'th_rest_week' = "Remaining"
  'week_note' = "Plan from the week plan &middot; current shift from {0} &middot; earlier days from {1} (missing hours there make 'made' too low)."
  'warn_weekplan' = "week plan per flavour not calculated: {0}"
  'warn_no_rcdb' = "sheet {0} not found &mdash; no history."
  'warn_hist_read' = "history not read: {0}"
  'total' = "Total"
  'last_box' = "Last box: {0} &middot; parsed rows: {1}{2}"
  'last_box_txt' = "{0} - {1} (box {2})"
  'skip_txt' = " &middot; row 11 = window start value (not counted)"
  'skip_blank' = " &middot; {0} row(s) without a label skipped"
  'empty_state' = "No boxes in this shift interval. Parsed rows: {0}. Target shift {1} &middot; {2}."
  'ts_plan' = "plan {0} (week {1}), {2} shift {3}"
  'ts_planfallback' = "plan {0} &mdash; nothing produced yet, expecting {1}"
  'ts_param' = "parameter -ShiftTarget"
  'ts_noplan' = "config (-NoPlan)"
  'ts_unknown' = "parameter/config"
  'warn_planfile_missing' = "specified plan file not found: {0}"
  'warn_no_planfile' = "no plan file (daily shift NDwk*.xls) in {0} &mdash; target from config."
  'warn_plan_read' = "weekly plan not read: {0}"
  'warn_day_not_in_plan' = "production day {0} (shift {1}) not in {2} (week {3}, {4}) &mdash; target from config."
  'warn_no_plan_dayshift' = "no plan for this day/shift in {0} &mdash; target from config."
  'warn_no_data_row11' = "No data from row 11 in {0}."
  'svg_target_line' = "target {0}/shift = {1}/min"
  'svg_tempo_line' = "rate {0}/min"
  'svg_clip_note' = "peak {0}/min - off scale: {1}"
  'svg_line' = "line"
  'svg_min' = "min"
  'svg_on_target' = "on target rate"
  'svg_shift_end_word' = "shift end"
  'stops_show_all' = "Show all {0} stops"
  'stops_collapse' = "Collapse"
  # --- line 11: machines (tag 8) ---
  'sec_machines' = "Machines"
  'mach_bron' = "({0} machines &middot; {1} running, {2} stopped &middot; machine = tag 8 on the label)"
  'th_machine' = "Machine"
  'th_flavour' = "Flavour"
  'th_tempo' = "Rate"
  'th_runtime_col' = "Uptime"
  'th_stop_col' = "Downtime"
  'th_status' = "Status"
  'st_running' = "running"
  'st_still' = "stopped {0} min"
  'weak_run_l11' = "Only {0} boxes in {1} min &mdash; forecast still provisional."
  'tt_tempo_mach' = "{0} boxes &middot; {1} machine(s): {2}"
  'card_stops_sub_l11' = "machine-min in {0} stops &middot; stopped now: {1} of {2}"
  'card_runtime_sub_l11' = "% &middot; {0} of {1} machine-min"
  'card_net_sub_l11' = "boxes/min while the line is producing"
  'card_behind' = "Ahead/behind target"
  'card_behind_sub' = "boxes versus a flat target pace"
  'st_lag' = "data is {0} min behind"
  'tt_machine' = "first box {0} - last {1} - net {2}/min - longest stop {3}"
  'nowstill_l11' = "The WHOLE line is stopped: {0} min without a single box (last {1})."
  'stops_head_line' = "Whole line stopped &mdash; all {0} interruptions, chronological"
  'warn_no_boxsheet' = "sheet {0} not found in {1} &mdash; that part of the line is missing."
  'svg_sec' = "sec"
  'svg_pct_run' = "running"
  'svg_pct_stop' = "stopped"
  'tt_strip_pct' = "{0}: running {1} min, stopped {2} min ({3} - {4})"
  'tt_next_at' = "planned in the week schedule from {0}"
  'nxt_from' = "from {0}"
  'tt_next_from' = "scheduled in the week plan since {0}; the shift plan is not reached yet"
  'order_note' = "Order: running now first, then what comes next in the week schedule (by planned start), then what has not been made this week, and at the bottom what has already been made this week."

  # ---- sleutels die alleen lijn 9 gebruikt + de per-lijn titel ----
  'h1_l9' = "Boxes per product type &mdash; current shift"
  'line_word' = "Line {0}"
  'h1_machines' = "Line {0} &mdash; boxes per flavour and machine, current shift"
  'sec_forecast' = "End-of-shift forecast &mdash; main product {0} "
  'forecast_bron' = "(last row &middot; data through {0})"
  'weak_run' = "Short run ({0} boxes in {1} min) &mdash; forecast still provisional."
  'wdesc_run' = "Run of {0}: {1} boxes since {2} ({3} min) &middot; {4} min to end of shift &middot; expected for this product {5} boxes"
  'nowstill' = "Line is stopped now: {0} min without a box (last {1})."
  'card_stops_sub' = "min in {0} stops &middot; longest {1} {2}"
  'card_runtime_sub' = "% &middot; {0} of {1} min"
  'card_net_sub' = "boxes/min while the line runs"
  'card_loss' = "Downtime loss"
  'card_loss_sub' = "boxes at net rate"
  'stops_head_all' = "Downtime &mdash; all {0} stops, chronological"
  'sec_products' = "Boxes per product type"
  'plan_bron' = "(plan: {0})"
  'th_todo' = "To do"
  'sec_weekplan' = "Week plan per flavour"
  'th_plan_week' = "Plan week"
  'svg_target_short' = "target {0}"
  'svg_forecast' = "forecast {0}"
}
'ru' = @{
  'interval' = "интервал"
  'shift_word' = "смена"
  'foot_loaded' = "Обновлено: {0} &middot; автообновление каждые {1} с"
  'err_prefix' = "ОШИБКА"
  'card_made' = "Произведено (смена)"
  'card_made_sub' = "сохранено {0}"
  'card_made_pct' = "{0}&nbsp;% от плана"
  'card_target' = "Цель (смена)"
  'card_target_sub' = "{0}&nbsp;% &middot; темп {1} коробок/мин"
  'card_expected_end' = "Ожидаемый итог"
  'card_expected_end_sub' = "{0}&nbsp;% от цели &middot; {1}"
  'card_tempo_now' = "Текущий темп"
  'card_tempo_now_sub' = "коробок/мин &middot; {0} в час"
  'sec_per_minute' = "Коробок в минуту"
  'card_last_win' = "Последние {0} мин"
  'card_eta' = "Цель достигнута в"
  'card_eta_sub' = "при текущем темпе"
  'eta_done' = "цель уже достигнута"
  'eta_over' = "смена окончена"
  'eta_impossible' = "недостижимо (темп 0)"
  'eta_after' = " (после конца смены)"
  'stops_word' = "Простой"
  'stops_bron' = "(разрыв &gt; {0} мин считается простоем)"
  'card_runtime' = "Время работы"
  'card_net' = "Чистый темп"
  'sec_behind' = "Отставание от целевого темпа"
  'th_duration' = "Длит."
  'th_kind' = "Тип"
  'kind_startup' = "запуск смены"
  'kind_nowstill' = "сейчас стоит"
  'kind_stop' = "остановка"
  'kind_longest' = "дольше всего"
  'kind_nowmark' = "сейчас"
  'kind_notstarted' = "ещё не начали"
  'kind_weekdone' = "план недели выполнен"
  'weekdone_note' = "{0}: недельный план был выполнен уже к началу смены ({1} из {2}). Этот вкус больше делать не нужно; остаток дневного плана по нему ({3} кор.) будет сделан другим вкусом. Цель смены — {4} кор. — не меняется."
  'tt_tempo' = "{0} коробок за {1} мин"
  'per_hour_short' = "/ч"
  'sec_all_products' = "Продукты &mdash; неделя и смена в одной строке"
  'grp_week' = "За неделю"
  'grp_shift' = "Текущая смена"
  'th_plan' = "План"
  'th_prognose' = "Прогноз смены"
  'th_producttype' = "Тип продукта"
  'th_product' = "Продукт"
  'th_plan_shift' = "План смены"
  'th_boxes_now' = "Коробок сейчас"
  'th_expected_end' = "Ожидаемый итог"
  'th_progress' = "Выполнение"
  'sec_history' = "История"
  'hist_bron' = "{0} &ndash; {1} &middot; лист {2}"
  'hist_week' = "неделя {0}"
  'hist_more' = "Показать ещё 7 дней"
  'th_day' = "День"
  'th_shift' = "Смена"
  'th_boxes' = "Коробок"
  'th_products' = "Продукты"
  'week_bron' = "неделя {0} &middot; {1} &ndash; {2} &middot; главный план — дневной"
  'th_made_week' = "Сделано"
  'th_rest_week' = "Осталось"
  'week_note' = "План — из недельного плана &middot; текущая смена — из {0} &middot; прошлые дни — из {1} (если там нет часов, «сделано» занижено)."
  'warn_weekplan' = "план недели по вкусам не посчитан: {0}"
  'warn_no_rcdb' = "лист {0} не найден &mdash; истории нет."
  'warn_hist_read' = "история не прочитана: {0}"
  'total' = "Итого"
  'last_box' = "Последняя коробка: {0} &middot; разобрано строк: {1}{2}"
  'last_box_txt' = "{0} - {1} (коробка {2})"
  'skip_txt' = " &middot; строка 11 = стартовое значение окна (не учтена)"
  'skip_blank' = " &middot; строк без этикетки пропущено: {0}"
  'empty_state' = "Нет коробок в этом интервале смены. Разобрано строк: {0}. Цель смены {1} &middot; {2}."
  'ts_plan' = "план {0} (неделя {1}), {2} смена {3}"
  'ts_planfallback' = "план {0} &mdash; ещё ничего не произведено, ожидается {1}"
  'ts_param' = "параметр -ShiftTarget"
  'ts_noplan' = "конфиг (-NoPlan)"
  'ts_unknown' = "параметр/конфиг"
  'warn_planfile_missing' = "указанный файл плана не найден: {0}"
  'warn_no_planfile' = "нет файла плана (daily shift NDwk*.xls) в {0} &mdash; цель из конфига."
  'warn_plan_read' = "недельный план не прочитан: {0}"
  'warn_day_not_in_plan' = "производственный день {0} (смена {1}) отсутствует в {2} (неделя {3}, {4}) &mdash; цель из конфига."
  'warn_no_plan_dayshift' = "нет плана на этот день/смену в {0} &mdash; цель из конфига."
  'warn_no_data_row11' = "Нет данных начиная со строки 11 в {0}."
  'svg_target_line' = "цель {0}/смена = {1}/мин"
  'svg_tempo_line' = "темп {0}/мин"
  'svg_clip_note' = "пик {0}/мин - вне шкалы: {1}"
  'svg_line' = "линия"
  'svg_min' = "мин"
  'svg_on_target' = "на целевом темпе"
  'svg_shift_end_word' = "конец смены"
  'stops_show_all' = "Показать все ({0})"
  'stops_collapse' = "Свернуть"
  # --- линия 11: машины (тег 8) ---
  'sec_machines' = "Машины"
  'mach_bron' = "({0} машин &middot; работают {1}, стоят {2} &middot; машина = тег 8 на этикетке)"
  'th_machine' = "Машина"
  'th_flavour' = "Вкус"
  'th_tempo' = "Темп"
  'th_runtime_col' = "В работе"
  'th_stop_col' = "Простой"
  'th_status' = "Статус"
  'st_running' = "работает"
  'st_still' = "стоит {0} мин"
  'weak_run_l11' = "Пока лишь {0} коробок за {1} мин &mdash; прогноз предварительный."
  'tt_tempo_mach' = "{0} коробок &middot; машин: {1} ({2})"
  'card_stops_sub_l11' = "машино-мин в {0} остановках &middot; сейчас стоят: {1} из {2}"
  'card_runtime_sub_l11' = "% &middot; {0} из {1} машино-мин"
  'card_net_sub_l11' = "коробок/мин, пока линия производит"
  'card_behind' = "Опережение/отставание"
  'card_behind_sub' = "коробок относительно ровного целевого темпа"
  'st_lag' = "данные отстают на {0} мин"
  'tt_machine' = "первая коробка {0} - последняя {1} - нетто {2}/мин - самый долгий простой {3}"
  'nowstill_l11' = "ВСЯ линия стоит: уже {0} мин ни одной коробки (последняя {1})."
  'stops_head_line' = "Стояла вся линия &mdash; все {0} перерывов, по времени"
  'warn_no_boxsheet' = "лист {0} не найден в {1} &mdash; эта часть линии отсутствует."
  'svg_sec' = "сек"
  'svg_pct_run' = "работа"
  'svg_pct_stop' = "простой"
  'tt_strip_pct' = "{0}: работала {1} мин, стояла {2} мин ({3} - {4})"
  'tt_next_at' = "по графику недели с {0}"
  'nxt_from' = "с {0}"
  'tt_next_from' = "по графику недели идёт с {0}; план смены ещё не выполнен"
  'order_note' = "Порядок: сначала то, что идёт сейчас, затем то, что дальше по графику недели (по плановому старту), затем то, что на этой неделе ещё не делали, внизу — уже сделанное на этой неделе."

  # ---- sleutels die alleen lijn 9 gebruikt + de per-lijn titel ----
  'h1_l9' = "Коробки по типу продукта &mdash; текущая смена"
  'line_word' = "Линия {0}"
  'h1_machines' = "Линия {0} &mdash; коробки по вкусам и машинам, текущая смена"
  'sec_forecast' = "Прогноз на конец смены &mdash; основной продукт {0} "
  'forecast_bron' = "(последняя строка &middot; данные до {0})"
  'weak_run' = "Короткий прогон ({0} коробок за {1} мин) &mdash; прогноз пока предварительный."
  'wdesc_run' = "Прогон {0}: {1} коробок с {2} ({3} мин) &middot; ещё {4} мин до конца смены &middot; ожидается по этому продукту {5} коробок"
  'nowstill' = "Линия сейчас стоит: уже {0} мин нет коробок (последняя {1})."
  'card_stops_sub' = "мин в {0} остановках &middot; дольше всего {1} {2}"
  'card_runtime_sub' = "% &middot; {0} из {1} мин"
  'card_net_sub' = "коробок/мин когда линия работает"
  'card_loss' = "Потери простоя"
  'card_loss_sub' = "коробок при чистом темпе"
  'stops_head_all' = "Простой &mdash; все {0} остановок, по времени"
  'sec_products' = "Коробки по типу продукта"
  'plan_bron' = "(план: {0})"
  'th_todo' = "Осталось"
  'sec_weekplan' = "План недели по вкусам"
  'th_plan_week' = "План недели"
  'svg_target_short' = "цель {0}"
  'svg_forecast' = "прогноз {0}"
}
}
function T([string]$k) {
    $tab = $script:I18N[$script:Lang]
    if ($tab -and $tab.ContainsKey($k)) { return $tab[$k] }
    $nlt = $script:I18N['nl']
    if ($nlt.ContainsKey($k)) { return $nlt[$k] }
    return $k
}
function Get-ReqLang([string]$req) {
    if ($req -match '[?&]lang=([A-Za-z]{2})') { $l = $matches[1].ToLower(); if ($script:Langs -contains $l) { return $l } }
    return $script:DefaultLang
}
# ===========================================================================

# Waarschuwingen stapelen in plaats van overschrijven. De TEKST komt pas in Render-Html uit de
# taaltabel (sleutel + waarden), zodat elke taal zijn eigen zin krijgt.
function Add-Warn($d, [string]$key, $vals = @()) {
    if ([string]::IsNullOrWhiteSpace($key)) { return }
    $d.WarnList += [pscustomobject]@{ Key = $key; Vals = @($vals) }
}

# Leest HET ENE configuratiebestand (config.txt). Formaat: 'sleutel = waarde', '#' = commentaar,
# en een regel '[6]' opent het blok van EEN lijn. Alles VOOR het eerste blok geldt voor ALLE
# lijnen; een sleutel binnen '[6]' overschrijft dat algemene blok, maar alleen voor lijn 6.
# Een kop zonder cijfer ('[algemeen]', '[общее]') keert terug naar het algemene blok.
# Resultaat: @{ '' = <algemeen>; '6' = <blok lijn 6>; ... } - de lege sleutel is het algemene blok.
function Read-ConfigFile([string]$path) {
    $all = @{ '' = @{} }
    $sec = ''
    if (Test-Path -LiteralPath $path) {
        foreach ($line in (Get-Content -LiteralPath $path)) {
            $t = $line.Trim()
            if ($t -eq '' -or $t.StartsWith('#')) { continue }
            if ($t.StartsWith('[') -and $t.EndsWith(']')) {
                # '[6]', '[lijn 6]', '[линия 6]' -> '6'
                $m = [regex]::Match($t, '\d+')
                $sec = if ($m.Success) { $m.Value } else { '' }
                if (-not $all.ContainsKey($sec)) { $all[$sec] = @{} }
                continue
            }
            $i = $t.IndexOf('=')
            if ($i -gt 0) {
                $k = $t.Substring(0, $i).Trim()
                $v = $t.Substring($i + 1).Trim().Trim('"').Trim("'")
                if ($k) { $all[$sec][$k] = $v }
            }
        }
    }
    return $all
}

# Waarde van EEN sleutel voor EEN lijn: eerst het eigen blok van de lijn, dan het algemene blok.
# Niets gevonden (of leeg) -> $null, zodat de aanroeper zijn eigen standaard kan houden.
function Get-CfgVal($all, [string]$id, [string]$key) {
    foreach ($s in @($id, '')) {
        if ($all.ContainsKey($s) -and $all[$s].ContainsKey($key)) {
            $v = $all[$s][$key]
            if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
        }
    }
    return $null
}

# Producttype = waarde van tag 14 in de etiket-string (tussen |14| en de volgende |)
function Get-ProductType([string]$label) {
    if ([string]::IsNullOrEmpty($label)) { return $null }
    $m = [regex]::Match($label, '\|14\|([^|]*)')
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
    return $null
}
# Machine = waarde van tag 8 in de etiket-string. OP LIJN 11 IS DIT DE KERN: elk van de acht
# verpakkingsmachines (1101..1108) draait zijn eigen smaak, dus zonder deze tag valt niet te
# zien welke smaak stilstaat. Leeg (etiket zonder tags, ~0,1 % van de dozen) -> '(onbekend)'.
function Get-Machine([string]$label) {
    if ([string]::IsNullOrEmpty($label)) { return $null }
    $m = [regex]::Match($label, '\|8\|([^|]*)')
    if ($m.Success) { return $m.Groups[1].Value.Trim() }
    return $null
}
# 'Boxruw111,Boxruw112' -> @('Boxruw111','Boxruw112'). Lijn 11 heeft TWEE doosprinters en dus
# twee bladen; komma/puntkomma/spatie mogen allemaal als scheidingsteken.
function Split-SheetList([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return @() }
    return @($s -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# Doosteller = laatste veld na de laatste | (controltekens verwijderd)
function Get-Counter([string]$label) {
    if ([string]::IsNullOrEmpty($label)) { return '' }
    $clean = $label.Replace([string][char]1, '').Replace([string][char]26, '')
    $parts = $clean -split '\|'
    if ($parts.Length -gt 0) { return ($parts[$parts.Length - 1]).Trim() }
    return ''
}

# HORIZON VAN DE OPHALING (ALLE lijnen samen). De cache wordt met vertraging weggeschreven: in
# twee snapshots liep de laatste doos van een NORMAAL draaiende lijn 4 a 6 minuten achter op het
# bestand. Met -StopMinutes 2 betekende 'laatste doos ouder dan nu - 2 min' dus bij ELKE render
# 'de lijn staat stil'. Die gedeelde vertraging is hier te meten: de JONGSTE doos in het HELE
# bestand (alle Boxruw-bladen) is het verste dat de ophaling gekomen is. Ligt de laatste doos van
# DEZE lijn daar meer dan -StopMinutes achter, dan gaat het niet om de ophaling maar staat de lijn
# echt stil - andere lijnen printten immers wel door. Voorbeeld 18/09 20:39: lijn 8 tot 20:15,
# lijn 11 tot 20:35 -> lijn 8 stond echt stil; lijn 11 liep zelf maar 4,6 min achter (= ophaling).
# Geeft $null als geen enkel blad leesbaar is; de aanroeper valt dan terug op de oude regel.
function Get-CollectEnd($sheets, [datetime]$cutoff) {
    $best = $null
    foreach ($s in $sheets) {
        if (([string]$s.Name) -notmatch '^Boxruw') { continue }
        try {
            $cells = $s.Cells; $anchor = $cells.Item($xlMaxRows, 2); $lastCell = $anchor.End($xlUp)
            $lastRow = [int]$lastCell.Row
            Rel $lastCell; Rel $anchor; Rel $cells
            if ($lastRow -lt 11) { continue }
            $vals  = $s.Range("A11:A$lastRow").Value2      # EEN marshaling-call, alleen de tijdkolom
            $isArr = ($vals -is [array])
            $hi    = if ($isArr) { $vals.GetUpperBound(0) } else { 1 }
            for ($i = $hi; $i -ge 1; $i--) {
                $v = if ($isArr) { $vals.GetValue($i, 1) } else { $vals }
                if ($v -isnot [double]) { continue }
                $ts = $null
                try { $ts = [DateTime]::FromOADate([double]$v) } catch { continue }
                if ($ts -gt $cutoff) { continue }          # tijdmachine (-Now): niets uit de toekomst
                if ($null -eq $best -or $ts -gt $best) { $best = $ts }
                break
            }
        }
        catch { }
    }
    return $best
}

# PLOEGROOSTER. Door de week DRIE ploegen van 8 u (05-13 / 13-21 / 21-05), maar op ZATERDAG en
# ZONDAG TWEE ploegen van 12 u: 05-17 ('Vr') en 17-05 ('La'). Zo staat het in het planbestand
# (blad 'Line Schedule Data': Schedule_Shift 1 = 05:00-17:00, 2 = 17:00-05:00; op 'daily shift
# dpp' is de 3e kolom van een weekenddag altijd 0) en zo telt SAPSTATus (kolommen Zo Vr/Zo La,
# Za Vr/Za La). Vroeger sneed het dashboard ook het weekend in 3x8 u en legde het de dozen van
# 05-13 naast het plan van 05-17: lijn 11 op zo 13/09 kwam zo op 65 % terwijl er ~97 % gemaakt was.
# De productiedag begint om 05:00, dus zaterdag 17-05 loopt door tot zondag 05:00.
function Test-WeekendDay([datetime]$prodDate) {
    return ($prodDate.DayOfWeek -eq [DayOfWeek]::Saturday -or $prodDate.DayOfWeek -eq [DayOfWeek]::Sunday)
}
# Ploegnummer (1..3, weekend 1..2) voor een uur van een PRODUCTIEDAG (uur 00-04 = de nacht erna).
function Get-ShiftNoForHour([datetime]$prodDate, [int]$hour) {
    if (Test-WeekendDay $prodDate) { if ($hour -ge 5 -and $hour -lt 17) { return 1 } else { return 2 } }
    if ($hour -ge 5 -and $hour -lt 13) { return 1 }
    if ($hour -ge 13 -and $hour -lt 21) { return 2 }
    return 3
}
# Huidige ploeg-interval op basis van 'nu' (zie het ploegrooster hierboven).
function Get-ShiftWindow([datetime]$now) {
    $prod = if ($now.Hour -lt 5) { $now.Date.AddDays(-1) } else { $now.Date }
    $no   = Get-ShiftNoForHour $prod $now.Hour
    $b    = Get-ShiftBounds $prod $no
    $lab  = '{0}-{1}' -f $b[0].ToString('HH:mm'), $b[1].ToString('HH:mm')
    return [pscustomobject]@{ Start = $b[0]; End = $b[1]; Label = $lab; Code = [string]$no }
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

# Goedkope 'is het bronbestand veranderd?'-stempel voor de webserver-bewaking.
# Lost het pad OPNIEUW op (zo telt ook een nieuwe snapshot-bestandsnaam mee) en leest ALLEEN de
# metadata (schrijftijd + grootte) - NIET de 4 MB inhoud; dat is een piep-kleine SMB-call.
# Schrijftijd EN grootte samen = robuuster tegen een enkel veld dat niet meebeweegt.
# LET OP CACHING: Windows/SMB cachet mapmetadata (standaard ~10s), dus een wijziging kan met die
# vertraging pas 'gezien' worden. $fi.Refresh() forceert een verse query; de ~10s vertraging is voor
# een ploegweergave prima. Netwerkhik -> $null (dan blijft de laatste goede cache staan).
function Get-BoxFileStamp {
    try { $path = Resolve-BoxFile } catch { return $null }
    try {
        $fi = New-Object System.IO.FileInfo($path)
        $fi.Refresh()
        if (-not $fi.Exists) { return $null }
        return ('{0}|{1}|{2}' -f $path, $fi.LastWriteTimeUtc.Ticks, $fi.Length)
    } catch { return $null }
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
# Kandidaat-planbestanden, nieuwste week eerst. LET OP: sorteren op weeknummer ALLEEN gaat in
# januari mis - week 52 van vorig jaar zou dan boven week 1 van dit jaar komen te staan en met
# 'Select -First N' viel het juiste bestand er zelfs helemaal uit. Daarom eerst op JAAR sorteren
# (bestandsnaam 'daily shift NDwk<week>-<jaar>.xls'); geen jaar in de naam -> 0, dus achteraan.
function Get-PlanCandidates([string]$folder) {
    if ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder)) { return @() }
    $list = Get-ChildItem -LiteralPath $folder -Filter 'daily shift NDwk*.xls*' -File -ErrorAction SilentlyContinue |
            ForEach-Object {
                $w = 0; $y = 0
                if ($_.Name -match 'NDwk0*(\d+)\s*-\s*(\d{4})') { $w = [int]$Matches[1]; $y = [int]$Matches[2] }
                elseif ($_.Name -match 'NDwk0*(\d+)')           { $w = [int]$Matches[1] }
                [pscustomobject]@{ File = $_.FullName; Year = $y; Week = $w; Modified = $_.LastWriteTime }
            }
    return @($list | Sort-Object Year, Week, Modified -Descending)
}

# ---------- LIJNFILTER UIT HET PLAN: blad 'Line Schedule Data' ----------
# Het weekplan op 'daily shift dpp' is FABRIEKBREED. Om er de target van EEN lijn uit te halen
# moet je weten welke smaak op welke lijn hoort - en dat staat in hetzelfde bestand, op blad
# 'Line Schedule Data': per planregel een Platform_ID (5 = BAKED, 6 = Extruder, 7 = Fryer,
# 4 = Multipak, 10 = Doritos 2) met begin- en eindtijd. Dat is exact en tijdgebonden, terwijl
# de oude filter (alle smaken die de lijn in 31 dagen draaide) te ruim is zodra twee lijnen
# hetzelfde assortiment delen.
# Resultaat: lijst van regels @{ Platform; Sku; Start; End } - of een lege lijst als het blad
# ontbreekt (oudere planbestanden), waarna alles terugvalt op de oude filter.
function Read-LineSchedule($sheets) {
    $out = @()
    $ws = $null
    foreach ($s in $sheets) { if ($s.Name -eq 'Line Schedule Data') { $ws = $s; break } }
    if ($null -eq $ws) { return $out }
    try {
        # kop zoeken (kolomnamen staan in rij 1) en daarna alleen de vier kolommen lezen die
        # we nodig hebben; het blad is ~1.000 regels breed 64 kolommen, alles inlezen is zonde.
        $hdr = $ws.Range("A1:BL1").Value2
        $colPlat = 0; $colStart = 0; $colEnd = 0; $colSku = 0; $colTgt = 0
        $cMax = $hdr.GetUpperBound(1)
        for ($c = 1; $c -le $cMax; $c++) {
            switch (([string]($hdr.GetValue(1, $c))).Trim()) {
                'Platform_ID'       { $colPlat  = $c }
                'Start_Timestamp'   { $colStart = $c }
                'End_Timestamp'     { $colEnd   = $c }
                'Product_Code'      { $colSku   = $c }
                'Target_Case_Count' { $colTgt   = $c }
            }
        }
        if ($colPlat -eq 0 -or $colStart -eq 0 -or $colEnd -eq 0 -or $colSku -eq 0) { return $out }
        $last = [int]($ws.UsedRange.Rows.Count)
        if ($last -lt 2) { return $out }
        if ($last -gt 20000) { $last = 20000 }
        $lo = [Math]::Min([Math]::Min($colPlat, $colStart), [Math]::Min($colEnd, $colSku))
        $hi = [Math]::Max([Math]::Max($colPlat, $colStart), [Math]::Max($colEnd, $colSku))
        # Target_Case_Count (optioneel): 0 = standby-/stilstandregel, geen echte productie.
        # Alleen de volgorde 'wat komt er nog' in de producttabel kijkt ernaar.
        if ($colTgt -gt 0) { $lo = [Math]::Min($lo, $colTgt); $hi = [Math]::Max($hi, $colTgt) }
        $vals = $ws.Range($ws.Cells(2, $lo), $ws.Cells($last, $hi)).Value2
        $rMax = $vals.GetUpperBound(0)
        for ($r = 1; $r -le $rMax; $r++) {
            $sap = Format-Sap ($vals.GetValue($r, $colSku - $lo + 1))
            if (-not (Is-Sku $sap)) { continue }
            $p  = $vals.GetValue($r, $colPlat  - $lo + 1)
            $st = $vals.GetValue($r, $colStart - $lo + 1)
            $en = $vals.GetValue($r, $colEnd   - $lo + 1)
            if ($p -isnot [double] -or $st -isnot [double] -or $en -isnot [double]) { continue }
            $tg = $null
            if ($colTgt -gt 0) { $tv = $vals.GetValue($r, $colTgt - $lo + 1); if ($tv -is [double]) { $tg = [double]$tv } }
            $out += [pscustomobject]@{
                Platform = [int]$p; Sku = $sap
                Start = [DateTime]::FromOADate($st); End = [DateTime]::FromOADate($en)
                Target = $tg
            }
        }
    }
    catch { }
    finally { Rel $ws }
    return $out
}

# Smaken die op EEN platform gepland staan in het venster [$from, $to) - de regels mogen er
# gedeeltelijk overheen lopen (een run loopt vaak over de ploeggrens heen).
function Get-PlatformSkus($sched, $platform, [datetime]$from, [datetime]$to) {
    $set = @{}
    if ($null -eq $sched -or $null -eq $platform) { return $set }
    foreach ($row in @($sched)) {
        if ($row.Platform -ne [int]$platform) { continue }
        if ($row.End -le $from -or $row.Start -ge $to) { continue }
        $set[$row.Sku] = $true
    }
    return $set
}

# VOLGORDE 'WAT KOMT ER NOG' (producttabel lijnen 5/6/8/11): per smaak het vroegste geplande begin
# dat nog OPEN staat binnen de lopende productieweek. Open = het begin ligt na 'nu', OF de run is
# gepland in de LOPENDE ploeg en de smaak heeft haar ploegplan nog niet gehaald (bv. 'nog niet
# gestart' gepland om 13:00 terwijl het nu 20:28 is: die hoort vooraan in de wachtrij, niet weg).
# Bron: blad 'Line Schedule Data' (tijd op de minuut; regels met Target_Case_Count 0 zijn standby
# of stilstand en tellen niet). Heeft het planbestand dat blad niet (of niets voor dit platform),
# dan het ploegraster van 'daily shift dpp': eerste ploeg met plan > 0, tijd = begin van die ploeg.
# $shiftPlan / $shiftCount = plan en dozen van de LOPENDE ploeg per smaak.
function Get-NextPlannedStart($sched, $grid, $platform, $skus, [datetime]$nowDt, $win, [datetime]$weekEnd, $shiftPlan, $shiftCount) {
    $next = @{}
    $want = @{}; foreach ($s in @($skus)) { if ($s) { $want[[string]$s] = $true } }
    $useSched = $false
    if ($null -ne $platform -and $null -ne $sched) {
        foreach ($row in @($sched)) {
            if ($row.Platform -ne [int]$platform) { continue }
            if ($row.Start -ge $weekEnd -or $row.End -le $win.Start) { continue }
            $useSched = $true
            if (-not $want.ContainsKey($row.Sku)) { continue }
            if ($null -ne $row.Target -and $row.Target -le 0) { continue }
            $open = ($row.Start -ge $nowDt)
            if (-not $open -and $row.Start -lt $win.End) {
                $pl = 0.0; if ($shiftPlan.ContainsKey($row.Sku))  { $pl = [double]$shiftPlan[$row.Sku] }
                $ct = 0;   if ($shiftCount.ContainsKey($row.Sku)) { $ct = [int]$shiftCount[$row.Sku] }
                $open = ($pl -gt 0 -and $ct -lt $pl)
            }
            if (-not $open) { continue }
            if (-not $next.ContainsKey($row.Sku) -or $row.Start -lt $next[$row.Sku]) { $next[$row.Sku] = $row.Start }
        }
    }
    if ($useSched -or $null -eq $grid) { return $next }

    # terugval: ploeg voor ploeg door het dagplanrooster, van de lopende ploeg tot het einde van de week
    $day = $win.Start.Date; $no = [int]$win.Code
    $rowMax = $grid.GetUpperBound(0)
    for ($k = 0; $k -lt 30; $k++) {
        $b = Get-ShiftBounds $day $no
        if ($b[0] -ge $weekEnd) { break }
        $col = Find-PlanColumn $grid $day $no
        if ($col -gt 0) {
            for ($r = 3; $r -le $rowMax; $r++) {
                $sap = Format-Sap ($grid.GetValue($r, 2))
                if (-not $want.ContainsKey($sap) -or $next.ContainsKey($sap)) { continue }
                $v = $grid.GetValue($r, $col)
                if ($v -isnot [double] -or $v -le 0) { continue }
                if ($k -eq 0) {
                    $pl = 0.0; if ($shiftPlan.ContainsKey($sap))  { $pl = [double]$shiftPlan[$sap] }
                    $ct = 0;   if ($shiftCount.ContainsKey($sap)) { $ct = [int]$shiftCount[$sap] }
                    if (-not ($pl -gt 0 -and $ct -lt $pl)) { continue }
                }
                $next[$sap] = $b[0]
            }
        }
        $maxNo = if (Test-WeekendDay $day) { 2 } else { 3 }
        if ($no -lt $maxNo) { $no++ } else { $no = 1; $day = $day.AddDays(1) }
    }
    return $next
}

# Per smaak het begin van de LAATSTE ploeg van deze week waarin ze gemaakt is: eerdere ploegen uit
# de historie (WeekIdx 0), de lopende ploeg uit het Boxruw-blad ($counts) - die is altijd de laatste.
function Get-LastMadeMap($history, $counts, [datetime]$shiftStart) {
    $last = @{}
    foreach ($h in @($history)) {
        if ($h.WeekIdx -ne 0) { continue }
        $hb = Get-ShiftBounds $h.Date $h.ShiftNo
        foreach ($s in $h.Skus) { if (-not $last.ContainsKey($s.Sku) -or $hb[0] -gt $last[$s.Sku]) { $last[$s.Sku] = $hb[0] } }
    }
    foreach ($p in $counts.Keys) { if ([int]$counts[$p] -gt 0) { $last[$p] = $shiftStart } }
    return $last
}

# PLAATS VAN EEN SMAAKREGEL in de producttabel (lijnen 5/6/8/11) en in 'Weekplan per smaak' (lijn 9).
# Vraag gebruiker 17/09/2026: bovenaan wat NU loopt, dan wat volgens het rooster van deze week nog
# komt (op gepland begin), dan de rest. Groepen:
#   0 = nu op een draaiende machine (meeste dozen deze ploeg eerst)
#   1 = wacht in het weekrooster, op gepland begin ($nextAt uit Get-NextPlannedStart)
#   2 = deze week nog helemaal niet gemaakt en niet meer ingepland
#   3 = deze week al gemaakt en niet meer ingepland (laatst gemaakt eerst)
#   9 = geen SAP-code ('(onbekend)')
# Staat een smaak in het plan van de LOPENDE ploeg, nog niet gehaald, maar kent het rooster er geen
# tijd voor (regel ontbreekt op 'Line Schedule Data'), dan telt het begin van de ploeg.
function Get-SkuOrder([string]$sku, [bool]$isCur, $nextAt, $lastMade, [double]$shiftCount, [double]$shiftPlan, [double]$weekMade, [double]$weekPlan, [datetime]$shiftStart) {
    $na = $null; if ($nextAt.ContainsKey($sku))   { $na = $nextAt[$sku] }
    $lm = $null; if ($lastMade.ContainsKey($sku)) { $lm = $lastMade[$sku] }
    if (-not $isCur -and $null -eq $na -and $shiftPlan -gt 0 -and $shiftCount -lt $shiftPlan) { $na = $shiftStart }
    if (-not (Is-Sku $sku))  { $g = 9; $k1 = 0.0; $k2 = 0.0 }
    elseif ($isCur)          { $g = 0; $k1 = -$shiftCount; $k2 = -$shiftPlan }
    elseif ($null -ne $na)   { $g = 1; $k1 = [double]$na.Ticks; $k2 = -$weekPlan }
    elseif ($weekMade -le 0 -and $shiftCount -le 0) { $g = 2; $k1 = -$weekPlan; $k2 = 0.0 }
    else {
        $g = 3; $k2 = -$weekMade
        $k1 = if ($null -ne $lm) { -[double]$lm.Ticks } else { 0.0 }
    }
    return [pscustomobject]@{ NextAt = $na; LastMade = $lm; Group = $g; Key = $k1; Key2 = $k2 }
}

# Venster van een productiedag + ploeg: door de week 05-13 / 13-21 / 21-05, in het weekend
# 05-17 / 17-05 (dag begint 05:00). Ploeg N is ook de N-de kolom van die dag op 'daily shift dpp'.
function Get-ShiftBounds([datetime]$prodDate, [int]$shiftNo) {
    $len = if (Test-WeekendDay $prodDate) { 12 } else { 8 }
    $st = $prodDate.Date.AddHours(5 + $len * ($shiftNo - 1))
    return @($st, $st.AddHours($len))
}

# Leest het plan voor productiedag + ploeg. $skus = producten die deze ploeg draaiden,
# $lineSkus = terugvalfilter (alle producten van de lijn uit de 31-daagse *RCDB-tabel); als
# het blad 'Line Schedule Data' er is EN de lijn een Platform heeft, wint die exacte filter.
function Read-PlanTargets($excel, [string]$planPath, [datetime]$prodDate, [int]$shiftNo, $skus, $lineSkus) {
    $res = [ordered]@{
        Ok = $false; File = (Split-Path $planPath -Leaf); WeekNo = $null; Period = ''
        Covered = $false; Total = 0.0; PerSku = @{}; Desc = @{}; FallbackSku = $null; Error = $null
        Grid = $null; SkuNames = @{}; Sched = @(); PlatformUsed = $false
    }
    $wb = $null; $ws = $null; $sheets = $null
    try {
        $wb = $excel.Workbooks.Open($planPath, 0, $true)          # UpdateLinks=0, ReadOnly=True
        if (-not $wb.ReadOnly) { throw "Planbestand niet in alleen-lezen geopend - afgebroken." }
        $sheets = $wb.Worksheets
        foreach ($s in $sheets) { if ($s.Name -eq 'daily shift dpp') { $ws = $s; break } }
        if ($null -eq $ws) { foreach ($s in $sheets) { if ($s.Name -like '*dpp*') { $ws = $s; break } } }
        if ($null -eq $ws) { throw "Werkblad 'daily shift dpp' niet gevonden in $($res.File)." }

        # VOLLEDIGE naamlijst uit blad 'sku' (de hele catalogus, ~2.500 regels; B = SAP, C = tekst).
        # Nodig omdat 'daily shift dpp' alleen de smaken van DEZE week bevat: een smaak die eerder
        # deze week wel gedraaid heeft maar nu niet gepland staat, bleef anders naamloos in de
        # tabel staan. Eenmalig per sessie ophalen; de catalogus verandert nauwelijks.
        if ($script:SkuNameCache -and $script:SkuNameCache.Count -gt 0) {
            $res.SkuNames = $script:SkuNameCache
        }
        else {
            $wsSku = $null
            foreach ($s2 in $sheets) { if ($s2.Name -eq 'sku') { $wsSku = $s2; break } }
            if ($null -eq $wsSku) { foreach ($s2 in $sheets) { if ($s2.Name -like 'sku*') { $wsSku = $s2; break } } }
            if ($wsSku) {
                try {
                    $sv    = $wsSku.Range("B1:C3000").Value2
                    $svMax = $sv.GetUpperBound(0)
                    for ($r2 = 2; $r2 -le $svMax; $r2++) {
                        $sp = Format-Sap ($sv.GetValue($r2, 1))
                        if (-not (Is-Sku $sp)) { continue }
                        if ($res.SkuNames.ContainsKey($sp)) { continue }
                        $nm = ([string]($sv.GetValue($r2, 2))).Trim()
                        if ($nm) { $res.SkuNames[$sp] = $nm }
                    }
                    if ($res.SkuNames.Count -gt 0) { $script:SkuNameCache = $res.SkuNames }
                }
                catch { }
                Rel $wsSku
            }
        }

        # exacte lijnfilter uit hetzelfde bestand (zie Read-LineSchedule hierboven)
        $res.Sched = @(Read-LineSchedule $sheets)
        if ($res.Sched.Count -gt 0 -and $null -ne $script:Platform) {
            $b = Get-ShiftBounds $prodDate $shiftNo
            $exact = Get-PlatformSkus $res.Sched $script:Platform $b[0] $b[1]
            if ($exact.Count -gt 0) { $lineSkus = $exact; $res.PlatformUsed = $true }
        }

        $vals   = $ws.Range("A1:AD400").Value2
        $rowMax = $vals.GetUpperBound(0); $colMax = $vals.GetUpperBound(1)
        # ruwe planroostercache: de historie hergebruikt hem (scheelt heropenen). METEEN zetten,
        # want een bestand van een ANDERE week valt hieronder uit bij 'dag niet in dit bestand'
        # terwijl de historie juist dat rooster nodig heeft voor de oudere weken.
        $res.Grid = $vals

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
            # telt mee voor deze ploeg: al gedraaid OF gepland voor een product dat op DEZE lijn loopt
            # (het plan is fabriekbreed, vandaar de filter op $lineSkus)
            $mine   = $want.ContainsKey($sap)
            $onPlan = ($lineSkus -and $lineSkus.ContainsKey($sap))
            # MET de exacte platformfilter telt alleen wat op DEZE lijn gepland staat. Anders
            # sleept een handvol dozen van een naburige lijn (omstelling, restje na een wissel)
            # het VOLLEDIGE dagplan van die smaak mee: di 08/09 gaf 97 dozen 340018067 op lijn 6
            # en daarmee 936 target die in werkelijkheid op lijn 5 hoorde.
            $onLine = if ($res.PlatformUsed) { $onPlan } else { $mine -or $onPlan }
            if (($mine -or $onLine) -and -not $res.Desc.ContainsKey($sap)) {
                $res.Desc[$sap] = ([string]($vals.GetValue($r, 3))).Trim()
            }
            $v = $vals.GetValue($r, $col)
            if ($v -isnot [double]) { continue }
            # dezelfde SAP-code kan meerdere planregels hebben -> optellen.
            # Een smaak die deze ploeg nog NIET gestart is telt ook mee (anders groeit de target
            # pas als de omstelling gebeurd is); daarvoor moet er wel echt iets gepland staan.
            if (($mine -and -not $res.PlatformUsed) -or ($onLine -and $v -gt 0)) {
                if ($res.PerSku.ContainsKey($sap)) { $res.PerSku[$sap] += [double]$v } else { $res.PerSku[$sap] = [double]$v }
            }
            if ($lineSkus -and $lineSkus.ContainsKey($sap) -and $v -gt $bestVal) { $bestVal = [double]$v; $best = $sap }
        }
        $tot = 0.0; foreach ($k2 in $res.PerSku.Keys) { $tot += [double]$res.PerSku[$k2] }
        # nog niets gemaakt deze ploeg -> plan van de lijn gebruiken (en melden welk product verwacht wordt)
        if ($want.Count -eq 0 -and $best -and $bestVal -gt 0) {
            $res.FallbackSku = $best
            if ($tot -le 0) { $res.PerSku[$best] = $bestVal; $tot = $bestVal }
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

# ---------- HISTORIE: dozen per productiedag/ploeg uit het *RCDB-blad ----------
# Kolom in het planrooster voor (productiedag, ploeg) - zelfde telling als in Read-PlanTargets.
function Find-PlanColumn($vals, [datetime]$prodDate, [int]$shiftNo) {
    if ($null -eq $vals) { return 0 }
    $colMax = $vals.GetUpperBound(1)
    $key = $prodDate.ToString('yyyy-MM-dd'); $prevKey = $null; $shift = 0
    for ($c = 4; $c -le $colMax; $c++) {
        $dt = Convert-Serial ($vals.GetValue(2, $c))
        if ($null -eq $dt) { continue }
        $k = $dt.ToString('yyyy-MM-dd')
        if ($k -ne $prevKey) { $shift = 0; $prevKey = $k } else { $shift++ }
        if ($k -eq $key -and $shift -eq ($shiftNo - 1)) { return $c }
    }
    return 0
}
# Plan uit die kolom, gefilterd op producten van DEZE lijn (het plan is fabriekbreed).
# $strict = de filter komt uit 'Line Schedule Data' en is exact; dan telt UITSLUITEND wat op
# deze lijn gepland stond, ook al zijn er losse dozen van een andere lijn gedraaid.
function Get-PlanForColumn($vals, [int]$col, $ranSkus, $lineSkus, [bool]$strict = $false) {
    $per = @{}
    if ($null -ne $vals -and $col -gt 0) {
        $rowMax = $vals.GetUpperBound(0)
        for ($r = 3; $r -le $rowMax; $r++) {
            $sap = Format-Sap ($vals.GetValue($r, 2))
            if (-not (Is-Sku $sap)) { continue }
            $onPlan = ($lineSkus -and $lineSkus.ContainsKey($sap))
            $mine   = if ($strict) { $onPlan -and $ranSkus.ContainsKey($sap) } else { $ranSkus.ContainsKey($sap) }
            if (-not ($mine -or $onPlan)) { continue }
            $v = $vals.GetValue($r, $col)
            if ($v -isnot [double]) { continue }
            if (-not $mine -and $v -le 0) { continue }
            if ($per.ContainsKey($sap)) { $per[$sap] += [double]$v } else { $per[$sap] = [double]$v }
        }
    }
    $t = 0.0; foreach ($k in $per.Keys) { $t += [double]$per[$k] }
    return [pscustomobject]@{ Total = $t; PerSku = $per }
}
# Alle producttypes die op DEZE lijn gedraaid hebben volgens het *RCDB-blad (31 dagen).
# Dit is de filter waarmee het FABRIEKBREDE weekplan wordt teruggebracht tot deze lijn.
# Bewust ruimer dan het Boxruw-blad (dat reikt maar ~2 dagen terug): een smaak die deze ploeg
# GEPLAND staat maar nog niet gestart is, moet ook meetellen in de ploegtarget.
function Get-RcdbSkus($vals) {
    $set = @{}
    if ($null -eq $vals) { return $set }
    $rowMax = $vals.GetUpperBound(0); $colMax = $vals.GetUpperBound(1)
    for ($b = 1; $b + 4 -le $colMax; $b += 5) {
        for ($r = 4; $r -le $rowMax; $r++) {
            $sap = Format-Sap ($vals.GetValue($r, $b + 1))
            if (-not (Is-Sku $sap)) { continue }
            $n = $vals.GetValue($r, $b + 4)
            if ($n -isnot [double] -or $n -le 0) { continue }
            $set[$sap] = $true
        }
    }
    return $set
}
# Weekplan per smaak: som van ALLE ploegen van de productieweek (zondag t/m zaterdag) uit
# hetzelfde 'daily shift dpp'-rooster. Zelfde lijnfilter als de ploegtarget (plan is fabriekbreed).
function Get-PlanForWeek($vals, [datetime]$weekStart, $lineSkus) {
    $per = @{}; $desc = @{}
    if ($null -ne $vals) {
        $rowMax = $vals.GetUpperBound(0); $colMax = $vals.GetUpperBound(1)
        $weekEnd = $weekStart.AddDays(6)
        $cols = @()
        for ($c = 4; $c -le $colMax; $c++) {
            $dt = Convert-Serial ($vals.GetValue(2, $c))
            if ($null -eq $dt) { continue }
            if ($dt.Date -ge $weekStart -and $dt.Date -le $weekEnd) { $cols += $c }
        }
        if ($cols.Count -gt 0) {
            for ($r = 3; $r -le $rowMax; $r++) {
                $sap = Format-Sap ($vals.GetValue($r, 2))
                if (-not (Is-Sku $sap)) { continue }
                if (-not ($lineSkus -and $lineSkus.ContainsKey($sap))) { continue }
                foreach ($c in $cols) {
                    $v = $vals.GetValue($r, $c)
                    if ($v -isnot [double] -or $v -le 0) { continue }
                    if ($per.ContainsKey($sap)) { $per[$sap] += [double]$v } else { $per[$sap] = [double]$v }
                    if (-not $desc.ContainsKey($sap)) { $desc[$sap] = ([string]($vals.GetValue($r, 3))).Trim() }
                }
            }
        }
    }
    return [pscustomobject]@{ PerSku = $per; Desc = $desc }
}
# Leest het blad 'L9 datatabel voor 31 dagen': per DAG VAN DE MAAND een blok van 5 kolommen
# [uur, SAP, doosprint, Machine, aantal]; rij 2 boven het blok = dagnummer, rij 3 = kopjes,
# data vanaf rij 4. Het blok van dag D bevat de uren 05..23 van D EN 00..04 van D+1 -> dat is
# precies de PRODUCTIEDAG. Ploeg: 1 = 5-12, 2 = 13-20, 3 = 21-4; weekend 1 = 5-16, 2 = 17-4
# (Get-ShiftNoForHour; letter: zie Get-ShiftLetter).
# ISO-weeknummer (System.Globalization.ISOWeek bestaat niet in Windows PowerShell 5.1).
function Get-IsoWeek([datetime]$dt) {
    $dow = [int]$dt.DayOfWeek; if ($dow -eq 0) { $dow = 7 }
    $thu = $dt.AddDays(4 - $dow)
    return [int][Math]::Floor(($thu.DayOfYear - 1) / 7) + 1
}
# Ploegletter X/Y/Z. De nacht (21-05) is ALTIJD Z; X en Y wisselen per week:
#   even week -> X = ochtend (05-13), Y = namiddag (13-21)
#   oneven week -> Y = ochtend, X = namiddag
# Het weeknummer is dat van de productieweek (die op zondag start, dus de ISO-week
# van de maandag erna - zelfde nummering als in de kop van de historie).
# Weekend (2 ploegen van 12 u): 'Vr' = 05-17 en 'La' = 17-05, net als SAPSTATus (Zo Vr / Za La).
function Get-ShiftLetter([datetime]$prodDate, [int]$shiftNo) {
    if (Test-WeekendDay $prodDate) { if ($shiftNo -eq 1) { return 'Vr' } else { return 'La' } }
    if ($shiftNo -eq 3) { return 'Z' }
    $weekStart = $prodDate.AddDays(-[int]$prodDate.DayOfWeek)   # zondag
    $wk = Get-IsoWeek $weekStart.AddDays(1)                     # maandag van die week
    $evenWeek = (($wk % 2) -eq 0)
    if ($shiftNo -eq 1) { if ($evenWeek) { return 'X' } else { return 'Y' } }
    if ($evenWeek) { return 'Y' } else { return 'X' }
}
# De tabel is een rollend venster op dagnummer: dagen NA vandaag horen bij de vorige maand.
# $first = oudste dag die meetelt, $curWeekStart = zondag van de LOPENDE productieweek
# (elke rij krijgt WeekIdx: 0 = deze week, 1 = vorige week, ...). $planGrids = alle gelezen
# planroosters; voor oudere weken zit het plan in een ANDER 'daily shift NDwk*'-bestand.
# LIJN 11: $valsList is een LIJST roosters (L11P1RCDB + L11P2RCDB). Ze worden opgeteld -
# precies zoals SAPSTATus doet (de som klopt 1:1 met het dag/ploeg-raster op blad 'W <week>').
function Build-History($valsList, [datetime]$curProdDate, [datetime]$first, [datetime]$curWeekStart, $planGrids, $planScheds) {
    $out = @()
    $grids = @($valsList | Where-Object { $null -ne $_ })
    if ($grids.Count -eq 0) { return $out }
    $agg = @{}          # "yyyy-MM-dd|ploeg" -> @{ sap = dozen }
    $lineSkus = @{}
    foreach ($vals in $grids) {
    $rowMax = $vals.GetUpperBound(0); $colMax = $vals.GetUpperBound(1)
    for ($b = 1; $b + 4 -le $colMax; $b += 5) {
        $dayV = $vals.GetValue(2, $b)
        if ($dayV -isnot [double]) { continue }
        $day = [int]$dayV
        if ($day -lt 1 -or $day -gt 31) { continue }
        # dagnummer -> echte datum (rollend venster van 31 dagen)
        $base = if ($day -le $curProdDate.Day) { $curProdDate } else { $curProdDate.AddMonths(-1) }
        if ($day -gt [DateTime]::DaysInMonth($base.Year, $base.Month)) { continue }
        # LET OP: New-Object/[datetime]::new - Get-Date -Hour 0 laat de milliseconden staan,
        # waardoor 'vandaag' groter is dan de productiedatum en de dag van vandaag wegviel.
        $date = New-Object DateTime($base.Year, $base.Month, $day)
        if ($date -lt $first -or $date -gt $curProdDate) { continue }
        for ($r = 4; $r -le $rowMax; $r++) {
            $uur = $vals.GetValue($r, $b)
            if ($uur -isnot [double]) { continue }
            $h = [int]$uur
            if ($h -lt 0 -or $h -gt 23) { continue }
            $sap = Format-Sap ($vals.GetValue($r, $b + 1))
            if (-not (Is-Sku $sap)) { continue }
            $n = $vals.GetValue($r, $b + 4)
            if ($n -isnot [double] -or $n -le 0) { continue }
            $lineSkus[$sap] = $true
            $pl  = Get-ShiftNoForHour $date $h          # weekend: 2 ploegen van 12 u
            $key = '{0}|{1}' -f $date.ToString('yyyy-MM-dd'), $pl
            if (-not $agg.ContainsKey($key)) { $agg[$key] = @{} }
            if ($agg[$key].ContainsKey($sap)) { $agg[$key][$sap] += [int]$n } else { $agg[$key][$sap] = [int]$n }
        }
    }
    }

    foreach ($key in $agg.Keys) {
        $parts = $key -split '\|'
        $date = [datetime]::ParseExact($parts[0], 'yyyy-MM-dd', $null)
        $pl   = [int]$parts[1]
        # losse dozen (<= 2 per producttype) zijn ruis rond een omstelling - SAPSTATus laat ze ook weg
        $skus = @()
        $tot  = 0
        foreach ($sap in ($agg[$key].Keys | Sort-Object { $agg[$key][$_] } -Descending)) {
            $n = [int]$agg[$key][$sap]
            if ($n -le 2) { continue }
            $skus += [pscustomobject]@{ Sku = $sap; Count = $n }
            $tot += $n
        }
        if ($skus.Count -eq 0) { continue }
        $ran = @{}; foreach ($s in $skus) { $ran[$s.Sku] = $true }
        # plan uit het rooster dat DEZE datum bevat (oudere weken = ander planbestand).
        # De lijnfilter komt bij voorkeur uit 'Line Schedule Data' van HETZELFDE bestand
        # (exact en per ploeg), anders uit de 31-daagse *RCDB-lijst hierboven.
        $pln = [pscustomobject]@{ Total = 0.0; PerSku = @{} }
        $gl = @($planGrids); $sl = @($planScheds)
        for ($gi = 0; $gi -lt $gl.Count; $gi++) {
            $col = Find-PlanColumn $gl[$gi] $date $pl
            if ($col -le 0) { continue }
            $flt = $lineSkus; $strict = $false
            if ($null -ne $script:Platform -and $gi -lt $sl.Count) {
                $b  = Get-ShiftBounds $date $pl
                $ex = Get-PlatformSkus $sl[$gi] $script:Platform $b[0] $b[1]
                if ($ex.Count -gt 0) { $flt = $ex; $strict = $true }
            }
            $pln = Get-PlanForColumn $gl[$gi] $col $ran $flt $strict
            break
        }
        $wk = [int][Math]::Floor(($curWeekStart - $date.AddDays(-[int]$date.DayOfWeek)).TotalDays / 7)
        $bnd = Get-ShiftBounds $date $pl
        $out += [pscustomobject]@{
            Date = $date; ShiftNo = $pl; Letter = (Get-ShiftLetter $date $pl); Range = ('{0:HH}-{1:HH}' -f $bnd[0], $bnd[1])
            Total = $tot; Skus = $skus; Target = $pln.Total; PlanPerSku = $pln.PerSku
            WeekIdx = $wk; WeekStart = $date.AddDays(-[int]$date.DayOfWeek)
        }
    }
    return @($out | Sort-Object @{ Expression = 'Date'; Descending = $true }, @{ Expression = 'ShiftNo'; Descending = $false })
}

# ---- leest Boxruw9 uit de cache en telt dozen per producttype + per minuut voor de huidige ploeg ----

# ======================= LIJNREGISTER (het hart van de samenvoeging) =======================
# Dit werktuig bedient BEIDE lijnen vanuit een pagina. Wat per lijn verschilt staat hier;
# de rest van het script is gedeeld. Per lijn:
#   BoxSheets  = blad(en) met de rauwe doosregistraties (lijn 11 heeft er twee: twee doseerders)
#   RcdbSheets = blad(en) met de 31-daagse uurtabel (lijn 11 ook twee; ze worden OPGETELD)
#   Engine     = 'single'   -> lijn 9: een verpakkingsmachine, hoofdproduct = laatste rij
#                'machines' -> lijn 11: acht machines naast elkaar, elk een eigen smaak
#   Platform   = Platform_ID van deze lijn op blad 'Line Schedule Data' in het planbestand.
#                Daarmee wordt het FABRIEKBREDE weekplan teruggebracht tot deze lijn. Zonder
#                die sleutel valt het terug op de oude filter (alle smaken die de lijn in 31
#                dagen draaide) - en die is te ruim zodra twee lijnen hetzelfde assortiment
#                delen: lijn 5 (BAKED) en lijn 6 (Extruder) draaien allebei Oven Baked, en
#                daardoor kreeg lijn 6 het plan van lijn 5 erbij (target 12.865 i.p.v. 4.387).
$script:LineOrder = @('5','6','8','9','11')
$script:LineDefs  = @{
    '5'  = @{ Id = '5';  BoxSheets = @('Boxruw5');                RcdbSheets = @('L5RCDB');
              Engine = 'machines'; Platform = 5 }     # BAKED
    '6'  = @{ Id = '6';  BoxSheets = @('Boxruw6');                RcdbSheets = @('L6RCDB');
              Engine = 'machines'; Platform = 6 }     # Extruder
    '8'  = @{ Id = '8';  BoxSheets = @('Boxruw8');                RcdbSheets = @('L8RCDB');
              Engine = 'machines'; Platform = 7 }     # Fryer
    '9'  = @{ Id = '9';  BoxSheets = @('Boxruw9');                RcdbSheets = @('L9RCDB');
              Engine = 'single';   Platform = 4 }     # Multipak
    '11' = @{ Id = '11'; BoxSheets = @('Boxruw111','Boxruw112');  RcdbSheets = @('L11P1RCDB','L11P2RCDB');
              Engine = 'machines'; Platform = 10 }    # Doritos 2
}
# Een lijn toevoegen = EEN regel hierboven. Een eigen blok in config.txt mag, maar hoeft niet:
# de motor 'machines' leest het aantal machines uit tag 8 van de etiketten zelf (Get-Machine),
# de kop van de pagina komt uit de sleutel 'h1_machines' met het lijnnummer erin, en de knop
# in de lijnbalk uit 'line_word'. Nergens staat een vast machinenummer of een vast aantal.
#   lijn 5  -> machines 500..505     lijn 6  -> 600..604     lijn 8  -> 801..806
#   lijn 9  -> een machine (901)     lijn 11 -> 1101..1108 verdeeld over twee bladen
$script:Line     = '9'      # lijn die NU gerenderd wordt
$script:LineCfg  = @{}      # per lijn de ingelezen configuratie (zie onderaan)

# Zet de script-scope instellingen op die van EEN lijn. De motoren lezen deze variabelen
# (zo hoefden ze voor de samenvoeging niet herschreven te worden).
function Set-LineContext([string]$id) {
    if (-not $script:LineDefs.ContainsKey($id)) { $id = $script:LineOrder[0] }
    $script:Line = $id
    $c = $script:LineCfg[$id]
    if ($null -eq $c) { return }
    $script:BoxSheets  = @($c.BoxSheets)
    $script:RcdbSheets = @($c.RcdbSheets)
    $script:Platform   = $script:LineDefs[$id].Platform
    $script:RcdbSheet  = ($c.RcdbSheets -join ',')
    Set-Variable -Scope Script -Name BoxSheet      -Value ($c.BoxSheets -join ',')
    Set-Variable -Scope Script -Name RcdbSheet     -Value ($c.RcdbSheets -join ',')
    Set-Variable -Scope Script -Name ShiftTarget   -Value ([int]$c.ShiftTarget)
    Set-Variable -Scope Script -Name RecentMinutes -Value ([int]$c.RecentMinutes)
    Set-Variable -Scope Script -Name StopMinutes   -Value ([double]$c.StopMinutes)
    Set-Variable -Scope Script -Name HistoryDays   -Value ([int]$c.HistoryDays)
    Set-Variable -Scope Script -Name PlanFile      -Value ([string]$c.PlanFile)
    $script:HasPlanFile = -not [string]::IsNullOrWhiteSpace($c.PlanFile)
    $script:HasTarget   = [bool]$c.HasTarget
    $script:HasRcdb     = $true
}

# De juiste motor voor de gekozen lijn. Beide leveren hetzelfde soort object op.
function Get-BoxDataFor([string]$id) {
    Set-LineContext $id
    if ($script:LineDefs[$id].Engine -eq 'machines') { return Get-BoxData11 }
    return Get-BoxData9
}

# De juiste weergave voor de gekozen lijn.
function Render-HtmlFor($d, [string]$lang, [string]$id) {
    Set-LineContext $id
    if ($script:LineDefs[$id].Engine -eq 'machines') { return Render-Html11 $d $lang }
    return Render-Html9 $d $lang
}

# ?line=9 / ?line=11 uit het verzoek halen (net als ?lang=)
function Get-ReqLine([string]$req) {
    if ($req -match '[?&]line=(\d{1,2})') { $l = $matches[1]; if ($script:LineOrder -contains $l) { return $l } }
    return $script:Line
}

# ============================ MOTOR LIJN 9 (een machine) ============================
function Get-BoxData9 {
    $nowDt = if ($script:HasNow) { $Now } else { Get-Date }
    $win = Get-ShiftWindow $nowDt
    $shiftMin = [int][math]::Round(($win.End - $win.Start).TotalMinutes)
    if ($shiftMin -le 0) { $shiftMin = 480 }
    $minutes = New-Object 'int[]' $shiftMin
    $targetPerMin = if ($ShiftTarget -gt 0) { [double]$ShiftTarget / $shiftMin } else { 0 }

    $d = [ordered]@{
        Ok = $true; Error = $null
        NowText = $nowDt.ToString('dd/MM/yyyy HH:mm')
        ShiftRange = $win.Label
        ShiftLetter = (Get-ShiftLetter $win.Start.Date ([int]$win.Code))
        WindowText = ('{0} -> {1}' -f $win.Start.ToString('dd/MM HH:mm'), $win.End.ToString('dd/MM HH:mm'))
        Sheet = $BoxSheet; BoxFile = ''; FileTimeText = '-'
        TargetMode = 'unknown'; PlanDate = $null; WarnList = @(); PlanFileName = ''; PlanWeek = $null
        PlanSku = $null; ShiftNo = 0
        Rows = @(); Tempo = @{}; Total = 0; ParsedRows = 0; LastText = ''; StartRowSkipped = $false
        BlankRows = 0                      # dozen zonder etiketstring: overgeslagen, niet geteld
        # losse onderdelen van 'laatste doos' - de zin zelf wordt PAS in Render-Html gezet (taal!)
        LastTimeText = ''; LastProduct = ''; LastCounter = ''
        ShiftStart = $win.Start; ShiftEnd = $win.End; ShiftMin = $shiftMin; Minutes = $minutes; MaxPerMin = 0
        Target = $ShiftTarget; TargetPerMin = $targetPerMin; Pct = 0
        # --- prognose einde ploeg ---
        HasForecast = $false; FcWeak = $false
        MainProduct = $null                       # hoofdproduct = producttype van de LAATSTE rij
        RunCount = 0; RunStartText = ''           # huidige aaneengesloten run van dat product
        RefNowText = '-'; ElapsedMin = 0; RemainMin = 0; NowOffsetMin = 0
        PerMin = 0; PerHour = 0
        RecentWin = 0; RecentCount = 0; HasRecent = $false; RecentPerMin = 0; ProjRecent = 0
        ProjTotal = 0; ProjMain = 0; ProjPct = 0; ProjDiff = 0
        EtaText = ''; EtaKind = ''; EtaTimeText = ''; EtaAfterShift = $false
        # --- stilstand + achterstand ---
        HasStops = $false; Stops = @(); StopCount = 0; StopMin = 0
        LongestMin = 0; LongestText = ''; NowStill = $false; StillMin = 0
        ElapsedShiftMin = 0; RunMin = 0; AvailPct = 0; NetPerMin = 0; LostBoxes = 0
        BehindNow = 0; BehindEnd = 0; StopLimit = $StopMinutes
        # --- historie per dag/ploeg (uit het RCDB-blad) ---
        HasHistory = $false; History = @(); HistSheet = $script:RcdbSheet
        HistFrom = $null; HistTo = $null; HistWeekStart = $null; HistMaxWeek = 0
        # --- weekplan per smaak (dagplan blijft leidend; dit is de laag eronder) ---
        HasWeekPlan = $false; WeekRows = @(); WeekPlanTotal = 0.0; WeekMadeTotal = 0; WeekNo = $null
        # smaken waarvan het WEEKPLAN bij ploegstart al rond was (zie 'weekplan al rond' verderop)
        WeekDoneNotes = @()
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

        if ($lastRow -lt 11) { Add-Warn $d 'warn_no_data_row11' @($BoxSheet); return [pscustomobject]$d }

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
            # LEGE B = EEN DOOS ZONDER ETIKET, GEEN EINDE VAN DE GEGEVENS. Zo'n rij heeft wel een
            # tijd en C='DELTA', alleen de etiketstring is leeg (18/09 gebeurde dat op lijn 6 om
            # 09:03:11 en 16:21:06, midden in een normale reeks). Dit LAS VROEGER ALS 'break' en
            # gooide alles daarna weg: van de 12.022 dozen van lijn 6 bleven er 3.021 over en de
            # hele ploeg 13-21 u was leeg, terwijl SAPSTATus gewoon productie toonde.
            # Niet meetellen als doos: het uurblad (*RCDB) telt zo'n rij ook niet mee (uur 09 en
            # uur 16 klopten tot op de doos met wat wij zonder deze rijen tellen).
            # Het leesbereik is toch al begrensd door de laatste gevulde rij van kolom B, dus
            # doorlopen kan geen losgeslagen staart opleveren.
            if ($null -eq $b -or ([string]$b).Trim() -eq '') { $d.BlankRows++; continue }
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
        if ($lastTs) {
            $d.LastTimeText = $lastTs.ToString('HH:mm'); $d.LastProduct = [string]$lastProd; $d.LastCounter = [string]$lastCtr
            $d.LastText = ('{0} - {1} (doos {2})' -f $d.LastTimeText, $d.LastProduct, $d.LastCounter)   # nl-terugval
        }

        # ---- producten van DEZE lijn: filter voor het fabriekbrede weekplan ----
        # Het *RCDB-blad wordt hier AL gelezen (en verderop hergebruikt voor de historie), want
        # zonder die 31-daagse lijst telt een smaak die deze ploeg gepland staat maar nog NIET
        # gestart is niet mee in de target: do 30/07 gaf 540 (alleen 340062773, het draaiende
        # product) i.p.v. 1.157, omdat de omstelling naar 340056956 (617) er niet gekomen was.
        $histVals = $null
        $lineProds = @{}
        foreach ($p in $allProds.Keys) { $lineProds[$p] = $true }
        $ws2 = $null
        foreach ($s in $sheets) { if ($s.Name -eq $script:RcdbSheet) { $ws2 = $s; break } }
        if ($null -eq $ws2) {
            Add-Warn $d 'warn_no_rcdb' @($script:RcdbSheet)
        }
        else {
            try {
                $histVals = $ws2.Range("A1:FI300").Value2      # 31 blokken van 5 kolommen = t/m kolom FI
                foreach ($p in (Get-RcdbSkus $histVals).Keys) { $lineProds[$p] = $true }
            }
            catch { Add-Warn $d 'warn_hist_read' @($_.Exception.Message) }
            Rel $ws2
        }

        # ---------------- TARGET UIT HET WEEKPLAN ----------------
        $shiftNo = [int]$win.Code; $prodDate = $win.Start.Date
        $d.ShiftNo = $shiftNo
        $effTarget = [double]$ShiftTarget
        $planPerSku = @{}; $planDesc = @{}; $planGrids = @(); $planScheds = @()
        if (-not $NoPlan -and -not $script:HasTarget) {
            $cands = @()
            if ($script:HasPlanFile) {
                if (Test-Path -LiteralPath $PlanFile) { $cands = @([pscustomobject]@{ File = (Resolve-Path -LiteralPath $PlanFile).Path }) }
                else { Add-Warn $d 'warn_planfile_missing' @($PlanFile) }
            }
            else { $cands = @(Get-PlanCandidates $script:PlanFolder | Select-Object -First 3) }

            $pt = $null; $read = 0
            foreach ($c in $cands) {
                $try = Read-PlanTargets $excel $c.File $prodDate $shiftNo (@($counts.Keys)) $lineProds
                $read++
                if ($try.Grid) { $planGrids += ,$try.Grid; $planScheds += ,$try.Sched }   # hergebruik voor de historie
                if ($null -eq $pt -or $try.Covered) { $pt = $try }
                if ($try.Covered) { break }
            }
            # planbestanden van OUDERE weken (voor de historie-knop '+7 dagen'); alleen als ze bestaan
            if ($HistoryDays -ge 0 -and $cands.Count -gt $read) {
                foreach ($c in ($cands | Select-Object -Skip $read)) {
                    $g = Read-PlanTargets $excel $c.File $prodDate $shiftNo (@($counts.Keys)) $lineProds
                    if ($g.Grid) { $planGrids += ,$g.Grid; $planScheds += ,$g.Sched }
                }
            }
            if ($null -eq $pt) {
                if ($cands.Count -eq 0) { Add-Warn $d 'warn_no_planfile' @($script:PlanFolder) }
            }
            elseif ($pt.Error)   { Add-Warn $d 'warn_plan_read' @($pt.Error) }
            else {
                $d.PlanFileName = $pt.File; $d.PlanWeek = $pt.WeekNo
                if (-not $pt.Covered) {
                    Add-Warn $d 'warn_day_not_in_plan' @($prodDate.ToString('dd/MM'), $shiftNo, $pt.File, $pt.WeekNo, $pt.Period)
                }
                elseif ($pt.Total -le 0) {
                    Add-Warn $d 'warn_no_plan_dayshift' @($pt.File)
                }
                else {
                    $effTarget  = [double]$pt.Total
                    $planPerSku = $pt.PerSku; $planDesc = $pt.Desc; $d.PlanSku = $pt.FallbackSku
                    $d.TargetMode = if ($pt.FallbackSku) { 'planfallback' } else { 'plan' }
                    $d.PlanDate = $prodDate
                }
            }
        }
        elseif ($script:HasTarget) { $d.TargetMode = 'param' }
        elseif ($NoPlan)           { $d.TargetMode = 'noplan' }

        $d.Target = $effTarget
        $targetPerMin = if ($effTarget -gt 0) { $effTarget / $shiftMin } else { 0 }
        $d.TargetPerMin = $targetPerMin
        $d.Pct = if ($effTarget -gt 0) { 100.0 * $tot / $effTarget } else { 0 }

        $rowsOut = @()
        foreach ($p in ($counts.Keys | Sort-Object { $counts[$_] } -Descending)) {
            $pl = 0.0; if ($planPerSku.ContainsKey($p)) { $pl = [double]$planPerSku[$p] }
            $ds = '';  if ($planDesc.ContainsKey($p))   { $ds = [string]$planDesc[$p] }
            $rowsOut += [pscustomobject]@{ Product = $p; Count = [int]$counts[$p]; IsMain = ($p -eq $mainProd); Plan = $pl; Desc = $ds
                                           WeekPlan = 0.0; WeekMadeBefore = 0; WeekDone = $false }
        }
        # smaken die deze ploeg wel gepland staan maar nog niet gestart zijn: achteraan, 0 dozen
        # (alleen als er al iets gedraaid is - anders blijft de 'lege ploeg'-melding staan)
        if ($counts.Count -gt 0) {
            foreach ($p in ($planPerSku.Keys | Where-Object { -not $counts.ContainsKey($_) } | Sort-Object { [double]$planPerSku[$_] } -Descending)) {
                $pl = [double]$planPerSku[$p]
                if ($pl -le 0) { continue }
                $ds = ''; if ($planDesc.ContainsKey($p)) { $ds = [string]$planDesc[$p] }
                $rowsOut += [pscustomobject]@{ Product = $p; Count = 0; IsMain = $false; Plan = $pl; Desc = $ds
                                               WeekPlan = 0.0; WeekMadeBefore = 0; WeekDone = $false }
            }
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
                    $stopList += [pscustomobject]@{ From = $prevTs; To = $sTs[$k]; Min = $gap; Kind = $kind; IsLongest = $false }
                }
                $prevTs = $sTs[$k]
            }
            # ---- STAAT DE LIJN NU ECHT STIL, of loopt alleen de OPHALING achter? ----
            # De cache wordt met vertraging weggeschreven (gemeten: 4 a 6 min bij lijnen die gewoon
            # draaiden), dus 'laatste doos ouder dan nu - 2 min' betekende in de praktijk bij ELKE
            # render 'stil'. Maatstaf is daarom de jongste doos in het HELE bestand (alle lijnen
            # samen, zie Get-CollectEnd): printten andere lijnen intussen wel door, dan is de stilte
            # van DEZE lijn echt. Lukt die meting niet, dan geldt de oude regel.
            $tail       = ($refNow - $prevTs).TotalMinutes
            $collectEnd = Get-CollectEnd $sheets $refNow
            $lineLag    = if ($null -ne $collectEnd) { ($collectEnd - $prevTs).TotalMinutes } else { $tail }
            if ($tail -gt $StopMinutes -and $lineLag -gt $StopMinutes) {
                $stopList += [pscustomobject]@{ From = $prevTs; To = $refNow; Min = $tail; Kind = 'nu'; IsLongest = $false }
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
                $lg.IsLongest  = $true          # markering: de lijst zelf blijft op TIJD gesorteerd
                $d.LongestMin  = $lg.Min
                $d.LongestText = ('{0}-{1}' -f $lg.From.ToString('HH:mm'), $lg.To.ToString('HH:mm'))
            }

            # ---- tempo PER PRODUCTTYPE: dozen van dat type / minuten dat dat type effectief liep ----
            # Een type kan meerdere blokken draaien (omstellen heen en terug), dus per blok de tijd van
            # de eerste tot de laatste doos van dat blok optellen; voor het LAATSTE blok tot 'nu'
            # (net als het tempo van de lopende run). Omsteltijd hoort zo bij geen van beide types.
            $tempo = @{}
            $bStart = 0
            for ($k = 1; $k -le $sProd.Count; $k++) {
                if ($k -lt $sProd.Count -and $sProd[$k] -eq $sProd[$bStart]) { continue }
                $p = $sProd[$bStart]
                $endTs = if ($k -ge $sProd.Count) { $refNow } else { $sTs[$k - 1] }
                $blkMin = ($endTs - $sTs[$bStart]).TotalMinutes
                if ($blkMin -lt 0) { $blkMin = 0 }
                if (-not $tempo.ContainsKey($p)) {
                    $tempo[$p] = [pscustomobject]@{ Product = $p; Count = 0; Min = 0.0; Blocks = 0; PerMin = 0.0; HasRate = $false }
                }
                $tempo[$p].Count  += ($k - $bStart)
                $tempo[$p].Min    += $blkMin
                $tempo[$p].Blocks += 1
                $bStart = $k
            }
            foreach ($p in @($tempo.Keys)) {
                $t = $tempo[$p]
                if ($t.Min -ge 1) { $t.PerMin = $t.Count / $t.Min; $t.HasRate = $true }
            }
            $d.Tempo = $tempo

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
                    if ($todo -le 0)      { $d.EtaText = 'target al gehaald'; $d.EtaKind = 'done' }
                    elseif ($remain -le 0){ $d.EtaText = 'ploeg voorbij'; $d.EtaKind = 'over' }
                    else {
                        if ($perMin -gt 0) {
                            $eta = $refNow.AddMinutes($todo / $perMin)
                            $d.EtaTimeText = $eta.ToString('HH:mm')
                            $d.EtaText = $eta.ToString('HH:mm')
                            $d.EtaKind = 'time'
                            if ($eta -gt $win.End) { $d.EtaText += ' (na ploegeinde)'; $d.EtaAfterShift = $true }
                        }
                        else { $d.EtaText = 'niet haalbaar (tempo 0)'; $d.EtaKind = 'impossible' }
                    }
                }
            }
        }

        # ---------------- HISTORIE (dozen per dag/ploeg, uit hetzelfde bestand) ----------------
        # Het *RCDB-blad houdt 31 dagen per UUR bij; dat is precies waar SAPSTATus zijn
        # dag/ploeg-raster mee vult (macro dozen4). Zelfde werkmap, dus geen extra Excel-opening.
        if ($HistoryDays -ge 0 -and $null -ne $histVals) {
            try {
                # HistoryDays = 0 -> de LOPENDE productieweek, die (net als in SAPSTATus) op
                # ZONDAG begint; > 0 -> een rollend venster van zoveel dagen.
                # De tabel bevat 31 dagen, dus lezen we ALLES in een keer en toont de pagina de
                # oudere weken pas na een klik op '+7 dagen' (geen tweede Excel-lees nodig).
                $prodDay   = $win.Start.Date
                $weekStart = $prodDay.AddDays(-[int]$prodDay.DayOfWeek)
                $histFrom  = if ($HistoryDays -gt 0) { $prodDay.AddDays(-($HistoryDays - 1)) } else { $weekStart }
                $oldest    = if ($HistoryDays -gt 0) { $histFrom } else { $weekStart.AddDays(-7 * $script:HistExtraWeeks) }
                $d.HistFrom = $histFrom; $d.HistTo = $prodDay; $d.HistWeekStart = $weekStart
                # roosterlijst zoals lijn 11: de komma houdt het 2D-rooster EEN element
                $histGrids = @(); if ($null -ne $histVals) { $histGrids += ,$histVals }
                $d.History = @(Build-History $histGrids $prodDay $oldest $weekStart $planGrids $planScheds)
                $d.HasHistory = ($d.History.Count -gt 0)
                $mw = 0; foreach ($h in $d.History) { if ($h.WeekIdx -gt $mw) { $mw = $h.WeekIdx } }
                $d.HistMaxWeek = $mw
            }
            catch { Add-Warn $d 'warn_hist_read' @($_.Exception.Message) }
        }

        # ---------------- WEEKPLAN PER SMAAK ----------------
        # Het DAGPLAN uit 'daily shift NDwk*' blijft leidend (dat is de ploegtarget hierboven);
        # dit is de laag eronder: per smaak het plan van de HELE productieweek en wat er al van
        # gemaakt is. Verklaart waarom een omstelling soms niet komt: eerst het weekplan van de
        # lopende smaak op 100% afwerken, desnoods pas in een volgende ploeg.
        # Bronnen: plan = weekplan; HUIDIGE ploeg = Boxruw-blad (doos per doos, actueelst);
        # EERDERE dagen = *RCDB-blad. LET OP: in dat blad ontbreken soms uren van een dag
        # (bron-bug in de Historian-ophaling) - 'gemaakt' staat dan te laag.
        if ($d.HasHistory -and $planGrids.Count -gt 0 -and $d.HistWeekStart) {
            try {
                $wkStart = $d.HistWeekStart
                $wp = $null
                $gl = @($planGrids); $sl = @($planScheds)
                for ($gi = 0; $gi -lt $gl.Count; $gi++) {
                    # lijnfilter voor de HELE week: alles wat op dit platform gepland staat
                    # van zondag 05:00 tot de zondag erop (zie Read-LineSchedule)
                    $flt = $lineProds
                    if ($null -ne $script:Platform -and $gi -lt $sl.Count) {
                        $ex = Get-PlatformSkus $sl[$gi] $script:Platform $wkStart.Date.AddHours(5) $wkStart.Date.AddDays(7).AddHours(5)
                        if ($ex.Count -gt 0) { $flt = $ex }
                    }
                    $cand = Get-PlanForWeek $gl[$gi] $wkStart $flt
                    if ($cand.PerSku.Count -gt 0) { $wp = $cand; break }
                }
                # gemaakt deze week: eerdere ploegen uit het RCDB-blad ...
                $made = @{}
                foreach ($h in $d.History) {
                    if ($h.WeekIdx -ne 0) { continue }
                    if ($h.Date -eq $prodDate -and $h.ShiftNo -eq $shiftNo) { continue }   # huidige ploeg: zie hieronder
                    foreach ($s in $h.Skus) {
                        if ($made.ContainsKey($s.Sku)) { $made[$s.Sku] += [int]$s.Count } else { $made[$s.Sku] = [int]$s.Count }
                    }
                }
                # ... plus de LOPENDE ploeg uit het Boxruw-blad (fijner en actueler dan het uurraster)
                foreach ($p in $counts.Keys) {
                    if (-not (Is-Sku $p)) { continue }
                    if ($made.ContainsKey($p)) { $made[$p] += [int]$counts[$p] } else { $made[$p] = [int]$counts[$p] }
                }

                $keys = @{}
                if ($wp) { foreach ($k in $wp.PerSku.Keys) { $keys[$k] = $true } }
                foreach ($k in $made.Keys) { $keys[$k] = $true }
                $wrows = @()
                foreach ($k in ($keys.Keys | Sort-Object { if ($wp -and $wp.PerSku.ContainsKey($_)) { [double]$wp.PerSku[$_] } else { 0.0 } } -Descending)) {
                    $pl = 0.0; if ($wp -and $wp.PerSku.ContainsKey($k)) { $pl = [double]$wp.PerSku[$k] }
                    $mk = 0;   if ($made.ContainsKey($k))               { $mk = [int]$made[$k] }
                    if ($pl -le 0 -and $mk -le 0) { continue }
                    $ds = ''
                    if     ($wp -and $wp.Desc.ContainsKey($k)) { $ds = [string]$wp.Desc[$k] }
                    elseif ($planDesc.ContainsKey($k))         { $ds = [string]$planDesc[$k] }
                    $wrows += [pscustomobject]@{ Sku = $k; Desc = $ds; Plan = $pl; Made = $mk; IsCur = ($k -eq $mainProd)
                                                 NextAt = $null; NextPast = $false; LastMade = $null; OrderGroup = 0; OrderKey = 0.0; OrderKey2 = 0.0 }
                }
                # ---- VOLGORDE, net als de producttabel van de machinelijnen (vraag gebruiker 17/09/2026) ----
                # nu -> wat volgens het weekrooster nog komt (op gepland begin) -> nog niet gemaakt ->
                # al gemaakt. 'Nu' is hier het product van de laatste doos (lijn 9 draait een smaak tegelijk).
                if ($wrows.Count -gt 0) {
                    $wkEnd  = $wkStart.Date.AddDays(7).AddHours(5)       # productieweek: zondag 05:00 -> zondag 05:00
                    $pSched = $null; $pGrid = $null
                    if ($null -ne $pt) { $pSched = $pt.Sched; $pGrid = $pt.Grid }
                    $nextAt   = Get-NextPlannedStart $pSched $pGrid $script:Platform @($wrows | ForEach-Object { $_.Sku }) $nowDt $win $wkEnd $planPerSku $counts
                    $lastMade = Get-LastMadeMap $d.History $counts $win.Start
                    foreach ($w in $wrows) {
                        $sc = 0.0; if ($counts.ContainsKey($w.Sku))     { $sc = [double]$counts[$w.Sku] }
                        $sp = 0.0; if ($planPerSku.ContainsKey($w.Sku)) { $sp = [double]$planPerSku[$w.Sku] }
                        $o = Get-SkuOrder $w.Sku $w.IsCur $nextAt $lastMade $sc $sp $w.Made $w.Plan $win.Start
                        $w.NextAt = $o.NextAt; $w.LastMade = $o.LastMade
                        # gepland begin dat AL voorbij is = de run loopt, alleen het ploegplan is nog niet gehaald
                        $w.NextPast = ($null -ne $o.NextAt -and $o.NextAt -le $nowDt)
                        $w.OrderGroup = $o.Group; $w.OrderKey = $o.Key; $w.OrderKey2 = $o.Key2
                    }
                    $wrows = @($wrows | Sort-Object OrderGroup, OrderKey, OrderKey2, Sku)
                }
                if ($wrows.Count -gt 0) {
                    $tp = 0.0; $tm = 0
                    foreach ($w in $wrows) { $tp += [double]$w.Plan; $tm += [int]$w.Made }
                    $d.WeekRows = @($wrows); $d.WeekPlanTotal = $tp; $d.WeekMadeTotal = $tm
                    $d.WeekNo = Get-IsoWeek $wkStart.AddDays(1)      # productieweek start zondag -> maandag bepaalt het ISO-nummer
                    $d.HasWeekPlan = $true

                    # ---- WEEKPLAN AL ROND VOOR DEZE PLOEG ----
                    # Het weekplan gaat voor. Was het weekplan van een smaak bij PLOEGSTART al
                    # gehaald, dan heeft het geen zin haar deze ploeg nog te draaien - ook al
                    # staat ze in het dagplan. Het DAGPLAN IN DOZEN verandert daar NIET van: de
                    # ploeg moet evenveel dozen maken, alleen in een andere smaak. Hier merken we
                    # zulke smaken; de uitleg komt onder de producttabel.
                    # 'gemaakt bij ploegstart' = weektotaal min wat DEZE ploeg er zelf van maakte.
                    $wmap = @{}
                    foreach ($w in $wrows) { $wmap[$w.Sku] = $w }
                    $notes = @()
                    foreach ($r in $d.Rows) {
                        if ($r.Plan -le 0 -or -not $wmap.ContainsKey($r.Product)) { continue }
                        $wpl = [double]$wmap[$r.Product].Plan
                        $mb  = [int]$wmap[$r.Product].Made - [int]$r.Count
                        if ($mb -lt 0) { $mb = 0 }
                        $r.WeekPlan = $wpl; $r.WeekMadeBefore = $mb
                        if ($wpl -gt 0 -and $mb -ge $wpl) {
                            $r.WeekDone = $true
                            $left = [double]$r.Plan - [double]$r.Count
                            if ($left -lt 0) { $left = 0 }
                            $notes += [pscustomobject]@{ Sku = $r.Product; MadeBefore = $mb; WeekPlan = $wpl; Left = $left }
                        }
                    }
                    $d.WeekDoneNotes = @($notes)
                }
            }
            catch { Add-Warn $d 'warn_weekplan' @($_.Exception.Message) }
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

# ==================== MOTOR LIJN 11 (acht machines naast elkaar) ====================
function Get-BoxData11 {
    $nowDt = if ($script:HasNow) { $Now } else { Get-Date }
    $win = Get-ShiftWindow $nowDt
    $shiftMin = [int][math]::Round(($win.End - $win.Start).TotalMinutes)
    if ($shiftMin -le 0) { $shiftMin = 480 }
    $minutes = New-Object 'int[]' $shiftMin
    $targetPerMin = if ($ShiftTarget -gt 0) { [double]$ShiftTarget / $shiftMin } else { 0 }

    $d = [ordered]@{
        Ok = $true; Error = $null
        NowText = $nowDt.ToString('dd/MM/yyyy HH:mm')
        ShiftRange = $win.Label
        ShiftLetter = (Get-ShiftLetter $win.Start.Date ([int]$win.Code))
        WindowText = ('{0} -> {1}' -f $win.Start.ToString('dd/MM HH:mm'), $win.End.ToString('dd/MM HH:mm'))
        Sheet = $BoxSheet; BoxFile = ''; FileTimeText = '-'
        TargetMode = 'unknown'; PlanDate = $null; WarnList = @(); PlanFileName = ''; PlanWeek = $null
        PlanSku = $null; ShiftNo = 0
        Rows = @(); Tempo = @{}; Total = 0; ParsedRows = 0; LastText = ''; StartRowSkipped = $false
        BlankRows = 0                      # dozen zonder etiketstring: overgeslagen, niet geteld
        # losse onderdelen van 'laatste doos' - de zin zelf wordt PAS in Render-Html gezet (taal!)
        LastTimeText = ''; LastProduct = ''; LastCounter = ''
        ShiftStart = $win.Start; ShiftEnd = $win.End; ShiftMin = $shiftMin; Minutes = $minutes; MaxPerMin = 0
        Target = $ShiftTarget; TargetPerMin = $targetPerMin; Pct = 0
        # --- prognose einde ploeg (LIJN 11: som van de machines, geen enkel 'hoofdproduct') ---
        HasForecast = $false; FcWeak = $false
        RefNowText = '-'; ElapsedMin = 0; RemainMin = 0; NowOffsetMin = 0
        PerMin = 0; PerHour = 0
        # --- machines (tag 8): het hart van de lijn-11-versie ---
        HasMach = $false; Machines = @(); MachTotal = 0; MachRunning = 0; MachStopped = 0
        MachStopMin = 0; MachStopCount = 0; MachElapsedMin = 0
        RecentWin = 0; RecentCount = 0; HasRecent = $false; RecentPerMin = 0; ProjRecent = 0
        ProjTotal = 0; ProjPct = 0; ProjDiff = 0
        EtaText = ''; EtaKind = ''; EtaTimeText = ''; EtaAfterShift = $false
        # --- stilstand + achterstand ---
        HasStops = $false; Stops = @(); StopCount = 0; StopMin = 0
        LongestMin = 0; LongestText = ''; NowStill = $false; StillMin = 0
        ElapsedShiftMin = 0; RunMin = 0; AvailPct = 0; NetPerMin = 0
        BehindNow = 0; BehindEnd = 0; StopLimit = $StopMinutes
        # --- historie per dag/ploeg (uit het RCDB-blad) ---
        HasHistory = $false; History = @(); HistSheet = $script:RcdbSheet
        HistFrom = $null; HistTo = $null; HistWeekStart = $null; HistMaxWeek = 0
        # --- weekplan per smaak (dagplan blijft leidend; dit is de laag eronder) ---
        HasWeekPlan = $false; WeekRows = @(); WeekPlanTotal = 0.0; WeekMadeTotal = 0; WeekNo = $null
        # smaken waarvan het WEEKPLAN bij ploegstart al rond was (zie 'weekplan al rond' verderop)
        WeekDoneNotes = @()
        # --- een regel per smaak: week EN ploeg samen (voedt de hoofdtabel van de pagina) ---
        HasCombined = $false; Combined = @(); SkuNames = @{}
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

        $cutoff = if ($script:HasNow) { $nowDt } else { [datetime]::MaxValue }

        # LIJN 11 heeft TWEE doosprinters -> twee bladen. Ze worden achter elkaar gelezen en
        # daarna OP TIJD gesorteerd tot een enkele dozenstroom: welke printer een doos etiketteert
        # zegt niets over de machine (tag 8) die hem gemaakt heeft - machine 1106 duikt zelfs in
        # beide bladen op. Per doos houden we tijd + smaak (tag 14) + machine (tag 8) bij.
        $sTs   = New-Object 'System.Collections.Generic.List[datetime]'
        $sProd = New-Object 'System.Collections.Generic.List[string]'
        $sMach = New-Object 'System.Collections.Generic.List[string]'

        $counts = @{}; $allProds = @{}; $parsed = 0; $lastTs = $null; $lastProd = $null; $lastCtr = $null
        $rawRows = New-Object 'System.Collections.Generic.List[object]'
        $sheetsRead = 0

        foreach ($sheetName in $script:BoxSheets) {
            $ws = $null
            foreach ($s in $sheets) { if ($s.Name -eq $sheetName) { $ws = $s; break } }
            if ($null -eq $ws) { Add-Warn $d 'warn_no_boxsheet' @($sheetName, $d.BoxFile); continue }

            # laatste gevulde rij in kolom B (om het leesbereik te begrenzen)
            $allCells = $ws.Cells
            $anchor   = $allCells.Item($xlMaxRows, 2)
            $lastCell = $anchor.End($xlUp)
            $lastRow  = [int]$lastCell.Row
            Rel $lastCell; Rel $anchor; Rel $allCells

            if ($lastRow -lt 11) { Add-Warn $d 'warn_no_data_row11' @($sheetName); Rel $ws; $ws = $null; continue }
            $sheetsRead++

            # A=tijd, B=etiket, C=lijn ; alles in EEN marshaling-call (leest cache, geen herberekening)
            $rng  = $ws.Range("A11:C$lastRow").Value2
            $rmax = $rng.GetUpperBound(0)

            for ($i = 1; $i -le $rmax; $i++) {
                $b = $rng.GetValue($i, 2)
                # LEGE B = EEN DOOS ZONDER ETIKET, GEEN EINDE VAN DE GEGEVENS - zie de uitleg bij
                # dezelfde regel in Get-BoxData9. Vroeger stopte het lezen hier ('break'), waardoor
                # een enkele etiketloze doos de rest van de ploeg onzichtbaar maakte.
                if ($null -eq $b -or ([string]$b).Trim() -eq '') { $d.BlankRows++; continue }
                $parsed++
                $label = [string]$b
                $prod  = Get-ProductType $label
                if ([string]::IsNullOrWhiteSpace($prod)) { $prod = '(onbekend)' }
                $mach  = Get-Machine $label
                if ([string]::IsNullOrWhiteSpace($mach)) { $mach = '(onbekend)' }

                $tsRaw = $rng.GetValue($i, 1); $ts = $null
                if ($tsRaw -is [double]) { try { $ts = [DateTime]::FromOADate([double]$tsRaw) } catch {} }

                # De EERSTE rij van ELK blad is de startwaarde van de DELTA-ophaling (stand op het
                # begin van het ophaalvenster, exact op het hele uur) - die doos is EERDER geprint,
                # dus niet meetellen. Beide printerbladen hebben zo'n rij.
                if ($i -eq 1 -and $null -ne $ts -and $ts.Minute -eq 0 -and $ts.Second -eq 0 -and $ts.Millisecond -eq 0) {
                    $d.StartRowSkipped = $true
                    continue
                }

                # alle producten van het blad (= wat op deze lijn draait) - filter voor het fabriekbrede plan
                if (-not $allProds.ContainsKey($prod)) { $allProds[$prod] = $true }

                # tellen alleen als het tijdstip in de HUIDIGE ploeg-interval valt
                # (met -Now ook niet verder tellen dan dat gesimuleerde moment: tijdmachine)
                if ($null -ne $ts -and $ts -ge $win.Start -and $ts -lt $win.End -and $ts -le $cutoff) {
                    $rawRows.Add([pscustomobject]@{ Ts = $ts; Prod = $prod; Mach = $mach; Ctr = (Get-Counter $label) })
                }
            }
            $rng = $null
            Rel $ws; $ws = $null
        }
        if ($sheetsRead -eq 0) { throw ("Geen enkel doosblad ({0}) gevonden in {1}." -f ($script:BoxSheets -join ', '), $d.BoxFile) }
        $d.ParsedRows = $parsed

        # de twee printerstromen tot EEN chronologische reeks samenvoegen
        foreach ($r in ($rawRows | Sort-Object Ts)) {
            if ($counts.ContainsKey($r.Prod)) { $counts[$r.Prod]++ } else { $counts[$r.Prod] = 1 }
            $off = [int][math]::Floor(($r.Ts - $win.Start).TotalMinutes)
            if ($off -ge 0 -and $off -lt $shiftMin) { $minutes[$off]++ }
            $sTs.Add($r.Ts); $sProd.Add($r.Prod); $sMach.Add($r.Mach)
            $lastTs = $r.Ts; $lastProd = $r.Prod; $lastCtr = $r.Ctr
        }

        # GEEN 'hoofdproduct' op lijn 11: er lopen standaard vier tot zes smaken tegelijk, elk op
        # zijn eigen machine. Wat er 'nu loopt' wordt verderop PER MACHINE bepaald.
        $tot = 0; foreach ($p in $counts.Keys) { $tot += [int]$counts[$p] }
        $d.Total = $tot
        $mx = 0; foreach ($v in $minutes) { if ($v -gt $mx) { $mx = $v } }
        $d.MaxPerMin = $mx
        if ($lastTs) {
            $d.LastTimeText = $lastTs.ToString('HH:mm'); $d.LastProduct = [string]$lastProd; $d.LastCounter = [string]$lastCtr
            $d.LastText = ('{0} - {1} (doos {2})' -f $d.LastTimeText, $d.LastProduct, $d.LastCounter)   # nl-terugval
        }

        # ---- producten van DEZE lijn: filter voor het fabriekbrede weekplan ----
        # Het *RCDB-blad wordt hier AL gelezen (en verderop hergebruikt voor de historie), want
        # zonder die 31-daagse lijst telt een smaak die deze ploeg gepland staat maar nog NIET
        # gestart is niet mee in de target: do 30/07 gaf 540 (alleen 340062773, het draaiende
        # product) i.p.v. 1.157, omdat de omstelling naar 340056956 (617) er niet gekomen was.
        # LIJN 11 heeft TWEE van die tabellen (L11P1RCDB + L11P2RCDB) - beide lezen en optellen.
        $histGrids = @()
        $lineProds = @{}
        foreach ($p in $allProds.Keys) { $lineProds[$p] = $true }
        foreach ($rcdbName in $script:RcdbSheets) {
            $ws2 = $null
            foreach ($s in $sheets) { if ($s.Name -eq $rcdbName) { $ws2 = $s; break } }
            if ($null -eq $ws2) { Add-Warn $d 'warn_no_rcdb' @($rcdbName); continue }
            try {
                $g = $ws2.Range("A1:FI300").Value2             # 31 blokken van 5 kolommen = t/m kolom FI
                $histGrids += ,$g
                foreach ($p in (Get-RcdbSkus $g).Keys) { $lineProds[$p] = $true }
            }
            catch { Add-Warn $d 'warn_hist_read' @($_.Exception.Message) }
            Rel $ws2; $ws2 = $null
        }

        # ---------------- TARGET UIT HET WEEKPLAN ----------------
        $shiftNo = [int]$win.Code; $prodDate = $win.Start.Date
        $d.ShiftNo = $shiftNo
        $effTarget = [double]$ShiftTarget
        $planPerSku = @{}; $planDesc = @{}; $planGrids = @(); $planScheds = @(); $skuNames = @{}
        if (-not $NoPlan -and -not $script:HasTarget) {
            $cands = @()
            if ($script:HasPlanFile) {
                if (Test-Path -LiteralPath $PlanFile) { $cands = @([pscustomobject]@{ File = (Resolve-Path -LiteralPath $PlanFile).Path }) }
                else { Add-Warn $d 'warn_planfile_missing' @($PlanFile) }
            }
            else { $cands = @(Get-PlanCandidates $script:PlanFolder | Select-Object -First 3) }

            $pt = $null; $read = 0
            foreach ($c in $cands) {
                $try = Read-PlanTargets $excel $c.File $prodDate $shiftNo (@($counts.Keys)) $lineProds
                foreach ($kn in $try.SkuNames.Keys) { if (-not $skuNames.ContainsKey($kn)) { $skuNames[$kn] = $try.SkuNames[$kn] } }
                $read++
                if ($try.Grid) { $planGrids += ,$try.Grid; $planScheds += ,$try.Sched }   # hergebruik voor de historie
                if ($null -eq $pt -or $try.Covered) { $pt = $try }
                if ($try.Covered) { break }
            }
            # planbestanden van OUDERE weken (voor de historie-knop '+7 dagen'); alleen als ze bestaan
            if ($HistoryDays -ge 0 -and $cands.Count -gt $read) {
                foreach ($c in ($cands | Select-Object -Skip $read)) {
                    $g = Read-PlanTargets $excel $c.File $prodDate $shiftNo (@($counts.Keys)) $lineProds
                    foreach ($kn in $g.SkuNames.Keys) { if (-not $skuNames.ContainsKey($kn)) { $skuNames[$kn] = $g.SkuNames[$kn] } }
                    if ($g.Grid) { $planGrids += ,$g.Grid; $planScheds += ,$g.Sched }
                }
            }
            if ($null -eq $pt) {
                if ($cands.Count -eq 0) { Add-Warn $d 'warn_no_planfile' @($script:PlanFolder) }
            }
            elseif ($pt.Error)   { Add-Warn $d 'warn_plan_read' @($pt.Error) }
            else {
                $d.PlanFileName = $pt.File; $d.PlanWeek = $pt.WeekNo
                if (-not $pt.Covered) {
                    Add-Warn $d 'warn_day_not_in_plan' @($prodDate.ToString('dd/MM'), $shiftNo, $pt.File, $pt.WeekNo, $pt.Period)
                }
                elseif ($pt.Total -le 0) {
                    Add-Warn $d 'warn_no_plan_dayshift' @($pt.File)
                }
                else {
                    $effTarget  = [double]$pt.Total
                    $planPerSku = $pt.PerSku; $planDesc = $pt.Desc; $d.PlanSku = $pt.FallbackSku
                    $d.TargetMode = if ($pt.FallbackSku) { 'planfallback' } else { 'plan' }
                    $d.PlanDate = $prodDate
                }
            }
        }
        elseif ($script:HasTarget) { $d.TargetMode = 'param' }
        elseif ($NoPlan)           { $d.TargetMode = 'noplan' }

        $d.Target = $effTarget
        $targetPerMin = if ($effTarget -gt 0) { $effTarget / $shiftMin } else { 0 }
        $d.TargetPerMin = $targetPerMin
        $d.Pct = if ($effTarget -gt 0) { 100.0 * $tot / $effTarget } else { 0 }

        $rowsOut = @()
        foreach ($p in ($counts.Keys | Sort-Object { $counts[$_] } -Descending)) {
            $pl = 0.0; if ($planPerSku.ContainsKey($p)) { $pl = [double]$planPerSku[$p] }
            $ds = '';  if ($planDesc.ContainsKey($p))   { $ds = [string]$planDesc[$p] }
            # Proj/PerMin/Machines/HasProj worden verderop PER SMAAK ingevuld door de machine-motor
            $rowsOut += [pscustomobject]@{ Product = $p; Count = [int]$counts[$p]; IsMain = $false; Plan = $pl; Desc = $ds
                                           Proj = 0.0; HasProj = $false; PerMin = 0.0; Machines = @() }
        }
        # smaken die deze ploeg wel gepland staan maar nog niet gestart zijn: achteraan, 0 dozen
        # (alleen als er al iets gedraaid is - anders blijft de 'lege ploeg'-melding staan)
        if ($counts.Count -gt 0) {
            foreach ($p in ($planPerSku.Keys | Where-Object { -not $counts.ContainsKey($_) } | Sort-Object { [double]$planPerSku[$_] } -Descending)) {
                $pl = [double]$planPerSku[$p]
                if ($pl -le 0) { continue }
                $ds = ''; if ($planDesc.ContainsKey($p)) { $ds = [string]$planDesc[$p] }
                $rowsOut += [pscustomobject]@{ Product = $p; Count = 0; IsMain = $false; Plan = $pl; Desc = $ds
                                               Proj = 0.0; HasProj = $false; PerMin = 0.0; Machines = @() }
            }
        }
        $d.Rows = $rowsOut

        # ---------------- MACHINES + PROGNOSE EINDE PLOEG ----------------
        # HIER ZIT HET GROTE VERSCHIL MET LIJN 9. Lijn 11 draait vier tot zes smaken TEGELIJK,
        # elk op een eigen verpakkingsmachine (tag 8 = 1101..1108). Daarom:
        #   * elke machine krijgt zijn EIGEN tempo, stilstanden en prognose;
        #   * de ploegprognose is de SOM van die machines (niet: een enkele lopende run
        #     doortrekken, zoals op lijn 9 waar er maar een smaak tegelijk loopt);
        #   * een smaak die op twee machines loopt telt beide machines bij elkaar op;
        #   * stilstand meet je PER MACHINE - de lijn als geheel staat vrijwel nooit stil
        #     (er draait altijd wel iets), dus lijnbreed meten verbergt het echte verlies.
        if ($sTs.Count -gt 0) {
            # 'nu' voor de data = niet later dan de opslagtijd van het bestand en niet na ploegeinde
            $refNow = $nowDt
            if (Test-Path -LiteralPath $boxFile) {
                $ft = (Get-Item -LiteralPath $boxFile).LastWriteTime
                if ($ft -lt $refNow) { $refNow = $ft }
            }
            if ($refNow -gt $win.End)   { $refNow = $win.End }
            if ($refNow -lt $lastTs)    { $refNow = $lastTs }   # data kan niet ouder zijn dan de laatste doos

            $elapsedShift = ($refNow - $win.Start).TotalMinutes
            $remain       = ($win.End - $refNow).TotalMinutes
            if ($remain -lt 0) { $remain = 0 }

            # DATA-HORIZON: verder dan de LAATSTE doos van de lijn weten we niets. Het bestand
            # wordt met vertraging weggeschreven (snapshot: laatste doos 20:03, bestand 20:08),
            # dus die staart is GEDEELDE achterstand van de ophaling - geen stilstand van een
            # afzonderlijke machine. Een machine heet daarom pas 'stil' als ze meer dan
            # -StopMinutes achterloopt op de horizon, d.w.z. terwijl ANDERE machines wel doorgaan.
            # Ligt de HELE lijn stil, dan vangt de rode balk 'de hele lijn staat stil' dat af.
            $dataEnd = $lastTs
            $machElapsed = ($dataEnd - $win.Start).TotalMinutes
            if ($machElapsed -lt 0) { $machElapsed = 0 }

            # ---- LIJNBREED stil: geen ENKELE machine printte nog een doos ----
            # Dit is bewust iets anders dan op lijn 9: hier betekent het 'de HELE lijn ligt plat'.
            # Zeldzaam, maar juist daarom het signaal dat bovenaan hoort.
            $stopList = @(); $prevTs = $win.Start
            for ($k = 0; $k -lt $sTs.Count; $k++) {
                $gap = ($sTs[$k] - $prevTs).TotalMinutes
                if ($gap -gt $StopMinutes) {
                    $kind = if ($k -eq 0) { 'opstart' } else { 'stop' }
                    $stopList += [pscustomobject]@{ From = $prevTs; To = $sTs[$k]; Min = $gap; Kind = $kind; IsLongest = $false }
                }
                $prevTs = $sTs[$k]
            }
            # ---- STAAT DE LIJN NU ECHT STIL, of loopt alleen de OPHALING achter? ----
            # Zie Get-CollectEnd: de jongste doos in het HELE bestand is het verste dat de ophaling
            # gekomen is. Alleen als DEZE lijn daar meer dan -StopMinutes bij achterblijft ligt ze
            # echt stil; anders is de staart gewoon de gedeelde schrijfvertraging (4 a 6 min).
            $tail       = ($refNow - $prevTs).TotalMinutes
            $collectEnd = Get-CollectEnd $sheets $refNow
            $lineLag    = if ($null -ne $collectEnd) { ($collectEnd - $prevTs).TotalMinutes } else { $tail }
            if ($tail -gt $StopMinutes -and $lineLag -gt $StopMinutes) {
                $stopList += [pscustomobject]@{ From = $prevTs; To = $refNow; Min = $tail; Kind = 'nu'; IsLongest = $false }
                $d.NowStill = $true; $d.StillMin = [Math]::Round($tail)
            }
            $lineStopSum = 0.0; foreach ($s in $stopList) { $lineStopSum += $s.Min }
            $d.HasStops  = $true
            $d.Stops     = $stopList
            $d.StopCount = $stopList.Count
            $d.StopMin   = $lineStopSum
            if ($stopList.Count -gt 0) {
                $lg = $stopList | Sort-Object Min -Descending | Select-Object -First 1
                $lg.IsLongest  = $true          # markering: de lijst zelf blijft op TIJD gesorteerd
                $d.LongestMin  = $lg.Min
                $d.LongestText = ('{0}-{1}' -f $lg.From.ToString('HH:mm'), $lg.To.ToString('HH:mm'))
            }

            # ---- PER MACHINE: dozen groeperen op tag 8 ----
            $machRows = @{}
            for ($k = 0; $k -lt $sTs.Count; $k++) {
                $m = $sMach[$k]
                if (-not $machRows.ContainsKey($m)) { $machRows[$m] = New-Object 'System.Collections.Generic.List[object]' }
                [void]$machRows[$m].Add([pscustomobject]@{ Ts = $sTs[$k]; Prod = $sProd[$k] })
            }
            # machinenummers oplopend; een doos zonder etiketgegevens ('(onbekend)') achteraan
            $machNames = @($machRows.Keys | Sort-Object { if ($_ -match '^\d+$') { '0{0:D8}' -f [int]$_ } else { "z$_" } })

            $machines  = @()
            $projAdd   = @{}      # per smaak: hoeveel dozen er tot ploegeinde nog bij komen
            $prodRate  = @{}      # per smaak: het gecombineerde tempo van zijn machines nu
            $prodMach  = @{}      # per smaak: welke machines er nu op staan
            $sumPerMin = 0.0; $sumStopMin = 0.0; $sumRunMin = 0.0
            $nRunning  = 0;   $nStopped = 0

            foreach ($mName in $machNames) {
                $lst = $machRows[$mName]
                $cnt = $lst.Count
                $mFirst = $lst[0].Ts; $mLast = $lst[$cnt - 1].Ts

                # stilstanden van DEZE machine: gat tussen twee dozen, plus het gat ploegstart ->
                # eerste doos (opstart) en laatste doos -> nu (staat deze machine NU stil).
                $mStops = @(); $prev = $win.Start
                for ($k = 0; $k -lt $cnt; $k++) {
                    $gap = ($lst[$k].Ts - $prev).TotalMinutes
                    if ($gap -gt $StopMinutes) {
                        $kind = if ($k -eq 0) { 'opstart' } else { 'stop' }
                        $mStops += [pscustomobject]@{ From = $prev; To = $lst[$k].Ts; Min = $gap; Kind = $kind }
                    }
                    $prev = $lst[$k].Ts
                }
                # staart = stilte t.o.v. de data-horizon; loopt de machine gewoon door, dan is de
                # laatste (gedeelde) minuten geen stilstand van HAAR maar vertraging van de ophaling.
                # Het verschil met de horizon telt alleen mee zolang de LIJN zelf draaide: lag de
                # lijn in dat venster stil, dan viel er voor deze machine niets te maken. Zonder die
                # correctie bepaalt een enkele late doos van een andere machine het oordeel - 18/09
                # printten 801/804/805 elk nog EEN doos om 20:15 nadat alles om 19:46 was gestopt,
                # waardoor 803 (laatste doos 20:04) 'stil 11 min' kreeg terwijl er lijnbreed niets
                # liep. Met deze correctie: 803 -> 0,1 min (draait), 806 -> 6,4 min (echt gestopt).
                $mLag = ($dataEnd - $prev).TotalMinutes
                foreach ($ls in $stopList) {
                    $lf = if ($ls.From -gt $prev)    { $ls.From } else { $prev }
                    $lt = if ($ls.To   -lt $dataEnd) { $ls.To }   else { $dataEnd }
                    $ov = ($lt - $lf).TotalMinutes
                    if ($ov -gt 0) { $mLag -= $ov }
                }
                if ($mLag -lt 0) { $mLag = 0 }
                $mStill = ($mLag -gt $StopMinutes)
                $mTail  = ($refNow - $prev).TotalMinutes
                if ($mStill) {
                    $mStops += [pscustomobject]@{ From = $prev; To = $refNow; Min = $mTail; Kind = 'nu' }
                }
                # Getoonde stilstand loopt tot NU; de draaitijd rekent tot de data-horizon, anders
                # zou de gedeelde ophaalvertraging bij elke machine apart worden afgetrokken.
                $mStopSum = 0.0; $mStopIn = 0.0
                foreach ($x in $mStops) {
                    $mStopSum += $x.Min
                    $to = if ($x.To -gt $dataEnd) { $dataEnd } else { $x.To }
                    $clip = ($to - $x.From).TotalMinutes
                    if ($clip -gt 0) { $mStopIn += $clip }
                }
                $mRunMin  = $machElapsed - $mStopIn; if ($mRunMin -lt 0) { $mRunMin = 0 }
                $mNet     = if ($mRunMin -ge 1) { $cnt / $mRunMin } else { 0.0 }
                $mLongest = 0.0
                foreach ($x in $mStops) { if ($x.Min -gt $mLongest) { $mLongest = $x.Min } }

                # smaken van deze machine (meestal een; bij een omstelling binnen de ploeg twee)
                $mProdCnt = @{}
                foreach ($x in $lst) { if ($mProdCnt.ContainsKey($x.Prod)) { $mProdCnt[$x.Prod]++ } else { $mProdCnt[$x.Prod] = 1 } }

                # LOPENDE run van deze machine = laatste aaneengesloten blok met dezelfde smaak.
                # Net als op lijn 9, maar dan per machine: zo telt een omstelling niet mee in het
                # tempo, en wordt de resterende tijd aan de JUISTE (huidige) smaak toegerekend.
                $curProd = $lst[$cnt - 1].Prod
                $idx = $cnt - 1
                while ($idx -gt 0 -and $lst[$idx - 1].Prod -eq $curProd) { $idx-- }
                $runStart = $lst[$idx].Ts
                $runCount = $cnt - $idx
                $runEl    = ($refNow - $runStart).TotalMinutes
                # tempo van de lopende run INCLUSIEF haar stilstanden: dat is de eerlijke
                # verwachting voor de rest van de ploeg (stopt de machine vaak, dan zakt het tempo).
                $mPerMin  = if ($runEl -ge 1) { $runCount / $runEl } else { 0.0 }
                # Staat de machine stil, dan maakt ze NU niets: dan telt ze niet mee in 'tempo nu',
                # niet in de prognose van haar smaak en niet in die van de lijn. Anders kreeg een
                # smaak een tempo en een oplopende prognose van een machine die al een half uur
                # zweeg (18/09: 340051174 stond op 828 dozen, prognose 868, tempo 1,89).
                $mEff     = if ($mStill) { 0.0 } else { $mPerMin }
                $mAdd     = $mEff * $remain
                $mProj    = $cnt + $mAdd

                # dozen met een leeg etiket ('(onbekend)') zijn echte dozen, maar geen machine:
                # ze tellen wel mee in de aantallen, niet in 'hoeveel machines draaien'.
                # Ligt de HELE lijn stil, dan draait er niets - ook niet de machines die tot de
                # horizon meeliepen. Anders meldde de kop '4 draaien' terwijl de rode strook
                # erboven zei dat de lijn stilstond.
                if ($mName -match '^\d+$') { if ($mStill -or $d.NowStill) { $nStopped++ } else { $nRunning++ } }
                $sumPerMin  += $mEff
                $sumStopMin += $mStopSum
                $sumRunMin  += $mRunMin

                if ($projAdd.ContainsKey($curProd))  { $projAdd[$curProd]  += $mAdd }   else { $projAdd[$curProd]  = $mAdd }
                if ($prodRate.ContainsKey($curProd)) { $prodRate[$curProd] += $mEff } else { $prodRate[$curProd] = $mEff }
                if (-not $prodMach.ContainsKey($curProd)) { $prodMach[$curProd] = @() }
                $prodMach[$curProd] += $mName

                $machines += [pscustomobject]@{
                    Machine    = $mName
                    Count      = $cnt
                    Product    = $curProd
                    ProdCount  = $mProdCnt.Count
                    ProdList   = @($mProdCnt.Keys | Sort-Object { $mProdCnt[$_] } -Descending |
                                   ForEach-Object { [pscustomobject]@{ Sku = $_; Count = [int]$mProdCnt[$_] } })
                    First      = $mFirst; Last = $mLast
                    FirstText  = $mFirst.ToString('HH:mm'); LastText = $mLast.ToString('HH:mm')
                    Stops      = $mStops
                    StopCount  = $mStops.Count
                    StopMin    = $mStopSum
                    LongestMin = $mLongest
                    RunMin     = $mRunMin
                    AvailPct   = if ($machElapsed -gt 0) { 100.0 * $mRunMin / $machElapsed } else { 0.0 }
                    IsReal     = ($mName -match '^\d+$')
                    NetPerMin  = $mNet
                    PerMin     = $mEff      # 0 zodra de machine stilstaat; haar netto tempo staat in de tooltip
                    RunCount   = $runCount
                    RunStartText = $runStart.ToString('HH:mm')
                    IsStill    = $mStill
                    StillMin   = [Math]::Round($mTail)
                    Proj       = $mProj
                }
            }

            $d.HasMach        = ($machines.Count -gt 0)
            $d.Machines       = $machines
            $d.MachTotal      = @($machines | Where-Object { $_.IsReal }).Count
            $d.MachRunning    = $nRunning
            $d.MachStopped    = $nStopped
            $d.MachStopMin    = $sumStopMin
            $d.MachStopCount  = 0
            foreach ($mm in $machines) { $d.MachStopCount += [int]$mm.StopCount }
            $d.MachElapsedMin = $machElapsed

            # ---- lijncijfers ----
            # Draaitijd telt in MACHINE-minuten (som over de machines); netto tempo blijft in
            # dozen/min van de LIJN, zodat het rechtstreeks met het targettempo te vergelijken is
            # (een som van machinetempo's zou een fantasiegetal geven: de meeste machines staan
            # een groot deel van de ploeg bewust stil, elke smaak heeft zijn eigen vraag).
            $lineRunMin        = $elapsedShift - $lineStopSum; if ($lineRunMin -lt 1) { $lineRunMin = $elapsedShift }
            $d.ElapsedShiftMin = $d.MachTotal * $machElapsed
            $d.RunMin          = $sumRunMin
            $d.AvailPct        = if ($d.ElapsedShiftMin -gt 0) { 100.0 * $sumRunMin / $d.ElapsedShiftMin } else { 0 }
            $d.NetPerMin       = if ($lineRunMin -gt 0) { $tot / $lineRunMin } else { 0 }
            $d.BehindNow       = $tot - $targetPerMin * $elapsedShift

            # ---- tempo per smaak = som van de machines die die smaak NU draaien ----
            $tempo = @{}
            foreach ($p in $prodRate.Keys) {
                $c = 0; if ($counts.ContainsKey($p)) { $c = [int]$counts[$p] }
                $tempo[$p] = [pscustomobject]@{
                    Product = $p; Count = $c; PerMin = [double]$prodRate[$p]
                    Machines = @($prodMach[$p]); MachCount = @($prodMach[$p]).Count
                    HasRate = ([double]$prodRate[$p] -gt 0)
                }
            }
            $d.Tempo = $tempo

            # ---- prognose ploeg = wat er staat + wat de machines er nog bij maken ----
            if ($elapsedShift -ge 1) {
                $d.HasForecast  = $true
                $d.FcWeak       = ($elapsedShift -lt 5 -or $tot -lt 5)
                if ($d.FcWeak) { Add-Warn $d 'weak_run_l11' @($tot, [Math]::Round($elapsedShift)) }
                $d.RefNowText   = $refNow.ToString('dd/MM HH:mm')
                $d.ElapsedMin   = [Math]::Round($elapsedShift)
                $d.RemainMin    = [Math]::Round($remain)
                $d.NowOffsetMin = $elapsedShift
                $d.PerMin       = $sumPerMin
                $d.PerHour      = $sumPerMin * 60
                $d.ProjTotal    = $tot + $sumPerMin * $remain
                $d.BehindEnd    = $d.ProjTotal - $effTarget

                # per smaak: eindstand van de ploeg + markering 'draait nu'
                foreach ($r in $rowsOut) {
                    $add = 0.0; if ($projAdd.ContainsKey($r.Product)) { $add = [double]$projAdd[$r.Product] }
                    $r.Proj     = [double]$r.Count + $add
                    $r.HasProj  = $true
                    $r.PerMin   = if ($prodRate.ContainsKey($r.Product)) { [double]$prodRate[$r.Product] } else { 0.0 }
                    $r.Machines = if ($prodMach.ContainsKey($r.Product)) { @($prodMach[$r.Product]) } else { @() }
                    # 'draait nu' = er staat minstens EEN machine op deze smaak die niet stilstaat.
                    # Ligt de HELE lijn stil, dan draait er per definitie niets en hoort er bij geen
                    # enkele smaak 'nu' te staan. De machinetabel deed dat al ($d.NowStill), de
                    # producttabel niet - daardoor kon een smaak 'nu' krijgen terwijl de rode strook
                    # erboven meldde dat de lijn stilstond.
                    $r.IsMain   = $false
                    if (-not $d.NowStill) {
                        foreach ($mm in $machines) { if ($mm.Product -eq $r.Product -and -not $mm.IsStill) { $r.IsMain = $true; break } }
                    }
                }

                # tweede schatting: tempo van de laatste RecentMinutes minuten (hele lijn)
                $recWin = [Math]::Min([double]$RecentMinutes, $elapsedShift)
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
                    if ($todo -le 0)      { $d.EtaText = 'target al gehaald'; $d.EtaKind = 'done' }
                    elseif ($remain -le 0){ $d.EtaText = 'ploeg voorbij'; $d.EtaKind = 'over' }
                    else {
                        if ($sumPerMin -gt 0) {
                            $eta = $refNow.AddMinutes($todo / $sumPerMin)
                            $d.EtaTimeText = $eta.ToString('HH:mm')
                            $d.EtaText = $eta.ToString('HH:mm')
                            $d.EtaKind = 'time'
                            if ($eta -gt $win.End) { $d.EtaText += ' (na ploegeinde)'; $d.EtaAfterShift = $true }
                        }
                        else { $d.EtaText = 'niet haalbaar (tempo 0)'; $d.EtaKind = 'impossible' }
                    }
                }
            }
        }
        # ---------------- HISTORIE (dozen per dag/ploeg, uit hetzelfde bestand) ----------------
        # Het *RCDB-blad houdt 31 dagen per UUR bij; dat is precies waar SAPSTATus zijn
        # dag/ploeg-raster mee vult (macro dozen4). Zelfde werkmap, dus geen extra Excel-opening.
        if ($HistoryDays -ge 0 -and $histGrids.Count -gt 0) {
            try {
                # HistoryDays = 0 -> de LOPENDE productieweek, die (net als in SAPSTATus) op
                # ZONDAG begint; > 0 -> een rollend venster van zoveel dagen.
                # De tabel bevat 31 dagen, dus lezen we ALLES in een keer en toont de pagina de
                # oudere weken pas na een klik op '+7 dagen' (geen tweede Excel-lees nodig).
                $prodDay   = $win.Start.Date
                $weekStart = $prodDay.AddDays(-[int]$prodDay.DayOfWeek)
                $histFrom  = if ($HistoryDays -gt 0) { $prodDay.AddDays(-($HistoryDays - 1)) } else { $weekStart }
                $oldest    = if ($HistoryDays -gt 0) { $histFrom } else { $weekStart.AddDays(-7 * $script:HistExtraWeeks) }
                $d.HistFrom = $histFrom; $d.HistTo = $prodDay; $d.HistWeekStart = $weekStart
                $d.History = @(Build-History $histGrids $prodDay $oldest $weekStart $planGrids $planScheds)
                $d.HasHistory = ($d.History.Count -gt 0)
                $mw = 0; foreach ($h in $d.History) { if ($h.WeekIdx -gt $mw) { $mw = $h.WeekIdx } }
                $d.HistMaxWeek = $mw
            }
            catch { Add-Warn $d 'warn_hist_read' @($_.Exception.Message) }
        }

        # ---------------- WEEKPLAN PER SMAAK ----------------
        # Het DAGPLAN uit 'daily shift NDwk*' blijft leidend (dat is de ploegtarget hierboven);
        # dit is de laag eronder: per smaak het plan van de HELE productieweek en wat er al van
        # gemaakt is. Verklaart waarom een omstelling soms niet komt: eerst het weekplan van de
        # lopende smaak op 100% afwerken, desnoods pas in een volgende ploeg.
        # Bronnen: plan = weekplan; HUIDIGE ploeg = Boxruw-blad (doos per doos, actueelst);
        # EERDERE dagen = *RCDB-blad. LET OP: in dat blad ontbreken soms uren van een dag
        # (bron-bug in de Historian-ophaling) - 'gemaakt' staat dan te laag.
        # smaken die op DIT moment op minstens een draaiende machine staan (markering '<< nu')
        $activeProds = @{}
        foreach ($r in $d.Rows) { if ($r.IsMain) { $activeProds[$r.Product] = $true } }

        if ($d.HasHistory -and $planGrids.Count -gt 0 -and $d.HistWeekStart) {
            try {
                $wkStart = $d.HistWeekStart
                $wp = $null
                $gl = @($planGrids); $sl = @($planScheds)
                for ($gi = 0; $gi -lt $gl.Count; $gi++) {
                    # lijnfilter voor de HELE week: alles wat op dit platform gepland staat
                    # van zondag 05:00 tot de zondag erop (zie Read-LineSchedule)
                    $flt = $lineProds
                    if ($null -ne $script:Platform -and $gi -lt $sl.Count) {
                        $ex = Get-PlatformSkus $sl[$gi] $script:Platform $wkStart.Date.AddHours(5) $wkStart.Date.AddDays(7).AddHours(5)
                        if ($ex.Count -gt 0) { $flt = $ex }
                    }
                    $cand = Get-PlanForWeek $gl[$gi] $wkStart $flt
                    if ($cand.PerSku.Count -gt 0) { $wp = $cand; break }
                }
                # gemaakt deze week: eerdere ploegen uit het RCDB-blad ...
                $made = @{}
                foreach ($h in $d.History) {
                    if ($h.WeekIdx -ne 0) { continue }
                    if ($h.Date -eq $prodDate -and $h.ShiftNo -eq $shiftNo) { continue }   # huidige ploeg: zie hieronder
                    foreach ($s in $h.Skus) {
                        if ($made.ContainsKey($s.Sku)) { $made[$s.Sku] += [int]$s.Count } else { $made[$s.Sku] = [int]$s.Count }
                    }
                }
                # ... plus de LOPENDE ploeg uit het Boxruw-blad (fijner en actueler dan het uurraster)
                foreach ($p in $counts.Keys) {
                    if (-not (Is-Sku $p)) { continue }
                    if ($made.ContainsKey($p)) { $made[$p] += [int]$counts[$p] } else { $made[$p] = [int]$counts[$p] }
                }

                $keys = @{}
                if ($wp) { foreach ($k in $wp.PerSku.Keys) { $keys[$k] = $true } }
                foreach ($k in $made.Keys) { $keys[$k] = $true }
                $wrows = @()
                foreach ($k in ($keys.Keys | Sort-Object { if ($wp -and $wp.PerSku.ContainsKey($_)) { [double]$wp.PerSku[$_] } else { 0.0 } } -Descending)) {
                    $pl = 0.0; if ($wp -and $wp.PerSku.ContainsKey($k)) { $pl = [double]$wp.PerSku[$k] }
                    $mk = 0;   if ($made.ContainsKey($k))               { $mk = [int]$made[$k] }
                    if ($pl -le 0 -and $mk -le 0) { continue }
                    $ds = ''
                    if     ($wp -and $wp.Desc.ContainsKey($k)) { $ds = [string]$wp.Desc[$k] }
                    elseif ($planDesc.ContainsKey($k))         { $ds = [string]$planDesc[$k] }
                    $wrows += [pscustomobject]@{ Sku = $k; Desc = $ds; Plan = $pl; Made = $mk; IsCur = $activeProds.ContainsKey($k) }
                }
                if ($wrows.Count -gt 0) {
                    $tp = 0.0; $tm = 0
                    foreach ($w in $wrows) { $tp += [double]$w.Plan; $tm += [int]$w.Made }
                    $d.WeekRows = @($wrows); $d.WeekPlanTotal = $tp; $d.WeekMadeTotal = $tm
                    $d.WeekNo = Get-IsoWeek $wkStart.AddDays(1)      # productieweek start zondag -> maandag bepaalt het ISO-nummer
                    $d.HasWeekPlan = $true
                }
            }
            catch { Add-Warn $d 'warn_weekplan' @($_.Exception.Message) }
        }

        # ---------------- EEN REGEL PER SMAAK (week + ploeg samen) ----------------
        # De pagina toonde deze cijfers eerst als ingeklemde lijstjes in vier kaarten en daarnaast
        # in twee losse tabellen (ploeg / week). Dat las slecht bij vijf smaken tegelijk. Nu staat
        # per smaak alles op EEN regel: wat er deze week gemaakt is, wat er nog te maken is, wat
        # deze ploeg gemaakt/gepland is, de verwachte eindstand en het tempo.
        $combi = @{}
        foreach ($r in $d.Rows) {
            $combi[$r.Product] = [pscustomobject]@{
                Sku      = $r.Product
                Desc     = [string]$r.Desc
                WeekMade = 0; WeekPlan = 0.0; HasWeek = $false
                WeekMadeBefore = 0; WeekDone = $false
                Count    = [int]$r.Count
                Plan     = [double]$r.Plan
                Proj     = [double]$r.Proj
                HasProj  = [bool]$r.HasProj
                PerMin   = [double]$r.PerMin
                Machines = @($r.Machines)
                IsCur    = [bool]$r.IsMain
                # volgorde in de tabel (zie hieronder)
                NextAt   = $null; NextPast = $false; LastMade = $null; OrderGroup = 0; OrderKey = 0.0; OrderKey2 = 0.0
            }
        }
        foreach ($w in $d.WeekRows) {
            if (-not $combi.ContainsKey($w.Sku)) {
                $combi[$w.Sku] = [pscustomobject]@{
                    Sku      = $w.Sku; Desc = [string]$w.Desc
                    WeekMade = 0; WeekPlan = 0.0; HasWeek = $false
                    WeekMadeBefore = 0; WeekDone = $false
                    Count    = 0; Plan = 0.0; Proj = 0.0; HasProj = $false; PerMin = 0.0
                    Machines = @(); IsCur = $false
                    NextAt   = $null; NextPast = $false; LastMade = $null; OrderGroup = 0; OrderKey = 0.0; OrderKey2 = 0.0
                }
            }
            $c = $combi[$w.Sku]
            if ([string]::IsNullOrWhiteSpace($c.Desc) -and $w.Desc) { $c.Desc = [string]$w.Desc }
            $c.WeekMade = [int]$w.Made
            $c.WeekPlan = [double]$w.Plan
            $c.HasWeek  = $true
            # ---- WEEKPLAN AL ROND VOOR DEZE PLOEG ----
            # Het weekplan gaat voor. Was het weekplan van een smaak bij PLOEGSTART al gehaald,
            # dan heeft het geen zin haar deze ploeg nog te draaien - ook al staat ze in het
            # dagplan. Het DAGPLAN IN DOZEN verandert daar NIET van: de ploeg moet evenveel
            # dozen maken, alleen in een andere smaak.
            # 'gemaakt bij ploegstart' = weektotaal min wat DEZE ploeg er zelf van maakte.
            $mb = [int]$w.Made - [int]$c.Count
            if ($mb -lt 0) { $mb = 0 }
            $c.WeekMadeBefore = $mb
            $c.WeekDone = ($c.WeekPlan -gt 0 -and $c.Plan -gt 0 -and $mb -ge $c.WeekPlan)
        }
        # naam erbij voor alles wat het weekplan niet noemt (bv. een smaak die maandag draaide
        # maar deze week niet meer gepland staat) - anders blijft de kolom 'Product' leeg
        foreach ($cv in $combi.Values) {
            if ([string]::IsNullOrWhiteSpace($cv.Desc) -and $skuNames.ContainsKey($cv.Sku)) { $cv.Desc = [string]$skuNames[$cv.Sku] }
        }

        # ---- VOLGORDE: nu -> volgens het weekrooster -> rest (zie Get-SkuOrder) ----
        $wkSunday = $prodDate.AddDays(-[int]$prodDate.DayOfWeek)
        $weekEnd  = $wkSunday.AddDays(7).AddHours(5)             # productieweek: zondag 05:00 -> zondag 05:00
        $pSched = $null; $pGrid = $null
        if ($null -ne $pt) { $pSched = $pt.Sched; $pGrid = $pt.Grid }
        $nextAt   = Get-NextPlannedStart $pSched $pGrid $script:Platform @($combi.Keys) $nowDt $win $weekEnd $planPerSku $counts
        $lastMade = Get-LastMadeMap $d.History $counts $win.Start
        foreach ($cv in $combi.Values) {
            $o = Get-SkuOrder $cv.Sku $cv.IsCur $nextAt $lastMade $cv.Count $cv.Plan $cv.WeekMade $cv.WeekPlan $win.Start
            $cv.NextAt = $o.NextAt; $cv.LastMade = $o.LastMade
            # gepland begin dat AL voorbij is = de run loopt, alleen het ploegplan is nog niet gehaald
            $cv.NextPast = ($null -ne $o.NextAt -and $o.NextAt -le $nowDt)
            $cv.OrderGroup = $o.Group; $cv.OrderKey = $o.Key; $cv.OrderKey2 = $o.Key2
        }
        $d.SkuNames    = $skuNames
        $d.Combined    = @($combi.Values | Sort-Object OrderGroup, OrderKey, OrderKey2, Sku)
        $d.HasCombined = ($d.Combined.Count -gt 0)
        # uitleg onder de tabel voor elke smaak waarvan het weekplan bij ploegstart al rond was
        $wdNotes = @()
        foreach ($c in $d.Combined) {
            if (-not $c.WeekDone) { continue }
            $left = [double]$c.Plan - [double]$c.Count
            if ($left -lt 0) { $left = 0 }
            $wdNotes += [pscustomobject]@{ Sku = $c.Sku; MadeBefore = [int]$c.WeekMadeBefore; WeekPlan = [double]$c.WeekPlan; Left = $left }
        }
        $d.WeekDoneNotes = @($wdNotes)
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

# ================================ GRAFIEKEN (gedeeld) ================================
function New-MinuteChartSvg($d) {
    $W = 1040; $L = 46; $R = 14; $T = 14; $B = 30; $H = 280
    $plotW = $W - $L - $R; $plotH = $H - $T - $B
    $n = [int]$d.ShiftMin; if ($n -le 0) { $n = 480 }
    $mins = $d.Minutes
    $tgt  = [double]$d.TargetPerMin
    $maxV = [double]$d.MaxPerMin
    $baseY = $T + $plotH

    # tot waar reikt de data ('nu'); daarna geen trendlijn tekenen
    $nowIdx = [int][Math]::Round([double]$d.NowOffsetMin)
    if ($nowIdx -le 0 -or $nowIdx -gt $n) {
        $nowIdx = 0
        for ($i = 0; $i -lt $n; $i++) { if ($mins[$i] -gt 0) { $nowIdx = $i + 1 } }
    }

    # voortschrijdend gemiddelde (venster +/- $maR min) = vloeiende trend van dozen/min.
    # Eerst berekenen, want de trendlijn moet sowieso binnen de y-schaal vallen.
    $maR = 7
    $ma = New-Object System.Collections.Generic.List[double]
    $maxMA = 0.0
    if ($nowIdx -ge 2 -and $maxV -gt 0) {
        for ($i = 0; $i -lt $nowIdx; $i++) {
            $lo = [Math]::Max(0, $i - $maR); $hi = [Math]::Min($nowIdx - 1, $i + $maR)
            $sum = 0.0; $cnt = 0
            for ($k = $lo; $k -le $hi; $k++) { $sum += $mins[$k]; $cnt++ }
            $avg = if ($cnt -gt 0) { $sum / $cnt } else { 0.0 }
            $ma.Add($avg)
            if ($avg -gt $maxMA) { $maxMA = $avg }
        }
    }

    # ---- robuuste y-schaal ------------------------------------------------------------
    # Loopt de lijn vast en wordt de prop daarna in een minuut leeggetrokken, dan staat er
    # 1 staaf van 10-15 dozen/min tussen honderden staven van 2-3. Schalen op die absolute
    # piek duwde de hele grafiek (en de target-lijn) plat tegen de onderrand. Daarom schaalt
    # de as nu op de NORMALE spreiding van de gedraaide minuten: bovengrens = P75 + 1,5*IQR
    # (klassieke uitschieter-grens) en minstens P95. Wat daarboven uitkomt wordt afgekapt en
    # met een roze pijl + label gemarkeerd, zodat de uitschieter zichtbaar blijft.
    $nz = New-Object System.Collections.Generic.List[double]
    for ($i = 0; $i -lt $nowIdx; $i++) { if ($mins[$i] -gt 0) { $nz.Add([double]$mins[$i]) } }
    $srt = @($nz | Sort-Object)
    $Pct = {
        param($sorted, [double]$p)
        if ($sorted.Count -eq 0) { return 0.0 }
        $i = [int][Math]::Ceiling($p * $sorted.Count) - 1
        if ($i -lt 0) { $i = 0 }
        if ($i -ge $sorted.Count) { $i = $sorted.Count - 1 }
        return [double]$sorted[$i]
    }
    $q1 = & $Pct $srt 0.25; $q3 = & $Pct $srt 0.75; $p95 = & $Pct $srt 0.95
    $yTop = [Math]::Max(($q3 + 1.5 * ($q3 - $q1)), $p95)
    if ($yTop -le 0) { $yTop = $maxV }
    $yTop = [Math]::Max($yTop, $maxMA * 1.25)                                   # trendlijn moet passen
    $yTop = $yTop * 1.1
    if ($maxV -gt 0) { $yTop = [Math]::Min($yTop, $maxV * 1.15) }               # nooit meer lucht dan er data is
    $yTop = [Math]::Max($yTop, [Math]::Max($tgt, [double]$d.PerMin) * 1.25)     # target-/tempo-lijn moet passen
    if ($yTop -le 0) { $yTop = 1 }
    $yTopLabel = [int][Math]::Ceiling($yTop)
    $yTop = [double]$yTopLabel                                                   # hele dozen op de as

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

    # ---- ruwe staven per minuut (licht doorschijnend) + vloeiende trendlijn erbovenop ----
    $barW = [Math]::Max(1.2, ($plotW / $n) * 0.85)
    $bws  = SvgN $barW
    $clipMax = 0; $clipCnt = 0
    for ($i = 0; $i -lt $n; $i++) {
        $v = $mins[$i]; if ($v -le 0) { continue }
        $x = $L + ($i / $n) * $plotW
        if ($v -gt $yTop) {
            # buiten schaal (bv. leeggetrokken opstopping): afkappen + pijl, met de echte waarde in de tooltip
            $clipCnt++; if ($v -gt $clipMax) { $clipMax = $v }
            $tip = '{0} - {1}' -f $d.ShiftStart.AddMinutes($i).ToString('HH:mm'), $v
            [void]$sb.Append("<rect x='$(SvgN $x)' y='$T' width='$bws' height='$(SvgN $plotH)' fill='#38bdf8' opacity='0.28'><title>$(HtmlEnc $tip)</title></rect>")
            $cx = $x + $barW / 2; $tw = [Math]::Max(3.0, $barW * 0.9)
            [void]$sb.Append("<path d='M $(SvgN ($cx - $tw)) $(SvgN ($T + 8)) L $(SvgN $cx) $(SvgN ($T + 1)) L $(SvgN ($cx + $tw)) $(SvgN ($T + 8)) Z' fill='#f472b6'/>")
        }
        else {
            $h = ($v / $yTop) * $plotH
            $y = $baseY - $h
            [void]$sb.Append("<rect x='$(SvgN $x)' y='$(SvgN $y)' width='$bws' height='$(SvgN $h)' fill='#38bdf8' opacity='0.28'/>")
        }
    }

    # vloeiende trend (voortschrijdend gemiddelde, hierboven berekend)
    if ($ma.Count -ge 2) {
        $maPts = New-Object System.Text.StringBuilder
        for ($i = 0; $i -lt $ma.Count; $i++) {
            $x = $L + ($i / $n) * $plotW
            $y = $baseY - ([Math]::Min($ma[$i], $yTop) / $yTop) * $plotH
            [void]$maPts.Append($(if ($i -eq 0) { 'M ' } else { ' L ' }))
            [void]$maPts.Append("$(SvgN $x) $(SvgN $y)")
        }
        [void]$sb.Append("<path d='$($maPts.ToString())' fill='none' stroke='#38bdf8' stroke-width='2.4' stroke-linejoin='round' stroke-linecap='round'/>")
    }

    # horizontale target-lijn
    $ty = $baseY - ($tgt / $yTop) * $plotH
    [void]$sb.Append("<line x1='$L' y1='$(SvgN $ty)' x2='$($L + $plotW)' y2='$(SvgN $ty)' stroke='#fbbf24' stroke-width='2' stroke-dasharray='6 4'/>")
    $tgtTxt = (T 'svg_target_line') -f (NF $d.Target), (PF2 $tgt)
    [void]$sb.Append("<text x='$($L + $plotW - 6)' y='$(SvgN ($ty - 6))' fill='#fbbf24' font-size='12' text-anchor='end'>$(HtmlEnc $tgtTxt)</text>")

    # gemiddeld tempo van de lopende run (= de lijn waarmee de prognose rekent)
    if ($d.HasForecast -and [double]$d.PerMin -gt 0) {
        $py = $baseY - ([double]$d.PerMin / $yTop) * $plotH
        [void]$sb.Append("<line x1='$L' y1='$(SvgN $py)' x2='$($L + $plotW)' y2='$(SvgN $py)' stroke='#34d399' stroke-width='2'/>")
        $pTxt = (T 'svg_tempo_line') -f (PF2 $d.PerMin)
        $pLab = if ([Math]::Abs($py - $ty) -lt 14) { $py + 14 } else { $py - 6 }
        [void]$sb.Append("<text x='$($L + 6)' y='$(SvgN $pLab)' fill='#34d399' font-size='12'>$(HtmlEnc $pTxt)</text>")
    }

    # y-as labels (0 en max)
    [void]$sb.Append("<text x='$($L - 6)' y='$baseY' fill='#94a3b8' font-size='11' text-anchor='end'>0</text>")
    [void]$sb.Append("<text x='$($L - 6)' y='$($T + 10)' fill='#94a3b8' font-size='11' text-anchor='end'>$yTopLabel</text>")

    # melding bij afgekapte staven: hoeveel en hoe hoog de piek werkelijk was
    if ($clipMax -gt 0) {
        $cTxt = (T 'svg_clip_note') -f (NF $clipMax), (NF $clipCnt)
        [void]$sb.Append("<text x='$($L + 8)' y='$($T + 14)' fill='#f472b6' font-size='11'>$(HtmlEnc $cTxt)</text>")
    }

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

# Rechterkolom naast een stilstandstrook: hoeveel % van de VERSTREKEN ploegtijd er gedraaid en
# stilgestaan is. $runPct = draaitijd in % ($null = er is nog niets verstreken -> streepje).
# Eerst op 1 decimaal afronden en stilstand = 100 - draaitijd, dan tellen de twee cijfers altijd
# precies op tot 100,0. De tooltip geeft de minuten en het venster ($from + $elMin): voor een machine
# loopt dat tot de laatste doos van de lijn (zoals de kolom 'Draaitijd'), niet tot 'nu' - staat de
# HELE lijn al lang stil, dan zie je daar dus bv. 13:00 - 19:01. $dim = grijs ('(onbekend)': geen machine).
function Get-StripPctSvg([double]$xRun, [double]$xStop, [double]$yText, $runPct, [double]$runMin, [double]$elMin, [datetime]$from, [string]$label, [bool]$bold, [bool]$dim) {
    if ($null -eq $runPct -or $elMin -lt 1) {
        return "<text x='$(SvgN $xStop)' y='$(SvgN $yText)' fill='#64748b' font-size='11' text-anchor='end'>&#8212;</text>"
    }
    $pRun  = [Math]::Round([Math]::Min(100.0, [Math]::Max(0.0, [double]$runPct)), 1, [MidpointRounding]::AwayFromZero)
    $pStop = 100.0 - $pRun
    $fw    = if ($bold) { '700' } else { '600' }
    $cRun  = if ($dim) { '#64748b' } else { '#34d399' }
    $cStop = if ($dim) { '#64748b' } else { '#f87171' }
    $tip   = (T 'tt_strip_pct') -f $label, (NF $runMin), (NF ([Math]::Max(0.0, $elMin - $runMin))),
                                  $from.ToString('HH:mm'), $from.AddMinutes($elMin).ToString('HH:mm')
    return "<g><title>$(HtmlEnc $tip)</title>" +
           "<text x='$(SvgN $xRun)' y='$(SvgN $yText)' fill='$cRun' font-size='11' font-weight='$fw' text-anchor='end'>$(PF $pRun) %</text>" +
           "<text x='$(SvgN $xStop)' y='$(SvgN $yText)' fill='$cStop' font-size='11' font-weight='$fw' text-anchor='end'>$(PF $pStop) %</text></g>"
}
# Kopjes boven die kolom ('draait' / 'stil').
function Get-StripPctHeadSvg([double]$xRun, [double]$xStop, [double]$y) {
    return "<text x='$(SvgN $xRun)' y='$(SvgN $y)' fill='#94a3b8' font-size='10' text-anchor='end'>$(HtmlEnc (T 'svg_pct_run'))</text>" +
           "<text x='$(SvgN $xStop)' y='$(SvgN $y)' fill='#94a3b8' font-size='10' text-anchor='end'>$(HtmlEnc (T 'svg_pct_stop'))</text>"
}

function New-StopStripSvg($d) {
    # rechts van de strook: % draaitijd / % stilstand (zelfde cijfer als de kaart 'Draaitijd')
    $W = 1040; $L = 46; $R = 14; $T = 22; $H = 70; $barH = 26; $pctW = 116
    $plotW = $W - $L - $R - $pctW
    $xStop = $W - $R; $xRun = $xStop - 58
    $n = [int]$d.ShiftMin; if ($n -le 0) { $n = 480 }
    $baseY = $T + $barH

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<svg viewBox='0 0 $W $H' width='100%' preserveAspectRatio='xMidYMid meet' xmlns='http://www.w3.org/2000/svg' font-family='Segoe UI,system-ui,Arial,sans-serif'>")
    [void]$sb.Append((Get-StripPctHeadSvg $xRun $xStop ($T - 8)))
    # afgeronde baan: hele ploeg = nog te gaan (grijs)
    [void]$sb.Append("<rect x='$L' y='$T' width='$plotW' height='$barH' rx='7' fill='#172033' stroke='#334155'/>")
    $nowOff = [double]$d.NowOffsetMin; if ($nowOff -lt 0) { $nowOff = 0 }; if ($nowOff -gt $n) { $nowOff = $n }
    $wNow = ($nowOff / $n) * $plotW
    # alle segmenten binnen de afgeronde baan clippen (blijven netjes binnen de rand)
    [void]$sb.Append("<clipPath id='stripClip'><rect x='$L' y='$T' width='$plotW' height='$barH' rx='7'/></clipPath>")
    [void]$sb.Append("<g clip-path='url(#stripClip)'>")
    # verstreken deel = draait (groen)
    [void]$sb.Append("<rect x='$L' y='$T' width='$(SvgN $wNow)' height='$barH' fill='#10b981'/>")

    # stilstanden er overheen
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
            [void]$sb.Append("<text x='$(SvgN ($x + $w / 2))' y='$($T + 17)' fill='#450a0a' font-size='11' font-weight='600' text-anchor='middle'>$([int][Math]::Round($s.Min)) $(T 'svg_min')</text>")
        }
    }
    [void]$sb.Append("</g>")

    # uur-ticks
    for ($m = 0; $m -le $n; $m += 60) {
        $x = $L + ($m / $n) * $plotW
        $lab = $d.ShiftStart.AddMinutes($m).ToString('HH:mm')
        [void]$sb.Append("<line x1='$(SvgN $x)' y1='$baseY' x2='$(SvgN $x)' y2='$($baseY + 4)' stroke='#475569'/>")
        [void]$sb.Append("<text x='$(SvgN $x)' y='$($baseY + 16)' fill='#94a3b8' font-size='11' text-anchor='middle'>$lab</text>")
    }
    [void]$sb.Append("<text x='$($L - 6)' y='$($T + 17)' fill='#94a3b8' font-size='11' text-anchor='end'>$(T 'svg_line')</text>")
    $pct = if ([double]$d.ElapsedShiftMin -ge 1) { [double]$d.AvailPct } else { $null }
    [void]$sb.Append((Get-StripPctSvg $xRun $xStop ($T + 17) $pct ([double]$d.RunMin) ([double]$d.ElapsedShiftMin) $d.ShiftStart (T 'svg_line') $true $false))
    [void]$sb.Append("</svg>")
    return $sb.ToString()
}

function New-MachineStripSvg($d) {
    # rechts van elke strook: % draaitijd / % stilstand van de verstreken ploegtijd
    $W = 1040; $L = 78; $R = 14; $T = 24; $rowH = 20; $gapY = 5; $B = 22; $pctW = 116
    $plotW = $W - $L - $R - $pctW
    $xStop = $W - $R; $xRun = $xStop - 58
    $n = [int]$d.ShiftMin; if ($n -le 0) { $n = 480 }
    $rows = @()
    # lijn = wat de bovenste strook tekent: stil zolang GEEN ENKELE machine een doos maakt (tot 'nu')
    $lineEl  = [double]$d.NowOffsetMin
    $lineRun = [Math]::Max(0.0, $lineEl - [double]$d.StopMin)
    $linePct = if ($lineEl -ge 1) { 100.0 * $lineRun / $lineEl } else { $null }
    $rows += [pscustomobject]@{ Label = (T 'svg_line'); Stops = $d.Stops; Wide = $true
                                RunPct = $linePct; RunMin = $lineRun; ElMin = $lineEl; Dim = $false }
    # machine = hetzelfde cijfer als de kolom 'Draaitijd' in de tabel eronder (gerekend tot de data-horizon)
    foreach ($m in $d.Machines) {
        $mEl  = [double]$d.MachElapsedMin
        $mPct = if ($mEl -ge 1) { [double]$m.AvailPct } else { $null }
        $rows += [pscustomobject]@{ Label = $m.Machine; Stops = $m.Stops; Wide = $false
                                    RunPct = $mPct; RunMin = [double]$m.RunMin; ElMin = $mEl; Dim = (-not $m.IsReal) }
    }
    $H = $T + $rows.Count * ($rowH + $gapY) + $B

    $nowOff = [double]$d.NowOffsetMin; if ($nowOff -lt 0) { $nowOff = 0 }; if ($nowOff -gt $n) { $nowOff = $n }
    $wNow = ($nowOff / $n) * $plotW

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<svg viewBox='0 0 $W $H' width='100%' preserveAspectRatio='xMidYMid meet' xmlns='http://www.w3.org/2000/svg' font-family='Segoe UI,system-ui,Arial,sans-serif'>")
    [void]$sb.Append("<clipPath id='mstripClip'><rect x='$L' y='0' width='$plotW' height='$H'/></clipPath>")
    [void]$sb.Append((Get-StripPctHeadSvg $xRun $xStop ($T - 9)))

    $y = $T
    foreach ($r in $rows) {
        $h = $rowH
        # baan: hele ploeg grijs, verstreken deel groen
        [void]$sb.Append("<rect x='$L' y='$(SvgN $y)' width='$plotW' height='$h' rx='5' fill='#172033' stroke='#334155'/>")
        [void]$sb.Append("<g clip-path='url(#mstripClip)'>")
        [void]$sb.Append("<rect x='$L' y='$(SvgN $y)' width='$(SvgN $wNow)' height='$h' fill='$(if ($r.Wide) { '#0ea5e9' } else { '#10b981' })'/>")
        foreach ($s in @($r.Stops)) {
            $a = ($s.From - $d.ShiftStart).TotalMinutes; if ($a -lt 0) { $a = 0 }
            $b = ($s.To   - $d.ShiftStart).TotalMinutes; if ($b -gt $n) { $b = $n }
            if ($b -le $a) { continue }
            $x = $L + ($a / $n) * $plotW
            $w = [Math]::Max(1.5, (($b - $a) / $n) * $plotW)
            $fill = if ($s.Kind -eq 'nu') { '#f97316' } elseif ($s.Kind -eq 'opstart') { '#7c2d12' } else { '#ef4444' }
            [void]$sb.Append("<rect x='$(SvgN $x)' y='$(SvgN $y)' width='$(SvgN $w)' height='$h' fill='$fill'/>")
            # alleen de lange stops krijgen een tijdlabel (anders wordt het een kluwen)
            if ($s.Min -ge 20 -and $w -ge 26) {
                [void]$sb.Append("<text x='$(SvgN ($x + $w / 2))' y='$(SvgN ($y + $h - 6))' fill='#450a0a' font-size='10' font-weight='600' text-anchor='middle'>$([int][Math]::Round($s.Min))</text>")
            }
        }
        [void]$sb.Append("</g>")
        $fw = if ($r.Wide) { '700' } else { '400' }
        [void]$sb.Append("<text x='$($L - 8)' y='$(SvgN ($y + $h - 6))' fill='#cbd5e1' font-size='11' font-weight='$fw' text-anchor='end'>$(HtmlEnc $r.Label)</text>")
        [void]$sb.Append((Get-StripPctSvg $xRun $xStop ($y + $h - 6) $r.RunPct $r.RunMin $r.ElMin $d.ShiftStart ([string]$r.Label) $r.Wide $r.Dim))
        $y += $rowH + $gapY
    }

    # uur-ticks onder de laatste strook
    for ($m = 0; $m -le $n; $m += 60) {
        $x = $L + ($m / $n) * $plotW
        $lab = $d.ShiftStart.AddMinutes($m).ToString('HH:mm')
        [void]$sb.Append("<line x1='$(SvgN $x)' y1='$(SvgN $y)' x2='$(SvgN $x)' y2='$(SvgN ($y + 4))' stroke='#475569'/>")
        [void]$sb.Append("<text x='$(SvgN $x)' y='$(SvgN ($y + 16))' fill='#94a3b8' font-size='11' text-anchor='middle'>$lab</text>")
    }
    [void]$sb.Append("</svg>")
    return $sb.ToString()
}

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
    # De SCHAAL komt UITSLUITEND uit het werkelijke verloop ($vals); de prognose telt hier
    # BEWUST NIET mee. Anders drukt een uitschieter (eindstand +200 terwijl de ploeg de hele
    # tijd binnen +/-50 bleef) het echte verloop plat tot een streepje en zie je niet meer
    # waar de rode en de groene zone liggen. Valt de prognose buiten deze schaal, dan wordt
    # de stippellijn verderop afgeknipt op de plotrand (met een driehoekje als teken).
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
    # diverging vlak: groen boven de nullijn (voor op schema), rood eronder (achter)
    $segW = SvgN (($plotW / $n) * 2 + 0.6)
    for ($i = 0; $i -lt $vals.Count; $i += 2) {
        $yv = & $MapY $vals[$i]
        $yA = [Math]::Min($yv, $y0); $hh = [Math]::Abs($yv - $y0)
        if ($hh -lt 0.4) { continue }
        $col2 = if ($vals[$i] -ge 0) { '#34d399' } else { '#f87171' }
        [void]$sb.Append("<rect x='$(SvgN (& $MapX ($i + 1)))' y='$(SvgN $yA)' width='$segW' height='$(SvgN $hh)' fill='$col2' opacity='0.4'/>")
    }
    [void]$sb.Append("<line x1='$L' y1='$(SvgN $y0)' x2='$($L + $plotW)' y2='$(SvgN $y0)' stroke='#fbbf24' stroke-width='2' stroke-dasharray='6 4'/>")
    [void]$sb.Append("<text x='$($L + 6)' y='$(SvgN ($y0 - 6))' fill='#fbbf24' font-size='12'>$(T 'svg_on_target')</text>")

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
        [void]$sb.Append("<polyline points='$($pts.ToString().Trim())' fill='none' stroke='#e2e8f0' stroke-width='1.8' opacity='0.85' stroke-linejoin='round'/>")
    }

    # Doortrekken naar ploegeinde volgens de prognose. Ligt het eindpunt buiten de schaal,
    # dan knippen we de stippellijn af op de plotrand en zetten daar een driehoekje: de lijn
    # wijst naar buiten, het cijfer zelf blijft in het label leesbaar.
    $xNow = & $MapX $nowOff; $yNow = & $MapY $behind
    if ($d.HasForecast) {
        $xEnd = $L + $plotW; $yEnd = & $MapY $end
        $yTop = $T + 3.0; $yBot = $T + $plotH - 3.0
        $offTop = ($yEnd -lt $yTop); $offBot = ($yEnd -gt $yBot); $off = ($offTop -or $offBot)
        $xTip = $xEnd; $yTip = $yEnd
        if ($off) {
            # snijpunt van (xNow,yNow)-(xEnd,yEnd) met de boven- of onderrand van het plotvlak
            $yLim = if ($offTop) { $yTop } else { $yBot }
            $dy = $yEnd - $yNow
            $t = if ([Math]::Abs($dy) -gt 0.001) { ($yLim - $yNow) / $dy } else { 1.0 }
            if ($t -lt 0) { $t = 0.0 }; if ($t -gt 1) { $t = 1.0 }
            $xTip = $xNow + ($xEnd - $xNow) * $t
            $yTip = $yLim
        }
        [void]$sb.Append("<line x1='$(SvgN $xNow)' y1='$(SvgN $yNow)' x2='$(SvgN $xTip)' y2='$(SvgN $yTip)' stroke='#34d399' stroke-width='2' stroke-dasharray='7 5'/>")
        $endTxt = if ($end -ge 0) { "$(T 'svg_shift_end_word') +$(NF $end)" } else { "$(T 'svg_shift_end_word') $(NF $end)" }
        if ($off) {
            $bs = if ($offTop) { 9.0 } else { -9.0 }
            $tri = "$(SvgN $xTip),$(SvgN $yTip) $(SvgN ($xTip - 5)),$(SvgN ($yTip + $bs)) $(SvgN ($xTip + 5)),$(SvgN ($yTip + $bs))"
            [void]$sb.Append("<polygon points='$tri' fill='#34d399'/>")
            $ly = if ($offTop) { $yTip + 26 } else { $yTip - 18 }
            if (($xTip + 130) -lt ($L + $plotW)) {
                [void]$sb.Append("<text x='$(SvgN ($xTip + 9))' y='$(SvgN $ly)' fill='#34d399' font-size='13' font-weight='600' text-anchor='start'>$(HtmlEnc $endTxt)</text>")
            } else {
                [void]$sb.Append("<text x='$($L + $plotW - 6)' y='$(SvgN $ly)' fill='#34d399' font-size='13' font-weight='600' text-anchor='end'>$(HtmlEnc $endTxt)</text>")
            }
        } else {
            [void]$sb.Append("<circle cx='$(SvgN $xEnd)' cy='$(SvgN $yEnd)' r='4' fill='#34d399'/>")
            $ly = if ($yEnd -lt ($T + 18)) { $yEnd + 16 } else { $yEnd - 8 }
            [void]$sb.Append("<text x='$($L + $plotW - 6)' y='$(SvgN $ly)' fill='#34d399' font-size='13' font-weight='600' text-anchor='end'>$(HtmlEnc $endTxt)</text>")
        }
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

# ============================== WEERGAVE LIJN 9 ==============================
function Render-Html9($d, [string]$lang = 'nl') {
    if ($script:Langs -notcontains $lang) { $lang = $script:DefaultLang }
    $script:Lang = $lang
    # de server houdt de data zelf vers (leest bij bestandswijziging), dus de pagina mag vaker en
    # goedkoop verversen; begrensd op [5..15]s zodat een wijziging snel zichtbaar wordt
    $refresh = [Math]::Max(5, [Math]::Min([int]$IntervalSeconds, 15))
    $load = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')

    $langbar = "<div class='langbar'>"
    foreach ($lc in $script:Langs) {
        $cls = if ($lc -eq $lang) { 'lang active' } else { 'lang' }
        $langbar += "<a class='$cls' href='/?lang=$lc&amp;line=$($script:Line)'>$($lc.ToUpper())</a>"
    }
    $langbar += "</div>"
    # ---- LIJNKIEZER: zelfde pagina, andere lijn (?line=9 / ?line=11) ----
    $linebar = "<div class='linebar'>"
    foreach ($ld in $script:LineOrder) {
        $lcls = if ($ld -eq $script:Line) { 'lineb active' } else { 'lineb' }
        $linebar += "<a class='$lcls' href='/?lang=$lang&amp;line=$ld'>$((T 'line_word') -f $ld)</a>"
    }
    $linebar += "</div>"

    # ploegletter (X/Y/Z, wisselt per week) i.p.v. het nummer; let op: $shiftLabel gaat door
    # HtmlEnc, dus hier GEEN html-entiteiten gebruiken
    $shiftLabel = "$(T 'shift_word') $($d.ShiftLetter) ($($d.ShiftRange))"
    $meta = "$(HtmlEnc $d.NowText) &middot; $(HtmlEnc $shiftLabel) &middot; $(T 'interval') $(HtmlEnc $d.WindowText)"
    $meta += " &middot; $(HtmlEnc $d.Sheet) &middot; $(HtmlEnc $d.BoxFile)"

    # herkomst van de target (weekplan / parameter / config), taalafhankelijk opgebouwd
    $tsrc = switch ($d.TargetMode) {
        'plan' {
            $cn = $script:CultMap[$lang]; $dt = ''
            if ($d.PlanDate) { try { $dt = $d.PlanDate.ToString('ddd dd/MM', [System.Globalization.CultureInfo]::GetCultureInfo($cn)) } catch { $dt = $d.PlanDate.ToString('dd/MM') } }
            (T 'ts_plan') -f (HtmlEnc $d.PlanFileName), $d.PlanWeek, (HtmlEnc $dt), $d.ShiftNo
        }
        'planfallback' { (T 'ts_planfallback') -f (HtmlEnc $d.PlanFileName), (HtmlEnc $d.PlanSku) }
        'param'  { T 'ts_param' }
        'noplan' { T 'ts_noplan' }
        default  { T 'ts_unknown' }
    }

    # gestapelde waarschuwingen (taalafhankelijk)
    $warnHtml = ""
    if ($d.WarnList -and @($d.WarnList).Count -gt 0) {
        foreach ($w in $d.WarnList) {
            $vals = @($w.Vals | ForEach-Object { HtmlEnc ([string]$_) })
            $wtxt = if ($vals.Count -gt 0) { (T $w.Key) -f $vals } else { (T $w.Key) }
            $warnHtml += "<div class='warn'>$wtxt</div>"
        }
    }

    if ($d.Error) {
        $bodyHtml = "<div class='err'>$(T 'err_prefix'): $(HtmlEnc $d.Error)</div>"
    }
    else {
        $pct = PF $d.Pct
        # draaide er meer dan 1 producttype, dan onder het totaal per type uitsplitsen
        # (bij 1 product zegt zo'n regel niets - dan blijft de kaart zoals hij was)
        # kleurdrempels voor 'hoeveel % van het plan': groen >= 100, oranje >= 80, anders rood
        $PctCls = { param([double]$p) if ($p -ge 100) { 'g' } elseif ($p -ge 80) { 'a' } else { 'r' } }
        # balkje 'hoeveel van het plan is gemaakt' (producttabel EN historie).
        # De balk loopt tot 100 %; wat erboven zit staat in het cijfer (bv. 181 %).
        $Bar = {
            param([double]$count, [double]$plan, [bool]$bold)
            if ($plan -le 0) { return "<span class='dash'>&mdash;</span>" }
            $p = 100.0 * $count / $plan
            $cls = & $PctCls $p
            $w = [Math]::Min(100.0, [Math]::Max(0.0, $p))
            $bcls = if ($bold) { " strong" } else { "" }
            return "<span class='pw$bcls'><span class='pv $cls'>$(PF $p)&nbsp;%</span><span class='bar'><i class='$cls' style='width:$(SvgN $w)%'></i></span></span>"
        }
        # 'laatste doos' in de taal van de pagina (staat al html-veilig, dus NIET nog eens encoderen)
        $lastBoxTxt = if ($d.LastTimeText) {
            (T 'last_box_txt') -f (HtmlEnc $d.LastTimeText), (HtmlEnc $d.LastProduct), (HtmlEnc $d.LastCounter)
        } else { HtmlEnc $d.LastText }

        $madeRows = @($d.Rows | Where-Object { $_.Count -gt 0 })
        $madeSplit = ""
        if ($madeRows.Count -gt 1) {
            $madeSplit = "<div class='split'>"
            foreach ($r in $madeRows) {
                $mark = if ($r.IsMain) { " <span class='nu'>$(T 'kind_nowmark')</span>" } else { "" }
                $ttl  = if ($r.Desc) { " title='$(HtmlEnc $r.Desc)'" } else { "" }
                # per smaak ook het percentage van het plan van die smaak
                $rp = ""
                if ($r.Plan -gt 0) {
                    $pv = 100.0 * $r.Count / $r.Plan
                    $rp = "<span class='pcs pcl $(& $PctCls $pv)'>$(PF $pv)&nbsp;%</span>"
                }
                $madeSplit += "<div$ttl><span class='sk'>$(HtmlEnc $r.Product)$mark</span><span class='tv'><b>$(NF $r.Count)</b>$rp</span></div>"
            }
            $madeSplit += "</div>"
        }
        # naast het grote totaal: hoeveel procent van het plan van de hele ploeg is dat
        $madePct = ""
        if ($d.Target -gt 0) {
            $madePct = "<span class='vpct pcl $(& $PctCls ([double]$d.Pct))'>$((T 'card_made_pct') -f $pct)</span>"
        }
        # target ploeg = alle smaken die deze ploeg gepland staan (ook wat nog niet gestart is)
        $planRows = @($d.Rows | Where-Object { $_.Plan -gt 0 })
        $tgtSplit = ""
        if ($planRows.Count -gt 1) {
            $tgtSplit = "<div class='split'>"
            foreach ($r in $planRows) {
                $mark = if ($r.IsMain) { " <span class='nu'>$(T 'kind_nowmark')</span>" }
                        elseif ($r.Count -le 0) { " <span class='soon'>$(T 'kind_notstarted')</span>" }
                        else { "" }
                if ($r.WeekDone) { $mark += " <span class='soon'>$(T 'kind_weekdone')</span>" }
                $ttl  = if ($r.Desc) { " title='$(HtmlEnc $r.Desc)'" } else { "" }
                $tgtSplit += "<div$ttl><span class='sk'>$(HtmlEnc $r.Product)$mark</span><b>$(NF $r.Plan)</b></div>"
            }
            $tgtSplit += "</div>"
        }
        # tempo per producttype: dozen van dat type / minuten dat dat type liep (omstellen telt niet mee)
        $tempoRows = @($madeRows | Where-Object { $d.Tempo.ContainsKey($_.Product) })
        $tempoSplitRows = ""
        if ($tempoRows.Count -gt 1) {
            foreach ($r in $tempoRows) {
                $t = $d.Tempo[$r.Product]
                $mark = if ($r.IsMain) { " <span class='nu'>$(T 'kind_nowmark')</span>" } else { "" }
                $ttl = (T 'tt_tempo') -f (NF $t.Count), (NF ([Math]::Round($t.Min)))
                if ($r.Desc) { $ttl = "$($r.Desc) - $ttl" }
                $val = if ($t.HasRate) {
                    "<b>$(PF2 $t.PerMin)</b><span class='ph'>$(NF ($t.PerMin * 60))$(T 'per_hour_short')</span>"
                } else { "<b>&mdash;</b>" }
                $tempoSplitRows += "<div title='$(HtmlEnc $ttl)'><span class='sk'>$(HtmlEnc $r.Product)$mark</span><span class='tv'>$val</span></div>"
            }
        }
        # tempo van de laatste -RecentMinutes minuten hoort bij 'Tempo nu' (de grote waarde = hele ploeg)
        if ($d.HasRecent) {
            $recTtl = (T 'tt_tempo') -f (NF $d.RecentCount), (NF $d.RecentWin)
            $tempoSplitRows += "<div title='$(HtmlEnc $recTtl)'><span class='sk'>$((T 'card_last_win') -f (NF $d.RecentWin))</span>" +
                               "<span class='tv'><b>$(PF2 $d.RecentPerMin)</b><span class='ph'>$(NF ($d.RecentPerMin * 60))$(T 'per_hour_short')</span></span></div>"
        }
        $tempoSplit = if ($tempoSplitRows) { "<div class='split'>$tempoSplitRows</div>" } else { "" }
        $cards = "<div class='cards'>" +
            "<div class='card'><div class='lbl'>$(T 'card_made')</div><div class='val done'>$(NF $d.Total)$madePct</div>$madeSplit<div class='sub'>$((T 'card_made_sub') -f (HtmlEnc $d.FileTimeText))</div></div>" +
            "<div class='card'><div class='lbl'>$(T 'card_target')</div><div class='val'>$(NF $d.Target)</div>$tgtSplit<div class='sub'>$((T 'card_target_sub') -f $pct, (PF2 $d.TargetPerMin))<br><span class='bron'>$tsrc</span></div></div>"
        if ($d.HasForecast) {
            $projCls = if ($d.ProjDiff -ge 0) { "done" } else { "behind" }
            $diffTxt = if ($d.ProjDiff -ge 0) { "+$(NF $d.ProjDiff)" } else { (NF $d.ProjDiff) }
            # ETA ("target bereikt om") hoort bij de verwachte eindstand -> als regel in dezelfde kaart
            $etaSplit = ""
            if ($d.EtaText) {
                # Een enkele regel LINKS uitgelijnd (geen label-links/waarde-rechts: met maar een
                # waarde oogt dat scheef). Tijd = vet, '(na ploegeinde)' = kleine grijze noot.
                # Bij een tekstuitkomst vervalt het label - 'Target bereikt om target al gehaald'
                # zou onzin zijn - en kleurt de uitkomst zelf.
                $etaTxt = switch ($d.EtaKind) {
                    'done'       { "<span class='done'>$(T 'eta_done')</span>" }
                    'over'       { "<b>$(T 'eta_over')</b>" }
                    'impossible' { "<span class='behind'>$(T 'eta_impossible')</span>" }
                    'time'       { "$(T 'card_eta') <b>$(HtmlEnc $d.EtaTimeText)</b>$(if ($d.EtaAfterShift) { "<span class='ph'>$((T 'eta_after').Trim())</span>" } else { '' })" }
                    default      { "<b>$(HtmlEnc $d.EtaText)</b>" }
                }
                $etaSplit = "<div class='eta' title='$(HtmlEnc (T 'card_eta_sub'))'>$etaTxt</div>"
            }
            $cards += "<div class='card hi'><div class='lbl'>$(T 'card_expected_end')</div><div class='val $projCls'>$(NF $d.ProjTotal)</div>$etaSplit" +
                      "<div class='sub'>$((T 'card_expected_end_sub') -f (PF $d.ProjPct), $diffTxt)</div></div>" +
                      "<div class='card'><div class='lbl'>$(T 'card_tempo_now')</div><div class='val'>$(PF2 $d.PerMin)</div>$tempoSplit<div class='sub'>$((T 'card_tempo_now_sub') -f (NF $d.PerHour))</div></div>"
        }
        $cards += "</div>"
        $chart = "<h2 class='sec'>$(T 'sec_per_minute')</h2><div class='chart'>$(New-MinuteChartSvg $d)</div>"

        # De sectie 'prognose einde ploeg' (kop + cumulatieve grafiek + runregel) is
        # bewust weggelaten: de kaart 'verwachte eindstand' bovenaan toont dezelfde
        # prognose in een cijfer. De BEREKENING blijft wel staan - die voedt die kaart,
        # de achterstandgrafiek en de kolom 'verwachte eindstand' in de tabel.
        $fc = ""

        $st = ""
        $bh = ""
        if ($d.HasStops) {
            $stillCls = if ($d.StopMin -gt 0) { "behind" } else { "done" }
            $availCls = if ($d.AvailPct -ge 95) { "done" } else { "behind" }
            $nowStillHtml = if ($d.NowStill) { "<div class='err'>$((T 'nowstill') -f (NF $d.StillMin), $lastBoxTxt)</div>" } else { "" }
            $allStops = @($d.Stops | Sort-Object From)
            $stopTotal = $allStops.Count
            $stopCap = 24
            $stopRows = ""
            for ($si = 0; $si -lt $stopTotal; $si++) {
                $s = $allStops[$si]
                $kindTxt = switch ($s.Kind) { 'opstart' { T 'kind_startup' } 'nu' { T 'kind_nowstill' } default { T 'kind_stop' } }
                if ($s.IsLongest) { $kindTxt += " <span class='nu'>$(T 'kind_longest')</span>" }
                $rowCls = if ($si -ge $stopCap) { " class='stop-extra'" } else { "" }
                $stopRows += "<tr$rowCls><td class='sku'>$($s.From.ToString('HH:mm')) &ndash; $($s.To.ToString('HH:mm'))</td><td class='dur'>$(Format-Dur $s.Min)</td><td>$kindTxt</td></tr>"
            }
            $stopHead = (T 'stops_head_all') -f $stopTotal
            $st = "<h2 class='sec'>$(T 'stops_word') <span class='bron'>$((T 'stops_bron') -f (PF $d.StopLimit))</span></h2>$nowStillHtml" +
                  "<div class='chart'>$(New-StopStripSvg $d)</div>" +
                  "<div class='cards'>" +
                  "<div class='card'><div class='lbl'>$(T 'stops_word')</div><div class='val $stillCls'>$(NF $d.StopMin)</div><div class='sub'>$((T 'card_stops_sub') -f $d.StopCount, (Format-Dur $d.LongestMin), (HtmlEnc $d.LongestText))</div></div>" +
                  "<div class='card'><div class='lbl'>$(T 'card_runtime')</div><div class='val $availCls'>$(PF $d.AvailPct)</div><div class='sub'>$((T 'card_runtime_sub') -f (NF $d.RunMin), (NF $d.ElapsedShiftMin))</div></div>" +
                  "<div class='card'><div class='lbl'>$(T 'card_net')</div><div class='val'>$(PF2 $d.NetPerMin)</div><div class='sub'>$(T 'card_net_sub')</div></div>" +
                  "<div class='card'><div class='lbl'>$(T 'card_loss')</div><div class='val behind'>$(NF $d.LostBoxes)</div><div class='sub'>$(T 'card_loss_sub')</div></div>" +
                  "</div>"
            # achterstand-grafiek staat hoger op de pagina: direct na 'dozen per minuut'
            $bh = "<h2 class='sec'>$(T 'sec_behind')</h2><div class='chart'>$(New-BehindChartSvg $d)</div>"
            if ($stopRows) {
                $tbl = "<div class='tw'><table class='shift'><thead><tr><th>$stopHead</th><th>$(T 'th_duration')</th><th>$(T 'th_kind')</th></tr></thead><tbody>$stopRows</tbody></table></div>"
                if ($stopTotal -gt $stopCap) {
                    $moreLbl = (T 'stops_show_all') -f $stopTotal
                    $lessLbl = T 'stops_collapse'
                    $btn = "<button type='button' class='showall' data-more='$moreLbl' data-less='$lessLbl'>$moreLbl</button>"
                    $scr = "<script>(function(){var w=document.getElementById('stopsWrap');if(!w)return;var b=w.querySelector('.showall');if(!b)return;function set(o){w.classList.toggle('open',o);b.textContent=o?b.dataset.less:b.dataset.more;}try{set(localStorage.getItem('bc_stops_open')==='1');}catch(e){}b.addEventListener('click',function(){var o=!w.classList.contains('open');set(o);try{localStorage.setItem('bc_stops_open',o?'1':'0');}catch(e){}});})();</script>"
                    $st += "<div id='stopsWrap'>$tbl$btn</div>$scr"
                } else {
                    $st += $tbl
                }
            }
        }

        if ($d.Rows.Count -gt 0) {
            $rows = ""
            foreach ($r in $d.Rows) {
                $cls  = if ($r.IsMain) { "cur" } else { "" }
                $mark = if ($r.IsMain) { " <span class='nu'>$(T 'kind_nowmark')</span>" }
                        elseif ($r.Count -le 0) { " <span class='soon'>$(T 'kind_notstarted')</span>" }
                        else { "" }
                if ($r.WeekDone) { $mark += " <span class='soon'>$(T 'kind_weekdone')</span>" }
                $prog = if ($r.IsMain -and $d.HasForecast) { NF $d.ProjMain } else { "&mdash;" }
                $planTxt = if ($r.Plan -gt 0) { NF $r.Plan } else { "&mdash;" }
                $restVal = $r.Plan - $r.Count
                # Weekplan al rond -> wat er van DIE smaak nog in het dagplan staat is geen
                # achterstand: die dozen komen uit een andere smaak. Dus neutraal, niet rood.
                $restCls = if ($r.WeekDone) { 'ph' } elseif ($restVal -gt 0) { 'behind' } else { 'done' }
                $restTxt = if ($r.Plan -gt 0) { "<span class='$restCls'>$(NF $restVal)</span>" } else { "&mdash;" }
                $rows += "<tr class='$cls'><td class='sku'>$(HtmlEnc $r.Product)$mark</td><td class='prd'>$(HtmlEnc $r.Desc)</td><td class='num'>$(NF $r.Count)</td><td class='num'>$planTxt</td><td class='num'>$restTxt</td><td class='num'>$prog</td><td class='pct'>$(& $Bar ([double]$r.Count) ([double]$r.Plan) $false)</td></tr>"
            }
            $totProj = if ($d.HasForecast) { NF $d.ProjTotal } else { "&mdash;" }
            $totRest = $d.Target - $d.Total
            $totalRow = "<tr class='tot'><td colspan='2'>$(T 'total')</td><td class='num'>$(NF $d.Total)</td><td class='num'>$(NF $d.Target)</td><td class='num'>$(NF $totRest)</td><td class='num'>$totProj</td><td class='pct'>$(& $Bar ([double]$d.Total) ([double]$d.Target) $true)</td></tr>"
            $table = "<h2 class='sec'>$(T 'sec_products') <span class='bron'>$((T 'plan_bron') -f $tsrc)</span></h2>" +
                     "<div class='tw'><table class='shift'><thead><tr><th>$(T 'th_producttype')</th><th class='prd'>$(T 'th_product')</th><th class='num'>$(T 'th_boxes_now')</th><th class='num'>$(T 'th_plan_shift')</th><th class='num'>$(T 'th_todo')</th><th class='num wr'>$(T 'th_expected_end')</th><th>$(T 'th_progress')</th></tr></thead>" +
                     "<tbody>$rows$totalRow</tbody></table></div>"
            # per smaak die het weekplan al rond had: waarom haar dagplan blijft staan maar niet gedraaid wordt
            foreach ($wn in @($d.WeekDoneNotes)) {
                $table += "<div class='wdesc'>$((T 'weekdone_note') -f (HtmlEnc $wn.Sku), (NF $wn.MadeBefore), (NF $wn.WeekPlan), (NF $wn.Left), (NF $d.Target))</div>"
            }
            $skipTxt  = if ($d.StartRowSkipped) { T 'skip_txt' } else { "" }
            if ($d.BlankRows -gt 0) { $skipTxt += (T 'skip_blank') -f $d.BlankRows }
            $lastHtml = if ($d.LastText) { "<div class='wdesc'>$((T 'last_box') -f $lastBoxTxt, $d.ParsedRows, $skipTxt)</div>" } else { "" }
        }
        else {
            $table = "<div class='warn'>$((T 'empty_state') -f $d.ParsedRows, (NF $d.Target), $tsrc)</div>"
            $lastHtml = ""
        }

        # ---- weekplan per smaak: waarom een omstelling soms wacht tot het weekplan rond is ----
        $wkp = ""
        if ($d.HasWeekPlan) {
            $cn2 = $script:CultMap[$lang]
            $cu2 = try { [System.Globalization.CultureInfo]::GetCultureInfo($cn2) } catch { $script:nl }
            $wRows = ""
            foreach ($w in $d.WeekRows) {
                $cls  = if ($w.IsCur) { "cur" } else { "" }
                $mark = if ($w.IsCur) { " <span class='nu'>$(T 'kind_nowmark')</span>" }
                        elseif ($w.Made -le 0) { " <span class='soon'>$(T 'kind_notstarted')</span>" }
                        else { "" }
                # wacht in het rooster: wanneer de smaak gepland staat (daarop is de volgorde gebaseerd)
                # Ligt dat geplande begin AL achter ons, dan leest een kale 'vr 13:00' als een
                # afspraak in de TOEKOMST terwijl de run in deze ploeg hoort te lopen. In dat geval
                # 'vanaf ...' met een eigen kleur.
                if ($w.OrderGroup -eq 1 -and $null -ne $w.NextAt) {
                    $nxTxt = $w.NextAt.ToString('ddd HH:mm', $cu2)
                    if ($w.NextPast) { $mark += " <span class='nxt nxtr' title='$(HtmlEnc ((T 'tt_next_from') -f $nxTxt))'>$(HtmlEnc ((T 'nxt_from') -f $nxTxt))</span>" }
                    else             { $mark += " <span class='nxt' title='$(HtmlEnc ((T 'tt_next_at') -f $nxTxt))'>$(HtmlEnc $nxTxt)</span>" }
                }
                $plTxt = if ($w.Plan -gt 0) { NF $w.Plan } else { "&mdash;" }
                $rest  = [double]$w.Plan - [double]$w.Made
                $rTxt  = if ($w.Plan -gt 0) { "<span class='$(if ($rest -gt 0) { 'behind' } else { 'done' })'>$(NF $rest)</span>" } else { "&mdash;" }
                $wRows += "<tr class='$cls'><td class='sku'>$(HtmlEnc $w.Sku)$mark</td><td class='prd'>$(HtmlEnc $w.Desc)</td>" +
                          "<td class='num'>$(NF $w.Made)</td><td class='num'>$plTxt</td><td class='num'>$rTxt</td>" +
                          "<td class='pct'>$(& $Bar ([double]$w.Made) ([double]$w.Plan) $false)</td></tr>"
            }
            $wRest = [double]$d.WeekPlanTotal - [double]$d.WeekMadeTotal
            $wRows += "<tr class='tot'><td colspan='2'>$(T 'total')</td><td class='num'>$(NF $d.WeekMadeTotal)</td>" +
                      "<td class='num'>$(NF $d.WeekPlanTotal)</td><td class='num'>$(NF $wRest)</td>" +
                      "<td class='pct'>$(& $Bar ([double]$d.WeekMadeTotal) ([double]$d.WeekPlanTotal) $true)</td></tr>"
            $wFrom = $d.HistWeekStart.ToString('dd/MM', $cu2)
            $wTo   = $d.HistWeekStart.AddDays(6).ToString('dd/MM', $cu2)
            $wkp = "<h2 class='sec'>$(T 'sec_weekplan') <span class='bron'>$((T 'week_bron') -f $d.WeekNo, (HtmlEnc $wFrom), (HtmlEnc $wTo))</span></h2>" +
                   "<div class='tw'><table class='shift'><thead><tr><th>$(T 'th_producttype')</th><th class='prd'>$(T 'th_product')</th>" +
                   "<th class='num'>$(T 'th_made_week')</th><th class='num'>$(T 'th_plan_week')</th><th class='num'>$(T 'th_rest_week')</th>" +
                   "<th>$(T 'th_progress')</th></tr></thead><tbody>$wRows</tbody></table></div>" +
                   "<div class='wdesc'>$(T 'order_note')</div>" +
                   "<div class='wdesc'>$((T 'week_note') -f (HtmlEnc $d.Sheet), (HtmlEnc $d.HistSheet))</div>"
        }

        # ---- historie: dag/ploeg-raster uit het RCDB-blad (zelfde bron als SAPSTATus) ----
        $hist = ""
        if ($d.HasHistory) {
            $cn = $script:CultMap[$lang]
            $cult = try { [System.Globalization.CultureInfo]::GetCultureInfo($cn) } catch { $script:nl }
            $hrows = ""; $prevDay = $null; $prevWeek = -1
            foreach ($h in $d.History) {
                # kop bij elke OUDERE week (de lopende week staat al in de sectietitel)
                if ($h.WeekIdx -ne $prevWeek) {
                    if ($h.WeekIdx -gt 0) {
                        $wEnd = $h.WeekStart.AddDays(6)
                        $wTxt = ((T 'hist_week') -f (Get-IsoWeek $h.WeekStart.AddDays(1))) +
                                (' &middot; {0} &ndash; {1}' -f $h.WeekStart.ToString('dd/MM', $cult), $wEnd.ToString('dd/MM', $cult))
                        $hrows += "<tr class='wh' data-w='$($h.WeekIdx)' style='display:none'><td colspan='6'>$wTxt</td></tr>"
                    }
                    $prevWeek = $h.WeekIdx; $prevDay = $null
                }
                $dayKey = $h.Date.ToString('yyyy-MM-dd')
                $newDay = ($dayKey -ne $prevDay)
                $prevDay = $dayKey
                $isNow = ($h.Date -eq $d.ShiftStart.Date -and $h.ShiftNo -eq $d.ShiftNo)
                $cls = @(); if ($newDay) { $cls += 'daybreak' }; if ($isNow) { $cls += 'cur' }
                $dayTxt = if ($newDay) { $h.Date.ToString('ddd dd/MM', $cult) } else { "" }
                $mark = if ($isNow) { " <span class='nu'>$(T 'kind_nowmark')</span>" } else { "" }
                $prodTxt = (@($h.Skus | ForEach-Object {
                    $pl = 0.0; if ($h.PlanPerSku.ContainsKey($_.Sku)) { $pl = [double]$h.PlanPerSku[$_.Sku] }
                    $t = if ($pl -gt 0) { "$($_.Sku) $(NF $_.Count)/$(NF $pl)" } else { "$($_.Sku) $(NF $_.Count)" }
                    "<span class='hsku'>$(HtmlEnc $t)</span>"
                }) -join " ")
                $planTxt = if ($h.Target -gt 0) { NF $h.Target } else { "&mdash;" }
                $hide = if ($h.WeekIdx -gt 0) { " style='display:none'" } else { "" }
                $hrows += "<tr class='$($cls -join ' ')' data-w='$($h.WeekIdx)'$hide><td class='sku'>$dayTxt</td>" +
                          "<td class='sku'>$($h.Letter) <span class='bron'>$($h.Range)</span>$mark</td>" +
                          "<td class='num'>$(NF $h.Total)</td><td class='num'>$planTxt</td>" +
                          "<td class='pct'>$(& $Bar ([double]$h.Total) ([double]$h.Target) $false)</td>" +
                          "<td>$prodTxt</td></tr>"
            }
            $fromTxt = if ($d.HistFrom) { $d.HistFrom.ToString('ddd dd/MM', $cult) } else { '' }
            $toTxt   = if ($d.HistTo)   { $d.HistTo.ToString('ddd dd/MM', $cult) }   else { '' }
            $wkTxt   = if ($d.PlanWeek) { ((T 'hist_week') -f $d.PlanWeek) + ' &middot; ' } else { '' }
            $hist = "<h2 class='sec'>$(T 'sec_history') <span class='bron'>$wkTxt$((T 'hist_bron') -f (HtmlEnc $fromTxt), (HtmlEnc $toTxt), (HtmlEnc $d.HistSheet))</span></h2>" +
                    "<div class='tw'><table class='shift hist'><thead><tr><th>$(T 'th_day')</th><th>$(T 'th_shift')</th>" +
                    "<th class='num'>$(T 'th_boxes')</th><th class='num'>$(T 'th_plan_shift')</th><th>$(T 'th_progress')</th>" +
                    "<th>$(T 'th_products')</th></tr></thead><tbody>$hrows</tbody></table></div>"
            # knop '+7 dagen': de oudere weken staan al in de pagina (verborgen), dus dit kost
            # GEEN nieuwe Excel-lees - een klik toont gewoon een week extra.
            if ($d.HistMaxWeek -gt 0) {
                $hist = "<div id='histWrap' data-max='$($d.HistMaxWeek)' data-shown='0'>$hist" +
                        "<button type='button' class='showall'>$(T 'hist_more')</button></div>" +
                        "<script>(function(){var w=document.getElementById('histWrap');if(!w)return;" +
                        "var b=w.querySelector('.showall');if(!b)return;var max=parseInt(w.dataset.max,10)||0;" +
                        "function set(n){if(n>max)n=max;if(n<0)n=0;w.dataset.shown=n;" +
                        "w.querySelectorAll('tr[data-w]').forEach(function(tr){var i=parseInt(tr.dataset.w,10)||0;tr.style.display=(i<=n)?'':'none';});" +
                        "b.style.display=(n>=max)?'none':'';try{localStorage.setItem('bc_hist_weeks',n);}catch(e){}}" +
                        "var s=0;try{s=parseInt(localStorage.getItem('bc_hist_weeks'),10)||0;}catch(e){}set(s);" +
                        "b.addEventListener('click',function(){set((parseInt(w.dataset.shown,10)||0)+1);});})();</script>"
            }
        }
        $bodyHtml = "$cards$chart$fc$bh$st$table$lastHtml$wkp$hist"
    }

    $css = "*{box-sizing:border-box}body{margin:0;background:#0f172a;color:#e2e8f0;font-family:Segoe UI,system-ui,Arial,sans-serif}" +
           ".wrap{max-width:960px;margin:0 auto;padding:24px}h1{font-size:22px;margin:0 0 4px}.meta{color:#94a3b8;font-size:13px;margin-bottom:16px}" +
           ".langbar{display:flex;gap:6px;justify-content:flex-end}" +
           ".topbar{display:flex;gap:14px;align-items:center;justify-content:space-between;margin-bottom:10px;flex-wrap:wrap}" +
           ".linebar{display:flex;gap:6px}" +
           ".lineb{padding:4px 14px;border-radius:8px;background:#1e293b;color:#cbd5e1;text-decoration:none;font-size:13px;font-weight:600;border:1px solid #334155}" +
           ".lineb:hover{background:#273449}" +
           ".lineb.active{background:#2563eb;color:#fff;border-color:#2563eb}" +
           ".lang{display:inline-block;padding:4px 11px;border-radius:8px;background:#1e293b;color:#94a3b8;font-size:12px;font-weight:600;text-decoration:none;border:1px solid #334155}" +
           ".lang:hover{color:#e2e8f0;border-color:#475569}.lang.active{background:#22c55e;color:#06210f;border-color:#22c55e}" +
           ".sec{font-size:14px;color:#cbd5e1;margin:18px 0 8px;text-transform:uppercase;letter-spacing:.03em}" +
           "table.shift{width:100%;border-collapse:collapse;background:#1e293b;border-radius:10px;overflow:hidden}" +
           ".shift th{text-align:left;font-size:11px;text-transform:uppercase;color:#94a3b8;padding:10px 12px;background:#172033}" +
           ".shift td{padding:11px 12px;border-top:1px solid #334155;font-size:16px}.shift td.num{text-align:right;font-variant-numeric:tabular-nums}" +
           # data-cellen altijd op EEN regel; alleen de lange koptekst mag afbreken (scheelt kolombreedte)
           ".shift th.num{text-align:right}.shift td{white-space:nowrap}" +
           ".shift th{white-space:nowrap}.shift th.wr{white-space:normal}" +
           # De productomschrijving is de ENIGE cel die mag afbreken; zij vangt de resterende
           # breedte op zodat de tabel past en er GEEN horizontale schuifbalk komt.
           ".shift td.prd{white-space:normal;line-height:1.25;word-break:break-word}" +
           ".shift th.prd{white-space:normal;width:99%}" +
           ".shift td.num,.shift th.num{padding-left:10px;padding-right:10px}" +
           ".shift td.dur{text-align:left;font-variant-numeric:tabular-nums}" +
           ".shift td.sku{font-weight:600}.done{color:#34d399;font-weight:600}.behind{color:#fbbf24;font-weight:600}" +
           ".shift tr.cur td{background:#14321f}" +
           ".nu{background:#22c55e;color:#06210f;font-size:11px;padding:1px 7px;border-radius:8px;margin-left:4px;vertical-align:middle}" +
           ".tot td{font-weight:700;border-top:2px solid #475569;color:#cbd5e1}" +
           ".bron{color:#94a3b8;font-weight:400;font-size:13px;text-transform:none;letter-spacing:0}" +
           ".wdesc{color:#94a3b8;font-size:13px;margin-top:10px}.cards{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:8px}" +
           # Kaarten: kolom-flex zodat de grijze 'sub'-regel van ALLE kaarten op dezelfde hoogte
           # (onderaan) staat; niets mag buiten de kaart vallen -> lange woorden mogen breken.
           ".card{flex:1;min-width:min(200px,100%);background:#1e293b;border-radius:10px;padding:14px 16px;display:flex;flex-direction:column;overflow-wrap:break-word}" +
           ".card.hi{outline:1px solid #334155;background:#172033}" +
           ".card .lbl{font-size:11px;color:#94a3b8;text-transform:uppercase}" +
           # het grote getal + het percentage ernaast: flex met wrap, anders steekt het percentage
           # bij smalle kaarten buiten de rand
           ".card .val{font-size:34px;font-weight:700;margin-top:4px;font-variant-numeric:tabular-nums;display:flex;flex-wrap:wrap;align-items:baseline;gap:0 9px;line-height:1.12}" +
           ".card .sub{font-size:13px;color:#94a3b8;margin-top:auto;padding-top:6px}" +
           ".card .split{margin-top:8px;display:flex;flex-direction:column;gap:3px}" +
           # label links, waarde rechts; past het niet naast elkaar, dan zakt de waarde netjes
           # naar de volgende regel (rechts uitgelijnd) i.p.v. de kaart uit te rekken
           ".card .split>div{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:baseline;gap:0 12px;font-size:13px;line-height:1.5;font-variant-numeric:tabular-nums}" +
           ".card .split>div>:last-child{margin-left:auto;text-align:right}" +
           ".card .split .sk{color:#94a3b8;min-width:0}.card .split b{font-weight:600;color:#e2e8f0}" +
           ".card .val .vpct{font-size:13px;font-weight:600;white-space:normal}" +
           ".card .split .pcs{font-size:12px;margin-left:6px}" +
           ".card .pcl.g{color:#22c55e}.card .pcl.a{color:#fbbf24}.card .pcl.r{color:#ef4444}" +
           ".soon{background:#334155;color:#cbd5e1;font-size:11px;padding:1px 7px;border-radius:8px;margin-left:4px;vertical-align:middle}" +
           # gepland begin van een smaak die in het weekrooster wacht (bv. 'do 14:30')
           ".nxt{border:1px solid #475569;color:#cbd5e1;font-size:11px;font-weight:400;padding:0 6px;border-radius:8px;margin-left:4px;vertical-align:middle;white-space:nowrap;font-variant-numeric:tabular-nums}" +
           ".nxt.nxtr{border-color:#a16207;color:#fbbf24}" +
           ".card .split .tv{white-space:nowrap}.card .split .ph{color:#94a3b8;font-size:12px;margin-left:5px}" +
           # losse regel onder het grote getal (bv. 'Target bereikt om 21:19'): gewoon links
           # uitgelijnd meelopende tekst, breekt netjes af als de kaart smal is
           ".card .eta{margin-top:8px;font-size:13px;color:#94a3b8;line-height:1.5}" +
           ".card .eta b{color:#e2e8f0;font-weight:600;font-variant-numeric:tabular-nums}" +
           ".card .eta .ph{color:#94a3b8;font-size:12px;margin-left:4px}" +
           # uitvoering-kolom: cijfer BOVEN de balk (smalle kolom, anders wordt de tabel te breed)
           ".tw{overflow-x:auto}.shift td.pct{width:138px}.pw{display:flex;align-items:center;gap:8px}" +
           ".pw .pv{font-size:13px;font-variant-numeric:tabular-nums;white-space:nowrap;min-width:50px;text-align:right}" +
           ".pw.strong .pv{font-weight:700}" +
           ".pw .bar{flex:1;min-width:48px;height:8px;background:#334155;border-radius:5px;overflow:hidden}" +
           ".pw .bar i{display:block;height:100%;border-radius:5px}" +
           ".pw .pv.g{color:#22c55e}.pw .pv.a{color:#fbbf24}.pw .pv.r{color:#ef4444}" +
           ".pw .bar i.g{background:#22c55e}.pw .bar i.a{background:#fbbf24}.pw .bar i.r{background:#ef4444}" +
           ".shift td.pct .dash{color:#64748b}" +
           ".hist td{font-size:14px}.hist tr.daybreak td{border-top:2px solid #475569}" +
           ".hist tr.wh td{background:#172033;color:#cbd5e1;font-size:12px;text-transform:uppercase;letter-spacing:.04em;padding:7px 12px;border-top:2px solid #475569}" +
           "#histWrap .showall{margin:10px 0 0}" +
           ".hist .hsku{display:inline-block;background:#172033;border:1px solid #334155;border-radius:7px;padding:1px 7px;margin:1px 4px 1px 0;font-size:12px;font-variant-numeric:tabular-nums;white-space:nowrap}" +
           ".chart{background:#0b1220;border-radius:10px;padding:8px 8px 2px;overflow-x:auto}.chart svg{display:block}" +
           ".warn{background:#3a2e10;color:#fcd34d;padding:10px 12px;border-radius:8px;margin:10px 0}" +
           ".err{background:#3a1212;color:#fca5a5;padding:10px 12px;border-radius:8px;margin:10px 0}" +
           "#stopsWrap .stop-extra{display:none}#stopsWrap.open .stop-extra{display:table-row}" +
           ".showall{margin:10px 0 0;background:#1e293b;border:1px solid #334155;color:#cbd5e1;font-size:13px;padding:7px 14px;border-radius:8px;cursor:pointer}.showall:hover{border-color:#475569;color:#e2e8f0}" +
           ".foot{color:#64748b;font-size:12px;margin-top:22px}" +
           # --- RESPONSIEF: tabellen moeten PASSEN, geen horizontale schuifbalk ---
           # Staat bewust HELEMAAL achteraan: media-queries verhogen de specificiteit niet, dus
           # regels als '.hist td{font-size:14px}' zouden ze anders weer overrulen.
           # Trapsgewijs: padding/letters kleiner -> koppen mogen afbreken -> balkje weg, enkel %.
           "@media(max-width:1100px){.shift th{font-size:10px;padding:8px 8px}.shift td,.hist td{font-size:15px;padding:10px 8px}" +
           ".shift td.num,.shift th.num{padding-left:7px;padding-right:7px}.shift td.pct{width:118px}" +
           ".pw{gap:6px}.pw .pv{min-width:44px;font-size:12px}.pw .bar{min-width:30px}}" +
           "@media(max-width:800px){.shift th{font-size:9px;padding:7px 5px;white-space:normal}" +
           ".shift td,.hist td{font-size:13px;padding:8px 5px}.shift td.sku{white-space:normal}.nu,.soon,.nxt{margin-left:0}" +
           ".shift td.num,.shift th.num{padding-left:4px;padding-right:4px}.shift td.pct{width:86px}" +
           ".pw .pv{min-width:36px;font-size:11px}.pw .bar{min-width:16px;height:6px}" +
           ".hist .hsku{font-size:11px;padding:1px 5px}}" +
           "@media(max-width:560px){.shift th{font-size:8px;padding:6px 3px;letter-spacing:0}" +
           ".shift td,.hist td{font-size:11.5px;padding:6px 3px}" +
           ".shift td.num,.shift th.num{padding-left:3px;padding-right:3px}" +
           # op een telefoon is het balkje luxe: alleen het percentage, kolom zo smal mogelijk
           ".shift td.pct{width:1%}.pw .bar{display:none}.pw .pv{min-width:0;font-size:11px}" +
           ".nu,.soon,.nxt{font-size:9px;padding:0 4px}.hist .hsku{font-size:10px;padding:0 4px;margin:1px 2px 1px 0}}"

    $langScript = "<script>(function(){var p=new URLSearchParams(location.search);var l=p.get('lang');if(l){try{localStorage.setItem('bc_lang',l)}catch(e){}}else{try{var s=localStorage.getItem('bc_lang');if(s&&s!=='$($script:DefaultLang)'){location.replace('/?lang='+s)}}catch(e){}}})();</script>"

    return "<!doctype html><html lang='$lang'><head><meta charset='utf-8'><meta http-equiv='refresh' content='$refresh; url=/?lang=$lang'>" +
           "<meta name='viewport' content='width=device-width, initial-scale=1'><title>BoxCount $(HtmlEnc $d.Sheet)</title><style>$css</style></head>" +
           "<body><div class='wrap'><div class='topbar'>$linebar$langbar</div><h1>$(T 'h1_l9')</h1><div class='meta'>$meta</div>$warnHtml$bodyHtml" +
           "<div class='foot'>$((T 'foot_loaded') -f $load, $refresh)</div>$langScript</div></body></html>"
}

# ============================== WEERGAVE LIJN 11 ==============================
function Render-Html11($d, [string]$lang = 'nl') {
    if ($script:Langs -notcontains $lang) { $lang = $script:DefaultLang }
    $script:Lang = $lang
    # de server houdt de data zelf vers (leest bij bestandswijziging), dus de pagina mag vaker en
    # goedkoop verversen; begrensd op [5..15]s zodat een wijziging snel zichtbaar wordt
    $refresh = [Math]::Max(5, [Math]::Min([int]$IntervalSeconds, 15))
    $load = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')

    $langbar = "<div class='langbar'>"
    foreach ($lc in $script:Langs) {
        $cls = if ($lc -eq $lang) { 'lang active' } else { 'lang' }
        $langbar += "<a class='$cls' href='/?lang=$lc&amp;line=$($script:Line)'>$($lc.ToUpper())</a>"
    }
    $langbar += "</div>"
    # ---- LIJNKIEZER: zelfde pagina, andere lijn (?line=9 / ?line=11) ----
    $linebar = "<div class='linebar'>"
    foreach ($ld in $script:LineOrder) {
        $lcls = if ($ld -eq $script:Line) { 'lineb active' } else { 'lineb' }
        $linebar += "<a class='$lcls' href='/?lang=$lang&amp;line=$ld'>$((T 'line_word') -f $ld)</a>"
    }
    $linebar += "</div>"

    # ploegletter (X/Y/Z, wisselt per week) i.p.v. het nummer; let op: $shiftLabel gaat door
    # HtmlEnc, dus hier GEEN html-entiteiten gebruiken
    $shiftLabel = "$(T 'shift_word') $($d.ShiftLetter) ($($d.ShiftRange))"
    $meta = "$(HtmlEnc $d.NowText) &middot; $(HtmlEnc $shiftLabel) &middot; $(T 'interval') $(HtmlEnc $d.WindowText)"
    $meta += " &middot; $(HtmlEnc $d.Sheet) &middot; $(HtmlEnc $d.BoxFile)"

    # herkomst van de target (weekplan / parameter / config), taalafhankelijk opgebouwd
    $tsrc = switch ($d.TargetMode) {
        'plan' {
            $cn = $script:CultMap[$lang]; $dt = ''
            if ($d.PlanDate) { try { $dt = $d.PlanDate.ToString('ddd dd/MM', [System.Globalization.CultureInfo]::GetCultureInfo($cn)) } catch { $dt = $d.PlanDate.ToString('dd/MM') } }
            (T 'ts_plan') -f (HtmlEnc $d.PlanFileName), $d.PlanWeek, (HtmlEnc $dt), $d.ShiftNo
        }
        'planfallback' { (T 'ts_planfallback') -f (HtmlEnc $d.PlanFileName), (HtmlEnc $d.PlanSku) }
        'param'  { T 'ts_param' }
        'noplan' { T 'ts_noplan' }
        default  { T 'ts_unknown' }
    }

    # gestapelde waarschuwingen (taalafhankelijk)
    $warnHtml = ""
    if ($d.WarnList -and @($d.WarnList).Count -gt 0) {
        foreach ($w in $d.WarnList) {
            $vals = @($w.Vals | ForEach-Object { HtmlEnc ([string]$_) })
            $wtxt = if ($vals.Count -gt 0) { (T $w.Key) -f $vals } else { (T $w.Key) }
            $warnHtml += "<div class='warn'>$wtxt</div>"
        }
    }

    if ($d.Error) {
        $bodyHtml = "<div class='err'>$(T 'err_prefix'): $(HtmlEnc $d.Error)</div>"
    }
    else {
        $pct = PF $d.Pct
        # draaide er meer dan 1 producttype, dan onder het totaal per type uitsplitsen
        # (bij 1 product zegt zo'n regel niets - dan blijft de kaart zoals hij was)
        # kleurdrempels voor 'hoeveel % van het plan': groen >= 100, oranje >= 80, anders rood
        $PctCls = { param([double]$p) if ($p -ge 100) { 'g' } elseif ($p -ge 80) { 'a' } else { 'r' } }
        # balkje 'hoeveel van het plan is gemaakt' (producttabel EN historie).
        # De balk loopt tot 100 %; wat erboven zit staat in het cijfer (bv. 181 %).
        $Bar = {
            param([double]$count, [double]$plan, [bool]$bold)
            if ($plan -le 0) { return "<span class='dash'>&mdash;</span>" }
            $p = 100.0 * $count / $plan
            $cls = & $PctCls $p
            $w = [Math]::Min(100.0, [Math]::Max(0.0, $p))
            $bcls = if ($bold) { " strong" } else { "" }
            return "<span class='pw$bcls'><span class='pv $cls'>$(PF $p)&nbsp;%</span><span class='bar'><i class='$cls' style='width:$(SvgN $w)%'></i></span></span>"
        }
        # 'laatste doos' in de taal van de pagina (staat al html-veilig, dus NIET nog eens encoderen)
        $lastBoxTxt = if ($d.LastTimeText) {
            (T 'last_box_txt') -f (HtmlEnc $d.LastTimeText), (HtmlEnc $d.LastProduct), (HtmlEnc $d.LastCounter)
        } else { HtmlEnc $d.LastText }

        # De KAARTEN tonen alleen nog de grote totalen. De uitsplitsing per smaak stond hier eerst
        # als lijstje IN de kaart; bij vijf smaken tegelijk brak dat af en werd het onleesbaar.
        # Alles per smaak staat nu in de tabel 'Producten' hieronder, een regel per smaak.
        # naast het grote totaal: hoeveel procent van het plan van de hele ploeg is dat
        $madePct = ""
        if ($d.Target -gt 0) {
            $madePct = "<span class='vpct pcl $(& $PctCls ([double]$d.Pct))'>$((T 'card_made_pct') -f $pct)</span>"
        }
        $tempoSplitRows = ""
        # tempo van de laatste -RecentMinutes minuten hoort bij 'Tempo nu' (de grote waarde = hele ploeg)
        if ($d.HasRecent) {
            $recTtl = (T 'tt_tempo') -f (NF $d.RecentCount), (NF $d.RecentWin)
            $tempoSplitRows += "<div title='$(HtmlEnc $recTtl)'><span class='sk'>$((T 'card_last_win') -f (NF $d.RecentWin))</span>" +
                               "<span class='tv'><b>$(PF2 $d.RecentPerMin)</b><span class='ph'>$(NF ($d.RecentPerMin * 60))$(T 'per_hour_short')</span></span></div>"
        }
        $tempoSplit = if ($tempoSplitRows) { "<div class='split'>$tempoSplitRows</div>" } else { "" }
        $cards = "<div class='cards'>" +
            "<div class='card'><div class='lbl'>$(T 'card_made')</div><div class='val done'>$(NF $d.Total)$madePct</div><div class='sub'>$((T 'card_made_sub') -f (HtmlEnc $d.FileTimeText))</div></div>" +
            "<div class='card'><div class='lbl'>$(T 'card_target')</div><div class='val'>$(NF $d.Target)</div><div class='sub'>$((T 'card_target_sub') -f $pct, (PF2 $d.TargetPerMin))<br><span class='bron'>$tsrc</span></div></div>"
        if ($d.HasForecast) {
            $projCls = if ($d.ProjDiff -ge 0) { "done" } else { "behind" }
            $diffTxt = if ($d.ProjDiff -ge 0) { "+$(NF $d.ProjDiff)" } else { (NF $d.ProjDiff) }
            # ETA ("target bereikt om") hoort bij de verwachte eindstand -> als regel in dezelfde kaart
            $etaSplit = ""
            if ($d.EtaText) {
                # Een enkele regel LINKS uitgelijnd (geen label-links/waarde-rechts: met maar een
                # waarde oogt dat scheef). Tijd = vet, '(na ploegeinde)' = kleine grijze noot.
                # Bij een tekstuitkomst vervalt het label - 'Target bereikt om target al gehaald'
                # zou onzin zijn - en kleurt de uitkomst zelf.
                $etaTxt = switch ($d.EtaKind) {
                    'done'       { "<span class='done'>$(T 'eta_done')</span>" }
                    'over'       { "<b>$(T 'eta_over')</b>" }
                    'impossible' { "<span class='behind'>$(T 'eta_impossible')</span>" }
                    'time'       { "$(T 'card_eta') <b>$(HtmlEnc $d.EtaTimeText)</b>$(if ($d.EtaAfterShift) { "<span class='ph'>$((T 'eta_after').Trim())</span>" } else { '' })" }
                    default      { "<b>$(HtmlEnc $d.EtaText)</b>" }
                }
                $etaSplit = "<div class='eta' title='$(HtmlEnc (T 'card_eta_sub'))'>$etaTxt</div>"
            }
            $cards += "<div class='card hi'><div class='lbl'>$(T 'card_expected_end')</div><div class='val $projCls'>$(NF $d.ProjTotal)</div>$etaSplit" +
                      "<div class='sub'>$((T 'card_expected_end_sub') -f (PF $d.ProjPct), $diffTxt)</div></div>" +
                      "<div class='card'><div class='lbl'>$(T 'card_tempo_now')</div><div class='val'>$(PF2 $d.PerMin)</div>$tempoSplit<div class='sub'>$((T 'card_tempo_now_sub') -f (NF $d.PerHour))</div></div>"
        }
        $cards += "</div>"
        $chart = "<h2 class='sec'>$(T 'sec_per_minute')</h2><div class='chart'>$(New-MinuteChartSvg $d)</div>"

        # GEEN cumulatieve prognosegrafiek op lijn 11 (weggehaald op verzoek): de verwachte
        # eindstand staat als getal in de kaart bovenaan en per smaak/machine in de tabellen,
        # en de grafiek 'achterstand t.o.v. target-tempo' hieronder toont dezelfde afwijking
        # scherper. De melding 'prognose nog voorlopig' loopt nu via de waarschuwingsstrook.

        # ---- MACHINES: het lijn-11-blok. Per machine smaak, tempo, draaitijd en prognose ----
        $mach = ""
        if ($d.HasMach) {
            $mRows = ""
            foreach ($m in $d.Machines) {
                $stCls  = if ($m.IsStill) { "behind" } elseif ($d.NowStill) { "" } else { "done" }
                $stTxt  = if ($m.IsStill)    { (T 'st_still') -f (NF $m.StillMin) }
                          elseif ($d.NowStill) { (T 'st_still') -f (NF $d.StillMin) }
                          else                 { T 'st_running' }
                $badge  = if ($m.IsStill -or $d.NowStill) { "" } else { " <span class='nu'>$(T 'kind_nowmark')</span>" }
                $mName  = if ($d.SkuNames.ContainsKey($m.Product)) { [string]$d.SkuNames[$m.Product] } else { '' }
                $flav   = "<span class='hsku'>$(HtmlEnc $m.Product)</span>"
                if ($mName) { $flav += " $(HtmlEnc $mName)" }
                if ($m.ProdCount -gt 1) {
                    # binnen de ploeg omgesteld: alle smaken tonen, de lopende voorop
                    $others = @($m.ProdList | Where-Object { $_.Sku -ne $m.Product } |
                                ForEach-Object { "<span class='hsku old'>$(HtmlEnc $_.Sku) $(NF $_.Count)</span>" })
                    $flav += ' ' + ($others -join ' ')
                }
                $avCls = if ($m.AvailPct -ge 90) { "done" } else { "behind" }
                $mTtl = (T 'tt_machine') -f (HtmlEnc $m.FirstText), (HtmlEnc $m.LastText), (PF2 $m.NetPerMin), (Format-Dur $m.LongestMin)
                $mRows += "<tr title='$(HtmlEnc $mTtl)'><td class='sku'>$(HtmlEnc $m.Machine)$badge</td><td class='prd'>$flav</td>" +
                          "<td class='num'>$(NF $m.Count)</td><td class='num'>$(PF2 $m.PerMin)</td>" +
                          "<td class='num'><span class='$avCls'>$(PF $m.AvailPct)&nbsp;%</span></td>" +
                          "<td class='num'>$(NF $m.StopMin) <span class='bron'>($($m.StopCount))</span></td>" +
                          "<td class='num'>$(NF $m.Proj)</td><td><span class='$stCls'>$(HtmlEnc $stTxt)</span></td></tr>"
            }
            $mRows += "<tr class='tot'><td colspan='2'>$(T 'total')</td><td class='num'>$(NF $d.Total)</td>" +
                      "<td class='num'>$(PF2 $d.PerMin)</td><td class='num'>$(PF $d.AvailPct)&nbsp;%</td>" +
                      "<td class='num'>$(NF $d.MachStopMin)</td><td class='num'>$(if ($d.HasForecast) { NF $d.ProjTotal } else { '&mdash;' })</td>" +
                      "<td>$($d.MachRunning)/$($d.MachTotal)</td></tr>"
            $mach = "<h2 class='sec'>$(T 'sec_machines') <span class='bron'>$((T 'mach_bron') -f $d.MachTotal, $d.MachRunning, $d.MachStopped)</span></h2>" +
                    "<div class='chart'>$(New-MachineStripSvg $d)</div>" +
                    "<div class='tw'><table class='shift'><thead><tr><th>$(T 'th_machine')</th><th class='prd'>$(T 'th_flavour')</th>" +
                    "<th class='num'>$(T 'th_boxes_now')</th><th class='num'>$(T 'th_tempo')</th>" +
                    "<th class='num'>$(T 'th_runtime_col')</th><th class='num'>$(T 'th_stop_col')</th>" +
                    "<th class='num wr'>$(T 'th_expected_end')</th><th>$(T 'th_status')</th></tr></thead>" +
                    "<tbody>$mRows</tbody></table></div>"
        }

        $st = ""
        $bh = ""
        if ($d.HasStops) {
            $stillCls = if ($d.MachStopMin -gt 0) { "behind" } else { "done" }
            $availCls = if ($d.AvailPct -ge 90) { "done" } else { "behind" }
            # 'nu stil' geldt hier voor de HELE lijn - een enkele machine die stilstaat staat in
            # de machinetabel, niet als rode balk bovenaan.
            $nowStillHtml = if ($d.NowStill) { "<div class='err'>$((T 'nowstill_l11') -f (NF $d.StillMin), $lastBoxTxt)</div>" } else { "" }
            $allStops = @($d.Stops | Sort-Object From)
            $stopTotal = $allStops.Count
            $stopCap = 24
            $stopRows = ""
            for ($si = 0; $si -lt $stopTotal; $si++) {
                $s = $allStops[$si]
                $kindTxt = switch ($s.Kind) { 'opstart' { T 'kind_startup' } 'nu' { T 'kind_nowstill' } default { T 'kind_stop' } }
                if ($s.IsLongest) { $kindTxt += " <span class='nu'>$(T 'kind_longest')</span>" }
                $rowCls = if ($si -ge $stopCap) { " class='stop-extra'" } else { "" }
                $stopRows += "<tr$rowCls><td class='sku'>$($s.From.ToString('HH:mm')) &ndash; $($s.To.ToString('HH:mm'))</td><td class='dur'>$(Format-Dur $s.Min)</td><td>$kindTxt</td></tr>"
            }
            $stopHead = (T 'stops_head_line') -f $stopTotal
            $st = "<h2 class='sec'>$(T 'stops_word') <span class='bron'>$((T 'stops_bron') -f (PF $d.StopLimit))</span></h2>$nowStillHtml" +
                  "<div class='cards'>" +
                  "<div class='card'><div class='lbl'>$(T 'stops_word')</div><div class='val $stillCls'>$(NF $d.MachStopMin)</div><div class='sub'>$((T 'card_stops_sub_l11') -f $d.MachStopCount, $d.MachStopped, $d.MachTotal)</div></div>" +
                  "<div class='card'><div class='lbl'>$(T 'card_runtime')</div><div class='val $availCls'>$(PF $d.AvailPct)</div><div class='sub'>$((T 'card_runtime_sub_l11') -f (NF $d.RunMin), (NF $d.ElapsedShiftMin))</div></div>" +
                  "<div class='card'><div class='lbl'>$(T 'card_net')</div><div class='val'>$(PF2 $d.NetPerMin)</div><div class='sub'>$(T 'card_net_sub_l11')</div></div>" +
                  "<div class='card'><div class='lbl'>$(T 'card_behind')</div><div class='val $(if ($d.BehindNow -ge 0) { 'done' } else { 'behind' })'>$(if ($d.BehindNow -ge 0) { '+' } else { '' })$(NF $d.BehindNow)</div><div class='sub'>$(T 'card_behind_sub')</div></div>" +
                  "</div>"
            # achterstand-grafiek staat hoger op de pagina: direct na de prognose einde ploeg
            $bh = "<h2 class='sec'>$(T 'sec_behind')</h2><div class='chart'>$(New-BehindChartSvg $d)</div>"
            # tabel = alleen de momenten dat de HELE lijn stillag (zeldzaam, dus kort)
            if ($stopRows) {
                $tbl = "<div class='tw'><table class='shift'><thead><tr><th>$stopHead</th><th>$(T 'th_duration')</th><th>$(T 'th_kind')</th></tr></thead><tbody>$stopRows</tbody></table></div>"
                if ($stopTotal -gt $stopCap) {
                    $moreLbl = (T 'stops_show_all') -f $stopTotal
                    $lessLbl = T 'stops_collapse'
                    $btn = "<button type='button' class='showall' data-more='$moreLbl' data-less='$lessLbl'>$moreLbl</button>"
                    $scr = "<script>(function(){var w=document.getElementById('stopsWrap');if(!w)return;var b=w.querySelector('.showall');if(!b)return;function set(o){w.classList.toggle('open',o);b.textContent=o?b.dataset.less:b.dataset.more;}try{set(localStorage.getItem('bc_stops_open')==='1');}catch(e){}b.addEventListener('click',function(){var o=!w.classList.contains('open');set(o);try{localStorage.setItem('bc_stops_open',o?'1':'0');}catch(e){}});})();</script>"
                    $st += "<div id='stopsWrap'>$tbl$btn</div>$scr"
                } else {
                    $st += $tbl
                }
            }
        }

        # ---- EEN TABEL, EEN REGEL PER SMAAK ----
        # Verving de twee losse tabellen ('per producttype' = ploeg, 'weekplan per smaak' = week)
        # en de ingeklemde lijstjes in de kaarten: bij vijf gelijktijdige smaken viel daar niet in
        # een oogopslag uit te lezen wat een smaak deze week nog moet en wat ze deze ploeg doet.
        # Staat er deze ploeg nog niets, dan komt de melding erboven - de weekkolommen tonen dan
        # nog steeds wat er deze week al gemaakt is.
        $table = ""
        if ($d.Rows.Count -eq 0) {
            $table = "<div class='warn'>$((T 'empty_state') -f $d.ParsedRows, (NF $d.Target), $tsrc)</div>"
        }
        $skipTxt  = if ($d.StartRowSkipped) { T 'skip_txt' } else { "" }
        if ($d.BlankRows -gt 0) { $skipTxt += (T 'skip_blank') -f $d.BlankRows }
        $lastHtml = if ($d.LastText) { "<div class='wdesc'>$((T 'last_box') -f $lastBoxTxt, $d.ParsedRows, $skipTxt)</div>" } else { "" }

        if ($d.HasCombined) {
            $cn2 = $script:CultMap[$lang]
            $cu2 = try { [System.Globalization.CultureInfo]::GetCultureInfo($cn2) } catch { $script:nl }
            $rows = ""
            foreach ($r in $d.Combined) {
                $cls  = if ($r.IsCur) { "cur" } else { "" }
                $mark = if ($r.IsCur) { " <span class='nu'>$(T 'kind_nowmark')</span>" }
                        elseif ($r.Plan -gt 0 -and $r.Count -le 0) { " <span class='soon'>$(T 'kind_notstarted')</span>" }
                        else { "" }
                if ($r.WeekDone) { $mark += " <span class='soon'>$(T 'kind_weekdone')</span>" }
                # wacht in het rooster: wanneer de smaak gepland staat (daarop is de volgorde gebaseerd)
                # Ligt dat geplande begin AL achter ons, dan leest een kale 'vr 13:00' als een
                # afspraak in de TOEKOMST terwijl de run in deze ploeg hoort te lopen. In dat geval
                # 'vanaf ...' met een eigen kleur.
                if ($r.OrderGroup -eq 1 -and $null -ne $r.NextAt) {
                    $nxTxt = $r.NextAt.ToString('ddd HH:mm', $cu2)
                    if ($r.NextPast) { $mark += " <span class='nxt nxtr' title='$(HtmlEnc ((T 'tt_next_from') -f $nxTxt))'>$(HtmlEnc ((T 'nxt_from') -f $nxTxt))</span>" }
                    else             { $mark += " <span class='nxt' title='$(HtmlEnc ((T 'tt_next_at') -f $nxTxt))'>$(HtmlEnc $nxTxt)</span>" }
                }
                # --- deze week ---
                $wPct = ""
                if ($r.WeekPlan -gt 0) {
                    $pv = 100.0 * $r.WeekMade / $r.WeekPlan
                    $wPct = "<span class='pcs pcl $(& $PctCls $pv)'>$(PF $pv)&nbsp;%</span>"
                }
                $wMade = if ($r.HasWeek -or $r.WeekMade -gt 0) { NF $r.WeekMade } else { "&mdash;" }
                $wPlan = if ($r.WeekPlan -gt 0) { NF $r.WeekPlan } else { "&mdash;" }
                $wRest = "&mdash;"
                if ($r.WeekPlan -gt 0) {
                    $rv = [double]$r.WeekPlan - [double]$r.WeekMade
                    $wRest = "<span class='$(if ($rv -gt 0) { 'behind' } else { 'done' })'>$(NF $rv)</span>"
                }
                # --- deze ploeg --- (eigen eindstand per smaak = som van de machines die haar draaien)
                $sPct = ""
                if ($r.Plan -gt 0) {
                    $pv2 = 100.0 * $r.Count / $r.Plan
                    $sPct = "<span class='pcs pcl $(& $PctCls $pv2)'>$(PF $pv2)&nbsp;%</span>"
                }
                $sPlan = if ($r.Plan -gt 0) { NF $r.Plan } else { "&mdash;" }
                $sProj = if ($d.HasForecast -and $r.HasProj -and ($r.Count -gt 0)) { NF $r.Proj } else { "&mdash;" }
                $sTmp  = "&mdash;"
                if ($r.PerMin -gt 0) {
                    $tTtl = (T 'tt_tempo_mach') -f (NF $r.Count), (@($r.Machines).Count), (@($r.Machines) -join ', ')
                    $sTmp = "<span title='$(HtmlEnc $tTtl)'>$(PF2 $r.PerMin)</span>"
                }
                # kolomgroepen: eerst DEZE PLOEG, dan DEZE WEEK (volgorde op vraag van de gebruiker)
                $rows += "<tr class='$cls'><td class='sku'>$(HtmlEnc $r.Sku)$mark</td><td class='prd'>$(HtmlEnc $r.Desc)</td>" +
                         "<td class='num gsep'>$(NF $r.Count)$sPct</td><td class='num'>$sPlan</td><td class='num'>$sProj</td><td class='num'>$sTmp</td>" +
                         "<td class='num gsep'>$wMade$wPct</td><td class='num'>$wPlan</td><td class='num'>$wRest</td></tr>"
            }
            $wtRest  = [double]$d.WeekPlanTotal - [double]$d.WeekMadeTotal
            $wtRTxt  = if ($d.WeekPlanTotal -gt 0) { "<span class='$(if ($wtRest -gt 0) { 'behind' } else { 'done' })'>$(NF $wtRest)</span>" } else { "&mdash;" }
            $wtMade  = if ($d.HasWeekPlan) { NF $d.WeekMadeTotal } else { "&mdash;" }
            $wtPlan  = if ($d.WeekPlanTotal -gt 0) { NF $d.WeekPlanTotal } else { "&mdash;" }
            $totProj = if ($d.HasForecast) { NF $d.ProjTotal } else { "&mdash;" }
            $rows += "<tr class='tot'><td colspan='2'>$(T 'total')</td>" +
                     "<td class='num gsep'>$(NF $d.Total)</td><td class='num'>$(NF $d.Target)</td><td class='num'>$totProj</td>" +
                     "<td class='num'>$(PF2 $d.PerMin)</td>" +
                     "<td class='num gsep'>$wtMade</td><td class='num'>$wtPlan</td><td class='num'>$wtRTxt</td></tr>"

            $wkBron = ""
            if ($d.HasWeekPlan -and $d.HistWeekStart) {
                $wkBron = (T 'week_bron') -f $d.WeekNo, (HtmlEnc $d.HistWeekStart.ToString('dd/MM', $cu2)), (HtmlEnc $d.HistWeekStart.AddDays(6).ToString('dd/MM', $cu2))
            }
            $table += "<h2 class='sec'>$(T 'sec_all_products') <span class='bron'>$wkBron</span></h2>" +
                      "<div class='tw'><table class='shift'><thead>" +
                      "<tr><th rowspan='2'>$(T 'th_producttype')</th><th rowspan='2' class='prd'>$(T 'th_product')</th>" +
                      "<th colspan='4' class='grp gsep'>$(T 'grp_shift')</th><th colspan='3' class='grp gsep'>$(T 'grp_week')</th></tr>" +
                      "<tr><th class='num gsep'>$(T 'th_made_week')</th><th class='num'>$(T 'th_plan')</th><th class='num wr'>$(T 'th_prognose')</th>" +
                      "<th class='num'>$(T 'th_tempo')</th>" +
                      "<th class='num gsep'>$(T 'th_made_week')</th><th class='num'>$(T 'th_plan')</th><th class='num'>$(T 'th_rest_week')</th></tr></thead>" +
                      "<tbody>$rows</tbody></table></div>" +
                      "<div class='wdesc'>$(T 'order_note')</div>" +
                      "<div class='wdesc'>$((T 'week_note') -f (HtmlEnc $d.Sheet), (HtmlEnc $d.HistSheet))</div>"
            # per smaak die het weekplan al rond had: waarom haar dagplan blijft staan maar niet gedraaid wordt
            foreach ($wn in @($d.WeekDoneNotes)) {
                $table += "<div class='wdesc'>$((T 'weekdone_note') -f (HtmlEnc $wn.Sku), (NF $wn.MadeBefore), (NF $wn.WeekPlan), (NF $wn.Left), (NF $d.Target))</div>"
            }
        }
        # ---- historie: dag/ploeg-raster uit het RCDB-blad (zelfde bron als SAPSTATus) ----
        $hist = ""
        if ($d.HasHistory) {
            $cn = $script:CultMap[$lang]
            $cult = try { [System.Globalization.CultureInfo]::GetCultureInfo($cn) } catch { $script:nl }
            $hrows = ""; $prevDay = $null; $prevWeek = -1
            foreach ($h in $d.History) {
                # kop bij elke OUDERE week (de lopende week staat al in de sectietitel)
                if ($h.WeekIdx -ne $prevWeek) {
                    if ($h.WeekIdx -gt 0) {
                        $wEnd = $h.WeekStart.AddDays(6)
                        $wTxt = ((T 'hist_week') -f (Get-IsoWeek $h.WeekStart.AddDays(1))) +
                                (' &middot; {0} &ndash; {1}' -f $h.WeekStart.ToString('dd/MM', $cult), $wEnd.ToString('dd/MM', $cult))
                        $hrows += "<tr class='wh' data-w='$($h.WeekIdx)' style='display:none'><td colspan='6'>$wTxt</td></tr>"
                    }
                    $prevWeek = $h.WeekIdx; $prevDay = $null
                }
                $dayKey = $h.Date.ToString('yyyy-MM-dd')
                $newDay = ($dayKey -ne $prevDay)
                $prevDay = $dayKey
                $isNow = ($h.Date -eq $d.ShiftStart.Date -and $h.ShiftNo -eq $d.ShiftNo)
                $cls = @(); if ($newDay) { $cls += 'daybreak' }; if ($isNow) { $cls += 'cur' }
                $dayTxt = if ($newDay) { $h.Date.ToString('ddd dd/MM', $cult) } else { "" }
                $mark = if ($isNow) { " <span class='nu'>$(T 'kind_nowmark')</span>" } else { "" }
                $prodTxt = (@($h.Skus | ForEach-Object {
                    $pl = 0.0; if ($h.PlanPerSku.ContainsKey($_.Sku)) { $pl = [double]$h.PlanPerSku[$_.Sku] }
                    $t = if ($pl -gt 0) { "$($_.Sku) $(NF $_.Count)/$(NF $pl)" } else { "$($_.Sku) $(NF $_.Count)" }
                    "<span class='hsku'>$(HtmlEnc $t)</span>"
                }) -join " ")
                $planTxt = if ($h.Target -gt 0) { NF $h.Target } else { "&mdash;" }
                $hide = if ($h.WeekIdx -gt 0) { " style='display:none'" } else { "" }
                $hrows += "<tr class='$($cls -join ' ')' data-w='$($h.WeekIdx)'$hide><td class='sku'>$dayTxt</td>" +
                          "<td class='sku'>$($h.Letter) <span class='bron'>$($h.Range)</span>$mark</td>" +
                          "<td class='num'>$(NF $h.Total)</td><td class='num'>$planTxt</td>" +
                          "<td class='pct'>$(& $Bar ([double]$h.Total) ([double]$h.Target) $false)</td>" +
                          "<td>$prodTxt</td></tr>"
            }
            $fromTxt = if ($d.HistFrom) { $d.HistFrom.ToString('ddd dd/MM', $cult) } else { '' }
            $toTxt   = if ($d.HistTo)   { $d.HistTo.ToString('ddd dd/MM', $cult) }   else { '' }
            $wkTxt   = if ($d.PlanWeek) { ((T 'hist_week') -f $d.PlanWeek) + ' &middot; ' } else { '' }
            $hist = "<h2 class='sec'>$(T 'sec_history') <span class='bron'>$wkTxt$((T 'hist_bron') -f (HtmlEnc $fromTxt), (HtmlEnc $toTxt), (HtmlEnc $d.HistSheet))</span></h2>" +
                    "<div class='tw'><table class='shift hist'><thead><tr><th>$(T 'th_day')</th><th>$(T 'th_shift')</th>" +
                    "<th class='num'>$(T 'th_boxes')</th><th class='num'>$(T 'th_plan_shift')</th><th>$(T 'th_progress')</th>" +
                    "<th>$(T 'th_products')</th></tr></thead><tbody>$hrows</tbody></table></div>"
            # knop '+7 dagen': de oudere weken staan al in de pagina (verborgen), dus dit kost
            # GEEN nieuwe Excel-lees - een klik toont gewoon een week extra.
            if ($d.HistMaxWeek -gt 0) {
                $hist = "<div id='histWrap' data-max='$($d.HistMaxWeek)' data-shown='0'>$hist" +
                        "<button type='button' class='showall'>$(T 'hist_more')</button></div>" +
                        "<script>(function(){var w=document.getElementById('histWrap');if(!w)return;" +
                        "var b=w.querySelector('.showall');if(!b)return;var max=parseInt(w.dataset.max,10)||0;" +
                        "function set(n){if(n>max)n=max;if(n<0)n=0;w.dataset.shown=n;" +
                        "w.querySelectorAll('tr[data-w]').forEach(function(tr){var i=parseInt(tr.dataset.w,10)||0;tr.style.display=(i<=n)?'':'none';});" +
                        "b.style.display=(n>=max)?'none':'';try{localStorage.setItem('bc_hist_weeks',n);}catch(e){}}" +
                        "var s=0;try{s=parseInt(localStorage.getItem('bc_hist_weeks'),10)||0;}catch(e){}set(s);" +
                        "b.addEventListener('click',function(){set((parseInt(w.dataset.shown,10)||0)+1);});})();</script>"
            }
        }
        $bodyHtml = "$cards$chart$bh$mach$st$table$lastHtml$hist"
    }

    $css = "*{box-sizing:border-box}body{margin:0;background:#0f172a;color:#e2e8f0;font-family:Segoe UI,system-ui,Arial,sans-serif}" +
           ".wrap{max-width:1200px;margin:0 auto;padding:24px}h1{font-size:22px;margin:0 0 4px}.meta{color:#94a3b8;font-size:13px;margin-bottom:16px}" +
           ".langbar{display:flex;gap:6px;justify-content:flex-end}" +
           ".topbar{display:flex;gap:14px;align-items:center;justify-content:space-between;margin-bottom:10px;flex-wrap:wrap}" +
           ".linebar{display:flex;gap:6px}" +
           ".lineb{padding:4px 14px;border-radius:8px;background:#1e293b;color:#cbd5e1;text-decoration:none;font-size:13px;font-weight:600;border:1px solid #334155}" +
           ".lineb:hover{background:#273449}" +
           ".lineb.active{background:#2563eb;color:#fff;border-color:#2563eb}" +
           ".lang{display:inline-block;padding:4px 11px;border-radius:8px;background:#1e293b;color:#94a3b8;font-size:12px;font-weight:600;text-decoration:none;border:1px solid #334155}" +
           ".lang:hover{color:#e2e8f0;border-color:#475569}.lang.active{background:#22c55e;color:#06210f;border-color:#22c55e}" +
           ".sec{font-size:14px;color:#cbd5e1;margin:18px 0 8px;text-transform:uppercase;letter-spacing:.03em}" +
           "table.shift{width:100%;border-collapse:collapse;background:#1e293b;border-radius:10px;overflow:hidden}" +
           ".shift th{text-align:left;font-size:11px;text-transform:uppercase;color:#94a3b8;padding:10px 12px;background:#172033}" +
           ".shift td{padding:11px 12px;border-top:1px solid #334155;font-size:16px}.shift td.num{text-align:right;font-variant-numeric:tabular-nums}" +
           # data-cellen altijd op EEN regel; alleen de lange koptekst mag afbreken (scheelt kolombreedte)
           ".shift th.num{text-align:right}.shift td{white-space:nowrap}" +
           ".shift th{white-space:nowrap}.shift th.wr{white-space:normal}" +
           # De productomschrijving is de ENIGE cel die mag afbreken; zij vangt de resterende
           # breedte op zodat de tabel past en er GEEN horizontale schuifbalk komt.
           ".shift td.prd{white-space:normal;line-height:1.25;word-break:break-word}" +
           ".shift th.prd{white-space:normal;width:99%}" +
           ".shift td.num,.shift th.num{padding-left:10px;padding-right:10px}" +
           ".shift td.dur{text-align:left;font-variant-numeric:tabular-nums}" +
           ".shift td.sku{font-weight:600}.done{color:#34d399;font-weight:600}.behind{color:#fbbf24;font-weight:600}" +
           ".shift tr.cur td{background:#14321f}" +
           ".nu{background:#22c55e;color:#06210f;font-size:11px;padding:1px 7px;border-radius:8px;margin-left:4px;vertical-align:middle}" +
           ".tot td{font-weight:700;border-top:2px solid #475569;color:#cbd5e1}" +
           ".bron{color:#94a3b8;font-weight:400;font-size:13px;text-transform:none;letter-spacing:0}" +
           ".wdesc{color:#94a3b8;font-size:13px;margin-top:10px}.cards{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:8px}" +
           # Kaarten: kolom-flex zodat de grijze 'sub'-regel van ALLE kaarten op dezelfde hoogte
           # (onderaan) staat; niets mag buiten de kaart vallen -> lange woorden mogen breken.
           ".card{flex:1;min-width:min(200px,100%);background:#1e293b;border-radius:10px;padding:14px 16px;display:flex;flex-direction:column;overflow-wrap:break-word}" +
           ".card.hi{outline:1px solid #334155;background:#172033}" +
           ".card .lbl{font-size:11px;color:#94a3b8;text-transform:uppercase}" +
           # het grote getal + het percentage ernaast: flex met wrap, anders steekt het percentage
           # bij smalle kaarten buiten de rand
           ".card .val{font-size:34px;font-weight:700;margin-top:4px;font-variant-numeric:tabular-nums;display:flex;flex-wrap:wrap;align-items:baseline;gap:0 9px;line-height:1.12}" +
           ".card .sub{font-size:13px;color:#94a3b8;margin-top:auto;padding-top:6px}" +
           ".card .split{margin-top:8px;display:flex;flex-direction:column;gap:3px}" +
           # label links, waarde rechts; past het niet naast elkaar, dan zakt de waarde netjes
           # naar de volgende regel (rechts uitgelijnd) i.p.v. de kaart uit te rekken
           ".card .split>div{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:baseline;gap:0 12px;font-size:13px;line-height:1.5;font-variant-numeric:tabular-nums}" +
           ".card .split>div>:last-child{margin-left:auto;text-align:right}" +
           ".card .split .sk{color:#94a3b8;min-width:0}.card .split b{font-weight:600;color:#e2e8f0}" +
           ".card .val .vpct{font-size:13px;font-weight:600;white-space:normal}" +
           ".card .split .pcs{font-size:12px;margin-left:6px}" +
           ".card .pcl.g{color:#22c55e}.card .pcl.a{color:#fbbf24}.card .pcl.r{color:#ef4444}" +
           ".soon{background:#334155;color:#cbd5e1;font-size:11px;padding:1px 7px;border-radius:8px;margin-left:4px;vertical-align:middle}" +
           ".card .split .tv{white-space:nowrap}.card .split .ph{color:#94a3b8;font-size:12px;margin-left:5px}" +
           # losse regel onder het grote getal (bv. 'Target bereikt om 21:19'): gewoon links
           # uitgelijnd meelopende tekst, breekt netjes af als de kaart smal is
           ".card .eta{margin-top:8px;font-size:13px;color:#94a3b8;line-height:1.5}" +
           ".card .eta b{color:#e2e8f0;font-weight:600;font-variant-numeric:tabular-nums}" +
           ".card .eta .ph{color:#94a3b8;font-size:12px;margin-left:4px}" +
           # uitvoering-kolom: cijfer BOVEN de balk (smalle kolom, anders wordt de tabel te breed)
           ".tw{overflow-x:auto}.shift td.pct{width:138px}.pw{display:flex;align-items:center;gap:8px}" +
           ".pw .pv{font-size:13px;font-variant-numeric:tabular-nums;white-space:nowrap;min-width:50px;text-align:right}" +
           ".pw.strong .pv{font-weight:700}" +
           ".pw .bar{flex:1;min-width:48px;height:8px;background:#334155;border-radius:5px;overflow:hidden}" +
           ".pw .bar i{display:block;height:100%;border-radius:5px}" +
           ".pw .pv.g{color:#22c55e}.pw .pv.a{color:#fbbf24}.pw .pv.r{color:#ef4444}" +
           ".pw .bar i.g{background:#22c55e}.pw .bar i.a{background:#fbbf24}.pw .bar i.r{background:#ef4444}" +
           ".shift td.pct .dash{color:#64748b}" +
           # gegroepeerde kop (WEEK | PLOEG) met een lijntje tussen de groepen
           ".shift th.grp{text-align:center;background:#131c2e;color:#cbd5e1;letter-spacing:.06em}" +
           ".shift td.gsep,.shift th.gsep{border-left:1px solid #475569}" +
           # gepland begin van een smaak die in het weekrooster wacht (bv. 'do 14:30')
           ".nxt{border:1px solid #475569;color:#cbd5e1;font-size:11px;font-weight:400;padding:0 6px;border-radius:8px;margin-left:4px;vertical-align:middle;white-space:nowrap;font-variant-numeric:tabular-nums}" +
           ".nxt.nxtr{border-color:#a16207;color:#fbbf24}" +
           ".shift td .pcs{display:inline-block;min-width:54px;text-align:right;font-size:12px;margin-left:6px;font-weight:600}" +
           ".shift td .pcs.g{color:#22c55e}.shift td .pcs.a{color:#fbbf24}.shift td .pcs.r{color:#ef4444}" +
           ".hist td{font-size:14px}.hist tr.daybreak td{border-top:2px solid #475569}" +
           ".hist tr.wh td{background:#172033;color:#cbd5e1;font-size:12px;text-transform:uppercase;letter-spacing:.04em;padding:7px 12px;border-top:2px solid #475569}" +
           "#histWrap .showall{margin:10px 0 0}" +
           ".shift .hsku{display:inline-block;background:#172033;border:1px solid #334155;border-radius:7px;padding:1px 7px;margin:1px 4px 1px 0;font-size:12px;font-variant-numeric:tabular-nums;white-space:nowrap}" +
           ".shift .hsku.old{color:#94a3b8;border-style:dashed}" +
           ".hist .hsku{display:inline-block;background:#172033;border:1px solid #334155;border-radius:7px;padding:1px 7px;margin:1px 4px 1px 0;font-size:12px;font-variant-numeric:tabular-nums;white-space:nowrap}" +
           ".chart{background:#0b1220;border-radius:10px;padding:8px 8px 2px;overflow-x:auto}.chart svg{display:block}" +
           ".warn{background:#3a2e10;color:#fcd34d;padding:10px 12px;border-radius:8px;margin:10px 0}" +
           ".err{background:#3a1212;color:#fca5a5;padding:10px 12px;border-radius:8px;margin:10px 0}" +
           "#stopsWrap .stop-extra{display:none}#stopsWrap.open .stop-extra{display:table-row}" +
           ".showall{margin:10px 0 0;background:#1e293b;border:1px solid #334155;color:#cbd5e1;font-size:13px;padding:7px 14px;border-radius:8px;cursor:pointer}.showall:hover{border-color:#475569;color:#e2e8f0}" +
           ".foot{color:#64748b;font-size:12px;margin-top:22px}" +
           # --- RESPONSIEF: tabellen moeten PASSEN, geen horizontale schuifbalk ---
           # Staat bewust HELEMAAL achteraan: media-queries verhogen de specificiteit niet, dus
           # regels als '.hist td{font-size:14px}' zouden ze anders weer overrulen.
           # Trapsgewijs: padding/letters kleiner -> koppen mogen afbreken -> balkje weg, enkel %.
           "@media(max-width:1100px){.shift th{font-size:10px;padding:8px 8px}.shift td,.hist td{font-size:15px;padding:10px 8px}" +
           ".shift td.num,.shift th.num{padding-left:7px;padding-right:7px}.shift td.pct{width:118px}" +
           ".pw{gap:6px}.pw .pv{min-width:44px;font-size:12px}.pw .bar{min-width:30px}}" +
           "@media(max-width:800px){.shift th{font-size:9px;padding:7px 5px;white-space:normal}" +
           ".shift td,.hist td{font-size:13px;padding:8px 5px}.shift td.sku{white-space:normal}.nu,.soon,.nxt{margin-left:0}" +
           ".shift td.num,.shift th.num{padding-left:4px;padding-right:4px}.shift td.pct{width:86px}" +
           ".pw .pv{min-width:36px;font-size:11px}.pw .bar{min-width:16px;height:6px}" +
           ".hist .hsku{font-size:11px;padding:1px 5px}}" +
           "@media(max-width:560px){.shift th{font-size:8px;padding:6px 3px;letter-spacing:0}" +
           ".shift td,.hist td{font-size:11.5px;padding:6px 3px}" +
           ".shift td.num,.shift th.num{padding-left:3px;padding-right:3px}" +
           # op een telefoon is het balkje luxe: alleen het percentage, kolom zo smal mogelijk
           ".shift td.pct{width:1%}.pw .bar{display:none}.pw .pv{min-width:0;font-size:11px}" +
           ".nu,.soon,.nxt{font-size:9px;padding:0 4px}.hist .hsku{font-size:10px;padding:0 4px;margin:1px 2px 1px 0}}" +
           # 9 kolommen moeten op een smal scherm passen: percentages zijn dan luxe
           "@media(max-width:900px){.shift td .pcs{display:none}}"

    $langScript = "<script>(function(){var p=new URLSearchParams(location.search);var l=p.get('lang');if(l){try{localStorage.setItem('bc_lang',l)}catch(e){}}else{try{var s=localStorage.getItem('bc_lang');if(s&&s!=='$($script:DefaultLang)'){location.replace('/?lang='+s)}}catch(e){}}})();</script>"

    return "<!doctype html><html lang='$lang'><head><meta charset='utf-8'><meta http-equiv='refresh' content='$refresh; url=/?lang=$lang'>" +
           "<meta name='viewport' content='width=device-width, initial-scale=1'><title>BoxCount L$($script:Line)</title><style>$css</style></head>" +
           "<body><div class='wrap'><div class='topbar'>$linebar$langbar</div><h1>$((T 'h1_machines') -f $script:Line)</h1><div class='meta'>$meta</div>$warnHtml$bodyHtml" +
           "<div class='foot'>$((T 'foot_loaded') -f $load, $refresh)</div>$langScript</div></body></html>"
}

# ================================== WEBSERVER ==================================
function Start-WebServer([int]$port) {
    $listener = New-Object System.Net.Sockets.TcpListener ([System.Net.IPAddress]::Loopback, $port)
    try { $listener.Start() }
    catch { Write-Host "FOUT: kan poort $port niet openen: $($_.Exception.Message)" -ForegroundColor Red; return }
    $url = "http://127.0.0.1:$port/?line=$($script:Line)"
    Write-Host "Webserver actief op $url" -ForegroundColor Green
    Write-Host "(lijn 9 en lijn 11 op dezelfde pagina; wisselen met de knoppen linksboven)" -ForegroundColor DarkGray
    Write-Host "(server bewaakt het bronbestand en herleest ALLEEN bij wijziging; Ctrl+C om te stoppen)" -ForegroundColor DarkGray
    if (-not $NoBrowser) { try { Start-Process $url } catch {} }

    # De cache is nu PER LIJN. Beide lijnen komen uit hetzelfde bronbestand, dus een wijziging
    # (of een ploegwissel) maakt ze allebei ongeldig; de lijn die gevraagd wordt, wordt herlezen.
    # Een lijn die nooit bekeken wordt, kost dus ook nooit een Excel-lees.
    $cache = @{}; $lastStamp = $null; $lastCheck = [datetime]::MinValue

    function Get-Cached([string]$id) {
        if (-not $cache.ContainsKey($id) -or $null -eq $cache[$id]) {
            $cache[$id] = Get-BoxDataFor $id       # DUUR: opent Excel + leest de 4 MB
            $script:LastReadStamp = Get-BoxFileStamp
        }
        return $cache[$id]
    }

    try {
        while ($true) {
            # --- 1) bronbestand bewaken; bij wijziging/ploegwissel de HELE cache legen ---
            if (([datetime]::Now - $lastCheck).TotalSeconds -ge 15) {
                $lastCheck = [datetime]::Now
                $stamp  = Get-BoxFileStamp
                $nowRef = if ($script:HasNow) { $Now } else { Get-Date }
                $winNow = Get-ShiftWindow $nowRef
                $newShift = $false
                foreach ($k in @($cache.Keys)) {
                    if ($null -ne $cache[$k] -and $cache[$k].ShiftStart -ne $winNow.Start) { $newShift = $true }
                }
                if ($newShift -or ($null -ne $stamp -and $null -ne $lastStamp -and $stamp -ne $lastStamp)) {
                    $cache = @{}
                }
                if ($null -ne $stamp) { $lastStamp = $stamp }
            }
            # --- 2) wachtend verzoek bedienen ---
            if ($listener.Pending()) {
                $client = $listener.AcceptTcpClient()
                try {
                    $client.ReceiveTimeout = 1500
                    $stream = $client.GetStream()
                    $buf = New-Object byte[] 4096
                    $reqLen = 0
                    try { $reqLen = $stream.Read($buf, 0, $buf.Length) } catch {}
                    $reqTxt = if ($reqLen -gt 0) { [System.Text.Encoding]::ASCII.GetString($buf, 0, $reqLen) } else { '' }
                    $lang = Get-ReqLang $reqTxt
                    $line = Get-ReqLine $reqTxt
                    $d    = Get-Cached $line
                    if ($null -eq $lastStamp) { $lastStamp = Get-BoxFileStamp }
                    $cacheHtml = Render-HtmlFor $d $lang $line
                    $body = [System.Text.Encoding]::UTF8.GetBytes($cacheHtml)
                    $head = "HTTP/1.1 200 OK`r`nContent-Type: text/html; charset=utf-8`r`nContent-Length: $($body.Length)`r`nCache-Control: no-cache`r`nConnection: close`r`n`r`n"
                    $hb = [System.Text.Encoding]::ASCII.GetBytes($head)
                    $stream.Write($hb, 0, $hb.Length); $stream.Write($body, 0, $body.Length); $stream.Flush()
                }
                catch {}
                finally { $client.Close() }
            }
            else {
                Start-Sleep -Milliseconds 300
            }
        }
    }
    finally { $listener.Stop() }
}

# ==================== configuratie laden: EEN bestand, alle lijnen ====================
# config.txt naast dit script bevat ALLES: bovenaan het algemene blok (bronbestand, planmap,
# standaardinstellingen), daaronder eventueel een blok '[5]' ... '[11]' dat alleen zijn eigen
# lijn overschrijft. Voorrang: opdrachtregel > blok van de lijn > algemeen blok > het register.
# Wat op de opdrachtregel is meegegeven geldt voor ALLE lijnen (ze lezen hetzelfde bronbestand).
$script:BoxFolder = $here
$cfgAll = Read-ConfigFile $ConfigFile

# bronbestand + planmap zijn lijn-overstijgend en staan dus in het algemene blok
if (-not $script:HasBox) {
    $v = Get-CfgVal $cfgAll '' 'BoxPrintingFile'
    if ($v) { $BoxPrintingFile = $v }
}
if (-not $script:HasPlanDir) {
    $v = Get-CfgVal $cfgAll '' 'PlanFolder'
    if ($v) { $script:PlanFolder = $v }
}

foreach ($id in $script:LineOrder) {
    $def = $script:LineDefs[$id]

    # bladen: opdrachtregel > blok van de lijn > algemeen blok > het register hierboven
    $cBs = Get-CfgVal $cfgAll $id 'BoxSheet'
    $cRs = Get-CfgVal $cfgAll $id 'RcdbSheet'
    $bs = if ($script:HasBoxSheet) { $BoxSheet }
          elseif ($cBs) { $cBs }
          else { $def.BoxSheets -join ',' }
    $rs = if ($script:HasRcdb) { $RcdbSheet }
          elseif ($cRs) { $cRs }
          else { $def.RcdbSheets -join ',' }

    $cTg = Get-CfgVal $cfgAll $id 'ShiftTarget'
    $cRm = Get-CfgVal $cfgAll $id 'RecentMinutes'
    $cSm = Get-CfgVal $cfgAll $id 'StopMinutes'
    $cHd = Get-CfgVal $cfgAll $id 'HistoryDays'
    $cPf = Get-CfgVal $cfgAll $id 'PlanFile'

    $tg = if ($script:HasTarget) { [int]$ShiftTarget }
          elseif ($cTg) { [int]$cTg }
          else { [int]$ShiftTarget }
    $rm = if ($script:HasRecentMin) { [int]$RecentMinutes }
          elseif ($cRm) { [int]$cRm }
          else { [int]$RecentMinutes }
    if ($rm -lt 1) { $rm = 30 }
    $sm = if ($script:HasStopMin) { [double]$StopMinutes }
          elseif ($cSm) { [double]$cSm }
          else { [double]$StopMinutes }
    if ($sm -le 0) { $sm = 2 }
    $hd = if ($script:HasHistDays) { [int]$HistoryDays }
          elseif ($cHd) { [int]$cHd }
          else { [int]$HistoryDays }
    $pf = if ($script:HasPlanFile) { $PlanFile }
          elseif ($cPf) { $cPf }
          else { '' }

    $script:LineCfg[$id] = @{
        BoxSheets     = @(Split-SheetList $bs)
        RcdbSheets    = @(Split-SheetList $rs)
        ShiftTarget   = $tg
        RecentMinutes = $rm
        StopMinutes   = $sm
        HistoryDays   = $hd
        PlanFile      = $pf
        HasTarget     = [bool]$script:HasTarget
    }
    if ($script:LineCfg[$id].BoxSheets.Count  -eq 0) { $script:LineCfg[$id].BoxSheets  = @($def.BoxSheets) }
    if ($script:LineCfg[$id].RcdbSheets.Count -eq 0) { $script:LineCfg[$id].RcdbSheets = @($def.RcdbSheets) }
}

if (-not $script:HasBox -and [string]::IsNullOrWhiteSpace($BoxPrintingFile)) {
    $BoxPrintingFile = Join-Path $here "Data_boxprintingbin3v7.xlsb"
}
$script:BoxFolder = if ($BoxPrintingFile) { Split-Path -Parent $BoxPrintingFile } else { $here }
if ([string]::IsNullOrWhiteSpace($script:BoxFolder)) { $script:BoxFolder = $here }
if ([string]::IsNullOrWhiteSpace($script:PlanFolder)) { $script:PlanFolder = $here }
# hoeveel WEKEN er extra (achter de knop '+7 dagen') uit de 31-daagse tabel gehaald worden
$script:HistExtraWeeks = 4

# welke lijn de pagina bij het openen toont
if (-not $script:LineDefs.ContainsKey($Line)) { $Line = $script:LineOrder[0] }
Set-LineContext $Line

# ============================ UITVOEREN ============================
# Dit werktuig heeft maar EEN weergave: het webdashboard. De console toont alleen serverlogs.
Start-WebServer $Port
