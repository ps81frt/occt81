&{Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Header($text) {
    Write-Host ("`n{0}`n{1}`n{0}" -f ('=' * 40), $text) -ForegroundColor Cyan
}

function Get-StatusColor($status) {
    switch ($status) {
        'OK'   { return 'Green' }
        'WARN' { return 'Yellow' }
        default { return 'Red' }
    }
}

Write-Header "  SYSTEM DIAGNOSTIC TOOL"

# ── 1. RAM TEST ──────────────────────────────────────────────────────────────
Write-Host "`n[1] RAM TEST" -ForegroundColor Yellow

$size      = 20MB
$passes    = 5
$ramErrors = 0
$rng       = [System.Random]::new()

# Référence fixe écrite puis relue — test de cohérence mémoire réelle
$ref = [byte[]]::new($size)
$rng.NextBytes($ref)

for ($i = 1; $i -le $passes; $i++) {
    Write-Host "  Pass $i/$passes..." -NoNewline
    $copy = [byte[]]::new($size)
    [Array]::Copy($ref, $copy, $size)  # copie mémoire, pas re-random

    $match = $true
    for ($j = 0; $j -lt $size; $j++) {
        if ($copy[$j] -ne $ref[$j]) { $ramErrors++; $match = $false }
    }
    Write-Host "$(if ($match) { " OK" } else { " FAIL" })" -ForegroundColor "$(if ($match) { 'Green' } else { 'Red' })"
}

# ── 2. LATENCY TEST ──────────────────────────────────────────────────────────
Write-Host "`n[2] LATENCY TEST" -ForegroundColor Yellow

$samples = 200
$lat     = [double[]]::new($samples)
$sw      = [System.Diagnostics.Stopwatch]::new()

for ($i = 0; $i -lt $samples; $i++) {
    $sw.Restart()
    $x = 0
    for ($j = 0; $j -lt 15000; $j++) { $x = $x -bxor ($j * 7) }
    $sw.Stop()
    $lat[$i] = $sw.Elapsed.TotalMilliseconds
}

$stats  = $lat | Measure-Object -Average -Minimum -Maximum
$avg    = $stats.Average
$max    = $stats.Maximum
$sorted = $lat | Sort-Object
$p95    = $sorted[[int]($samples * 0.95)]
$p99    = $sorted[[int]($samples * 0.99)]

Write-Host ("  Avg={0:N2}ms  P95={1:N2}ms  P99={2:N2}ms  Max={3:N2}ms" -f $avg, $p95, $p99, $max)

# ── 3. WHEA CHECK ────────────────────────────────────────────────────────────
Write-Host "`n[3] WHEA CHECK" -ForegroundColor Yellow

$wheaEvents = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    ProviderName = 'Microsoft-Windows-WHEA-Logger'
    Id      = 17, 18, 19, 20, 41, 4101
} -MaxEvents 50 -ErrorAction SilentlyContinue

$wheaCount   = ($wheaEvents | Measure-Object).Count
$wheaCritical= ($wheaEvents | Where-Object { $_.Id -eq 41 } | Measure-Object).Count

if ($wheaCount -gt 0) {
    Write-Host "  Derniers événements :"
    $wheaEvents | Select-Object -First 3 | ForEach-Object {
        Write-Host ("  [{0}] Id={1} {2}" -f $_.TimeCreated.ToString('dd/MM HH:mm'), $_.Id, $_.Message.Substring(0, [Math]::Min(60, $_.Message.Length))) -ForegroundColor DarkYellow
    }
}

# ── 4. CPU TEMP (via WMI/CIM) ────────────────────────────────────────────────
Write-Host "`n[4] CPU TEMP" -ForegroundColor Yellow

$tempStatus = 'N/A'
$tempVal    = '–'
try {
    $raw = (Get-CimInstance -Namespace 'root/WMI' -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction Stop).CurrentTemperature
    $celsius = [Math]::Round(($raw - 2732) / 10.0, 1)
    $tempVal    = "$celsius°C"
    $tempStatus = if ($celsius -lt 85) { 'OK' } else { 'WARN' }
    Write-Host "  CPU: $tempVal"
} catch {
    Write-Host "  Non disponible via WMI (normal sur certains systèmes)" -ForegroundColor DarkGray
}

# ── 5. RÉSUMÉ ─────────────────────────────────────────────────────────────────
Write-Header "  RÉSUMÉ"

$results = @(
    [PSCustomObject]@{
        Test   = 'RAM'
        Status = if ($ramErrors -eq 0) { 'OK' } else { 'FAIL' }
        Valeur = "$ramErrors erreur(s)"
    }
    [PSCustomObject]@{
        Test   = 'Latence moy.'
        Status = if ($avg -lt 20) { 'OK' } else { 'WARN' }
        Valeur = ("{0:N2} ms" -f $avg)
    }
    [PSCustomObject]@{
        Test   = 'Latence P99'
        Status = if ($p99 -lt 100) { 'OK' } else { 'WARN' }
        Valeur = ("{0:N2} ms" -f $p99)
    }
    [PSCustomObject]@{
        Test   = 'WHEA total'
        Status = if ($wheaCount -eq 0) { 'OK' } else { 'WARN' }
        Valeur = "$wheaCount événement(s)"
    }
    [PSCustomObject]@{
        Test   = 'WHEA critique (id41)'
        Status = if ($wheaCritical -eq 0) { 'OK' } else { 'FAIL' }
        Valeur = "$wheaCritical événement(s)"
    }
    [PSCustomObject]@{
        Test   = 'Température CPU'
        Status = $tempStatus
        Valeur = $tempVal
    }
)

$results | ForEach-Object {
    $color = Get-StatusColor $_.Status
    Write-Host ("  {0,-25} [{1,-4}]  {2}" -f $_.Test, $_.Status, $_.Valeur) -ForegroundColor $color
}

# ── 6. CONCLUSION ─────────────────────────────────────────────────────────────
Write-Host "`n" ('=' * 40)

$critical = $ramErrors -gt 0 -or $wheaCritical -gt 0
$warnings = $wheaCount -gt 0 -or $avg -ge 20 -or $p99 -ge 100

if ($critical) {
    Write-Host "  !! PROBLÈME MATÉRIEL DÉTECTÉ — vérifiez RAM et journaux WHEA" -ForegroundColor Red
} elseif ($warnings) {
    Write-Host "  !! AVERTISSEMENTS — surveillance recommandée" -ForegroundColor Yellow
} else {
    Write-Host "  Système stable — aucun problème détecté" -ForegroundColor Green
}

Write-Host ('=' * 40) "`n"
}