#Requires -Version 5.1

<#
.SYNOPSIS
    occt81 v3.0 — Outil de diagnostic universel Windows

.DESCRIPTION
    Teste RAM (patterns reels), latence CPU, erreurs WHEA, temperature (fallback chain),
    disques (SMART + vitesse I/O), GPU (nvidia-smi / OHM / LHM / WMI) et uptime.
    Fonctionne en mode CLI ou GUI (WPF avec sparkline latence).
    Historique JSON automatique, mode Watch, comparaison de runs, config JSON.

.PARAMETER Help
    Affiche l'aide courte.

.PARAMETER Man
    Affiche le manuel complet.

.PARAMETER GUI
    Lance l'interface graphique WPF.

.PARAMETER Silent
    Supprime toute sortie console.

.PARAMETER Export
    Chemin du fichier de rapport. Formats : .txt, .csv, .html

.PARAMETER Tests
    Liste des tests (virgules). Valeurs : RAM, Latence, WHEA, Temp, Disque, GPU, Uptime, Tout.

.PARAMETER Passes
    Nombre de passes RAM (defaut : 5)

.PARAMETER RamSize
    Taille buffer RAM en Mo (defaut : 512)

.PARAMETER Watch
    Mode surveillance continue — relance les tests legers toutes les N secondes.

.PARAMETER Compare
    Chemin d'un rapport JSON precedent pour comparer avec le run actuel.

.PARAMETER Config
    Chemin d'un fichier occt81.config.json (seuils personnalises).

.EXAMPLE
    .\occt81.ps1
    .\occt81.ps1 -GUI
    .\occt81.ps1 -Tests "RAM,WHEA" -Export rapport.html
    .\occt81.ps1 -Watch 60
    .\occt81.ps1 -Compare "$env:APPDATA\occt81\history\2024-01-01.json"
#>

[CmdletBinding()]
param(
    [switch]$Help,
    [switch]$Man,
    [switch]$GUI,
    [switch]$Silent,
    [string]$Export  = '',
    [string]$Tests   = 'Tout',
    [int]   $Passes  = 5,
    [int]   $RamSize = 512,
    [int]   $Watch   = 0,
    [string]$Compare = '',
    [string]$Config  = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK — demande confirmation avant elevation
# ─────────────────────────────────────────────────────────────────────────────
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    $choice = Read-Host "occt81 necessite les droits admin pour WHEA, Temp et SMART. Relancer en admin ? [O/N]"
    if ($choice -match '^[Oo]') {
        $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
        foreach ($k in $PSBoundParameters.Keys) {
            $v = $PSBoundParameters[$k]
            if ($v -is [switch]) { if ($v) { $argList += "-$k" } }
            else { $argList += "-$k"; $argList += "$v" }
        }
        Start-Process powershell.exe -ArgumentList $argList -Verb RunAs
        exit
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION — defaults + chargement config.json si present
# ─────────────────────────────────────────────────────────────────────────────
$cfg = @{
    LatMoyMax      = 20
    LatP99Max      = 100
    TempCPUMax     = 85
    TempGPUMax     = 90
    DiskPctMax     = 85
    DiskWriteMin   = 100
    RamPctWarn     = 80
    RamPctFail     = 90
    UptimeDaysWarn = 30
}

$defaultConfig = Join-Path $env:APPDATA 'occt81\occt81.config.json'
$configPath    = if ($Config -and (Test-Path $Config)) { $Config }
                 elseif (Test-Path $defaultConfig)     { $defaultConfig }
                 else                                  { $null }

if ($configPath) {
    try {
        $loaded = Get-Content $configPath -Raw | ConvertFrom-Json
        foreach ($k in $cfg.Keys) {
            if ($null -ne $loaded.$k) { $cfg[$k] = [int]$loaded.$k }
        }
    } catch { Write-Warning "Config invalide : $configPath" }
}

# ─────────────────────────────────────────────────────────────────────────────
# AIDE
# ─────────────────────────────────────────────────────────────────────────────
if ($Help) {
    Write-Host @'

occt81 v3.0 — Diagnostic systeme universel Windows
=======================================================
USAGE    .\occt81.ps1 [options]

OPTIONS
  -GUI                Interface graphique WPF (sparkline, watch, compare)
  -Tests <liste>      RAM, Latence, WHEA, Temp, Disque, GPU, Uptime, Tout
  -Export <fichier>   Rapport .txt / .csv / .html
  -Silent             Pas de sortie console
  -Passes <n>         Passes RAM (defaut : 5)
  -RamSize <Mo>       Buffer RAM (defaut : 512 Mo)
  -Watch <n>          Mode surveillance toutes les n secondes
  -Compare <json>     Compare avec un run precedent
  -Config <json>      Fichier de seuils personnalises
  -Help               Cette aide  |  -Man Manuel complet

EXEMPLES
  .\occt81.ps1
  .\occt81.ps1 -GUI
  .\occt81.ps1 -Tests "RAM,WHEA" -Export rapport.html
  .\occt81.ps1 -Watch 60
  .\occt81.ps1 -Compare "$env:APPDATA\occt81\history\2024-01-01.json"

'@ -ForegroundColor Cyan
    exit 0
}
if ($Man) { Get-Help $MyInvocation.MyCommand.Path -Full; exit 0 }

# ─────────────────────────────────────────────────────────────────────────────
# UTILITAIRES CLI
# ─────────────────────────────────────────────────────────────────────────────
function Write-Header([string]$text) {
    if ($Silent) { return }
    $bar = '=' * 58
    Write-Host "`n$bar" -ForegroundColor DarkCyan
    Write-Host "  $text" -ForegroundColor Cyan
    Write-Host "$bar" -ForegroundColor DarkCyan
}

function Write-Section([string]$text) {
    if ($Silent) { return }
    Write-Host "`n -- $text" -ForegroundColor Yellow
}

function Write-Info([string]$text, [string]$color = 'Gray') {
    if ($Silent) { return }
    Write-Host "    $text" -ForegroundColor $color
}

function Get-StatusColor([string]$status) {
    switch ($status) {
        'OK'   { return 'Green' }
        'WARN' { return 'Yellow' }
        'FAIL' { return 'Red' }
        default{ return 'DarkGray' }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# TESTS A EXECUTER
# ─────────────────────────────────────────────────────────────────────────────
$allTests   = @('RAM','Latence','WHEA','Temp','Disque','GPU','Uptime')
$watchTests = @('Latence','Temp','Disque','Uptime')
$testsToRun = if ($Tests -eq 'Tout') { $allTests } else {
    $Tests -split ',' | ForEach-Object { $_.Trim() }
}

function Set-Should-Run([string]$name) { $testsToRun -contains $name }

# ─────────────────────────────────────────────────────────────────────────────
# COLLECTE RESULTATS (liste partagee CLI + GUI)
# ─────────────────────────────────────────────────────────────────────────────
$results = [System.Collections.Generic.List[PSCustomObject]]::new()

function Add-Result([string]$test,[string]$status,[string]$valeur,[string]$detail='') {
    $r = [PSCustomObject]@{
        Test   = $test
        Status = $status
        Valeur = $valeur
        Detail = $detail
        Heure  = (Get-Date -Format 'HH:mm:ss')
    }
    $results.Add($r)
    return $r
}

# ─────────────────────────────────────────────────────────────────────────────
# TEST 1 — RAM avec patterns reels
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-RamTest {
    Write-Section "RAM — patterns memoire (${RamSize} Mo, $Passes passes)"

    $size      = [int]($RamSize * 1MB)
    $ramErrors = 0
    $patterns  = @([byte]0x00,[byte]0xFF,[byte]0xAA,[byte]0x55,[byte]0xCC,[byte]0x33)

    for ($pass = 1; $pass -le $Passes; $pass++) {
        $pat = $patterns[($pass - 1) % $patterns.Count]
        Write-Info "Pass $pass/$Passes — pattern 0x$('{0:X2}' -f $pat)" -color 'DarkGray'
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        $buf = [byte[]]::new($size)
        for ($i = 0; $i -lt $size; $i++) { $buf[$i] = $pat }

        $ok = $true
        for ($i = 0; $i -lt $size; $i++) {
            if ($buf[$i] -ne $pat) { $ok = $false; break }
        }
        if (-not $ok) { $ramErrors++ }

        # Checkerboard sur passes paires
        if ($pass % 2 -eq 0) {
            $alt = if ($pat -eq [byte]0xAA) { [byte]0x55 } else { [byte]0xAA }
            for ($i = 0; $i -lt $size; $i += 2) { $buf[$i] = $alt }
            $errChk = $false
            for ($i = 0; $i -lt $size; $i++) {
                $exp = if ($i % 2 -eq 0) { $alt } else { $pat }
                if ($buf[$i] -ne $exp) { $errChk = $true; break }
            }
            if ($errChk) { $ramErrors++ }
        }
    }

    $st = if ($ramErrors -eq 0) { 'OK' } else { 'FAIL' }
    $v  = if ($ramErrors -eq 0) { '0 erreur' } else { "$ramErrors erreur(s)" }
    $d  = "Patterns 0x00/FF/AA/55/CC/33 + checkerboard | ${RamSize} Mo x $Passes passes | NOTE: ne remplace pas MemTest86"

    Write-Info "Resultat : $v" -color (Get-StatusColor $st)
    Write-Info "NOTE : test de coherence memoire logiciel. Pour hardware complet -> MemTest86 hors OS." -color 'DarkYellow'
    Add-Result 'RAM' $st $v $d | Out-Null
}

# ─────────────────────────────────────────────────────────────────────────────
# TEST 2 — LATENCE CPU (mesure temps execution PS)
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-LatenceTest {
    Write-Section "LATENCE execution PowerShell (200 echantillons)"

    $warmup = [System.Diagnostics.Stopwatch]::StartNew()
    while ($warmup.ElapsedMilliseconds -lt 500) { $null = 1 + 1 }
    $warmup.Stop()

    $samples = 200
    $lat     = [double[]]::new($samples)
    $sw      = [System.Diagnostics.Stopwatch]::new()

    for ($i = 0; $i -lt $samples; $i++) {
        $sw.Restart()
        $x = 0
        for ($j = 0; $j -lt 50000; $j++) { $x = $x -bxor ($j * 7) }
        $sw.Stop()
        $lat[$i] = $sw.Elapsed.TotalMilliseconds
    }

    $avg    = ($lat | Measure-Object -Average).Average
    $sorted = $lat | Sort-Object
    $p95    = $sorted[[Math]::Min([int]($samples * 0.95), $samples - 1)]
    $p99    = $sorted[[Math]::Min([int]($samples * 0.99), $samples - 1)]
    $maxL   = ($lat | Measure-Object -Maximum).Maximum
    $txt    = "Avg={0:N2}ms P95={1:N2}ms P99={2:N2}ms Max={3:N2}ms" -f $avg,$p95,$p99,$maxL

    Write-Info $txt -color 'Gray'
    Write-Info "NOTE : mesure le temps d'execution PS, pas la latence CPU hardware." -color 'DarkYellow'

    $stAvg = if ($avg -lt $cfg.LatMoyMax) { 'OK' } else { 'WARN' }
    $stP99 = if ($p99 -lt $cfg.LatP99Max) { 'OK' } else { 'WARN' }

    Add-Result 'Latence (moy)' $stAvg ("{0:N2} ms" -f $avg) $txt | Out-Null
    Add-Result 'Latence (P99)' $stP99 ("{0:N2} ms" -f $p99) $txt | Out-Null

    $script:LatSamples = $lat
}

# ─────────────────────────────────────────────────────────────────────────────
# TEST 3 — WHEA
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-WheaTest {
    Write-Section "WHEA — Erreurs materielles"

    if (-not $IsAdmin) {
        Write-Info "! Droits admin requis" -color 'DarkYellow'
        Add-Result 'WHEA total'    'N/A' 'Admin requis' '' | Out-Null
        Add-Result 'WHEA critique' 'N/A' 'Admin requis' '' | Out-Null
        return
    }

    $wheaEvents = @()
    try {
        $wheaEvents = @(Get-WinEvent -FilterHashtable @{
            LogName      = 'System'
            ProviderName = 'Microsoft-Windows-WHEA-Logger'
            Id           = 17,18,19,20,41,4101
        } -MaxEvents 50 -ErrorAction SilentlyContinue)
    } catch { $wheaEvents = @() }

    $wheaCount    = $wheaEvents.Count
    $wheaCritical = @($wheaEvents | Where-Object { $_.Id -eq 41 }).Count

    if ($wheaCount -gt 0) {
        $wheaEvents | Select-Object -First 3 | ForEach-Object {
            $msg = if ($_.Message) { $_.Message.Substring(0,[Math]::Min(80,$_.Message.Length)) } else { '(no message)' }
            Write-Info "[$(  $_.TimeCreated.ToString('dd/MM HH:mm'))] Id=$($_.Id) $msg" -color 'DarkYellow'
        }
    }

    $stTotal = if ($wheaCount    -eq 0) { 'OK' } else { 'WARN' }
    $stCrit  = if ($wheaCritical -eq 0) { 'OK' } else { 'FAIL' }

    Write-Info "Total: $wheaCount | Critique (id=41): $wheaCritical" -color (Get-StatusColor $stTotal)
    Add-Result 'WHEA total'    $stTotal "$wheaCount evenement(s)"          '' | Out-Null
    Add-Result 'WHEA critique' $stCrit  "$wheaCritical evenement(s) id=41" '' | Out-Null
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER OHM — lance OHM si present a cote du script, attend le WMI
# Retourne le Process lance (ou $null si non necessaire / deja actif)
# ─────────────────────────────────────────────────────────────────────────────
function Start-OhmIfNeeded {
    # Deja actif ?
    try {
        $test = Get-CimInstance -Namespace 'root/OpenHardwareMonitor' -ClassName Sensor -ErrorAction Stop |
                Select-Object -First 1
        if ($test) { return $null }   # WMI OHM deja expose, rien a faire
    } catch { }

    # Cherche OHM.exe dans le meme dossier que le script
    #$ohmExe = Join-Path (Split-Path $PSCommandPath -Parent) 'OpenHardwareMonitor.exe'
    $ohmExe = Join-Path $PSScriptRoot 'OpenHardwareMonitor.exe'
    if (-not (Test-Path $ohmExe)) { return $null }

    if (-not $IsAdmin) {
        Write-Info "! OHM detecte mais admin requis pour le lancer (WMI)" -color 'DarkYellow'
        return $null
    }

    Write-Info "OHM detecte — lancement automatique..." -color 'DarkGray'
    try {
        $proc = Start-Process -FilePath $ohmExe -PassThru -WindowStyle Minimized -ErrorAction Stop
        # Attente WMI OHM — max 8 secondes
        $deadline = (Get-Date).AddSeconds(8)
        $ready = $false
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 500
            try {
                $chk = Get-CimInstance -Namespace 'root/OpenHardwareMonitor' -ClassName Sensor -ErrorAction Stop |
                       Select-Object -First 1
                if ($chk) { $ready = $true; break }
            } catch { }
        }
        if ($ready) {
            Write-Info "OHM WMI actif." -color 'DarkGray'
            return $proc
        } else {
            Write-Info "! OHM lance mais WMI non disponible apres 8s." -color 'DarkYellow'
            try { $proc.Kill() } catch { }
            return $null
        }
    } catch {
        Write-Info "! Impossible de lancer OHM : $($_.Exception.Message)" -color 'DarkYellow'
        return $null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# TEST 4 — TEMPERATURE avec chaine de fallback complete
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-TempTest {
    Write-Section "TEMPERATURE CPU"

    $celsius  = $null
    $src      = 'inconnu'
    $ohmProc  = Start-OhmIfNeeded   # lance OHM si present et WMI pas encore dispo

    # 1. MSAcpi_ThermalZoneTemperature (ACPI natif, admin requis)
    if ($IsAdmin) {
        try {
            $raw = (Get-CimInstance -Namespace 'root/WMI' -ClassName MSAcpi_ThermalZoneTemperature `
                    -ErrorAction Stop | Select-Object -First 1).CurrentTemperature
            if ($null -ne $raw -and $raw -gt 0) {
                $celsius = [Math]::Round(($raw - 2732) / 10.0, 1)
                $src     = 'ACPI (MSAcpi)'
            }
        } catch { }
    }

    # 2. Win32_PerfFormattedData_Counters_ThermalZoneInformation (Windows 11 natif)
    if ($null -eq $celsius) {
        try {
            $tz = Get-CimInstance -Namespace 'root/CIMV2' `
                  -ClassName 'Win32_PerfFormattedData_Counters_ThermalZoneInformation' `
                  -ErrorAction Stop | Select-Object -First 1
            if ($null -ne $tz -and $tz.Temperature -gt 0) {
                $celsius = [Math]::Round($tz.Temperature / 10.0 - 273.15, 1)
                $src     = 'ThermalZone (Win11 natif)'
            }
        } catch { }
    }

    # 3. OpenHardwareMonitor WMI (bundlé dans le repo)
    if ($null -eq $celsius) {
        try {
            $ohm = Get-CimInstance -Namespace 'root/OpenHardwareMonitor' -ClassName Sensor `
                   -ErrorAction Stop |
                   Where-Object { $_.SensorType -eq 'Temperature' -and $_.Name -match 'CPU|Package|Core|Tdie|Tctl' } |
                   Sort-Object Value -Descending | Select-Object -First 1
            if ($ohm -and $ohm.Value -gt 0) {
                $celsius = [Math]::Round($ohm.Value, 1)
                $src     = "OHM: $($ohm.Name)"
            }
        } catch { }
    }

    # 4. LibreHardwareMonitor WMI (si installé)
    if ($null -eq $celsius) {
        try {
            $lhm = Get-CimInstance -Namespace 'root/LibreHardwareMonitor' -ClassName Sensor `
                   -ErrorAction Stop |
                   Where-Object { $_.SensorType -eq 'Temperature' -and $_.Name -match 'CPU|Package|Core' } |
                   Sort-Object Value -Descending | Select-Object -First 1
            if ($lhm -and $lhm.Value -gt 0) {
                $celsius = [Math]::Round($lhm.Value, 1)
                $src     = "LHM: $($lhm.Name)"
            }
        } catch { }
    }

    if ($null -ne $celsius -and $celsius -gt 0 -and $celsius -lt 150) {
        $st = if ($celsius -lt $cfg.TempCPUMax) { 'OK' }
              elseif ($celsius -lt ($cfg.TempCPUMax + 10)) { 'WARN' }
              else { 'FAIL' }
        Write-Info "CPU : ${celsius}°C  [source: $src]" -color (Get-StatusColor $st)
        Add-Result 'Temperature CPU' $st "${celsius}°C" "Source: $src" | Out-Null
    } else {
        Write-Info "! Temperature indisponible. Lancez OpenHardwareMonitor.exe (inclus) ou LibreHardwareMonitor." -color 'DarkYellow'
        Add-Result 'Temperature CPU' 'N/A' 'Source indisponible' 'Lancer OHM ou LHM avec WMI active' | Out-Null
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# TEST 5 — DISQUES : espace + SMART via StorageReliabilityCounter + vitesse I/O
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-DisqueTest {
    Write-Section "DISQUES — Espace + SMART + vitesse I/O"

    # Espace
    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $null -ne $_.Used -and $null -ne $_.Free }
    foreach ($d in $drives) {
        $total = $d.Used + $d.Free
        if ($total -le 0) { continue }
        $pct   = [Math]::Round($d.Used / $total * 100, 1)
        $freeG = [Math]::Round($d.Free / 1GB, 1)
        $st    = if ($pct -lt $cfg.DiskPctMax) { 'OK' } elseif ($pct -lt 95) { 'WARN' } else { 'FAIL' }
        Write-Info "$($d.Name): $pct% utilise — ${freeG} Go libres" -color (Get-StatusColor $st)
        Add-Result "Disque $($d.Name):" $st "${pct}% utilise" "${freeG} Go libres" | Out-Null
    }

    # SMART — Get-PhysicalDisk + Get-StorageReliabilityCounter
    Write-Info "--- SMART ---" -color 'DarkGray'
    try {
        $physDisks = Get-PhysicalDisk -ErrorAction Stop
        foreach ($disk in $physDisks) {
            try {
                $rel  = $disk | Get-StorageReliabilityCounter -ErrorAction Stop
                $name = if ($disk.FriendlyName) { $disk.FriendlyName } else { $disk.DeviceId }
                $health = $disk.HealthStatus

                $parts = @()
                if ($rel.Wear -gt 0) { $parts += "Wear:$($rel.Wear)%" }
                if ($rel.Temperature -gt 0) { $parts += "Temp:$($rel.Temperature)°C" }
                if ($rel.ReadErrorsTotal -gt 0) { $parts += "RdErr:$($rel.ReadErrorsTotal)" }
                if ($null -ne $rel.WriteErrorsUncorrected -and $rel.WriteErrorsUncorrected -gt 0) {
                    $parts += "WrErr:$($rel.WriteErrorsUncorrected)"
                }
                if ($null -ne $rel.MediaErrors -and $rel.MediaErrors -gt 0) {
                    $parts += "MediaErr:$($rel.MediaErrors)"
                }
                $detailStr = if ($parts.Count -gt 0) { $parts -join ' | ' } else { 'Pas d anomalie detectee' }

                $st = 'OK'
                if ($health -ne 'Healthy') { $st = 'WARN' }
                if ($null -ne $rel.WriteErrorsUncorrected -and $rel.WriteErrorsUncorrected -gt 0) { $st = 'FAIL' }
                if ($null -ne $rel.MediaErrors -and $rel.MediaErrors -gt 5) { $st = 'FAIL' }
                if ($rel.Wear -gt 90) { $st = 'FAIL' } elseif ($rel.Wear -gt 75) { $st = 'WARN' }

                $valStr = "[$health]"
                if ($rel.Wear -gt 0) { $valStr += " Wear=$($rel.Wear)%" }

                Write-Info "$name : $valStr — $detailStr" -color (Get-StatusColor $st)
                Add-Result "SMART: $name" $st $valStr $detailStr | Out-Null
            } catch {
                Add-Result "SMART: $($disk.FriendlyName)" 'N/A' 'Non supporte' 'Driver incompatible' | Out-Null
            }
        }
    } catch {
        Write-Info "! Get-PhysicalDisk indisponible." -color 'DarkYellow'
        Add-Result 'SMART global' 'N/A' 'Indisponible' 'Get-PhysicalDisk non supporte' | Out-Null
    }

    # Vitesse I/O
    Write-Info "--- Vitesse I/O ---" -color 'DarkGray'
    $tmpFile = [System.IO.Path]::GetTempFileName()
    try {
        $buf = [byte[]]::new([int](50 * 1MB))
        [System.Random]::new().NextBytes($buf)
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        [System.IO.File]::WriteAllBytes($tmpFile, $buf)
        $sw.Stop()
        $mbps = [Math]::Round(50 / $sw.Elapsed.TotalSeconds, 0)
        $st   = if ($mbps -gt $cfg.DiskWriteMin) { 'OK' } elseif ($mbps -gt 30) { 'WARN' } else { 'FAIL' }
        Write-Info "Ecriture sequentielle : ${mbps} Mo/s" -color (Get-StatusColor $st)
        Add-Result 'Disque — Ecriture' $st "${mbps} Mo/s" 'Fichier temp 50 Mo' | Out-Null
    } catch {
        Add-Result 'Disque — Ecriture' 'N/A' 'Erreur I/O' $_.Exception.Message | Out-Null
    } finally {
        Remove-Item $tmpFile -ErrorAction SilentlyContinue
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# TEST 6 — GPU : nvidia-smi → OHM → LHM → WMI fallback
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-GpuTest {
    Write-Section "GPU"

    $found = $false

    # 1. nvidia-smi
    $smiPath = @(
        "$env:ProgramFiles\NVIDIA Corporation\NVSMI\nvidia-smi.exe",
        "$env:SystemRoot\System32\nvidia-smi.exe"
    ) | Where-Object { Test-Path $_ -ErrorAction SilentlyContinue } | Select-Object -First 1

    if (-not $smiPath) {
        try { $smiPath = (Get-Command 'nvidia-smi.exe' -ErrorAction Stop).Source } catch { $smiPath = $null }
    }

    if ($smiPath) {
        try {
            $smiOut = & $smiPath `
                --query-gpu=name,temperature.gpu,utilization.gpu,memory.used,memory.total,power.draw,pstate `
                --format=csv,noheader,nounits 2>$null
            foreach ($line in ($smiOut -split "`n" | Where-Object { $_.Trim() -ne '' })) {
                $p = $line -split ',' | ForEach-Object { $_.Trim() }
                if ($p.Count -lt 5) { continue }
                $gpuName = $p[0]
                $tempG   = if ($p[1] -match '^\d') { [int]$p[1] } else { $null }
                $utilG   = if ($p[2] -match '^\d') { [int]$p[2] } else { $null }
                $mU      = if ($p[3] -match '^\d') { [int]$p[3] } else { $null }
                $mT      = if ($p[4] -match '^\d') { [int]$p[4] } else { $null }
                $pw      = if ($p.Count -gt 5 -and $p[5] -match '^\d') { [double]$p[5] } else { $null }

                $st = 'OK'
                if ($tempG -and $tempG -gt $cfg.TempGPUMax)          { $st = 'FAIL' }
                elseif ($tempG -and $tempG -gt ($cfg.TempGPUMax-10)) { $st = 'WARN' }

                $val = if ($tempG) { "${tempG}°C" } else { 'N/A' }
                $det = ''
                if ($null -ne $utilG) { $det += "Load:${utilG}% " }
                if ($mU -and $mT)     { $det += "VRAM:${mU}/${mT}MiB " }
                if ($null -ne $pw)    { $det += "Pwr:$([Math]::Round($pw,1))W " }
                $det += '| nvidia-smi'

                Write-Info "$gpuName : $val | $det" -color (Get-StatusColor $st)
                Add-Result "GPU: $gpuName" $st $val $det | Out-Null
                $found = $true
            }
        } catch { }
    }

    # 2. OHM WMI
    if (-not $found) {
        try {
            $ohmSensors = @(Get-CimInstance -Namespace 'root/OpenHardwareMonitor' -ClassName Sensor -ErrorAction Stop |
                Where-Object { $_.SensorType -eq 'Temperature' -and $_.Name -match 'GPU' })
            foreach ($g in $ohmSensors) {
                if ($g.Value -gt 0) {
                    $st = if ($g.Value -lt $cfg.TempGPUMax) { 'OK' } else { 'WARN' }
                    Write-Info "$($g.Name) : $([Math]::Round($g.Value,1))°C [OHM]" -color (Get-StatusColor $st)
                    Add-Result "GPU Temp: $($g.Name)" $st "$([Math]::Round($g.Value,1))°C" 'Source: OHM' | Out-Null
                    $found = $true
                }
            }
        } catch { }
    }

    # 3. LHM WMI
    if (-not $found) {
        try {
            $lhmSensors = @(Get-CimInstance -Namespace 'root/LibreHardwareMonitor' -ClassName Sensor -ErrorAction Stop |
                Where-Object { $_.SensorType -eq 'Temperature' -and $_.Name -match 'GPU' })
            foreach ($g in $lhmSensors) {
                if ($g.Value -gt 0) {
                    $st = if ($g.Value -lt $cfg.TempGPUMax) { 'OK' } else { 'WARN' }
                    Add-Result "GPU Temp: $($g.Name)" $st "$([Math]::Round($g.Value,1))°C" 'Source: LHM' | Out-Null
                    $found = $true
                }
            }
        } catch { }
    }

    # 4. Fallback Win32_VideoController (info driver, pas de temperature)
    if (-not $found) {
        try {
            $gpus = Get-CimInstance -ClassName Win32_VideoController -ErrorAction Stop
            foreach ($g in $gpus) {
                $vram = if ($g.AdapterRAM -gt 0) { [Math]::Round($g.AdapterRAM / 1MB, 0) } else { '?' }
                $st   = if ($g.Status -eq 'OK') { 'OK' } else { 'WARN' }
                Write-Info "$($g.Name) — VRAM ${vram} Mo [pas de temperature — installez nvidia-smi ou OHM]" -color (Get-StatusColor $st)
                Add-Result "GPU: $($g.Name)" $st "VRAM ${vram} Mo" "Driver: $($g.DriverVersion) | Info driver seul" | Out-Null
            }
        } catch {
            Add-Result 'GPU' 'N/A' 'Indisponible' $_.Exception.Message | Out-Null
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# TEST 7 — UPTIME & SYSTEME
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-UptimeTest {
    Write-Section "UPTIME & SYSTEME"

    $os      = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cpu     = Get-CimInstance Win32_Processor       -ErrorAction SilentlyContinue | Select-Object -First 1
    $ramInfo = Get-CimInstance Win32_PhysicalMemory  -ErrorAction SilentlyContinue | Measure-Object -Property Capacity -Sum

    if (-not $os) {
        Add-Result 'Uptime' 'N/A' 'Indisponible' '' | Out-Null
        return
    }

    $uptime  = (Get-Date) - $os.LastBootUpTime
    $days    = [Math]::Floor($uptime.TotalDays)
    $ramGB   = [Math]::Round($ramInfo.Sum / 1GB, 1)
    $ramUsed = [Math]::Round(($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / 1MB, 1)
    $ramTot  = [Math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
    $ramPct  = [Math]::Round($ramUsed / $ramTot * 100, 1)

    Write-Info "OS     : $($os.Caption) build $($os.BuildNumber)" -color 'Gray'
    Write-Info "CPU    : $($cpu.Name)" -color 'Gray'
    Write-Info "RAM    : ${ramGB} Go installes — ${ramPct}% utilises" -color 'Gray'
    Write-Info "Uptime : ${days}j $($uptime.Hours)h $($uptime.Minutes)m" -color 'Gray'

    $stUptime = if ($days -lt $cfg.UptimeDaysWarn) { 'OK' } else { 'WARN' }
    $stRAM    = if ($ramPct -lt $cfg.RamPctWarn) { 'OK' }
                elseif ($ramPct -lt $cfg.RamPctFail) { 'WARN' }
                else { 'FAIL' }

    Add-Result 'Uptime'       $stUptime "${days}j $($uptime.Hours)h $($uptime.Minutes)m" "OS: $($os.Caption)" | Out-Null
    Add-Result 'RAM utilisee' $stRAM    "${ramPct}%"                                      "Physique: ${ramGB} Go" | Out-Null
    Add-Result 'CPU info'     'OK'      $cpu.Name                                         "Cores: $($cpu.NumberOfCores) / Logiques: $($cpu.NumberOfLogicalProcessors)" | Out-Null
}

# ─────────────────────────────────────────────────────────────────────────────
# MOTEUR PRINCIPAL (unique, partagé CLI + GUI via runspace)
# ─────────────────────────────────────────────────────────────────────────────
function Invoke-AllTests {
    $results.Clear()
    if (Set-Should-Run 'RAM')     { Invoke-RamTest    }
    if (Set-Should-Run 'Latence') { Invoke-LatenceTest }
    if (Set-Should-Run 'WHEA')    { Invoke-WheaTest   }
    if (Set-Should-Run 'Temp')    { Invoke-TempTest   }
    if (Set-Should-Run 'Disque')  { Invoke-DisqueTest }
    if (Set-Should-Run 'GPU')     { Invoke-GpuTest    }
    if (Set-Should-Run 'Uptime')  { Invoke-UptimeTest }
}

# ─────────────────────────────────────────────────────────────────────────────
# HISTORIQUE JSON
# ─────────────────────────────────────────────────────────────────────────────
$historyDir = Join-Path $env:APPDATA 'occt81\history'

function Save-History {
    try {
        if (-not (Test-Path $historyDir)) {
            New-Item -ItemType Directory -Path $historyDir -Force | Out-Null
        }
        $stamp   = Get-Date -Format 'yyyy-MM-ddTHH-mm-ss'
        $outFile = Join-Path $historyDir "$stamp.json"
        $payload = @{
            Date    = (Get-Date -Format 'o')
            Machine = $env:COMPUTERNAME
            Results = @($results | ForEach-Object {
                @{ Test=$_.Test; Status=$_.Status; Valeur=$_.Valeur; Detail=$_.Detail; Heure=$_.Heure }
            })
        }
        $payload | ConvertTo-Json -Depth 5 | Set-Content $outFile -Encoding UTF8
        if (-not $Silent) { Write-Info "Historique : $outFile" -color 'DarkGray' }
        return $outFile
    } catch {
        if (-not $Silent) { Write-Info "! Historique non sauvegarde : $($_.Exception.Message)" -color 'DarkYellow' }
        return $null
    }
}

function Compare-History([string]$jsonPath) {
    if (-not (Test-Path $jsonPath)) {
        Write-Info "! Fichier introuvable : $jsonPath" -color 'Red'
        return
    }
    try {
        $prev    = Get-Content $jsonPath -Raw | ConvertFrom-Json
        $prevDate = if ($prev.Date) { $prev.Date } else { 'date inconnue' }
        Write-Header "COMPARAISON avec run du $prevDate — $($prev.Machine)"

        $prevMap = @{}
        foreach ($r in $prev.Results) { $prevMap[$r.Test] = $r }

        $changes = 0
        foreach ($cur in $results) {
            $p = $prevMap[$cur.Test]
            if (-not $p) {
                Write-Info "[NOUVEAU]  $($cur.Test) : $($cur.Status) — $($cur.Valeur)" -color 'Cyan'
                $changes++
                continue
            }
            if ($p.Status -ne $cur.Status) {
                $changes++
                $arrow = switch ($cur.Status) {
                    'OK'   { '<-- AMELIORE' }
                    'FAIL' { '<-- DEGRADE !!' }
                    default { '<-- change' }
                }
                $col = switch ($cur.Status) {
                    'OK'   { 'Green' }
                    'FAIL' { 'Red' }
                    default { 'Yellow' }
                }
                Write-Info "$($cur.Test) : $($p.Status) -> $($cur.Status)  $arrow" -color $col
            }
        }
        if ($changes -eq 0) { Write-Info 'Aucun changement de statut detecte.' -color 'Green' }
    } catch {
        Write-Info "! Erreur comparaison : $($_.Exception.Message)" -color 'Red'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# EXPORT (TXT / CSV / HTML)
# ─────────────────────────────────────────────────────────────────────────────
function Export-Report {
    if (-not $Export) { return }
    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Export)
    $ext = [System.IO.Path]::GetExtension($resolvedPath).ToLower()

    switch ($ext) {
        '.csv' {
            $results | Select-Object Heure,Test,Status,Valeur,Detail |
                Export-Csv -Path $resolvedPath -NoTypeInformation -Encoding UTF8
        }
        '.html' {
            $osH   = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
            $cpuH  = Get-CimInstance Win32_Processor       -ErrorAction SilentlyContinue | Select-Object -First 1
            $upH   = if ($osH) { (Get-Date) - $osH.LastBootUpTime } else { $null }
            $dispOS  = if ($osH)  { $osH.Caption } else { 'Inconnu' }
            $dispCPU = if ($cpuH) { $cpuH.Name   } else { 'Inconnu' }
            $dispUp  = if ($upH)  { "$([Math]::Floor($upH.TotalDays))j $($upH.Hours)h $($upH.Minutes)m" } else { 'N/A' }

            $rows = @($results | ForEach-Object {
                $bg  = switch ($_.Status) { 'OK'{'#0a1628'} 'WARN'{'#1e1b16'} 'FAIL'{'#1c1917'} default{'#0a1628'} }
                $col = switch ($_.Status) { 'OK'{'#4ade80'} 'WARN'{'#fbbf24'} 'FAIL'{'#f87171'} default{'#94a3b8'} }
                "<tr style='background:$bg'><td style='color:#64748b'>$($_.Heure)</td><td style='color:#e2e8f0'>$($_.Test)</td><td style='color:$col;font-weight:bold'>$($_.Status)</td><td style='color:#e2e8f0'>$($_.Valeur)</td><td style='color:#475569;font-size:12px'>$($_.Detail)</td></tr>"
            })

            $crit  = @($results | Where-Object { $_.Status -eq 'FAIL' }).Count
            $warns = @($results | Where-Object { $_.Status -eq 'WARN' }).Count
            $concl = if ($crit -gt 0) { '&#9940; Probleme materiel detecte' }
                     elseif ($warns -gt 0) { '&#9888; Avertissements' }
                     else { '&#9989; Systeme stable' }
            $cCol  = if ($crit -gt 0) { '#f87171' } elseif ($warns -gt 0) { '#fbbf24' } else { '#4ade80' }

            $html = @"
<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><title>occt81 v3.0</title>
<style>
body{font-family:Consolas,monospace;background:#020617;color:#e2e8f0;margin:0;padding:32px}
h1{color:#38bdf8;font-size:22px;letter-spacing:1px;margin-bottom:4px}
.meta{color:#94a3b8;font-size:13px;margin-bottom:24px;line-height:1.8;border-left:3px solid #1e293b;padding-left:15px}
table{width:100%;border-collapse:collapse;font-size:13px;border:1px solid #1e293b}
th{background:#050d1a;color:#475569;padding:12px 14px;text-align:left;font-weight:bold;border-bottom:1px solid #1e293b}
td{padding:10px 14px;border-bottom:1px solid #0a1628}
.concl{margin-top:24px;padding:16px 20px;border-radius:4px;background:#0a1628;border:1px solid #1e293b;font-size:15px;color:$cCol;font-weight:bold}
.foot{margin-top:12px;color:#1e293b;font-size:11px}
</style></head><body>
<h1>&#9881; occt81 v3.0</h1>
<div class="meta">
<strong>Machine :</strong> $env:COMPUTERNAME<br>
<strong>Systeme :</strong> $dispOS<br>
<strong>CPU :</strong> $dispCPU<br>
<strong>Uptime :</strong> $dispUp<br>
<strong>Date :</strong> $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
</div>
<table>
<tr><th>HEURE</th><th>TEST</th><th>STATUT</th><th>VALEUR</th><th>DETAIL</th></tr>
$($rows -join "`n")
</table>
<div class="concl">$concl</div>
<div class="foot">Genere par occt81 v3.0 — github.com/ps81frt/occt81</div>
</body></html>
"@
            [System.IO.File]::WriteAllText($resolvedPath, $html, [System.Text.Encoding]::UTF8)
        }
        default {
            $osH  = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
            $cpuH = Get-CimInstance Win32_Processor       -ErrorAction SilentlyContinue | Select-Object -First 1
            $upH  = if ($osH) { (Get-Date) - $osH.LastBootUpTime } else { $null }
            $lines = @(
                "occt81 v3.0 — $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')",
                "Machine : $env:COMPUTERNAME",
                "Systeme : $(if($osH){$osH.Caption}else{'Inconnu'})",
                "CPU     : $(if($cpuH){$cpuH.Name}else{'Inconnu'})",
                "Uptime  : $(if($upH){"$([Math]::Floor($upH.TotalDays))j $($upH.Hours)h $($upH.Minutes)m"}else{'N/A'})",
                ('-' * 90)
            )
            foreach ($r in $results) {
                $lines += "{0,-32} [{1,-4}] {2,-22} {3}" -f $r.Test, $r.Status, $r.Valeur, $r.Detail
            }
            $lines | Set-Content -Path $resolvedPath -Encoding UTF8
        }
    }
    if (-not $Silent) { Write-Info "Export : $resolvedPath" -color 'Green' }
}

# ─────────────────────────────────────────────────────────────────────────────
# RESUME CLI
# ─────────────────────────────────────────────────────────────────────────────
function Write-Summary {
    Write-Header 'RESUME'
    foreach ($r in $results) {
        $color  = Get-StatusColor $r.Status
        $name   = $r.Test.PadRight(30)
        $status = "[$($r.Status.PadRight(4))]"
        $val    = $r.Valeur.PadRight(20)
        Write-Host "  $name $status $val $($r.Detail)" -ForegroundColor $color
    }

    $bar      = '=' * 58
    $critical = @($results | Where-Object { $_.Status -eq 'FAIL' }).Count
    $warns    = @($results | Where-Object { $_.Status -eq 'WARN' }).Count

    Write-Host "`n  $bar" -ForegroundColor DarkCyan
    if ($critical -gt 0) {
        Write-Host "  !! PROBLEME MATERIEL — verifiez RAM et journaux WHEA" -ForegroundColor Red
    } elseif ($warns -gt 0) {
        Write-Host "  !! AVERTISSEMENTS — surveillance recommandee" -ForegroundColor Yellow
    } else {
        Write-Host "  Systeme stable — aucun probleme detecte" -ForegroundColor Green
    }
    Write-Host "  $bar`n" -ForegroundColor DarkCyan
}

# ─────────────────────────────────────────────────────────────────────────────
# MODE WATCH CLI
# ─────────────────────────────────────────────────────────────────────────────
function Start-WatchMode([int]$intervalSec) {
    Write-Header "WATCH MODE — intervalle ${intervalSec}s | Ctrl+C pour arreter"
    Write-Info "Tests : $($watchTests -join ', ')" -color 'DarkGray'

    $savedTests = $script:testsToRun
    $script:testsToRun = $watchTests

    $run = 0
    try {
        while ($true) {
            $run++
            Write-Host "`n  [$(Get-Date -Format 'HH:mm:ss')] Run #$run" -ForegroundColor Cyan
            $results.Clear()
            Invoke-AllTests
            foreach ($r in $results) {
                Write-Host ("    {0,-28} [{1,-4}] {2}" -f $r.Test, $r.Status, $r.Valeur) `
                    -ForegroundColor (Get-StatusColor $r.Status)
            }
            $crit = @($results | Where-Object { $_.Status -eq 'FAIL' }).Count
            if ($crit -gt 0) { Write-Host "    !! ALERTE : $crit test(s) FAIL" -ForegroundColor Red }
            Save-History | Out-Null
            Start-Sleep -Seconds $intervalSec
        }
    } finally {
        $script:testsToRun = $savedTests
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# GUI WPF — sparkline latence, watch mode, compare, export, historique auto
# ─────────────────────────────────────────────────────────────────────────────
function Show-Gui {
    Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Windows.Forms

    [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="occt81 v3.0" Height="840" Width="1000"
        Background="#020617" FontFamily="Consolas"
        WindowStartupLocation="CenterScreen" MinWidth="700" MinHeight="640">
  <Window.Resources>
    <Style TargetType="Button" x:Key="Btn">
      <Setter Property="Background"      Value="#0a1e3a"/>
      <Setter Property="Foreground"      Value="#38bdf8"/>
      <Setter Property="BorderBrush"     Value="#1e4976"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding"         Value="12,6"/>
      <Setter Property="FontFamily"      Value="Consolas"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="3" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#142d56"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.25"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#7dd3fc"/>
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize"   Value="12"/>
      <Setter Property="Margin"     Value="0,2"/>
    </Style>
  </Window.Resources>

  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="3"/>
      <RowDefinition Height="110"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Titre -->
    <StackPanel Grid.Row="0" Margin="0,0,0,12">
      <TextBlock Text="&#9881;  occt81  v3.0" FontSize="20" FontWeight="Bold" Foreground="#38bdf8"/>
      <TextBlock x:Name="lblMachine" FontSize="11" Foreground="#1e3a5f" Margin="2,3,0,0"/>
    </StackPanel>

    <!-- Checkboxes -->
    <Border Grid.Row="1" Background="#070f1e" CornerRadius="3"
            Padding="12,8" Margin="0,0,0,10"
            BorderBrush="#0f2140" BorderThickness="1">
      <StackPanel>
        <TextBlock Text="TESTS" FontSize="9" Foreground="#1e3a5f"
                   Margin="0,0,0,7" FontWeight="Bold"/>
        <WrapPanel>
          <CheckBox x:Name="chkRAM"     Content="RAM"     IsChecked="True" Margin="0,0,16,0"/>
          <CheckBox x:Name="chkLatence" Content="Latence" IsChecked="True" Margin="0,0,16,0"/>
          <CheckBox x:Name="chkWHEA"    Content="WHEA"    IsChecked="True" Margin="0,0,16,0"/>
          <CheckBox x:Name="chkTemp"    Content="Temp"    IsChecked="True" Margin="0,0,16,0"/>
          <CheckBox x:Name="chkDisque"  Content="Disque"  IsChecked="True" Margin="0,0,16,0"/>
          <CheckBox x:Name="chkGPU"     Content="GPU"     IsChecked="True" Margin="0,0,16,0"/>
          <CheckBox x:Name="chkUptime"  Content="Uptime"  IsChecked="True"/>
        </WrapPanel>
      </StackPanel>
    </Border>

    <!-- Boutons -->
    <WrapPanel Grid.Row="2" Margin="0,0,0,10">
      <Button x:Name="btnRun"     Content="&#9654; LANCER"    Style="{StaticResource Btn}" Margin="0,0,6,0"/>
      <Button x:Name="btnWatch"   Content="&#128262; WATCH"   Style="{StaticResource Btn}" Margin="0,0,6,0"/>
      <Button x:Name="btnStop"    Content="&#9646;&#9646; STOP" Style="{StaticResource Btn}" Margin="0,0,6,0" IsEnabled="False"/>
      <Button x:Name="btnExport"  Content="&#128190; EXPORT"  Style="{StaticResource Btn}" Margin="0,0,6,0" IsEnabled="False"/>
      <Button x:Name="btnCompare" Content="&#128203; COMPARE" Style="{StaticResource Btn}" Margin="0,0,6,0" IsEnabled="False"/>
      <Button x:Name="btnConfig"  Content="&#9881; CONFIG"    Style="{StaticResource Btn}" Margin="0,0,6,0"/>
      <Button x:Name="btnClear"   Content="&#10005; EFFACER"  Style="{StaticResource Btn}"/>
      <TextBlock x:Name="lblStatus" VerticalAlignment="Center"
                 Foreground="#1e3a5f" FontSize="11" Margin="12,0,0,0"/>
    </WrapPanel>

    <!-- DataGrid -->
    <Border Grid.Row="3" Background="#070f1e" CornerRadius="3"
            BorderBrush="#0f2140" BorderThickness="1">
      <DataGrid x:Name="dgResults" AutoGenerateColumns="False"
                Background="#070f1e" Foreground="#e2e8f0"
                GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#050d1a"
                RowBackground="#070f1e" AlternatingRowBackground="#0a1628"
                BorderThickness="0" CanUserAddRows="False" IsReadOnly="True"
                ColumnHeaderHeight="28" FontSize="12">
        <DataGrid.ColumnHeaderStyle>
          <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background"      Value="#030810"/>
            <Setter Property="Foreground"      Value="#1e3a5f"/>
            <Setter Property="FontFamily"      Value="Consolas"/>
            <Setter Property="FontSize"        Value="10"/>
            <Setter Property="Padding"         Value="10,0"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Setter Property="BorderBrush"     Value="#0f2140"/>
          </Style>
        </DataGrid.ColumnHeaderStyle>
        <DataGrid.Columns>
          <DataGridTextColumn Header="HEURE"  Binding="{Binding Heure}"  Width="62"/>
          <DataGridTextColumn Header="TEST"   Binding="{Binding Test}"   Width="200"/>
          <DataGridTemplateColumn Header="STATUT" Width="72">
            <DataGridTemplateColumn.CellTemplate>
              <DataTemplate>
                <TextBlock Text="{Binding Status}" FontWeight="Bold" Padding="10,4">
                  <TextBlock.Style>
                    <Style TargetType="TextBlock">
                      <Style.Triggers>
                        <DataTrigger Binding="{Binding Status}" Value="OK">
                          <Setter Property="Foreground" Value="#4ade80"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding Status}" Value="WARN">
                          <Setter Property="Foreground" Value="#fbbf24"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding Status}" Value="FAIL">
                          <Setter Property="Foreground" Value="#f87171"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding Status}" Value="N/A">
                          <Setter Property="Foreground" Value="#1e3a5f"/>
                        </DataTrigger>
                      </Style.Triggers>
                    </Style>
                  </TextBlock.Style>
                </TextBlock>
              </DataTemplate>
            </DataGridTemplateColumn.CellTemplate>
          </DataGridTemplateColumn>
          <DataGridTextColumn Header="VALEUR" Binding="{Binding Valeur}" Width="130"/>
          <DataGridTextColumn Header="DETAIL" Binding="{Binding Detail}" Width="*"/>
        </DataGrid.Columns>
      </DataGrid>
    </Border>

    <!-- Progress -->
    <ProgressBar x:Name="pbProgress" Grid.Row="4"
                 Background="#070f1e" Foreground="#38bdf8" BorderThickness="0"
                 Minimum="0" Maximum="100" Value="0"/>

    <!-- Sparkline latence -->
    <Border Grid.Row="5" Background="#070f1e" CornerRadius="3"
            BorderBrush="#0f2140" BorderThickness="1" Padding="10,6" Margin="0,6,0,6">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="LATENCE — SPARKLINE (ms)"
                   FontSize="9" Foreground="#1e3a5f" FontWeight="Bold"
                   Margin="0,0,0,3"/>
        <Canvas x:Name="sparkCanvas" Grid.Row="1" ClipToBounds="True" Background="Transparent"/>
      </Grid>
    </Border>

    <!-- Conclusion -->
    <Border x:Name="borderConcl" Grid.Row="6" CornerRadius="3" Padding="14,10"
            Background="#070f1e" BorderBrush="#0f2140" BorderThickness="1" Margin="0,0,0,0">
      <TextBlock x:Name="lblConcl" FontSize="13" FontWeight="Bold"
                 Foreground="#1e3a5f" Text="En attente du diagnostic..."/>
    </Border>

    <!-- Footer hint -->
    <TextBlock Grid.Row="7" FontSize="9" Foreground="#0a1628" Margin="0,4,0,0"
               Text="Historique : %APPDATA%\occt81\history\  |  Config : %APPDATA%\occt81\occt81.config.json"/>
  </Grid>
</Window>
'@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Controles
    $dgResults   = $window.FindName('dgResults')
    $btnRun      = $window.FindName('btnRun')
    $btnWatch    = $window.FindName('btnWatch')
    $btnStop     = $window.FindName('btnStop')
    $btnExport   = $window.FindName('btnExport')
    $btnCompare  = $window.FindName('btnCompare')
    $btnConfig   = $window.FindName('btnConfig')
    $btnClear    = $window.FindName('btnClear')
    $lblStatus   = $window.FindName('lblStatus')
    $lblConcl    = $window.FindName('lblConcl')
    $lblMachine  = $window.FindName('lblMachine')
    $pbProgress  = $window.FindName('pbProgress')
    $borderConcl = $window.FindName('borderConcl')
    $sparkCanvas = $window.FindName('sparkCanvas')

    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $lblMachine.Text = "$env:COMPUTERNAME  |  $($osInfo.Caption)  |  $(if($IsAdmin){'Admin'}else{'Standard'})"

    $observableResults = [System.Collections.ObjectModel.ObservableCollection[PSCustomObject]]::new()
    $dgResults.ItemsSource = $observableResults

    $script:watchCts  = $null
    $script:currentPs = $null

    # ── Lancement runspace ─────────────────────────────────────────────────
    function Start-GuiRunspace([string[]]$selTests, [bool]$isWatch, [int]$watchSec) {

        $btnRun.IsEnabled     = $false
        $btnWatch.IsEnabled   = $false
        $btnExport.IsEnabled  = $false
        $btnCompare.IsEnabled = $false
        $btnStop.IsEnabled    = $true
        $pbProgress.Value     = 0
        $lblStatus.Text       = if ($isWatch) { "Watch en cours..." } else { "Diagnostic en cours..." }

        if (-not $isWatch) { $observableResults.Clear(); $results.Clear() }

        $script:watchCts = [System.Threading.CancellationTokenSource]::new()

        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = 'STA'; $rs.ThreadOptions = 'ReuseThread'; $rs.Open()

        foreach ($vn in @('RamSize','Passes','IsAdmin','cfg','historyDir')) {
            $rs.SessionStateProxy.SetVariable($vn, (Get-Variable $vn -Scope Script -ErrorAction SilentlyContinue).Value)
        }
        $rs.SessionStateProxy.SetVariable('selTests',       $selTests)
        $rs.SessionStateProxy.SetVariable('isWatch',        $isWatch)
        $rs.SessionStateProxy.SetVariable('watchSec',       $watchSec)
        $rs.SessionStateProxy.SetVariable('dispatcher',     $window.Dispatcher)
        $rs.SessionStateProxy.SetVariable('observableR',    $observableResults)
        $rs.SessionStateProxy.SetVariable('resultsRef',     $results)
        $rs.SessionStateProxy.SetVariable('pbProgress',     $pbProgress)
        $rs.SessionStateProxy.SetVariable('btnRun',         $btnRun)
        $rs.SessionStateProxy.SetVariable('btnWatch',       $btnWatch)
        $rs.SessionStateProxy.SetVariable('btnStop',        $btnStop)
        $rs.SessionStateProxy.SetVariable('btnExport',      $btnExport)
        $rs.SessionStateProxy.SetVariable('btnCompare',     $btnCompare)
        $rs.SessionStateProxy.SetVariable('lblStatus',      $lblStatus)
        $rs.SessionStateProxy.SetVariable('lblConcl',       $lblConcl)
        $rs.SessionStateProxy.SetVariable('borderConcl',    $borderConcl)
        $rs.SessionStateProxy.SetVariable('sparkCanvas',    $sparkCanvas)
        $rs.SessionStateProxy.SetVariable('cts',            $script:watchCts)

        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs
        $script:currentPs = $ps

        $null = $ps.AddScript({

            function Add-GR([string]$t,[string]$s,[string]$v,[string]$d='') {
                $obj=[PSCustomObject]@{Test=$t;Status=$s;Valeur=$v;Detail=$d;Heure=(Get-Date -Format 'HH:mm:ss')}
                $resultsRef.Add($obj)
                $dispatcher.Invoke([Action]{ $observableR.Add($obj) })
            }
            function SetProg([int]$v) { $dispatcher.Invoke([Action]{ $pbProgress.Value=$v }) }

            function Set-Run-TestsOnce([string[]]$tests) {
                $total=$tests.Count; $done=0; $latSmp=$null

                if ($tests -contains 'RAM') {
                    $sz=[int]($RamSize*1MB); $err=0
                    $pats=@([byte]0x00,[byte]0xFF,[byte]0xAA,[byte]0x55,[byte]0xCC,[byte]0x33)
                    for ($p=1;$p -le $Passes;$p++) {
                        $pat=$pats[($p-1)%$pats.Count]
                        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
                        $buf=[byte[]]::new($sz)
                        for ($i=0;$i -lt $sz;$i++) { $buf[$i]=$pat }
                        $ok=$true
                        for ($i=0;$i -lt $sz;$i++) { if ($buf[$i] -ne $pat) { $ok=$false;break } }
                        if (-not $ok) { $err++ }
                        if ($p%2 -eq 0) {
                            $alt=if($pat -eq [byte]0xAA){[byte]0x55}else{[byte]0xAA}
                            for ($i=0;$i -lt $sz;$i+=2) { $buf[$i]=$alt }
                            $ec=$false
                            for ($i=0;$i -lt $sz;$i++) {
                                $exp=if($i%2 -eq 0){$alt}else{$pat}
                                if ($buf[$i] -ne $exp) { $ec=$true;break }
                            }
                            if ($ec) { $err++ }
                        }
                    }
                    Add-GR 'RAM' (if($err -eq 0){'OK'}else{'FAIL'}) (if($err -eq 0){'0 erreur'}else{"$err erreur(s)"}) "Patterns+checkerboard | ${RamSize}Mo x ${Passes} passes | NOTE: ne remplace pas MemTest86"
                    $done++; SetProg([int]($done/$total*100))
                }

                if ($tests -contains 'Latence') {
                    $w=[System.Diagnostics.Stopwatch]::StartNew(); while($w.ElapsedMilliseconds -lt 500){$null=1+1}; $w.Stop()
                    $n=200; $lat=[double[]]::new($n); $sw=[System.Diagnostics.Stopwatch]::new()
                    for ($i=0;$i -lt $n;$i++) {
                        $sw.Restart(); $x=0
                        for ($j=0;$j -lt 50000;$j++) { $x=$x -bxor ($j*7) }
                        $sw.Stop(); $lat[$i]=$sw.Elapsed.TotalMilliseconds
                    }
                    $avg=($lat|Measure-Object -Average).Average
                    $s2=$lat|Sort-Object
                    $p99=$s2[[Math]::Min([int]($n*0.99),$n-1)]
                    $p95=$s2[[Math]::Min([int]($n*0.95),$n-1)]
                    $mxL=($lat|Measure-Object -Maximum).Maximum
                    $txt="Avg={0:N2}ms P95={1:N2}ms P99={2:N2}ms Max={3:N2}ms" -f $avg,$p95,$p99,$mxL
                    Add-GR 'Latence (moy)' (if($avg -lt $cfg.LatMoyMax){'OK'}else{'WARN'}) ("{0:N2} ms" -f $avg) $txt
                    Add-GR 'Latence (P99)' (if($p99 -lt $cfg.LatP99Max){'OK'}else{'WARN'}) ("{0:N2} ms" -f $p99) $txt
                    $latSmp=$lat
                    $done++; SetProg([int]($done/$total*100))
                }

                if ($tests -contains 'WHEA') {
                    if (-not $IsAdmin) {
                        Add-GR 'WHEA' 'N/A' 'Admin requis' ''
                    } else {
                        $ev=@(); try { $ev=@(Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='Microsoft-Windows-WHEA-Logger';Id=17,18,19,20,41,4101} -MaxEvents 50 -ErrorAction SilentlyContinue) } catch {}
                        $wc=$ev.Count; $cc=@($ev|Where-Object{$_.Id -eq 41}).Count
                        Add-GR 'WHEA total'    (if($wc -eq 0){'OK'}else{'WARN'}) "$wc evenement(s)" ''
                        Add-GR 'WHEA critique' (if($cc -eq 0){'OK'}else{'FAIL'}) "$cc id=41" ''
                    }
                    $done++; SetProg([int]($done/$total*100))
                }

                if ($tests -contains 'Temp') {
                    $cel=$null; $src='inconnu'
                    if ($IsAdmin) {
                        try { $rw=(Get-CimInstance -Namespace 'root/WMI' -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction Stop|Select-Object -First 1).CurrentTemperature; if($rw -gt 0){$cel=[Math]::Round(($rw-2732)/10.0,1);$src='ACPI'} } catch {}
                    }
                    if ($null -eq $cel) {
                        try { $tz=Get-CimInstance -Namespace 'root/CIMV2' -ClassName 'Win32_PerfFormattedData_Counters_ThermalZoneInformation' -ErrorAction Stop|Select-Object -First 1; if($tz -and $tz.Temperature -gt 0){$cel=[Math]::Round($tz.Temperature/10.0-273.15,1);$src='ThermalZone'} } catch {}
                    }
                    if ($null -eq $cel) {
                        try { $o=Get-CimInstance -Namespace 'root/OpenHardwareMonitor' -ClassName Sensor -ErrorAction Stop|Where-Object{$_.SensorType -eq 'Temperature' -and $_.Name -match 'CPU|Package|Core'}|Sort-Object Value -Descending|Select-Object -First 1; if($o -and $o.Value -gt 0){$cel=[Math]::Round($o.Value,1);$src="OHM:$($o.Name)"} } catch {}
                    }
                    if ($null -eq $cel) {
                        try { $l=Get-CimInstance -Namespace 'root/LibreHardwareMonitor' -ClassName Sensor -ErrorAction Stop|Where-Object{$_.SensorType -eq 'Temperature' -and $_.Name -match 'CPU|Package|Core'}|Sort-Object Value -Descending|Select-Object -First 1; if($l -and $l.Value -gt 0){$cel=[Math]::Round($l.Value,1);$src="LHM:$($l.Name)"} } catch {}
                    }
                    if ($null -ne $cel -and $cel -gt 0 -and $cel -lt 150) {
                        $st=if($cel -lt $cfg.TempCPUMax){'OK'}elseif($cel -lt ($cfg.TempCPUMax+10)){'WARN'}else{'FAIL'}
                        Add-GR 'Temperature CPU' $st "${cel}°C" "Source: $src"
                    } else {
                        Add-GR 'Temperature CPU' 'N/A' 'Source indisponible' 'Lancer OHM ou LHM'
                    }
                    $done++; SetProg([int]($done/$total*100))
                }

                if ($tests -contains 'Disque') {
                    $drv=Get-PSDrive -PSProvider FileSystem|Where-Object{$null -ne $_.Used -and $null -ne $_.Free}
                    foreach ($d in $drv) {
                        $tot=$d.Used+$d.Free; if($tot -le 0){continue}
                        $pct=[Math]::Round($d.Used/$tot*100,1); $fg=[Math]::Round($d.Free/1GB,1)
                        Add-GR "Disque $($d.Name):" (if($pct -lt $cfg.DiskPctMax){'OK'}elseif($pct -lt 95){'WARN'}else{'FAIL'}) "${pct}% utilise" "${fg} Go libres"
                    }
                    try {
                        $pd=Get-PhysicalDisk -ErrorAction Stop
                        foreach ($dk in $pd) {
                            try {
                                $rl=$dk|Get-StorageReliabilityCounter -ErrorAction Stop
                                $h=$dk.HealthStatus; $pts=@()
                                if($rl.Wear -gt 0){$pts+="Wear:$($rl.Wear)%"}
                                if($rl.Temperature -gt 0){$pts+="Temp:$($rl.Temperature)°C"}
                                if($rl.ReadErrorsTotal -gt 0){$pts+="RdErr:$($rl.ReadErrorsTotal)"}
                                if($null -ne $rl.WriteErrorsUncorrected -and $rl.WriteErrorsUncorrected -gt 0){$pts+="WrErr:$($rl.WriteErrorsUncorrected)"}
                                if($null -ne $rl.MediaErrors -and $rl.MediaErrors -gt 0){$pts+="MdErr:$($rl.MediaErrors)"}
                                $det=if($pts.Count -gt 0){$pts -join ' | '}else{'OK'}
                                $st='OK'
                                if($h -ne 'Healthy'){$st='WARN'}
                                if($null -ne $rl.WriteErrorsUncorrected -and $rl.WriteErrorsUncorrected -gt 0){$st='FAIL'}
                                if($rl.Wear -gt 90){$st='FAIL'}elseif($rl.Wear -gt 75){$st='WARN'}
                                $val="[$h]"; if($rl.Wear -gt 0){$val+=" W=$($rl.Wear)%"}
                                Add-GR "SMART: $($dk.FriendlyName)" $st $val $det
                            } catch { Add-GR "SMART: $($dk.FriendlyName)" 'N/A' 'Non supporte' '' }
                        }
                    } catch { Add-GR 'SMART' 'N/A' 'Indisponible' '' }
                    $tmp=[System.IO.Path]::GetTempFileName()
                    try {
                        $buf=[byte[]]::new([int](50*1MB)); [System.Random]::new().NextBytes($buf)
                        $sw2=[System.Diagnostics.Stopwatch]::StartNew()
                        [System.IO.File]::WriteAllBytes($tmp,$buf); $sw2.Stop()
                        $mbps=[Math]::Round(50/$sw2.Elapsed.TotalSeconds,0)
                        Add-GR 'Disque — Ecriture' (if($mbps -gt $cfg.DiskWriteMin){'OK'}elseif($mbps -gt 30){'WARN'}else{'FAIL'}) "${mbps} Mo/s" '50 Mo temp'
                    } catch { Add-GR 'Disque — Ecriture' 'N/A' 'Erreur I/O' '' }
                    finally { Remove-Item $tmp -ErrorAction SilentlyContinue }
                    $done++; SetProg([int]($done/$total*100))
                }

                if ($tests -contains 'GPU') {
                    $fnd=$false
                    $sm=@("$env:ProgramFiles\NVIDIA Corporation\NVSMI\nvidia-smi.exe","$env:SystemRoot\System32\nvidia-smi.exe")|Where-Object{Test-Path $_ -EA SilentlyContinue}|Select-Object -First 1
                    if (-not $sm) { try{$sm=(Get-Command 'nvidia-smi.exe' -ErrorAction Stop).Source}catch{} }
                    if ($sm) {
                        try {
                            $so=& $sm --query-gpu=name,temperature.gpu,utilization.gpu,memory.used,memory.total,power.draw,pstate --format=csv,noheader,nounits 2>$null
                            foreach ($ln in ($so -split "`n"|Where-Object{$_.Trim()-ne''})) {
                                $p=$ln -split ','|ForEach-Object{$_.Trim()}
                                if($p.Count -lt 5){continue}
                                $tg=if($p[1]-match'^\d'){[int]$p[1]}else{$null}
                                $ug=if($p[2]-match'^\d'){[int]$p[2]}else{$null}
                                $mu=if($p[3]-match'^\d'){[int]$p[3]}else{$null}
                                $mt=if($p[4]-match'^\d'){[int]$p[4]}else{$null}
                                $pw2=if($p.Count -gt 5 -and $p[5]-match'^\d'){[double]$p[5]}else{$null}
                                $st='OK'
                                if($tg -and $tg -gt $cfg.TempGPUMax){$st='FAIL'}elseif($tg -and $tg -gt ($cfg.TempGPUMax-10)){$st='WARN'}
                                $vl=if($tg){"${tg}°C"}else{'N/A'}
                                $dt=""; if($null -ne $ug){$dt+="Load:${ug}% "}; if($mu -and $mt){$dt+="VRAM:${mu}/${mt}MiB "}; if($null -ne $pw2){$dt+="Pwr:$([Math]::Round($pw2,1))W "}; $dt+="| nvidia-smi"
                                Add-GR "GPU: $($p[0])" $st $vl $dt
                                $fnd=$true
                            }
                        } catch {}
                    }
                    if (-not $fnd) {
                        try {
                            $og=@(Get-CimInstance -Namespace 'root/OpenHardwareMonitor' -ClassName Sensor -ErrorAction Stop|Where-Object{$_.SensorType -eq 'Temperature' -and $_.Name -match 'GPU'})
                            foreach ($g in $og) { if($g.Value -gt 0){$st=if($g.Value -lt $cfg.TempGPUMax){'OK'}else{'WARN'}; Add-GR "GPU:$($g.Name)" $st "$([Math]::Round($g.Value,1))°C" 'OHM'; $fnd=$true} }
                        } catch {}
                    }
                    if (-not $fnd) {
                        try {
                            $gc2=Get-CimInstance -ClassName Win32_VideoController -ErrorAction Stop
                            foreach ($g in $gc2) {
                                $vr=if($g.AdapterRAM -gt 0){[Math]::Round($g.AdapterRAM/1MB,0)}else{'?'}
                                Add-GR "GPU: $($g.Name)" (if($g.Status -eq 'OK'){'OK'}else{'WARN'}) "VRAM ${vr}Mo" "Driver:$($g.DriverVersion) | Info driver seul"
                            }
                        } catch { Add-GR 'GPU' 'N/A' 'Indisponible' '' }
                    }
                    $done++; SetProg([int]($done/$total*100))
                }

                if ($tests -contains 'Uptime') {
                    $os2=Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
                    $cp2=Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue|Select-Object -First 1
                    $ri2=Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue|Measure-Object -Property Capacity -Sum
                    if ($os2) {
                        $up=(Get-Date)-$os2.LastBootUpTime; $dy=[Math]::Floor($up.TotalDays)
                        $rGB=[Math]::Round($ri2.Sum/1GB,1)
                        $rU=[Math]::Round(($os2.TotalVisibleMemorySize-$os2.FreePhysicalMemory)/1MB,1)
                        $rT=[Math]::Round($os2.TotalVisibleMemorySize/1MB,1)
                        $rP=[Math]::Round($rU/$rT*100,1)
                        Add-GR 'Uptime'       (if($dy -lt $cfg.UptimeDaysWarn){'OK'}else{'WARN'}) "${dy}j $($up.Hours)h $($up.Minutes)m" "OS: $($os2.Caption)"
                        Add-GR 'RAM utilisee' (if($rP -lt $cfg.RamPctWarn){'OK'}elseif($rP -lt $cfg.RamPctFail){'WARN'}else{'FAIL'}) "${rP}%" "Physique: ${rGB} Go"
                        Add-GR 'CPU info'     'OK' $cp2.Name "Cores:$($cp2.NumberOfCores)/Logiques:$($cp2.NumberOfLogicalProcessors)"
                    }
                    $done++; SetProg([int]($done/$total*100))
                }

                return $latSmp
            }

            # ── Boucle runs ────────────────────────────────────────────────
            $loopN = 0
            do {
                $loopN++
                if ($loopN -gt 1) {
                    $dispatcher.Invoke([Action]{ $observableR.Clear(); $resultsRef.Clear() })
                }

                $latSamples = Set-Run-TestsOnce $selTests

                # Sauvegarde historique
                try {
                    if (-not (Test-Path $historyDir)) { New-Item -ItemType Directory -Path $historyDir -Force|Out-Null }
                    $f=Join-Path $historyDir "$(Get-Date -Format 'yyyy-MM-ddTHH-mm-ss').json"
                    @{Date=(Get-Date -Format 'o');Machine=$env:COMPUTERNAME;Results=@($resultsRef|ForEach-Object{@{Test=$_.Test;Status=$_.Status;Valeur=$_.Valeur;Detail=$_.Detail}})} |
                        ConvertTo-Json -Depth 5 | Set-Content $f -Encoding UTF8
                } catch {}

                $crit=@($resultsRef|Where-Object{$_.Status -eq 'FAIL'}).Count
                $warn=@($resultsRef|Where-Object{$_.Status -eq 'WARN'}).Count
                $ls=$latSamples

                $dispatcher.Invoke([Action]{
                    $pbProgress.Value=100
                    if (-not $isWatch) {
                        $btnRun.IsEnabled=$true; $btnWatch.IsEnabled=$true
                        $btnStop.IsEnabled=$false; $btnExport.IsEnabled=$true; $btnCompare.IsEnabled=$true
                    }
                    $ts=Get-Date -Format 'HH:mm:ss'
                    $lblStatus.Text=if($isWatch){"Watch run $loopN — $ts"}else{"Termine $ts"}

                    # Conclusion
                    if ($crit -gt 0) {
                        $lblConcl.Text='!! PROBLEME MATERIEL DETECTE — verifiez RAM et WHEA'
                        $lblConcl.Foreground=[Windows.Media.Brushes]::Salmon
                        $borderConcl.Background=[Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(58,16,16))
                    } elseif ($warn -gt 0) {
                        $lblConcl.Text='!! AVERTISSEMENTS — surveillance recommandee'
                        $lblConcl.Foreground=[Windows.Media.Brushes]::Gold
                        $borderConcl.Background=[Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(55,45,10))
                    } else {
                        $lblConcl.Text='Systeme stable — aucun probleme detecte'
                        $lblConcl.Foreground=[Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(74,222,128))
                        $borderConcl.Background=[Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(12,40,12))
                    }

                    # Sparkline
                    $sparkCanvas.Children.Clear()
                    if ($ls -and $ls.Count -gt 0) {
                        $w=$sparkCanvas.ActualWidth; $h=$sparkCanvas.ActualHeight
                        if($w -le 0){$w=900;$h=72}
                        $maxV=($ls|Measure-Object -Maximum).Maximum
                        $minV=($ls|Measure-Object -Minimum).Minimum
                        $rng=if($maxV-$minV -lt 0.01){1}else{$maxV-$minV}
                        $n=$ls.Count; $sx=$w/[Math]::Max($n-1,1)
                        # Grille
                        foreach ($pct in @(0.25,0.5,0.75)) {
                            $gl=[Windows.Shapes.Line]::new()
                            $gl.X1=0;$gl.X2=$w;$gl.Y1=$h-$pct*$h;$gl.Y2=$h-$pct*$h
                            $gl.Stroke=[Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(10,22,40))
                            $gl.StrokeThickness=1; $sparkCanvas.Children.Add($gl)|Out-Null
                        }
                        # Aire remplie sous la courbe
                        $poly=[Windows.Shapes.Polygon]::new()
                        $poly.Fill=[Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromArgb(30,56,189,248))
                        $poly.Stroke=[Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(56,189,248))
                        $poly.StrokeThickness=1.5
                        $pts=[Windows.Media.PointCollection]::new()
                        $pts.Add([Windows.Point]::new(0,$h))|Out-Null
                        for ($i=0;$i -lt $n;$i++) {
                            $x=$i*$sx; $nr=($ls[$i]-$minV)/$rng; $y=$h-($nr*($h-6))-3
                            $pts.Add([Windows.Point]::new($x,$y))|Out-Null
                        }
                        $pts.Add([Windows.Point]::new(($n-1)*$sx,$h))|Out-Null
                        $poly.Points=$pts; $sparkCanvas.Children.Add($poly)|Out-Null
                        # Labels
                        $lmx=[Windows.Controls.TextBlock]::new(); $lmx.Text=("{0:N2}ms"-f $maxV)
                        $lmx.Foreground=[Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(51,65,85))
                        $lmx.FontSize=9; [Windows.Controls.Canvas]::SetLeft($lmx,2); [Windows.Controls.Canvas]::SetTop($lmx,0)
                        $sparkCanvas.Children.Add($lmx)|Out-Null
                        $lmn=[Windows.Controls.TextBlock]::new(); $lmn.Text=("{0:N2}ms"-f $minV)
                        $lmn.Foreground=[Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(51,65,85))
                        $lmn.FontSize=9; [Windows.Controls.Canvas]::SetLeft($lmn,2); [Windows.Controls.Canvas]::SetTop($lmn,$h-14)
                        $sparkCanvas.Children.Add($lmn)|Out-Null
                    }
                })

                if ($isWatch -and -not $cts.IsCancellationRequested) {
                    Start-Sleep -Seconds $watchSec
                }

            } while ($isWatch -and -not $cts.IsCancellationRequested)

            if ($isWatch) {
                $dispatcher.Invoke([Action]{
                    $btnRun.IsEnabled=$true; $btnWatch.IsEnabled=$true
                    $btnStop.IsEnabled=$false; $btnExport.IsEnabled=$true; $btnCompare.IsEnabled=$true
                    $lblStatus.Text="Watch arrete"
                })
            }
        })

        $null = $ps.BeginInvoke()
    }

    # ── Evenements ─────────────────────────────────────────────────────────

    $btnRun.Add_Click({
        $sel=@(); foreach ($t in @('RAM','Latence','WHEA','Temp','Disque','GPU','Uptime')) {
            $chk=$window.FindName("chk$t"); if($chk -and $chk.IsChecked){$sel+=$t}
        }
        if ($sel.Count -eq 0) { $sel=@('RAM','Latence','WHEA','Temp','Disque','GPU','Uptime') }
        Start-GuiRunspace -selTests $sel -isWatch $false -watchSec 0
    })

    $btnWatch.Add_Click({
        $sel=@(); foreach ($t in @('Latence','Temp','Disque','Uptime')) {
            $chk=$window.FindName("chk$t"); if($chk -and $chk.IsChecked){$sel+=$t}
        }
        if ($sel.Count -eq 0) { $sel=@('Latence','Temp','Disque','Uptime') }
        Start-GuiRunspace -selTests $sel -isWatch $true -watchSec 30
    })

    $btnStop.Add_Click({
        if ($script:watchCts) { $script:watchCts.Cancel() }
        $lblStatus.Text='Arret demande...'
        $btnStop.IsEnabled=$false
    })

    $btnClear.Add_Click({
        $observableResults.Clear(); $results.Clear()
        $pbProgress.Value=0; $sparkCanvas.Children.Clear()
        $lblStatus.Text=''; $lblConcl.Text='En attente du diagnostic...'
        $lblConcl.Foreground=[Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(30,58,95))
        $borderConcl.Background=[Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(7,15,30))
        $btnExport.IsEnabled=$false; $btnCompare.IsEnabled=$false
    })

    $btnExport.Add_Click({
        $dlg=[Microsoft.Win32.SaveFileDialog]::new()
        $dlg.Title='Exporter le rapport'
        $dlg.Filter='HTML (*.html)|*.html|CSV (*.csv)|*.csv|Texte (*.txt)|*.txt'
        $dlg.DefaultExt='html'
        $dlg.FileName="occt81_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        if ($dlg.ShowDialog()) {
            $script:Export=$dlg.FileName; Export-Report
            $lblStatus.Text="Exporte : $($dlg.SafeFileName)"
        }
    })

    $btnCompare.Add_Click({
        $dlg=[Microsoft.Win32.OpenFileDialog]::new()
        $dlg.Title='Choisir un rapport JSON precedent'
        $dlg.Filter='JSON (*.json)|*.json'
        if (Test-Path $historyDir) { $dlg.InitialDirectory=$historyDir }
        if ($dlg.ShowDialog()) {
            Compare-History $dlg.FileName
            [System.Windows.MessageBox]::Show('Comparaison affichee dans la console.','occt81','OK','Information')|Out-Null
        }
    })

    $btnConfig.Add_Click({
        # Crée un config par défaut si absent
        $cfgDir = Join-Path $env:APPDATA 'occt81'
        if (-not (Test-Path $cfgDir)) { New-Item -ItemType Directory $cfgDir -Force|Out-Null }
        $cfgFile = Join-Path $cfgDir 'occt81.config.json'
        if (-not (Test-Path $cfgFile)) {
            @{
                LatMoyMax=20; LatP99Max=100; TempCPUMax=85; TempGPUMax=90
                DiskPctMax=85; DiskWriteMin=100; RamPctWarn=80; RamPctFail=90; UptimeDaysWarn=30
            } | ConvertTo-Json | Set-Content $cfgFile -Encoding UTF8
        }
        Start-Process notepad.exe $cfgFile
    })

    $null = $window.ShowDialog()
}

# ─────────────────────────────────────────────────────────────────────────────
# POINT D'ENTREE
# ─────────────────────────────────────────────────────────────────────────────
if ($GUI) { Show-Gui; exit 0 }

if ($Watch -gt 0) {
    Write-Header "occt81 v3.0 — $env:COMPUTERNAME"
    Start-WatchMode -intervalSec $Watch
    exit 0
}

# CLI standard
Write-Header "occt81 v3.0 — $env:COMPUTERNAME"
if (-not $Silent) {
    $osName = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
    Write-Info "OS     : $osName" -color 'DarkGray'
    Write-Info "Admin  : $(if($IsAdmin){'Oui'}else{'Non — WHEA/Temp/SMART indisponibles'})" `
        -color $(if ($IsAdmin) { 'DarkGray' } else { 'DarkYellow' })
    Write-Info "Tests  : $($testsToRun -join ', ')" -color 'DarkGray'
    if ($configPath) { Write-Info "Config : $configPath" -color 'DarkGray' }
}

Invoke-AllTests
Write-Summary
Save-History | Out-Null
if ($Compare) { Compare-History -jsonPath $Compare }
Export-Report
