#Requires -Version 5.1
<#
.SYNOPSIS
    occt81 — Outil de diagnostic universel Windows

.DESCRIPTION
    Teste RAM, latence CPU, erreurs WHEA, temperature, disques, GPU et uptime.
    Fonctionne en mode CLI ou GUI (WPF). Compatible Windows 10/11, user standard ou admin.

.PARAMETER Help
    Affiche l'aide courte.

.PARAMETER Man
    Affiche le manuel complet (man page).

.PARAMETER GUI
    Lance l'interface graphique WPF.

.PARAMETER Silent
    Supprime toute sortie console (utile pour export ou tache planifiee).

.PARAMETER Export
    Chemin du fichier de rapport. Formats : .txt, .csv, .html
    Exemple : -Export "C:\Rapports\diag.html"

.PARAMETER Tests
    Liste des tests a executer (virgule-separes).
    Valeurs : RAM, Latence, WHEA, Temp, Disque, GPU, Uptime, Tout (defaut)
    Exemple : -Tests "RAM,WHEA,Disque"

.PARAMETER Passes
    Nombre de passes pour le test RAM (defaut : 5)

.PARAMETER RamSize
    Taille du buffer RAM en Mo (defaut : 1024)

.EXAMPLE
    .\occt81.ps1
    Lance le diagnostic complet en mode CLI.

.EXAMPLE
    .\occt81.ps1 -GUI
    Lance l'interface graphique.

.EXAMPLE
    .\occt81.ps1 -Tests "RAM,WHEA" -Export "rapport.html"
    Teste uniquement RAM et WHEA, exporte en HTML.

.EXAMPLE
    .\occt81.ps1 -Silent -Export "C:\Logs\diag.csv"
    Mode silencieux avec export CSV pour tache planifiee.

.NOTES
    Version : 2.0
    Requis  : PowerShell 5.1+, Windows 10/11
    Droits  : Certains tests (WHEA, Temp WMI) necessitent des droits admin.
              Le script fonctionne sans admin mais ces tests seront marques N/A.
#>


[CmdletBinding()]
param(
    [switch]$Help,
    [switch]$Man,
    [switch]$GUI,
    [switch]$Silent,
    [string]$Export   = '',
    [string]$Tests    = 'Tout',
    [int]   $Passes   = 5,
    [int]   $RamSize  = 1024
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
# UAC
# ─────────────────────────────────────────────────────────────────────────────

if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File",$PSCommandPath -Verb RunAs
    exit
}


# ─────────────────────────────────────────────────────────────────────────────
#  AIDE
# ─────────────────────────────────────────────────────────────────────────────

if ($Help) {
    Write-Host @'

  occt81 v2.0 — Diagnostic systeme universel Windows
  =======================================================

  USAGE
      .\occt81.ps1 [options]

  OPTIONS
      -GUI                  Interface graphique WPF
      -Tests  <liste>       RAM, Latence, WHEA, Temp, Disque, GPU, Uptime, Tout
      -Export <fichier>     Rapport .txt / .csv / .html
      -Silent               Pas de sortie console
      -Passes <n>           Passes RAM (defaut : 5)
      -RamSize <Mo>         Buffer RAM en Mo (defaut : 20)
      -Help                 Cette aide
      -Man                  Manuel complet (Get-Help)

  EXEMPLES
      .\occt81.ps1
      .\occt81.ps1 -GUI
      .\occt81.ps1 -Tests "RAM,WHEA" -Export rapport.html
      .\occt81.ps1 -Silent -Export C:\Logs\diag.csv

'@ -ForegroundColor Cyan
    exit 0
}

if ($Man) {
    Get-Help $MyInvocation.MyCommand.Path -Full
    exit 0
}

# ─────────────────────────────────────────────────────────────────────────────
#  UTILITAIRES CLI
# ─────────────────────────────────────────────────────────────────────────────

function Write-Header($text) {
    if ($Silent) { return }
    $bar = '=' * 50
    Write-Host "`n$bar" -ForegroundColor DarkCyan
    Write-Host "  $text" -ForegroundColor Cyan
    Write-Host "$bar" -ForegroundColor DarkCyan
}

function Write-Section($text) {
    if ($Silent) { return }
    Write-Host "`n  -- $text" -ForegroundColor Yellow
}

function Write-Info($text, $color = 'Gray') {
    if ($Silent) { return }
    Write-Host "     $text" -ForegroundColor $color
}

function Get-StatusColor($status) {
    switch ($status) {
        'OK'   { 'Green' }
        'WARN' { 'Yellow' }
        'FAIL' { 'Red' }
        default { 'DarkGray' }
    }
}

$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)

# ─────────────────────────────────────────────────────────────────────────────
#  LISTE DES TESTS DEMANDES
# ─────────────────────────────────────────────────────────────────────────────

$allTests   = @('RAM','Latence','WHEA','Temp','Disque','GPU','Uptime')
$testsToRun = if ($Tests -eq 'Tout') { $allTests } else { $Tests -split ',' | ForEach-Object { $_.Trim() } }

function Should-Run($name) { $testsToRun -contains $name }

# ─────────────────────────────────────────────────────────────────────────────
#  COLLECTE DES RESULTATS (partagee CLI + GUI)
# ─────────────────────────────────────────────────────────────────────────────

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

function Add-Result($test, $status, $valeur, $detail = '') {
    $results.Add([PSCustomObject]@{
        Test   = $test
        Status = $status
        Valeur = $valeur
        Detail = $detail
        Heure  = (Get-Date -Format 'HH:mm:ss')
    })
}

# ─────────────────────────────────────────────────────────────────────────────
#  TEST 1 — RAM
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-RamTest {
    Write-Section "RAM TEST (buffer ${RamSize} Mo, $Passes passes)"
    $size      = [int]($RamSize * 1MB)
    $ramErrors = 0
    $rng       = [System.Random]::new()
    
    # On pré-alloue pour éviter les lags d'allocation durant le test
    $ref = [byte[]]::new($size)
    $rng.NextBytes($ref)

    for ($i = 1; $i -le $Passes; $i++) {
        Write-Info "Pass $i/$Passes..." -color 'DarkGray'
        
        # Force le nettoyage mémoire pour éviter les faux positifs de swap
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        $copy = [byte[]]::new($size)
        [Array]::Copy($ref, $copy, $size)

        # Comparaison ultra-rapide via .NET LINQ au lieu d'une boucle PowerShell
        if (-not [System.Linq.Enumerable]::SequenceEqual($ref, $copy)) {
            $ramErrors++
        }
    }

    $st = if ($ramErrors -eq 0) { 'OK' } else { 'FAIL' }
    $v  = if ($ramErrors -eq 0) { '0 erreur' } else { "$ramErrors passe(s) corrompue(s)" }
    Write-Info "Resultat : $v" -color (Get-StatusColor $st)
    Add-Result 'RAM' $st $v "Buffer ${RamSize} Mo x $Passes passes"
}

# ─────────────────────────────────────────────────────────────────────────────
#  TEST 2 — LATENCE CPU
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-LatenceTest {
    Write-Section "LATENCE CPU (200 echantillons)"
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

    $stats  = $lat | Measure-Object -Average -Maximum
    $avg    = $stats.Average
    $max    = $stats.Maximum
    $sorted = $lat | Sort-Object
    $p95    = $sorted[[Math]::Min([int]($samples * 0.95), $samples - 1)]
    $p99    = $sorted[[Math]::Min([int]($samples * 0.99), $samples - 1)]
    $txt    = "Avg={0:N2}ms  P95={1:N2}ms  P99={2:N2}ms  Max={3:N2}ms" -f $avg, $p95, $p99, $max

    Write-Info $txt -color 'Gray'
    $statusAvg = if ($avg -lt 20) { 'OK' } else { 'WARN' }
    $statusP99 = if ($p99 -lt 100) { 'OK' } else { 'WARN' }

    Add-Result 'Latence (moy)' $statusAvg ("{0:N2} ms" -f $avg) $txt
    Add-Result 'Latence (P99)' $statusP99 ("{0:N2} ms" -f $p99) $txt
}
# ─────────────────────────────────────────────────────────────────────────────
#  TEST 3 — WHEA
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-WheaTest {
    Write-Section "WHEA — Erreurs materielles"

    if (-not $IsAdmin) {
        Write-Info "! Droits admin requis — test ignore" -color 'DarkYellow'
        Add-Result 'WHEA total'    'N/A' 'Admin requis' ''
        Add-Result 'WHEA critique' 'N/A' 'Admin requis' ''
        return
    }

    $wheaEvents = $null
    try {
        $wheaEvents = Get-WinEvent -FilterHashtable @{
            LogName      = 'System'
            ProviderName = 'Microsoft-Windows-WHEA-Logger'
            Id           = 17, 18, 19, 20, 41, 4101
        } -MaxEvents 50 -ErrorAction SilentlyContinue
    } catch { $wheaEvents = @() }

    $wheaCount    = ($wheaEvents | Measure-Object).Count
    $wheaCritical = ($wheaEvents | Where-Object { $_.Id -eq 41 } | Measure-Object).Count

    if ($wheaCount -gt 0) {
        Write-Info "Derniers evenements :" -color 'DarkYellow'
        $wheaEvents | Select-Object -First 3 | ForEach-Object {
            $msg = if ($_.Message) { $_.Message.Substring(0, [Math]::Min(80, $_.Message.Length)) } else { '(no message)' }
            Write-Info ("[{0}] Id={1} {2}" -f $_.TimeCreated.ToString('dd/MM HH:mm'), $_.Id, $msg) -color 'DarkYellow'
        }
    }

        $statusTotal = if ($wheaCount -eq 0) { 'OK' } else { 'WARN' }
        $statusCrit  = if ($wheaCritical -eq 0) { 'OK' } else { 'FAIL' }

        Add-Result 'WHEA total' $statusTotal "$wheaCount evenement(s)" ''
        Add-Result 'WHEA critique' $statusCrit "$wheaCritical evenement(s) id=41" ''

}

function Invoke-TempTest {
    Write-Section "TEMPERATURE CPU"
    $celsius = $null

    try {
        $raw = (Get-CimInstance -Namespace 'root/WMI' -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction Stop | Select-Object -First 1).CurrentTemperature
        if ($raw -gt 0) { $celsius = [Math]::Round(($raw - 2732) / 10.0, 1) }
    } catch {}

    if ($null -eq $celsius) {
        try {
            $sensors = Get-CimInstance -Namespace 'root/OpenHardwareMonitor' -ClassName Sensor -ErrorAction Stop | 
                       Where-Object { $_.SensorType -eq 'Temperature' }
            
            $ohm = $sensors | Where-Object { $_.Name -match 'CPU|Package|Core|Tdie|Tctl' } | 
                   Sort-Object Value -Descending | Select-Object -First 1
            
            if ($ohm) { $celsius = [Math]::Round($ohm.Value, 1) }
        } catch {}
    }

    if ($null -ne $celsius -and $celsius -gt 0) {
        $st = if ($celsius -lt 85) { 'OK' } elseif ($celsius -lt 95) { 'WARN' } else { 'FAIL' }
        Write-Info "CPU : ${celsius}°C" -color (Get-StatusColor $st)
        Add-Result 'Temperature CPU' $st "${celsius}°C" "Capteur: $($ohm.Name)"
    } else {
        Write-Host "  [!] Activez 'WMI Support' dans les options d'OpenHardwareMonitor" -ForegroundColor Yellow
        Add-Result 'Temperature CPU' 'N/A' 'Source indisponible' 'Lancer OHM en Admin + Option WMI'
    }
}
# ─────────────────────────────────────────────────────────────────────────────
#  TEST 5 — DISQUES
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-DisqueTest {
    Write-Section "DISQUES"

    $disks = Get-PSDrive -PSProvider FileSystem | Where-Object { $null -ne $_.Used -and $null -ne $_.Free }
    foreach ($d in $disks) {
        $total = $d.Used + $d.Free
        if ($total -le 0) { continue }
        $pct   = [Math]::Round($d.Used / $total * 100, 1)
        $freeG = [Math]::Round($d.Free / 1GB, 1)
        $st    = if ($pct -lt 85) { 'OK' } elseif ($pct -lt 95) { 'WARN' } else { 'FAIL' }
        Write-Info "$($d.Name): — $pct% utilise — ${freeG} Go libres" -color (Get-StatusColor $st)
        Add-Result "Disque $($d.Name):" $st "${pct}% utilise" "${freeG} Go libres"
    }

    # Vitesse I/O basique (fichier temp 50 Mo)
    $tmpFile = [System.IO.Path]::GetTempFileName()
    try {
        $buf = [byte[]]::new([int](50 * 1MB))
        [System.Random]::new().NextBytes($buf)
        $sw2 = [System.Diagnostics.Stopwatch]::StartNew()
        [System.IO.File]::WriteAllBytes($tmpFile, $buf)
        $sw2.Stop()
        $mbps = [Math]::Round(50 / $sw2.Elapsed.TotalSeconds, 0)
        $st   = if ($mbps -gt 100) { 'OK' } elseif ($mbps -gt 30) { 'WARN' } else { 'FAIL' }
        Write-Info "Ecriture sequentielle : ${mbps} Mo/s" -color (Get-StatusColor $st)
        Add-Result 'Disque — Ecriture' $st "${mbps} Mo/s" 'Fichier temp 50 Mo'
    } catch {
        Add-Result 'Disque — Ecriture' 'N/A' 'Erreur I/O' $_.Exception.Message
    } finally {
        Remove-Item $tmpFile -ErrorAction SilentlyContinue
    }
}

# ─────────────────────────────────────────────────────────────────────────────
#  TEST 6 — GPU
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-GpuTest {
    Write-Section "GPU"
    try {
        $gpus = Get-CimInstance -ClassName Win32_VideoController -ErrorAction Stop
        foreach ($g in $gpus) {
            $vram = if ($g.AdapterRAM -gt 0) { [Math]::Round($g.AdapterRAM / 1MB, 0) } else { '?' }
            $st   = if ($g.Status -eq 'OK') { 'OK' } else { 'WARN' }
            Write-Info "$($g.Name) — VRAM ${vram} Mo" -color (Get-StatusColor $st)
            Add-Result "GPU — $($g.Name)" $st "VRAM ${vram} Mo" "Driver: $($g.DriverVersion)"
        }
    } catch {
        Add-Result 'GPU' 'N/A' 'Indisponible' $_.Exception.Message
    }
}

# ─────────────────────────────────────────────────────────────────────────────
#  TEST 7 — UPTIME & INFOS SYSTEME
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-UptimeTest {
    Write-Section "UPTIME & SYSTEME"
    $os      = Get-CimInstance Win32_OperatingSystem  -ErrorAction SilentlyContinue
    $cpu     = Get-CimInstance Win32_Processor        -ErrorAction SilentlyContinue | Select-Object -First 1
    $ramInfo = Get-CimInstance Win32_PhysicalMemory   -ErrorAction SilentlyContinue | Measure-Object -Property Capacity -Sum

    if ($os) {
        $uptime  = (Get-Date) - $os.LastBootUpTime
        $days    = [Math]::Floor($uptime.TotalDays)
        $ramGB   = [Math]::Round($ramInfo.Sum / 1GB, 1)
        $ramUsed = [Math]::Round(($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / 1MB, 1)
        $ramTot  = [Math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
        $ramPct  = [Math]::Round($ramUsed / $ramTot * 100, 1)

        Write-Info "OS      : $($os.Caption) $($os.Version)" -color 'Gray'
        Write-Info "CPU     : $($cpu.Name)" -color 'Gray'
        Write-Info "RAM     : ${ramGB} Go installes — ${ramPct}% utilises" -color 'Gray'
        Write-Info "Uptime  : ${days}j $($uptime.Hours)h $($uptime.Minutes)m" -color 'Gray'

        $statusUptime = if ($days -lt 30) { 'OK' } else { 'WARN' }

        $statusRAM = if ($ramPct -lt 80) { 'OK' }
                     elseif ($ramPct -lt 90) { 'WARN' }
                                  else { 'FAIL' }

                                  Add-Result 'Uptime' $statusUptime "${days}j $($uptime.Hours)h $($uptime.Minutes)m" "OS: $($os.Caption)"
                                  Add-Result 'RAM utilisee' $statusRAM "${ramPct}%" "Physique: ${ramGB} Go"
                                  Add-Result 'CPU info'     'OK' $cpu.Name "Cores: $($cpu.NumberOfCores) / Logiques: $($cpu.NumberOfLogicalProcessors)"
            } else {
                Add-Result 'Uptime' 'N/A' 'Indisponible' ''
            }
        }

# ─────────────────────────────────────────────────────────────────────────────
#  MOTEUR PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-AllTests {
    $results.Clear()
    if (Should-Run 'RAM')     { Invoke-RamTest }
    if (Should-Run 'Latence') { Invoke-LatenceTest }
    if (Should-Run 'WHEA')    { Invoke-WheaTest }
    if (Should-Run 'Temp')    { Invoke-TempTest }
    if (Should-Run 'Disque')  { Invoke-DisqueTest }
    if (Should-Run 'GPU')     { Invoke-GpuTest }
    if (Should-Run 'Uptime')  { Invoke-UptimeTest }
}

# ─────────────────────────────────────────────────────────────────────────────
#  EXPORT (TXT / CSV / HTML)
# ─────────────────────────────────────────────────────────────────────────────

function Export-Report {
    if (-not $Export) { return }
    $ext = [System.IO.Path]::GetExtension($Export).ToLower()

    switch ($ext) {
        '.csv' {
            $results | Select-Object Heure, Test, Status, Valeur, Detail |
                Export-Csv -Path $Export -NoTypeInformation -Encoding UTF8
        }
        '.html' {
            $rows = $results | ForEach-Object {
                $bg  = switch ($_.Status) { 'OK'{'#1a3a1a'} 'WARN'{'#3a3010'} 'FAIL'{'#3a1010'} default{'#1a1a2a'} }
                $col = switch ($_.Status) { 'OK'{'#4ade80'} 'WARN'{'#fbbf24'} 'FAIL'{'#f87171'} default{'#94a3b8'} }
                "<tr style='background:$bg'><td style='color:#94a3b8'>$($_.Heure)</td><td style='color:#e2e8f0'>$($_.Test)</td><td style='color:$col;font-weight:bold'>$($_.Status)</td><td style='color:#e2e8f0'>$($_.Valeur)</td><td style='color:#64748b;font-size:12px'>$($_.Detail)</td></tr>"
            }
            $crit    = ($results | Where-Object { $_.Status -eq 'FAIL' } | Measure-Object).Count
            $warns   = ($results | Where-Object { $_.Status -eq 'WARN' } | Measure-Object).Count
            $concl   = if ($crit -gt 0) { '&#9940; Probleme materiel detecte' } elseif ($warns -gt 0) { '&#9888; Avertissements' } else { '&#9989; Systeme stable' }
            $cColor  = if ($crit -gt 0) { '#f87171' } elseif ($warns -gt 0) { '#fbbf24' } else { '#4ade80' }
            $osName  = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
            $html    = @"
<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"><title>occt81</title>
<style>body{font-family:Consolas,monospace;background:#0f172a;color:#e2e8f0;margin:0;padding:32px}
h1{color:#38bdf8;font-size:22px;letter-spacing:2px}.meta{color:#64748b;font-size:13px;margin-bottom:24px}
table{width:100%;border-collapse:collapse;font-size:14px}
th{background:#1e293b;color:#475569;padding:10px 14px;text-align:left;font-weight:400}
td{padding:8px 14px;border-bottom:1px solid #1e293b}
.concl{margin-top:24px;padding:14px 20px;border-radius:6px;background:#1e293b;font-size:15px;color:$cColor;font-weight:bold}
</style></head><body>
<h1>occt81 v2.0</h1>
<div class="meta">$env:COMPUTERNAME &nbsp;|&nbsp; $osName &nbsp;|&nbsp; $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')</div>
<table><tr><th>HEURE</th><th>TEST</th><th>STATUT</th><th>VALEUR</th><th>DETAIL</th></tr>
$($rows -join "`n")
</table>
<div class="concl">$concl</div>
</body></html>
"@
            [System.IO.File]::WriteAllText($Export, $html, [System.Text.Encoding]::UTF8)
        }
        default {
            $lines = @("occt81 v2.0 — $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')", "Machine : $env:COMPUTERNAME", ('-' * 60))
            $results | ForEach-Object { $lines += "{0,-26} [{1,-4}]  {2}  {3}" -f $_.Test, $_.Status, $_.Valeur, $_.Detail }
            $lines | Set-Content -Path $Export -Encoding UTF8
        }
    }

    if (-not $Silent) {
        Write-Host "`n  Export : $Export" -ForegroundColor Green
    }
}

# ─────────────────────────────────────────────────────────────────────────────
#  RESUME CLI
# ─────────────────────────────────────────────────────────────────────────────

function Write-Summary {
    Write-Header 'RESUME'
    foreach ($r in $results) {
        $color = Get-StatusColor $r.Status
        $name   = $r.Test.PadRight(26)
        $status = "[$($r.Status.PadRight(4))]"
        $val    = $r.Valeur.PadRight(12)
        $detail = if ($r.Detail) { $r.Detail } else { "" }
        
        Write-Host "  $name $status  $val  $detail" -ForegroundColor $color
    }
    $bar      = '=' * 50
    $critical = ($results | Where-Object { $_.Status -eq 'FAIL' } | Measure-Object).Count
    $warns    = ($results | Where-Object { $_.Status -eq 'WARN' } | Measure-Object).Count
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
#  GUI WPF
# ─────────────────────────────────────────────────────────────────────────────

function Show-Gui {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

    [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="occt81 v2.0" Height="680" Width="860"
        Background="#0f172a" FontFamily="Consolas" WindowStartupLocation="CenterScreen">
  <Window.Resources>
    <Style TargetType="Button" x:Key="Btn">
      <Setter Property="Background" Value="#1e3a5f"/>
      <Setter Property="Foreground" Value="#38bdf8"/>
      <Setter Property="BorderBrush" Value="#38bdf8"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="14,7"/>
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#164e8a"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.35"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#94a3b8"/>
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Margin" Value="0,3"/>
    </Style>
  </Window.Resources>
  <Grid Margin="20">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <!-- Titre -->
    <StackPanel Grid.Row="0" Margin="0,0,0,16">
      <TextBlock Text="occt81 v2.0" FontSize="22" FontWeight="Bold" Foreground="#38bdf8"/>
      <TextBlock x:Name="lblMachine" FontSize="12" Foreground="#475569" Margin="0,4,0,0"/>
    </StackPanel>
    <!-- Checkboxes tests -->
    <Border Grid.Row="1" Background="#1e293b" CornerRadius="6" Padding="14,10" Margin="0,0,0,12">
      <StackPanel>
        <TextBlock Text="TESTS A EXECUTER" FontSize="11" Foreground="#475569" Margin="0,0,0,8"/>
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
    <WrapPanel Grid.Row="2" Margin="0,0,0,14">
      <Button x:Name="btnRun"    Content="LANCER"    Style="{StaticResource Btn}" Margin="0,0,10,0"/>
      <Button x:Name="btnExport" Content="EXPORTER"  Style="{StaticResource Btn}" Margin="0,0,10,0" IsEnabled="False"/>
      <Button x:Name="btnClear"  Content="EFFACER"   Style="{StaticResource Btn}"/>
      <TextBlock x:Name="lblStatus" VerticalAlignment="Center" Foreground="#64748b" FontSize="13" Margin="16,0,0,0"/>
    </WrapPanel>
    <!-- DataGrid -->
    <Border Grid.Row="3" Background="#1e293b" CornerRadius="6">
      <DataGrid x:Name="dgResults" AutoGenerateColumns="False"
                Background="#1e293b" Foreground="#e2e8f0"
                GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#0f172a"
                RowBackground="#1e293b" AlternatingRowBackground="#162032"
                BorderThickness="0" CanUserAddRows="False" IsReadOnly="True"
                ColumnHeaderHeight="32" FontSize="13">
        <DataGrid.ColumnHeaderStyle>
          <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#0f172a"/>
            <Setter Property="Foreground" Value="#475569"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize"   Value="11"/>
            <Setter Property="Padding"    Value="12,0"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Setter Property="BorderBrush" Value="#1e293b"/>
          </Style>
        </DataGrid.ColumnHeaderStyle>
        <DataGrid.Columns>
          <DataGridTextColumn Header="HEURE"  Binding="{Binding Heure}"  Width="70"/>
          <DataGridTextColumn Header="TEST"   Binding="{Binding Test}"   Width="190"/>
          <DataGridTemplateColumn Header="STATUT" Width="80">
            <DataGridTemplateColumn.CellTemplate>
              <DataTemplate>
                <TextBlock Text="{Binding Status}" FontWeight="Bold" Padding="12,4">
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
                          <Setter Property="Foreground" Value="#475569"/>
                        </DataTrigger>
                      </Style.Triggers>
                    </Style>
                  </TextBlock.Style>
                </TextBlock>
              </DataTemplate>
            </DataGridTemplateColumn.CellTemplate>
          </DataGridTemplateColumn>
          <DataGridTextColumn Header="VALEUR"  Binding="{Binding Valeur}" Width="160"/>
          <DataGridTextColumn Header="DETAIL"  Binding="{Binding Detail}" Width="*"/>
        </DataGrid.Columns>
      </DataGrid>
    </Border>
    <!-- Barre de progression -->
    <ProgressBar x:Name="pbProgress" Grid.Row="4" Height="4" Margin="0,10,0,6"
                 Background="#1e293b" Foreground="#38bdf8" BorderThickness="0"
                 Minimum="0" Maximum="100" Value="0"/>
    <!-- Conclusion -->
    <Border x:Name="borderConcl" Grid.Row="5" CornerRadius="6" Padding="14,10"
            Background="#1e293b" Margin="0,4,0,0">
      <TextBlock x:Name="lblConcl" FontSize="14" FontWeight="Bold"
                 Foreground="#475569" Text="En attente du diagnostic..."/>
    </Border>
  </Grid>
</Window>
'@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $dgResults   = $window.FindName('dgResults')
    $btnRun      = $window.FindName('btnRun')
    $btnExport   = $window.FindName('btnExport')
    $btnClear    = $window.FindName('btnClear')
    $lblStatus   = $window.FindName('lblStatus')
    $lblConcl    = $window.FindName('lblConcl')
    $lblMachine  = $window.FindName('lblMachine')
    $pbProgress  = $window.FindName('pbProgress')
    $borderConcl = $window.FindName('borderConcl')

    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $lblMachine.Text = "$env:COMPUTERNAME  |  $($osInfo.Caption)  |  $(if($IsAdmin){'Admin'}else{'Utilisateur standard'})"

    $observableResults = [System.Collections.ObjectModel.ObservableCollection[PSCustomObject]]::new()
    $dgResults.ItemsSource = $observableResults

    $btnRun.Add_Click({
        $btnRun.IsEnabled    = $false
        $btnExport.IsEnabled = $false
        $lblStatus.Text      = 'Diagnostic en cours...'
        $pbProgress.Value    = 0
        $observableResults.Clear()
        $results.Clear()

        $sel = @()
        foreach ($t in @('RAM','Latence','WHEA','Temp','Disque','GPU','Uptime')) {
            $chk = $window.FindName("chk$t")
            if ($chk -and $chk.IsChecked) { $sel += $t }
        }
        $script:testsToRun = if ($sel.Count -gt 0) { $sel } else { $allTests }

        $dispatcher = $window.Dispatcher
        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = 'STA'
        $rs.ThreadOptions  = 'ReuseThread'
        $rs.Open()

        foreach ($v in @('RamSize','Passes','IsAdmin')) { $rs.SessionStateProxy.SetVariable($v, (Get-Variable $v).Value) }
        $rs.SessionStateProxy.SetVariable('testsToRun',  $script:testsToRun)
        $rs.SessionStateProxy.SetVariable('allTests',    $allTests)
        $rs.SessionStateProxy.SetVariable('dispatcher',  $dispatcher)
        $rs.SessionStateProxy.SetVariable('observableR', $observableResults)
        $rs.SessionStateProxy.SetVariable('pbProgress',  $pbProgress)
        $rs.SessionStateProxy.SetVariable('btnRun',      $btnRun)
        $rs.SessionStateProxy.SetVariable('btnExport',   $btnExport)
        $rs.SessionStateProxy.SetVariable('lblStatus',   $lblStatus)
        $rs.SessionStateProxy.SetVariable('lblConcl',    $lblConcl)
        $rs.SessionStateProxy.SetVariable('borderConcl', $borderConcl)
        $rs.SessionStateProxy.SetVariable('resultsRef',  $results)

        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs

        $null = $ps.AddScript({

            function Add-GR($test, $status, $valeur, $detail = '') {
                $obj = [PSCustomObject]@{ Test=$test; Status=$status; Valeur=$valeur; Detail=$detail; Heure=(Get-Date -Format 'HH:mm:ss') }
                $resultsRef.Add($obj)
                $dispatcher.Invoke([Action]{ $observableR.Add($obj) })
            }

            function Upd-Prog($val) { $dispatcher.Invoke([Action]{ $pbProgress.Value = $val }) }

            $total = $testsToRun.Count
            $done  = 0

            # RAM
            if ($testsToRun -contains 'RAM') {
                $size = [int]($RamSize * 1MB); $ramErrors = 0
                $rng = [System.Random]::new(); $ref = [byte[]]::new($size); $rng.NextBytes($ref)
                for ($i = 1; $i -le $Passes; $i++) {
                    [System.GC]::Collect() 
                    [System.GC]::WaitForPendingFinalizers()
                    $copy = [byte[]]::new($size)
                    [Array]::Copy($ref, $copy, $size)
                    if (-not [System.Linq.Enumerable]::SequenceEqual($ref, $copy)) { $ramErrors++ }
                }
                Add-GR 'RAM' (if ($ramErrors -eq 0) { 'OK' } else { 'FAIL' }) (if ($ramErrors -eq 0) { '0 erreur' } else { "$ramErrors passe(s) corrompue(s)" }) "Buffer ${RamSize} Mo x $Passes passes"
                $done++; Upd-Prog ([int]($done / $total * 100))
            }

            # LATENCE
            if ($testsToRun -contains 'Latence') {
                $warmup = [System.Diagnostics.Stopwatch]::StartNew()
                while($warmup.ElapsedMilliseconds -lt 300) { $null = 1 + 1 }
                $warmup.Stop()

                $samples = 200; $lat = [double[]]::new($samples)
                $sw = [System.Diagnostics.Stopwatch]::new()
                for ($i = 0; $i -lt $samples; $i++) {
                    $sw.Restart(); $x = 0
                    for ($j = 0; $j -lt 50000; $j++) { $x = $x -bxor ($j * 7) } # Charge augmentée
                    $sw.Stop(); $lat[$i] = $sw.Elapsed.TotalMilliseconds
                }
                $avg    = ($lat | Measure-Object -Average).Average
                $sorted = $lat | Sort-Object
                $p95    = $sorted[[Math]::Min([int]($samples * 0.95), $samples - 1)]
                $p99    = $sorted[[Math]::Min([int]($samples * 0.99), $samples - 1)]
                $max    = ($lat | Measure-Object -Maximum).Maximum
                $txt    = "Avg={0:N2}ms P95={1:N2}ms P99={2:N2}ms Max={3:N2}ms" -f $avg, $p95, $p99, $max
                
                $statusAvg = if ($avg -lt 10) { 'OK' } else { 'WARN' }
                $statusP99 = if ($p99 -lt 30) { 'OK' } else { 'WARN' }
                
                Add-GR 'Latence (moy)' $statusAvg ("{0:N2} ms" -f $avg) $txt
                Add-GR 'Latence (P99)' $statusP99 ("{0:N2} ms" -f $p99) $txt
                $done++; Upd-Prog ([int]($done / $total * 100))
            }

            # WHEA
            if ($testsToRun -contains 'WHEA') {
                if (-not $IsAdmin) {
                    Add-GR 'WHEA' 'N/A' 'Admin requis' ''
                } else {
                    $ev = $null
                    try { $ev = Get-WinEvent -FilterHashtable @{ LogName='System'; ProviderName='Microsoft-Windows-WHEA-Logger'; Id=17,18,19,20,41,4101 } -MaxEvents 50 -ErrorAction SilentlyContinue } catch { $ev = @() }
                    $wc = ($ev | Measure-Object).Count
                    $cc = ($ev | Where-Object { $_.Id -eq 41 } | Measure-Object).Count
                    Add-GR 'WHEA total'    (if ($wc -eq 0) { 'OK' } else { 'WARN' }) "$wc evenement(s)" ''
                    Add-GR 'WHEA critique' (if ($cc -eq 0) { 'OK' } else { 'FAIL' }) "$cc evenement(s) id=41" ''
                }
                $done++; Upd-Prog ([int]($done / $total * 100))
            }

            # TEMP
            if ($testsToRun -contains 'Temp') {
                $celsius = $null
                try { $raw = (Get-CimInstance -Namespace 'root/WMI' -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction Stop | Select-Object -First 1).CurrentTemperature; $celsius = [Math]::Round(($raw-2732)/10.0,1) } catch {}
                if ($null -eq $celsius) {
                    try { $ohm = Get-CimInstance -Namespace 'root/OpenHardwareMonitor' -ClassName Sensor -ErrorAction Stop | Where-Object { $_.SensorType -eq 'Temperature' -and $_.Name -match 'CPU' } | Select-Object -First 1; if ($ohm) { $celsius = [Math]::Round($ohm.Value,1) } } catch {}
                }
                if ($null -ne $celsius) {
                    $st = if ($celsius -lt 85) { 'OK' } elseif ($celsius -lt 95) { 'WARN' } else { 'FAIL' }
                    Add-GR 'Temperature CPU' $st "${celsius}°C" ''
                } else {
                    Add-GR 'Temperature CPU' 'N/A' 'Source indisponible' 'Installer OpenHardwareMonitor'
                }
                $done++; Upd-Prog ([int]($done / $total * 100))
            }

            # DISQUE
            if ($testsToRun -contains 'Disque') {
                $disks = Get-PSDrive -PSProvider FileSystem | Where-Object { $null -ne $_.Used -and $null -ne $_.Free }
                foreach ($d in $disks) {
                    $tot = $d.Used + $d.Free; if ($tot -le 0) { continue }
                    $pct = [Math]::Round($d.Used/$tot*100,1); $freeG = [Math]::Round($d.Free/1GB,1)
                    Add-GR "Disque $($d.Name):" (if ($pct -lt 85) { 'OK' } elseif ($pct -lt 95) { 'WARN' } else { 'FAIL' }) "${pct}% utilise" "${freeG} Go libres"
                }
                $tmpFile = [System.IO.Path]::GetTempFileName()
                try {
                    $buf = [byte[]]::new([int](50*1MB)); [System.Random]::new().NextBytes($buf)
                    $sw2 = [System.Diagnostics.Stopwatch]::StartNew()
                    [System.IO.File]::WriteAllBytes($tmpFile, $buf); $sw2.Stop()
                    $mbps = [Math]::Round(50/$sw2.Elapsed.TotalSeconds,0)
                    Add-GR 'Disque — Ecriture' (if ($mbps -gt 100) { 'OK' } elseif ($mbps -gt 30) { 'WARN' } else { 'FAIL' }) "${mbps} Mo/s" '50 Mo temp'
                } catch { Add-GR 'Disque — Ecriture' 'N/A' 'Erreur I/O' '' }
                finally { Remove-Item $tmpFile -ErrorAction SilentlyContinue }
                $done++; Upd-Prog ([int]($done / $total * 100))
            }

            # GPU
            if ($testsToRun -contains 'GPU') {
                try {
                    $gpus = Get-CimInstance -ClassName Win32_VideoController -ErrorAction Stop
                    foreach ($g in $gpus) {
                        $vr = if ($g.AdapterRAM -gt 0) { [Math]::Round($g.AdapterRAM/1MB,0) } else { '?' }
                        Add-GR "GPU — $($g.Name)" (if ($g.Status -eq 'OK') { 'OK' } else { 'WARN' }) "VRAM $vr Mo" "Driver: $($g.DriverVersion)"
                    }
                } catch { Add-GR 'GPU' 'N/A' 'Indisponible' '' }
                $done++; Upd-Prog ([int]($done / $total * 100))
            }

            # UPTIME
            if ($testsToRun -contains 'Uptime') {
                $os2 = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
                $cpu2 = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
                $ri   = Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue | Measure-Object -Property Capacity -Sum
                if ($os2) {
                    $up   = (Get-Date) - $os2.LastBootUpTime
                    $days = [Math]::Floor($up.TotalDays)
                    $rGB  = [Math]::Round($ri.Sum/1GB,1)
                    $rU   = [Math]::Round(($os2.TotalVisibleMemorySize-$os2.FreePhysicalMemory)/1MB,1)
                    $rT   = [Math]::Round($os2.TotalVisibleMemorySize/1MB,1)
                    $rP   = [Math]::Round($rU/$rT*100,1)
                    Add-GR 'Uptime' (if ($days -lt 30) { 'OK' } else { 'WARN' }) "${days}j $($up.Hours)h $($up.Minutes)m" "OS: $($os2.Caption)"
                    Add-GR 'RAM utilisee' (if ($rP -lt 80) { 'OK' } elseif ($rP -lt 90) { 'WARN' } else { 'FAIL' }) "${rP}%" "Physique: ${rGB} Go"
                    Add-GR 'CPU info'     'OK' $cpu2.Name "Cores: $($cpu2.NumberOfCores) / Logiques: $($cpu2.NumberOfLogicalProcessors)"
                }
                $done++; Upd-Prog ([int]($done / $total * 100))
            }

            # Conclusion finale
            $crit2 = ($resultsRef | Where-Object { $_.Status -eq 'FAIL' } | Measure-Object).Count
            $warn2 = ($resultsRef | Where-Object { $_.Status -eq 'WARN' } | Measure-Object).Count

            $dispatcher.Invoke([Action]{
                $pbProgress.Value    = 100
                $btnRun.IsEnabled    = $true
                $btnExport.IsEnabled = $true
                $lblStatus.Text      = "Termine $(Get-Date -Format 'HH:mm:ss')"

                if ($crit2 -gt 0) {
                    $lblConcl.Text       = '!! PROBLEME MATERIEL DETECTE — verifiez RAM et journaux WHEA'
                    $lblConcl.Foreground = [Windows.Media.Brushes]::Salmon
                    $borderConcl.Background = [Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(58,16,16))
                } elseif ($warn2 -gt 0) {
                    $lblConcl.Text       = '!! AVERTISSEMENTS — surveillance recommandee'
                    $lblConcl.Foreground = [Windows.Media.Brushes]::Gold
                    $borderConcl.Background = [Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(58,48,16))
                } else {
                    $lblConcl.Text       = 'Systeme stable — aucun probleme detecte'
                    $lblConcl.Foreground = [Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(74,222,128))
                    $borderConcl.Background = [Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(26,58,26))
                }
            })
        })

        $null = $ps.BeginInvoke()
    })

    $btnClear.Add_Click({
        $observableResults.Clear()
        $results.Clear()
        $pbProgress.Value    = 0
        $lblStatus.Text      = ''
        $lblConcl.Text       = 'En attente du diagnostic...'
        $lblConcl.Foreground = [Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(71,85,105))
        $borderConcl.Background = [Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(30,41,59))
        $btnExport.IsEnabled = $false
    })

    $btnExport.Add_Click({
        $dlg = [Microsoft.Win32.SaveFileDialog]::new()
        $dlg.Title      = 'Exporter le rapport'
        $dlg.Filter     = 'Rapport HTML (*.html)|*.html|Fichier CSV (*.csv)|*.csv|Texte (*.txt)|*.txt'
        $dlg.DefaultExt = 'html'
        if ($dlg.ShowDialog()) {
            $script:Export = $dlg.FileName
            Export-Report
        }
    })

    $null = $window.ShowDialog()
}

# ─────────────────────────────────────────────────────────────────────────────
#  POINT D'ENTREE
# ─────────────────────────────────────────────────────────────────────────────

if ($GUI) {
    Show-Gui
} else {
    Write-Header "occt81 v2.0 — $env:COMPUTERNAME"

    if (-not $Silent) {
        $osName = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
        Write-Info "OS     : $osName" -color 'DarkGray'
        Write-Info "Admin  : $(if($IsAdmin){'Oui'}else{'Non — certains tests indisponibles'})" -color $(if($IsAdmin){'DarkGray'}else{'DarkYellow'})
        Write-Info "Tests  : $($testsToRun -join ', ')" -color 'DarkGray'
    }

    Invoke-AllTests
    Write-Summary
    Export-Report
}
