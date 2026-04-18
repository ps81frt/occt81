#Requires -Version 5.1
<#
.SYNOPSIS
occt81 v4.2 CLI — Outil de diagnostic universel Windows
.DESCRIPTION
Test RAM (patterns complets + walking bits + bande passante), scheduling CPU (jitter + freq effective),
erreurs WHEA, temperature (fallback chain),
disques (SMART + vitesse I/O), GPU (nvidia-smi / OHM / WMI) et uptime.
Mode CLI uniquement. Historique JSON automatique, mode Watch, comparaison de runs, config JSON.
.PARAMETER Help
Affiche l aide courte.
.PARAMETER Man
Affiche le manuel complet.
.PARAMETER Silent
Supprime toute sortie console.
.PARAMETER Export
Chemin du fichier de rapport. Formats : .txt, .csv, .html, .json
.PARAMETER Tests
Liste des tests (virgules). Valeurs : RAM, Latence, WHEA, Temp, Disque, GPU, Uptime, Tout.
.PARAMETER Passes
Nombre de passes RAM (defaut : 5)
.PARAMETER RamSize
Taille buffer RAM en Mo (defaut : 512)
.PARAMETER Watch
Mode surveillance continue — relance les tests legers toutes les N secondes.
.PARAMETER Compare
Chemin d un rapport JSON precedent pour comparer avec le run actuel.
.PARAMETER Config
Chemin d un fichier occt81.config.json (seuils personnalises).
.PARAMETER UploadDPaste
Upload automatique du rapport sur DPaste et affiche le lien.
.PARAMETER UploadGoFile
Upload automatique du rapport sur GoFile et affiche le lien.
.EXAMPLE
.\occt81.ps1
.\occt81.ps1 -Tests "RAM,WHEA" -Export rapport.html -UploadDPaste
.\occt81.ps1 -Watch 60
.\occt81.ps1 -Compare "$env:APPDATA\occt81\history\2024-01-01.json"
#>
[CmdletBinding()]
param(
    [switch]$Help,
    [switch]$Man,
    [switch]$Silent,
    [string]$Export   = '',
    [string[]]$Tests,
    [int]   $Passes   = 5,
    [int]   $RamSize  = 512,
    [int]   $Watch    = 0,
    [string]$Compare  = '',
    [string]$Config   = '',
    [switch]$UploadDPaste,
    [switch]$UploadGoFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#─────────────────────────────────────────────────────────────────────────────
# PARSING $Tests — accepte "RAM,WHEA" ou @('RAM','WHEA') ou "RAM WHEA"
#─────────────────────────────────────────────────────────────────────────────
if ($Tests) {
    $Tests = $Tests |
        ForEach-Object { $_ -split '[,\s]+' } |
        ForEach-Object { $_.Trim() } |
        Where-Object   { $_ -ne '' }
}

#─────────────────────────────────────────────────────────────────────────────
# ADMIN CHECK — demande confirmation avant elevation
#─────────────────────────────────────────────────────────────────────────────
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

#─────────────────────────────────────────────────────────────────────────────
# RESOLUTION DES TESTS
#─────────────────────────────────────────────────────────────────────────────
$allTests   = @('RAM', 'Latence', 'WHEA', 'Temp', 'Disque', 'GPU', 'Uptime')
$watchTests = @('Latence', 'Temp', 'Disque', 'Uptime')

function Resolve-Tests {
    param([string[]]$raw)

    if (-not $raw) { return @('RAM', 'Uptime') }

    if ($raw | Where-Object { $_ -match '^(Tout|All)$' }) {
        return @($allTests)
    }

    $resolved = @()
    foreach ($item in $raw) {
        foreach ($t in ($item -split '[,\s]+')) {
            $t = $t.Trim()
            if (-not $t) { continue }

            $match = $allTests | Where-Object { $_ -ieq $t }
            if ($match) {
                $resolved += $match
            }
            elseif (-not $Silent) {
                Write-Warning "Test inconnu ignore : '$t' (valeurs valides : $($allTests -join ', '))"
            }
        }
    }

    if ($resolved.Count -eq 0) { return @('RAM', 'Uptime') }
    return $resolved | Select-Object -Unique
}

$script:testsToRun = Resolve-Tests $Tests

function Set-ShouldRun { param([string]$name); return $script:testsToRun -contains $name }

#─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION — defaults + chargement config.json si present
#─────────────────────────────────────────────────────────────────────────────
$cfg = @{
    # Seuils RAM
    RamPctWarn      = 80
    RamPctFail      = 90
    RamSizeAutoMB   = 0
    RamPasses       = 0
    RamMaxThreads   = 8
    RamOsReserveMB  = 256
    RamPerThreadMin = 64
    # Seuils CPU scheduling
    LatMoyMax       = 20
    LatP99Max       = 100
    LatJitterWarn   = 3.0
    LatJitterFail   = 8.0
    LatSamples      = 500
    # Seuils temperature
    TempCPUMax      = 85
    TempGPUMax      = 90
    # Seuils disque
    DiskPctMax      = 85
    DiskWriteMin    = 100
    DiskReadMin     = 100
    DiskTestSizeMB  = 1024
    # Uptime
    UptimeDaysWarn  = 30
    # OHM
    OhmTimeoutSec   = 25
}

$defaultConfig = Join-Path $env:APPDATA 'occt81\occt81.config.json'
$configPath    = if ($Config -and (Test-Path $Config)) { $Config }
                 elseif (Test-Path $defaultConfig)     { $defaultConfig }
                 else                                  { $null }

if ($configPath) {
    try {
        $loaded = Get-Content $configPath -Raw | ConvertFrom-Json
        # Copier les cles dans un tableau fixe AVANT d iterер :
        # foreach ($k in $cfg.Keys) modifie la collection pendant l enumeration
        # => "Collection was modified; enumeration operation may not execute."
        $cfgKeys = @($cfg.Keys)
        foreach ($k in $cfgKeys) {
            if ($null -ne $loaded.$k) {
                # Conserver le type : float pour les ratios, int pour le reste
                $cfg[$k] = if ($cfg[$k] -is [double] -or $loaded.$k -match '\.')
                               { [double]$loaded.$k }
                           else
                               { [int]$loaded.$k }
            }
        }
    }
    catch {
        if (-not $Silent -and (Test-Path $configPath)) {
            Write-Warning "Config invalide (JSON corrompu ?) : $configPath — $($_.Exception.Message)"
        }
    }
}

# Generation du fichier config exemple si absent
if (-not (Test-Path $defaultConfig) -and -not $Silent) {
    try {
        $cfgDir = Split-Path $defaultConfig
        if (-not (Test-Path $cfgDir)) { New-Item -ItemType Directory $cfgDir -Force | Out-Null }
        $cfgTemplate = [ordered]@{
            '_aide'          = "Toutes les valeurs sont optionnelles. Supprimer une cle = valeur par defaut."
            RamSizeAutoMB    = 0
            '_RamSizeAutoMB' = "0=auto (25% RAM physique, min 512, max 4096). >0=fixe en Mo. Priorite sur -RamSize CLI."
            RamPasses        = 0
            '_RamPasses'     = "0=utilise -Passes CLI. >0 override."
            RamMaxThreads    = $cfg.RamMaxThreads
            '_RamMaxThreads' = "Nb max threads RAM (1-32). Defaut: 8."
            RamOsReserveMB   = $cfg.RamOsReserveMB
            '_RamOsReserveMB'= "Mo reserves pour l OS (defaut: 256)."
            RamPerThreadMin  = $cfg.RamPerThreadMin
            RamPctWarn       = $cfg.RamPctWarn
            RamPctFail       = $cfg.RamPctFail
            LatMoyMax        = $cfg.LatMoyMax
            '_LatMoyMax'     = "ms moyenne scheduling -> WARN si depasse."
            LatP99Max        = $cfg.LatP99Max
            LatJitterWarn    = $cfg.LatJitterWarn
            LatJitterFail    = $cfg.LatJitterFail
            LatSamples       = $cfg.LatSamples
            '_LatSamples'    = "Nb echantillons scheduling (defaut: 500)."
            TempCPUMax       = $cfg.TempCPUMax
            TempGPUMax       = $cfg.TempGPUMax
            DiskPctMax       = $cfg.DiskPctMax
            DiskWriteMin     = $cfg.DiskWriteMin
            DiskReadMin      = $cfg.DiskReadMin
            DiskTestSizeMB   = $cfg.DiskTestSizeMB
            '_DiskTestSizeMB'= "Mo utilises par le test sequentiel dd (defaut: 1024)."
            UptimeDaysWarn   = $cfg.UptimeDaysWarn
            OhmTimeoutSec    = $cfg.OhmTimeoutSec
        }
        $cfgTemplate | ConvertTo-Json | Set-Content $defaultConfig -Encoding UTF8
        Write-Host "    Config template cree : $defaultConfig" -ForegroundColor DarkGray
    } catch { <# non bloquant #> }
}

#─────────────────────────────────────────────────────────────────────────────
# AIDE
#─────────────────────────────────────────────────────────────────────────────
if ($Help) {
    Write-Host @'
occt81 v4.2 CLI — Diagnostic systeme universel Windows
USAGE    .\occt81.ps1 [options]
OPTIONS
  -Tests       RAM, Latence (scheduling CPU), WHEA, Temp, Disque, GPU, Uptime, Tout
  -Export      Rapport .txt / .csv / .html / .json
  -Silent      Pas de sortie console
  -Passes      Passes RAM (defaut : 5)
  -RamSize     Buffer RAM (defaut : 512 Mo)
  -Watch       Mode surveillance toutes les N secondes
  -Compare     Compare avec un run precedent (JSON)
  -Config      Fichier de seuils personnalises
  -UploadDPaste   Upload du rapport sur DPaste
  -UploadGoFile   Upload du rapport sur GoFile
  -Help        Cette aide
  -Man         Manuel complet
EXEMPLES
  .\occt81.ps1
  .\occt81.ps1 -Tests "RAM,WHEA" -Export rapport.html
  .\occt81.ps1 -Watch 60
  .\occt81.ps1 -Tests Tout -Export rapport.json -UploadDPaste
  .\occt81.ps1 -Compare "$env:APPDATA\occt81\history\2024-01-01.json"
'@ -ForegroundColor Cyan
    exit 0
}
if ($Man) { Get-Help $MyInvocation.MyCommand.Path -Full; exit 0 }

#─────────────────────────────────────────────────────────────────────────────
# UTILITAIRES CLI
#─────────────────────────────────────────────────────────────────────────────
function Write-Header {
    param([string]$text)
    if ($Silent) { return }
    $bar = '=' * 58
    Write-Host "`n$bar" -ForegroundColor DarkCyan
    Write-Host "  $text"  -ForegroundColor Cyan
    Write-Host "$bar"     -ForegroundColor DarkCyan
}

function Write-Section {
    param([string]$text)
    if ($Silent) { return }
    Write-Host "`n -- $text" -ForegroundColor Yellow
}

function Write-Info {
    param([string]$text, [string]$color = 'Gray')
    if ($Silent) { return }
    Write-Host "    $text" -ForegroundColor $color
}

function Get-StatusColor {
    param([string]$status)
    switch ($status) {
        'OK'    { return 'Green' }
        'WARN'  { return 'Yellow' }
        'FAIL'  { return 'Red' }
        default { return 'DarkGray' }
    }
}

#─────────────────────────────────────────────────────────────────────────────
# COLLECTE RESULTATS
#─────────────────────────────────────────────────────────────────────────────
$results = [System.Collections.Generic.List[PSCustomObject]]::new()

function Add-Result {
    param(
        [string]$test,
        [string]$status,
        [string]$valeur,
        [string]$detail = ''
    )
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

#─────────────────────────────────────────────────────────────────────────────
# INTEROP + MOTEUR RAM — tout en C# natif
# - VirtualLock / VirtualUnlock (pin pages physiques)
# - SetProcessWorkingSetSize (elargit le quota de lock du process)
# - SeLockMemoryPrivilege
# - Toutes les boucles fill/check/walking en unsafe C# pour la vitesse
# - Parallel.For pour multi-thread natif (pas de runspaces PS)
# - Progress reporting via volatile int partage
#─────────────────────────────────────────────────────────────────────────────
$RamEngineCode = @'
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;

public static class RamEngine {

    // ── P/Invoke ──────────────────────────────────────────────────────────────
    [DllImport("kernel32.dll", SetLastError=true)]
    static extern bool VirtualLock(IntPtr addr, UIntPtr size);
    [DllImport("kernel32.dll", SetLastError=true)]
    static extern bool VirtualUnlock(IntPtr addr, UIntPtr size);
    [DllImport("kernel32.dll", SetLastError=true)]
    static extern IntPtr GetCurrentProcess();
    [DllImport("kernel32.dll", SetLastError=true)]
    static extern bool SetProcessWorkingSetSize(IntPtr hProcess, UIntPtr min, UIntPtr max);
    [DllImport("advapi32.dll", SetLastError=true)]
    static extern bool OpenProcessToken(IntPtr hProc, uint access, out IntPtr hToken);
    [DllImport("advapi32.dll", SetLastError=true)]
    static extern bool LookupPrivilegeValue(string sys, string name, out LUID luid);
    [DllImport("advapi32.dll", SetLastError=true)]
    static extern bool AdjustTokenPrivileges(IntPtr hToken, bool disable,
        ref TOKEN_PRIVILEGES tp, uint len, IntPtr prev, IntPtr retLen);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
    [DllImport("kernel32.dll")] static extern int GetLastError();

    [StructLayout(LayoutKind.Sequential)] struct LUID { public uint Low; public int High; }
    [StructLayout(LayoutKind.Sequential)] struct LUID_ATTR { public LUID Luid; public uint Attr; }
    [StructLayout(LayoutKind.Sequential)] struct TOKEN_PRIVILEGES { public uint Count; public LUID_ATTR P; }

    const uint TOK_ADJUST = 0x0020, TOK_QUERY = 0x0008, SE_ENABLE = 2;

    // ── Privilege + Working Set ────────────────────────────────────────────────
    public static bool TryEnableLockMemory() {
        IntPtr tok;
        if (!OpenProcessToken(GetCurrentProcess(), TOK_ADJUST | TOK_QUERY, out tok)) return false;
        try {
            LUID luid;
            if (!LookupPrivilegeValue(null, "SeLockMemoryPrivilege", out luid)) return false;
            var tp = new TOKEN_PRIVILEGES { Count=1, P=new LUID_ATTR{Luid=luid, Attr=SE_ENABLE} };
            AdjustTokenPrivileges(tok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
            return GetLastError() == 0;
        } finally { CloseHandle(tok); }
    }

    // Elargit le Working Set du process pour permettre VirtualLock sur de gros buffers
    // min/max en octets
    public static bool ExpandWorkingSet(long minBytes, long maxBytes) {
        return SetProcessWorkingSetSize(GetCurrentProcess(),
            (UIntPtr)(ulong)minBytes, (UIntPtr)(ulong)maxBytes);
    }

    // ── Structures de resultat ────────────────────────────────────────────────
    public class ThreadResult {
        public int  ThreadId;
        public int  Errors;
        public bool Locked;
        public long LockedBytes;
        public long TestedBytes;
        public double ElapsedSec;
    }

    public class RunResult {
        public ThreadResult[] Threads;
        public int  TotalErrors;
        public long TotalLockedMB;
        public long TotalTestedMB;
        public int  BW_MBs;
        public string LockStatus;
    }

    // ── Progress partagé (volatile) ───────────────────────────────────────────
    static volatile int _passesCompleted = 0;
    static volatile int _totalPassesExpected = 0;
    public static int  PassesCompleted   { get { return _passesCompleted; } }
    public static int  TotalPassesExpected { get { return _totalPassesExpected; } }
    public static void ResetProgress(int total) { _passesCompleted = 0; _totalPassesExpected = total; }

    // ── Moteur de test principal ──────────────────────────────────────────────
    // patterns : tableau de byte patterns a tester
    // sizePerThreadBytes : taille du buffer par thread
    // passes : nombre de passes solid
    // numThreads : parallelisme
    // lockPriv : true si SeLockMemoryPrivilege disponible
    public static RunResult RunTest(
        int    numThreads,
        long   sizePerThreadBytes,
        int    passes,
        bool   lockPriv)
    {
        byte[] solidPatterns = new byte[] { 0x00, 0xFF, 0xAA, 0x55, 0xCC, 0x33 };
        // phases par thread : passes solid + floor(passes/2) checkerboard + 16 walking
        int walkPhases = 16; // 8 bits x 2 (1s + 0s)
        int cbPhases   = passes / 2;
        int totalPhases = passes + cbPhases + walkPhases;
        ResetProgress(numThreads * totalPhases);

        var results = new ThreadResult[numThreads];
        var sw = Stopwatch.StartNew();

        Parallel.For(0, numThreads, new ParallelOptions { MaxDegreeOfParallelism = numThreads }, t => {
            var res = new ThreadResult { ThreadId = t };
            results[t] = res;

            long sz = sizePerThreadBytes;
            byte[] buf = new byte[sz];

            // Pin GC + VirtualLock
            var handle = GCHandle.Alloc(buf, GCHandleType.Pinned);
            bool locked = false;
            try {
                if (lockPriv) {
                    IntPtr addr = handle.AddrOfPinnedObject();
                    locked = VirtualLock(addr, (UIntPtr)(ulong)sz);
                    res.Locked = locked;
                    res.LockedBytes = locked ? sz : 0;
                }

                long testedBytes = 0;

                // ── Solid patterns ────────────────────────────────────────────
                for (int pass = 0; pass < passes; pass++) {
                    byte pat = solidPatterns[pass % solidPatterns.Length];

                    // Fill
                    unsafe {
                        fixed (byte* p = buf) {
                            // Remplissage 8 octets a la fois (ulong)
                            ulong val64 = (ulong)pat * 0x0101010101010101UL;
                            long  words = sz / 8;
                            ulong* pw = (ulong*)p;
                            for (long i = 0; i < words; i++) pw[i] = val64;
                            for (long i = words*8; i < sz; i++) p[i] = pat;
                        }
                    }
                    testedBytes += sz;

                    // Verify
                    bool ok = true;
                    unsafe {
                        fixed (byte* p = buf) {
                            ulong val64 = (ulong)pat * 0x0101010101010101UL;
                            long  words = sz / 8;
                            ulong* pw = (ulong*)p;
                            for (long i = 0; i < words; i++) {
                                if (pw[i] != val64) { ok = false; break; }
                            }
                            if (ok) for (long i = words*8; i < sz; i++) {
                                if (p[i] != pat) { ok = false; break; }
                            }
                        }
                    }
                    testedBytes += sz;
                    if (!ok) Interlocked.Increment(ref res.Errors);
                    Interlocked.Increment(ref _passesCompleted);

                    // ── Checkerboard (passes paires) ──────────────────────────
                    if (pass % 2 == 1) {
                        byte alt = (pat == 0xAA) ? (byte)0x55 : (byte)0xAA;
                        unsafe {
                            fixed (byte* p = buf) {
                                // Remplissage checkerboard 2 octets a la fois
                                ushort cb = (ushort)((alt << 8) | pat);
                                long pairs = sz / 2;
                                ushort* pw = (ushort*)p;
                                for (long i = 0; i < pairs; i++) pw[i] = cb;
                                if (sz % 2 != 0) p[sz-1] = alt;
                            }
                        }
                        bool okCb = true;
                        unsafe {
                            fixed (byte* p = buf) {
                                for (long i = 0; i < sz-1; i += 2) {
                                    if (p[i] != alt || p[i+1] != pat) { okCb=false; break; }
                                }
                            }
                        }
                        testedBytes += sz;
                        if (!okCb) Interlocked.Increment(ref res.Errors);
                        Interlocked.Increment(ref _passesCompleted);
                    }
                }

                // ── Walking 1s / 0s ───────────────────────────────────────────
                for (int bit = 0; bit < 8; bit++) {
                    byte p1 = (byte)(1 << bit);
                    byte p0 = (byte)(~p1);

                    foreach (byte wp in new byte[]{p1, p0}) {
                        unsafe {
                            fixed (byte* p = buf) {
                                ulong val64 = (ulong)wp * 0x0101010101010101UL;
                                long words = sz / 8; ulong* pw = (ulong*)p;
                                for (long i = 0; i < words; i++) pw[i] = val64;
                                for (long i = words*8; i < sz; i++) p[i] = wp;
                            }
                        }
                        bool ok = true;
                        unsafe {
                            fixed (byte* p = buf) {
                                ulong val64 = (ulong)wp * 0x0101010101010101UL;
                                long words = sz/8; ulong* pw = (ulong*)p;
                                for (long i = 0; i < words; i++) {
                                    if (pw[i] != val64) { ok=false; break; }
                                }
                                if (ok) for (long i = words*8; i < sz; i++) {
                                    if (p[i] != wp) { ok=false; break; }
                                }
                            }
                        }
                        testedBytes += sz * 2;
                        if (!ok) Interlocked.Increment(ref res.Errors);
                        Interlocked.Increment(ref _passesCompleted);
                    }
                }

                res.TestedBytes = testedBytes;

            } finally {
                if (locked) {
                    try { VirtualUnlock(handle.AddrOfPinnedObject(), (UIntPtr)(ulong)sz); } catch {}
                }
                if (handle.IsAllocated) handle.Free();
                buf = null;
                GC.Collect();
            }
        });

        sw.Stop();

        var run = new RunResult();
        run.Threads = results;
        foreach (var r in results) {
            if (r == null) continue;
            run.TotalErrors   += r.Errors;
            run.TotalLockedMB += r.LockedBytes / (1024*1024);
            run.TotalTestedMB += r.TestedBytes / (1024*1024);
        }
        double sec = sw.Elapsed.TotalSeconds;
        run.BW_MBs = sec > 0 ? (int)(run.TotalTestedMB / sec) : 0;
        bool anyLocked = run.TotalLockedMB > 0;
        run.LockStatus = anyLocked
            ? string.Format("VirtualLock OK — {0} Mo epingles en RAM physique", run.TotalLockedMB)
            : (lockPriv ? "VirtualLock: privilege OK mais refuse par Windows (Hyper-V? quota?)"
                        : "VirtualLock: non disponible (pas admin ou GPO)");
        return run;
    }
}
'@

# Version safe (sans unsafe/fixed/pointeurs) — fallback si /unsafe indisponible
$RamEngineCodeSafe = @'
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;

public static class RamEngine {

    [DllImport("kernel32.dll", SetLastError=true)]
    static extern bool VirtualLock(IntPtr addr, UIntPtr size);
    [DllImport("kernel32.dll", SetLastError=true)]
    static extern bool VirtualUnlock(IntPtr addr, UIntPtr size);
    [DllImport("kernel32.dll", SetLastError=true)]
    static extern IntPtr GetCurrentProcess();
    [DllImport("kernel32.dll", SetLastError=true)]
    static extern bool SetProcessWorkingSetSize(IntPtr hProcess, UIntPtr min, UIntPtr max);
    [DllImport("advapi32.dll", SetLastError=true)]
    static extern bool OpenProcessToken(IntPtr hProc, uint access, out IntPtr hToken);
    [DllImport("advapi32.dll", SetLastError=true)]
    static extern bool LookupPrivilegeValue(string sys, string name, out LUID luid);
    [DllImport("advapi32.dll", SetLastError=true)]
    static extern bool AdjustTokenPrivileges(IntPtr hToken, bool disable,
        ref TOKEN_PRIVILEGES tp, uint len, IntPtr prev, IntPtr retLen);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
    [DllImport("kernel32.dll")] static extern int GetLastError();

    [StructLayout(LayoutKind.Sequential)] struct LUID { public uint Low; public int High; }
    [StructLayout(LayoutKind.Sequential)] struct LUID_ATTR { public LUID Luid; public uint Attr; }
    [StructLayout(LayoutKind.Sequential)] struct TOKEN_PRIVILEGES { public uint Count; public LUID_ATTR P; }

    const uint TOK_ADJUST = 0x0020, TOK_QUERY = 0x0008, SE_ENABLE = 2;

    public static bool TryEnableLockMemory() {
        IntPtr tok;
        if (!OpenProcessToken(GetCurrentProcess(), TOK_ADJUST | TOK_QUERY, out tok)) return false;
        try {
            LUID luid;
            if (!LookupPrivilegeValue(null, "SeLockMemoryPrivilege", out luid)) return false;
            var tp = new TOKEN_PRIVILEGES { Count=1, P=new LUID_ATTR{Luid=luid, Attr=SE_ENABLE} };
            AdjustTokenPrivileges(tok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
            return GetLastError() == 0;
        } finally { CloseHandle(tok); }
    }

    public static bool ExpandWorkingSet(long minBytes, long maxBytes) {
        return SetProcessWorkingSetSize(GetCurrentProcess(),
            (UIntPtr)(ulong)minBytes, (UIntPtr)(ulong)maxBytes);
    }

    public class ThreadResult {
        public int    ThreadId;
        public int    Errors;
        public bool   Locked;
        public long   LockedBytes;
        public long   TestedBytes;
        public double ElapsedSec;
    }

    public class RunResult {
        public ThreadResult[] Threads;
        public int    TotalErrors;
        public long   TotalLockedMB;
        public long   TotalTestedMB;
        public int    BW_MBs;
        public string LockStatus;
    }

    static volatile int _passesCompleted = 0;
    static volatile int _totalPassesExpected = 0;
    public static int  PassesCompleted      { get { return _passesCompleted; } }
    public static int  TotalPassesExpected  { get { return _totalPassesExpected; } }
    public static void ResetProgress(int total) { _passesCompleted = 0; _totalPassesExpected = total; }

    // Remplissage safe via Buffer.BlockCopy (pas de pointeurs)
    static void FillSolid(byte[] buf, byte pat) {
        if (buf.Length == 0) return;
        buf[0] = pat;
        int filled = 1;
        while (filled < buf.Length) {
            int copy = Math.Min(filled, buf.Length - filled);
            Buffer.BlockCopy(buf, 0, buf, filled, copy);
            filled += copy;
        }
    }

    static bool VerifySolid(byte[] buf, byte pat) {
        for (long i = 0; i < buf.Length; i++)
            if (buf[i] != pat) return false;
        return true;
    }

    static void FillCheckerboard(byte[] buf, byte a, byte b) {
        for (long i = 0; i < buf.Length - 1; i += 2) { buf[i] = a; buf[i+1] = b; }
        if (buf.Length % 2 != 0) buf[buf.Length-1] = a;
    }

    static bool VerifyCheckerboard(byte[] buf, byte a, byte b) {
        for (long i = 0; i < buf.Length - 1; i += 2)
            if (buf[i] != a || buf[i+1] != b) return false;
        return true;
    }

    public static RunResult RunTest(int numThreads, long sizePerThreadBytes, int passes, bool lockPriv) {
        byte[] solidPatterns = new byte[] { 0x00, 0xFF, 0xAA, 0x55, 0xCC, 0x33 };
        int walkPhases  = 16;
        int cbPhases    = passes / 2;
        int totalPhases = passes + cbPhases + walkPhases;
        ResetProgress(numThreads * totalPhases);

        var results = new ThreadResult[numThreads];
        var sw = Stopwatch.StartNew();

        Parallel.For(0, numThreads, new ParallelOptions { MaxDegreeOfParallelism = numThreads }, t => {
            var res = new ThreadResult { ThreadId = t };
            results[t] = res;

            long sz  = sizePerThreadBytes;
            byte[] buf = new byte[sz];

            var handle = GCHandle.Alloc(buf, GCHandleType.Pinned);
            bool locked = false;
            try {
                if (lockPriv) {
                    IntPtr addr = handle.AddrOfPinnedObject();
                    locked = VirtualLock(addr, (UIntPtr)(ulong)sz);
                    res.Locked      = locked;
                    res.LockedBytes = locked ? sz : 0;
                }

                long testedBytes = 0;

                // Solid patterns
                for (int pass = 0; pass < passes; pass++) {
                    byte pat = solidPatterns[pass % solidPatterns.Length];
                    FillSolid(buf, pat);   testedBytes += sz;
                    if (!VerifySolid(buf, pat)) Interlocked.Increment(ref res.Errors);
                    testedBytes += sz;
                    Interlocked.Increment(ref _passesCompleted);

                    // Checkerboard (passes paires)
                    if (pass % 2 == 1) {
                        byte alt = (pat == 0xAA) ? (byte)0x55 : (byte)0xAA;
                        FillCheckerboard(buf, alt, pat);
                        if (!VerifyCheckerboard(buf, alt, pat)) Interlocked.Increment(ref res.Errors);
                        testedBytes += sz;
                        Interlocked.Increment(ref _passesCompleted);
                    }
                }

                // Walking 1s / 0s
                for (int bit = 0; bit < 8; bit++) {
                    byte p1 = (byte)(1 << bit);
                    byte p0 = (byte)(~p1);
                    foreach (byte wp in new byte[]{ p1, p0 }) {
                        FillSolid(buf, wp);
                        if (!VerifySolid(buf, wp)) Interlocked.Increment(ref res.Errors);
                        testedBytes += sz * 2;
                        Interlocked.Increment(ref _passesCompleted);
                    }
                }

                res.TestedBytes = testedBytes;

            } finally {
                if (locked) {
                    try { VirtualUnlock(handle.AddrOfPinnedObject(), (UIntPtr)(ulong)sz); } catch {}
                }
                if (handle.IsAllocated) handle.Free();
                buf = null;
                GC.Collect();
            }
        });

        sw.Stop();

        var run = new RunResult { Threads = results };
        foreach (var r in results) {
            if (r == null) continue;
            run.TotalErrors   += r.Errors;
            run.TotalLockedMB += r.LockedBytes / (1024*1024);
            run.TotalTestedMB += r.TestedBytes / (1024*1024);
        }
        double sec = sw.Elapsed.TotalSeconds;
        run.BW_MBs = sec > 0 ? (int)(run.TotalTestedMB / sec) : 0;
        bool anyLocked = run.TotalLockedMB > 0;
        run.LockStatus = anyLocked
            ? string.Format("VirtualLock OK — {0} Mo epingles en RAM physique", run.TotalLockedMB)
            : (lockPriv ? "VirtualLock: privilege OK mais refuse par Windows (Hyper-V? quota?)"
                        : "VirtualLock: non disponible (pas admin ou GPO)");
        return run;
    }
}
'@

#─────────────────────────────────────────────────────────────────────────────
# COMPILATION RAMENGINE — 3 tentatives par ordre de preference
#  1. CSharpCodeProvider direct avec /unsafe /optimize+ (plus portable)
#  2. Add-Type -CompilerParameters (PS 5.1 standard, peut manquer sur certains hosts)
#  3. Fallback code safe (pas de pointeurs, Buffer.BlockCopy)
#─────────────────────────────────────────────────────────────────────────────
$script:RamEngineCompileError = $null
$script:RamEngineSafeMode     = $false

if (-not ([System.Management.Automation.PSTypeName]'RamEngine').Type) {

    # Tentative 1 : CSharpCodeProvider en memoire — charge l'assembly directement
    $compiled = $false
    try {
        $provider = New-Object Microsoft.CSharp.CSharpCodeProvider
        $cp = New-Object System.CodeDom.Compiler.CompilerParameters
        $cp.GenerateInMemory = $true
        $cp.CompilerOptions  = '/unsafe /optimize+'
        $results_csc = $provider.CompileAssemblyFromSource($cp, $RamEngineCode)
        if ($results_csc.Errors.HasErrors) {
            $msgs = @($results_csc.Errors | ForEach-Object { $_.ToString() })
            throw "Erreurs CSC : $($msgs -join ' ; ')"
        }
        # L'assembly est deja en memoire — verifier le type et forcer chargement dans le domaine PS
        $asm = $results_csc.CompiledAssembly
        if (-not $asm.GetType('RamEngine')) { throw "Type RamEngine introuvable dans l'assembly compile" }
        $null = [System.AppDomain]::CurrentDomain.Load($asm.GetName())
        $compiled = $true
    } catch {
        $errTry1 = $_.Exception.Message
    }

    # Tentative 2 : compilation vers DLL temporaire sur disque, puis Add-Type -Path
    # (contourne les restrictions de certains hosts PS sur -CompilerParameters)
    if (-not $compiled -and -not ([System.Management.Automation.PSTypeName]'RamEngine').Type) {
        $tmpCs  = $null
        $tmpDll = $null
        try {
            $tmpCs  = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "RamEngine_$([System.Guid]::NewGuid().ToString('N')).cs")
            $tmpDll = $tmpCs -replace '\.cs$','.dll'
            [System.IO.File]::WriteAllText($tmpCs, $RamEngineCode, [System.Text.Encoding]::UTF8)
            $provider2 = New-Object Microsoft.CSharp.CSharpCodeProvider
            $cp2 = New-Object System.CodeDom.Compiler.CompilerParameters
            $cp2.GenerateInMemory = $false
            $cp2.OutputAssembly   = $tmpDll
            $cp2.CompilerOptions  = '/unsafe /optimize+'
            $r2 = $provider2.CompileAssemblyFromFile($cp2, $tmpCs)
            if ($r2.Errors.HasErrors) {
                $msgs = @($r2.Errors | ForEach-Object { $_.ToString() })
                throw "Erreurs CSC fichier : $($msgs -join ' ; ')"
            }
            Add-Type -Path $tmpDll -ErrorAction Stop
            $compiled = $true
        } catch {
            $errTry2 = $_.Exception.Message
        } finally {
            if ($tmpCs  -and (Test-Path $tmpCs))  { Remove-Item $tmpCs  -Force -ErrorAction SilentlyContinue }
            # Ne pas supprimer $tmpDll : Add-Type en a besoin en memoire tant que la session tourne
        }
    }

    # Tentative 3 : code safe (Buffer.BlockCopy, pas de pointeurs)
    if (-not $compiled -and -not ([System.Management.Automation.PSTypeName]'RamEngine').Type) {
        try {
            Add-Type -TypeDefinition $RamEngineCodeSafe -Language CSharp -ErrorAction Stop
            $compiled = $true
            $script:RamEngineSafeMode = $true
            if (-not $Silent) {
                Write-Warning "RamEngine compile en mode safe (pas de /unsafe) — performances legerement reduites."
            }
        } catch {
            $errTry3 = $_.Exception.Message
            $script:RamEngineCompileError = @(
                "Tentative 1 (CSharpCodeProvider /unsafe) : $errTry1",
                "Tentative 2 (Add-Type /unsafe)           : $errTry2",
                "Tentative 3 (code safe)                  : $errTry3"
            ) -join " | "
            if (-not $Silent) {
                Write-Warning "RamEngine : echec total compilation.`n  $($script:RamEngineCompileError -replace ' \| ', "`n  ")"
            }
        }
    }
}

#─────────────────────────────────────────────────────────────────────────────
# TEST 1 — RAM
#─────────────────────────────────────────────────────────────────────────────
function Invoke-RamTest {
    # Calcul RamSizeFinal / PassesFinal : priorite config JSON > CLI > auto (25% RAM physique)
    $osQ          = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $freePhysQMB  = if ($osQ) { [Math]::Floor($osQ.FreePhysicalMemory / 1KB) } else { $RamSize }
    $totalPhysQMB = if ($osQ) { [Math]::Floor($osQ.TotalVisibleMemorySize / 1KB) } else { 2048 }
    if ($cfg.RamSizeAutoMB -gt 0) {
        $RamSizeFinal = $cfg.RamSizeAutoMB
    } elseif ($PSBoundParameters.ContainsKey('RamSize')) {
        $RamSizeFinal = $RamSize
    } else {
        $RamSizeFinal = [Math]::Max(512, [Math]::Min(4096, [int]($totalPhysQMB * 0.25)))
    }
    $script:RamSizeFinal = $RamSizeFinal
    $PassesFinal = if ($cfg.RamPasses -gt 0) { $cfg.RamPasses } else { $Passes }

    $modeHeader = if ($script:RamEngineSafeMode) { 'safe/Buffer.BlockCopy' } else { 'unsafe/natif' }
    Write-Section "RAM — VirtualLock + C# $modeHeader + patterns complets ($RamSizeFinal Mo req., $PassesFinal passes)"

    if (-not ([System.Management.Automation.PSTypeName]'RamEngine').Type) {
        $cause = if ($script:RamEngineCompileError) { $script:RamEngineCompileError } else { 'Add-Type RamEngine indisponible' }
        Write-Info "! RamEngine non compile — test RAM annule" -color 'Red'
        Write-Info "  Cause : $cause" -color 'DarkYellow'
        Write-Info "  Solutions : lancer en admin, verifier .NET Framework 4.x, ou utiliser PowerShell 7+" -color 'DarkYellow'
        Add-Result 'RAM' 'N/A' 'Compilation echouee' $cause | Out-Null
        return
    }

    # ── Privilege SeLockMemoryPrivilege ───────────────────────────────────────
    $lockPrivOk = $false
    if ($IsAdmin) {
        try { $lockPrivOk = [RamEngine]::TryEnableLockMemory() } catch {}
    }

    if ($lockPrivOk) {
        Write-Info "SeLockMemoryPrivilege : OK" -color 'DarkGray'
    } else {
        Write-Info "SeLockMemoryPrivilege : non disponible (pas admin ou GPO bloque)" -color 'DarkYellow'
    }

       # Calcul usable ($freePhysQMB et $RamSizeFinal deja calcules en tete de fonction)
    $freePhysMB = $freePhysQMB
    $usableMB   = [Math]::Max($cfg.RamPerThreadMin, $freePhysMB - $cfg.RamOsReserveMB)

    $cpuLogical = (Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue |
                   Measure-Object -Property NumberOfLogicalProcessors -Sum).Sum
    # Max RamMaxThreads threads (configurable)
    $numThreads = [Math]::Max(1, [Math]::Min([int]$cpuLogical, $cfg.RamMaxThreads))

    $totalTestMB  = [Math]::Min($RamSizeFinal, $usableMB)
    $perThreadMB  = [Math]::Max($cfg.RamPerThreadMin, [Math]::Floor($totalTestMB / $numThreads))
    $perThreadB   = [long]($perThreadMB * 1MB)
    $testedMB     = $perThreadMB * $numThreads

    Write-Info ("RAM libre : {0} Mo | Test : {1} Mo ({2} threads x {3} Mo)" `
        -f $freePhysMB, $testedMB, $numThreads, $perThreadMB) -color 'Gray'

    if ($totalTestMB -lt $RamSizeFinal) {
        Write-Info ("  ! RAM libre ({0} Mo) < RamSize ({1} Mo) — test reduit" `
            -f $usableMB, $RamSizeFinal) -color 'DarkYellow'
    }

    # ── Elargir le Working Set pour permettre VirtualLock ────────────────────
    if ($lockPrivOk) {
        $minWS = [long]($testedMB + 512) * 1MB
        $maxWS = [long]($testedMB + [Math]::Max(512, $cfg.RamOsReserveMB * 4)) * 1MB
        try {
            $wsOk = [RamEngine]::ExpandWorkingSet($minWS, $maxWS)
            if ($wsOk) {
                Write-Info ("Working Set elargi a {0} Mo min pour VirtualLock" -f ($minWS/1MB)) -color 'DarkGray'
            } else {
                Write-Info "  ! ExpandWorkingSet refuse — VirtualLock peut echouer sur gros buffers" -color 'DarkYellow'
            }
        } catch {
            Write-Info "  ! ExpandWorkingSet exception : $($_.Exception.Message)" -color 'DarkYellow'
        }
    }

    # ── Lancement avec progress ───────────────────────────────────────────────
    $modeLabel = if ($script:RamEngineSafeMode) { 'mode safe — Buffer.BlockCopy' } else { 'mode unsafe — pointeurs natifs' }
    Write-Info "Lancement ($numThreads threads, $modeLabel)..." -color 'DarkGray'

    # Lancer le test en arriere-plan PS pour pouvoir afficher la progression
    $runJob = [System.Management.Automation.PowerShell]::Create()
    $null   = $runJob.AddScript({
        param($nt, $ptb, $passes, $lock)
        return [RamEngine]::RunTest($nt, $ptb, $passes, $lock)
    }).AddArgument($numThreads).AddArgument($perThreadB).AddArgument($PassesFinal).AddArgument($lockPrivOk)

    $rsAsync = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rsAsync.Open()
    $runJob.Runspace = $rsAsync
    $asyncHandle = $runJob.BeginInvoke()

    # Boucle de progression dans le thread principal
    $spinner = @('|','/','-','\')
    $si = 0
    while (-not $asyncHandle.IsCompleted) {
        $done  = [RamEngine]::PassesCompleted
        $total = [RamEngine]::TotalPassesExpected
        $pct   = if ($total -gt 0) { [Math]::Round($done / $total * 100, 0) } else { 0 }
        $spin  = $spinner[$si % 4]; $si++
        if (-not $Silent) {
            Write-Host ("`r    {0} Progression : {1}% ({2}/{3} phases)" `
                -f $spin, $pct, $done, $total) -NoNewline -ForegroundColor DarkGray
        }
        Start-Sleep -Milliseconds 400
    }
    if (-not $Silent) { Write-Host "`r    " -NoNewline }  # efface la ligne spinner

    # Collecte resultat
    $runResult = $null
    try {
        $raw = $runJob.EndInvoke($asyncHandle)
        $runResult = $raw | Select-Object -First 1
    } catch {
        Write-Info "! Erreur RunTest : $($_.Exception.Message)" -color 'Red'
    } finally {
        $runJob.Dispose()
        $rsAsync.Close()
    }

    if (-not $runResult) {
        Add-Result 'RAM' 'N/A' 'Erreur execution' 'RunTest a leve une exception' | Out-Null
        return
    }

    # ── Affichage resultats ───────────────────────────────────────────────────
    $allErrors = $runResult.TotalErrors
    $lockInfo  = $runResult.LockStatus
    $bw        = $runResult.BW_MBs
    $lockedMB  = $runResult.TotalLockedMB

    Write-Info ("Termine : {0} Mo testes | {1} | BW: {2} Mo/s" `
        -f $testedMB, $lockInfo, $bw) -color 'Gray'

    $st = if ($allErrors -eq 0) { 'OK' } else { 'FAIL' }
    $v  = if ($allErrors -eq 0) { '0 erreur' } else { "$allErrors erreur(s)" }
    $engineMode = if ($script:RamEngineSafeMode) { 'C# safe (Buffer.BlockCopy)' } else { 'C# natif+unsafe' }
    $lockDetail = if ($lockedMB -gt 0) { "Epingles: ${lockedMB} Mo" } else { "VirtualLock: non actif" }
    $d  = "$engineMode | ${testedMB} Mo / $numThreads threads / $PassesFinal passes | $lockDetail | BW: ${bw} Mo/s | Solid+Checkerboard+Walking1s/0s | Scan complet 64-bit"

    Write-Info "Resultat : $v" -color (Get-StatusColor $st)
    Write-Host ""
    Write-Host "    ┌─ LIMITE CONNUE ────────────────────────────────────────────────────┐" -ForegroundColor DarkYellow
    Write-Host "    │ VirtualLock epingle les pages en RAM physique (si privilege OK).    │" -ForegroundColor DarkYellow
    Write-Host "    │ Ce test NE peut PAS couvrir : pages reservees OS/drivers/BIOS,     │" -ForegroundColor DarkYellow
    Write-Host "    │ erreurs liees a la temperature ou la tension des modules,           │" -ForegroundColor DarkYellow
    Write-Host "    │ ni adresser directement les rangees physiques (row hammer, etc).    │" -ForegroundColor DarkYellow
    Write-Host "    │ Pour validation hardware complete (overclock, stabilite DIMM) :     │" -ForegroundColor DarkYellow
    Write-Host "    │   -> MemTest86 ou TestMem5  (hors OS, acces physique direct)       │" -ForegroundColor DarkYellow
    Write-Host "    └────────────────────────────────────────────────────────────────────┘" -ForegroundColor DarkYellow
    Write-Host ""

    Add-Result 'RAM' $st $v $d | Out-Null
}

#─────────────────────────────────────────────────────────────────────────────
# TEST 2 — CHARGE CPU / SCHEDULING (mesure scheduling OS + freq effective)
#─────────────────────────────────────────────────────────────────────────────
function Invoke-LatenceTest {
    Write-Section "SCHEDULING CPU — regularite d'execution ($($cfg.LatSamples) echantillons)"

    # ── Explications pedagogiques ────────────────────────────────────────────
    if (-not $Silent) {
        Write-Host "    Ce test NE mesure PAS la latence hardware du CPU." -ForegroundColor DarkYellow
        Write-Host "    Il mesure la REGULARITE avec laquelle Windows execute ce processus." -ForegroundColor DarkYellow
        Write-Host "    Chaque echantillon = $([int]1e5) ops XOR/MUL chronometrees." -ForegroundColor DarkYellow
        Write-Host "    Un jitter eleve = l'OS preempte souvent ce thread (DPC, IRQ, autre process)." -ForegroundColor DarkYellow
        Write-Host ""
    }

    # ── Frequence CPU effective ───────────────────────────────────────────────
    $freqInfo = ''
    $freqPct  = 0
    try {
        $proc = Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1
        $maxMHz = $proc.MaxClockSpeed
        $curMHz = $proc.CurrentClockSpeed
        if ($maxMHz -gt 0 -and $curMHz -gt 0) {
            $freqPct  = [Math]::Round($curMHz / $maxMHz * 100, 0)
            $freqInfo = "Freq CPU: ${curMHz} MHz / ${maxMHz} MHz (${freqPct}%)"
            $freqColor = if ($freqPct -ge 95) { 'Green' } elseif ($freqPct -ge 80) { 'Yellow' } else { 'Red' }
            Write-Info $freqInfo -color $freqColor
            if ($freqPct -lt 80) {
                Write-Info "  ! Freq < 80% — throttling possible (temperature ? plan d alimentation ?)" -color 'DarkYellow'
            }
        }
    }
    catch { Write-Verbose "Freq CPU WMI : $($_.Exception.Message)" }

    # ── Warmup JIT ────────────────────────────────────────────────────────────
    if (-not $Silent) { Write-Host "    Warmup JIT..." -ForegroundColor DarkGray -NoNewline }
    $warmup = [System.Diagnostics.Stopwatch]::StartNew()
    while ($warmup.ElapsedMilliseconds -lt 1000) { $null = [Math]::Sqrt(12345.6789) }
    $warmup.Stop()
    if (-not $Silent) { Write-Host " OK" -ForegroundColor DarkGray }

    # ── Mesure ────────────────────────────────────────────────────────────────
    $samples = $cfg.LatSamples
    $lat     = [double[]]::new($samples)
    $sw      = [System.Diagnostics.Stopwatch]::new()

    if (-not $Silent) { Write-Host "    Mesure en cours..." -ForegroundColor DarkGray -NoNewline }
    for ($i = 0; $i -lt $samples; $i++) {
        $sw.Restart()
        $x = [long]0
        for ($j = 0; $j -lt 100000; $j++) { $x = $x -bxor ([long]$j * 2654435761L) }
        $sw.Stop()
        $lat[$i] = $sw.Elapsed.TotalMilliseconds
        if ($x -eq 0) { Write-Verbose "prevent-opt" }
    }
    if (-not $Silent) { Write-Host " OK ($samples echantillons)" -ForegroundColor DarkGray }

    # ── Calculs statistiques ─────────────────────────────────────────────────
    $avg    = ($lat | Measure-Object -Average).Average
    $sorted = $lat | Sort-Object
    $p50    = $sorted[[Math]::Min([int]($samples * 0.50), $samples - 1)]
    $p95    = $sorted[[Math]::Min([int]($samples * 0.95), $samples - 1)]
    $p99    = $sorted[[Math]::Min([int]($samples * 0.99), $samples - 1)]
    $maxL   = ($lat | Measure-Object -Maximum).Maximum
    $minL   = ($lat | Measure-Object -Minimum).Minimum

    # Jitter = P99 / Min  (1.0 = parfait, >3 = problematique, >8 = critique)
    $jitter = if ($minL -gt 0) { [Math]::Round($p99 / $minL, 1) } else { 0 }

    # Outliers = echantillons > 2x la moyenne
    $outliers = @($lat | Where-Object { $_ -gt ($avg * 2) }).Count
    $outlierPct = [Math]::Round($outliers / $samples * 100, 1)

    # ── Tableau de resultats detaille ────────────────────────────────────────
    if (-not $Silent) {
        Write-Host ""
        Write-Host "    ┌─ RESULTATS SCHEDULING ────────────────────────────────────────────┐" -ForegroundColor Cyan
        Write-Host ("    │  Min  (meilleur cas, sans interruption)  : {0,8:N2} ms              │" -f $minL) -ForegroundColor Cyan
        Write-Host ("    │  Moy  (moyenne des {0} echantillons)    : {1,8:N2} ms              │" -f $samples, $avg) -ForegroundColor Cyan
        Write-Host ("    │  P50  (mediane — 50% des runs)           : {0,8:N2} ms              │" -f $p50) -ForegroundColor Cyan
        Write-Host ("    │  P95  (95% des runs terminent avant)     : {0,8:N2} ms              │" -f $p95) -ForegroundColor Cyan
        Write-Host ("    │  P99  (99% des runs terminent avant)     : {0,8:N2} ms              │" -f $p99) -ForegroundColor Cyan
        Write-Host ("    │  Max  (pire interruption observee)       : {0,8:N2} ms              │" -f $maxL) -ForegroundColor Cyan

        # Couleur jitter selon seuils
        $jitterColor = if ($jitter -lt $cfg.LatJitterWarn) { 'Green' }
                       elseif ($jitter -lt $cfg.LatJitterFail) { 'Yellow' }
                       else { 'Red' }
        $jitterLabel = if ($jitter -lt $cfg.LatJitterWarn)  { 'Excellent' }
                       elseif ($jitter -lt $cfg.LatJitterFail) { 'Moyen — surveiller' }
                       else { 'Eleve — probleme probable' }

        Write-Host ("    │  Jitter P99/Min ({0,-22})    : {1,8:N1} x              │" -f $jitterLabel, $jitter) -ForegroundColor $jitterColor
        Write-Host ("    │  Outliers (>2x moy)                      : {0,5} echant. ({1:N1}%)    │" -f $outliers, $outlierPct) -ForegroundColor Cyan
        Write-Host "    │                                                                    │" -ForegroundColor DarkGray
        Write-Host "    │  LEGENDE :                                                         │" -ForegroundColor DarkGray
        Write-Host "    │  Min  = baseline hardware (sans preemption OS)                     │" -ForegroundColor DarkGray
        Write-Host "    │  P95/P99 = queue de latence, revele les interruptions rares         │" -ForegroundColor DarkGray
        Write-Host "    │  Jitter = P99/Min, ratio de regularite (1x=parfait, >3x=warn)      │" -ForegroundColor DarkGray
        Write-Host "    │  Outliers = pics de latence > 2x moy (DPC, IRQ, autre processus)   │" -ForegroundColor DarkGray
        Write-Host "    └────────────────────────────────────────────────────────────────────┘" -ForegroundColor Cyan

        # ── Mini histogramme ASCII ────────────────────────────────────────────
        Write-Host ""
        Write-Host "    Distribution (histogramme, $samples points) :" -ForegroundColor DarkGray
        $buckets = 10
        $bMin = $minL
        $bMax = [Math]::Max($maxL, $minL + 0.001)
        $bWidth = ($bMax - $bMin) / $buckets
        $counts = @(0) * $buckets
        foreach ($v in $lat) {
            $b = [Math]::Min([int](($v - $bMin) / $bWidth), $buckets - 1)
            $counts[$b]++
        }
        $maxCount = ($counts | Measure-Object -Maximum).Maximum
        for ($b = 0; $b -lt $buckets; $b++) {
            $lo   = $bMin + $b * $bWidth
            $hi   = $lo + $bWidth
            $barLen = if ($counts[$b] -gt 0) { [Math]::Max(1, [int]($counts[$b] / $maxCount * 40)) } else { 0 }
            $bar  = '#' * $barLen
            $pct  = [Math]::Round($counts[$b] / $samples * 100, 0)
            if ($counts[$b] -gt 0) {
                Write-Host ("    {0,6:N2}-{1,6:N2}ms | {2,-40} {3,3}% ({4})" `
                    -f $lo, $hi, $bar, $pct, $counts[$b]) -ForegroundColor DarkGray
            }
        }
        Write-Host ""
    }

    # ── Statuts ───────────────────────────────────────────────────────────────
    $stAvg    = if ($avg -lt $cfg.LatMoyMax) { 'OK' } else { 'WARN' }
    $stJitter = if ($jitter -lt $cfg.LatJitterWarn)  { 'OK' }
                elseif ($jitter -lt $cfg.LatJitterFail) { 'WARN' }
                else { 'FAIL' }

    $detailFull = ("Min={0:N2}ms Avg={1:N2}ms P50={2:N2}ms P95={3:N2}ms P99={4:N2}ms Max={5:N2}ms Jitter={6:N1}x Outliers={7} ({8:N1}%)" `
        -f $minL, $avg, $p50, $p95, $p99, $maxL, $jitter, $outliers, $outlierPct)
    if ($freqInfo) { $detailFull += " | $freqInfo" }

    Add-Result 'CPU Scheduling (moy)'    $stAvg    ("{0:N2} ms" -f $avg)        $detailFull | Out-Null
    Add-Result 'CPU Scheduling (jitter)' $stJitter ("Jitter {0:N1}x" -f $jitter) $detailFull | Out-Null

    $script:LatSamples = $lat
}

#─────────────────────────────────────────────────────────────────────────────
# TEST 3 — WHEA
#─────────────────────────────────────────────────────────────────────────────
function Invoke-WheaTest {
    Write-Section "WHEA — Erreurs materielles"

    if (-not $IsAdmin) {
        Write-Info "! Droits admin requis" -color 'DarkYellow'
        Add-Result 'WHEA total'    'N/A' 'Admin requis' '' | Out-Null
        Add-Result 'WHEA critique' 'N/A' 'Admin requis' '' | Out-Null
        return
    }

    $wheaCount    = 0
    $wheaCritical = 0
    $wheaDetail   = ''

    # 1. psloglist64.exe (optionnel, dans Tools\)
    # Notes techniques :
    #   - $ErrorActionPreference='Stop' global ferait planter sur exit-code non-zero
    #     => on force 'Continue' localement + on capture stderr via 2>&1
    #   - psloglist64 scanne un journal par NOM (pas par provider) => "System"
    #   - On filtre ensuite les lignes contenant "WHEA" dans la sortie texte
    #   - -accepteula enregistre la licence en registre sans popup bloquant
    $psLog = Join-Path $PSScriptRoot 'Tools\psloglist64.exe'
    if (Test-Path $psLog) {
        Write-Info "Scan WHEA via psloglist64.exe..." -color 'DarkGray'
        try {
            # Neutraliser Stop localement : psloglist64 peut retourner exit!=0
            # meme en cas de succes (EULA 1er lancement, journal vide, etc.)
            $savedEAP = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            $logOut = & $psLog -accepteula System -n 200 2>&1
            $ErrorActionPreference = $savedEAP

            # $logOut peut contenir des ErrorRecord (stderr) — on les ignore
            $logLines = $logOut | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] }

            if ($logLines.Count -eq 0) {
                # Peut arriver si EULA vient d etre acceptee (1er lancement)
                # On relance une fois sans -accepteula
                $ErrorActionPreference = 'Continue'
                $logOut   = & $psLog System -n 200 2>&1
                $ErrorActionPreference = $savedEAP
                $logLines = $logOut | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] }
            }

            # Filtrer les lignes WHEA-Logger dans la sortie texte
            $wheaLines    = $logLines | Where-Object { $_ -match 'WHEA' }
            $wheaCount    = $wheaLines.Count
            $wheaCritical = ($wheaLines | Where-Object { $_ -match '(Critical|Error)' }).Count
            $wheaDetail   = if ($wheaLines.Count -gt 0) {
                "psloglist64: $wheaCount evt(s) WHEA"
            } elseif ($logLines.Count -gt 0) {
                "psloglist64 OK — aucun evt WHEA"
            } else {
                # Vide apres 2 tentatives => laisser Get-WinEvent prendre le relais
                ''
            }
            Write-Verbose "psloglist64 : $($logLines.Count) lignes, $wheaCount WHEA"
        }
        catch {
            # Ne jamais bloquer sur psloglist64 — Get-WinEvent prend le relais
            Write-Verbose "psloglist64 exception : $($_.Exception.Message)"
            $wheaDetail = ''
        }
    }

    # 2. Fallback Get-WinEvent si psloglist64 absent, vide ou en echec
    # wheaDetail='' = psloglist64 n a pas fourni de resultat => fallback necessaire
    # wheaDetail='psloglist64 OK' = scan reussi, 0 erreur WHEA => pas de fallback
    $pslogReussi = $wheaDetail -ne '' -and $wheaDetail -notmatch '^$'
    if (-not $pslogReussi) {
        try {
            Write-Info "Fallback WHEA via Get-WinEvent..." -color 'DarkGray'
            $wheaEvents = @(Get-WinEvent -FilterHashtable @{
                LogName      = 'System'
                ProviderName = 'Microsoft-Windows-WHEA-Logger'
                Id           = 17, 18, 19, 20, 41, 4101
            } -MaxEvents 100 -ErrorAction SilentlyContinue)

            $wheaCount    = $wheaEvents.Count
            $wheaCritical = @($wheaEvents | Where-Object { $_.Id -eq 41 -or $_.Level -eq 1 }).Count

            if ($wheaCount -gt 0) {
                $wheaEvents | Select-Object -First 3 | ForEach-Object {
                    $msg = if ($_.Message) {
                        $_.Message.Substring(0, [Math]::Min(80, $_.Message.Length))
                    } else { '(no message)' }
                    $wheaDetail += "[$($_.TimeCreated.ToString('dd/MM HH:mm'))] Id=$($_.Id) $msg; "
                }
            }
        }
        catch {
            Write-Verbose "Get-WinEvent WHEA echec : $($_.Exception.Message)"
            $wheaDetail += " WinEvent fallback echec"
        }
    }

    $stTotal = if ($wheaCount    -eq 0) { 'OK' } else { 'WARN' }
    $stCrit  = if ($wheaCritical -eq 0) { 'OK' } else { 'FAIL' }
    $vTot    = "Total: $wheaCount | Critique: $wheaCritical"

    Write-Info "Resultat WHEA : $vTot" -color (Get-StatusColor $stTotal)
    Add-Result 'WHEA total'    $stTotal $vTot             $wheaDetail       | Out-Null
    Add-Result 'WHEA critique' $stCrit  "$wheaCritical evenement(s)" 'Seuil zero tolerance' | Out-Null
}

#─────────────────────────────────────────────────────────────────────────────
# LANCEMENT OHM SI PRESENT
#─────────────────────────────────────────────────────────────────────────────
function Start-OhmIfNeeded {
    # Verifier si capteurs WMI deja actifs
    try {
        $sensor = Get-CimInstance -Namespace 'root/OpenHardwareMonitor' -ClassName Sensor -ErrorAction Stop |
                  Where-Object { $_.SensorType -eq 'Temperature' -and $_.Identifier -match '/cpu/' } |
                  Select-Object -First 1
        if ($sensor -and $sensor.Value -gt 0) {
            Write-Info "WMI capteurs deja actifs (CPU OK)." -color 'DarkGray'
            return $null
        }
    }
    catch { Write-Verbose "OHM WMI check : $($_.Exception.Message)" }

    # Chercher OHM
    $ohmExe = Join-Path $PSScriptRoot 'OpenHardwareMonitor.exe'
    if (-not (Test-Path $ohmExe)) {
        $ohmExe = Join-Path (Join-Path $PSScriptRoot 'Tools') 'OpenHardwareMonitor.exe'
    }
    if (-not (Test-Path $ohmExe)) {
        Write-Info "! OpenHardwareMonitor.exe introuvable (optionnel)." -color 'DarkGray'
        return $null
    }
    if (-not $IsAdmin) {
        Write-Info "! Admin requis pour OHM." -color 'DarkYellow'
        return $null
    }

    Write-Info "OHM lancement..." -color 'DarkGray'
    # Tuer un eventuel zombie
    Get-Process OpenHardwareMonitor -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500

    $proc = Start-Process -FilePath $ohmExe -ArgumentList "/remote" -WindowStyle Hidden -PassThru -ErrorAction Stop
    Write-Info "OHM demarre (PID $($proc.Id)) — attente capteurs CPU (max $($cfg.OhmTimeoutSec)s)..." -color 'DarkGray'

    $deadline = (Get-Date).AddSeconds($cfg.OhmTimeoutSec)
    $ready    = $false
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        try {
            $sensor = Get-CimInstance -Namespace 'root/OpenHardwareMonitor' -ClassName Sensor -ErrorAction Stop |
                      Where-Object { $_.SensorType -eq 'Temperature' -and $_.Identifier -match '/cpu/' } |
                      Select-Object -First 1
            if ($sensor -and $sensor.Value -gt 0) { $ready = $true; break }
        }
        catch { Write-Verbose "OHM attente : $($_.Exception.Message)" }
    }

    if ($ready) {
        Write-Info "OHM OK — capteurs CPU disponibles." -color 'DarkGray'
    }
    else {
        Write-Info "! OHM lance mais capteurs non prets (timeout)." -color 'DarkYellow'
    }
    return $proc
}

#─────────────────────────────────────────────────────────────────────────────
# TEST 4 — TEMPERATURE avec chaine de fallback complete
#─────────────────────────────────────────────────────────────────────────────
function Invoke-TempTest {
    Write-Section "TEMPERATURE CPU"

    $celsius = $null
    $src     = 'inconnu'

    # 1. MSAcpi_ThermalZoneTemperature (ACPI natif, admin requis)
    if ($IsAdmin) {
        try {
            $raw = (Get-CimInstance -Namespace 'root/WMI' -ClassName MSAcpi_ThermalZoneTemperature `
                    -ErrorAction Stop | Select-Object -First 1).CurrentTemperature
            if ($null -ne $raw -and $raw -gt 0) {
                # FIX : valeur en dixiemes de Kelvin — division par 10, pas soustraction de 2732
                $celsius = [Math]::Round($raw / 10.0 - 273.15, 1)
                $src     = 'ACPI (MSAcpi)'
            }
        }
        catch { Write-Verbose "ACPI ThermalZone : $($_.Exception.Message)" }
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
        }
        catch { Write-Verbose "ThermalZoneInfo : $($_.Exception.Message)" }
    }

    # 3. OpenHardwareMonitor WMI
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
        }
        catch { Write-Verbose "OHM WMI Temp : $($_.Exception.Message)" }
    }

    if ($null -ne $celsius -and $celsius -gt 0 -and $celsius -lt 150) {
        $st = if ($celsius -lt $cfg.TempCPUMax)           { 'OK' }
              elseif ($celsius -lt ($cfg.TempCPUMax + 10)) { 'WARN' }
              else                                          { 'FAIL' }
        Write-Info "CPU : ${celsius}°C  [source: $src]" -color (Get-StatusColor $st)
        Add-Result 'Temperature CPU' $st "${celsius}°C" "Source: $src" | Out-Null
    }
    else {
        Write-Info "! Temperature indisponible. Placez OpenHardwareMonitor.exe dans le dossier Tools\." -color 'DarkYellow'
        Add-Result 'Temperature CPU' 'N/A' 'Source indisponible' 'Lancer OpenHardwareMonitor.exe avec WMI active' | Out-Null
    }
}

#─────────────────────────────────────────────────────────────────────────────
# TEST 5 — DISQUES : espace + SMART + vitesse I/O
#─────────────────────────────────────────────────────────────────────────────
function Invoke-DisqueTest {
    Write-Section "DISQUES — Espace + SMART + vitesse I/O"

    # ── Espace libre ──────────────────────────────────────────────────────────
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

    # ── SMART ─────────────────────────────────────────────────────────────────
    Write-Info "--- SMART ---" -color 'DarkGray'

    if (-not $IsAdmin) {
        Write-Info "! Droits admin requis pour SMART" -color 'DarkYellow'
        Add-Result 'SMART global' 'N/A' 'Admin requis' 'Relancer en administrateur' | Out-Null
    }
    else {
        $toolsDir    = Join-Path $PSScriptRoot 'Tools'
        $smartctlExe = Join-Path $toolsDir 'smartctl.exe'
        $smartDone   = $false

        # 1. smartctl.exe
        if (Test-Path $smartctlExe) {
            try {
                $scanOut  = & $smartctlExe --scan-open 2>&1
                $devLines = @($scanOut | Where-Object { $_ -match '^/dev/' })

                if ($devLines.Count -eq 0) {
                    $scanOut  = & $smartctlExe --scan 2>&1
                    $devLines = @($scanOut | Where-Object { $_ -match '^/dev/' })
                }

                foreach ($devLine in $devLines) {
                    $dev = ($devLine -split '\s')[0]
                    try {
                        $infoJson = & $smartctlExe -a -j $dev 2>&1
                        $info     = $infoJson | ConvertFrom-Json -ErrorAction Stop

                        $dname  = if ($info.model_name)    { $info.model_name }
                                  elseif ($info.device.name) { $info.device.name }
                                  else                        { $dev }

                        $health = if ($info.smart_status.passed -eq $true) { 'PASSED' } else { 'FAILED' }
                        $temp   = if ($info.temperature.current)  { "$($info.temperature.current)°C" } else { '?' }

                        $reallocatedAttr = $info.ata_smart_attributes.table | Where-Object { $_.id -eq 5   } | Select-Object -First 1
                        $pendingAttr     = $info.ata_smart_attributes.table | Where-Object { $_.id -eq 197 } | Select-Object -First 1
                        $udmaAttr        = $info.ata_smart_attributes.table | Where-Object { $_.id -eq 199 } | Select-Object -First 1

                        $reallocated = if ($reallocatedAttr -and $reallocatedAttr.raw.value -gt 0) {
                            $reallocatedAttr.raw.value
                        } elseif ($reallocatedAttr -and $reallocatedAttr.raw.string) {
                            [long]($reallocatedAttr.raw.string -replace '[^\d]', '')
                        } else { 0 }

                        $pending = if ($pendingAttr -and $pendingAttr.raw.value -gt 0) {
                            $pendingAttr.raw.value
                        } elseif ($pendingAttr -and $pendingAttr.raw.string) {
                            [long]($pendingAttr.raw.string -replace '[^\d]', '')
                        } else { 0 }

                        $udmaCrc = if ($udmaAttr -and $udmaAttr.raw.value -gt 0) {
                            $udmaAttr.raw.value
                        } elseif ($udmaAttr -and $udmaAttr.raw.string) {
                            [long]($udmaAttr.raw.string -replace '[^\d]', '')
                        } else { 0 }

                        $nvmeLog = $info.nvme_smart_health_information_log
                        $nvmePE  = if ($nvmeLog) { $nvmeLog.percentage_used } else { $null }
                        $nvmeME  = if ($nvmeLog) { $nvmeLog.media_errors }    else { $null }

                        $pts = @("Sante:$health", "Temp:$temp")
                        if ($reallocated -gt 0) { $pts += "Reallocated:$reallocated" }
                        if ($pending     -gt 0) { $pts += "Pending:$pending" }
                        if ($udmaCrc     -gt 0) { $pts += "UDMA_CRC:$udmaCrc" }
                        if ($nvmeME -and [long]$nvmeME -gt 0) { $pts += "NVMe-MediaErr:$nvmeME" }
                        if ($nvmePE -and [int]$nvmePE  -gt 0) { $pts += "UsedLife:${nvmePE}%" }

                        $detailStr = $pts -join ' | '

                        $st = 'OK'
                        if ($health      -eq 'FAILED')                          { $st = 'FAIL' }
                        if ($pending     -gt 5)                                  { $st = 'FAIL' }
                        elseif ($reallocated -gt 0)                              { if ($st -ne 'FAIL') { $st = 'WARN' } }
                        if ($udmaCrc     -gt 1)                                  { if ($st -ne 'FAIL') { $st = 'WARN' } }
                        if ($nvmeME -and [long]$nvmeME -gt 0)                   { $st = 'FAIL' }
                        if ($nvmePE -and [int]$nvmePE  -gt 90)                  { $st = 'FAIL' }
                        elseif ($nvmePE -and [int]$nvmePE -gt 75 -and $st -ne 'FAIL') { $st = 'WARN' }

                        Write-Info "$dname : [$health] $temp — $detailStr" -color (Get-StatusColor $st)
                        Add-Result "SMART: $dname" $st "[$health] $temp" $detailStr | Out-Null
                        $smartDone = $true
                    }
                    catch {
                        # Fallback texte si JSON echoue
                        try {
                            $txtOut  = & $smartctlExe -a $dev 2>&1
                            $pending = ($txtOut | Select-String "Current_Pending_Sector"  | ForEach-Object { ($_ -split '\s+')[-1] }) -as [int]
                            $realloc = ($txtOut | Select-String "Reallocated_Sector_Ct"   | ForEach-Object { ($_ -split '\s+')[-1] }) -as [int]
                            $uncorr  = ($txtOut | Select-String "Offline_Uncorrectable"   | ForEach-Object { ($_ -split '\s+')[-1] }) -as [int]

                            $pts = @()
                            if ($realloc -gt 0) { $pts += "Reallocated:$realloc" }
                            if ($pending -gt 0) { $pts += "Pending:$pending" }
                            if ($uncorr  -gt 0) { $pts += "Uncorrectable:$uncorr" }
                            $detailStr = if ($pts.Count -gt 0) { $pts -join ' | ' } else { 'OK' }

                            $st = 'OK'
                            if ($pending -gt 5 -or $uncorr -gt 0) { $st = 'FAIL' }
                            elseif ($realloc -gt 0)                { $st = 'WARN' }

                            Write-Info "$dev : $detailStr" -color (Get-StatusColor $st)
                            Add-Result "SMART: $dev" $st $detailStr "smartctl -a fallback" | Out-Null
                            $smartDone = $true
                        }
                        catch {
                            Write-Info "! smartctl echec pour $dev : $($_.Exception.Message)" -color 'DarkYellow'
                            Add-Result "SMART: $dev" 'N/A' 'smartctl echec' $_.Exception.Message | Out-Null
                        }
                    }
                }

                if ($devLines.Count -eq 0) {
                    Write-Info "! smartctl --scan : aucun disque detecte" -color 'DarkYellow'
                }
            }
            catch {
                Write-Info "! smartctl.exe erreur : $($_.Exception.Message)" -color 'DarkYellow'
            }
        }

        # 2. Fallback StorageReliabilityCounter si smartctl absent ou echec total
        if (-not $smartDone) {
            try {
                $physDisks = Get-PhysicalDisk -ErrorAction Stop
                if (-not $physDisks) {
                    Write-Info "! Aucun disque detecte via Get-PhysicalDisk." -color 'DarkYellow'
                    Add-Result 'SMART global' 'N/A' 'Aucun disque detecte' 'Placez smartctl.exe dans Tools\' | Out-Null
                }
                foreach ($disk in $physDisks) {
                    $dname = if ($disk.FriendlyName) { $disk.FriendlyName } else { "Disk $($disk.DeviceId)" }
                    try {
                        $rel    = $disk | Get-StorageReliabilityCounter -ErrorAction Stop
                        $health = $disk.HealthStatus
                        $parts  = @()
                        if ($rel.PSObject.Properties['Wear']                   -and $rel.Wear                   -gt 0) { $parts += "Wear:$($rel.Wear)%" }
                        if ($rel.PSObject.Properties['Temperature']             -and $rel.Temperature             -gt 0) { $parts += "Temp:$($rel.Temperature)°C" }
                        if ($rel.PSObject.Properties['ReadErrorsTotal']         -and $rel.ReadErrorsTotal         -gt 0) { $parts += "RdErr:$($rel.ReadErrorsTotal)" }
                        if ($rel.PSObject.Properties['WriteErrorsUncorrected']  -and $null -ne $rel.WriteErrorsUncorrected -and $rel.WriteErrorsUncorrected -gt 0) { $parts += "WrErr:$($rel.WriteErrorsUncorrected)" }
                        if ($rel.PSObject.Properties['MediaErrors']             -and $null -ne $rel.MediaErrors   -and $rel.MediaErrors -gt 0) { $parts += "MediaErr:$($rel.MediaErrors)" }
                        $detailStr = if ($parts.Count -gt 0) { $parts -join ' | ' } else { 'Pas d anomalie detectee' }

                        $st = 'OK'
                        if ($health -ne 'Healthy') { $st = 'WARN' }
                        if ($rel.PSObject.Properties['WriteErrorsUncorrected'] -and $null -ne $rel.WriteErrorsUncorrected -and $rel.WriteErrorsUncorrected -gt 0) { $st = 'FAIL' }
                        if ($rel.PSObject.Properties['MediaErrors']            -and $null -ne $rel.MediaErrors   -and $rel.MediaErrors -gt 5) { $st = 'FAIL' }
                        if ($rel.PSObject.Properties['Wear'] -and $rel.Wear -gt 90)                             { $st = 'FAIL' }
                        elseif ($rel.PSObject.Properties['Wear'] -and $rel.Wear -gt 75 -and $st -ne 'FAIL')    { $st = 'WARN' }

                        $valStr = "[$health]"
                        if ($rel.PSObject.Properties['Wear'] -and $rel.Wear -gt 0) { $valStr += " Wear=$($rel.Wear)%" }

                        Write-Info "$dname : $valStr — $detailStr" -color (Get-StatusColor $st)
                        Add-Result "SMART: $dname" $st $valStr $detailStr | Out-Null
                    }
                    catch {
                        Write-Info "! SMART indispo pour $dname : $($_.Exception.Message)" -color 'DarkYellow'
                        Add-Result "SMART: $dname" 'N/A' 'Non supporte' "Placez smartctl.exe dans Tools\ pour eviter cette erreur" | Out-Null
                    }
                }
            }
            catch {
                $errMsg = $_.Exception.Message
                Write-Info "! Get-PhysicalDisk indisponible : $errMsg" -color 'DarkYellow'
                Add-Result 'SMART global' 'N/A' 'Indisponible' $errMsg | Out-Null
            }
        }
    }

    # ── Vitesse I/O ───────────────────────────────────────────────────────────
    Write-Info "--- Vitesse I/O ---" -color 'DarkGray'

    $ddExe    = Join-Path (Join-Path $PSScriptRoot 'Tools') 'dd.exe'
    $testFile = Join-Path $env:TEMP "occt81_disktest_$([System.Diagnostics.Process]::GetCurrentProcess().Id).dat"
    $testSizeBytes = [long]($cfg.DiskTestSizeMB) * 1MB
    $count    = $cfg.DiskTestSizeMB

    if (Test-Path $ddExe) {
        # ── dd.exe present ────────────────────────────────────────────────────
        try {
            Write-Info "Ecriture sequentielle (dd.exe, $($cfg.DiskTestSizeMB) Mo)..." -color 'DarkGray'
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $ddWrite = & $ddExe if=/dev/zero of=$testFile bs=1M count=$count 2>&1
            Start-Sleep -Milliseconds 200
            $sw.Stop()

            $ddOut     = $ddWrite -join ' '
            $writeMbps = $null
            if ($ddOut -match '(\d[\d,\.]+)\s*(bytes|octets)/sec') {
                $writeMbps = [Math]::Round([double]($Matches[1] -replace ',', '') / 1MB, 0)
            }
            elseif ($sw.Elapsed.TotalSeconds -gt 0 -and (Test-Path $testFile)) {
                $sz = (Get-Item $testFile).Length
                $writeMbps = [Math]::Round($sz / 1MB / $sw.Elapsed.TotalSeconds, 0)
            }

            if ($null -ne $writeMbps -and $writeMbps -gt 0) {
                $st = if ($writeMbps -gt $cfg.DiskWriteMin) { 'OK' } elseif ($writeMbps -gt 30) { 'WARN' } else { 'FAIL' }
                Write-Info "Ecriture : ${writeMbps} Mo/s" -color (Get-StatusColor $st)
                Add-Result 'Disque — Ecriture (dd)' $st "${writeMbps} Mo/s" "dd.exe — $($cfg.DiskTestSizeMB) Mo dans TEMP" | Out-Null
            }
            else {
                Write-Info "! dd.exe ecriture : debit illisible" -color 'DarkYellow'
                Add-Result 'Disque — Ecriture (dd)' 'N/A' 'Debit illisible' $ddOut | Out-Null
            }
        }
        catch {
            Write-Info "! dd.exe ecriture echouee : $($_.Exception.Message)" -color 'DarkYellow'
            Add-Result 'Disque — Ecriture (dd)' 'N/A' 'Erreur dd' $_.Exception.Message | Out-Null
        }

        if (Test-Path $testFile) {
            try {
                Write-Info "Lecture sequentielle (dd.exe, $($cfg.DiskTestSizeMB) Mo)..." -color 'DarkGray'
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                $ddRead = & $ddExe if=$testFile of=/dev/null bs=1M 2>&1
                Start-Sleep -Milliseconds 200
                $sw.Stop()

                $ddOut    = $ddRead -join ' '
                $readMbps = $null
                if ($ddOut -match '(\d[\d,\.]+)\s*(bytes|octets)/sec') {
                    $readMbps = [Math]::Round([double]($Matches[1] -replace ',', '') / 1MB, 0)
                }
                elseif ($sw.Elapsed.TotalSeconds -gt 0) {
                    $readMbps = [Math]::Round($testSizeBytes / 1MB / $sw.Elapsed.TotalSeconds, 0)
                }

                if ($null -ne $readMbps -and $readMbps -gt 0) {
                    $st = if ($readMbps -gt $cfg.DiskReadMin) { 'OK' } elseif ($readMbps -gt 50) { 'WARN' } else { 'FAIL' }
                    Write-Info "Lecture : ${readMbps} Mo/s" -color (Get-StatusColor $st)
                    Add-Result 'Disque — Lecture (dd)' $st "${readMbps} Mo/s" "dd.exe — $($cfg.DiskTestSizeMB) Mo dans TEMP" | Out-Null
                }
                else {
                    Write-Info "! dd.exe lecture : debit illisible" -color 'DarkYellow'
                    Add-Result 'Disque — Lecture (dd)' 'N/A' 'Debit illisible' $ddOut | Out-Null
                }
            }
            catch {
                Write-Info "! dd.exe lecture echouee : $($_.Exception.Message)" -color 'DarkYellow'
                Add-Result 'Disque — Lecture (dd)' 'N/A' 'Erreur dd' $_.Exception.Message | Out-Null
            }
            Remove-Item -LiteralPath $testFile -Force -ErrorAction SilentlyContinue
        }
    }
    else {
        # ── Fallback PS (dd.exe absent) ───────────────────────────────────────
        Write-Info "(dd.exe absent — fallback ecriture PS, 50 Mo)" -color 'DarkGray'
        try {
            $sizeMB = 50
            $buf    = [byte[]]::new([int]($sizeMB * 1MB))
            [System.Random]::new().NextBytes($buf)

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            [System.IO.File]::WriteAllBytes($testFile, $buf)
            $sw.Stop()

            $mbps = [Math]::Round($sizeMB / $sw.Elapsed.TotalSeconds, 0)
            $st   = if ($mbps -gt $cfg.DiskWriteMin) { 'OK' } elseif ($mbps -gt 30) { 'WARN' } else { 'FAIL' }
            Write-Info "Ecriture : ${mbps} Mo/s" -color (Get-StatusColor $st)
            Add-Result 'Disque — Ecriture' $st "${mbps} Mo/s" "PS fallback — dd.exe recommande" | Out-Null
        }
        catch {
            Write-Info "! Ecriture echouee : $($_.Exception.Message)" -color 'DarkYellow'
            Add-Result 'Disque — Ecriture' 'N/A' 'Erreur I/O' $_.Exception.Message | Out-Null
        }
        finally {
            Remove-Item -LiteralPath $testFile -Force -ErrorAction SilentlyContinue
        }
    }
}

#─────────────────────────────────────────────────────────────────────────────
# TEST 6 — GPU : nvidia-smi → OHM → WMI fallback
#─────────────────────────────────────────────────────────────────────────────
function Invoke-GpuTest {
    Write-Section "GPU"
    $found = $false

    # 1. nvidia-smi
    $smiPath = @(
        "$env:ProgramFiles\NVIDIA Corporation\NVSMI\nvidia-smi.exe",
        "$env:SystemRoot\System32\nvidia-smi.exe"
    ) | Where-Object { Test-Path $_ -ErrorAction SilentlyContinue } | Select-Object -First 1

    if (-not $smiPath) {
        try { $smiPath = (Get-Command 'nvidia-smi.exe' -ErrorAction Stop).Source }
        catch { Write-Verbose "nvidia-smi non trouve dans PATH" }
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
                $tempG   = if ($p[1] -match '^\d') { [int]$p[1] }    else { $null }
                $utilG   = if ($p[2] -match '^\d') { [int]$p[2] }    else { $null }
                $mU      = if ($p[3] -match '^\d') { [int]$p[3] }    else { $null }
                $mT      = if ($p[4] -match '^\d') { [int]$p[4] }    else { $null }
                $pw      = if ($p.Count -gt 5 -and $p[5] -match '^\d') { [double]$p[5] } else { $null }

                $st = 'OK'
                if ($tempG -and $tempG -gt $cfg.TempGPUMax)              { $st = 'FAIL' }
                elseif ($tempG -and $tempG -gt ($cfg.TempGPUMax - 10))   { $st = 'WARN' }

                $val = if ($tempG) { "${tempG}°C" } else { 'N/A' }
                $det = ''
                if ($null -ne $utilG) { $det += "Load:${utilG}%  " }
                if ($mU -and $mT)     { $det += "VRAM:${mU}/${mT}MiB  " }
                if ($null -ne $pw)    { $det += "Pwr:$([Math]::Round($pw,1))W  " }
                $det += '| nvidia-smi'

                Write-Info "$gpuName : $val | $det" -color (Get-StatusColor $st)
                Add-Result "GPU: $gpuName" $st $val $det | Out-Null
                $found = $true
            }
        }
        catch { Write-Verbose "nvidia-smi echec : $($_.Exception.Message)" }
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
        }
        catch { Write-Verbose "OHM GPU : $($_.Exception.Message)" }
    }

    # 3. Fallback Win32_VideoController (info driver, pas de temperature)
    if (-not $found) {
        try {
            $gpus = Get-CimInstance -ClassName Win32_VideoController -ErrorAction Stop
            foreach ($g in $gpus) {
                $vram = if ($g.AdapterRAM -gt 0) { [Math]::Round($g.AdapterRAM / 1MB, 0) } else { '?' }
                $st   = if ($g.Status -eq 'OK') { 'OK' } else { 'WARN' }
                Write-Info "$($g.Name) — VRAM ${vram} Mo [pas de temperature — installez nvidia-smi ou OHM]" -color (Get-StatusColor $st)
                Add-Result "GPU: $($g.Name)" $st "VRAM ${vram} Mo" "Driver: $($g.DriverVersion) | Info driver seul" | Out-Null
            }
        }
        catch {
            Write-Verbose "Win32_VideoController : $($_.Exception.Message)"
            Add-Result 'GPU' 'N/A' 'Indisponible' $_.Exception.Message | Out-Null
        }
    }
}

#─────────────────────────────────────────────────────────────────────────────
# TEST 7 — UPTIME & SYSTEME
#─────────────────────────────────────────────────────────────────────────────
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
    $stRAM    = if ($ramPct -lt $cfg.RamPctWarn)   { 'OK' }
                elseif ($ramPct -lt $cfg.RamPctFail) { 'WARN' }
                else                                  { 'FAIL' }

    Add-Result 'Uptime'       $stUptime "${days}j $($uptime.Hours)h $($uptime.Minutes)m" "OS: $($os.Caption)" | Out-Null
    Add-Result 'RAM utilisee' $stRAM    "${ramPct}%"  "Physique: ${ramGB} Go" | Out-Null
    Add-Result 'CPU info'     'OK'      $cpu.Name     "Cores: $($cpu.NumberOfCores) / Logiques: $($cpu.NumberOfLogicalProcessors)" | Out-Null
}

#─────────────────────────────────────────────────────────────────────────────
# MOTEUR PRINCIPAL
#─────────────────────────────────────────────────────────────────────────────
function Invoke-AllTests {
    $results.Clear()
    if (Set-ShouldRun 'RAM')     { Invoke-RamTest     }
    if (Set-ShouldRun 'Latence') { Invoke-LatenceTest }
    if (Set-ShouldRun 'WHEA')    { Invoke-WheaTest    }
    if (Set-ShouldRun 'Temp')    { Invoke-TempTest    }
    if (Set-ShouldRun 'Disque')  { Invoke-DisqueTest  }
    if (Set-ShouldRun 'GPU')     { Invoke-GpuTest     }
    if (Set-ShouldRun 'Uptime')  { Invoke-UptimeTest  }
}

#─────────────────────────────────────────────────────────────────────────────
# HISTORIQUE JSON
#─────────────────────────────────────────────────────────────────────────────
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
    }
    catch {
        if (-not $Silent) { Write-Info "! Historique non sauvegarde : $($_.Exception.Message)" -color 'DarkYellow' }
        return $null
    }
}

function Compare-History {
    param([string]$jsonPath)

    if (-not (Test-Path $jsonPath)) {
        Write-Info "! Fichier introuvable : $jsonPath" -color 'Red'
        return
    }
    try {
        $prev     = Get-Content $jsonPath -Raw | ConvertFrom-Json
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
                    'OK'    { '<-- AMELIORE' }
                    'FAIL'  { '<-- DEGRADE !!' }
                    default { '<-- change' }
                }
                $col = switch ($cur.Status) {
                    'OK'    { 'Green' }
                    'FAIL'  { 'Red' }
                    default { 'Yellow' }
                }
                Write-Info "$($cur.Test) : $($p.Status) -> $($cur.Status)  $arrow" -color $col
            }
        }
        if ($changes -eq 0) { Write-Info 'Aucun changement de statut detecte.' -color 'Green' }
    }
    catch {
        Write-Info "! Erreur comparaison : $($_.Exception.Message)" -color 'Red'
    }
}

#─────────────────────────────────────────────────────────────────────────────
# UPLOADS
#─────────────────────────────────────────────────────────────────────────────
function Invoke-DPasteUpload {
    param([string]$Text, [string]$Title = "occt81 Report")
    try {
        $body = @{
            content     = $Text
            title       = $Title
            syntax      = 'powershell'
            expiry_days = 30
        }
        $res = Invoke-RestMethod -Uri "https://dpaste.com/api/v2/" -Method Post -Body $body -ErrorAction Stop
        Write-Host "DPaste : $($res.ToString().Trim())" -ForegroundColor Cyan
        return $res.ToString().Trim()
    }
    catch {
        Write-Warning "DPaste upload echoue : $($_.Exception.Message)"
        return $null
    }
}

function Invoke-GoFileUpload {
    param([string]$FilePath)
    try {
        if (-not (Test-Path $FilePath)) { throw "Fichier introuvable : $FilePath" }

        Write-Host "Upload GoFile..." -ForegroundColor Cyan -NoNewline

        $resp = curl.exe -s -F "file=@$FilePath" "https://upload.gofile.io/uploadfile" | ConvertFrom-Json

        if ($resp.status -eq "ok" -and $resp.data) {
            $dl = if ($resp.data.downloadPage) { $resp.data.downloadPage }
                  elseif ($resp.data.code)     { "https://gofile.io/d/$($resp.data.code)" }
                  elseif ($resp.data.fileId)   { "https://gofile.io/d/$($resp.data.fileId)" }
                  else { throw "Lien introuvable dans la reponse" }
            Write-Host " OK" -ForegroundColor Green
            Write-Host "   LIEN GOFILE : $dl" -ForegroundColor Yellow
            return $dl
        }
        else { throw "Upload echoue : $($resp.status)" }
    }
    catch {
        Write-Warning "GoFile upload echoue : $($_.Exception.Message)"
        return $null
    }
}

#─────────────────────────────────────────────────────────────────────────────
# CONCLUSION PARTAGEE (unique source de verite — CLI + export)
#─────────────────────────────────────────────────────────────────────────────
function Get-DiagConclusion {
    $critical  = @($results | Where-Object { $_.Status -eq 'FAIL' }).Count
    $warns     = @($results | Where-Object { $_.Status -eq 'WARN' }).Count
    $failItems = @($results | Where-Object { $_.Status -eq 'FAIL' })

    $lines = @()
    if ($critical -gt 0) {
        $smartFail = $failItems | Where-Object { $_.Test -match 'SMART' }
        if ($smartFail) {
            $lines += " !! DISQUE EN DANGER — sauvegarde URGENTE recommandee"
            foreach ($f in $smartFail) {
                $dev    = ($f.Test -replace 'SMART:\s*', '').Trim()
                $detail = "$($f.Valeur) $($f.Detail)"
                $lines += "   Disque : $dev"
                $lines += "   Details: $detail"
                if ($detail -match 'Pending:\s*(\d+)') {
                    $p = [int]$Matches[1]
                    if ($p -gt 100) { $lines += "   Cause: secteurs instables critiques ($p)" }
                    elseif ($p -gt 0) { $lines += "   Cause: secteurs instables ($p)" }
                }
                if ($detail -match 'Uncorrectable:\s*(\d+)') {
                    $u = [int]$Matches[1]
                    if ($u -gt 0) { $lines += "   Cause: erreurs non recuperables ($u)" }
                }
            }
        }
        elseif ($failItems | Where-Object { $_.Test -match '^RAM' }) {
            $lines += " !! ERREURS RAM detectees — verifier la memoire"
            foreach ($f in ($failItems | Where-Object { $_.Test -match '^RAM' })) {
                $lines += "   Details: $($f.Valeur)"
            }
        }
        elseif ($failItems | Where-Object { $_.Test -match 'WHEA' }) {
            $lines += " !! ERREURS MATERIELLES CPU/PCIe (WHEA)"
            foreach ($f in ($failItems | Where-Object { $_.Test -match 'WHEA' })) {
                $lines += "   Details: $($f.Valeur)"
            }
        }
        else {
            $lines += " !! PROBLEME MATERIEL detecte"
            foreach ($f in $failItems) { $lines += "   $($f.Test.TrimEnd()) -> $($f.Valeur.TrimEnd())" }
        }
    }
    elseif ($warns -gt 0) {
        $lines += " !! AVERTISSEMENTS — surveillance recommandee"
    }
    else {
        $lines += " Systeme stable — aucun probleme detecte"
    }

    return @{
        Lines    = $lines
        Critical = $critical
        Warns    = $warns
    }
}

#─────────────────────────────────────────────────────────────────────────────
# RESUME CLI
#─────────────────────────────────────────────────────────────────────────────
function Write-Summary {
    Write-Header 'RESUME DIAGNOSTIQUE'

    # ── Tableau principal ────────────────────────────────────────────────────
    Write-Host ""
    Write-Host ("  {0,-32} {1,-8} {2,-22} {3}" -f 'TEST', 'STATUT', 'VALEUR', 'DETAIL') -ForegroundColor DarkGray
    Write-Host ("  {0}" -f ('-' * 90)) -ForegroundColor DarkGray

    foreach ($r in $results) {
        $color  = Get-StatusColor $r.Status
        $icon   = switch ($r.Status) { 'OK'{'  '} 'WARN'{'!!'} 'FAIL'{'XX'} default{'  '} }
        $name   = $r.Test.TrimEnd().PadRight(32)
        $status = ("[{0}]" -f $r.Status.PadRight(4)).PadRight(8)
        $val    = $r.Valeur.TrimEnd().PadRight(22)
        $detail = $r.Detail.TrimEnd()

        # Tronquer le detail si trop long (le résumé complet est dans l'export)
        if ($detail.Length -gt 80) { $detail = $detail.Substring(0, 77) + '...' }

        Write-Host ("  $icon $name $status $val $detail") -ForegroundColor $color
    }

    Write-Host ("  {0}" -f ('-' * 90)) -ForegroundColor DarkGray

    # ── Legende statuts ──────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "  LEGENDE STATUTS :" -ForegroundColor DarkGray
    Write-Host "    [OK  ]  Tout va bien, dans les seuils normaux"                          -ForegroundColor Green
    Write-Host "    [WARN]  Valeur hors seuil — a surveiller mais pas critique"             -ForegroundColor Yellow
    Write-Host "    [FAIL]  Probleme detecte — intervention recommandee"                    -ForegroundColor Red
    Write-Host "    [N/A ]  Test impossible (droits insuffisants ou outil manquant)"        -ForegroundColor DarkGray

    # ── Legende par test ─────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "  LEGENDE PAR TEST :" -ForegroundColor DarkGray
    $testLegend = @(
        @{ Test='RAM';                    Desc='Erreurs de bits detectees dans la RAM physique (patterns fill+verify+walking)' }
        @{ Test='CPU Scheduling (moy)';   Desc='Temps moyen par iteration — reflete la frequence CPU effective' }
        @{ Test='CPU Scheduling (jitter)';Desc='Jitter = P99/Min — irregularite du scheduling Windows (1x=parfait, >3x=warn, >8x=fail)' }
        @{ Test='WHEA total';             Desc='Erreurs hardware signalees par le firmware/CPU (0 = normal)' }
        @{ Test='WHEA critique';          Desc='Sous-ensemble WHEA critique — zero tolerance (crash potentiel)' }
        @{ Test='Temp CPU';               Desc='Temperature CPU en charge — au-dela de 85C = risque de throttling' }
        @{ Test='Temp GPU';               Desc='Temperature GPU — au-dela de 90C = risque de throttling' }
        @{ Test='Disque';                 Desc='SMART + vitesse I/O sequentielle (lecture/ecriture en Mo/s)' }
        @{ Test='GPU';                    Desc='Presence et charge GPU detectees via nvidia-smi / OHM / WMI' }
        @{ Test='Uptime';                 Desc='Temps depuis le dernier redemarrage (>30j = WARN)' }
    )
    foreach ($tl in $testLegend) {
        Write-Host ("    {0,-30} : {1}" -f $tl.Test, $tl.Desc) -ForegroundColor DarkGray
    }

    # ── Conclusion ───────────────────────────────────────────────────────────
    $bar  = '=' * 58
    $conc = Get-DiagConclusion

    Write-Host ""
    Write-Host "$bar" -ForegroundColor DarkCyan
    foreach ($line in $conc.Lines) {
        $col = if ($conc.Critical -gt 0) { 'Red' } elseif ($conc.Warns -gt 0) { 'Yellow' } else { 'Green' }
        Write-Host $line -ForegroundColor $col
    }
    Write-Host "$bar`n" -ForegroundColor DarkCyan
}

#─────────────────────────────────────────────────────────────────────────────
# EXPORT (TXT / CSV / HTML / JSON)
#─────────────────────────────────────────────────────────────────────────────
function Export-Report {
    # Preparer le fichier upload TXT si necessaire
    $uploadFile   = $null
    $isTempUpload = $false

    if ($UploadDPaste -or $UploadGoFile) {
        if ($Export -ne '' -and [System.IO.Path]::GetExtension($Export).ToLower() -eq '.txt') {
            $uploadFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Export)
        }
        if (-not $uploadFile) {
            $uploadFile   = "$env:TEMP\occt81_upload_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            $isTempUpload = $true
        }
    }

    if ($Export -eq '' -and -not ($UploadDPaste -or $UploadGoFile)) { return }

    $resolvedPath = $null
    $ext          = '.txt'

    if ($Export -ne '') {
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Export)
        $ext          = [System.IO.Path]::GetExtension($resolvedPath).ToLower()
    }
    elseif ($isTempUpload) {
        $resolvedPath = $uploadFile
    }

    # Infos systeme pour les rapports
    $osH    = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cpuH   = Get-CimInstance Win32_Processor       -ErrorAction SilentlyContinue | Select-Object -First 1
    $upH    = if ($osH) { (Get-Date) - $osH.LastBootUpTime } else { $null }
    $dispOS  = if ($osH)  { $osH.Caption } else { 'Inconnu' }
    $dispCPU = if ($cpuH) { $cpuH.Name   } else { 'Inconnu' }
    $dispUp  = if ($upH)  { "$([Math]::Floor($upH.TotalDays))j $($upH.Hours)h $($upH.Minutes)m" } else { 'N/A' }

    $conc        = Get-DiagConclusion
    $contentTxt  = ''

    function Build-TextContent {
        $header = @(
            "occt81 v4.2 — $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')",
            "Machine : $env:COMPUTERNAME",
            "Systeme : $dispOS",
            "CPU     : $dispCPU",
            "Uptime  : $dispUp"
        )
        $bar   = '=' * 90
        $sep   = '-' * 90
        $lines = @(
            $bar,
            "  RESUME DIAGNOSTIQUE",
            $bar,
            ("  {0,-30} {1,-8} {2,-22} {3}" -f 'TEST','STATUT','VALEUR','DETAIL'),
            $sep
        )
        foreach ($r in $results) {
            $icon = switch ($r.Status) { 'OK'{'  '} 'WARN'{'!!'} 'FAIL'{'XX'} default{'  '} }
            $n = $r.Test.TrimEnd().PadRight(30)
            $s = ("[{0}]" -f $r.Status.PadRight(4)).PadRight(8)
            $v = $r.Valeur.TrimEnd().PadRight(22)
            $lines += "  $icon $n $s $v $($r.Detail.TrimEnd())"
        }
        $lines += $sep
        $lines += $conc.Lines
        $lines += $bar
        $lines += ""
        $lines += "  LEGENDE STATUTS :"
        $lines += "    [OK  ]  Dans les seuils normaux"
        $lines += "    [WARN]  A surveiller — pas critique"
        $lines += "    [FAIL]  Probleme detecte — intervention recommandee"
        $lines += "    [N/A ]  Test impossible (droits insuffisants ou outil manquant)"
        $lines += ""
        $lines += "  LEGENDE TESTS :"
        $lines += "    RAM                       : Erreurs de bits RAM (fill+verify+walking). 0 = normal."
        $lines += "    CPU Scheduling (moy)      : Temps moyen/iteration = frequence CPU effective (pas latence hardware)"
        $lines += "    CPU Scheduling (jitter)   : Jitter = P99/Min. 1x=parfait, <3x=OK, 3-8x=WARN, >8x=FAIL"
        $lines += "    WHEA total                : Erreurs hardware CPU/firmware. >0 = investigation recommandee"
        $lines += "    WHEA critique             : Severite critique WHEA. Zero tolerance."
        $lines += "    Temp CPU                  : Temperature CPU sous charge. Seuil WARN=85C"
        $lines += "    Temp GPU                  : Temperature GPU. Seuil WARN=90C"
        $lines += "    Disque                    : SMART + vitesse sequentielle I/O (Mo/s)"
        $lines += "    GPU                       : Presence/charge GPU (nvidia-smi / OHM / WMI)"
        $lines += "    Uptime                    : Duree depuis dernier redemarrage. >30j = WARN"
        $lines += $bar
        return ($header + $lines) -join "`n"
    }

    if ($ext -eq '.csv') {
        # CSV enrichi : metadata en tete + donnees + legende en pied
        # UTF-8 BOM pour compatibilite Excel (evite les problemes d accents)
        $csvLines = @()

        # En-tete metadata (lignes commentees, ignorees par les parseurs standard)
        $csvLines += "# occt81 v4.2 — Rapport de diagnostic"
        $csvLines += "# Date,$( Get-Date -Format 'dd/MM/yyyy HH:mm:ss')"
        $csvLines += "# Machine,$env:COMPUTERNAME"
        $csvLines += "# Systeme,$dispOS"
        $csvLines += "# CPU,$dispCPU"
        $csvLines += "# Uptime,$dispUp"
        $csvLines += "# Conclusion,$($conc.Lines -join ' | ')"
        $csvLines += ""

        # Donnees principales
        $csvLines += '"HEURE","TEST","STATUT","VALEUR","DETAIL"'
        foreach ($r in $results) {
            $h = $r.Heure  -replace '"','""'
            $t = $r.Test.TrimEnd()   -replace '"','""'
            $s = $r.Status -replace '"','""'
            $v = $r.Valeur.TrimEnd() -replace '"','""'
            $d = $r.Detail.TrimEnd() -replace '"','""'
            $csvLines += "`"$h`",`"$t`",`"$s`",`"$v`",`"$d`""
        }
        $csvLines += ""

        # Legende en pied
        $csvLines += "# --- LEGENDE STATUTS ---"
        $csvLines += "# OK,Dans les seuils normaux"
        $csvLines += "# WARN,A surveiller — pas critique"
        $csvLines += "# FAIL,Probleme detecte — intervention recommandee"
        $csvLines += "# N/A,Test impossible (droits insuffisants ou outil manquant)"
        $csvLines += ""
        $csvLines += "# --- LEGENDE TESTS ---"
        $csvLines += "# RAM,Erreurs de bits RAM (fill+verify+walking). 0 = normal."
        $csvLines += "# CPU Scheduling (moy),Temps moyen/iteration — frequence CPU effective (pas latence hardware)"
        $csvLines += "# CPU Scheduling (jitter),Jitter = P99/Min. 1x=parfait | <3x=OK | 3-8x=WARN | >8x=FAIL"
        $csvLines += "# WHEA total,Erreurs hardware CPU/firmware. >0 = investigation recommandee"
        $csvLines += "# WHEA critique,Severite critique WHEA. Zero tolerance."
        $csvLines += "# Temp CPU,Temperature CPU sous charge. Seuil WARN=85C"
        $csvLines += "# Temp GPU,Temperature GPU. Seuil WARN=90C"
        $csvLines += "# Disque,SMART + vitesse sequentielle I/O (Mo/s)"
        $csvLines += "# GPU,Presence/charge GPU (nvidia-smi / OHM / WMI)"
        $csvLines += "# Uptime,Duree depuis dernier redemarrage. >30j = WARN"

        # UTF-8 BOM + ecriture
        $bom     = [System.Text.Encoding]::UTF8.GetPreamble()
        $content = $csvLines -join "`r`n"
        $bytes   = $bom + [System.Text.Encoding]::UTF8.GetBytes($content)
        [System.IO.File]::WriteAllBytes($resolvedPath, $bytes)

        if ($isTempUpload) {
            $contentTxt = Build-TextContent
            [System.IO.File]::WriteAllText($uploadFile, $contentTxt, [System.Text.Encoding]::UTF8)
        }
    }
    elseif ($ext -eq '.html') {
        $rows = @($results | ForEach-Object {
            $bg  = switch ($_.Status) { 'OK'{'#0a1628'} 'WARN'{'#1e1b16'} 'FAIL'{'#1c1917'} default{'#0a1628'} }
            $cls = switch ($_.Status) { 'OK'{'ok'} 'WARN'{'warn'} 'FAIL'{'fail'} default{'na'} }
            $icon = switch ($_.Status) { 'OK'{'✓'} 'WARN'{'⚠'} 'FAIL'{'✗'} default{'—'} }
            # Tooltip : detail complet au survol
            $tipDetail = [System.Web.HttpUtility]::HtmlEncode($_.Detail.TrimEnd())
            # Formater le detail en badges cle=valeur lisibles
            $detailRaw = $_.Detail.TrimEnd()
            $detailHtml = if ($detailRaw -match '\w+=\S+') {
                $tokens = $detailRaw -split '\s*\|\s*'
                $parts  = @()
                foreach ($tok in $tokens) {
                    $subtoks = $tok.Trim() -split '\s+'
                    foreach ($s in $subtoks) {
                        if ($s -match '^([\w][\w ]+)=(.+)$') {
                            $parts += "<span class='badge'><b>$([System.Web.HttpUtility]::HtmlEncode($Matches[1]))</b>&thinsp;$([System.Web.HttpUtility]::HtmlEncode($Matches[2]))</span>"
                        } elseif ($s.Trim() -ne '') {
                            $parts += "<span class='badge-plain'>$([System.Web.HttpUtility]::HtmlEncode($s.Trim()))</span>"
                        }
                    }
                }
                $parts -join ' '
            } else {
                [System.Web.HttpUtility]::HtmlEncode($detailRaw)
            }
            "<tr style='background:$bg' title='$tipDetail'><td style='color:#64748b;white-space:nowrap'>$($_.Heure)</td><td style='color:#e2e8f0;white-space:nowrap'>$($_.Test.TrimEnd())</td><td class='$cls' style='text-align:center;white-space:nowrap'>$icon $($_.Status)</td><td style='color:#e2e8f0;white-space:nowrap'>$($_.Valeur.TrimEnd())</td><td class='detail-cell'>$detailHtml</td></tr>"
        })
        $conclusionHtml = ($conc.Lines | ForEach-Object {
            $_ -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
        }) -join '<br>'
        $cCol  = if ($conc.Critical -gt 0) { '#f87171' } elseif ($conc.Warns -gt 0) { '#fbbf24' } else { '#4ade80' }

        $html = @"
<!DOCTYPE html><html lang='fr'><head><meta charset='UTF-8'><title>occt81 Report — $env:COMPUTERNAME</title>
<style>
*{box-sizing:border-box}
body{font-family:Consolas,monospace;background:#020617;color:#e2e8f0;margin:0;padding:32px;font-size:13px}
h1{color:#38bdf8;font-size:22px;letter-spacing:1px;margin-bottom:4px}
.meta{color:#94a3b8;font-size:13px;margin-bottom:24px;line-height:1.8;border-left:3px solid #1e293b;padding-left:15px}
table{width:100%;border-collapse:collapse;font-size:13px;border:1px solid #1e293b;margin-bottom:16px}
th{background:#050d1a;color:#64748b;padding:10px 14px;text-align:left;font-weight:bold;border-bottom:2px solid #1e293b;white-space:nowrap}
td{padding:9px 14px;border-bottom:1px solid #0a1628;vertical-align:middle}
tr:hover td{background:#0f172a}
.ok{color:#4ade80;font-weight:bold}.warn{color:#fbbf24;font-weight:bold}.fail{color:#f87171;font-weight:bold}.na{color:#64748b}
.concl{margin-top:24px;padding:16px 20px;border-radius:4px;background:#0a1628;border:1px solid #1e293b;font-size:15px;color:$cCol;font-weight:bold;line-height:1.6}
.legend{margin-top:24px;padding:16px 20px;background:#050d1a;border:1px solid #1e293b;border-radius:4px}
.legend h2{color:#38bdf8;font-size:14px;margin:0 0 12px 0;letter-spacing:1px}
.legend-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(420px,1fr));gap:6px 24px}
.legend-item{color:#64748b;font-size:12px;line-height:1.6}
.legend-item b{color:#94a3b8}
.legend-status{margin-top:16px;display:flex;gap:24px;flex-wrap:wrap}
.ls{font-size:12px}.ls .ok{color:#4ade80}.ls .warn{color:#fbbf24}.ls .fail{color:#f87171}.ls .na{color:#64748b}
.foot{margin-top:12px;color:#334155;font-size:11px}
.detail-cell{font-size:11px;color:#475569;line-height:1.8;max-width:500px}
.badge{display:inline-block;background:#0f172a;border:1px solid #1e293b;border-radius:3px;padding:1px 6px;margin:1px 2px;white-space:nowrap;font-size:11px}
.badge b{color:#94a3b8;margin-right:2px}
.badge-plain{display:inline-block;color:#475569;font-size:11px;padding:1px 3px}
</style></head><body>
<h1>⚙ occt81 v4.2 — Rapport de diagnostic</h1>
<div class='meta'>
<b>Machine</b> : $env:COMPUTERNAME &nbsp;|&nbsp;
<b>Systeme</b> : $dispOS &nbsp;|&nbsp;
<b>CPU</b> : $dispCPU<br>
<b>Uptime</b> : $dispUp &nbsp;|&nbsp;
<b>Date</b> : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
</div>

<table>
<thead><tr>
  <th>HEURE</th>
  <th>TEST</th>
  <th style='text-align:center'>STATUT</th>
  <th>VALEUR</th>
  <th>DETAIL <span style='font-weight:normal;color:#334155'>(survoler pour voir le detail complet)</span></th>
</tr></thead>
<tbody>
$($rows -join "`n")
</tbody></table>

<div class='concl'>$conclusionHtml</div>

<div class='legend'>
  <h2>LEGENDE DES TESTS</h2>
  <div class='legend-grid'>
    <div class='legend-item'><b>RAM</b> — Erreurs de bits dans la RAM physique (patterns fill+verify+walking bits). 0 erreur = normal.</div>
    <div class='legend-item'><b>CPU Scheduling (moy)</b> — Temps moyen par iteration. Reflete la frequence CPU effective, pas la latence hardware.</div>
    <div class='legend-item'><b>CPU Scheduling (jitter)</b> — Irregularite du scheduling Windows. Jitter = P99/Min. 1x=parfait · &lt;3x=OK · 3-8x=WARN · &gt;8x=FAIL</div>
    <div class='legend-item'><b>WHEA total</b> — Erreurs hardware signalees par CPU/firmware. Toute valeur &gt;0 merite investigation.</div>
    <div class='legend-item'><b>WHEA critique</b> — Sous-ensemble WHEA de severite critique. Zero tolerance — peut causer des crashs.</div>
    <div class='legend-item'><b>Temp CPU</b> — Temperature CPU sous charge. Seuil WARN : 85°C (throttling probable au-dela).</div>
    <div class='legend-item'><b>Temp GPU</b> — Temperature GPU. Seuil WARN : 90°C.</div>
    <div class='legend-item'><b>Disque</b> — Sante SMART + vitesse sequentielle (Mo/s). SMART KO = risque de perte de donnees.</div>
    <div class='legend-item'><b>GPU</b> — Presence et utilisation GPU (nvidia-smi / OHM / WMI).</div>
    <div class='legend-item'><b>Uptime</b> — Duree depuis le dernier redemarrage. &gt;30 jours = WARN (mises a jour en attente ?).</div>
  </div>
  <div class='legend-status'>
    <div class='ls'><span class='ok'>✓ OK</span> — Dans les seuils normaux</div>
    <div class='ls'><span class='warn'>⚠ WARN</span> — A surveiller, pas critique</div>
    <div class='ls'><span class='fail'>✗ FAIL</span> — Probleme detecte, intervention recommandee</div>
    <div class='ls'><span class='na'>— N/A</span> — Test impossible (droits ou outil manquant)</div>
  </div>
</div>

<p class='foot'>Genere par occt81 v4.2 — $env:COMPUTERNAME — $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')</p>
</body></html>
"@
        [System.IO.File]::WriteAllText($resolvedPath, $html, [System.Text.Encoding]::UTF8)
        if ($isTempUpload) {
            $contentTxt = Build-TextContent
            [System.IO.File]::WriteAllText($uploadFile, $contentTxt, [System.Text.Encoding]::UTF8)
        }
    }
    elseif ($ext -eq '.json') {
        # JSON enrichi : metadata complete + conclusion + legende integrée
        $legendeTests = @(
            @{ Test='RAM';                     Description='Erreurs de bits RAM (fill+verify+walking). 0 = normal.' }
            @{ Test='CPU Scheduling (moy)';    Description='Temps moyen/iteration — frequence CPU effective (pas latence hardware)' }
            @{ Test='CPU Scheduling (jitter)'; Description='Jitter = P99/Min. 1x=parfait, <3x=OK, 3-8x=WARN, >8x=FAIL' }
            @{ Test='WHEA total';              Description='Erreurs hardware CPU/firmware. >0 = investigation recommandee' }
            @{ Test='WHEA critique';           Description='Severite critique WHEA. Zero tolerance.' }
            @{ Test='Temp CPU';                Description='Temperature CPU sous charge. Seuil WARN=85C' }
            @{ Test='Temp GPU';                Description='Temperature GPU. Seuil WARN=90C' }
            @{ Test='Disque';                  Description='SMART + vitesse sequentielle I/O (Mo/s)' }
            @{ Test='GPU';                     Description='Presence/charge GPU (nvidia-smi / OHM / WMI)' }
            @{ Test='Uptime';                  Description='Duree depuis dernier redemarrage. >30j = WARN' }
        )
        $legendeStatuts = @(
            @{ Statut='OK';   Signification='Dans les seuils normaux' }
            @{ Statut='WARN'; Signification='A surveiller — pas critique' }
            @{ Statut='FAIL'; Signification='Probleme detecte — intervention recommandee' }
            @{ Statut='N/A';  Signification='Test impossible (droits insuffisants ou outil manquant)' }
        )
        $payload2 = [ordered]@{
            _outil     = 'occt81 v4.2'
            Date       = (Get-Date -Format 'o')
            DateFr     = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
            Machine    = $env:COMPUTERNAME
            Systeme    = $dispOS
            CPU        = $dispCPU
            Uptime     = $dispUp
            Conclusion = [ordered]@{
                Statut   = if ($conc.Critical -gt 0) { 'FAIL' } elseif ($conc.Warns -gt 0) { 'WARN' } else { 'OK' }
                Critical = $conc.Critical
                Warns    = $conc.Warns
                Resume   = $conc.Lines -join ' | '
            }
            Results    = @($results | ForEach-Object {
                [ordered]@{
                    Heure  = $_.Heure
                    Test   = $_.Test.TrimEnd()
                    Status = $_.Status
                    Valeur = $_.Valeur.TrimEnd()
                    Detail = $_.Detail.TrimEnd()
                }
            })
            Legende    = [ordered]@{
                Statuts = $legendeStatuts
                Tests   = $legendeTests
            }
        }
        $payload2 | ConvertTo-Json -Depth 6 | Set-Content $resolvedPath -Encoding UTF8
        if ($isTempUpload) {
            $contentTxt = Build-TextContent
            [System.IO.File]::WriteAllText($uploadFile, $contentTxt, [System.Text.Encoding]::UTF8)
        }
    }
    else {
        # .txt (defaut)
        $contentTxt = Build-TextContent
        [System.IO.File]::WriteAllText($resolvedPath, $contentTxt, [System.Text.Encoding]::UTF8)
    }

    if (-not $Silent) { Write-Info "Export : $resolvedPath" -color 'Green' }

    # Uploads
    if ($UploadDPaste -or $UploadGoFile) {
        if (-not (Test-Path $uploadFile)) {
            Write-Warning "Fichier upload introuvable : $uploadFile"
        }
        else {
            Write-Info "Upload en cours ($([System.IO.Path]::GetFileName($uploadFile)))..." -color 'Cyan'
            if ($UploadDPaste) {
                if ($contentTxt -eq '') { $contentTxt = Get-Content $uploadFile -Raw }
                Invoke-DPasteUpload -Text $contentTxt -Title "occt81_$env:COMPUTERNAME"
            }
            if ($UploadGoFile) {
                Invoke-GoFileUpload -FilePath $uploadFile
            }
            if ($isTempUpload -and (Test-Path $uploadFile)) {
                Remove-Item $uploadFile -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

#─────────────────────────────────────────────────────────────────────────────
# MODE WATCH CLI
#─────────────────────────────────────────────────────────────────────────────
function Start-WatchMode {
    param([int]$intervalSec)

    Write-Header "WATCH MODE — intervalle ${intervalSec}s | Ctrl+C pour arreter"
    Write-Info "Tests : $($watchTests -join ', ')" -color 'DarkGray'

    $savedTests            = $script:testsToRun
    $script:testsToRun     = $watchTests
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
    }
    finally {
        $script:testsToRun = $savedTests
    }
}

#─────────────────────────────────────────────────────────────────────────────
# POINT D ENTREE
#─────────────────────────────────────────────────────────────────────────────
if ($Watch -gt 0) {
    Write-Header "occt81 v4.2 — $env:COMPUTERNAME"
    Start-WatchMode -intervalSec $Watch
    exit 0
}

# CLI standard
Write-Header "occt81 v4.2 — $env:COMPUTERNAME"

if (-not $Silent) {
    $osName = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
    Write-Info "OS     : $osName" -color 'DarkGray'
    Write-Info "Admin  : $(if ($IsAdmin) { 'Oui' } else { 'Non — WHEA/Temp/SMART indisponibles' })" `
        -color $(if ($IsAdmin) { 'DarkGray' } else { 'DarkYellow' })
    Write-Info "Tests  : $($script:testsToRun -join ', ')" -color 'DarkGray'
    if ($configPath) { Write-Info "Config : $configPath" -color 'DarkGray' }
}

$ohmProc = $null

try {
    # Lancement OHM UNE SEULE FOIS avant les tests
    $ohmProc = Start-OhmIfNeeded

    Invoke-AllTests
    Write-Summary
    Save-History | Out-Null
    if ($Compare) { Compare-History -jsonPath $Compare }
    Export-Report
}
catch {
    # Crash log enrichi
    try {
        $crashDir = Join-Path $env:APPDATA 'occt81'
        if (-not (Test-Path $crashDir)) { New-Item -ItemType Directory $crashDir -Force | Out-Null }
        $crashLog = Join-Path $crashDir "crashlog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

        # Chaine d exceptions complete (inner exceptions)
        $exChain = @()
        $ex = $_.Exception
        $depth = 0
        while ($ex -and $depth -lt 10) {
            $exChain += "  [$depth] $($ex.GetType().FullName) : $($ex.Message)"
            $ex = $ex.InnerException
            $depth++
        }

        # Infos systeme
        $osInfo  = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
        $cpuInfo = Get-CimInstance Win32_Processor       -ErrorAction SilentlyContinue | Select-Object -First 1
        $ramInfo = Get-CimInstance Win32_PhysicalMemory  -ErrorAction SilentlyContinue |
                   Measure-Object -Property Capacity -Sum
        $uptime  = if ($osInfo) { (Get-Date) - $osInfo.LastBootUpTime } else { $null }

        $sysLines = @(
            "OS          : $(if ($osInfo)  { $osInfo.Caption + ' (Build ' + $osInfo.BuildNumber + ')' } else { 'Inconnu' })",
            "Architecture: $($env:PROCESSOR_ARCHITECTURE)",
            "CPU         : $(if ($cpuInfo) { $cpuInfo.Name } else { 'Inconnu' })",
            "Coeurs log. : $(if ($cpuInfo) { $cpuInfo.NumberOfLogicalProcessors } else { '?' })",
            "RAM totale  : $(if ($ramInfo.Sum) { [Math]::Round($ramInfo.Sum/1GB,1).ToString() + ' Go' } else { 'Inconnu' })",
            "RAM libre   : $(if ($osInfo) { [Math]::Round($osInfo.FreePhysicalMemory/1MB,0).ToString() + ' Mo' } else { 'Inconnu' })",
            "Uptime      : $(if ($uptime) { "$([Math]::Floor($uptime.TotalDays))j $($uptime.Hours)h $($uptime.Minutes)m" } else { 'N/A' })",
            "PowerShell  : $($PSVersionTable.PSVersion)",
            "CLR         : $([System.Runtime.InteropServices.RuntimeEnvironment]::GetSystemVersion())",
            "Admin       : $(if ($IsAdmin) { 'Oui' } else { 'Non' })",
            "User        : $($env:USERDOMAIN)\$($env:USERNAME)"
        )

        # Etat compilation RamEngine
        $ramEngineStatus = if (([System.Management.Automation.PSTypeName]'RamEngine').Type) {
            if ($script:RamEngineSafeMode) { 'Compile en mode SAFE (Buffer.BlockCopy, sans /unsafe)' }
            else { 'Compile avec succes (mode unsafe natif)' }
        } elseif ($script:RamEngineCompileError) {
            "ECHEC — $script:RamEngineCompileError"
        } else {
            'Non compile (raison inconnue)'
        }

        # Resultats partiels deja collectes
        $partialResults = if ($results -and $results.Count -gt 0) {
            $results | ForEach-Object {
                "  $($_.Heure)  $($_.Test.PadRight(20)) [$($_.Status.PadRight(4))] $($_.Valeur)  $($_.Detail)"
            }
        } else { @('  (aucun resultat collecte avant le crash)') }

        # Parametres passes au script
        $paramLines = @(
            "Tests      : $($script:testsToRun -join ', ')",
            "Passes     : $Passes",
            "RamSize    : $RamSize Mo",
            "Export     : $(if ($Export) { $Export } else { '(aucun)' })",
            "Config     : $(if ($configPath) { $configPath } else { '(defaut)' })",
            "Watch      : $Watch",
            "Compare    : $(if ($Compare) { $Compare } else { '(aucun)' })"
        )

        $errContent = (@(
            "╔══════════════════════════════════════════════════════════╗",
            "  occt81 v4.2 — CRASH REPORT",
            "  $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')",
            "╚══════════════════════════════════════════════════════════╝",
            "",
            "── MACHINE ──────────────────────────────────────────────────",
            "Nom         : $env:COMPUTERNAME"
        ) + $sysLines + @(
            "",
            "── ERREUR ───────────────────────────────────────────────────",
            "Type        : $($_.Exception.GetType().FullName)",
            "Message     : $($_.Exception.Message)",
            "Position    : $($_.InvocationInfo.PositionMessage)",
            "",
            "Chaine d exceptions :"
        ) + $exChain + @(
            "",
            "Stack PowerShell :",
            $_.ScriptStackTrace,
            "",
            "── PARAMETRES ───────────────────────────────────────────────"
        ) + $paramLines + @(
            "",
            "── ETAT RAMENGINE ───────────────────────────────────────────",
            "Statut      : $ramEngineStatus",
            "",
            "── RESULTATS PARTIELS (avant crash) ─────────────────────────"
        ) + $partialResults + @("")) -join "`n"

        [System.IO.File]::WriteAllText($crashLog, $errContent, [System.Text.Encoding]::UTF8)

        if (-not $Silent) {
            Write-Host "`n!! ERREUR FATALE : $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "   Type     : $($_.Exception.GetType().FullName)" -ForegroundColor DarkYellow
            Write-Host "   Position : $($_.InvocationInfo.PositionMessage)" -ForegroundColor DarkYellow
            if ($script:RamEngineCompileError) {
                Write-Host "   RamEngine: $script:RamEngineCompileError" -ForegroundColor DarkYellow
            }
            Write-Host "   Crash log: $crashLog" -ForegroundColor Cyan
        }
    }
    catch { <# ne pas planter le handler de crash #> }
    throw
}
finally {
    # Cleanup OHM
    try {
        if ($ohmProc) {
            if (-not $Silent) { Write-Info "Arret OHM..." -color 'DarkGray' }
            $ohmProc.CloseMainWindow() | Out-Null
            Start-Sleep -Seconds 1
            if (-not $ohmProc.HasExited) { $ohmProc | Stop-Process -Force }
        }
        # Securite : tuer tout OHM residuel lance par ce script
        Get-Process OpenHardwareMonitor -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    catch { <# cleanup ne doit jamais planter #> }
}
