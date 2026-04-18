```powershell
& {
$occt = {
    $file    = "$env:USERPROFILE\Desktop\Rapport_occt81.txt"
    $zip     = "$env:TEMP\occt81.zip"
    $extract = "$env:TEMP\occt81_run"
    if (Test-Path $extract) { Remove-Item $extract -Recurse -Force -ErrorAction SilentlyContinue }

    Write-Host "`nTéléchargement occt81..." -ForegroundColor Cyan
    Invoke-WebRequest "https://github.com/ps81frt/occt81/archive/refs/heads/main.zip" -OutFile $zip -UseBasicParsing
    Expand-Archive $zip $extract -Force
    Remove-Item $zip -Force -ErrorAction SilentlyContinue

    $dir = (Get-ChildItem $extract -Recurse -Filter "occt81.ps1" | Select-Object -First 1).Directory.FullName
    Write-Host "Dossier : $dir" -ForegroundColor DarkGray

    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    Write-Host "Lancement occt81.ps1..." -ForegroundColor Cyan
    & (Join-Path $dir "occt81.ps1") -Export $file -Tests Tout -UploadGoFile -UploadDPaste

    Write-Host "`nDOSSIER DES OUTILS :" -ForegroundColor Cyan
    if ($env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode') {
        $esc = [char]27; $url = "file:///" + ($dir.Replace('\','/'))
        Write-Host "$esc]8;;$url$esc\$dir$esc]8;;$esc\" -ForegroundColor Yellow
    } else { Write-Host $dir -ForegroundColor Yellow }

    Write-Host "`nTerminé." -ForegroundColor Green
}
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $bytes   = [System.Text.Encoding]::Unicode.GetBytes("& {$($occt.ToString())}")
    $encoded = [Convert]::ToBase64String($bytes)
    $exe     = if ($PSVersionTable.PSVersion.Major -ge 6) { "pwsh.exe" } else { "powershell.exe" }
    Start-Process $exe "-NoProfile -ExecutionPolicy Bypass -NoExit -EncodedCommand $encoded" -Verb RunAs
    exit
}
& $occt
}

```
