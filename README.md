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
    $ohmExe = Get-ChildItem $dir -Recurse -Filter "OpenHardwareMonitor.exe" | Select-Object -First 1
    $ohmProc = $null
    if ($ohmExe) {
        Write-Host "Démarrage OHM..." -ForegroundColor DarkGray
        $ohmProc = Start-Process $ohmExe.FullName -WindowStyle Minimized -PassThru -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
    }

    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    Write-Host "Lancement occt81.ps1..." -ForegroundColor Cyan
    & (Join-Path $dir "occt81.ps1") -Export $file

    if ($ohmProc -and -not $ohmProc.HasExited) {
        $ohmProc.CloseMainWindow() | Out-Null
        Start-Sleep -Milliseconds 800
        if (-not $ohmProc.HasExited) { $ohmProc | Stop-Process -Force -ErrorAction SilentlyContinue }
    }
    Get-Process OpenHardwareMonitor -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

    Write-Host "`nDOSSIER DES OUTILS :" -ForegroundColor Cyan
    if ($env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode') {
        $esc = [char]27; $url = "file:///" + ($dir.Replace('\','/'))
        Write-Host "$esc]8;;$url$esc\$dir$esc]8;;$esc\" -ForegroundColor Yellow
    } else { Write-Host $dir -ForegroundColor Yellow }

    if (Test-Path $file) {
        Write-Host "`nUpload GoFile..." -ForegroundColor Cyan
        try {
            $srv = ((Invoke-RestMethod "https://api.gofile.io/servers").data.servers | Get-Random).name
            $boundary = [System.Guid]::NewGuid().ToString()
            $enc = [System.Text.Encoding]::UTF8
            $fileBytes = [System.IO.File]::ReadAllBytes($file)
            $fileName  = [System.IO.Path]::GetFileName($file)
            $head = "--$boundary`r`nContent-Disposition: form-data; name=`"file`"; filename=`"$fileName`"`r`nContent-Type: text/plain`r`n`r`n"
            $bodyBytes = $enc.GetBytes($head) + $fileBytes + $enc.GetBytes("`r`n--$boundary--`r`n")
            $resp = Invoke-RestMethod -Uri "https://$srv.gofile.io/contents/uploadfile" -Method Post -ContentType "multipart/form-data; boundary=$boundary" -Body $bodyBytes
            $dl = if ($resp.data.downloadPage) { $resp.data.downloadPage } else { "https://gofile.io/d/$($resp.data.code)" }
            $entry = "Rapport_occt81.txt -> $dl"
            Write-Host "`n=== LIEN RAPPORT ===" -ForegroundColor Cyan
            Write-Host $entry -ForegroundColor Yellow
            $entry | Out-File "$env:USERPROFILE\Desktop\lien_occt81.txt" -Encoding UTF8
        } catch {
            Write-Host "Upload échoué : $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Rapport local : $file" -ForegroundColor DarkYellow
        }
    }
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
