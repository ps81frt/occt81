```powershell
& {
$occt = {
    $file = "$env:USERPROFILE\Desktop\Rapport_occt81.txt"
    $zip = "$env:TEMP\occt81.zip"
    irm "https://github.com/ps81frt/occt81/archive/refs/heads/main.zip" -OutFile $zip
    Expand-Archive $zip "$env:TEMP\occt81" -Force
    $dir = (Get-ChildItem "$env:TEMP\occt81" -Recurse -Filter "occt81.ps1" | Select-Object -First 1).Directory.FullName
    $ohm = Get-ChildItem $dir -Recurse -Filter "OpenHardwareMonitor.exe" | Select-Object -First 1
    if ($ohm) { $p = Start-Process $ohm.FullName -Verb RunAs -WindowStyle Minimized -PassThru; Start-Sleep -Seconds 3 }
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    & (Join-Path $dir "occt81.ps1") -Export "$env:USERPROFILE\Desktop\Rapport_occt81.txt"
    taskkill /IM OpenHardwareMonitor.exe /F /T > $null 2>&1    
    $esc = [char]27
    $url = "file:///" + ($dir.Replace('\','/'))
    $link = "$esc]8;;$url$esc\$dir$esc]8;;$esc\"
    
    Write-Host "`nDOSSIER DES OUTILS :" -ForegroundColor Cyan
    Write-Host $link -ForegroundColor Yellow
    
    if(Test-Path $file){
        $dl = (curl.exe -s -F "file=@$file" https://store1.gofile.io/uploadFile | ConvertFrom-Json).data.downloadPage
        $entry = "1. $(Split-Path $file -Leaf) -> $dl"
        Write-Host "`n=== Récap liens ===" -ForegroundColor Cyan
        Write-Host $entry -ForegroundColor Yellow
        $entry | Out-File "$env:USERPROFILE\Desktop\liens_upload.txt" -Encoding UTF8
    }

    }
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $bytes = [System.Text.Encoding]::Unicode.GetBytes("& {$occt}")
    $encoded = [Convert]::ToBase64String($bytes)
    $exe = if ($PSVersionTable.PSVersion.Major -ge 6) { "pwsh.exe" } else { "powershell.exe" }
    Start-Process $exe "-NoProfile -ExecutionPolicy Bypass -NoExit -EncodedCommand $encoded" -Verb RunAs
    exit
}
& $occt
}

```
