$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$source = Join-Path $PSScriptRoot "..\dist\納期回答書作成.exe"
$source = [System.IO.Path]::GetFullPath($source)

$targets = @(
    "\\flsv04\316京葉\納期回答書ツールフォルダ\納期回答書作成.exe",
    "C:\Users\kento.kashiwabara\Desktop\納期回答書ツールフォルダ（DT)\納期回答書作成.exe",
    "\\flsv04\001全社_共有\納期回答書作成ツール\納期回答書作成.exe"
)

Write-Host "=== 配布前 ==="
Write-Host ("Source: {0}" -f $source)
$sf = Get-Item $source
Write-Host ("  {0}  {1:N0} bytes" -f $sf.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $sf.Length)
Write-Host ""

foreach ($t in $targets) {
    $dir = Split-Path $t -Parent
    Write-Host ("Target: {0}" -f $t)
    Write-Host ("  Dir exists: {0}" -f (Test-Path $dir))
    if (Test-Path $t) {
        $f = Get-Item $t
        Write-Host ("  Current: {0}  {1:N0} bytes" -f $f.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $f.Length)
    } else {
        Write-Host "  Current: (not found)"
    }
}

Write-Host ""
Write-Host "=== Copying ==="
$ok = 0
foreach ($t in $targets) {
    $dir = Split-Path $t -Parent
    if (-not (Test-Path $dir)) {
        Write-Host ("SKIP (dir not found): {0}" -f $t)
        continue
    }
    try {
        Copy-Item -Path $source -Destination $t -Force -ErrorAction Stop
        Write-Host ("OK: {0}" -f $t)
        $ok++
    } catch {
        Write-Host ("FAIL: {0} -- {1}" -f $t, $_.Exception.Message)
    }
}

Write-Host ""
Write-Host "=== 配布後 ==="
foreach ($t in $targets) {
    if (Test-Path $t) {
        $f = Get-Item $t
        Write-Host ("{0}  {1:N0} bytes  {2}" -f $f.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $f.Length, $t)
    } else {
        Write-Host ("(not found)  {0}" -f $t)
    }
}
Write-Host ""
Write-Host ("{0}/{1} completed." -f $ok, $targets.Count)
