# Test-Loop: startet die installierte App, misst Lebensdauer, prüft Singleton, liest Logs
$ErrorActionPreference = "Continue"
$exe = "$env:LOCALAPPDATA\Programs\Magic-Voice\MagicVoice.exe"
$logFile = "$env:LOCALAPPDATA\Magic-Voice\logs\startup-crash.log"

function Pass($msg) { Write-Host "  [OK] $msg" -ForegroundColor Green }
function Fail($msg) { Write-Host "  [!!] $msg" -ForegroundColor Red; $script:fails++ }
function Info($msg) { Write-Host "  [i ] $msg" -ForegroundColor Cyan }
$script:fails = 0

if (-not (Test-Path $exe)) { Write-Host "App nicht installiert: $exe" -ForegroundColor Red; exit 1 }

# Test 1: App startet stabil und läuft mind. 15s
Write-Host "=== Test 1: Stabiler Lauf 15s ===" -ForegroundColor Cyan
Get-Process -Name "MagicVoice" -ErrorAction SilentlyContinue | Stop-Process -Force; Start-Sleep -Milliseconds 500
$logSizeBefore = if (Test-Path $logFile) { (Get-Item $logFile).Length } else { 0 }
$p = Start-Process -FilePath $exe -PassThru
$startTime = Get-Date
$liveAfter = @{}
foreach ($wait in @(2, 5, 10, 15)) {
    Start-Sleep -Seconds ($wait - ((Get-Date) - $startTime).TotalSeconds)
    $alive = Get-Process -Id $p.Id -ErrorAction SilentlyContinue
    $liveAfter[$wait] = if ($alive) { $true } else { $false }
}
foreach ($wait in @(2, 5, 10, 15)) {
    if ($liveAfter[$wait]) { Pass "Nach ${wait}s: lebt" } else { Fail "Nach ${wait}s: tot" }
}
$alive = Get-Process -Id $p.Id -ErrorAction SilentlyContinue
if ($alive) {
    Info ("Window-Handle: {0} (0=kein sichtbares Fenster)" -f $alive.MainWindowHandle)
    Info ("Threads: {0}, WS: {1} MB" -f $alive.Threads.Count, [Math]::Round($alive.WorkingSet64/1MB,1))
}

# Test 2: Singleton — zweiter Start muss sich beenden, erster muss leben
Write-Host "`n=== Test 2: Singleton (zweiter Start beenden, erster leben) ===" -ForegroundColor Cyan
$first = Get-Process -Name "MagicVoice" -ErrorAction SilentlyContinue | Select-Object -First 1
if ($first) {
    Info "Erste Instanz lebt: PID=$($first.Id)"
    $second = Start-Process -FilePath $exe -PassThru
    Start-Sleep -Seconds 4
    $secondAlive = Get-Process -Id $second.Id -ErrorAction SilentlyContinue
    $firstStill = Get-Process -Id $first.Id -ErrorAction SilentlyContinue
    if ($firstStill) { Pass "Erste Instanz lebt noch nach 2. Start" } else { Fail "Erste Instanz GESTORBEN" }
    if (-not $secondAlive) { Pass "Zweite Instanz hat sich beendet (Singleton greift)" }
    else { Fail "Zweite Instanz läuft (Singleton greift NICHT)"; Stop-Process -Id $second.Id -Force }
} else { Fail "Erste Instanz war beim Singleton-Test nicht mehr da" }

# Test 3: Logs nach Test
Write-Host "`n=== Test 3: Logs ===" -ForegroundColor Cyan
$logSizeAfter = if (Test-Path $logFile) { (Get-Item $logFile).Length } else { 0 }
Info ("Log-Wachstum: {0} → {1} Bytes" -f $logSizeBefore, $logSizeAfter)
if (Test-Path $logFile) {
    Write-Host "  Letzte Log-Einträge:"
    Get-Content $logFile -Tail 8 | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
}

# Crash-Spuren?
$crashEvents = Get-WinEvent -LogName Application -MaxEvents 30 -ErrorAction SilentlyContinue |
    Where-Object { $_.TimeCreated -gt $startTime -and $_.Message -match "MagicVoice" -and $_.LevelDisplayName -match "Fehler|Error" }
if ($crashEvents) {
    Fail ("Crash-Events seit Test-Start: {0}" -f $crashEvents.Count)
    $crashEvents | Select-Object -First 1 | ForEach-Object {
        Write-Host ("    {0}" -f ($_.Message -replace "`r`n", " | ").Substring(0, 300)) -ForegroundColor Red
    }
} else { Pass "Keine Crash-Events im Application-Log seit Test-Start" }

# Cleanup
Get-Process -Name "MagicVoice" -ErrorAction SilentlyContinue | Stop-Process -Force

Write-Host ""
if ($script:fails -eq 0) { Write-Host "RESULT: ALLE TESTS BESTANDEN" -ForegroundColor Green; exit 0 }
else { Write-Host "RESULT: $script:fails FEHLER" -ForegroundColor Red; exit $script:fails }
