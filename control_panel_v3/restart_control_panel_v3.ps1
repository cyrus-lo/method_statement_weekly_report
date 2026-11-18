[CmdletBinding()]
param([switch]$Force,[switch]$NoBrowser)
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$panelRoot = $PSScriptRoot
$expectedApp = 'control_panel_v3.main:app'
$venvPython = Join-Path $panelRoot '.venv\Scripts\python.exe'
function Get-V3Setting([string]$Name, [string]$Default) {
    $processValue = [Environment]::GetEnvironmentVariable($Name, 'Process')
    if (-not [string]::IsNullOrWhiteSpace($processValue)) { return $processValue.Trim() }
    $envFile = Join-Path $panelRoot '.env'
    if (Test-Path $envFile) {
        foreach ($line in Get-Content $envFile) {
            if ($line -match "^\s*$([regex]::Escape($Name))\s*=\s*(.*?)\s*$") { return $Matches[1].Trim().Trim('"').Trim("'") }
        }
    }
    return $Default
}
$hostAddress = Get-V3Setting 'CONTROL_PANEL_V3_HOST' '127.0.0.1'
$portText = Get-V3Setting 'CONTROL_PANEL_V3_PORT' '8767'
$port = 0
if ($hostAddress -notin @('127.0.0.1','localhost')) { throw 'V3 must bind to localhost only.' }
if (-not [int]::TryParse($portText, [ref]$port) -or $port -lt 1 -or $port -gt 65535 -or $port -eq 8766) { throw "Invalid or protected V3 port: $portText" }
$servers = @()
$listeners = @(Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue)
foreach ($listener in $listeners) {
    $owner = Get-CimInstance Win32_Process -Filter "ProcessId = $($listener.OwningProcess)" -ErrorAction SilentlyContinue
    if ($owner -and $owner.CommandLine -like "*$expectedApp*") { $servers += $owner }
    else { throw "Port $port is owned by another process; restart refused without touching it." }
}
function Get-Children([int]$ParentId) {
    $all = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue)
    $queue = [System.Collections.Generic.Queue[int]]::new(); $queue.Enqueue($ParentId); $found = @()
    while ($queue.Count -gt 0) { $parent = $queue.Dequeue(); foreach ($child in $all | Where-Object { $_.ParentProcessId -eq $parent }) { $found += $child; $queue.Enqueue([int]$child.ProcessId) } }
    return $found
}
foreach ($server in $servers) {
    $children = @(Get-Children ([int]$server.ProcessId))
    if ($children.Count -gt 0 -and -not $Force) { throw "V3 has active child processes; restart refused. Use -Force only if stopping workers is intended." }
    foreach ($child in ($children | Sort-Object ProcessId -Descending)) { Stop-Process -Id $child.ProcessId -Force -ErrorAction SilentlyContinue }
    Stop-Process -Id $server.ProcessId -Force -ErrorAction SilentlyContinue
}
$deadline = (Get-Date).AddSeconds(15)
do {
    $remaining = @(Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue)
    if ($remaining.Count -eq 0) { break }
    $unexpected = @($remaining | Where-Object { $listener = $_; $owner = Get-CimInstance Win32_Process -Filter "ProcessId = $($listener.OwningProcess)" -ErrorAction SilentlyContinue; -not ($owner -and $owner.CommandLine -like "*$expectedApp*") })
    if ($unexpected.Count -gt 0) { throw "Port $port was claimed by another process; V3 restart aborted." }
    Start-Sleep -Milliseconds 250
} while ((Get-Date) -lt $deadline)
if (@(Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue).Count -gt 0) { throw "V3 did not release port $port within 15 seconds; restart aborted." }
if (-not (Test-Path $venvPython)) { & py -3 -m venv (Join-Path $panelRoot '.venv'); if ($LASTEXITCODE -ne 0) { throw 'Failed to create the V3 virtual environment.' } }
& $venvPython -m pip install -r (Join-Path $panelRoot 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'Failed to install V3 requirements.' }
$server = Start-Process $venvPython -PassThru -WorkingDirectory $projectRoot -ArgumentList @('-m','uvicorn',$expectedApp,'--host',$hostAddress,'--port',"$port")
$deadline = (Get-Date).AddSeconds(30)
$ready = $false
do {
    if ($server.HasExited) { throw "V3 server exited during startup (code $($server.ExitCode))." }
    try {
        $health = Invoke-RestMethod -Uri "http://${hostAddress}:$port/health" -TimeoutSec 2
        if ($health.service -eq 'control_panel_v3' -and $health.status -eq 'ok') { $ready = $true; break }
    } catch { }
    Start-Sleep -Milliseconds 300
} while ((Get-Date) -lt $deadline)
if (-not $ready) { throw "V3 did not pass its HTTP health check on port $port. Inspect the V3 server process before retrying." }
if (-not $NoBrowser) { Start-Process "http://${hostAddress}:$port/" }
Write-Host "Restarted V3 control panel and passed /health at http://${hostAddress}:$port/ (PID $($server.Id))."
