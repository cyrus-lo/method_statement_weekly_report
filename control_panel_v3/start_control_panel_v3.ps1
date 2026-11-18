$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$panelRoot = $PSScriptRoot
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
$expectedApp = 'control_panel_v3.main:app'
$listeners = @(Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue)
if ($listeners.Count -gt 0) {
    $owners = @($listeners | ForEach-Object { Get-CimInstance Win32_Process -Filter "ProcessId = $($_.OwningProcess)" -ErrorAction SilentlyContinue })
    if ($owners | Where-Object { $_.CommandLine -like "*$expectedApp*" }) { throw "V3 already appears to be running on port $port. Use restart_control_panel_v3.ps1 instead." }
    throw "Port $port is already owned by another process; V3 startup refused."
}
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
Start-Process "http://${hostAddress}:$port/"
Write-Host "V3 control panel started and passed /health at http://${hostAddress}:$port/ (PID $($server.Id))."
