[CmdletBinding()]
param(
    [switch]$Force,
    [switch]$NoBrowser
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$panelRoot = $PSScriptRoot
$venvPython = Join-Path $panelRoot '.venv\Scripts\python.exe'
$port = 8766
$expectedApp = 'control_panel_v2.main:app'

function Get-PanelProcesses {
    $listeners = @(Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue)
    $processes = @()
    foreach ($listener in $listeners) {
        $process = Get-CimInstance Win32_Process -Filter "ProcessId = $($listener.OwningProcess)" -ErrorAction SilentlyContinue
        if ($process -and $process.CommandLine -like "*$expectedApp*") { $processes += $process }
    }
    return $processes
}

function Get-Descendants([int]$ParentId) {
    $all = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue)
    $found = @()
    $pending = [System.Collections.Generic.Queue[int]]::new()
    $pending.Enqueue($ParentId)
    while ($pending.Count -gt 0) {
        $parent = $pending.Dequeue()
        foreach ($process in $all | Where-Object { $_.ParentProcessId -eq $parent }) {
            $found += $process
            $pending.Enqueue([int]$process.ProcessId)
        }
    }
    return $found
}

function Stop-ProcessTree([int]$RootId) {
    $descendants = @(Get-Descendants $RootId)
    foreach ($process in ($descendants | Sort-Object { $_.ProcessId } -Descending)) {
        Stop-Process -Id $process.ProcessId -Force -ErrorAction SilentlyContinue
    }
    Stop-Process -Id $RootId -Force -ErrorAction SilentlyContinue
}

$existing = @(Get-PanelProcesses)
foreach ($server in $existing) {
    $descendants = @(Get-Descendants ([int]$server.ProcessId))
    if ($descendants.Count -gt 0 -and -not $Force) {
        $details = ($descendants | ForEach-Object { "$($_.ProcessId): $($_.Name)" }) -join ', '
        throw "Control panel $expectedApp has child processes ($details). Restart refused. Use -Force only if you intend to stop the active workers."
    }
    Stop-ProcessTree ([int]$server.ProcessId)
}

$deadline = (Get-Date).AddSeconds(15)
do {
    if (@(Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue).Count -eq 0) { break }
    Start-Sleep -Milliseconds 250
} while ((Get-Date) -lt $deadline)
if (@(Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue).Count -gt 0) { throw "Port $port was not released." }

if (-not (Test-Path $venvPython)) { & py -3 -m venv (Join-Path $panelRoot '.venv') }
& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -r (Join-Path $panelRoot 'requirements.txt')

Start-Process $venvPython -WorkingDirectory $projectRoot -ArgumentList @('-m', 'uvicorn', $expectedApp, '--host', '127.0.0.1', '--port', $port)
$deadline = (Get-Date).AddSeconds(15)
do {
    if (@(Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue).Count -gt 0) { break }
    Start-Sleep -Milliseconds 250
} while ((Get-Date) -lt $deadline)
if (@(Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue).Count -eq 0) { throw "Control panel did not start listening on port $port." }
if (-not $NoBrowser) { Start-Process "http://127.0.0.1:$port/" }
Write-Host "Restarted V2 control panel on http://127.0.0.1:$port/"
