$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $PSScriptRoot
$panelRoot = $PSScriptRoot
$venvPython = Join-Path $panelRoot '.venv\Scripts\python.exe'

if (-not (Test-Path $venvPython)) {
    & py -3 -m venv (Join-Path $panelRoot '.venv')
}

& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -r (Join-Path $panelRoot 'requirements.txt')

$port = 8766
Start-Process $venvPython -WorkingDirectory $projectRoot -ArgumentList @(
    '-m', 'uvicorn', 'control_panel_v2.main:app', '--host', '127.0.0.1', '--port', $port
)
Start-Process "http://127.0.0.1:$port/"
