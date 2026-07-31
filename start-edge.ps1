# Lanzador todo-en-uno (modo Edge + attach, el estable en este equipo):
#   1. actualiza a la última versión (git pull)
#   2. abre Microsoft Edge con un puerto de depuración (perfil aparte)
#   3. conecta la herramienta a ese Edge y entra directo al flujo de notas
#
# Para correrlo: doble clic en INICIAR.bat, o en PowerShell:
#   powershell -ExecutionPolicy Bypass -File .\start-edge.ps1

$ErrorActionPreference = "Continue"
Set-Location $PSScriptRoot

$port = 9222

Write-Host "Actualizando a la ultima version..." -ForegroundColor Cyan
try { git pull } catch { Write-Host "(no se pudo actualizar; sigo con la version actual)" -ForegroundColor Yellow }

$venvPy = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $venvPy)) {
    Write-Host "Falta la instalacion. Corre primero:" -ForegroundColor Yellow
    Write-Host "    powershell -ExecutionPolicy Bypass -File .\setup.ps1"
    Read-Host "Presiona ENTER para cerrar"
    exit 1
}

Write-Host "Abriendo Microsoft Edge..." -ForegroundColor Cyan
Start-Process msedge -ArgumentList "--remote-debugging-port=$port","--user-data-dir=$env:LOCALAPPDATA\edge-grades"
Start-Sleep -Seconds 4

Write-Host ""
Write-Host "En la ventana de Edge que se abrio:" -ForegroundColor Green
Write-Host "  inicia sesion en la plataforma y abre la planilla de notas." -ForegroundColor Green
Write-Host "Luego la herramienta te pedira ENTER para continuar." -ForegroundColor Green
Write-Host ""

& $venvPy poc.py --attach $port
