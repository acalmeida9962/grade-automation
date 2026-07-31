@echo off
REM Doble clic aqui para iniciar la herramienta (modo Edge + attach).
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -File ".\start-edge.ps1"
echo.
pause
