@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%bootstrap_and_run.ps1"
if errorlevel 1 (
  echo.
  echo Ocurrio un error durante la preparacion o ejecucion.
  pause
)
endlocal
