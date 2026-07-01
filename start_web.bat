@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PY_CMD=.venv\Scripts\python.exe"
if not exist "%PY_CMD%" (
    echo [ERROR] .venv not found. Run start_dfv.bat once to create it.
    goto :fail
)

echo [INFO] Starting DFV web app on http://localhost:8000 ...
"%PY_CMD%" dfv_tool\app.py
if errorlevel 1 goto :fail
exit /b 0

:fail
echo [INFO] Press any key to close.
pause >nul
exit /b 1
