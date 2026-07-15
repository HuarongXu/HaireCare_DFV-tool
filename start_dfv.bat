@echo off
setlocal EnableExtensions

cd /d "%~dp0"

set "PY_CMD="
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" --version >nul 2>&1
    if not errorlevel 1 set "PY_CMD=.venv\Scripts\python.exe"
)

if "%PY_CMD%"=="" (
    echo [INFO] Creating virtual environment...
    where py >nul 2>&1
    if not errorlevel 1 (
        py -3 -m venv .venv
    ) else (
        python -m venv .venv
    )
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        goto :fail
    )
    set "PY_CMD=.venv\Scripts\python.exe"
)

echo [INFO] Installing/updating dependencies...
"%PY_CMD%" -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Dependency installation failed.
    goto :fail
)

echo [INFO] Running full DFV workflow...
"%PY_CMD%" dfv_tool\run.py
if errorlevel 1 (
    echo [WARN] Full AO run failed. Trying pipeline with latest CSV...
    "%PY_CMD%" dfv_tool\pipeline.py
    if errorlevel 1 (
        echo [ERROR] Pipeline fallback failed.
        goto :fail
    )
)

rem --- Detect this computer's LAN IPv4 (the address other people use to reach you) ---
set "HOST_IP="
powershell -NoProfile -Command "$c=Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -ne $null -and $_.NetAdapter.Status -eq 'Up' } | Select-Object -First 1; if($c){ $c.IPv4Address.IPAddress } else { (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*' } | Select-Object -First 1).IPAddress }" > "%TEMP%\_dfv_ip.txt" 2>nul
set /p HOST_IP=<"%TEMP%\_dfv_ip.txt"
del "%TEMP%\_dfv_ip.txt" >nul 2>&1
if "%HOST_IP%"=="" set "HOST_IP=localhost"

echo.
echo [INFO] Starting DFV web app...
echo [INFO] This computer's address:  http://%HOST_IP%:8060
echo [INFO] Share that address with others on the same network to let them log in.
echo [INFO] KEEP THIS WINDOW OPEN - closing it stops the web app.
echo.

rem Open the local browser after a short delay so the server is ready.
start "" powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 3; Start-Process 'http://%HOST_IP%:8060'"

"%PY_CMD%" dfv_tool\app.py
if errorlevel 1 goto :fail
exit /b 0

:fail
echo [INFO] Press any key to close.
pause >nul
exit /b 1
