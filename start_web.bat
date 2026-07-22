@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PY_CMD=.venv\Scripts\python.exe"
if not exist "%PY_CMD%" (
    echo [ERROR] .venv not found. Run start_dfv.bat once to create it.
    goto :fail
)

rem --- Detect this computer's LAN IPv4 (the address other people use to reach you) ---
set "HOST_IP="
powershell -NoProfile -Command "$c=Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -ne $null -and $_.NetAdapter.Status -eq 'Up' } | Select-Object -First 1; if($c){ $c.IPv4Address.IPAddress } else { (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*' } | Select-Object -First 1).IPAddress }" > "%TEMP%\_dfv_ip.txt" 2>nul
set /p HOST_IP=<"%TEMP%\_dfv_ip.txt"
del "%TEMP%\_dfv_ip.txt" >nul 2>&1
if "%HOST_IP%"=="" set "HOST_IP=localhost"

rem --- Kill any stale server still holding port 8060 (avoids multiple ---
rem --- instances serving old code after the window was closed uncleanly). ---
echo [INFO] Freeing port 8060 (stopping any previous DFV server)...
powershell -NoProfile -Command "$pids = Get-NetTCPConnection -LocalPort 8060 -State Listen -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique; foreach ($procId in $pids) { try { Stop-Process -Id $procId -Force -ErrorAction Stop; Write-Host ('[INFO] Stopped old server PID ' + $procId) } catch { Write-Host ('[WARN] Could not stop PID ' + $procId + ' - ' + $_.Exception.Message) } }"

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
