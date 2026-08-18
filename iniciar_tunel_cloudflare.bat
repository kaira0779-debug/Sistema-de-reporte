@echo off
title Tunel Cloudflare - Acceso Remoto Seguro
cd /d "%~dp0\.."

echo ========================================================
echo  INICIANDO TUNEL CLOUDFLARE PARA EL PUERTO 5000
echo ========================================================
echo.

where cloudflared >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    if not exist "%CD%\cloudflared.exe" (
        echo Descargando cloudflared portable...
        powershell -Command "Invoke-WebRequest -Uri https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe -OutFile cloudflared.exe"
    )
    "%CD%\cloudflared.exe" tunnel --url http://127.0.0.1:5000
) else (
    cloudflared tunnel --url http://127.0.0.1:5000
)

pause