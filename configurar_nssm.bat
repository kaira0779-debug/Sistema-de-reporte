@echo off
:: Ejecutar como Administrador
title Configurando Servicio NSSM - Sistema de Reportes
cd /d "%~dp0\.."

set SERVICE_NAME=SistemaReportesDatinvoz
set PROJECT_DIR=%CD%
set PYTHON_EXE=%PROJECT_DIR%\.venv\Scripts\python.exe
set SCRIPT_RUN=%PROJECT_DIR%\run.py

if not exist "%PYTHON_EXE%" (
    set PYTHON_EXE=python.exe
)

echo Deteniendo y eliminando servicio previo si existe...
nssm stop %SERVICE_NAME% >nul 2>&1
nssm remove %SERVICE_NAME% confirm >nul 2>&1

if not exist "%PROJECT_DIR%\logs" mkdir "%PROJECT_DIR%\logs"

echo Instalando nuevo servicio %SERVICE_NAME%...
nssm install %SERVICE_NAME% "%PYTHON_EXE%" "%SCRIPT_RUN%"

echo Configurando directorio de trabajo...
nssm set %SERVICE_NAME% AppDirectory "%PROJECT_DIR%"

echo Configurando logs de salida y error...
nssm set %SERVICE_NAME% AppStdout "%PROJECT_DIR%\logs\servicio_output.log"
nssm set %SERVICE_NAME% AppStderr "%PROJECT_DIR%\logs\servicio_error.log"

echo Configurando reinicio automatico...
nssm set %SERVICE_NAME% AppExit Default Restart
nssm set %SERVICE_NAME% AppRestartDelay 5000

echo Configurando dependencia con servicio MySQL...
nssm set %SERVICE_NAME% DependOnService MySQL >nul 2>&1
nssm set %SERVICE_NAME% DependOnService mysql >nul 2>&1

echo Iniciando servicio...
nssm start %SERVICE_NAME%

echo.
echo ========================================================
echo  SERVICIO CONFIGURADO E INICIADO EXITOSAMENTE!
echo  Revisa los logs en: %PROJECT_DIR%\logs\
echo ========================================================
pause