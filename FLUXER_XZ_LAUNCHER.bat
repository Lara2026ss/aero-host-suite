@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
title FLUXER XZ - PUBLIC RELEASE V26-0.1

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

rem Si se pasa --cli o --console, forzar el lanzador de consola PowerShell
if /i "%1"=="--cli" goto :run_cli
if /i "%1"=="-cli" goto :run_cli
if /i "%1"=="--console" goto :run_cli
if /i "%1"=="-console" goto :run_cli

rem Verificar si existe el ejecutable nativo de la aplicacion de escritorio
if exist "%SCRIPT_DIR%FluxerXZ.exe" (
    start "" "%SCRIPT_DIR%FluxerXZ.exe" %*
    exit /b 0
)
if exist "%SCRIPT_DIR%FLUXER_XZ\FluxerXZ.exe" (
    start "" "%SCRIPT_DIR%FLUXER_XZ\FluxerXZ.exe" %*
    exit /b 0
)
if exist "%SCRIPT_DIR%Systems\FluxerXZ.exe" (
    start "" "%SCRIPT_DIR%Systems\FluxerXZ.exe" %*
    exit /b 0
)

:run_cli
set "PS_SCRIPT=%SCRIPT_DIR%FLUXER_XZ_LAUNCHER.ps1"
if not exist "%PS_SCRIPT%" (
    if exist "%SCRIPT_DIR%setup_fluxer_launcher.ps1" (
        set "PS_SCRIPT=%SCRIPT_DIR%setup_fluxer_launcher.ps1"
    ) else if exist "%SCRIPT_DIR%FLUXER_XZ\FLUXER_XZ_LAUNCHER.ps1" (
        set "PS_SCRIPT=%SCRIPT_DIR%FLUXER_XZ\FLUXER_XZ_LAUNCHER.ps1"
    ) else if exist "%SCRIPT_DIR%Systems\FLUXER_XZ_LAUNCHER.ps1" (
        set "PS_SCRIPT=%SCRIPT_DIR%Systems\FLUXER_XZ_LAUNCHER.ps1"
    ) else (
        echo [ERROR] No se encontro FluxerXZ.exe ni FLUXER_XZ_LAUNCHER.ps1 en "%SCRIPT_DIR%"
        echo Verifique que los archivos del paquete de Fluxer XZ esten presentes.
        pause
        exit /b 1
    )
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" %*

if %ERRORLEVEL% neq 0 (
    echo.
    echo [!] Finalizado con codigo de estado: %ERRORLEVEL%
    pause
)
