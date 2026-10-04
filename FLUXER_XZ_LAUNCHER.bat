@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
title FLUXER XZ - PUBLIC RELEASE V26-0.1

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

set "PS_SCRIPT=%SCRIPT_DIR%FLUXER_XZ_LAUNCHER.ps1"
if not exist "%PS_SCRIPT%" (
    if exist "%SCRIPT_DIR%setup_fluxer_launcher.ps1" (
        set "PS_SCRIPT=%SCRIPT_DIR%setup_fluxer_launcher.ps1"
    ) else if exist "%SCRIPT_DIR%FLUXER_XZ\FLUXER_XZ_LAUNCHER.ps1" (
        set "PS_SCRIPT=%SCRIPT_DIR%FLUXER_XZ\FLUXER_XZ_LAUNCHER.ps1"
    ) else if exist "%SCRIPT_DIR%Systems\FLUXER_XZ_LAUNCHER.ps1" (
        set "PS_SCRIPT=%SCRIPT_DIR%Systems\FLUXER_XZ_LAUNCHER.ps1"
    ) else (
        echo [ERROR] No se encontro FLUXER_XZ_LAUNCHER.ps1 en "%SCRIPT_DIR%"
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
