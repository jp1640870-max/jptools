@echo off
title JP Tools - Diagnostico Completo de Equipo
cd /d "%~dp0"
echo ========================================
echo  JP Tools - Diagnostico de Equipo
echo  Ejecutando como Administrador...
echo ========================================
timeout /t 2 /nobreak >nul
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\Invoke-ComputerDiagnostic.ps1"
pause
