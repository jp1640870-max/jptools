@echo off
title JP Tools - PowerShell Toolkit
cd /d "%~dp0"
echo ========================================
echo  JP Tools - PowerShell Toolkit
echo  Cargando menu de herramientas...
echo ========================================
timeout /t 2 /nobreak >nul
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\bootstrap.ps1"
pause
