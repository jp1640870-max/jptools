$rawBase = "https://raw.githubusercontent.com/jp1640870-max/jptools/master/scripts/"

$opciones = @{
    "1" = @{ nombre = "Inventario de hardware (copiar al portapapeles)"; archivo = "Get-HardwareInventory.ps1"; elevado = $false }
    "2" = @{ nombre = "Revisión CallCenter (herramientas + diagnostico)"; archivo = "Invoke-CallCenterReview.ps1"; elevado = $false }
    "3" = @{ nombre = "Instalar GLPI Agent"; archivo = "Install-GLPI-Agent.ps1"; elevado = $true }
    "4" = @{ nombre = "Instalar OCS Agent"; archivo = "Install-OCS-Agent.ps1"; elevado = $true }
    "5" = @{ nombre = "Diagnóstico completo de equipo (12 categorías)"; archivo = "Invoke-ComputerDiagnostic.ps1"; elevado = $false }
}

do {
    Clear-Host
    Write-Host "       ______  ______            __     " -ForegroundColor Green
    Write-Host "      / / __ \/_  __/___  ____  / /____" -ForegroundColor Green
    Write-Host " __  / / /_/ / / / / __ \/ __ \/ / ___/" -ForegroundColor Green
    Write-Host "/ /_/ / ____/ / / / /_/ / /_/ / (__  ) " -ForegroundColor Green
    Write-Host "\____/_/     /_/  \____/\____/_/____/  " -ForegroundColor Green
    Write-Host ""
    Write-Host "               [ PowerShell Toolkit ]" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "=== HERRAMIENTAS DE SOPORTE ===" -ForegroundColor Yellow
    Write-Host "1. Inventario de hardware (copiar al portapapeles)" -ForegroundColor White
    Write-Host "2. Revisión CallCenter (herramientas + diagnostico)" -ForegroundColor White
    Write-Host "3. Instalar GLPI Agent" -ForegroundColor White
    Write-Host "4. Instalar OCS Agent" -ForegroundColor White
    Write-Host "5. Diagnóstico completo de equipo (salud, rendimiento, seguridad)" -ForegroundColor Cyan
    Write-Host "6. Salir" -ForegroundColor White
    Write-Host ""
    $opcion = Read-Host "Selecciona una opcion"

    if ($opciones.ContainsKey($opcion)) {
        $opt = $opciones[$opcion]
        Write-Host "`nEjecutando: $($opt.nombre)..." -ForegroundColor Yellow

        if ($opt.elevado) {
            $scriptPath = Join-Path $PSScriptRoot $opt.archivo
            if (-not (Test-Path $scriptPath)) {
                Write-Host "[ERROR] No se encontro el script local: $scriptPath" -ForegroundColor Red
                Write-Host "Presiona cualquier tecla para volver al menu..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                continue
            }
            Start-Process powershell -Verb RunAs "-ExecutionPolicy Bypass -NoExit -File `"$scriptPath`""
        } else {
            $url = $rawBase + $opt.archivo
            try {
                $script = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction Stop
                iex $script.Content
            } catch {
                Write-Host "[ERROR] No se pudo descargar el script: $_" -ForegroundColor Red
                Write-Host "Verifica la conexion a internet y la URL: $url" -ForegroundColor Gray
            }
        }

        Write-Host "`nPresiona cualquier tecla para volver al menu..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } elseif ($opcion -eq "6") {
        Write-Host "Saliendo..." -ForegroundColor Green
    } else {
        Write-Host "Opcion invalida. Intenta de nuevo." -ForegroundColor Red
        Start-Sleep -Seconds 1
    }
} while ($opcion -ne "6")
