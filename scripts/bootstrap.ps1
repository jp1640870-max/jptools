$rawBase = "https://raw.githubusercontent.com/jp1640870-max/jptools/main/scripts/"

$opciones = @{
    "1" = @{ nombre = "Inventario de hardware (copiar al portapapeles)"; archivo = "Get-HardwareInventory.ps1" }
    "2" = @{ nombre = "Revisión CallCenter (herramientas + diagnostico)"; archivo = "Invoke-CallCenterReview.ps1" }
    "3" = @{ nombre = "Instalar GLPI Agent"; archivo = "Install-GLPI-Agent.ps1" }
}

do {
    Clear-Host
    Write-Host "██╗██████╗ ████████╗ ██████╗  ██████╗ ██╗     ███████╗" -ForegroundColor Green
    Write-Host "     ██║██╔══██╗╚══██╔══╝██╔═══██╗██╔═══██╗██║     ██╔════╝" -ForegroundColor Green
    Write-Host "     ██║██████╔╝   ██║   ██║   ██║██║   ██║██║     ███████╗" -ForegroundColor Green
    Write-Host "██   ██║██╔═══╝    ██║   ██║   ██║██║   ██║██║     ╚════██║" -ForegroundColor Green
    Write-Host "╚█████╔╝██║        ██║   ╚██████╔╝╚██████╔╝███████╗███████║" -ForegroundColor Green
    Write-Host " ╚════╝ ╚═╝        ╚═╝    ╚═════╝  ╚═════╝ ╚══════╝╚══════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "               [ PowerShell Toolkit ]" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "=== HERRAMIENTAS DE SOPORTE ===" -ForegroundColor Yellow
    Write-Host "1. Inventario de hardware (copiar al portapapeles)" -ForegroundColor White
    Write-Host "2. Revisión CallCenter (herramientas + diagnostico)" -ForegroundColor White
    Write-Host "3. Instalar GLPI Agent" -ForegroundColor White
    Write-Host "4. Salir" -ForegroundColor White
    Write-Host ""
    $opcion = Read-Host "Selecciona una opcion"

    if ($opciones.ContainsKey($opcion)) {
        $url = $rawBase + $opciones[$opcion].archivo
        Write-Host "`nEjecutando: $($opciones[$opcion].nombre)..." -ForegroundColor Yellow
        try {
            $script = Invoke-WebRequest -Uri $url -UseBasicParsing -ErrorAction Stop
            iex $script.Content
        } catch {
            Write-Host "[ERROR] No se pudo descargar el script: $_" -ForegroundColor Red
            Write-Host "Verifica la conexion a internet y la URL: $url" -ForegroundColor Gray
        }
        Write-Host "`nPresiona cualquier tecla para volver al menu..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } elseif ($opcion -eq "4") {
        Write-Host "Saliendo..." -ForegroundColor Green
    } else {
        Write-Host "Opcion invalida. Intenta de nuevo." -ForegroundColor Red
        Start-Sleep -Seconds 1
    }
} while ($opcion -ne "4")
