#Requires -RunAsAdministrator

$ModulosRAM = Get-CimInstance -ClassName Win32_PhysicalMemory
$RAM_Total_GB = [Math]::Round(($ModulosRAM | Measure-Object -Property Capacity -Sum).Sum / 1GB, 0)
$CPU = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1

$DiscosFisicos = Get-PhysicalDisk | Where-Object { $_.BusType -ne 'USB' }

$InfoDiscos = ""
foreach ($Disco in $DiscosFisicos) {
    $Tipo = switch ($Disco.MediaType) {
        "SSD" { "SSD" }
        "HDD" { "HDD" }
        default { "Desconocido" }
    }

    $Particiones = $Disco | Get-Partition | Where-Object { $_.DriveLetter }

    foreach ($Particion in $Particiones) {
        $Volumen = $Particion | Get-Volume
        $Total_GB = [Math]::Round($Volumen.Size / 1GB, 1)
        $Libre_GB = [Math]::Round($Volumen.SizeRemaining / 1GB, 1)
        $InfoDiscos += "  $($Particion.DriveLetter):`t$Tipo`t$Total_GB GB (Libre: $Libre_GB GB)`n"
    }
}

if ([string]::IsNullOrWhiteSpace($InfoDiscos)) {
    $Discos = Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
    foreach ($Disco in $Discos) {
        $Total_GB = [Math]::Round($Disco.Size / 1GB, 1)
        $Libre_GB = [Math]::Round($Disco.FreeSpace / 1GB, 1)
        $InfoDiscos += "  $($Disco.DeviceID):`tDesconocido`t$Total_GB GB (Libre: $Libre_GB GB)`n"
    }
}

function Mostrar-InfoEquipo {
    Write-Host "`n  INFORMACIÓN DEL EQUIPO" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    $Sistema = Get-CimInstance -ClassName Win32_ComputerSystem
    $BIOS = Get-CimInstance -ClassName Win32_BIOS
    $SO = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -First 1
    Write-Host "  Fabricante:  $($Sistema.Manufacturer)" -ForegroundColor White
    Write-Host "  Modelo:      $($Sistema.Model)" -ForegroundColor White
    Write-Host "  Serial:      $($BIOS.SerialNumber)" -ForegroundColor White
    Write-Host "  SO:          $($SO.Caption)" -ForegroundColor White
    Write-Host "  Nombre:      $($Sistema.Name)" -ForegroundColor White
    Write-Host "  PROCESADOR:  $($CPU.Name)" -ForegroundColor White
    Write-Host "  RAM TOTAL:   $RAM_Total_GB GB" -ForegroundColor White
    Write-Host "  DISCO DURO:  $InfoDiscos" -ForegroundColor White
    Write-Host "========================================`n" -ForegroundColor Cyan
}

function Desactivar-IPv6 {
    Write-Host "`n  DESACTIVAR IPv6" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    try {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters" `
            -Name "DisabledComponents" -Value 0xFF -ErrorAction Stop
        Get-NetAdapterBinding -ComponentID ms_tcpip6 -ErrorAction Stop | ForEach-Object {
            Disable-NetAdapterBinding -Name $_.Name -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue
        }
        $Valor = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters" -Name "DisabledComponents" -ErrorAction Stop
        if ($Valor.DisabledComponents -eq 0xFF) {
            Write-Host "[OK] IPv6 desactivado correctamente (registro: 0xFF)" -ForegroundColor Green
        } else {
            Write-Host "[!] Registro modificado pero valor no confirmado: $($Valor.DisabledComponents)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "[ERROR] No se pudo desactivar IPv6: $_" -ForegroundColor Red
    }
}

function Ejecutar-SFC {
    Write-Host "`n  SFC /SCANNOW" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "[*] Iniciando SFC en ventana separada..." -ForegroundColor Yellow
    Write-Host "    -> El escaneo se ejecutará en paralelo.`n" -ForegroundColor Gray
    Start-Process "sfc" -ArgumentList "/scannow"
}

function Ejecutar-CSSEmerg11005 {
    Write-Host "`n  CSSEmerg11005 - Diagnóstico de Red" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    $URL = "https://aka.ms/diag_cssemerg11005"
    $Archivo = Join-Path $env:TEMP "CSSEmerg11005.diagcab"
    try {
        Write-Host "[*] Descargando CSSEmerg11005.diagcab..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $URL -OutFile $Archivo -UseBasicParsing -ErrorAction Stop
        if ((Test-Path $Archivo) -and ((Get-Item $Archivo).Length -gt 0)) {
            Write-Host "[OK] Descarga completada.`n" -ForegroundColor Green
            Write-Host "[*] Abriendo diagnóstico de red..." -ForegroundColor Yellow
            Write-Host "    -> Sigue las instrucciones en pantalla.`n" -ForegroundColor Gray
            Start-Process "msdt.exe" -ArgumentList "/cab `"$Archivo`"" -Wait
            Write-Host "[OK] Diagnóstico finalizado." -ForegroundColor Green
        } else {
            Write-Host "[ERROR] El archivo descargado está vacío." -ForegroundColor Red
        }
    } catch {
        Write-Host "[ERROR] No se pudo descargar CSSEmerg11005: $_" -ForegroundColor Red
        Write-Host "    URL: $URL" -ForegroundColor Gray
        Write-Host "    Descárgalo manualmente y ejecuta: msdt.exe /cab CSSEmerg11005.diagcab" -ForegroundColor Gray
    }
}


Write-Host "`n╔════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║   HERRAMIENTAS DE REVISIÓN DE EQUIPO   ║" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════╝" -ForegroundColor Green

# 1. Información del equipo
Mostrar-InfoEquipo

# 2. Desactivar IPv6
Desactivar-IPv6

# 3. SFC en paralelo
Ejecutar-SFC

Start-Sleep -Seconds 2

# 4. Diagnóstico de red con CSSEmerg11005
Ejecutar-CSSEmerg11005

Start-Sleep -Seconds 2

# 5. Abrir herramientas de revisión
Write-Host "`n[*] Abriendo herramientas de revisión..." -ForegroundColor Cyan

# Información del sistema
Write-Host "  [+] Abriendo Información del Sistema..." -ForegroundColor White
Start-Process "msinfo32"

# Windows Update
Write-Host "  [+] Abriendo Windows Update..." -ForegroundColor White
Start-Process "ms-settings:windowsupdate"

# Panel de control
Write-Host "  [+] Abriendo Panel de Control..." -ForegroundColor White
Start-Process "control"

# Task Manager
Write-Host "  [+] Abriendo Administrador de Tareas..." -ForegroundColor White
Start-Process "taskmgr"

# GLPI
$GLPI = "C:\Program Files\GLPI-Agent\GLPI-AgentMonitor-x64.exe"
if (Test-Path $GLPI) {
    Write-Host "  [+] Abriendo GLPI Agent Monitor..." -ForegroundColor White
    Start-Process $GLPI
} else {
    Write-Host "  [-] GLPI Monitor no encontrado" -ForegroundColor Yellow
}

# Speedtest
Write-Host "  [+] Abriendo Speedtest (fast.com)..." -ForegroundColor White
Start-Process "https://fast.com/es/"

Write-Host "`n[OK] Todas las herramientas se abrieron correctamente.`n" -ForegroundColor Green
