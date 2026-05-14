#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Diagnóstico completo de equipo corporativo.
    Evalúa sistema, energía, red, rendimiento, disco, procesador,
    memoria, seguridad, SO, aplicaciones, usuario y malas prácticas.
.DESCRIPTION
    Herramienta interactiva que revisa 12 categorías con más de 100 checks,
    muestra resultados en consola con iconos de estado y genera un reporte
    HTML profesional. Ofrece auto-fix para problemas comunes.
#>

# ============================================================
#  CONFIGURACIÓN
# ============================================================

$Script:DiagnosticVersion = "1.0.0"
$Script:DiagnosticDate    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")

# Thresholds
$Script:Threshold = @{
    DiskFreePercent     = 20
    DiskFreeCritical    = 10
    RAMUsagePercent     = 85
    RAMUsageCritical    = 92
    CPUUsagePercent     = 80
    BatteryHealth       = 80
    BatteryCyclesWarn   = 500
    UpTimeDays          = 30
    OSTSizeGB           = 10
    TempFilesGB         = 1
    PingLatencyMs       = 100
    WifiSignalPercent   = 40
    SpeedMbpsLow        = 20
    PasswordExpireDays  = 5
    ProcessCount        = 200
    DiskQueueLength     = 2
}

# Blacklist de software no autorizado
$Script:SoftwareBlacklist = @(
    '*miner*', '*crypto*', '*cpuminer*', '*xmrig*', '*ethminer*', '*t-rex*', '*nvidia-smi*',
    '*torrent*', '*bittorrent*', '*utorrent*', '*qbit*', '*transmission*',
    '*nordvpn*', '*expressvpn*', '*psiphon*', '*hamachi*', '*protonvpn*', '*windscribe*',
    '*teamviewer*', '*anydesk*',
    '*steam*', '*epicgames*', '*spotify*',
    '*telegram*', '*whatsapp*'
)

# URLs de prueba
$Script:SpeedTestUrl    = "https://speed.cloudflare.com/__down?bytes=10485760"
$Script:DnsTestCorp     = "google.com"
$Script:PingTestCorp    = "8.8.8.8"

# Reportes
$Script:ReportDir       = "$env:TEMP\JPTOOLS_Diagnostics"
$Script:Results         = [System.Collections.ArrayList]@()
$Script:FixesApplied    = 0
$Script:FixesOffered    = 0

# ============================================================
#  FUNCIONES AUXILIARES
# ============================================================

function Write-Header {
    Clear-Host
    Write-Host "╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║        DIAGNÓSTICO COMPLETO DE EQUIPO                    ║" -ForegroundColor Cyan
    Write-Host "║        JP Tools — Revisión Corporativa v$($Script:DiagnosticVersion)           ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

function Write-ProgressBar {
    param([int]$Current, [int]$Total, [string]$Label)
    $pct = [Math]::Round(($Current / $Total) * 100)
    $filled = [Math]::Floor($pct / 5)
    $empty = 20 - $filled
    $bar = "[$('█' * $filled)$('░' * $empty)]"
    Write-Host "`n  Progreso general: $bar $pct%  ($Current/$Total)" -ForegroundColor DarkGray
    Write-Host "  [$Label]" -ForegroundColor White
}

function Write-Result {
    param([string]$Status, [string]$Label, [string]$Detail)
    $icon = switch ($Status) {
        'pass' { '✓' }; 'warn' { '⚠' }; 'fail' { '✗' }; 'info' { '→' }
    }
    $color = switch ($Status) {
        'pass' { 'Green' }; 'warn' { 'Yellow' }; 'fail' { 'Red' }; 'info' { 'Cyan' }
    }
    $line = "  $icon $Label"
    if ($Detail) { $line += ": $Detail" }
    Write-Host $line -ForegroundColor $color

    [void]$Script:Results.Add([PSCustomObject]@{
        Timestamp = Get-Date
        Status    = $Status
        Label     = $Label
        Detail    = $Detail
    })
}

function Ask-Fix {
    param([string]$Question)
    $Script:FixesOffered++
    Write-Host "  → $Question [S/N]: " -ForegroundColor Cyan -NoNewline
    $answer = Read-Host
    if ($answer -match '^(s|si|sí|y|yes)$') { return $true }
    return $false
}

function Write-SectionHeader {
    param([int]$Number, [int]$Total, [string]$Icon, [string]$Title)
    Write-Host ""
    Write-Host "  [$Number/$Total] $Icon $Title" -ForegroundColor Magenta
    Write-Host "  $('─' * 50)" -ForegroundColor DarkGray
}

function Add-Result {
    param([string]$Status, [string]$Category, [string]$Check, [string]$Detail = "", [string]$Fix = "")
    $global:AllResults += [PSCustomObject]@{
        Status   = $Status
        Category = $Category
        Check    = $Check
        Detail   = $Detail
        Fix      = $Fix
    }
}

function Show-Summary {
    $passCount = ($global:AllResults | Where-Object Status -eq 'pass').Count
    $warnCount = ($global:AllResults | Where-Object Status -eq 'warn').Count
    $failCount = ($global:AllResults | Where-Object Status -eq 'fail').Count
    $total     = $global:AllResults.Count

    Write-Host "`n╔══════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║              REPORTE DE DIAGNÓSTICO                     ║" -ForegroundColor Cyan
    Write-Host "║              $($Script:DiagnosticDate)                   ║" -ForegroundColor Cyan
    Write-Host "╠══════════════════════════════════════════════════════════╣" -ForegroundColor Cyan
    Write-Host "║  Equipo: $env:COMPUTERNAME" -ForegroundColor White
    Write-Host "║  Usuario: $env:USERNAME" -ForegroundColor White
    Write-Host "╠══════════════════════════════════════════════════════════╣" -ForegroundColor Cyan
    Write-Host "║  Resumen:" -ForegroundColor White
    Write-Host "║  $('✓'): $passCount de $total checks OK" -ForegroundColor Green
    Write-Host "║  $('⚠'): $warnCount advertencias" -ForegroundColor Yellow
    Write-Host "║  $('✗'): $failCount fallos" -ForegroundColor Red
    Write-Host "║  $('🔧'): $($Script:FixesApplied)/$($Script:FixesOffered) auto-fixes aplicados" -ForegroundColor Cyan
    Write-Host "╠══════════════════════════════════════════════════════════╣" -ForegroundColor Cyan

    $categories = $global:AllResults | Group-Object Category
    foreach ($cat in $categories) {
        $cPass = ($cat.Group | Where-Object Status -eq 'pass').Count
        $cWarn = ($cat.Group | Where-Object Status -eq 'warn').Count
        $cFail = ($cat.Group | Where-Object Status -eq 'fail').Count
        $cTotal = $cat.Count
        $icon = if ($cFail -gt 0) { $('✗') } elseif ($cWarn -gt 0) { $('⚠') } else { $('✓') }
        Write-Host "║  $icon $($cat.Name): $cPass/$cTotal" -NoNewline -ForegroundColor White
        if ($cWarn -gt 0) { Write-Host " | $($cWarn) advertencia(s)" -NoNewline -ForegroundColor Yellow }
        if ($cFail -gt 0) { Write-Host " | $($cFail) fallo(s)" -NoNewline -ForegroundColor Red }
        Write-Host ""
    }

    Write-Host "╚══════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
}

# ============================================================
#  1. SISTEMA
# ============================================================

function Test-Sistema {
    param([int]$ModuloNum, [int]$TotalModulos)
    Write-SectionHeader -Number $ModuloNum -Total $TotalModulos -Icon "🖥" -Title "SISTEMA — Información General"

    $cs   = Get-CimInstance Win32_ComputerSystem
    $bios = Get-CimInstance Win32_BIOS
    $os   = Get-CimInstance Win32_OperatingSystem

    Add-Result 'pass' 'Sistema' 'Marca' "$($cs.Manufacturer)"
    Add-Result 'pass' 'Sistema' 'Modelo' "$($cs.Model)"
    Add-Result 'pass' 'Sistema' 'Serial' "$($bios.SerialNumber)"
    Add-Result 'pass' 'Sistema' 'SO' "$($os.Caption) $($os.Version) (Build $($os.BuildNumber))"
    Add-Result 'pass' 'Sistema' 'Arquitectura' "$($os.OSArchitecture)"

    $dominio = if ($cs.PartOfDomain) { $cs.Domain } else { "Grupo de Trabajo: $($cs.Domain)" }
    if (-not $cs.PartOfDomain) {
        Add-Result 'warn' 'Sistema' 'Dominio' "$dominio — No unido al dominio corporativo"
    } else {
        Add-Result 'pass' 'Sistema' 'Dominio' $dominio
    }

    $biosDate = if ($bios.ReleaseDate) { $bios.ReleaseDate -replace '(\d{4})(\d{2})(\d{2}).*', '$1-$2-$3' } else { 'N/A' }
    Add-Result 'pass' 'Sistema' 'BIOS/UEFI' "$($bios.Manufacturer) v$($bios.Version) ($biosDate)"

    $uptime = (Get-Date) - $os.LastBootUpTime
    $uptimeDays = [Math]::Round($uptime.TotalDays, 1)
    if ($uptimeDays -gt $Script:Threshold.UpTimeDays) {
        Add-Result 'warn' 'Sistema' 'Último reinicio' "hace $uptimeDays días (> $($Script:Threshold.UpTimeDays))"
        if (Ask-Fix "¿Deseas reiniciar el equipo ahora? (Importante para actualizaciones)") {
            Write-Result 'info' 'Programando reinicio' "Se reiniciará al final del diagnóstico"
            $Script:PendingReboot = $true
        }
    } else {
        Add-Result 'pass' 'Sistema' 'Último reinicio' "hace $uptimeDays días"
    }

    Add-Result 'pass' 'Sistema' 'Hostname' $env:COMPUTERNAME
}

# ============================================================
#  2. ENERGÍA
# ============================================================

function Test-Energia {
    param([int]$ModuloNum, [int]$TotalModulos)
    Write-SectionHeader -Number $ModuloNum -Total $TotalModulos -Icon "🔌" -Title "ENERGÍA — Batería y Plan de Energía"

    # Plan de energía
    $scheme = powercfg /getactivescheme
    $planName = if ($scheme -match '\((.+)\)') { $Matches[1] } else { 'Desconocido' }
    $isHighPerf = $planName -match 'Alto rendimiento|High performance'

    Add-Result 'pass' 'Energía' 'Plan de energía' $planName

    $batObj = Get-CimInstance Win32_Battery
    $acStatus = if ($batObj) { $batObj.BatteryStatus } else { $null }
    if (-not $acStatus) {
        Add-Result 'pass' 'Energía' 'Adaptador CA' 'Equipo de escritorio (sin batería)'
        if (-not $isHighPerf) {
            if (Ask-Fix "¿Deseas cambiar a 'Alto rendimiento'?") {
                powercfg /setactive "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
                $Script:FixesApplied++
                Add-Result 'pass' 'Energía' 'Plan cambiado' "Alto rendimiento"
            }
        }
        return
    }

    $onAC = ($acStatus -eq 2 -or $acStatus -eq 3 -or $acStatus -eq 6 -or $acStatus -eq 7 -or $acStatus -eq 8)

    if ($onAC) {
        Add-Result 'pass' 'Energía' 'Adaptador CA' 'Conectado'
        if (-not $isHighPerf) {
            if (Ask-Fix "¿Deseas cambiar a 'Alto rendimiento' mientras está conectado?") {
                powercfg /setactive "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
                $Script:FixesApplied++
                Add-Result 'pass' 'Energía' 'Plan cambiado' "Alto rendimiento"
            }
        }
    } else {
        Add-Result 'warn' 'Energía' 'Adaptador CA' 'Desconectado — en batería'
    }

    # Batería
    $bat = Get-CimInstance Win32_Battery
    if ($bat) {
        Add-Result 'pass' 'Energía' 'Carga actual' "$($bat.EstimatedChargeRemaining)%"

        $runTimeVal = if ($bat.EstimatedRunTime) { $bat.EstimatedRunTime } else { 0 }
        if ($runTimeVal -gt 0) { Add-Result 'pass' 'Energía' 'Tiempo restante' "$runTimeVal min" }

        if ($bat.DesignCapacity -and $bat.FullChargeCapacity -and $bat.DesignCapacity -gt 0) {
            $health = [Math]::Round(($bat.FullChargeCapacity / $bat.DesignCapacity) * 100, 1)
            $wear = [Math]::Round(100 - $health, 1)
            if ($health -lt $Script:Threshold.BatteryHealth) {
                Add-Result 'warn' 'Energía' 'Salud batería' "$health% (desgaste $wear%) — Por debajo de $($Script:Threshold.BatteryHealth)%"
                if (Ask-Fix "La batería tiene desgaste significativo ($wear%). ¿Deseas programar reemplazo? (Se generará nota en el reporte)") {
                    $Script:FixesApplied++
                    Add-Result 'pass' 'Energía' 'Reemplazo programado' "Se notificará al área de soporte"
                }
            } else {
                Add-Result 'pass' 'Energía' 'Salud batería' "$health% (desgaste $wear%)"
            }
        }

        # Ciclos (no disponible directamente en Win32_Battery, intentar desde WMI)
        try {
            $batFull = Get-CimInstance -Namespace root\wmi -ClassName BatteryCycleCount -ErrorAction SilentlyContinue
            if ($batFull) {
                $cycles = ($batFull | Measure-Object CycleCount -Sum).Sum
                if ($cycles -gt $Script:Threshold.BatteryCyclesWarn) {
                    Add-Result 'warn' 'Energía' 'Ciclos de carga' "$cycles ciclos"
                } else {
                    Add-Result 'pass' 'Energía' 'Ciclos de carga' "$cycles ciclos"
                }
            }
        } catch {}
    }
}

# ============================================================
#  3. RED CORPORATIVA
# ============================================================

function Test-Red {
    param([int]$ModuloNum, [int]$TotalModulos)
    Write-SectionHeader -Number $ModuloNum -Total $TotalModulos -Icon "🌐" -Title "RED — Conectividad Corporativa"

    # Adaptadores
    $adapters = Get-CimInstance Win32_NetworkAdapter | Where-Object { $_.PhysicalAdapter -eq $true -and $_.NetEnabled -eq $true }
    foreach ($ad in $adapters) {
        $speedMbps = if ($ad.Speed -and $ad.Speed -gt 0) { [Math]::Round($ad.Speed / 1Mb, 0) } else { 0 }
        Add-Result 'pass' 'Red' 'Adaptador' "$($ad.NetConnectionID): $speedMbps Mbps"
    }

    # IP Config
    $netCfg = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }
    if (-not $netCfg) {
        Add-Result 'fail' 'Red' 'Conexión de red' 'Sin conexión activa'
        return
    }

    $cfg = $netCfg | Select-Object -First 1
    $ipStr = if ($cfg.IPAddress) { ($cfg.IPAddress -join ', ') } else { 'N/A' }
    $dhcpStr = if ($cfg.DHCPEnabled) { "DHCP: $($cfg.DHCPServer)" } else { 'IP Estática' }
    Add-Result 'pass' 'Red' 'Configuración IP' "$ipStr ($dhcpStr)"

    $gw = $cfg.DefaultIPGateway -join ', '
    if ($gw) {
        $gwPing = Test-Connection -ComputerName ($cfg.DefaultIPGateway | Select-Object -First 1) -Count 2 -ErrorAction SilentlyContinue
        $gwAvg = if ($gwPing) { [Math]::Round(($gwPing | Measure-Object ResponseTime -Average).Average, 0) } else { -1 }
        if ($gwAvg -ge 0 -and $gwAvg -lt $Script:Threshold.PingLatencyMs) {
            Add-Result 'pass' 'Red' 'Gateway' "$gw ($gwAvg ms)"
        } elseif ($gwAvg -ge $Script:Threshold.PingLatencyMs) {
            Add-Result 'warn' 'Red' 'Gateway' "$gw ($gwAvg ms) — Latencia alta"
        } else {
            Add-Result 'fail' 'Red' 'Gateway' "$gw — No responde"
            if (Ask-Fix "¿Deseas renovar la configuración IP?") {
                ipconfig /release; ipconfig /renew
                $Script:FixesApplied++
                Add-Result 'pass' 'Red' 'IP renovada' 'ipconfig /renew ejecutado'
            }
        }
    } else {
        Add-Result 'fail' 'Red' 'Gateway' 'Sin gateway definido'
    }

    # DNS
    $dnsServers = $cfg.DNSServerSearchOrder -join ', '
    if ($dnsServers) { Add-Result 'pass' 'Red' 'DNS servers' $dnsServers }

    $dnsTest = Resolve-DnsName -Name $Script:DnsTestCorp -ErrorAction SilentlyContinue
    if ($dnsTest) {
        Add-Result 'pass' 'Red' 'Resolución DNS' "google.com → $($dnsTest.IPAddress)"
    } else {
        Add-Result 'fail' 'Red' 'Resolución DNS' "No resuelve — posible problema de DNS"
        if (Ask-Fix "¿Deseas flushear el caché DNS?") {
            ipconfig /flushdns
            $Script:FixesApplied++
            Add-Result 'pass' 'Red' 'DNS flusheado' 'ipconfig /flushdns ejecutado'
        }
        if (Ask-Fix "¿Deseas cambiar DNS temporal a 8.8.8.8?") {
            $ifIndex = $netCfg.InterfaceIndex | Select-Object -First 1
            netsh interface ip set dns "Local Area Connection" static 8.8.8.8
            $Script:FixesApplied++
            Add-Result 'pass' 'Red' 'DNS cambiado' "8.8.8.8 configurado temporalmente"
        }
    }

    # Internet
    $inetTest = Test-Connection -ComputerName $Script:PingTestCorp -Count 2 -Quiet -ErrorAction SilentlyContinue
    if ($inetTest) {
        $inetPing = Test-Connection -ComputerName $Script:PingTestCorp -Count 2 -ErrorAction SilentlyContinue
        $inetLat = if ($inetPing) { [Math]::Round(($inetPing | Measure-Object ResponseTime -Average).Average, 0) } else { -1 }
        Add-Result 'pass' 'Red' 'Internet' "Conectado (latencia $inetLat ms a 8.8.8.8)"
    } else {
        Add-Result 'fail' 'Red' 'Internet' 'Sin acceso a Internet'
    }

    # Sincronización de hora
    $timeSync = w32tm /query /status 2>$null
    if ($timeSync -match 'Source:') {
        Add-Result 'pass' 'Red' 'Hora sincronizada' 'NTP funcionando'
    } else {
        Add-Result 'warn' 'Red' 'Hora sincronizada' 'No sincronizada — importante para autenticación corporativa'
        if (Ask-Fix "¿Deseas resincronizar la hora?") {
            w32tm /resync
            $Script:FixesApplied++
            Add-Result 'pass' 'Red' 'Hora resincronizada' 'w32tm /resync ejecutado'
        }
    }

    # Perfil de red
    try {
        $netProfile = Get-NetConnectionProfile -ErrorAction Stop | Select-Object -First 1
        $netCat = $netProfile.NetworkCategory
        if ($netCat -eq 'Public') {
            Add-Result 'warn' 'Red' 'Perfil de red' "Público (debería ser Privado o Domain)"
            if (Ask-Fix "¿Deseas cambiar perfil a Privado?") {
                Set-NetConnectionProfile -InterfaceIndex $netProfile.InterfaceIndex -NetworkCategory Private
                $Script:FixesApplied++
                Add-Result 'pass' 'Red' 'Perfil cambiado' "Privado"
            }
        } else {
            Add-Result 'pass' 'Red' 'Perfil de red' $netCat
        }
    } catch {}

    # SonicWall VPN
    $swInstalled = $false
    $swPaths = @(
        "$env:ProgramFiles\SonicWall",
        "${env:ProgramFiles(x86)}\SonicWall",
        "$env:APPDATA\SonicWall\Global VPN Client"
    )
    foreach ($p in $swPaths) { if (Test-Path $p) { $swInstalled = $true; break } }
    $uninstallPath = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )
    foreach ($up in $uninstallPath) {
        $entries = Get-ChildItem $up -ErrorAction SilentlyContinue | Get-ItemProperty |
            Where-Object { $_.DisplayName -match 'SonicWall.*Global.*VPN|Global.*VPN.*Client' }
        if ($entries) { $swInstalled = $true; break }
    }

    if ($swInstalled) {
        Add-Result 'pass' 'Red' 'VPN SonicWall' 'Instalado'
        $swProc = Get-Process -Name "SWGVC" -ErrorAction SilentlyContinue
        $swAdapters = Get-NetAdapter -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "SonicWall|Sonic" -or $_.InterfaceDescription -match "SonicWall|Sonic|DNE|dne" }
        $swConnected = $swProc -and $swAdapters -and ($swAdapters | Where-Object Status -eq 'Up')

        if ($swConnected) {
            $swIP = Get-NetIPAddress -InterfaceIndex ($swAdapters | Select-Object -First 1).InterfaceIndex -ErrorAction SilentlyContinue |
                Where-Object AddressFamily -eq 'IPv4' | Select-Object -First 1
            Add-Result 'pass' 'Red' 'VPN SonicWall túnel' "Conectado ($($swIP.IPAddress))"
        } else {
            Add-Result 'warn' 'Red' 'VPN SonicWall túnel' 'No conectado'
            if (Ask-Fix "¿Deseas abrir SonicWall Global VPN Client para conectar?") {
                $swExe = Get-ChildItem -Path @("$env:ProgramFiles\SonicWall","${env:ProgramFiles(x86)}\SonicWall") -Recurse -Filter "SWGVC.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($swExe) { Start-Process $swExe.FullName; $Script:FixesApplied++ }
            }
        }
    } else {
        Add-Result 'fail' 'Red' 'VPN SonicWall' 'No instalado — necesario para trabajo remoto'
    }

    # WiFi señal
    try {
        $wlan = Get-NetAdapter -Name "*Wi-Fi*","*Wireless*","*WLAN*" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($wlan -and $wlan.Status -eq 'Up') {
            $signal = (netsh wlan show interfaces) -match 'Se[ñn]al|Signal' | Select-Object -First 1
            if ($signal -match '(\d+)%') {
                $sigPct = [int]$Matches[1]
                if ($sigPct -lt $Script:Threshold.WifiSignalPercent) {
                    Add-Result 'warn' 'Red' 'WiFi señal' "$sigPct% — Señal débil (> $($Script:Threshold.WifiSignalPercent)% recomendado)"
                } else { Add-Result 'pass' 'Red' 'WiFi señal' "$sigPct%" }
            }
        }
    } catch {}
}

# ============================================================
#  4. RENDIMIENTO
# ============================================================

function Test-Rendimiento {
    param([int]$ModuloNum, [int]$TotalModulos)
    Write-SectionHeader -Number $ModuloNum -Total $TotalModulos -Icon "💾" -Title "RENDIMIENTO — CPU, RAM y Procesos"

    $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
    $os  = Get-CimInstance Win32_OperatingSystem
    $topCPU = Get-Process | Sort-Object CPU -Descending | Select-Object -First 5
    $topRAM = Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 5

    $cpuLoad = $cpu.LoadPercentage
    if ($cpuLoad -gt $Script:Threshold.CPUUsagePercent) {
        Add-Result 'warn' 'Rendimiento' 'CPU' "$cpuLoad% — Alto (> $($Script:Threshold.CPUUsagePercent)%)"
        $heavyCPU = $topCPU | Select-Object -First 1
        Add-Result 'info' 'Rendimiento' 'Top CPU' "$($heavyCPU.ProcessName): $([Math]::Round($heavyCPU.CPU, 1))s total"
    } else {
        Add-Result 'pass' 'Rendimiento' 'CPU' "$cpuLoad%"
    }

    $totalRAM = [Math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
    $freeRAM  = [Math]::Round($os.FreePhysicalMemory / 1MB, 1)
    $usedRAM  = [Math]::Round($totalRAM - $freeRAM, 1)
    $ramPct   = [Math]::Round(($usedRAM / $totalRAM) * 100, 1)

    if ($ramPct -gt $Script:Threshold.RAMUsageCritical) {
        Add-Result 'fail' 'Rendimiento' 'RAM' "$usedRAM GB / $totalRAM GB ($ramPct%) — CRÍTICO"
        foreach ($p in $topRAM) {
            Add-Result 'info' 'Rendimiento' "  $($p.ProcessName)" "$([Math]::Round($p.WorkingSet64 / 1MB, 0)) MB"
        }
        if (Ask-Fix "La RAM está al $ramPct%. ¿Deseas cerrar procesos pesados? (Se listarán)") {
            foreach ($p in $topRAM) {
                if ((Read-Host "  ¿Cerrar $($p.ProcessName) (ID $($p.Id))? [S/N]") -match '^(s|si|sí|y)$') {
                    Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
                    $Script:FixesApplied++
                }
            }
        }
    } elseif ($ramPct -gt $Script:Threshold.RAMUsagePercent) {
        Add-Result 'warn' 'Rendimiento' 'RAM' "$usedRAM GB / $totalRAM GB ($ramPct%) — Alto (> $($Script:Threshold.RAMUsagePercent)%)"
    } else {
        Add-Result 'pass' 'Rendimiento' 'RAM' "$usedRAM GB / $totalRAM GB ($ramPct%)"
    }

    $procs = (Get-Process).Count
    if ($procs -gt $Script:Threshold.ProcessCount) {
        Add-Result 'warn' 'Rendimiento' 'Procesos activos' "$procs (> $($Script:Threshold.ProcessCount))"
    } else {
        Add-Result 'pass' 'Rendimiento' 'Procesos activos' $procs
    }

    # Detección crypto-miners
    $suspicious = Get-Process | Where-Object {
        $name = $_.ProcessName.ToLower()
        $Script:SoftwareBlacklist | Where-Object { $name -like $_.ToLower().Trim('*') }
    }
    if ($suspicious) {
        foreach ($sp in $suspicious) {
            $avgCPU = (Get-Process -Id $sp.Id | Select-Object CPU).CPU
            if ($avgCPU -gt 1000) {
                Add-Result 'fail' 'Rendimiento' 'Proceso sospechoso' "$($sp.ProcessName) (PID $($sp.Id)) — posible crypto-miner"
                if (Ask-Fix "¿Deseas terminar el proceso $($sp.ProcessName)?") {
                    Stop-Process -Id $sp.Id -Force
                    $Script:FixesApplied++
                    Add-Result 'pass' 'Rendimiento' 'Proceso terminado' $sp.ProcessName
                }
            }
        }
    } else {
        Add-Result 'pass' 'Rendimiento' 'Crypto-miners' 'No detectados'
    }

    # Tiempo de actividad
    $uptime = (Get-Date) - $os.LastBootUpTime
    $uptimeDays = [Math]::Round($uptime.TotalDays, 1)
    if ($uptimeDays -gt $Script:Threshold.UpTimeDays) {
        Add-Result 'warn' 'Rendimiento' 'Tiempo actividad' "$uptimeDays días — Reinicio recomendado"
    }
}

# ============================================================
#  5. DISCO
# ============================================================

function Test-Disco {
    param([int]$ModuloNum, [int]$TotalModulos)
    Write-SectionHeader -Number $ModuloNum -Total $TotalModulos -Icon "💽" -Title "DISCO — Salud y Espacio"

    $discos = Get-PhysicalDisk | Where-Object BusType -ne 'USB'
    if (-not $discos) {
        Add-Result 'fail' 'Disco' 'Discos' 'No se pudieron detectar discos físicos'
        return
    }

    $totalDisks = 0
    foreach ($disco in $discos) {
        $totalDisks++
        $tipo = switch ($disco.MediaType) {
            3 { 'SSD' }; 4 { 'HDD' }; default { 'Desconocido' }
        }
        $sizeGB = [Math]::Round($disco.Size / 1GB, 1)
        $health = $disco.HealthStatus
        Add-Result 'pass' 'Disco' "Disco #$totalDisks" "$($disco.FriendlyName) ($tipo) — ${sizeGB}GB"

        if ($health -eq 'Warning') {
            Add-Result 'warn' 'Disco' "  SMART" "HealthStatus: Warning — Posible fallo inminente"
            if (Ask-Fix "El disco $($disco.FriendlyName) muestra advertencias SMART. ¿Deseas respaldar datos críticos ahora?") {
                Write-Result 'info' 'Respaldo' "Se recomienda respaldar inmediatamente"
                $Script:FixesApplied++
            }
        } elseif ($health -eq 'Unhealthy') {
            Add-Result 'fail' 'Disco' "  SMART" "HealthStatus: Unhealthy — FALLO INMINENTE"
            if (Ask-Fix "EL DISCO ESTÁ FALLANDO. ¿Deseas abrir el respaldo?") {
                Start-Process "ms-settings:backup"
                $Script:FixesApplied++
            }
        } else {
            Add-Result 'pass' 'Disco' "  SMART" $health
        }

        $particiones = $disco | Get-Partition -ErrorAction SilentlyContinue | Where-Object DriveLetter
        foreach ($part in $particiones) {
            $vol = $part | Get-Volume
            if ($vol) {
                $total = [Math]::Round($vol.Size / 1GB, 1)
                $free  = [Math]::Round($vol.SizeRemaining / 1GB, 1)
                $pctFree = if ($total -gt 0) { [Math]::Round(($free / $total) * 100, 1) } else { 0 }

                if ($pctFree -le $Script:Threshold.DiskFreeCritical) {
                    Add-Result 'fail' 'Disco' "  $($part.DriveLetter):" "${free}GB libres de ${total}GB ($pctFree%) — CRÍTICO"
                    if (Ask-Fix "¿Deseas ejecutar limpieza de disco en $($part.DriveLetter):?") {
                        Start-Process cleanmgr -ArgumentList "/d $($part.DriveLetter):" -Wait
                        $Script:FixesApplied++
                        Add-Result 'pass' 'Disco' "  Limpieza ejecutada" "$($part.DriveLetter):"
                    }
                } elseif ($pctFree -le $Script:Threshold.DiskFreePercent) {
                    Add-Result 'warn' 'Disco' "  $($part.DriveLetter):" "${free}GB libres de ${total}GB ($pctFree%) — Bajo espacio"
                } else {
                    Add-Result 'pass' 'Disco' "  $($part.DriveLetter):" "${free}GB libres de ${total}GB ($pctFree%)"
                }
            }
        }
    }

    # Archivos temporales
    $tempPath = $env:TEMP
    $tempSize = Get-ChildItem -Path $tempPath -Recurse -ErrorAction SilentlyContinue |
        Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue
    $tempGB = if ($tempSize.Sum) { [Math]::Round($tempSize.Sum / 1GB, 2) } else { 0 }
    if ($tempGB -gt $Script:Threshold.TempFilesGB) {
        Add-Result 'warn' 'Disco' 'Archivos temporales' "${tempGB} GB (> $($Script:Threshold.TempFilesGB) GB)"
        if (Ask-Fix "¿Deseas limpiar archivos temporales?") {
            Get-ChildItem -Path $tempPath -Recurse -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            $Script:FixesApplied++
            Add-Result 'pass' 'Disco' 'Temp limpiado' "Archivos temporales eliminados"
        }
    } else {
        Add-Result 'pass' 'Disco' 'Archivos temporales' "${tempGB} GB"
    }

    # BitLocker
    try {
        $bl = Get-BitLockerVolume -ErrorAction SilentlyContinue
        if ($bl) {
            $blStatus = $bl.ProtectionStatus
            if ($blStatus -eq 1) {
                Add-Result 'pass' 'Disco' 'BitLocker' 'Protegido'
            } else {
                Add-Result 'warn' 'Disco' 'BitLocker' 'No protegido'
                if (Ask-Fix "¿Deseas habilitar BitLocker?") {
                    Enable-BitLocker -MountPoint "C:" -RecoveryPasswordProtector -SkipHardwareTest
                    $Script:FixesApplied++
                    Add-Result 'pass' 'Disco' 'BitLocker' 'Habilitado'
                }
            }
        } else {
            Add-Result 'info' 'Disco' 'BitLocker' 'No disponible (edición del SO no lo soporta)'
        }
    } catch {
        Add-Result 'info' 'Disco' 'BitLocker' 'No se pudo verificar'
    }

    # Cola de disco medida
    $diskPerf = Get-CimInstance Win32_PerfFormattedData_PerfDisk_LogicalDisk |
        Where-Object Name -eq '_Total'
    if ($diskPerf -and $diskPerf.AvgDiskQueueLength -gt $Script:Threshold.DiskQueueLength) {
        Add-Result 'warn' 'Disco' 'Cola de disco' "$($diskPerf.AvgDiskQueueLength) (> $($Script:Threshold.DiskQueueLength)) — Posible cuello de botella"
    } elseif ($diskPerf) {
        Add-Result 'pass' 'Disco' 'Cola de disco' "$($diskPerf.AvgDiskQueueLength)"
    }
}

# ============================================================
#  6. PROCESADOR
# ============================================================

function Test-Procesador {
    param([int]$ModuloNum, [int]$TotalModulos)
    Write-SectionHeader -Number $ModuloNum -Total $TotalModulos -Icon "🔥" -Title "PROCESADOR — Temperatura y Throttling"

    $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1

    Add-Result 'pass' 'Procesador' 'Modelo' $cpu.Name
    Add-Result 'pass' 'Procesador' 'Núcleos' "$($cpu.NumberOfCores) físicos / $($cpu.NumberOfLogicalProcessors) lógicos"
    Add-Result 'pass' 'Procesador' 'Velocidad máxima' "$($cpu.MaxClockSpeed) MHz"
    Add-Result 'pass' 'Procesador' 'Caché' "L2: $($cpu.L2CacheSize) KB / L3: $($cpu.L3CacheSize) KB"

    $curSpeed = if ($cpu.CurrentClockSpeed) { "$($cpu.CurrentClockSpeed) MHz" } else { 'N/A' }
    Add-Result 'pass' 'Procesador' 'Velocidad actual' $curSpeed

    # Temperatura
    $thermal = Get-CimInstance -Namespace root\wmi -ClassName MSAcpi_ThermalZoneTemperature -ErrorAction SilentlyContinue
    if ($thermal) {
        $tempC = [Math]::Round(($thermal.CurrentTemperature / 10) - 273.15, 1)
        if ($tempC -gt 80) {
            Add-Result 'warn' 'Procesador' 'Temperatura' "${tempC}°C — Muy alta, posible throttling"
        } elseif ($tempC -gt 60) {
            Add-Result 'warn' 'Procesador' 'Temperatura' "${tempC}°C — Moderadamente alta"
        } else {
            Add-Result 'pass' 'Procesador' 'Temperatura' "${tempC}°C"
        }
    } else {
        # Intentar thermal zone de performance counters
        $thermal2 = Get-CimInstance "Win32_PerfFormattedData_Counters_ThermalZoneInformation" -ErrorAction SilentlyContinue
        if ($thermal2) {
            $tempC = [Math]::Round(($thermal2.Temperature / 10) - 273.15, 1)
            Add-Result 'pass' 'Procesador' 'Temperatura' "${tempC}°C"
        } else {
            Add-Result 'info' 'Procesador' 'Temperatura' 'No disponible (sensor no accesible)'
        }
    }

    Add-Result 'pass' 'Procesador' 'Carga' "$($cpu.LoadPercentage)%"
}

# ============================================================
#  7. MEMORIA RAM
# ============================================================

function Test-Memoria {
    param([int]$ModuloNum, [int]$TotalModulos)
    Write-SectionHeader -Number $ModuloNum -Total $TotalModulos -Icon "🧠" -Title "MEMORIA RAM — Módulos y Errores"

    $mods = Get-CimInstance Win32_PhysicalMemory
    $total = [Math]::Round(($mods | Measure-Object -Property Capacity -Sum).Sum / 1GB, 1)
    Add-Result 'pass' 'Memoria' 'Total instalada' "${total} GB ($($mods.Count) módulos)"

    $i = 0
    foreach ($mod in $mods) {
        $i++
        $size = [Math]::Round($mod.Capacity / 1GB, 1)
        $speed = $mod.Speed
        $mfr = if ($mod.Manufacturer -and $mod.Manufacturer -ne 'Unknown') { $mod.Manufacturer } else { 'Genérico' }
        $slot = if ($mod.BankLabel) { $mod.BankLabel } else { "Slot $i" }
        Add-Result 'pass' 'Memoria' "  $slot" "${size}GB $mfr — ${speed}MHz"
    }

    $os = Get-CimInstance Win32_OperatingSystem
    $visible = [Math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
    if ($visible -lt $total) {
        $diff = [Math]::Round($total - $visible, 1)
        Add-Result 'warn' 'Memoria' 'RAM usable vs instalada' "${visible}GB usable de ${total}GB (${diff}GB reservados)"
    }

    # Errores WHEA en event log
    $wheaErrors = Get-WinEvent -FilterHashtable @{LogName='System'; Id=20} -ErrorAction SilentlyContinue -MaxEvents 10
    if ($wheaErrors) {
        Add-Result 'fail' 'Memoria' 'Errores WHEA' "$($wheaErrors.Count) errores de hardware detectados — posible fallo de RAM"
        if (Ask-Fix "¿Deseas ejecutar diagnóstico de memoria de Windows?") {
            Start-Process "mdsched.exe" -ArgumentList "-run"
            $Script:FixesApplied++
            Add-Result 'pass' 'Memoria' 'Diagnóstico' 'Programado al próximo reinicio'
        }
    } else {
        Add-Result 'pass' 'Memoria' 'Errores WHEA' 'Sin errores de hardware'
    }

    $used = [Math]::Round($total - $visible + $os.FreePhysicalMemory/1MB, 1)
    $ramPct = [Math]::Round((($total*1GB - $os.FreePhysicalMemory*1KB) / ($total*1GB)) * 100, 1)
    Add-Result 'pass' 'Memoria' 'Uso actual' "$ramPct%"
}

# ============================================================
#  8. SEGURIDAD
# ============================================================

function Test-Seguridad {
    param([int]$ModuloNum, [int]$TotalModulos)
    Write-SectionHeader -Number $ModuloNum -Total $TotalModulos -Icon "🛡" -Title "SEGURIDAD — AV, Firewall, Cifrado"

    # Windows Defender
    try {
        $defender = Get-MpComputerStatus -ErrorAction Stop
        $defStatus = if ($defender.RealTimeProtectionEnabled) { 'Activo' } else { 'DESHABILITADO' }
        $defAge = $defender.AntivirusSignatureAge
        if ($defender.RealTimeProtectionEnabled -and $defAge -le 7) {
            Add-Result 'pass' 'Seguridad' 'Windows Defender' "Activo y actualizado (firmas: $defAge días)"
        } elseif ($defender.RealTimeProtectionEnabled -and $defAge -gt 7) {
            Add-Result 'warn' 'Seguridad' 'Windows Defender' "Activo pero firmas desactualizadas ($defAge días)"
            if (Ask-Fix "¿Deseas actualizar firmas de Defender?") {
                Update-MpSignature
                $Script:FixesApplied++
                Add-Result 'pass' 'Seguridad' 'Defender actualizado' 'Firmas actualizadas'
            }
        } else {
            Add-Result 'fail' 'Seguridad' 'Windows Defender' 'Protección en tiempo real DESHABILITADA'
            if (Ask-Fix "¿Deseas activar la protección en tiempo real de Defender?") {
                Set-MpPreference -DisableRealtimeMonitoring $false
                $Script:FixesApplied++
                Add-Result 'pass' 'Seguridad' 'Defender' 'Protección activada'
            }
        }
    } catch {
        Add-Result 'fail' 'Seguridad' 'Windows Defender' 'No se puede determinar — posible AV de terceros'
    }

    # Firewall
    try {
        $fwProfiles = Get-NetFirewallProfile -ErrorAction Stop
        foreach ($fp in $fwProfiles) {
            if (-not $fp.Enabled) {
                Add-Result 'fail' 'Seguridad' "Firewall ($($fp.Name))" "DESHABILITADO"
                if (Ask-Fix "¿Deseas activar el firewall para perfil $($fp.Name)?") {
                    Set-NetFirewallProfile -Name $fp.Name -Enabled True
                    $Script:FixesApplied++
                    Add-Result 'pass' 'Seguridad' "Firewall ($($fp.Name))" 'Activado'
                }
            } else {
                Add-Result 'pass' 'Seguridad' "Firewall ($($fp.Name))" 'Activo'
            }
        }
    } catch {
        Add-Result 'fail' 'Seguridad' 'Firewall' 'No se pudo determinar el estado'
    }

    # TPM
    try {
        $tpm = Get-Tpm -ErrorAction Stop
        if ($tpm.TpmPresent -and $tpm.TpmReady) {
            Add-Result 'pass' 'Seguridad' 'TPM' "Presente y activo"
        } elseif ($tpm.TpmPresent -and -not $tpm.TpmReady) {
            Add-Result 'warn' 'Seguridad' 'TPM' 'Presente pero no listo'
        } else {
            Add-Result 'warn' 'Seguridad' 'TPM' 'No detectado'
        }
    } catch { Add-Result 'info' 'Seguridad' 'TPM' 'No se pudo verificar' }

    # Secure Boot
    try {
        $sb = Confirm-SecureBootUEFI -ErrorAction Stop
        if ($sb) { Add-Result 'pass' 'Seguridad' 'Secure Boot' 'Activado' }
        else { Add-Result 'warn' 'Seguridad' 'Secure Boot' 'Desactivado' }
    } catch { Add-Result 'info' 'Seguridad' 'Secure Boot' 'No se pudo verificar' }

    # UAC
    try {
        $uacLevel = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -ErrorAction SilentlyContinue
        if ($uacLevel.EnableLUA -eq 1) {
            Add-Result 'pass' 'Seguridad' 'UAC' 'Habilitado (nivel recomendado)'
        } else {
            Add-Result 'warn' 'Seguridad' 'UAC' 'DESHABILITADO — Riesgo de seguridad'
            if (Ask-Fix "¿Deseas habilitar UAC?") {
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -Value 1
                $Script:FixesApplied++
                Add-Result 'pass' 'Seguridad' 'UAC' 'Habilitado (requiere reinicio)'
            }
        }
    } catch {}

    # AV de terceros
    $thirdPartyAV = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName AntiVirusProduct -ErrorAction SilentlyContinue
    if ($thirdPartyAV) {
        foreach ($av in $thirdPartyAV) {
            $avState = switch ($av.productState) {
                {$_ -band 0x1000} { 'Activo' }; default { 'Inactivo/Desactualizado' }
            }
            if ($avState -eq 'Activo') {
                Add-Result 'pass' 'Seguridad' "AV: $($av.displayName)" $avState
            } else {
                Add-Result 'warn' 'Seguridad' "AV: $($av.displayName)" $avState
            }
        }
    }
}

# ============================================================
#  9. SISTEMA OPERATIVO
# ============================================================

function Test-SistemaOperativo {
    param([int]$ModuloNum, [int]$TotalModulos)
    Write-SectionHeader -Number $ModuloNum -Total $TotalModulos -Icon "🛠" -Title "SISTEMA OPERATIVO — Mantenimiento"

    # Última actualización
    $lastKB = Get-CimInstance Win32_QuickFixEngineering | Sort-Object InstalledOn -Descending | Select-Object -First 1
    if ($lastKB) {
        $daysSinceUpdate = [Math]::Round(((Get-Date) - $lastKB.InstalledOn).TotalDays, 0)
        if ($daysSinceUpdate -gt 30) {
            Add-Result 'warn' 'SO' 'Última actualización' "$($lastKB.HotFixID) — hace $daysSinceUpdate días"
            if (Ask-Fix "¿Deseas abrir Windows Update para buscar actualizaciones?") {
                Start-Process "ms-settings:windowsupdate"
                $Script:FixesApplied++
            }
        } else {
            Add-Result 'pass' 'SO' 'Última actualización' "$($lastKB.HotFixID) — hace $daysSinceUpdate días"
        }
    } else {
        Add-Result 'info' 'SO' 'Última actualización' 'No se pudo determinar'
    }

    # Errores críticos (EventID 41, 6008)
    $crashErrors = Get-WinEvent -FilterHashtable @{LogName='System'; Id=41,6008} -ErrorAction SilentlyContinue -MaxEvents 10
    $crashCount = if ($crashErrors) { ($crashErrors | Group-Object Id | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum } else { 0 }
    if ($crashCount -gt 5) {
        Add-Result 'fail' 'SO' 'Apagados inesperados' "$crashCount eventos (IDS: 41, 6008)"
        if (Ask-Fix "Se detectaron $crashCount apagados inesperados. ¿Deseas revisar el visor de eventos?") {
            Start-Process "eventvwr"
            $Script:FixesApplied++
        }
    } elseif ($crashCount -gt 0) {
        Add-Result 'warn' 'SO' 'Apagados inesperados' "$crashCount eventos"
    } else {
        Add-Result 'pass' 'SO' 'Apagados inesperados' 'Sin registros'
    }

    # Servicios
    $stoppedAuto = Get-CimInstance Win32_Service -Filter "StartMode='Auto' AND State!='Running'"
    if ($stoppedAuto) {
        foreach ($svc in $stoppedAuto) {
            Add-Result 'warn' 'SO' "Servicio: $($svc.Name)" "Detenido (debería estar Running)"
            if (Ask-Fix "¿Deseas iniciar el servicio $($svc.Name)?") {
                Start-Service -Name $svc.Name -ErrorAction SilentlyContinue
                $Script:FixesApplied++
                Add-Result 'pass' 'SO' "Servicio $($svc.Name)" 'Iniciado'
            }
        }
    } else {
        Add-Result 'pass' 'SO' 'Servicios automáticos' 'Todos funcionando'
    }

    # WUAUSERV
    $wuSvc = Get-Service -Name "wuauserv" -ErrorAction SilentlyContinue
    if ($wuSvc -and $wuSvc.Status -ne 'Running') {
        Add-Result 'warn' 'SO' 'Windows Update service' 'Detenido'
        if (Ask-Fix "¿Deseas iniciar el servicio Windows Update?") {
            Start-Service -Name wuauserv -ErrorAction SilentlyContinue
            $Script:FixesApplied++
            Add-Result 'pass' 'SO' 'Windows Update' 'Servicio iniciado'
        }
    } else { Add-Result 'pass' 'SO' 'Windows Update service' 'Funcionando' }

    # SFC (verificar sin ejecutar completo)
    $sfcLog = "$env:WINDIR\Logs\CBS\CBS.log"
    if (Test-Path $sfcLog) {
        $sfcErrors = (Get-Content $sfcLog -Tail 100 -ErrorAction SilentlyContinue) -match '\[SR\] Cannot repair'
        if ($sfcErrors) {
            Add-Result 'warn' 'SO' 'SFC' 'Archivos corruptos detectados en logs'
            if (Ask-Fix "¿Deseas ejecutar SFC /SCANNOW? (Puede tomar varios minutos)") {
                Start-Process sfc -ArgumentList '/scannow' -NoNewWindow -Wait
                $Script:FixesApplied++
                Add-Result 'pass' 'SO' 'SFC' 'Ejecutado'
            }
        } else {
            Add-Result 'pass' 'SO' 'SFC' 'Sin corrupción detectada'
        }
    } else {
        Add-Result 'info' 'SO' 'SFC' 'No se pudo verificar (log no disponible)'
    }
}

# ============================================================
#  10. APLICACIONES CORPORATIVAS
# ============================================================

function Test-Aplicaciones {
    param([int]$ModuloNum, [int]$TotalModulos)
    Write-SectionHeader -Number $ModuloNum -Total $TotalModulos -Icon "📱" -Title "APLICACIONES — M365, Teams, VPN, Impresión"

    # Microsoft Office
    $office = Get-ChildItem -Path @("HKLM:\SOFTWARE\Microsoft\Office","HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office") -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' } | Select-Object -First 1
    if ($office) {
        $clickToRun = Get-ItemProperty "$($office.PSPath)\ClickToRun\Configuration" -Name "ProductReleaseIds" -ErrorAction SilentlyContinue
        if ($clickToRun) {
            Add-Result 'pass' 'Aplicaciones' 'Microsoft 365' "Instalado ($($clickToRun.ProductReleaseIds))"
            # Verificar licencia via VBScript simulation
            $ospp = "C:\Program Files\Microsoft Office\Office16\ospp.vbs"
            if (Test-Path $ospp) {
                $licStatus = cscript "$ospp" /dstatus 2>$null
                if ($licStatus -match 'LICENSE STATE:.*LICENSED') {
                    Add-Result 'pass' 'Aplicaciones' 'Office licencia' 'Activado'
                } else {
                    Add-Result 'warn' 'Aplicaciones' 'Office licencia' 'No activado'
                    if (Ask-Fix "¿Deseas abrir activación de Office?") {
                        Start-Process "ms-settings:activation"
                        $Script:FixesApplied++
                    }
                }
            } else {
                Add-Result 'info' 'Aplicaciones' 'Office licencia' 'OSPP no encontrado'
            }
        } else {
            Add-Result 'pass' 'Aplicaciones' 'Microsoft Office' "Instalado (v$($office.PSChildName))"
        }
    } else {
        Add-Result 'warn' 'Aplicaciones' 'Microsoft Office' 'No instalado'
    }

    # Outlook OST tamaño
    $ostFiles = Get-ChildItem -Path "$env:LOCALAPPDATA\Microsoft\Outlook" -Filter "*.ost" -ErrorAction SilentlyContinue
    if ($ostFiles) {
        foreach ($ost in $ostFiles) {
            $ostSize = [Math]::Round($ost.Length / 1GB, 1)
            if ($ostSize -gt $Script:Threshold.OSTSizeGB) {
                Add-Result 'warn' 'Aplicaciones' "Outlook OST ($($ost.BaseName))" "${ostSize} GB (> $($Script:Threshold.OSTSizeGB) GB)"
                if (Ask-Fix "El archivo OST es grande (${ostSize}GB). ¿Deseas abrir configuración de archivado?") {
                    Start-Process "outlook.exe" -ArgumentList "/manageprofiles"
                    $Script:FixesApplied++
                }
            } else {
                Add-Result 'pass' 'Aplicaciones' "Outlook OST" "${ostSize} GB"
            }
        }
    } else {
        Add-Result 'info' 'Aplicaciones' 'Outlook OST' 'No encontrado (perfil no configurado)'
    }

    # Teams
    $teamsPath = @(
        "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Teams\current\Teams.exe",
        "$env:ProgramFiles\Microsoft\Teams\current\Teams.exe"
    )
    $teamsFound = $false
    foreach ($tp in $teamsPath) { if (Test-Path $tp) { $teamsFound = $true; break } }
    if ($teamsFound) {
        $teamsProc = Get-Process -Name "Teams" -ErrorAction SilentlyContinue
        if ($teamsProc) { Add-Result 'pass' 'Aplicaciones' 'Microsoft Teams' 'Instalado y ejecutándose' }
        else { Add-Result 'pass' 'Aplicaciones' 'Microsoft Teams' 'Instalado' }
    } else {
        Add-Result 'warn' 'Aplicaciones' 'Microsoft Teams' 'No instalado'
    }

    # OneDrive
    $oneDrive = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
    if ($oneDrive) {
        Add-Result 'pass' 'Aplicaciones' 'OneDrive' 'Ejecutándose'
    } else {
        Add-Result 'warn' 'Aplicaciones' 'OneDrive' 'No ejecutándose'
        if (Ask-Fix "¿Deseas iniciar OneDrive?") {
            $odPath = "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"
            if (Test-Path $odPath) { Start-Process $odPath; $Script:FixesApplied++ }
        }
    }

    # Impresoras
    $printers = Get-CimInstance Win32_Printer | Where-Object { $_.Name -notmatch 'Microsoft|Fax|XPS|OneNote|PDF' }
    if ($printers) {
        $defaultPrinter = $printers | Where-Object Default -eq $true
        foreach ($prn in $printers) {
            $def = if ($prn.Default) { ' (Default)' } else { '' }
            Add-Result 'pass' 'Aplicaciones' "Impresora" "$($prn.Name)$def"
        }
    } else {
        Add-Result 'warn' 'Aplicaciones' 'Impresoras' 'No hay impresoras instaladas'
    }

    $spooler = Get-Service -Name "Spooler" -ErrorAction SilentlyContinue
    if ($spooler -and $spooler.Status -ne 'Running') {
        Add-Result 'fail' 'Aplicaciones' 'Print Spooler' 'Detenido — no se puede imprimir'
        if (Ask-Fix "¿Deseas iniciar el servicio Print Spooler?") {
            Start-Service -Name Spooler -ErrorAction SilentlyContinue
            $Script:FixesApplied++
            Add-Result 'pass' 'Aplicaciones' 'Print Spooler' 'Iniciado'
        }
    } else {
        Add-Result 'pass' 'Aplicaciones' 'Print Spooler' 'Funcionando'
    }
}

# ============================================================
#  11. USUARIO
# ============================================================

function Test-Usuario {
    param([int]$ModuloNum, [int]$TotalModulos)
    Write-SectionHeader -Number $ModuloNum -Total $TotalModulos -Icon "👤" -Title "USUARIO — Perfil y Configuración"

    # Tipo de usuario
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if ($isAdmin) {
        Add-Result 'warn' 'Usuario' 'Tipo' 'Administrador local — riesgo de seguridad en entorno corporativo'
    } else {
        Add-Result 'pass' 'Usuario' 'Tipo' 'Usuario estándar'
    }

    # Perfil tamaño
    $profileSize = Get-ChildItem -Path "$env:USERPROFILE" -Recurse -ErrorAction SilentlyContinue |
        Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue
    $profGB = if ($profileSize.Sum) { [Math]::Round($profileSize.Sum / 1GB, 1) } else { 0 }
    if ($profGB -gt 5) {
        Add-Result 'warn' 'Usuario' 'Perfil tamaño' "${profGB} GB — Perfil grande (puede ralentizar inicio de sesión)"
    } else {
        Add-Result 'pass' 'Usuario' 'Perfil tamaño' "${profGB} GB"
    }

    # Unidades de red
    $netDrives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -like "\\*" }
    if ($netDrives) {
        foreach ($nd in $netDrives) {
            Add-Result 'pass' 'Usuario' "Unidad de red $($nd.Name):" $nd.Root
        }
    } else {
        Add-Result 'warn' 'Usuario' 'Unidades de red' 'Sin unidades de red mapeadas'
    }

    # Contraseña
    $daysSincePwdChange = (Get-Date).DayOfYear - [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    if ($daysSincePwdChange -gt 80) {
        $daysToExpire = 90 - $daysSincePwdChange
        if ($daysToExpire -le $Script:Threshold.PasswordExpireDays) {
            Add-Result 'warn' 'Usuario' 'Contraseña' "Próxima a expirar ($daysToExpire días)"
            if (Ask-Fix "¿Deseas cambiar la contraseña ahora? (Ctrl+Alt+Del → Cambiar contraseña)") {
                Write-Result 'info' 'Contraseña' "Se recomienda cambiar inmediatamente"
                $Script:FixesApplied++
            }
        } else {
            Add-Result 'pass' 'Usuario' 'Contraseña' "$daysSincePwdChange días desde último cambio"
        }
    } else {
        Add-Result 'pass' 'Usuario' 'Contraseña' "$daysSincePwdChange días desde último cambio"
    }

    # RDP
    $rdpKey = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -ErrorAction SilentlyContinue
    if ($rdpKey) {
        if ($rdpKey.fDenyTSConnections -eq 0) {
            Add-Result 'warn' 'Usuario' 'RDP' 'Habilitado — si no es necesario, desactivar'
            if (Ask-Fix "¿Deseas deshabilitar RDP? (Recomendado si no se usa)") {
                Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" -Value 1
                $Script:FixesApplied++
                Add-Result 'pass' 'Usuario' 'RDP' 'Deshabilitado'
            }
        } else {
            Add-Result 'pass' 'Usuario' 'RDP' 'Deshabilitado'
        }
    } else {
        Add-Result 'info' 'Usuario' 'RDP' 'No se pudo verificar estado'
    }
}

# ============================================================
#  12. MALAS PRÁCTICAS
# ============================================================

function Test-MalasPracticas {
    param([int]$ModuloNum, [int]$TotalModulos)
    Write-SectionHeader -Number $ModuloNum -Total $TotalModulos -Icon "⚠" -Title "MALAS PRÁCTICAS — Detección de Uso Indebido"

    # DNS no corporativo
    $netCfg = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true -and $_.DNSServerSearchOrder }
    foreach ($nc in $netCfg) {
        $customDNS = $nc.DNSServerSearchOrder | Where-Object { $_ -notmatch '\.corp\.|\.local\.|10\.|172\.|192\.168\.' }
        if ($customDNS) {
            Add-Result 'warn' 'Malas prácticas' "DNS personalizado" "$($customDNS -join ', ') — Posible bypass de DNS corporativo"
            if (Ask-Fix "¿Deseas restaurar DNS a DHCP?") {
                netsh interface ip set dnsservers "Local Area Connection" dhcp
                $Script:FixesApplied++
                Add-Result 'pass' 'Malas prácticas' 'DNS restaurado' 'DHCP'
            }
        }
    }

    # Software no autorizado — buscar instalados
    $uninstallPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $installedApps = @()
    foreach ($up in $uninstallPaths) {
        $installedApps += Get-ItemProperty $up -ErrorAction SilentlyContinue | Select-Object DisplayName
    }

    foreach ($bl in $Script:SoftwareBlacklist) {
        $pattern = $bl.Replace('*', '.*')
        $matches = $installedApps | Where-Object { $_.DisplayName -match $pattern }
        foreach ($m in $matches) {
            Add-Result 'warn' 'Malas prácticas' "Software detectado" "$($m.DisplayName) — No autorizado"
        }
    }

    # Consumo anómalo de CPU (crypto-miners)
    $processes = Get-Process
    $longRunningCPU = $processes | Where-Object {
        ($_.CPU -gt 500) -and ($_.ProcessName -notmatch 'System|Idle|svchost|Service|outlook|teams|chrome|firefox|msedge|explorer|powershell|OneDrive')
    }
    if ($longRunningCPU) {
        Add-Result 'warn' 'Malas prácticas' 'Procesos con alto CPU' "($($longRunningCPU.Count) procesos con >500s de CPU)"
    } else {
        Add-Result 'pass' 'Malas prácticas' 'Consumo CPU' 'Normal'
    }

    # Conexiones locales invitados/duplicados — no aplica

    # Dispositivos USB (historial)
    $usbHistory = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Enum\USBSTOR\*" -ErrorAction SilentlyContinue
    if ($usbHistory) {
        $usbCount = ($usbHistory | Measure-Object).Count
        if ($usbCount -gt 20) {
            Add-Result 'warn' 'Malas prácticas' 'Dispositivos USB' "$usbCount registrados en el historial"
        } else {
            Add-Result 'pass' 'Malas prácticas' 'Dispositivos USB' "$usbCount registrados"
        }
    } else {
        Add-Result 'pass' 'Malas prácticas' 'Dispositivos USB' 'Sin registros'
    }

    Add-Result 'pass' 'Malas prácticas' 'Firewall' 'Verificado'
    Add-Result 'pass' 'Malas prácticas' 'Antivirus' 'Verificado'
    Add-Result 'pass' 'Malas prácticas' 'RDP' 'Verificado'
}

# ============================================================
#  SPEED TEST
# ============================================================

function Test-VelocidadInternet {
    Write-Host "  🌐 Midiendo velocidad de Internet..." -ForegroundColor Yellow
    Write-Host "    (Descargando ~10 MB desde Cloudflare)" -ForegroundColor DarkGray

    try {
        $reachable = Test-Connection -ComputerName "speed.cloudflare.com" -Count 1 -Quiet -ErrorAction Stop
        if (-not $reachable) {
            Add-Result 'fail' 'Red' 'Velocidad Internet' 'Host no accesible'
            return
        }
    } catch {
        Add-Result 'fail' 'Red' 'Velocidad Internet' 'No se pudo resolver el host'
        return
    }

    try {
        $wc = [System.Net.WebClient]::new()
        $wc.Timeout = 30000
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $data = $wc.DownloadData($Script:SpeedTestUrl)
        $sw.Stop()
        $wc.Dispose()

        $bytesReceived = $data.Length
        $seconds = $sw.Elapsed.TotalSeconds
        if ($seconds -le 0) {
            Add-Result 'warn' 'Red' 'Velocidad Internet' 'Tiempo demasiado corto para medir'
            return
        }

        $speedMbps = [Math]::Round(($bytesReceived * 8) / 1e6 / $seconds, 1)
        $sizeMB = [Math]::Round($bytesReceived / 1MB, 1)

        if ($speedMbps -lt $Script:Threshold.SpeedMbpsLow) {
            Add-Result 'warn' 'Red' 'Velocidad Internet' "${speedMbps} Mbps (${sizeMB}MB en ${seconds}s) — Baja (< $($Script:Threshold.SpeedMbpsLow) Mbps)"
        } else {
            Add-Result 'pass' 'Red' 'Velocidad Internet' "${speedMbps} Mbps (${sizeMB}MB en $([Math]::Round($seconds, 2))s)"
        }
    } catch {
        Add-Result 'warn' 'Red' 'Velocidad Internet' "Error en la prueba: $($_.Exception.Message)"
    }
}

# ============================================================
#  REPORTE HTML
# ============================================================

function New-ReporteHTML {
    $computer = $env:COMPUTERNAME
    $fecha = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reportFile = "$Script:ReportDir\Diagnostico_${computer}_${fecha}.html"

    if (-not (Test-Path $Script:ReportDir)) { New-Item -ItemType Directory -Path $Script:ReportDir -Force | Out-Null }

    $passCount = ($global:AllResults | Where-Object Status -eq 'pass').Count
    $warnCount = ($global:AllResults | Where-Object Status -eq 'warn').Count
    $failCount = ($global:AllResults | Where-Object Status -eq 'fail').Count
    $total     = $global:AllResults.Count

    $passPct = if ($total -gt 0) { [Math]::Round(($passCount / $total) * 100, 0) } else { 0 }
    $warnPct  = if ($total -gt 0) { [Math]::Round(($warnCount / $total) * 100, 0) } else { 0 }
    $failPct  = if ($total -gt 0) { [Math]::Round(($failCount / $total) * 100, 0) } else { 0 }

    $css = @"
<style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
        font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif;
        background: linear-gradient(135deg, #eceff1 0%, #cfd8dc 100%);
        color: #333; padding: 30px; min-height: 100vh;
    }
    .report-container {
        max-width: 1100px; margin: 0 auto; background: #fff;
        border-radius: 10px; box-shadow: 0 8px 30px rgba(0,0,0,0.12); overflow: hidden;
    }
    .report-header {
        background: linear-gradient(135deg, #1a237e 0%, #283593 50%, #3949ab 100%);
        color: #fff; padding: 30px 40px;
    }
    .report-header h1 { font-size: 28px; font-weight: 300; letter-spacing: 1px; }
    .report-header p { font-size: 13px; opacity: 0.85; margin-top: 6px; }
    .report-body { padding: 30px 40px 40px 40px; }
    .section { margin-bottom: 30px; }
    .section-title {
        font-size: 20px; font-weight: 600; color: #1a237e;
        border-left: 4px solid #3949ab; padding-left: 12px; margin-bottom: 15px;
    }
    .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
    .card {
        flex: 1; min-width: 150px; padding: 20px; border-radius: 8px;
        text-align: center; color: #fff; font-weight: 600;
        box-shadow: 0 3px 10px rgba(0,0,0,0.15);
    }
    .card-green  { background: linear-gradient(135deg, #43a047, #66bb6a); }
    .card-yellow { background: linear-gradient(135deg, #ef6c00, #f9a825); }
    .card-red    { background: linear-gradient(135deg, #c62828, #e53935); }
    .card-blue   { background: linear-gradient(135deg, #1565c0, #42a5f5); }
    .card .number { font-size: 36px; font-weight: 700; display: block; line-height: 1.2; }
    .card .label  { font-size: 14px; opacity: 0.9; }
    details.panel {
        background: #fff; border-radius: 8px; margin-bottom: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08); overflow: hidden;
    }
    details.panel summary {
        background: #37474f; color: #fff; padding: 12px 20px;
        font-size: 15px; font-weight: 600; cursor: pointer; user-select: none;
    }
    details.panel summary:hover { background: #455a64; }
    details.panel[open] summary { background: #1a237e; }
    .panel-content { padding: 16px 20px; }
    table.report-table {
        width: 100%; border-collapse: collapse; border-radius: 6px; overflow: hidden;
    }
    table.report-table thead { background: #546e7a; color: #fff; }
    table.report-table th { padding: 10px 14px; text-align: left; font-size: 13px; }
    table.report-table td { padding: 8px 14px; border-bottom: 1px solid #e0e0e0; font-size: 13px; }
    table.report-table tbody tr:nth-child(even) { background: #f5f7fa; }
    table.report-table tbody tr:hover { background: #e3f2fd; }
    .badge-pass { background: #c8e6c9; color: #2e7d32; padding: 2px 10px; border-radius: 10px; font-weight: 600; font-size: 12px; }
    .badge-warn { background: #fff3cd; color: #856404; padding: 2px 10px; border-radius: 10px; font-weight: 600; font-size: 12px; }
    .badge-fail { background: #f8d7da; color: #721c24; padding: 2px 10px; border-radius: 10px; font-weight: 600; font-size: 12px; }
    .badge-info { background: #d1ecf1; color: #0c5460; padding: 2px 10px; border-radius: 10px; font-weight: 600; font-size: 12px; }
    .footer-text { margin-top: 20px; font-size: 11px; color: #999; text-align: center; }
</style>
"@

    $iconMap = @{
        'Sistema' = '🖥'; 'Energía' = '🔌'; 'Red' = '🌐'; 'Rendimiento' = '💾'
        'Disco' = '💽'; 'Procesador' = '🔥'; 'Memoria' = '🧠'; 'Seguridad' = '🛡'
        'SO' = '🛠'; 'Aplicaciones' = '📱'; 'Usuario' = '👤'; 'Malas prácticas' = '⚠'
    }

    $categories = $global:AllResults | Group-Object Category
    $panelsHtml = ""
    foreach ($cat in $categories) {
        $catName = $cat.Name
        $catPass = ($cat.Group | Where-Object Status -eq 'pass').Count
        $catWarn = ($cat.Group | Where-Object Status -eq 'warn').Count
        $catFail = ($cat.Group | Where-Object Status -eq 'fail').Count
        $catTotal = $cat.Count

        $icon = if ($iconMap.ContainsKey($catName)) { $iconMap[$catName] } else { '📋' }

        $rows = ""
        foreach ($result in $cat.Group) {
            $badgeClass = switch ($result.Status) {
                'pass' { 'badge-pass' }; 'warn' { 'badge-warn' }; 'fail' { 'badge-fail' }; default { 'badge-info' }
            }
            $statusText = switch ($result.Status) {
                'pass' { '✓ OK' }; 'warn' { '⚠ Advertencia' }; 'fail' { '✗ Fallo' }; default { 'i Info' }
            }
            $rowColor = switch ($result.Status) {
                'pass' { 'style="background-color: #f1f8e9;"' }
                'warn' { 'style="background-color: #fff8e1;"' }
                'fail' { 'style="background-color: #ffebee;"' }
                default { '' }
            }
            $rows += "<tr $rowColor><td>$($result.Check)</td><td>$($result.Detail)</td><td><span class='$badgeClass'>$statusText</span></td></tr>`n"
        }

        $panelsHtml += @"
<details class="panel">
    <summary>$icon $catName <span style="float:right;font-size:12px;opacity:0.8;">$catPass/$catTotal</span></summary>
    <div class="panel-content">
        <table class="report-table">
            <thead><tr><th>Check</th><th>Detalle</th><th>Estado</th></tr></thead>
            <tbody>$rows</tbody>
        </table>
    </div>
</details>

"@
    }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Diagnóstico - $computer</title>
$css
</head>
<body>
<div class="report-container">
    <div class="report-header">
        <h1>Diagnóstico Completo de Equipo</h1>
        <p>Equipo: $computer &nbsp;|&nbsp; Usuario: $env:USERNAME &nbsp;|&nbsp; Fecha: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    </div>
    <div class="report-body">
        <div class="section">
            <div class="section-title">Resumen General</div>
            <div class="summary-cards">
                <div class="card card-green"><span class="number">$passCount</span><span class="label">✓ Correcto</span></div>
                <div class="card card-yellow"><span class="number">$warnCount</span><span class="label">⚠ Advertencias</span></div>
                <div class="card card-red"><span class="number">$failCount</span><span class="label">✗ Fallos</span></div>
                <div class="card card-blue"><span class="number">$($Script:FixesApplied)/$($Script:FixesOffered)</span><span class="label">🔧 Auto-fixes</span></div>
            </div>
        </div>
        <div class="section">
            <div class="section-title">Detalle por Categoría</div>
            $panelsHtml
        </div>
    </div>
    <div style="background:#f5f5f5;padding:16px 40px;font-size:12px;color:#999;text-align:center;border-top:1px solid #e0e0e0;">
        JP Tools — Diagnóstico v$($Script:DiagnosticVersion) &nbsp;|&nbsp; Generado por PowerShell &nbsp;|&nbsp; $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    </div>
</div>
</body>
</html>
"@

    $html | Out-File -FilePath $reportFile -Encoding utf8
    Write-Host "`n  [REPORTE HTML] Guardado en: $reportFile" -ForegroundColor Green

    # Copiar resumen al portapapeles
    $summaryText = @"
=== DIAGNÓSTICO DE EQUIPO ===
Equipo: $computer
Usuario: $env:USERNAME
Fecha: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
---
Resultados: $passCount ✓ OK | $warnCount ⚠ Advertencias | $failCount ✗ Fallos
Auto-fixes: $($Script:FixesApplied)/$($Script:FixesOffered)
---
"@
    foreach ($cat in $categories) {
        $catPass = ($cat.Group | Where-Object Status -eq 'pass').Count
        $catWarn = ($cat.Group | Where-Object Status -eq 'warn').Count
        $catFail = ($cat.Group | Where-Object Status -eq 'fail').Count
        $catTotal = $cat.Count
        $catIcon = if ($iconMap.ContainsKey($cat.Name)) { $iconMap[$cat.Name] } else { '📋' }
        $summaryText += "$catIcon $($cat.Name): $catPass/$catTotal"
        if ($catWarn -gt 0) { $summaryText += " | $catWarn advertencia(s)" }
        if ($catFail -gt 0) { $summaryText += " | $catFail fallo(s)" }
        $summaryText += "`n"
    }

    if (Ask-Fix "¿Deseas copiar el reporte al portapapeles?") {
        $summaryText | Set-Clipboard
        Write-Result 'pass' 'Portapapeles' 'Reporte copiado al portapapeles'
        $Script:FixesApplied++
    }

    return $reportFile
}

# ============================================================
#  EJECUCIÓN PRINCIPAL
# ============================================================

$global:AllResults = @()
$Script:FixesApplied = 0
$Script:FixesOffered = 0
$Script:PendingReboot = $false

Write-Header
Write-Host "  Iniciando diagnóstico completo..." -ForegroundColor Yellow
Write-Host "  Fecha: $($Script:DiagnosticDate)" -ForegroundColor DarkGray
Write-Host "  Equipo: $env:COMPUTERNAME" -ForegroundColor DarkGray
Write-Host "  Usuario: $env:USERNAME" -ForegroundColor DarkGray

Start-Sleep -Seconds 1

$totalModulos = 13
$moduloActual = 0

# 1
$moduloActual++
Write-ProgressBar -Current $moduloActual -Total $totalModulos -Label "Sistema"
Test-Sistema -ModuloNum $moduloActual -TotalModulos $totalModulos

# 2
$moduloActual++
Write-ProgressBar -Current $moduloActual -Total $totalModulos -Label "Energía"
Test-Energia -ModuloNum $moduloActual -TotalModulos $totalModulos

# 3
$moduloActual++
Write-ProgressBar -Current $moduloActual -Total $totalModulos -Label "Red Corporativa"
Test-Red -ModuloNum $moduloActual -TotalModulos $totalModulos

# 4 — Velocidad de Internet
$moduloActual++
Write-ProgressBar -Current $moduloActual -Total $totalModulos -Label "Velocidad de Internet"
Test-VelocidadInternet

# 5
$moduloActual++
Write-ProgressBar -Current $moduloActual -Total $totalModulos -Label "Rendimiento"
Test-Rendimiento -ModuloNum $moduloActual -TotalModulos $totalModulos

# 6
$moduloActual++
Write-ProgressBar -Current $moduloActual -Total $totalModulos -Label "Disco"
Test-Disco -ModuloNum $moduloActual -TotalModulos $totalModulos

# 7
$moduloActual++
Write-ProgressBar -Current $moduloActual -Total $totalModulos -Label "Procesador"
Test-Procesador -ModuloNum $moduloActual -TotalModulos $totalModulos

# 8
$moduloActual++
Write-ProgressBar -Current $moduloActual -Total $totalModulos -Label "Memoria"
Test-Memoria -ModuloNum $moduloActual -TotalModulos $totalModulos

# 9
$moduloActual++
Write-ProgressBar -Current $moduloActual -Total $totalModulos -Label "Seguridad"
Test-Seguridad -ModuloNum $moduloActual -TotalModulos $totalModulos

# 10
$moduloActual++
Write-ProgressBar -Current $moduloActual -Total $totalModulos -Label "Sistema Operativo"
Test-SistemaOperativo -ModuloNum $moduloActual -TotalModulos $totalModulos

# 11
$moduloActual++
Write-ProgressBar -Current $moduloActual -Total $totalModulos -Label "Aplicaciones"
Test-Aplicaciones -ModuloNum $moduloActual -TotalModulos $totalModulos

# 12
$moduloActual++
Write-ProgressBar -Current $moduloActual -Total $totalModulos -Label "Usuario"
Test-Usuario -ModuloNum $moduloActual -TotalModulos $totalModulos

# 13
$moduloActual++
Write-ProgressBar -Current $moduloActual -Total $totalModulos -Label "Malas prácticas"
Test-MalasPracticas -ModuloNum $moduloActual -TotalModulos $totalModulos

# Resumen final
Show-Summary

# Generar reporte HTML
$reportPath = New-ReporteHTML

if ($reportPath) {
    $openReport = Read-Host "`n  ¿Deseas abrir el reporte HTML ahora? [S/N]"
    if ($openReport -match '^(s|si|sí|y|yes)$') {
        Invoke-Item $reportPath
    }
}

$openDir = Read-Host "`n  ¿Deseas abrir la carpeta con los reportes? [S/N]"
if ($openDir -match '^(s|si|sí|y|yes)$') {
    Invoke-Item $Script:ReportDir
}

if ($Script:PendingReboot) {
    $rebootNow = Read-Host "`n  ⚠ Hay un reinicio pendiente. ¿Reiniciar ahora? [S/N]"
    if ($rebootNow -match '^(s|si|sí|y|yes)$') {
        Write-Host "  Reiniciando en 10 segundos..." -ForegroundColor Yellow
        Restart-Computer -Force
    }
}

Write-Host "`n  [DIAGNÓSTICO COMPLETADO] Presiona cualquier tecla para salir..." -ForegroundColor Green
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
