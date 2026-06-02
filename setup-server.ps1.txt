<#
.SYNOPSIS
    setup-server.ps1 — Configuración automática de Lenovo Legion como servidor helpdeskbot
.DESCRIPTION
    Ejecutar como ADMINISTRADOR: clic derecho > "Ejecutar con PowerShell" (o: powershell -ExecutionPolicy Bypass -File setup-server.ps1)
    Requiere Windows 11 Pro y al menos 20 GB libres en disco.
#>

$ErrorActionPreference = 'Stop'
$C = @{ G='[32m'; Y='[33m'; R='[31m'; C='[36m'; N='[0m' }

Write-Host "${C[C]}╔════════════════════════════════════════════════════╗${C[N]}"
Write-Host "${C[C]}║   🚀 SETUP SERVER — LENOVO LEGION (WINDOWS 11)    ║${C[N]}"
Write-Host "${C[C]}╚════════════════════════════════════════════════════╝${C[N]}"
Write-Host ""

# ─── 1. Admin check ─────────────────────────────────────────────────────────
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')) {
    Write-Host "${C[R]}ERROR: Ejecuta como ADMINISTRADOR${C[N]}"
    Write-Host "Clic derecho > Ejecutar con PowerShell (como admin)"
    Read-Host "Presiona Enter para salir"; exit 1
}
Write-Host "[✓] Administrador`n"

# ─── 2. Helper function ─────────────────────────────────────────────────────
function Install-IfMissing {
    param([string]$Name, [string]$WingetId)
    if (!(Get-Command $Name -ErrorAction SilentlyContinue)) {
        Write-Host "  Instalando $Name..."
        winget install --id $WingetId --silent --accept-package-agreements --accept-source-agreements
        refreshenv
        $env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
    } else {
        Write-Host "  [✓] $Name ya instalado"
    }
}

# ─── 3. Dependencias (winget) ───────────────────────────────────────────────
Write-Host "${C[Y]}▶ Instalando dependencias del sistema...${C[N]}"
Install-IfMissing -Name git -WingetId Git.Git
Install-IfMissing -Name node -WingetId OpenJS.NodeJS.LTS
Install-IfMissing -Name chrome -WingetId Google.Chrome
Write-Host ""

# ─── 4. Verificar Node/npm ──────────────────────────────────────────────────
$env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
refreshenv 2>$null
Write-Host "[✓] Node $(node -v) / npm $(npm -v)`n"

# ─── 5. Ollama ──────────────────────────────────────────────────────────────
Write-Host "${C[Y]}▶ Instalando Ollama...${C[N]}"
if (!(Get-Command ollama -ErrorAction SilentlyContinue)) {
    $installer = "$env:TEMP\OllamaSetup.exe"
    Invoke-WebRequest -Uri 'https://ollama.com/download/OllamaSetup.exe' -OutFile $installer
    Start-Process -Wait -FilePath $installer -ArgumentList '/SILENT'
    Remove-Item $installer -Force
    $env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
    Write-Host "  [✓] Ollama instalado"
} else {
    Write-Host "  [✓] Ollama ya instalado"
}
# Asegurar servicio corriendo
$svc = Get-Service Ollama -ErrorAction SilentlyContinue
if ($svc -and $svc.Status -ne 'Running') { Start-Service Ollama; Start-Sleep 2 }
Write-Host ""

# ─── 6. Modelo ──────────────────────────────────────────────────────────────
Write-Host "${C[Y]}▶ Descargando modelo qwen2.5vl:7b...${C[N]}"
Write-Host "  (puede tardar 10-30 min, no cerrar la ventana)"
ollama pull qwen2.5vl:7b
Write-Host "  [✓] Modelo descargado`n"

# ─── 7. GPU check ───────────────────────────────────────────────────────────
Write-Host "${C[Y]}▶ Verificando GPU...${C[N]}"
$gpu = Get-CimInstance Win32_VideoController | Where-Object { $_.Name -match 'NVIDIA|AMD|Intel.*Arc' }
if ($gpu) {
    Write-Host "  [✓] GPU: $($gpu.Name)"
    Write-Host "  Recomendación: instala drivers NVIDIA Studio desde"
    Write-Host "  https://www.nvidia.com/Download/index.aspx"
} else {
    Write-Host "  ${C[R]}⚠ No se detectó GPU dedicada — se usará CPU (lento con imágenes)${C[N]}"
}
# Exponer Ollama en LAN
[Environment]::SetEnvironmentVariable('OLLAMA_HOST','0.0.0.0','Machine')
Write-Host "  [✓] OLLAMA_HOST=0.0.0.0 (accesible desde LAN)`n"

# ─── 8. Clonar repo ─────────────────────────────────────────────────────────
Write-Host "${C[Y]}▶ Clonando repositorio...${C[N]}"
$repo = 'C:\helpdeskbot'
if (Test-Path $repo) {
    Push-Location $repo; git pull; Pop-Location
    Write-Host "  [✓] Repositorio actualizado"
} else {
    Push-Location C:\
    git clone https://github.com/jp1640870-max/helpdeskbot.git $repo
    Pop-Location
    Write-Host "  [✓] Repositorio clonado"
}
Write-Host ""

# ─── 9. .env ────────────────────────────────────────────────────────────────
Write-Host "${C[Y]}▶ Creando .env...${C[N]}"
@'
OLLAMA_HOST=http://localhost:11434
OLLAMA_MODEL=qwen2.5vl:7b

DASHBOARD_PORT=3001
DASHBOARD_HOST=0.0.0.0
DASHBOARD_USER=admin
DASHBOARD_PASSWORD=helpdesk123

PORT=3000
LOG_LEVEL=info
AGGREGATION_TIME=12000
'@ | Set-Content "$repo\.env" -Encoding UTF8
Write-Host "  [✓] .env creado`n"

# ─── 10. npm install ────────────────────────────────────────────────────────
Write-Host "${C[Y]}▶ Instalando dependencias npm...${C[N]}"
Push-Location $repo
npm install
Pop-Location
Write-Host "  [✓] npm install listo`n"

# ─── 11. PM2 ────────────────────────────────────────────────────────────────
Write-Host "${C[Y]}▶ Configurando PM2...${C[N]}"
if (!(Get-Command pm2 -ErrorAction SilentlyContinue)) {
    npm install -g pm2
}
Set-Location $repo
pm2 start index.js --name helpdeskbot
pm2 save
pm2 startup
Set-Location C:\
Write-Host "  [✓] PM2 — bot iniciado y auto-start configurado`n"

# ─── 12. OpenSSH Server ─────────────────────────────────────────────────────
Write-Host "${C[Y]}▶ Configurando OpenSSH Server...${C[N]}"
$svc = Get-Service sshd -ErrorAction SilentlyContinue
if (!$svc) {
    Add-WindowsCapability -Online -Name 'OpenSSH.Server~~~~0.0.1.0'
    Write-Host "  [✓] OpenSSH instalado"
}
Set-Service sshd -StartupType Automatic
Start-Service sshd -ErrorAction SilentlyContinue
Write-Host "  [✓] SSH en puerto 22`n"

# ─── 13. Firewall ───────────────────────────────────────────────────────────
Write-Host "${C[Y]}▶ Abriendo puertos en firewall...${C[N]}"
$rules = @(
    @{Name='SSH (22)';    Port=22;  Proto='TCP'}
    @{Name='Dashboard (3001)'; Port=3001; Proto='TCP'}
)
foreach ($r in $rules) {
    if (!(Get-NetFirewallRule -DisplayName $r.Name -ErrorAction SilentlyContinue)) {
        New-NetFirewallRule -DisplayName $r.Name -Direction Inbound -Protocol $r.Proto -LocalPort $r.Port -Action Allow
        Write-Host "  [✓] Puerto $($r.Port) abierto"
    } else {
        Write-Host "  [✓] Puerto $($r.Port) ya configurado"
    }
}
Write-Host ""

# ─── 14. IP estática (asistida) ─────────────────────────────────────────────
Write-Host "${C[Y]}▶ Configurar IP estática (opcional)${C[N]}"
Write-Host "  Para asignar 10.0.120.151/24 con gateway 10.0.120.1:"
Write-Host ""
$adapters = Get-NetAdapter | Where-Object Status -eq Up
$i = 0
foreach ($ad in $adapters) {
    $ip = (Get-NetIPAddress -InterfaceIndex $ad.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress
    Write-Host "  [$i] $($ad.Name)  → $($ip)"
    $global:adIdx = $i
    $i++
}
Write-Host ""
Write-Host "  Ejecuta MANUALMENTE si necesitas IP fija:"
Write-Host '    $ifIndex = (Get-NetAdapter | Where-Object Status -eq Up)[<NÚMERO>].ifIndex'
Write-Host '    New-NetIPAddress -InterfaceIndex $ifIndex -IPAddress 10.0.120.151 -PrefixLength 24 -DefaultGateway 10.0.120.1'
Write-Host '    Set-DnsClientServerAddress -InterfaceIndex $ifIndex -ServerAddresses 10.0.120.1'
Write-Host ""

# ─── 15. Fin ────────────────────────────────────────────────────────────────
Write-Host "${C[G]}╔════════════════════════════════════════════════════╗${C[N]}"
Write-Host "${C[G]}║   ✅ SETUP COMPLETO                               ║${C[N]}"
Write-Host "${C[G]}╚════════════════════════════════════════════════════╝${C[N]}"
Write-Host ""
Write-Host "  ${C[C]}Resumen:${C[N]}"
Write-Host "  • Repositorio: C:\helpdeskbot"
Write-Host "  • Dashboard:   http://localhost:3001  (admin / helpdesk123)"
Write-Host "  • Bot (PM2):   pm2 status / pm2 logs helpdeskbot"
Write-Host "  • SSH:         ssh ""soporte tecnico@10.0.120.151"""
Write-Host "  • Ollama:      http://localhost:11434"
Write-Host "  • Modelo:      qwen2.5vl:7b"
Write-Host ""
Write-Host "  ${C[Y]}Próximos pasos:${C[N]}"
Write-Host "  1. Abre http://localhost:3001 y escanea el QR desde WhatsApp"
Write-Host "  2. O desde WhatsApp > Dispositivos vinculados > escanea QR"
Write-Host "  3. Configura IP estática manualmente si lo requieres"
Write-Host ""

Read-Host "Presiona Enter para salir"
