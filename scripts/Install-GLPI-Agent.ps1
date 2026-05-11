# ================================
# Configuración
# ================================

$ServerUrl      = 'http://10.0.201.14:8088/glpi'
$Version        = '1.15'
$MonitorVersion = '1.4.1'

$InstallerUrl  = "https://github.com/glpi-project/glpi-agent/releases/download/$Version/GLPI-Agent-$Version-x64.msi"
$InstallerFile = Join-Path $env:TEMP "GLPI-Agent-$Version-x64.msi"
$MsiLog        = "C:\Windows\Temp\glpi-agent-$Version-install.log"

$MonitorUrl    = "https://github.com/glpi-project/glpi-agentmonitor/releases/download/$MonitorVersion/GLPI-AgentMonitor-x64.exe"

$ErrorActionPreference = 'Stop'

# ================================
# Funciones auxiliares
# ================================

function Assert-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object Security.Principal.WindowsPrincipal($id)

    if (-not $p.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
        throw "Este script debe ejecutarse con privilegios de administrador."
    }
}

function Get-InstallDir {

    foreach ($k in 'HKLM:\SOFTWARE\GLPI-Agent','HKLM:\SOFTWARE\WOW6432Node\GLPI-Agent') {

        try {
            $reg = Get-ItemProperty -Path $k -ErrorAction Stop

            foreach ($n in 'InstallDir','installdir','INSTALLDIR') {
                if ($reg.PSObject.Properties.Name -contains $n) {
                    return $reg.$n
                }
            }
        }
        catch { }
    }

    return 'C:\Program Files\GLPI-Agent'
}

function Show-FinalScreen {

    param(
        [string]$InstallDir,
        [string]$Server
    )

    Clear-Host

    $green = 'Green'
    $white = 'White'

Write-Host ""
Write-Host "  +--------------------------------------------------------------+" -ForegroundColor $green
Write-Host "  |                                                              |" -ForegroundColor $green
Write-Host "  |            GLPI AGENT - INSTALACION COMPLETADA               |" -ForegroundColor $green
Write-Host "  |                                                              |" -ForegroundColor $green
Write-Host "  +--------------------------------------------------------------+" -ForegroundColor $green
Write-Host "  | Estado        : OK                                           |" -ForegroundColor $white
Write-Host "  |                                                              |" -ForegroundColor $green
Write-Host "  | Servidor GLPI :                                              |" -ForegroundColor $white
Write-Host "  |   $Server" -ForegroundColor $green
Write-Host "  |                                                              |" -ForegroundColor $green
Write-Host "  | Directorio    :                                              |" -ForegroundColor $white
Write-Host "  |   $InstallDir" -ForegroundColor $green
Write-Host "  |                                                              |" -ForegroundColor $green
Write-Host "  | Componentes   : Servicio y Monitor configurados correctamente|" -ForegroundColor $white
Write-Host "  |                                                              |" -ForegroundColor $green
Write-Host "  +--------------------------------------------------------------+" -ForegroundColor $green
Write-Host ""
Write-Host "    Presiona [ ESC ] para cerrar esta ventana..." -ForegroundColor $green
Write-Host ""
    do {
        $key = [System.Console]::ReadKey($true)
    }
    until ($key.Key -eq 'Escape')
}

# Inicio

Assert-Admin

Write-Host "Descargando GLPI Agent..." -ForegroundColor Cyan

Invoke-WebRequest -Uri $InstallerUrl -OutFile $InstallerFile

if (-not (Test-Path $InstallerFile) -or (Get-Item $InstallerFile).Length -lt 100kb) {
    throw "La descarga del instalador de GLPI Agent fallo o esta incompleta."
}

# Instalación MSI

$msiArgs = @(
    '/i', "`"$InstallerFile`"",
    '/quiet',
    '/l*v', "`"$MsiLog`"",
    "SERVER=`"$ServerUrl`"",
    'EXECMODE=1',
    'AGENTMONITOR=1',
    'QUICKINSTALL=0',
    'RUNNOW=1',
    'ADD_FIREWALL_EXCEPTION=1'
) -join ' '

Write-Host "Instalando GLPI Agent..." -ForegroundColor Cyan

$proc = Start-Process msiexec.exe -ArgumentList $msiArgs -Wait -NoNewWindow -PassThru

if ($proc.ExitCode -ne 0) {
    throw "Error durante la instalacion de GLPI Agent. ExitCode: $($proc.ExitCode). Revisa el log: $MsiLog"
}

# Servicio

$svc = Get-Service -Name 'GLPI-Agent','glpi-agent' -ErrorAction SilentlyContinue |
       Select-Object -First 1

if ($svc) {

    if ($svc.Status -ne 'Running') {
        Start-Service -Name $svc.Name -ErrorAction Stop
    }

} else {
    throw "El servicio de GLPI Agent no fue encontrado despues de la instalacion."
}



$InstallDir = Get-InstallDir

if (-not (Test-Path $InstallDir)) {
    throw "No se encontro el directorio de instalacion de GLPI Agent."
}



$MonitorExe = Get-ChildItem -Path $InstallDir -Recurse -File -ErrorAction SilentlyContinue |
    Where-Object {
        $_.Name -match '^GLPI-AgentMonitor.*\.exe$' -or
        $_.Name -match '^glpi-agentmonitor.*\.exe$'
    } |
    Select-Object -ExpandProperty FullName -First 1



if (-not $MonitorExe) {

    $MonitorExe = Join-Path $InstallDir 'GLPI-AgentMonitor-x64.exe'

    Write-Host "Descargando GLPI Agent Monitor..." -ForegroundColor Cyan

    Invoke-WebRequest -Uri $MonitorUrl -OutFile $MonitorExe

    if (-not (Test-Path $MonitorExe) -or (Get-Item $MonitorExe).Length -lt 100kb) {
        throw "La descarga de GLPI Agent Monitor fallo."
    }

    Unblock-File -Path $MonitorExe -ErrorAction SilentlyContinue
}

if (-not (Test-Path $MonitorExe)) {
    throw "No se encontro el ejecutable del monitor."
}


$runKey  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
$runName = 'GLPI-AgentMonitor'
$runCmd  = "`"$MonitorExe`""

New-ItemProperty -Path $runKey -Name $runName -Value $runCmd -PropertyType String -Force | Out-Null



Get-Process -Name 'GLPI-AgentMonitor' -ErrorAction SilentlyContinue | ForEach-Object {
    try {
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    }
    catch { }
}

Start-Sleep -Seconds 1

Start-Process -FilePath $MonitorExe -WindowStyle Minimized -ErrorAction Stop



Show-FinalScreen -InstallDir $InstallDir -Server $ServerUrl

