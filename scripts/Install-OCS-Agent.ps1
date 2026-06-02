# ================================
# Configuración
# ================================

$ServerUrl = 'https://ocs.comunidadsd.org/ocsinventory'
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

function Show-InputDialog {
    param([string]$Title, [string]$Prompt, [string]$Default = '')

    Add-Type -AssemblyName System.Windows.Forms

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(450,200)
    $form.StartPosition = 'CenterScreen'
    $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Prompt
    $label.Location = New-Object System.Drawing.Point(20,20)
    $label.Size = New-Object System.Drawing.Size(400,25)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Text = $Default
    $textBox.Location = New-Object System.Drawing.Point(20,50)
    $textBox.Size = New-Object System.Drawing.Size(400,25)
    $form.Controls.Add($textBox)

    $okBtn = New-Object System.Windows.Forms.Button
    $okBtn.Text = 'Aceptar'
    $okBtn.Location = New-Object System.Drawing.Point(150,100)
    $okBtn.Size = New-Object System.Drawing.Size(100,30)
    $okBtn.Add_Click({ $form.DialogResult = 'OK'; $form.Close() })
    $form.Controls.Add($okBtn)

    $form.Add_Shown({ $textBox.Focus() })
    $null = $form.ShowDialog()

    return $textBox.Text
}

function Show-FinalScreen {
    param([string]$TAG, [string]$User)

    Clear-Host
    Write-Host ""
    Write-Host "  +--------------------------------------------------------------+" -ForegroundColor Green
    Write-Host "  |                                                              |" -ForegroundColor Green
    Write-Host "  |            OCS INVENTORY - INSTALACION COMPLETADA            |" -ForegroundColor Green
    Write-Host "  |                                                              |" -ForegroundColor Green
    Write-Host "  +--------------------------------------------------------------+" -ForegroundColor Green
    Write-Host "  | Estado        : OK                                           |" -ForegroundColor White
    Write-Host "  |                                                              |" -ForegroundColor Green
    Write-Host "  | Servidor OCS : $ServerUrl" -ForegroundColor Green
    Write-Host "  | TAG           : $TAG" -ForegroundColor Green
    Write-Host "  | Usuario AD    : $User" -ForegroundColor Green
    Write-Host "  |                                                              |" -ForegroundColor Green
    Write-Host "  +--------------------------------------------------------------+" -ForegroundColor Green
    Write-Host ""
    Write-Host "    Presiona [ ESC ] para cerrar esta ventana..." -ForegroundColor Green
    Write-Host ""

    do {
        $key = [System.Console]::ReadKey($true)
    } until ($key.Key -eq 'Escape')
}

# ================================
# Inicio
# ================================

Assert-Admin

# Buscar instalador
$searchPaths = @(
    Join-Path $PSScriptRoot "..\tools\OCS Inventory2.9.1.0"
    "D:\BASICO\OCS Inventory2.9.1.0"
)

$installDir = $null

foreach ($p in $searchPaths) {
    $resolved = Resolve-Path $p -ErrorAction SilentlyContinue
    if ($resolved) {
        $exe = Join-Path $resolved "OCS-Windows-Agent-Setup-x64.exe"
        $dll = Join-Path $resolved "libcurl.dll"

        if ((Test-Path $exe) -and (Test-Path $dll)) {
            $installDir = $resolved
            break
        }
    }
}

if (-not $installDir) {
    throw "No se encontro el instalador de OCS ni libcurl.dll en las rutas esperadas."
}

# Pedir TAG y Usuario AD (GUI)
$TAG  = Show-InputDialog -Title "OCS Inventory - Instalacion" -Prompt "Ingresa el TAG del equipo (ej: CSI3_SISTEMAS, CLN_PALIZA):"
$User = Show-InputDialog -Title "OCS Inventory - Instalacion" -Prompt "Ingresa el usuario de Active Directory:" -Default $env:USERNAME

if ([string]::IsNullOrWhiteSpace($TAG) -or [string]::IsNullOrWhiteSpace($User)) {
    throw "TAG y Usuario AD son obligatorios."
}

# Instalación silenciosa
$installer = Join-Path $installDir "OCS-Windows-Agent-Setup-x64.exe"
$installArgs = "/S /SERVER=`"$ServerUrl`" /TAG=`"$TAG`" /USER=`"$User`" /NOW"
Write-Host "Instalando OCS Inventory Agent..." -ForegroundColor Cyan

$proc = Start-Process -FilePath $installer -ArgumentList $installArgs -Wait -NoNewWindow -PassThru

if ($proc.ExitCode -ne 0) {
    throw "Error durante la instalacion. ExitCode: $($proc.ExitCode)."
}

# Post-instalación
$agentDir = "C:\Program Files\OCS Inventory Agent"

Write-Host "Deteniendo servicio OCS Inventory Service..." -ForegroundColor Cyan
Stop-Service -Name "OCS Inventory Service" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

Get-Process -Name "OcsSystray","OCSInventory" -ErrorAction SilentlyContinue | Stop-Process -Force

Write-Host "Copiando libcurl.dll..." -ForegroundColor Cyan
$dllSrc  = Join-Path $installDir "libcurl.dll"
$dllDest = Join-Path $agentDir "libcurl.dll"

Copy-Item -Path $dllSrc -Destination $dllDest -Force

Write-Host "Forzando SSL=0 (desactivar validacion de certificados)..." -ForegroundColor Cyan
$iniPath = "$env:ProgramData\OCS Inventory NG\Agent\ocsinventory.ini"
(Get-Content $iniPath) -replace 'SSL=1', 'SSL=0' -replace 'AuthRequired=1', 'AuthRequired=0' | Set-Content $iniPath

Write-Host "Iniciando servicio OCS Inventory Service..." -ForegroundColor Cyan
Start-Service -Name "OCS Inventory Service"

Start-Sleep -Seconds 3

Write-Host "Ejecutando OCSInventory.exe..." -ForegroundColor Cyan
$ocsExe = Join-Path $agentDir "OCSInventory.exe"
if (Test-Path $ocsExe) {
    Start-Process -FilePath $ocsExe -WindowStyle Minimized
}

Write-Host "Ejecutando OcsSystray.exe..." -ForegroundColor Cyan
$trayExe = Join-Path $agentDir "OcsSystray.exe"
if (Test-Path $trayExe) {
    Start-Process -FilePath $trayExe -WindowStyle Minimized
}

Write-Host "Verificando envio de inventario..." -ForegroundColor Cyan
$logPath = "$env:ProgramData\OCS Inventory NG\Agent\OCSInventory.log"
$maxWait = 60
$waited  = 0

while ($waited -lt $maxWait) {
    Start-Sleep -Seconds 2
    $waited += 2

    if (Test-Path $logPath) {
        $log = Get-Content $logPath -Tail 20
        if ($log -match "Inventory successfully sent") {
            Write-Host "[OK] Inventario enviado exitosamente." -ForegroundColor Green
            break
        }
    }
}

if ($waited -ge $maxWait) {
    Write-Host "[AVISO] No se pudo confirmar el envio en $maxWait segundos. Revisa el log manualmente." -ForegroundColor Yellow
}

Show-FinalScreen -TAG $TAG -User $User
