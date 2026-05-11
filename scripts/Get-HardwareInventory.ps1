
# Recopilar información
$Sistema = Get-CimInstance -ClassName Win32_ComputerSystem
$BIOS = Get-CimInstance -ClassName Win32_BIOS
$CPU = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
$ModulosRAM = Get-CimInstance -ClassName Win32_PhysicalMemory
# Sumar RAM total de todos los módulos
$RAM_Total_GB = [Math]::Round(($ModulosRAM | Measure-Object -Property Capacity -Sum).Sum / 1GB, 0)
# Obtener discos físicos con tipo
$DiscosFisicos = Get-PhysicalDisk | Where-Object { $_.BusType -ne 'USB' }
# Construir información de discos
$InfoDiscos = ""
foreach ($Disco in $DiscosFisicos) {
    $Disco_GB = [Math]::Round($Disco.Size / 1GB, 1)
    $Tipo = switch ($Disco.MediaType) {
        "SSD" { "SSD" }
        "HDD" { "HDD" }
        default { "Desconocido" }
    }
    $InfoDiscos += "  $($Disco.FriendlyName) ($Tipo) - $Disco_GB GB`n"
}
# Si no hay discos detectados por Get-PhysicalDisk, usar fallback
if ([string]::IsNullOrWhiteSpace($InfoDiscos)) {
    $Discos = Get-CimInstance -ClassName Win32_DiskDrive
    foreach ($Disco in $Discos) {
        $Disco_GB = [Math]::Round($Disco.Size / 1GB, 1)
        $InfoDiscos += "  $($Disco.Model) - $Disco_GB GB`n"
    }
}
# Crear salida
$Salida = @"
========================================
  INVENTARIO DE HARDWARE
========================================
  MARCA:         $($Sistema.Manufacturer)
  MODELO:        $($Sistema.Model)
  SERIAL:        $($BIOS.SerialNumber)
  PROCESADOR:    $($CPU.Name)
  RAM TOTAL:     $RAM_Total_GB GB
  DISCO DURO:    $InfoDiscos
========================================
"@
# Mostrar en pantalla
Write-Host $Salida
# Copiar al portapapeles
$Salida | Set-Clipboard
Write-Host "[OK] Información copiada al portapapeles.`n" -ForegroundColor Green