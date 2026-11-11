# Programa la revision automatica en ESTE computador con el Programador de tareas de Windows.
# No necesita Claude ni permisos de administrador. Solo descarga y archiva; no envia correo ni toca el wiki.
#   Instalar:  powershell -ExecutionPolicy Bypass -File scripts\instalar_tarea.ps1
#   Quitar:    powershell -ExecutionPolicy Bypass -File scripts\instalar_tarea.ps1 -Quitar
#   Otra hora: powershell -ExecutionPolicy Bypass -File scripts\instalar_tarea.ps1 -Hora 07:30
param([switch]$Quitar, [string]$Hora = '08:00')

$ErrorActionPreference = 'Stop'
$script  = Join-Path $PSScriptRoot 'revisar_tablas.ps1'
$accion  = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f $script
$tareas  = @(
    @{ nombre = 'AgenteLCR_Tablas_Lunes'; args = @('/SC', 'WEEKLY', '/D', 'MON') },
    @{ nombre = 'AgenteLCR_Tablas_Dia1';  args = @('/SC', 'MONTHLY', '/D', '1') }
)

foreach ($t in $tareas) {
    if ($Quitar) {
        & schtasks.exe /Delete /TN $t.nombre /F | Out-Null
        if ($LASTEXITCODE -eq 0) { Write-Output "Quitada: $($t.nombre)" } else { Write-Output "No existia: $($t.nombre)" }
    }
    else {
        & schtasks.exe /Create /TN $t.nombre /TR $accion /ST $Hora /F @($t.args) | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "No se pudo crear la tarea $($t.nombre)" }
        Write-Output "Programada: $($t.nombre) a las $Hora"
    }
}
if (-not $Quitar) {
    Write-Output "Listo. Revisa los lunes y el dia 1 de cada mes. El resumen de cada revision queda en estado\historial.log"
}
