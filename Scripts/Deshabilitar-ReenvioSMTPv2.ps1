<#
.SYNOPSIS
    Deshabilita el reenvío SMTP a nivel de buzón en Exchange Online para un listado de usuarios.

.DESCRIPTION
    Lee un listado de usuarios (TXT o CSV) y, por cada uno, limpia las propiedades de reenvío
    del buzón: ForwardingAddress, ForwardingSmtpAddress y DeliverToMailboxAndForward.

    Incluye modo de simulación (-Simular) y registro de auditoría en CSV con códigos de estado.

    NOTA: Este script actúa sobre el reenvío configurado a nivel de BUZÓN (Set-Mailbox).
    No toca las reglas de bandeja de entrada (Inbox Rules), que son otro mecanismo de reenvío.
    Ver la sección de notas al final para detectarlas.

.PARAMETER ListaUsuarios
    Ruta al archivo con el listado de usuarios (TXT con un usuario por línea, o CSV).

.PARAMETER Columna
    Nombre de la columna a usar cuando el archivo es CSV. Por defecto "UserPrincipalName".
    Si el archivo es TXT plano, se ignora.

.PARAMETER RutaLog
    Ruta del CSV de auditoría. Por defecto se genera en el directorio actual con fecha/hora.

.PARAMETER Simular
    Si se especifica, no aplica cambios; solo registra qué haría (estado SIMULADO).

.EXAMPLE
    .\Deshabilitar-ReenvioSMTP.ps1 -ListaUsuarios .\usuarios.txt -Simular

.EXAMPLE
    .\Deshabilitar-ReenvioSMTP.ps1 -ListaUsuarios .\usuarios.csv -Columna "Correo"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ListaUsuarios,

    [Parameter(Mandatory = $false)]
    [string]$Columna = "UserPrincipalName",

    [Parameter(Mandatory = $false)]
    [string]$RutaLog = ".\Log_ReenvioSMTP_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter(Mandatory = $false)]
    [switch]$Simular
)

# --- Validaciones iniciales ---
if (-not (Test-Path -Path $ListaUsuarios)) {
    Write-Error "No se encontró el archivo de listado: $ListaUsuarios"
    return
}

# --- Cargar el listado de usuarios (detecta TXT vs CSV) ---
$usuarios = @()
$extension = [System.IO.Path]::GetExtension($ListaUsuarios).ToLower()

if ($extension -eq ".csv") {
    $csv = Import-Csv -Path $ListaUsuarios
    if ($csv.Count -gt 0 -and -not ($csv[0].PSObject.Properties.Name -contains $Columna)) {
        Write-Error "La columna '$Columna' no existe en el CSV. Columnas disponibles: $($csv[0].PSObject.Properties.Name -join ', ')"
        return
    }
    $usuarios = $csv | ForEach-Object { $_.$Columna } | Where-Object { $_ -and $_.Trim() -ne "" }
}
else {
    # TXT plano: un usuario por línea
    $usuarios = Get-Content -Path $ListaUsuarios | Where-Object { $_ -and $_.Trim() -ne "" } | ForEach-Object { $_.Trim() }
}

if ($usuarios.Count -eq 0) {
    Write-Error "El listado no contiene usuarios válidos."
    return
}

Write-Host "Usuarios a procesar: $($usuarios.Count)" -ForegroundColor Cyan
if ($Simular) {
    Write-Host "MODO SIMULACIÓN ACTIVO: no se aplicarán cambios." -ForegroundColor Yellow
}

# --- Conexión a Exchange Online ---
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Error "Falta el módulo ExchangeOnlineManagement. Instálalo con: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    return
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Conectado a Exchange Online." -ForegroundColor Green
}
catch {
    Write-Error "No se pudo conectar a Exchange Online: $($_.Exception.Message)"
    return
}

# --- Procesamiento ---
$resultados = New-Object System.Collections.Generic.List[object]
$contador = 0

foreach ($usuario in $usuarios) {
    $contador++
    Write-Progress -Activity "Deshabilitando reenvío SMTP" -Status "$contador de $($usuarios.Count): $usuario" -PercentComplete (($contador / $usuarios.Count) * 100)

    $registro = [PSCustomObject]@{
        Usuario              = $usuario
        ForwardingAddress    = ""
        ForwardingSmtpAddress = ""
        DeliverAndForward    = ""
        Estado               = ""
        Detalle              = ""
        FechaHora            = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }

    try {
        $mbx = Get-Mailbox -Identity $usuario -ErrorAction Stop

        # Capturar el estado actual del reenvío (para el log)
        $registro.ForwardingAddress     = $mbx.ForwardingAddress
        $registro.ForwardingSmtpAddress = $mbx.ForwardingSmtpAddress
        $registro.DeliverAndForward     = $mbx.DeliverToMailboxAndForward

        $tieneReenvio = ($null -ne $mbx.ForwardingAddress) -or `
                        (-not [string]::IsNullOrWhiteSpace([string]$mbx.ForwardingSmtpAddress))

        if (-not $tieneReenvio) {
            $registro.Estado  = "SIN REENVIO"
            $registro.Detalle = "El buzón no tiene reenvío configurado a nivel de buzón."
        }
        elseif ($Simular) {
            $registro.Estado  = "SIMULADO"
            $registro.Detalle = "Se limpiaría el reenvío SMTP (modo simulación)."
        }
        else {
            Set-Mailbox -Identity $usuario `
                        -ForwardingAddress $null `
                        -ForwardingSmtpAddress $null `
                        -DeliverToMailboxAndForward $false `
                        -ErrorAction Stop

            $registro.Estado  = "DESHABILITADO"
            $registro.Detalle = "Reenvío SMTP eliminado correctamente."
        }
    }
    catch {
        if ($_.Exception.Message -match "couldn't be found|no se encontró|not found") {
            $registro.Estado  = "NO ENCONTRADO"
            $registro.Detalle = "No existe un buzón para este usuario."
        }
        else {
            $registro.Estado  = "ERROR"
            $registro.Detalle = $_.Exception.Message
        }
    }

    # Salida en consola con color según estado
    $color = switch ($registro.Estado) {
        "DESHABILITADO" { "Green" }
        "SIMULADO"      { "Yellow" }
        "SIN REENVIO"   { "Gray" }
        "NO ENCONTRADO" { "Magenta" }
        "ERROR"         { "Red" }
        default         { "White" }
    }
    Write-Host ("[{0}] {1} - {2}" -f $registro.Estado, $usuario, $registro.Detalle) -ForegroundColor $color

    $resultados.Add($registro)
}

# --- Exportar log ---
try {
    $resultados | Export-Csv -Path $RutaLog -NoTypeInformation -Encoding UTF8
    Write-Host "`nLog de auditoría generado en: $RutaLog" -ForegroundColor Cyan
}
catch {
    Write-Warning "No se pudo escribir el log en $RutaLog : $($_.Exception.Message)"
}

# --- Resumen ---
Write-Host "`n===== RESUMEN =====" -ForegroundColor Cyan
$resultados | Group-Object Estado | Sort-Object Name | ForEach-Object {
    Write-Host ("{0,-15}: {1}" -f $_.Name, $_.Count)
}

# --- Desconexión ---
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "`nDesconectado de Exchange Online." -ForegroundColor Green
