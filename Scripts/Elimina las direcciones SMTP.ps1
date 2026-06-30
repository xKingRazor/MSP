<#
.SYNOPSIS
    Elimina las direcciones SMTP secundarias con dominio @banistmo.com de todos los buzones del tenant.

.DESCRIPTION
    Recorre todos los buzones de Exchange Online, busca direcciones proxy (EmailAddresses)
    cuyo dominio sea exactamente @banistmo.com y las elimina.
    Por seguridad NO elimina la direccion principal (prefijo "SMTP:" en mayusculas) y
    soporta modo de simulacion con -WhatIf. Genera un log CSV con todo lo realizado.

.PARAMETER Dominio
    Dominio a buscar y eliminar. Por defecto: banistmo.com

.PARAMETER LogPath
    Ruta del archivo CSV de resultados.

.EXAMPLE
    # Primero SIEMPRE en modo simulacion para revisar que se va a tocar:
    .\Eliminar-SMTP-Banistmo.ps1 -WhatIf

.EXAMPLE
    # Ejecucion real:
    .\Eliminar-SMTP-Banistmo.ps1

.NOTES
    - Requiere el modulo ExchangeOnlineManagement y rol de Administrador de Exchange.
    - Si el tenant usa sincronizacion de directorio (Azure AD Connect / Entra Connect),
      las direcciones sincronizadas NO se pueden modificar aqui: hay que cambiarlas en el
      Active Directory local (proxyAddresses) y esperar la sincronizacion.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$Dominio = "banistmo.com",
    [string]$LogPath = ".\Eliminacion_SMTP_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# --- Conexion a Exchange Online ---
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Instalando modulo ExchangeOnlineManagement..." -ForegroundColor Yellow
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowBanner:$false

$resultados = @()

# --- Obtener todos los buzones ---
Write-Host "Obteniendo buzones..." -ForegroundColor Cyan
$buzones = Get-Mailbox -ResultSize Unlimited
$total   = $buzones.Count
$i       = 0

foreach ($buzon in $buzones) {
    $i++
    Write-Progress -Activity "Revisando buzones" `
                   -Status "$i de $total : $($buzon.UserPrincipalName)" `
                   -PercentComplete (($i / [math]::Max($total,1)) * 100)

    # Direcciones SMTP cuyo dominio sea EXACTAMENTE @banistmo.com
    # ("-like" es insensible a mayusculas, por eso captura SMTP: y smtp:)
    $aEliminar = $buzon.EmailAddresses | Where-Object {
        $_ -like "smtp:*@$Dominio"
    }

    foreach ($direccion in $aEliminar) {

        # El prefijo SMTP: en MAYUSCULAS indica la direccion principal -> no se toca
        if ($direccion -clike "SMTP:*") {
            Write-Warning "OMITIDA direccion principal en $($buzon.UserPrincipalName): $direccion. Cambiela manualmente antes de eliminarla."
            $resultados += [PSCustomObject]@{
                Buzon     = $buzon.UserPrincipalName
                Direccion = $direccion
                Accion    = "OMITIDA (es la principal)"
            }
            continue
        }

        if ($PSCmdlet.ShouldProcess($buzon.UserPrincipalName, "Eliminar $direccion")) {
            try {
                Set-Mailbox -Identity $buzon.Identity -EmailAddresses @{ remove = $direccion } -ErrorAction Stop
                $accion = "ELIMINADA"
            }
            catch {
                $accion = "ERROR: $($_.Exception.Message)"
            }
        }
        else {
            $accion = "SIMULADA (WhatIf)"
        }

        $resultados += [PSCustomObject]@{
            Buzon     = $buzon.UserPrincipalName
            Direccion = $direccion
            Accion    = $accion
        }
    }
}

# --- Exportar log ---
if ($resultados.Count -gt 0) {
    $resultados | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nProceso terminado. Direcciones procesadas: $($resultados.Count)" -ForegroundColor Green
    Write-Host "Log generado en: $LogPath" -ForegroundColor Green
}
else {
    Write-Host "`nNo se encontraron direcciones con el dominio @$Dominio." -ForegroundColor Green
}

Disconnect-ExchangeOnline -Confirm:$false
