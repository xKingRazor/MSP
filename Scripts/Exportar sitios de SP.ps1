<#
.SYNOPSIS
    Exporta un inventario de los sitios de SharePoint Online de un tenant.

.DESCRIPTION
    Genera un CSV con: Nombre, URL, Tipo de sitio (plantilla), Última modificación,
    Espacio en uso (GB), Espacio total / cuota (GB), Id de sitio y Propietarios.

    La autenticación es app-only (sin usuario interactivo) contra una App Registration
    de Entra ID (Azure AD) usando un CERTIFICADO. Es el método recomendado por
    Microsoft para operaciones de administración de SharePoint desatendidas.

.REQUISITOS PREVIOS
    1. Módulo PnP.PowerShell:
         Install-Module PnP.PowerShell -Scope CurrentUser

    2. App Registration en Entra ID (portal.azure.com > App registrations):
       - Subir un certificado (.cer) en "Certificates & secrets".
       - Conceder y dar consentimiento de administrador a estos permisos de APLICACIÓN:
           SharePoint  -> Sites.FullControl.All   (o Sites.Read.All si no usas -IncluirAdminsDeColeccion)
           Microsoft Graph -> Group.Read.All       (para leer propietarios de grupos M365)
           Microsoft Graph -> User.Read.All         (para resolver datos de usuario)
       - Anotar Application (client) ID y Directory (tenant) ID.

    3. El certificado (.pfx con su clave privada) debe estar instalado en el
       almacén de certificados del equipo donde corre el script (CurrentUser\My
       o LocalMachine\My) y se referencia por su Thumbprint.

.PARAMETER TenantName
    Nombre del tenant, p.ej. "banistmo" (la parte antes de .onmicrosoft.com)
    o el dominio completo "banistmo.onmicrosoft.com".

.PARAMETER AdminUrl
    URL del centro de administración de SharePoint, p.ej. https://banistmo-admin.sharepoint.com

.PARAMETER ClientId
    Application (client) ID de la App Registration.

.PARAMETER Thumbprint
    Huella digital (thumbprint) del certificado instalado localmente.

.PARAMETER RutaCsv
    Ruta del archivo CSV de salida.

.PARAMETER IncluirAdminsDeColeccion
    Si se especifica, en sitios NO conectados a grupo M365 se conecta a cada sitio
    para listar TODOS los administradores de colección de sitios (más lento).
    Sin este switch, en esos sitios se reporta solo el propietario principal.

.EXAMPLE
    .\Exportar-InventarioSharePoint.ps1 `
        -TenantName "banistmo" `
        -AdminUrl "https://banistmo-admin.sharepoint.com" `
        -ClientId "00000000-0000-0000-0000-000000000000" `
        -Thumbprint "A1B2C3D4E5F6...." `
        -RutaCsv "C:\Reportes\InventarioSharePoint.csv" `
        -IncluirAdminsDeColeccion
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantName,

    [Parameter(Mandatory = $true)]
    [string]$AdminUrl,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$Thumbprint,

    [Parameter(Mandatory = $false)]
    [string]$RutaCsv = ".\InventarioSharePoint_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter(Mandatory = $false)]
    [switch]$IncluirAdminsDeColeccion
)

#--------------------------------------------------------------------
# Funciones auxiliares
#--------------------------------------------------------------------

function ConvertTo-GB {
    # Recibe un valor en MB y lo devuelve en GB con 2 decimales
    param([Parameter()][double]$ValorMB)
    if ($null -eq $ValorMB) { return 0 }
    return [math]::Round($ValorMB / 1024, 2)
}

#--------------------------------------------------------------------
# Validaciones iniciales
#--------------------------------------------------------------------

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Error "No se encontró el módulo PnP.PowerShell. Instálalo con: Install-Module PnP.PowerShell -Scope CurrentUser"
    return
}

Import-Module PnP.PowerShell -ErrorAction Stop

#--------------------------------------------------------------------
# Conexión app-only al centro de administración de SharePoint
#--------------------------------------------------------------------

Write-Host "Conectando al tenant '$TenantName' (app-only, certificado)..." -ForegroundColor Cyan
try {
    Connect-PnPOnline -Url $AdminUrl `
        -ClientId $ClientId `
        -Thumbprint $Thumbprint `
        -Tenant $TenantName `
        -ErrorAction Stop
    Write-Host "Conexión establecida correctamente." -ForegroundColor Green
}
catch {
    Write-Error "Error al conectar: $($_.Exception.Message)"
    return
}

#--------------------------------------------------------------------
# Enumeración de sitios
#--------------------------------------------------------------------

Write-Host "Obteniendo el listado de sitios..." -ForegroundColor Cyan
try {
    # -Detailed trae StorageUsageCurrent, Owner, LastContentModifiedDate, GroupId, etc.
    $sitios = Get-PnPTenantSite -Detailed -ErrorAction Stop
}
catch {
    Write-Error "Error al obtener los sitios: $($_.Exception.Message)"
    Disconnect-PnPOnline
    return
}

$total = $sitios.Count
Write-Host "Se encontraron $total sitios. Procesando..." -ForegroundColor Cyan

$resultado = New-Object System.Collections.Generic.List[object]
$contador  = 0

foreach ($sitio in $sitios) {

    $contador++
    Write-Progress -Activity "Procesando sitios de SharePoint" `
        -Status "$contador de $total - $($sitio.Url)" `
        -PercentComplete (($contador / [math]::Max($total,1)) * 100)

    # ---- Propietarios ----
    $propietarios = @()

    $esGrupo = ($sitio.GroupId -and $sitio.GroupId -ne [Guid]::Empty)

    try {
        if ($esGrupo) {
            # Sitio conectado a un grupo M365: los propietarios son los owners del grupo
            $owners = Get-PnPMicrosoft365GroupOwner -Identity $sitio.GroupId.ToString() -ErrorAction Stop
            $propietarios = $owners | ForEach-Object {
                if ($_.Email) { $_.Email } elseif ($_.UserPrincipalName) { $_.UserPrincipalName } else { $_.DisplayName }
            }
        }
        elseif ($IncluirAdminsDeColeccion) {
            # Sitio clásico/comunicación: listar todos los admins de colección (requiere Sites.FullControl.All)
            Connect-PnPOnline -Url $sitio.Url -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $TenantName -ErrorAction Stop
            $admins = Get-PnPSiteCollectionAdmin -ErrorAction Stop
            $propietarios = $admins | ForEach-Object {
                if ($_.Email) { $_.Email } else { $_.LoginName }
            }
            # Volver al contexto de administración
            Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $TenantName -ErrorAction Stop
        }
        else {
            # Solo el propietario principal reportado por el tenant
            if ($sitio.Owner) { $propietarios = @($sitio.Owner) }
        }
    }
    catch {
        $propietarios = @("ERROR_PROPIETARIOS: $($_.Exception.Message)")
    }

    # ---- Id de sitio (puede no existir en versiones antiguas de PnP) ----
    $idSitio = $null
    if ($sitio.PSObject.Properties.Name -contains 'SiteId' -and $sitio.SiteId) {
        $idSitio = $sitio.SiteId.ToString()
    }

    # ---- Objeto de salida ----
    $resultado.Add([PSCustomObject]@{
        Nombre              = $sitio.Title
        URL                 = $sitio.Url
        TipoSitio           = $sitio.Template               # p.ej. GROUP#0, SITEPAGEPUBLISHING#0, STS#3
        ConectadoAGrupoM365 = if ($esGrupo) { "Sí" } else { "No" }
        UltimaModificacion  = $sitio.LastContentModifiedDate
        EspacioUsoGB        = ConvertTo-GB -ValorMB $sitio.StorageUsageCurrent
        EspacioTotalGB      = ConvertTo-GB -ValorMB $sitio.StorageQuota
        IdSitio             = $idSitio
        IdGrupoM365         = if ($esGrupo) { $sitio.GroupId.ToString() } else { $null }
        Propietarios        = ($propietarios -join "; ")
        Estado              = $sitio.Status
    })
}

Write-Progress -Activity "Procesando sitios de SharePoint" -Completed

#--------------------------------------------------------------------
# Exportación a CSV
#--------------------------------------------------------------------

try {
    # UTF8 con BOM para que los acentos se vean bien en Excel.
    # Si tu Excel usa ";" como separador, cambia -Delimiter "," por -Delimiter ";".
    $resultado | Export-Csv -Path $RutaCsv -NoTypeInformation -Encoding UTF8 -Delimiter ","
    Write-Host "Inventario exportado a: $RutaCsv" -ForegroundColor Green
    Write-Host "Total de sitios exportados: $($resultado.Count)" -ForegroundColor Green
}
catch {
    Write-Error "Error al exportar el CSV: $($_.Exception.Message)"
}
finally {
    Disconnect-PnPOnline
    Write-Host "Desconectado de SharePoint Online." -ForegroundColor Cyan
}
