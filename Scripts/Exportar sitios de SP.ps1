<#
.SYNOPSIS
    Exporta un inventario de los sitios de SharePoint Online usando Microsoft Graph
    con autenticacion app-only por CLIENT SECRET (Client Id + Tenant Id + Client Secret).

.DESCRIPTION
    Genera un CSV con: Nombre, URL, Tipo de sitio (plantilla raiz), Ultima modificacion,
    Espacio en uso (GB), Espacio total / cuota (GB), Id de sitio y Propietarios.

    IMPORTANTE - Por que Graph y no PnP/Get-PnPTenantSite:
    El metodo app-only de SharePoint con CLIENT SECRET (Azure ACS) fue deprecado por
    Microsoft y esta deshabilitado por defecto en tenants creados desde finales de 2024.
    Con un client secret la via soportada y estable es Microsoft Graph, que es la que
    usa este script.

    De donde sale cada dato:
    - Nombre, URL, Id de sitio y Ultima modificacion: del endpoint getAllSites de Graph.
    - Espacio en uso, espacio total/cuota, plantilla y propietario principal: del
      reporte "getSharePointSiteUsageDetail" (una sola llamada masiva), cruzado por URL.
    - TODOS los propietarios de sitios conectados a grupo M365: de los owners del grupo.

.REQUISITOS PREVIOS
    1. Modulos de Microsoft Graph PowerShell:
         Install-Module Microsoft.Graph -Scope CurrentUser
       (o como minimo: Microsoft.Graph.Authentication, Microsoft.Graph.Reports,
        Microsoft.Graph.Groups)

    2. App Registration en Entra ID con un CLIENT SECRET y permisos de APLICACION
       (con consentimiento de administrador) de Microsoft Graph:
         Reports.Read.All     -> reporte de uso de SharePoint
         Sites.Read.All       -> enumeracion y datos de sitios (getAllSites)
         Group.Read.All       -> propietarios de grupos M365

    3. En el Centro de administracion de M365: Settings > Org settings > Reports,
       DESACTIVAR "Mostrar nombres ocultos de usuario, grupo y sitio en todos los
       informes". Si esta activado, el reporte devuelve URLs y propietarios
       anonimizados (GUIDs) en lugar de los nombres reales.

.PARAMETER TenantId
    Directory (tenant) ID de la App Registration.

.PARAMETER ClientId
    Application (client) ID de la App Registration.

.PARAMETER ClientSecret
    Valor del client secret.

.PARAMETER RutaCsv
    Ruta del archivo CSV de salida.

.PARAMETER OmitirPropietariosGrupo
    Si se especifica, NO enumera los grupos M365 para sacar todos los propietarios
    (mas rapido). En ese caso se reporta solo el propietario principal del reporte.

.EXAMPLE
    .\Exportar-InventarioSharePoint-Graph.ps1 `
        -TenantId   "00000000-0000-0000-0000-000000000000" `
        -ClientId   "11111111-1111-1111-1111-111111111111" `
        -ClientSecret "tu-client-secret" `
        -RutaCsv "C:\Reportes\InventarioSharePoint.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]  [string]$TenantId,
    [Parameter(Mandatory = $true)]  [string]$ClientId,
    [Parameter(Mandatory = $true)]  [string]$ClientSecret,
    [Parameter(Mandatory = $false)] [string]$RutaCsv = ".\InventarioSharePoint_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
    [Parameter(Mandatory = $false)] [switch]$OmitirPropietariosGrupo
)

#--------------------------------------------------------------------
# Funciones auxiliares
#--------------------------------------------------------------------

function ConvertTo-GB {
    param([Parameter()][double]$Bytes)
    if (-not $Bytes) { return 0 }
    return [math]::Round($Bytes / 1GB, 2)
}

function Normalizar-Url {
    param([string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return "" }
    return $Url.TrimEnd('/').ToLowerInvariant()
}

#--------------------------------------------------------------------
# Validaciones iniciales
#--------------------------------------------------------------------

foreach ($m in @('Microsoft.Graph.Authentication','Microsoft.Graph.Reports','Microsoft.Graph.Groups')) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Error "Falta el modulo '$m'. Instala con: Install-Module Microsoft.Graph -Scope CurrentUser"
        return
    }
}

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Reports        -ErrorAction Stop
Import-Module Microsoft.Graph.Groups         -ErrorAction Stop

#--------------------------------------------------------------------
# Conexion app-only a Microsoft Graph con client secret
#--------------------------------------------------------------------

Write-Host "Conectando a Microsoft Graph (app-only, client secret)..." -ForegroundColor Cyan
try {
    $secure = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    $cred   = New-Object System.Management.Automation.PSCredential($ClientId, $secure)
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $cred -NoWelcome -ErrorAction Stop
    Write-Host "Conexion establecida correctamente." -ForegroundColor Green
}
catch {
    Write-Error "Error al conectar a Graph: $($_.Exception.Message)"
    return
}

#--------------------------------------------------------------------
# 1) Reporte de uso (almacenamiento, plantilla, owner principal) -> mapa por URL
#--------------------------------------------------------------------

Write-Host "Descargando el reporte de uso de sitios (getSharePointSiteUsageDetail)..." -ForegroundColor Cyan
$tmpCsv = Join-Path $env:TEMP "spo_usage_$(Get-Date -Format 'yyyyMMddHHmmss').csv"
$mapaStorage = @{}

try {
    # D7 = ventana de 7 dias. Otros valores: D30, D90, D180.
    Get-MgReportSharePointSiteUsageDetail -Period 'D7' -OutFile $tmpCsv -ErrorAction Stop
    $sitiosReporte = Import-Csv -Path $tmpCsv
    foreach ($s in $sitiosReporte) {
        $k = Normalizar-Url $s.'Site URL'
        if ($k) { $mapaStorage[$k] = $s }
    }
    Write-Host "Reporte cargado: $($mapaStorage.Count) sitios con metricas de almacenamiento." -ForegroundColor Green
}
catch {
    Write-Warning "No se pudo obtener el reporte de uso: $($_.Exception.Message). El almacenamiento saldra en blanco."
}
finally {
    if (Test-Path $tmpCsv) { Remove-Item $tmpCsv -Force -ErrorAction SilentlyContinue }
}

#--------------------------------------------------------------------
# 2) (Opcional) Propietarios de grupos M365 -> mapa por URL
#--------------------------------------------------------------------

$mapaOwnersGrupo = @{}

if (-not $OmitirPropietariosGrupo) {
    Write-Host "Enumerando grupos M365 para obtener todos los propietarios (puede tardar)..." -ForegroundColor Cyan
    try {
        $grupos = Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All -Property Id,DisplayName -ErrorAction Stop
        $g = 0
        foreach ($grupo in $grupos) {
            $g++
            Write-Progress -Activity "Procesando grupos M365" `
                -Status "$g de $($grupos.Count) - $($grupo.DisplayName)" `
                -PercentComplete (($g / [math]::Max($grupos.Count,1)) * 100)
            try {
                $site = Invoke-MgGraphRequest -Method GET `
                            -Uri "v1.0/groups/$($grupo.Id)/sites/root?`$select=webUrl" -ErrorAction Stop
                $url = $site.webUrl
                if ([string]::IsNullOrWhiteSpace($url)) { continue }

                $owners = Get-MgGroupOwner -GroupId $grupo.Id -All -ErrorAction Stop
                $lista = foreach ($o in $owners) {
                    $upn  = $o.AdditionalProperties['userPrincipalName']
                    $name = $o.AdditionalProperties['displayName']
                    if ($upn) { $upn } elseif ($name) { $name }
                }
                if ($lista) { $mapaOwnersGrupo[(Normalizar-Url $url)] = ($lista -join "; ") }
            }
            catch { continue }
        }
        Write-Progress -Activity "Procesando grupos M365" -Completed
        Write-Host "Propietarios mapeados para $($mapaOwnersGrupo.Count) sitios de grupo." -ForegroundColor Green
    }
    catch {
        Write-Warning "No se pudieron enumerar los grupos M365: $($_.Exception.Message). Se usara el propietario principal del reporte."
    }
}

#--------------------------------------------------------------------
# 3) Enumerar TODOS los sitios (nombre, url, id, ultima modificacion) via getAllSites
#--------------------------------------------------------------------

Write-Host "Enumerando todos los sitios (getAllSites)..." -ForegroundColor Cyan
$resultado     = New-Object System.Collections.Generic.List[object]
$urlsVistas    = New-Object System.Collections.Generic.HashSet[string]

function Procesar-Sitio {
    param($Nombre, $Url, $IdSitio, $UltimaModificacion)

    $urlNorm = Normalizar-Url $Url
    [void]$urlsVistas.Add($urlNorm)

    $rep = $null
    if ($mapaStorage.ContainsKey($urlNorm)) { $rep = $mapaStorage[$urlNorm] }

    $plantilla = if ($rep) { $rep.'Root Web Template' } else { $null }
    $esGrupo   = ($plantilla -like 'GROUP*')

    if ($mapaOwnersGrupo.ContainsKey($urlNorm)) {
        $propietarios = $mapaOwnersGrupo[$urlNorm]
    }
    elseif ($rep) {
        $propietarios = if ($rep.'Owner Principal Name') { $rep.'Owner Principal Name' } else { $rep.'Owner Display Name' }
    }
    else { $propietarios = $null }

    [PSCustomObject]@{
        Nombre              = $Nombre
        URL                 = $Url
        TipoSitio           = $plantilla
        ConectadoAGrupoM365 = if ($esGrupo) { "Si" } else { "No" }
        UltimaModificacion  = $UltimaModificacion
        EspacioUsoGB        = if ($rep) { ConvertTo-GB -Bytes ([double]$rep.'Storage Used (Byte)') } else { $null }
        EspacioTotalGB      = if ($rep) { ConvertTo-GB -Bytes ([double]$rep.'Storage Allocated (Byte)') } else { $null }
        IdSitio             = $IdSitio
        Propietarios        = $propietarios
    }
}

try {
    $uri = "v1.0/sites/getAllSites?`$select=id,name,displayName,webUrl,lastModifiedDateTime&`$top=200"
    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        foreach ($site in $resp.value) {
            $nombre = if ($site.displayName) { $site.displayName } elseif ($site.name) { $site.name } else { $site.webUrl }
            $resultado.Add( (Procesar-Sitio -Nombre $nombre -Url $site.webUrl -IdSitio $site.id -UltimaModificacion $site.lastModifiedDateTime) )
        }
        $uri = $resp.'@odata.nextLink'
    } while ($uri)
}
catch {
    Write-Warning "getAllSites no esta disponible o fallo: $($_.Exception.Message). Se construira el inventario solo desde el reporte de uso."
}

# Agregar sitios que esten en el reporte pero no aparecieron en getAllSites (p.ej. eliminados)
foreach ($s in $sitiosReporte) {
    $urlNorm = Normalizar-Url $s.'Site URL'
    if (-not $urlsVistas.Contains($urlNorm)) {
        $nombreDerivado = ($s.'Site URL' -split '/')[-1]
        $resultado.Add( (Procesar-Sitio -Nombre $nombreDerivado -Url $s.'Site URL' -IdSitio $s.'Site Id' -UltimaModificacion $s.'Last Activity Date') )
    }
}

#--------------------------------------------------------------------
# 4) Exportar a CSV
#--------------------------------------------------------------------

try {
    # UTF8 para acentos. Si tu Excel usa ";" como separador, cambia el -Delimiter ",".
    $resultado | Export-Csv -Path $RutaCsv -NoTypeInformation -Encoding UTF8 -Delimiter ","
    Write-Host "Inventario exportado a: $RutaCsv" -ForegroundColor Green
    Write-Host "Total de sitios exportados: $($resultado.Count)" -ForegroundColor Green
}
catch {
    Write-Error "Error al exportar el CSV: $($_.Exception.Message)"
}
finally {
    Disconnect-MgGraph | Out-Null
    Write-Host "Desconectado de Microsoft Graph." -ForegroundColor Cyan
}
