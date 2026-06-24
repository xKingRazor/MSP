# ==============================================================================
# CONFIGURACIÓN (Cambia estos valores según tu entorno)
# ==============================================================================
$tenantName = "tu-tenant" # Reemplaza "tu-tenant" con el nombre real de tu empresa
$adminCenterUrl = "https://$tenantName-admin.sharepoint.com"
$csvPath = "C:\temp\ReporteSitiosSharePoint.csv" # Asegúrate de que la carpeta "C:\temp" exista

# ==============================================================================
# EJECUCIÓN
# ==============================================================================

# 1. Iniciar sesión en el centro de administración (Te pedirá credenciales de Admin)
Write-Host "Conectando a SharePoint Online..." -ForegroundColor Cyan
Connect-SPOService -Url $adminCenterUrl

# 2. Obtener todos los sitios (Limit All asegura que traiga todo y no se detenga en 200)
Write-Host "Obteniendo todos los sitios. Esto puede tardar varios minutos..." -ForegroundColor Yellow
$sites = Get-SPOSite -Limit All

$resultados = @()

# 3. Recorrer cada sitio y mapear los datos solicitados
foreach ($site in $sites) {
    $resultados += [PSCustomObject]@{
        "Nombre"              = $site.Title
        "Propietarios"        = $site.Owner
        "Url"                 = $site.Url
        "Espacio en uso (MB)" = $site.StorageUsageCurrent
        "Espacio límite (MB)" = $site.StorageQuota
        "Template"            = $site.Template
        # Nota: El módulo estándar no extrae el GUID del sitio directamente para sitios puros, 
        # pero GroupId nos da el ID único para todos los sitios conectados a grupos de M365.
        "Id del sitio (Grupo)"= $site.GroupId 
    }
}

# 4. Exportar los datos a un archivo CSV
# Se usa delimitador ";" para que Excel en español lo separe en columnas correctamente
$resultados | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

Write-Host "¡Exportación finalizada con éxito! Puedes encontrar tu archivo en: $csvPath" -ForegroundColor Green