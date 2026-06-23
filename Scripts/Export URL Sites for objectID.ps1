# ==========================================
# CONFIGURACIÓN DE RUTAS Y CREDENCIALES
# ==========================================
$TenantId      = "TU-TENANT-ID-AQUÍ"          # ID del directorio de Azure
$ClientId      = "TU-CLIENT-ID-AQUÍ"          # ID de la aplicación (Client ID)
$ClientSecret  = "TU-CLIENT-SECRET-AQUÍ"      # El valor del secreto generado

# Rutas de los archivos (Ajusta estas rutas según tu caso)
$PathTxtInput  = "C:\Ruta\De\Tu\Archivo\grupos.txt"   # Archivo de texto con los IDs (uno por línea)
$PathCsvOutput = "C:\Ruta\De\Tu\Archivo\resultados_sitios.csv" # Archivo donde se guardará el resultado

# ==========================================
# 1. VALIDACIÓN DE ARCHIVO DE ENTRADA
# ==========================================
if (-not (Test-Path $PathTxtInput)) {
    Write-Error "No se encontró el archivo de texto en la ruta especificada: $PathTxtInput"
    exit
}

# Leer los IDs del archivo TXT (elimina líneas vacías)
$GroupIds = Get-Content -Path $PathTxtInput | Where-Object { $_ -match '\S' }

Write-Host "Se encontraron $($GroupIds.Count) IDs para procesar.`n" -ForegroundColor Cyan

# ==========================================
# 2. CONEXIÓN AUTENTICADA (Contexto de Aplicación)
# ==========================================
$SecretSecure = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$Credential   = New-Object System.Management.Automation.PSCredential($ClientId, $SecretSecure)

Write-Host "Conectando a Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -TenantId $TenantId -Credential $Credential

# ==========================================
# 3. PROCESAMIENTO EN BUCLE Y EXTRACCIÓN
# ==========================================
# Creamos un arreglo vacío para almacenar los resultados
$Resultados = @()

foreach ($GroupId in $GroupIds) {
    Write-Host "Procesando Grupo ID: $GroupId ... " -NoNewline -ForegroundColor Yellow
    
    try {
        # Intentamos obtener el sitio de SharePoint asociado
        $SiteUrl = (Get-MgGroupSite -GroupId $GroupId -SiteId "root").WebUrl
        
        # Guardamos el resultado exitoso
        $Resultados += [PSCustomObject]@{
            GroupId  = $GroupId
            SitioUrl = $SiteUrl
            Estado   = "Éxito"
        }
        Write-Host "¡Encontrado!" -ForegroundColor Green
    }
    catch {
        # Si da error (por ejemplo, el ID no existe o no tiene Teams), guardamos el fallo
        $Resultados += [PSCustomObject]@{
            GroupId  = $GroupId
            SitioUrl = "N/A"
            Estado   = "Error: $($_.Exception.Message)"
        }
        Write-Host "Error" -ForegroundColor Red
    }
}

# ==========================================
# 4. EXPORTAR RESULTADOS Y DESCONECTAR
# ==========================================
Write-Host "`nExportando resultados a: $PathCsvOutput" -ForegroundColor Cyan
# Exportamos a CSV con codificación UTF8 para evitar problemas de caracteres
$Resultados | Export-Csv -Path $PathCsvOutput -NoTypeInformation -Encoding UTF8 -Delimiter ","

Write-Host "Proceso finalizado por completo." -ForegroundColor Green

Disconnect-MgGraph | Out-Null
