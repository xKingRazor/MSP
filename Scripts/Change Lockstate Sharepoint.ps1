$TenantAdminUrl = ""
$TxtPath        = ""

if (-not (Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
    Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction SilentlyContinue
}

if (-not (Test-Path -Path $TxtPath)) {
    Write-Error "No se encontro el archivo: $TxtPath"
    Exit
}

try {
    Write-Host "Conectando a SharePoint"
    Connect-SPOService -Url $TenantAdminUrl
}
catch {
    Write-Error "Error al conectar a SharePoint"
    Exit
}

$Sites = Get-Content -Path $TxtPath | Where-Object { $_ -match "^https://" }

Write-Host "`nIniciando Aplicado a ($($Sites.Count)) Sites"
Write-Host "--------------------------------------------------------"

foreach ($Url in $Sites) {
    $Url = $Url.Trim()
    
    Write-Host "Aplicando: $Url"
    
    try {
        Set-SPOSite -Identity $Url -LockState "ReadOnly" -ErrorAction Stop
        Write-Host "INFO: Configurado como ReadOnly"
    }
    catch {
        Write-Host "ERROR: No se pudo modificar el site: $($_.Exception.Message)"
    }
    Write-Host "--------------------------------------------------------"
}

Write-Host "`nFinalizado."
