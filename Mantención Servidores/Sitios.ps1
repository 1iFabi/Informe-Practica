# Obtener la fecha actual
$currentDate = Get-Date -Format "dd/MM/yyyy"

# Obtener el nombre del equipo
$ComputerName = $env:COMPUTERNAME

# Crear contenido HTML
$htmlhead = @"
<html>
    <title>Sitios Web - $ComputerName</title>
    <style>
        BODY { font-family: Arial; font-size: 8pt; }
        H1 { font-size: 20px; }
        H2 { font-size: 18px; }
        TABLE { border: 1px solid black; border-collapse: collapse; font-size: 9pt; width: 90%; table-layout: fixed; word-wrap: break-word; }
        TH { border: 1px solid black; background: #00609c; padding: 5px; color: white; }
        TD { border: 1px solid black; padding: 5px; }
    </style>
    <body>
        <div align='center'>
            <h1>Información de Sitios Web en el Servidor: $ComputerName</h1>
            <h2>Generado el: $currentDate</h2>
            <table>
                <tr>
                    <th>Nombre del Sitio</th>
                    <th>Ruta Física</th>
                    <th>Espacio Utilizado</th>
                </tr>
"@

$htmlbody = ""

# Función para convertir bytes a un formato legible
function Format-FileSize {
    param ([long]$size)
    $sizes = 'B','KB','MB','GB','TB'
    $index = 0
    while ($size -gt 1024 -and $index -lt ($sizes.Count - 1)) {
        $size = $size / 1024
        $index++
    }
    return "{0:N2} {1}" -f $size, $sizes[$index]
}

# Función para obtener el tamaño de un directorio
function Get-FolderSize {
    param ([string]$path)
    try {
        $folderSize = Get-ChildItem -Path $path -Recurse -File -ErrorAction SilentlyContinue | 
            Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue
        return $folderSize.Sum
    }
    catch {
        return 0
    }
}

# Función para buscar sitios web en una ruta
function Find-WebSites {
    param (
        [string]$basePath,
        [string[]]$commonFolders = @('inetpub\wwwroot', 'Sitios')
    )

    $results = @()
    
    # Buscar en las carpetas comunes
    foreach ($folder in $commonFolders) {
        $searchPath = Join-Path -Path $basePath -ChildPath $folder
        if (Test-Path $searchPath) {
            Get-ChildItem -Path $searchPath -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $results += @{
                    Name = $_.Name
                    Path = $_.FullName
                }
            }
        }
    }
    
    return $results
}

# Obtener todos los discos del sistema
$disks = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
$sitesFound = $false

try {
    foreach ($disk in $disks) {
        $sites = Find-WebSites -basePath $disk.DeviceID
        
        foreach ($site in $sites) {
            $sitesFound = $true
            $siteSize = Get-FolderSize -path $site.Path
            $formattedSize = Format-FileSize -size $siteSize
            
            $htmlbody += @"
            <tr>
                <td>$($site.Name)</td>
                <td>$($site.Path)</td>
                <td>$formattedSize</td>
            </tr>
"@
        }
    }
    
    if (-not $sitesFound) {
        $htmlbody += "<tr><td colspan='3'>No se encontraron sitios web en ningún disco.</td></tr>"
    }
}
catch {
    Write-Warning $_.Exception.Message
    $htmlbody += "<tr><td colspan='3'>Error al buscar sitios web: $($_.Exception.Message)</td></tr>"
}

$htmltail = @"
            </table>
        </div>
    </body>
</html>
"@

$htmlreport = $htmlhead + $htmlbody + $htmltail

# Crear el nombre del archivo
$fileName = "$ComputerName-Sitios.html"

# Guardar en el escritorio
try {
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $htmlFilePath = Join-Path -Path $desktopPath -ChildPath $fileName
    $htmlreport | Out-File -FilePath $htmlFilePath -Encoding UTF8 -Force
    
    if (Test-Path $htmlFilePath) {
        Write-Host "Archivo creado exitosamente en: $htmlFilePath"
        Start-Process $htmlFilePath
    } else {
        throw "No se pudo verificar la creación del archivo"
    }
}
catch {
    Write-Warning "No se pudo guardar en el escritorio. Intentando en el directorio actual..."
    try {
        $currentPath = Get-Location
        $htmlFilePath = Join-Path -Path $currentPath -ChildPath $fileName
        $htmlreport | Out-File -FilePath $htmlFilePath -Encoding UTF8 -Force
        
        if (Test-Path $htmlFilePath) {
            Write-Host "Archivo creado exitosamente en: $htmlFilePath"
            Start-Process $htmlFilePath
        } else {
            throw "No se pudo crear el archivo en ninguna ubicación"
        }
    }
    catch {
        Write-Error "Error al crear el archivo: $($_.Exception.Message)"
    }
}