# Define la ruta del registro
$regPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server"
# Obtener la fecha actual
$currentDate = Get-Date -Format "dd/MM/yyyy"
# Obtener el nombre del equipo
$ComputerName = $env:COMPUTERNAME
# Cargar el ensamblado de SQL Client
[System.Reflection.Assembly]::LoadWithPartialName("System.Data.SqlClient") | Out-Null

# Función para verificar el estado del servicio SQL Server
function Test-SQLService {
    param (
        [string]$instanceName
    )
    
    $serviceName = if ($instanceName -eq "MSSQLSERVER") {
        "MSSQLSERVER"
    } else {
        "MSSQL`$$instanceName"
    }
    
    try {
        $service = Get-Service -Name $serviceName -ErrorAction Stop
        return @{
            Status = $service.Status
            Running = $service.Status -eq 'Running'
        }
    }
    catch {
        return @{
            Status = "No encontrado"
            Running = $false
        }
    }
}

# Función para obtener el collation desde el registro
function Get-SQLCollationFromRegistry {
    param (
        [string]$instanceName
    )
    
    try {
        # Obtener el nombre real de la instancia desde el registro
        $setupValues = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL" -ErrorAction Stop
        $instanceID = $setupValues.$instanceName
        
        if ($instanceID) {
            # Construir la ruta al registro donde se almacena el collation
            $collationPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$instanceID\Setup"
            
            if (Test-Path $collationPath) {
                $setupInfo = Get-ItemProperty -Path $collationPath -ErrorAction Stop
                if ($setupInfo.Collation) {
                    return $setupInfo.Collation
                }
            }
            
            # Ruta alternativa para el collation
            $collationPath2 = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$instanceID\MSSQLServer"
            if (Test-Path $collationPath2) {
                $collationInfo = Get-ItemProperty -Path $collationPath2 -ErrorAction Stop
                if ($collationInfo.Collation) {
                    return $collationInfo.Collation
                }
            }
        }
        return "No se pudo determinar"
    }
    catch {
        Write-Warning "Error al leer el collation del registro: $($_.Exception.Message)"
        return "Error al leer registro"
    }
}

# Función para obtener detalles de espacio de las instancias
function Get-SQLInstanceSpaceDetails {
    param (
        [string]$instanceName,
        [string]$instancePath
    )
    $mdfTotalSize = 0
    $ldfTotalSize = 0
    $bakTotalSize = 0
    $instanceSize = 0

    Write-Host "Buscando archivos para la instancia $instanceName en: $instancePath"

    try {
        # Calcular el tamaño total de archivos .mdf
        $mdfFiles = Get-ChildItem -Path $instancePath -Filter "*.mdf" -File -Recurse -ErrorAction Stop
        $mdfTotalSize = ($mdfFiles | Measure-Object -Property Length -Sum).Sum

        # Calcular el tamaño total de archivos .ldf
        $ldfFiles = Get-ChildItem -Path $instancePath -Filter "*.ldf" -File -Recurse -ErrorAction Stop
        $ldfTotalSize = ($ldfFiles | Measure-Object -Property Length -Sum).Sum

        # Calcular el tamaño total de archivos .bak
        $bakFiles = Get-ChildItem -Path $instancePath -Filter "*.bak" -File -Recurse -ErrorAction Stop
        $bakTotalSize = ($bakFiles | Measure-Object -Property Length -Sum).Sum

        # Calcular el tamaño total de la instancia
        $instanceSize = $mdfTotalSize + $ldfTotalSize + $bakTotalSize
    }
    catch {
        Write-Warning "Error al acceder a los archivos de la instancia $instanceName : $($_.Exception.Message)"
    }

    # Convertir bytes a GB y formatear
    $mdfGB = if ($mdfTotalSize -gt 0) { "{0:N2} GB" -f ($mdfTotalSize / 1GB) } else { "0.00 GB" }
    $ldfGB = if ($ldfTotalSize -gt 0) { "{0:N2} GB" -f ($ldfTotalSize / 1GB) } else { "0.00 GB" }
    $bakGB = if ($bakTotalSize -gt 0) { "{0:N2} GB" -f ($bakTotalSize / 1GB) } else { "0.00 GB" }
    $instanceSizeGB = if ($instanceSize -gt 0) { "{0:N2} GB" -f ($instanceSize / 1GB) } else { "0.00 GB" }

    return @{
        'MDFSize' = $mdfGB
        'LDFSize' = $ldfGB
        'BAKSize' = $bakGB
        'InstanceSize' = $instanceSizeGB
    }
}

# Crear contenido HTML
$htmlhead = @"
<html>
 <title>Instancias SQL - $ComputerName</title>
 <style>
 BODY { font-family: Arial; font-size: 8pt; }
 H1 { font-size: 20px; }
 H2 { font-size: 18px; }
 H3 { font-size: 16px; }
 TABLE { border: 1px solid black; border-collapse: collapse; font-size: 9pt; width: 90%; table-layout: fixed; word-wrap: break-word; }
 TH { border: 1px solid black; background: #00609c; padding: 5px; color: white; }
 TD { border: 1px solid black; padding: 5px; }
 td.pass { background: #7FFF00; }
 td.warn { background: #FFE600; }
 td.fail { background: #FF0000; color: #ffffff; }
 td.info { background: #85D4FF; }
 </style>
 <body>
 <div align='center'>
 <table>
 <tr>
 <th colspan='9'><h1>Información de Instancias SQL en el Servidor: $ComputerName</h1></th>
 </tr>
 <tr>
 <th colspan='9'><h2>Generado el: $currentDate</h2></th>
 </tr>
 </table>
"@
$htmlbody = ""

# Recolecta información de SQL Server
try {
    if (Test-Path $regPath) {
        # Obtener las instancias instaladas de SQL Server
        $instances = Get-ItemProperty -Path $regPath -Name InstalledInstances -ErrorAction SilentlyContinue

        if ($instances) {
            $htmlbody += @"
 <table>
 <tr>
 <th>Nombre de Instancia</th>
 <th>Estado del Servicio</th>
 <th>Versión</th>
 <th>Edición</th>
 <th>Collation</th>
 <th>Ruta de Instalación</th>
 <th>Tamaño .mdf</th>
 <th>Tamaño .ldf</th>
 <th>Tamaño .bak</th>
 <th>Tamaño Total</th>
 </tr>
"@

            foreach ($instance in $instances.InstalledInstances) {
                # Obtener el estado del servicio
                $serviceStatus = Test-SQLService -instanceName $instance
                $statusClass = switch ($serviceStatus.Status) {
                    'Running' { 'pass' }
                    'Stopped' { 'fail' }
                    default { 'warn' }
                }

                # Obtener información específica de la instancia
                $setupValues = Get-ItemProperty -Path "$regPath\Instance Names\SQL" -ErrorAction SilentlyContinue
                if ($setupValues) {
                    $instanceName = $setupValues.$instance
                    $setupPath = "$regPath\$instanceName\Setup"
                    if (Test-Path $setupPath) {
                        $setup = Get-ItemProperty -Path $setupPath -ErrorAction SilentlyContinue

                        # Obtener el Collation desde el registro
                        $collation = Get-SQLCollationFromRegistry -instanceName $instance

                        # Obtener información de espacio
                        $spaceInfo = Get-SQLInstanceSpaceDetails -instanceName $instance -instancePath $setup.SQLPath

                        $htmlbody += @"
 <tr>
 <td>$instance</td>
 <td class='$statusClass'>$($serviceStatus.Status)</td>
 <td>$($setup.Version)</td>
 <td>$($setup.Edition)</td>
 <td>$collation</td>
 <td>$($setup.SQLPath)</td>
 <td>$($spaceInfo.MDFSize)</td>
 <td>$($spaceInfo.LDFSize)</td>
 <td>$($spaceInfo.BAKSize)</td>
 <td>$($spaceInfo.InstanceSize)</td
 </tr>
"@
                    }
                }
            }

            $htmlbody += "</table>"
        } else {
            $htmlbody += "<p>No se encontraron instancias de SQL Server instaladas.</p>"
        }
    } else {
        $htmlbody += "<p>No se encontró el registro de SQL Server en el sistema.</p>"
    }
}
catch {
    Write-Warning $_.Exception.Message
    $htmlbody += "<p>Se ha encontrado un error al leer la información de SQL Server. $($_.Exception.Message)</p>"
}

$htmltail = @"
 </div>
 </body>
</html>
"@
$htmlreport = $htmlhead + $htmlbody + $htmltail

# Crear el nombre del archivo
$fileName = "$ComputerName-InstanciasTODO.html"

# Guardado en el escritorio
try {
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $htmlFilePath = Join-Path -Path $desktopPath -ChildPath $fileName
    $htmlreport | Out-File -FilePath $htmlFilePath -Encoding UTF8 -Force

    # Verificar si el archivo se creó correctamente
    if (Test-Path $htmlFilePath) {
        Write-Host "Archivo creado exitosamente en: $htmlFilePath"
        # Abrir el archivo HTML en el navegador predeterminado
        Start-Process $htmlFilePath
    } else {
        throw "No se pudo verificar la creación del archivo"
    }
}
catch {
    # Si falla el escritorio, intentar en el directorio actual
    Write-Warning "No se pudo guardar en el escritorio. Intentando en el directorio actual..."
    try {
        $currentPath = Get-Location
        $htmlFilePath = Join-Path -Path $currentPath -ChildPath $fileName
        $htmlreport | Out-File -FilePath $htmlFilePath -Encoding UTF8 -Force

        if (Test-Path $htmlFilePath) {
            Write-Host "Archivo creado exitosamente en: $htmlFilePath"
            # Abrir el archivo HTML en el navegador predeterminado
            Start-Process $htmlFilePath
        } else {
            throw "No se pudo crear el archivo en ninguna ubicación"
        }
    }
    catch {
        Write-Error "Error al crear el archivo: $($_.Exception.Message)"
    }
}