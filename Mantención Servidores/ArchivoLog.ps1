# Ruta para guardar el archivo HTML
$outputPath = "$env:USERPROFILE\Desktop\Logs.html"

# Obtener la fecha actual y el nombre del equipo
$currentDate = Get-Date -Format "dd/MM/yyyy"
$ComputerName = $env:COMPUTERNAME

# Crear encabezado HTML 
$htmlHead = @"
<html>
    <title>Conteo de Archivos - $ComputerName</title>
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
        td.total { background: #00609c; color: white; font-weight: bold; }
    </style>
</html>
<body>
    <div align='center'>
        <h1>Conteo de Archivos Residuales por Disco</h1>
        <h3>Equipo: $ComputerName</h3>
        <h3>Fecha: $currentDate</h3>
        <table>
            <tr>
                <th>Disco</th>
                <th>Tipo de Archivo</th>
                <th>Tamaño Total</th>
            </tr>
"@

# Función para convertir tamaño a formato legible
function Format-FileSize {
    param ([double]$size)
    if ($size -gt 1TB) { return "{0:N2} TB" -f ($size / 1TB) }
    if ($size -gt 1GB) { return "{0:N2} GB" -f ($size / 1GB) }
    if ($size -gt 1MB) { return "{0:N2} MB" -f ($size / 1MB) }
    if ($size -gt 1KB) { return "{0:N2} KB" -f ($size / 1KB) }
    return "{0:N2} B" -f $size
}

$Discos = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -gt 0 } | Select-Object -ExpandProperty Root

# Extensiones a buscar
$Extensiones = @(".zip", ".bak", ".ldf")


$totalArchivos = 0
$totalesPorTipo = @{}
foreach ($ext in $Extensiones) {
    $totalesPorTipo[$ext] = @{
        cantidad = 0
        tamano = 0
    }
}
$granTotalTamano = 0

# cuerpo HTML
$htmlBody = ""

foreach ($disco in $Discos) {
    Write-Host "Procesando disco: $disco" -ForegroundColor Yellow
    
    foreach ($extension in $Extensiones) {
        try {
            $archivos = Get-ChildItem -Path $disco -Recurse -Filter "*$extension" -ErrorAction SilentlyContinue
            $cantidad = ($archivos | Measure-Object).Count
            $tamanoTotal = ($archivos | Measure-Object -Property Length -Sum).Sum
            
            if ($cantidad -gt 0) {
                $tamanoFormateado = Format-FileSize $tamanoTotal
                $htmlBody += @"
                    <tr>
                        <td>$disco</td>
                        <td class='info'>$extension</td>
                        <td>$tamanoFormateado</td>
                    </tr>
"@
                
                $totalArchivos += $cantidad
                $totalesPorTipo[$extension].cantidad += $cantidad
                $totalesPorTipo[$extension].tamano += $tamanoTotal
                $granTotalTamano += $tamanoTotal
            }
        }
        catch {
            Write-Warning "Error al procesar $disco : $_"
        }
    }
}

if ($totalArchivos -eq 0) {
    $htmlBody = "<tr><td colspan='3' class='warn'>No se encontraron archivos</td></tr>"
}
else {
    # Agregar totales por tipo de archivo
    foreach ($ext in $Extensiones) {
        $cantidadExt = $totalesPorTipo[$ext].cantidad
        $tamanoExt = Format-FileSize $totalesPorTipo[$ext].tamano
        $htmlBody += @"
            <tr>
                <td class='total' colspan='2'>Total archivos $ext : $cantidadExt</td>
                <td class='total'>$tamanoExt</td>
            </tr>
"@
    }
    
    $granTotalFormateado = Format-FileSize $granTotalTamano
    $htmlBody += @"
        <tr>
            <td class='total' colspan='2'>TOTAL GENERAL: $totalArchivos archivos</td>
            <td class='total'>$granTotalFormateado</td>
        </tr>
"@
}

$htmlFooter = @"
        </table>
    </div>
</body>
</html>
"@

# Combinar todo el HTML
$htmlContent = $htmlHead + $htmlBody + $htmlFooter

# Guardar en el archivo
$htmlContent | Out-File -FilePath $outputPath -Encoding UTF8

Write-Host "`nInforme generado exitosamente:" -ForegroundColor Green
Write-Host $outputPath -ForegroundColor Yellow
Write-Host "`nResumen de archivos encontrados:" -ForegroundColor Cyan
foreach ($ext in $Extensiones) {
    Write-Host "Total $ext : $($totalesPorTipo[$ext].cantidad) archivos - $(Format-FileSize $totalesPorTipo[$ext].tamano)" -ForegroundColor White
}
Write-Host "TOTAL GENERAL: $totalArchivos archivos - $(Format-FileSize $granTotalTamano)" -ForegroundColor White
