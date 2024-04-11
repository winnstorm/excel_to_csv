Write-Host "Bienvenido al Conversor de Archivos XLSX a CSV"
Write-Host "Gabxcel2CSV v0.1"

$rutaScript = Split-Path -Parent $MyInvocation.MyCommand.Path
$directorioEntrada = $rutaScript
$directorioSalida = $rutaScript

do {
    $nombreArchivoEntrada = Read-Host "Ingrese el nombre del archivo Excel de entrada (sin extension) "
    $rutaCompletaEntrada = "$directorioEntrada\$nombreArchivoEntrada.xlsx"

    if (-not (Test-Path $rutaCompletaEntrada)) {
        Write-Host "El archivo de entrada especificado no existe en el directorio de entrada."
    }
} while (-not (Test-Path $rutaCompletaEntrada))

$infoArchivoEntrada = Get-Item $rutaCompletaEntrada

Write-Host "Se ha encontrado el archivo de entrada:"
Write-Host "Nombre: $($infoArchivoEntrada.Name)"
Write-Host "Tipo: $($infoArchivoEntrada.Extension)"
Write-Host "Size: $($infoArchivoEntrada.Length) bytes"

$confirmacion = Read-Host "¿Desea continuar con este archivo? (S/N) "

if ($confirmacion -ne "S") {
    Write-Host "Operacion cancelada."
    Exit
}

$nombreBaseArchivoSalida = Read-Host "Ingrese el nombre base para los archivos CSV de salida "

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

$workbook = $excel.Workbooks.Open($rutaCompletaEntrada)

$totalHojas = $workbook.Sheets.Count
$hojasProcesadas = 0

$progreso = 0
Write-Progress -Activity "Convirtiendo archivos..." -Status "Procesando hojas..." -PercentComplete $progreso

foreach ($sheet in $workbook.Sheets) {
    $nombreHoja = $sheet.Name

    $nombreArchivoCSV = Join-Path -Path $directorioSalida -ChildPath "$nombreBaseArchivoSalida`_$nombreHoja.csv"

    $sheet.SaveAs($nombreArchivoCSV, 6)  # 6 representa el formato CSV
    $sheet.SaveAs($nombreArchivoCSV, 65001)  # 65001 representa el formato UTF-8

    $hojasProcesadas++
    $progreso = ($hojasProcesadas / $totalHojas) * 100
    Write-Progress -Activity "Convirtiendo archivos..." -Status "Procesando hojas..." -PercentComplete $progreso
}

$workbook.Close()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "La conversion de archivos XLSX a CSV ha sido completada."

$rutaSalidaCompleta = Join-Path -Path $directorioSalida -ChildPath "$nombreBaseArchivoSalida`_*.csv"
Write-Host "Los archivos CSV de salida se encuentran en la siguiente ruta:"
Write-Host $rutaSalidaCompleta

Invoke-Item -Path $directorioSalida