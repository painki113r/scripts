# Importar la librería iTextSharp
Add-Type -Path "C:\temp\clasifica\powers\lib\itextsharp.dll"

# Ruta de la carpeta que contiene los archivos PDF
$carpetaPDFs = "C:\temp\clasifica\tst"

# Palabras clave para clasificar los PDFs
$palabrasClavePagar = "Libro", "Busqueda", "Ruta"
$palabrasClaveRetener = "Nombre", "Apellido", "Telefono"

# Rutas de las carpetas de destino
$rutaDestinoPagar = "C:\temp\clasifica\tst\pagar"
$rutaDestinoRetener = "C:\temp\clasifica\tst\oficios"
$rutaDestinoOtras = "C:\temp\clasifica\tst\otros"

# Comprobar si las carpetas de destino existen, si no, crearlas
if (!(Test-Path -Path $rutaDestinoPagar)) {
    New-Item -Path $rutaDestinoPagar -ItemType Directory
}
if (!(Test-Path -Path $rutaDestinoRetener)) {
    New-Item -Path $rutaDestinoRetener -ItemType Directory
}
if (!(Test-Path -Path $rutaDestinoOtras)) {
    New-Item -Path $rutaDestinoOtras -ItemType Directory
}

# Recorrer todos los archivos PDF en la carpeta
$archivosPDF = Get-ChildItem -Path $carpetaPDFs -Filter "*.pdf"
foreach ($archivoPDF in $archivosPDF) {
    # Crear un objeto PDFReader
    $pdfReader = New-Object iTextSharp.text.pdf.PdfReader($archivoPDF.FullName)
    
    # Inicializar el texto
    $textoPDF = ""
    
    # Recorrer las páginas del PDF
    for ($pagina = 1; $pagina -le $pdfReader.NumberOfPages; $pagina++) {
        $textoPDF += [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdfReader, $pagina)
    }
    
    # Verificar palabras clave para Pagar
    $cumplePalabrasClavePagar = $true
    foreach ($palabraClave in $palabrasClavePagar) {
        if (!($textoPDF -match $palabraClave)) {
            $cumplePalabrasClavePagar = $false
            break
        }
    }
    
    # Verificar palabras clave para Retener
    $cumplePalabrasClaveRetener = $true
    foreach ($palabraClave in $palabrasClaveRetener) {
        if (!($textoPDF -match $palabraClave)) {
            $cumplePalabrasClaveRetener = $false
            break
        }
    }
    
    # Mover el archivo a la carpeta correspondiente o a "Otras"
    if ($cumplePalabrasClavePagar) {
        #Move-Item -Path $archivoPDF.FullName -Destination $rutaDestinoPagar
        Copy-Item -Path $archivoPDF.FullName -Destination $rutaDestinoPagar
    }
    elseif ($cumplePalabrasClaveRetener) {
        #Move-Item -Path $archivoPDF.FullName -Destination $rutaDestinoRetener
        Copy-Item -Path $archivoPDF.FullName -Destination $rutaDestinoRetener
    }
    else {
        #Move-Item -Path $archivoPDF.FullName -Destination $rutaDestinoOtras
        Copy-Item -Path $archivoPDF.FullName -Destination $rutaDestinoOtras
    }
    
    # Cerrar el objeto PDFReader
    $pdfReader.Close()
}
