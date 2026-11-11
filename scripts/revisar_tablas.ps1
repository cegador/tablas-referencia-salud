# Revisa las paginas de INVIMA y Minsalud y descarga las tablas nuevas.
# Uso:  powershell -ExecutionPolicy Bypass -File scripts\revisar_tablas.ps1 [-Forzar]
#   -Forzar  revisa las dos fuentes aunque el calendario diga que hoy no toca.
# Deja el resumen en estado\ultimo_resultado.json. No envia correos ni toca el wiki.
# (Archivo solo con caracteres ASCII a proposito: Windows PowerShell 5.1 lo exige sin BOM.)
param([switch]$Forzar)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Add-Type -AssemblyName System.IO.Compression.FileSystem

$Raiz          = Split-Path -Parent $PSScriptRoot
$DirTablas     = Join-Path $Raiz 'tablas'
$DirHist       = Join-Path $Raiz 'historico'
$DirOriginales = Join-Path $Raiz 'originales'   # archivos tal cual se descargaron, cuando hubo que convertirlos a .xlsx
$DirEstado     = Join-Path $Raiz 'estado'
$DirTmp        = Join-Path $DirEstado 'tmp'
$RutaEstado    = Join-Path $DirEstado 'estado.json'
$RutaResultado = Join-Path $DirEstado 'ultimo_resultado.json'

$UA          = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
$UrlInvima   = 'https://www.invima.gov.co/productos-vigilados/medicamentos-y-productos-biologicos/autorizacion-de-comercializacion-registros'
$UrlMinsalud = 'https://www.minsalud.gov.co/salud/POS/Paginas/unidad-de-pago-por-capitacion-upc.aspx'

$Meses = @('enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre')
$TiposInvima = [ordered]@{
    vigentes     = 'Medicamentos Vigentes'
    vencidos     = 'Medicamentos Vencidos'
    renovacion   = 'Medicamentos en Tramite de Renovacion'
    otrosestados = 'Medicamentos Otros Estados'
}
# Tablas de Minsalud que nos interesan y como reconocerlas dentro del zip
$TablasMinsalud = [ordered]@{
    cups  = @{ nombre = 'CUPS';   patron = '*CUPS*' }
    cie10 = @{ nombre = 'CIE-10'; patron = '*CIE-10*' }
    reps  = @{ nombre = 'REPS';   patron = '*REPS*' }
}

function Decir([string]$Texto) {
    # Avance en pantalla, para quien este mirando la ejecucion
    Write-Host ("[{0}] {1}" -f (Get-Date).ToString('HH:mm:ss'), $Texto)
}

function Descargar([string]$Url, [string]$Destino, [string]$Que) {
    Decir "  Descargando $Que ..."
    $reloj = [Diagnostics.Stopwatch]::StartNew()
    $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -UserAgent $UA -TimeoutSec 900 -OutFile $Destino -PassThru
    $mb = (Get-Item $Destino).Length / 1MB
    Decir ("  Listo: {0:N1} MB en {1:N0} s" -f $mb, $reloj.Elapsed.TotalSeconds)
    return $resp
}

function Obtener-Pagina([string]$Url) {
    (Invoke-WebRequest -Uri $Url -UseBasicParsing -UserAgent $UA -TimeoutSec 120).Content
}

function Es-ArchivoOffice([string]$Ruta) {
    # xlsx, xlsb y zip empiezan con los bytes "PK"
    if (-not (Test-Path $Ruta)) { return $false }
    if ((Get-Item $Ruta).Length -lt 10000) { return $false }
    $fs = [IO.File]::OpenRead($Ruta)
    try { $b1 = $fs.ReadByte(); $b2 = $fs.ReadByte() } finally { $fs.Close() }
    return ($b1 -eq 0x50 -and $b2 -eq 0x4B)
}

function Poner($Objeto, [string]$Nombre, $Valor) {
    # Crea o cambia un dato del estado (los objetos leidos del JSON no aceptan datos nuevos por asignacion)
    $Objeto | Add-Member -NotePropertyName $Nombre -NotePropertyValue $Valor -Force
}

function Leer($Objeto, [string]$Nombre) {
    if ($Objeto.PSObject.Properties[$Nombre]) { return $Objeto.$Nombre }
    return $null
}

function Reemplazar-Tabla($Actual, [string]$CarpetaHistorico, [string]$RutaNueva, [string]$NombreNuevo, [datetime]$Hoy) {
    # Mueve la tabla vieja (y su original, si fue convertida) al historico y deja la nueva en "tablas"
    $viejos = @()
    if ($Actual.archivo) { $viejos += (Join-Path $DirTablas $Actual.archivo) }
    $origViejo = Leer $Actual 'original'
    if ($origViejo -and $origViejo -ne $Actual.archivo) { $viejos += (Join-Path $DirOriginales $origViejo) }
    foreach ($v in $viejos) {
        if (Test-Path $v) {
            New-Item -ItemType Directory -Force -Path $CarpetaHistorico | Out-Null
            Move-Item -Force -Path $v -Destination (Join-Path $CarpetaHistorico (Split-Path $v -Leaf))
        }
    }
    $destino = Join-Path $DirTablas $NombreNuevo
    Move-Item -Force -Path $RutaNueva -Destination $destino
    $Actual.archivo = $NombreNuevo
    $Actual.descargado = $Hoy.ToString('yyyy-MM-dd')
    Poner $Actual 'original' $NombreNuevo
    Poner $Actual 'hash_original' (Get-FileHash $destino -Algorithm SHA256).Hash
}

function Convertir-AXlsx([string]$Origen, [string]$Destino) {
    # Convierte un archivo de Excel (.xlsb, .xls, .xlsm) a .xlsx. Usa Excel si esta instalado; si no, LibreOffice.
    $excel = $null; $libro = $null
    try { $excel = New-Object -ComObject Excel.Application } catch { $excel = $null }
    if ($excel) {
        try {
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $excel.AutomationSecurity = 3          # no ejecutar macros
            $libro = $excel.Workbooks.Open($Origen, 0, $true)   # sin actualizar vinculos, solo lectura
            $libro.SaveAs($Destino, 51)            # 51 = libro .xlsx
        }
        finally {
            if ($libro) { $libro.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($libro) }
            $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
    }
    else {
        $soffice = @("$env:ProgramFiles\LibreOffice\program\soffice.exe", "${env:ProgramFiles(x86)}\LibreOffice\program\soffice.exe") |
            Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
        if (-not $soffice) { throw 'este computador no tiene Excel ni LibreOffice para hacer la conversion' }
        $carpeta = Split-Path $Destino -Parent
        $p = Start-Process -FilePath $soffice -ArgumentList @('--headless', '--convert-to', 'xlsx', '--outdir', "`"$carpeta`"", "`"$Origen`"") -Wait -PassThru -WindowStyle Hidden
        $generado = Join-Path $carpeta ([IO.Path]::GetFileNameWithoutExtension($Origen) + '.xlsx')
        if ($generado -ne $Destino -and (Test-Path $generado)) { Move-Item -Force $generado $Destino }
    }
    if (-not (Es-ArchivoOffice $Destino)) { throw 'la conversion no produjo un archivo .xlsx valido' }
}

function Asegurar-Xlsx($estado) {
    # Revisa que todas las tablas vigentes esten en .xlsx; convierte las que no y guarda el original aparte
    $res = [ordered]@{ conversiones = @(); avisos = @() }
    $entradas = @()
    foreach ($k in $TiposInvima.Keys)    { $entradas += $estado.invima.archivos.$k }
    foreach ($k in $TablasMinsalud.Keys) { $entradas += $estado.minsalud.archivos.$k }

    foreach ($a in $entradas) {
        if ($a.archivo -match '\.xlsx$') { continue }
        $origen = Join-Path $DirTablas $a.archivo
        if (-not (Test-Path $origen)) { $res.avisos += "No se encuentra el archivo $($a.archivo) en la carpeta tablas."; continue }
        $nombreXlsx = [IO.Path]::GetFileNameWithoutExtension($a.archivo) + '.xlsx'
        $destino = Join-Path $DirTablas $nombreXlsx
        try {
            $hash = (Get-FileHash $origen -Algorithm SHA256).Hash
            if (Test-Path $destino) { Remove-Item $destino -Force }
            Convertir-AXlsx $origen $destino
            New-Item -ItemType Directory -Force -Path $DirOriginales | Out-Null
            Move-Item -Force -Path $origen -Destination (Join-Path $DirOriginales $a.archivo)
            $res.conversiones += [ordered]@{ original = $a.archivo; convertido = $nombreXlsx }
            Poner $a 'original' $a.archivo
            Poner $a 'hash_original' $hash
            $a.archivo = $nombreXlsx
        }
        catch {
            if (Test-Path $destino) { Remove-Item $destino -Force -ErrorAction SilentlyContinue }
            $res.avisos += "La tabla $($a.archivo) no esta en .xlsx y no se pudo convertir: $($_.Exception.Message)"
        }
    }
    return $res
}

function Cerrar-Resultado($r) {
    if ($r.errores.Count -gt 0 -and $r.novedades.Count -gt 0) { $r.estado = 'actualizado_con_errores' }
    elseif ($r.errores.Count -gt 0) { $r.estado = 'error' }
    elseif ($r.novedades.Count -gt 0) { $r.estado = 'actualizado' }
    return $r
}

# ---------------------------------------------------------------- INVIMA
function Revisar-Invima($estado, [datetime]$hoy) {
    $r = [ordered]@{ fuente = 'INVIMA'; url = $UrlInvima; estado = 'sin_cambios'; detalle = ''; novedades = @(); errores = @() }
    $mesActual = $hoy.ToString('yyyy-MM')

    $pendientes = @($TiposInvima.Keys | Where-Object { $estado.invima.archivos.$_.mes -lt $mesActual })
    # Tablas que el estado dice tener pero no estan en la carpeta (por ejemplo, copia recien descargada de GitHub)
    $faltan = @($TiposInvima.Keys | Where-Object { -not (Test-Path (Join-Path $DirTablas $estado.invima.archivos.$_.archivo)) })
    if ($faltan.Count -gt 0) { $Forzar = $true }
    if (-not $Forzar -and $pendientes.Count -eq 0) {
        $r.estado = 'omitido'
        $r.detalle = 'Ya se tienen los cuatro listados de este mes; no hace falta revisar.'
        return $r
    }
    if (-not $Forzar -and $estado.invima.ultima_revision -eq $hoy.ToString('yyyy-MM-dd')) {
        $r.estado = 'omitido'
        $r.detalle = 'El INVIMA ya se reviso hoy.'
        return $r
    }

    Decir 'INVIMA: abriendo la pagina de registros ...'
    try { $html = Obtener-Pagina $UrlInvima }
    catch {
        $r.errores += "La pagina del INVIMA no funciona: $($_.Exception.Message)"
        return (Cerrar-Resultado $r)
    }
    $estado.invima.ultima_revision = $hoy.ToString('yyyy-MM-dd')

    # Enlaces del tipo /biblioteca/listadocodigounicovigentesoctubre2026xlsx
    $rx = [regex]("(?i)/biblioteca/listadocodigounico(" + ($TiposInvima.Keys -join '|') + ")(" + ($Meses -join '|') + ")(\d{4})xlsx")
    $publicados = @{}
    foreach ($m in $rx.Matches($html)) {
        $tipo = $m.Groups[1].Value.ToLower()
        $numMes = [array]::IndexOf($Meses, $m.Groups[2].Value.ToLower()) + 1
        $mes = '{0}-{1:00}' -f $m.Groups[3].Value, $numMes
        if (-not $publicados.ContainsKey($tipo) -or $mes -gt $publicados[$tipo].mes) {
            $publicados[$tipo] = @{ mes = $mes; enlace = $m.Value; etiqueta = ('{0} {1}' -f $m.Groups[2].Value.ToLower(), $m.Groups[3].Value) }
        }
    }

    foreach ($tipo in $TiposInvima.Keys) {
        $nombreTabla = "Listado Codigo Unico de $($TiposInvima[$tipo])"
        if (-not $publicados.ContainsKey($tipo)) {
            $r.errores += "No se encontro en la pagina el enlace de '$nombreTabla'. Puede que la pagina haya cambiado."
            continue
        }
        $pub = $publicados[$tipo]
        $actual = $estado.invima.archivos.$tipo
        if ($pub.mes -le $actual.mes -and ($faltan -notcontains $tipo)) { continue }

        try {
            $fichaUrl = 'https://www.invima.gov.co' + $pub.enlace
            $ficha = Obtener-Pagina $fichaUrl
            $md = [regex]::Match($ficha, '(?i)/biblioteca/download/\d+')
            if (-not $md.Success) { throw "no se encontro el boton de descarga en $fichaUrl" }

            New-Item -ItemType Directory -Force -Path $DirTmp | Out-Null
            $tmp = Join-Path $DirTmp "invima_$tipo.tmp"
            $resp = Descargar ('https://www.invima.gov.co' + $md.Value) $tmp "$($TiposInvima[$tipo]) ($($pub.etiqueta))"
            if (-not (Es-ArchivoOffice $tmp)) { throw 'lo descargado no es un archivo de Excel valido' }

            $nombre = $null
            $cd = [string]$resp.Headers['Content-Disposition']
            $mn = [regex]::Match($cd, 'filename="?([^";]+)"?')
            if ($mn.Success) { $nombre = $mn.Groups[1].Value.Trim() }
            if (-not $nombre -or $nombre -notmatch '\.xls[xb]$') {
                $nombre = 'ListadoCodigoUnico_{0}_{1}.xlsx' -f $tipo, $pub.mes
            }

            $nov = [ordered]@{
                tabla = $nombreTabla; clave = $tipo
                anterior = $actual.archivo; version_anterior = $actual.mes
                nuevo = $nombre; version_nueva = $pub.mes; etiqueta = $pub.etiqueta
            }
            Reemplazar-Tabla $actual (Join-Path (Join-Path $DirHist 'invima') $actual.mes) $tmp $nombre $hoy
            Decir "  Guardada en tablas: $nombre"
            $r.novedades += $nov
            $actual.mes = $pub.mes
        }
        catch {
            $r.errores += "No se pudo descargar '$nombreTabla' ($($pub.etiqueta)): $($_.Exception.Message)"
        }
    }

    if ($r.novedades.Count -eq 0 -and $r.errores.Count -eq 0) {
        $r.detalle = "El INVIMA todavia no ha publicado listados posteriores a los que ya tenemos."
    }
    return (Cerrar-Resultado $r)
}

# -------------------------------------------------------------- MINSALUD
function Revisar-Minsalud($estado, [datetime]$hoy) {
    $r = [ordered]@{ fuente = 'Minsalud'; url = $UrlMinsalud; estado = 'sin_cambios'; detalle = ''; novedades = @(); errores = @() }

    # De diciembre a febrero se revisa en cada ejecucion; el resto del ano, una vez al mes.
    $temporada = @(12, 1, 2) -contains $hoy.Month
    $yaEsteMes = ($estado.minsalud.ultima_revision -and $estado.minsalud.ultima_revision.Substring(0, 7) -eq $hoy.ToString('yyyy-MM'))
    # Tablas que el estado dice tener pero no estan en la carpeta (por ejemplo, copia recien descargada de GitHub)
    $faltan = @($TablasMinsalud.Keys | Where-Object { -not (Test-Path (Join-Path $DirTablas $estado.minsalud.archivos.$_.archivo)) })
    if ($faltan.Count -gt 0) { $Forzar = $true }
    if (-not $Forzar -and $estado.minsalud.ultima_revision -eq $hoy.ToString('yyyy-MM-dd')) {
        $r.estado = 'omitido'
        $r.detalle = 'Minsalud ya se reviso hoy.'
        return $r
    }
    if (-not $Forzar -and -not $temporada -and $yaEsteMes) {
        $r.estado = 'omitido'
        $r.detalle = 'Las tablas anuales ya se revisaron este mes.'
        return $r
    }

    Decir 'Minsalud: abriendo la pagina de UPC ...'
    try { $html = Obtener-Pagina $UrlMinsalud }
    catch {
        $r.errores += "La pagina de Minsalud no funciona: $($_.Exception.Message)"
        return (Cerrar-Resultado $r)
    }

    $enlace = $null; $vigencia = 0
    foreach ($m in [regex]::Matches($html, '(?i)href="([^"]*tablas-vigencia-(\d{4})\.zip)"')) {
        $anio = [int]$m.Groups[2].Value
        if ($anio -gt $vigencia) { $vigencia = $anio; $enlace = $m.Groups[1].Value }
    }
    if (-not $enlace) {
        $r.errores += "No se encontro en la pagina el enlace 'Tablas de referencia - Vigencia'. Puede que la pagina haya cambiado."
        return (Cerrar-Resultado $r)
    }
    $zipUrl = (New-Object Uri((New-Object Uri($UrlMinsalud)), $enlace)).AbsoluteUri

    try {
        $cab = Invoke-WebRequest -Uri $zipUrl -Method Head -UseBasicParsing -UserAgent $UA -TimeoutSec 120
        $modificado = [string]$cab.Headers['Last-Modified']
    }
    catch {
        $r.errores += "No se pudo consultar el archivo de tablas de Minsalud: $($_.Exception.Message)"
        return (Cerrar-Resultado $r)
    }
    $estado.minsalud.ultima_revision = $hoy.ToString('yyyy-MM-dd')

    $vigenciaVieja = [int]$estado.minsalud.vigencia
    if ($vigencia -lt $vigenciaVieja) {
        $r.errores += "La pagina ofrece la vigencia $vigencia, anterior a la que ya tenemos ($vigenciaVieja)."
        return (Cerrar-Resultado $r)
    }
    if ($vigencia -eq $vigenciaVieja -and $modificado -eq $estado.minsalud.zip_modificado -and $faltan.Count -eq 0) {
        $r.detalle = "Sigue publicada la vigencia $vigencia, sin cambios desde la ultima revision."
        return $r
    }

    # Hay vigencia nueva o el zip de la misma vigencia fue actualizado: se descarga y se compara tabla por tabla
    $zip = $null
    try {
        New-Item -ItemType Directory -Force -Path $DirTmp | Out-Null
        $tmpZip = Join-Path $DirTmp 'minsalud.zip'
        [void](Descargar $zipUrl $tmpZip "el paquete de tablas de referencia (vigencia $vigencia)")
        if (-not (Es-ArchivoOffice $tmpZip)) { throw 'lo descargado no es un archivo zip valido' }
        $zip = [IO.Compression.ZipFile]::OpenRead($tmpZip)

        $histMinsalud = Join-Path $DirHist 'minsalud'
        if ($vigencia -gt $vigenciaVieja) { $carpetaHist = Join-Path $histMinsalud "$vigenciaVieja" }
        else { $carpetaHist = Join-Path $histMinsalud ("{0}_reemplazado_{1}" -f $vigenciaVieja, $hoy.ToString('yyyy-MM-dd')) }

        foreach ($clave in $TablasMinsalud.Keys) {
            $def = $TablasMinsalud[$clave]
            $entrada = $zip.Entries | Where-Object {
                $_.Name -like $def.patron -and $_.Name -notlike '._*' -and $_.FullName -notlike '__MACOSX*' -and $_.Name -match '\.xls[xb]$'
            } | Select-Object -First 1
            if (-not $entrada) {
                $r.errores += "El archivo de Minsalud (vigencia $vigencia) no trae la tabla $($def.nombre)."
                continue
            }
            $tmp = Join-Path $DirTmp "minsalud_$clave.tmp"
            [IO.Compression.ZipFileExtensions]::ExtractToFile($entrada, $tmp, $true)

            $actual = $estado.minsalud.archivos.$clave
            # Se compara contra el archivo tal cual se descargo (el original), no contra el convertido a .xlsx
            $nombreOriginal = Leer $actual 'original'
            if (-not $nombreOriginal) { $nombreOriginal = $actual.archivo }
            $hashOriginal = Leer $actual 'hash_original'
            if (-not $hashOriginal) {
                $rutaActual = Join-Path $DirTablas $actual.archivo
                if (Test-Path $rutaActual) { $hashOriginal = (Get-FileHash $rutaActual -Algorithm SHA256).Hash }
            }
            $igual = ($faltan -notcontains $clave -and $entrada.Name -eq $nombreOriginal -and $hashOriginal -and
                      (Get-FileHash $tmp -Algorithm SHA256).Hash -eq $hashOriginal)
            if ($igual) { Remove-Item $tmp -Force; continue }

            $nov = [ordered]@{
                tabla = $def.nombre; clave = $clave
                anterior = $actual.archivo; version_anterior = "vigencia $vigenciaVieja"
                nuevo = $entrada.Name; version_nueva = "vigencia $vigencia"; etiqueta = "vigencia $vigencia"
            }
            Reemplazar-Tabla $actual $carpetaHist $tmp $entrada.Name $hoy
            Decir "  Guardada en tablas: $($entrada.Name)"
            $r.novedades += $nov
        }

        if ($r.errores.Count -eq 0) {
            $estado.minsalud.vigencia = $vigencia
            $estado.minsalud.zip_modificado = $modificado
        }
        if ($r.novedades.Count -eq 0 -and $r.errores.Count -eq 0) {
            $r.detalle = "Minsalud actualizo el paquete de la vigencia $vigencia, pero CUPS, CIE-10 y REPS siguen iguales."
        }
    }
    catch {
        $r.errores += "No se pudieron descargar las tablas de Minsalud: $($_.Exception.Message)"
    }
    finally {
        if ($zip) { $zip.Dispose() }
    }
    return (Cerrar-Resultado $r)
}

# ------------------------------------------------------------- PRINCIPAL
$hoy = Get-Date
$estado = Get-Content $RutaEstado -Raw -Encoding UTF8 | ConvertFrom-Json
New-Item -ItemType Directory -Force -Path $DirTablas | Out-Null

$resInvima   = Revisar-Invima $estado $hoy
$resMinsalud = Revisar-Minsalud $estado $hoy
$formato     = Asegurar-Xlsx $estado

if (Test-Path $DirTmp) { Remove-Item $DirTmp -Recurse -Force -ErrorAction SilentlyContinue }

$estado | ConvertTo-Json -Depth 8 | Set-Content -Path $RutaEstado -Encoding UTF8
$resultado = [ordered]@{
    fecha    = $hoy.ToString('yyyy-MM-dd HH:mm')
    forzado  = [bool]$Forzar
    fuentes  = @($resInvima, $resMinsalud)
    conversiones = @($formato.conversiones)
    avisos       = @($formato.avisos)
}
$resultado | ConvertTo-Json -Depth 8 | Set-Content -Path $RutaResultado -Encoding UTF8

# Resumen en pantalla y en estado\historial.log (sirve cuando el script corre solo, sin el agente)
$lineas = @()
foreach ($f in $resultado.fuentes) {
    $lineas += ("{0}: {1}. {2}" -f $f.fuente, $f.estado, $f.detalle)
    foreach ($n in $f.novedades) { $lineas += ("  NUEVO  {0}: {1} -> {2}" -f $n.tabla, $n.anterior, $n.nuevo) }
    foreach ($e in $f.errores)   { $lineas += ("  ERROR  {0}" -f $e) }
}
foreach ($c in $resultado.conversiones) { $lineas += ("Formato: {0} se convirtio a {1}" -f $c.original, $c.convertido) }
foreach ($a in $resultado.avisos)       { $lineas += ("AVISO de formato: {0}" -f $a) }
Write-Host ''
Write-Host '--- Resumen ---'
$lineas | Write-Output
$soloOmitidos = (@($resultado.fuentes | Where-Object { $_.estado -ne 'omitido' }).Count -eq 0) -and
                ($resultado.conversiones.Count -eq 0) -and ($resultado.avisos.Count -eq 0)
if (-not $soloOmitidos) {
    (@("[{0}]" -f $resultado.fecha) + $lineas) | Add-Content -Path (Join-Path $DirEstado 'historial.log') -Encoding UTF8
}
