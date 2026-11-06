# ============================================================
# iTop - DEVOLUCION MASIVA DE ACTIVOS
# Usuario: astrid.tejerina
# ============================================================

$ApiUrl = "https://yaocmiitopq2.ypfb-andina.com.bo/webservices/rest.php?version=1.3"

$ApiUser = Read-Host "Usuario iTop"
$ApiPass = Read-Host "Password iTop" -AsSecureString

$ApiPassPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ApiPass)
)

# ============================================================
# DATOS DE LA DEVOLUCION
# ============================================================

$LoginUsuario = "adriana.mercado"

$NuevaLocalidad = 9

$NuevaObservacion = "Devolución 26/10/2026 SIGDEV26339-ADRIANA.MERCADO -Conclusión Laboral"

# ============================================================
# FUNCION API
# ============================================================

function Invoke-iTopAPI {

    param(
        [hashtable]$Data
    )

    $Json = $Data | ConvertTo-Json -Depth 20 -Compress

    $Body = @{
        auth_user = $ApiUser
        auth_pwd  = $ApiPassPlain
        json_data = $Json
    }

    try {

        return Invoke-RestMethod `
            -Uri $ApiUrl `
            -Method Post `
            -Body $Body `
            -ContentType "application/x-www-form-urlencoded" `
            -ErrorAction Stop

    }
    catch {

        Write-Host ""
        Write-Host "ERROR DE CONEXION CON iTOP:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        return $null
    }
}

# ============================================================
# BUSCAR TODOS LOS ACTIVOS al usuario
# ============================================================

$OQL = @"
SELECT PhysicalDevice AS d
JOIN Person AS p ON d.pdasigned_id = p.id
JOIN User AS u ON u.contactid = p.id
WHERE u.login = '$LoginUsuario'
"@

Write-Host ""
Write-Host "============================================================"
Write-Host "BUSCANDO ACTIVOS ASIGNADOS A $LoginUsuario"
Write-Host "============================================================"
Write-Host ""

$Consulta = @{
    operation     = "core/get"
    class         = "PhysicalDevice"
    key           = $OQL
    output_fields = "*+"
}

$Resultado = Invoke-iTopAPI -Data $Consulta

if (-not $Resultado) {
    Write-Host "No se obtuvo respuesta de iTop." -ForegroundColor Red
    exit
}

if ($Resultado.code -ne 0) {

    Write-Host ""
    Write-Host "ERROR DE iTOP:" -ForegroundColor Red
    Write-Host $Resultado.message
    exit
}

if (-not $Resultado.objects) {

    Write-Host ""
    Write-Host "No se encontraron activos asignados a $LoginUsuario." -ForegroundColor Yellow
    exit
}

$Objetos = @($Resultado.objects.PSObject.Properties)

# ============================================================
# MOSTRAR ACTIVOS ENCONTRADOS
# ============================================================

Write-Host "ACTIVOS ENCONTRADOS: $($Objetos.Count)" -ForegroundColor Green
Write-Host ""

$Lista = foreach ($Objeto in $Objetos) {

    $Item = $Objeto.Value

    [PSCustomObject]@{
        ID        = $Item.key
        Clase     = $Item.class
        Equipo    = $Item.fields.name
        Serie     = $Item.fields.serialnumber
        Modelo    = $Item.fields.model_name
        Estado    = $Item.fields.status
        Localidad = $Item.fields.location_name
        Asignado  = $Item.fields.pdasigned_id_friendlyname
    }
}

$Lista | Format-Table -AutoSize

# ============================================================
# CONFIRMACION
# ============================================================

Write-Host ""
Write-Host "============================================================"
Write-Host "CAMBIOS QUE SE REALIZARAN"
Write-Host "============================================================"
Write-Host ""
Write-Host "Asignado a  -> VACIO"
Write-Host "Estatus     -> Disponible"
Write-Host "Localidad   -> DEP. BLOQUE C (ID 9)"
Write-Host "Contacto    -> VACIO"
Write-Host "Observacion -> $NuevaObservacion"
Write-Host ""

$Confirmar = Read-Host "Escriba SI para ejecutar la devolucion de TODOS los activos"

if ($Confirmar -ne "SI") {

    Write-Host ""
    Write-Host "Operacion cancelada. No se modifico ningun activo." -ForegroundColor Yellow
    exit
}

# ============================================================
# PROCESAR ACTIVOS
# ============================================================

$Correctos = 0
$Errores = 0

Write-Host ""
Write-Host "============================================================"
Write-Host "INICIANDO ACTUALIZACION"
Write-Host "============================================================"
Write-Host ""

foreach ($Objeto in $Objetos) {

    $Item = $Objeto.Value

    $ID = $Item.key
    $Clase = $Item.class
    $Nombre = $Item.fields.name

    Write-Host "Procesando: $Nombre | ID: $ID | Clase: $Clase ... " -NoNewline

    # --------------------------------------------------------
    # CAMPOS A MODIFICAR
    # --------------------------------------------------------

    $Campos = @{
        pdasigned_id  = 0
        location_id   = $NuevaLocalidad
        contacts_list = @()
        description   = $NuevaObservacion
        status        = "Disponible"
    }

    # --------------------------------------------------------
    # ACTUALIZACION
    # --------------------------------------------------------

    $Actualizacion = @{
        operation     = "core/update"
        class         = $Clase
        key           = $ID
        comment       = "Devolucion de equipo - liberacion de activo"
        fields        = $Campos
        output_fields = "name,pdasigned_id,pdasigned_id_friendlyname,status,location_id,location_name,contacts_list,description"
    }

    $Respuesta = Invoke-iTopAPI -Data $Actualizacion

    if ($Respuesta -and $Respuesta.code -eq 0) {

        Write-Host "OK" -ForegroundColor Green
        $Correctos++

    }
    else {

        Write-Host "ERROR" -ForegroundColor Red
        $Errores++

        if ($Respuesta) {
            Write-Host "    $($Respuesta.message)" -ForegroundColor Red
        }
    }
}

# ============================================================
# RESUMEN
# ============================================================

Write-Host ""
Write-Host "============================================================"
Write-Host "RESULTADO FINAL"
Write-Host "============================================================"
Write-Host ""
Write-Host "Activos encontrados : $($Objetos.Count)"
Write-Host "Actualizados        : $Correctos" -ForegroundColor Green
Write-Host "Con error           : $Errores" -ForegroundColor Red
Write-Host ""

if ($Errores -eq 0) {

    Write-Host "TODOS LOS ACTIVOS FUERON PROCESADOS CORRECTAMENTE." -ForegroundColor Green

}
else {

    Write-Host "ATENCION: algunos activos presentaron errores." -ForegroundColor Yellow
}

Write-Host ""