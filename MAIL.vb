Public Sub ActualizarCatalogos()

    Dim rutaCatalogo As String
    Dim wbCatalogo As Workbook
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim hoja As Variant

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    rutaCatalogo = "\\172.20.0.22\unidad_u\DOCUMENTOS\CAMBIO DE EQUIPOS\AUTOMATIZADO\DATA_SYNC\iTopDatos.xlsx"

    Set wbCatalogo = Workbooks.Open(rutaCatalogo, ReadOnly:=True)

    For Each hoja In Array("dataPersonas", "dataDispositivos")

        Set wsOrigen = wbCatalogo.Sheets(CStr(hoja))
        Set wsDestino = ThisWorkbook.Sheets(CStr(hoja))

        wsDestino.Cells.Clear

        wsOrigen.UsedRange.Copy

        wsDestino.Range("A1").PasteSpecial xlPasteValues

    Next hoja

    wbCatalogo.Close False

    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    MsgBox "Datos actualizados correctamente.", vbInformation

Exit Sub

ErrorHandler:

    Application.ScreenUpdating = True

    If Not wbCatalogo Is Nothing Then
        wbCatalogo.Close False
    End If

    MsgBox "No fue posible actualizar los catálogos." & vbCrLf & _
           "Verifique la conexión a la carpeta compartida.", vbCritical

End Sub






Function GenerarNroActa() As String

    Dim wbHist As Workbook
    Set wbHist = Workbooks.Open("\\172.20.0.22\Unidad_U\DOCUMENTOS\CAMBIO DE EQUIPOS\AUTOMATIZADO\DATA_SYNC\iTopHistorico.xlsx")
    

    Dim ws As Worksheet
    Dim ultimo As Long
    Dim correlativo As Long

    Set ws = wbHist.Sheets("Actas")

    ultimo = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If ultimo < 2 Then
        correlativo = 1
    Else
        correlativo = ultimo
    End If

    GenerarNroActa = "ACT-" & _
                     Format(Date, "yyyymmdd") & "-" & _
                     Format(correlativo, "0000")

End Function

Sub GenerarPDF(ByVal nroActa As String)

    Dim Ruta As String
    Dim NombrePDF As String

    'Carpeta donde guardar
    Ruta = "\\172.20.0.22\unidad_u\DOCUMENTOS\CAMBIO DE EQUIPOS\AUTOMATIZADO\ACTAS_FIRMADAS\"

    'Nombre del archivo
    NombrePDF = nroActa & ".pdf"

    'Exportar
    ThisWorkbook.Sheets("Formulario").ExportAsFixedFormat _
        Type:=0, _
        Filename:=Ruta & NombrePDF, _
        Quality:=0, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
End Sub

Sub GuardarActa()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim wbHist As Workbook
    Set wbHist = Workbooks.Open("\\172.20.0.22\Unidad_U\DOCUMENTOS\CAMBIO DE EQUIPOS\AUTOMATIZADO\DATA_SYNC\iTopHistorico.xlsx")
    Dim wsForm As Worksheet
    Dim wsActas As Worksheet
    Dim wsDetalle As Worksheet

    Dim nroActa As String
    Dim filaCab As Long
    Dim filaDet As Long

    Dim i As Long

    Set wsForm = ThisWorkbook.Sheets("Formulario")
    Set wsActas = wbHist.Sheets("Actas")
    Set wsDetalle = wbHist.Sheets("Detalles")

    '==========================================
    ' GENERAR NÚMERO DE ACTA
    '==========================================
    
    nroActa = GenerarNroActa()
    wsForm.Range("H4").Value = nroActa

    '==========================================
    ' GUARDAR CABECERA
    '==========================================

    filaCab = wsActas.Cells(wsActas.Rows.Count, 1).End(xlUp).Row + 1

    wsActas.Cells(filaCab, 1).Value = nroActa
    wsActas.Cells(filaCab, 2).Value = Now
    wsActas.Cells(filaCab, 3).Value = wsForm.Range("C9").Value
    wsActas.Cells(filaCab, 4).Value = wsForm.Range("C10").Value
    wsActas.Cells(filaCab, 5).Value = wsForm.Range("C11").Value
    wsActas.Cells(filaCab, 6).Value = wsForm.Range("C12").Value
    wsActas.Cells(filaCab, 7).Value = wsForm.Range("C13").Value

    '==========================================
    ' DETALLE ASIGNACIONES
    '==========================================

    For i = 17 To 23

        If Trim(wsForm.Cells(i, 1).Value) <> "" Then

            filaDet = wsDetalle.Cells(wsDetalle.Rows.Count, 1).End(xlUp).Row + 1

            wsDetalle.Cells(filaDet, 1).Value = nroActa
            wsDetalle.Cells(filaDet, 2).Value = "ASIGNACION"
            wsDetalle.Cells(filaDet, 3).Value = wsForm.Range("H10").Value
            wsDetalle.Cells(filaDet, 4).Value = wsForm.Cells(i, 1).Value
            wsDetalle.Cells(filaDet, 5).Value = wsForm.Cells(i, 3).Value
            wsDetalle.Cells(filaDet, 6).Value = wsForm.Cells(i, 4).Value
            wsDetalle.Cells(filaDet, 7).Value = wsForm.Cells(i, 5).Value
            wsDetalle.Cells(filaDet, 8).Value = wsForm.Cells(i, 6).Value
            wsDetalle.Cells(filaDet, 9).Value = wsForm.Range("A25").Value

        End If

    Next i

    '==========================================
    ' DETALLE DEVOLUCIONES
    '==========================================

    For i = 30 To 36

        If Trim(wsForm.Cells(i, 1).Value) <> "" Then

            filaDet = wsDetalle.Cells(wsDetalle.Rows.Count, 1).End(xlUp).Row + 1

            wsDetalle.Cells(filaDet, 1).Value = nroActa
            wsDetalle.Cells(filaDet, 2).Value = "DEVOLUCION"
            wsDetalle.Cells(filaDet, 3).Value = wsForm.Range("H10").Value
            wsDetalle.Cells(filaDet, 4).Value = wsForm.Cells(i, 1).Value
            wsDetalle.Cells(filaDet, 5).Value = wsForm.Cells(i, 3).Value
            wsDetalle.Cells(filaDet, 6).Value = wsForm.Cells(i, 4).Value
            wsDetalle.Cells(filaDet, 7).Value = wsForm.Cells(i, 5).Value
            wsDetalle.Cells(filaDet, 8).Value = wsForm.Cells(i, 6).Value
            wsDetalle.Cells(filaDet, 9).Value = wsForm.Range("A38").Value

        End If

    Next i

    '==========================================
    ' GUARDAR Y CERRAR HISTÓRICO
    '==========================================

    wbHist.Save
    wbHist.Close SaveChanges:=False

    Set wsDetalle = Nothing
    Set wsActas = Nothing
    Set wbHist = Nothing

    '==========================================
    ' GENERAR DOCUMENTO PDF
    '==========================================
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Acta registrada correctamente: " & nroActa, vbInformation
    Call GenerarPDF(nroActa)
End Sub

Function ObtenerRutaPDF(ByVal nroActa As String) As String

    Dim Ruta As String
    Dim NumeroActa As String

    NumeroActa = nroActa

    Ruta = "\\172.20.0.22\Unidad_U\DOCUMENTOS\CAMBIO DE EQUIPOS\AUTOMATIZADO\ACTAS_FIRMADAS\"

    ObtenerRutaPDF = Ruta & NumeroActa & ".pdf"

End Function

Sub EnviarMail(ByVal nroActa As String)

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim RutaPDF As String


    RutaPDF = ObtenerRutaPDF(nroActa)

    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)

    With OutlookMail

        .To = ThisWorkbook.Sheets("Formulario").Range("I10").Value

        .Subject = "Acta de Asignación/Devolución de Equipo - " & nroActa

        .Body = "Estimados," & vbCrLf & vbCrLf & _
                "Usuario: " & ThisWorkbook.Sheets("Formulario").Range("C10").Value & vbCrLf & _
                "Acta: " & nroActa & vbCrLf & vbCrLf & _
                "Se adjunta el acta de asignación/devolución generada para poder gestionar la firma digital." & vbCrLf & _
                "Una vez firmado, favor enviarlo por este medio." & vbCrLf & vbCrLf & _
                "Saludos."

        .Attachments.Add RutaPDF

        .Display
        '.Send

    End With

End Sub