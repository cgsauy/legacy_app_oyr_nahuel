Attribute VB_Name = "modImpresion"
Option Explicit

'Public jobnum As Integer
'Public JobSRep1 As Integer
'Public JobSRep2 As Integer
'Public Result As Integer
'Public CantForm   As Integer
'Public NombreFormula As String

'Bandejas de Impresoras----------------------------------
'Public paINContadoB As Integer
'Public paINContadoN As String
Public paPrintCtdoXDef As Boolean
'Public paPrintCtdoPaperSize As Integer

Public paPrintConfB As Integer      'Fichas de servicio (reimpresion)
Public paPrintConfD As String
Public paPrintConfPaperSize As Integer


Public Function ValidoDocumento(idDocumento As Long, HacerNota As Boolean, HacerAnulacion As Boolean, _
                                                DocumentoQFactura As Long, Optional DocumentoService As Long = 0) As Boolean
'   Valida si hay una factura pendiente para hacer el reclamo (para el id_Documento).
'   Dependiendo de la condición carga las variables HacerNota, HacerAnulacion y DocumentoQFactura para al grabar ver que se hace
'   Retorna: True o False para ver si continua con la accion

    On Error GoTo errValido
    Dim aTexto As String
    
    Screen.MousePointer = 11
    FechaDelServidor
    ValidoDocumento = True
    HacerNota = False: HacerAnulacion = False
    
    DocumentoQFactura = idDocumento
    'Valido Si el documento que factura el retiro esta en la tabla de pendientes para anular o emitir nota
    If idDocumento = 0 And DocumentoService = 0 Then Exit Function
    
    If idDocumento = 0 Then idDocumento = DocumentoService      'x si esta facturado todo en el doc. de la reparacion
    
    Cons = "Select * From DocumentoPendiente, Documento Where DPeDocumento = " & idDocumento & " And DPeDocumento = DocCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not RsAux!DocAnulado Then
            'Si es del mismo dia Anulo y Hago mov. Caja ------- Si no Hago Nota
            If Format(RsAux!DocFecha, sqlFormatoF) <> Format(gFechaServidor, sqlFormatoF) Then HacerNota = True Else HacerAnulacion = True
        End If
    End If
    RsAux.Close
    Screen.MousePointer = 0
    
    If HacerNota Then
        aTexto = "El servicio ya fue facturado (con distinta fecha al día de hoy)." & Chr(vbKeyReturn) & _
                      "Si ud. continúa con la operación, al grabar, se emitirá una nota para anular el documento." & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                      "Está seguro de continuar."
        If MsgBox(aTexto, vbQuestion + vbYesNo + vbDefaultButton2, "Emisión de Nota") = vbNo Then ValidoDocumento = False
    End If
    
    If HacerAnulacion Then
        aTexto = "El servicio ya fue facturado (en el día de hoy)." & Chr(vbKeyReturn) & _
                      "Si ud. continúa con la operación, al grabar, se emitirá un movimiento de caja para anular el documento." & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                      "Está seguro de continuar."
        If MsgBox(aTexto, vbQuestion + vbYesNo + vbDefaultButton2, "Anulación de Documento") = vbNo Then ValidoDocumento = False
    End If
    
    Exit Function

errValido:
    clsGeneral.OcurrioError "Ocurrió un error al validar el documento que factura el servicio.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function ArmoNombre(Ape1 As String, Ape2 As String, Nom1 As String, Nom2 As String) As String
    ArmoNombre = Trim(Ape1) & " " & Trim(Ape2)
    ArmoNombre = Trim(ArmoNombre) & ", " & Trim(Nom1) & " " & Trim(Nom2)
End Function

Public Sub prj_GetPrinter(ByVal bShowP As Boolean)
On Error GoTo errImp
    paPrintConfD = ""
    paPrintConfB = -1
    Dim objP As New clslPrintConfig
    With objP
        If bShowP Then
            If Not .ShowPrinterSetup("1,6", paCodigoDeTerminal) Then
                GoTo errImp
            End If
        End If
        If .LoadPrinterConfig("1,6", paCodigoDeTerminal) Then
            .GetPrinterDoc 6, paPrintConfD, paPrintConfB, paPrintCtdoXDef, paPrintConfPaperSize
        End If
    End With
    If paPrintConfD = "" Then MsgBox "Por favor verifique la configuración de impresión.", vbInformation, "Atención"
    
errImp:
    Set objP = Nothing
    Screen.MousePointer = 0
End Sub


