Attribute VB_Name = "modImpresion"
Option Explicit

Public jobnum As Integer
Public JobSRep1 As Integer
Public JobSRep2 As Integer
Public result As Integer
Public CantForm   As Integer
Public NombreFormula As String

'Bandejas de Impresoras----------------------------------
Public paINContadoB As Integer
Public paINContadoN As String
Public paPrintCtdoXDef As Boolean
Public paPrintCtdoPaperSize As Integer

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
            'If Format(RsAux!DocFecha, sqlFormatoF) <> Format(gFechaServidor, sqlFormatoF) Then HacerNota = True Else HacerAnulacion = True
            'cambio efactura elimine línea de arriba.
            HacerNota = True
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

Public Function ProcesoDocumentoFacturado(idDocumento As Long, HacerNota As Boolean, HacerAnulacion As Boolean, IdServicio As Long, _
                                        UsuarioSuceso As Long, DefensaSuceso As String, Optional DocumentoService As Long = 0) As Long
'   1) Graba la nota o anulacion, la asigna al documento y registra suceso
'   2) Registra mov. de caja y suceso
'   Retorna el Id de Nota si es que hay que hacerla

Dim auxTotal As Currency, auxIva As Currency
Dim aNumeroNota As String
Dim aTexto As String

Dim aIdNota As Long
Dim RsDoc As rdoResultset

    ProcesoDocumentoFacturado = 0
    If Not HacerNota And Not HacerAnulacion Then Exit Function
    If idDocumento = 0 And DocumentoService = 0 Then Exit Function
    If idDocumento = 0 Then idDocumento = DocumentoService
    auxTotal = 0: auxIva = 0
    
    Cons = "Select * from Documento Where DocCodigo = " & idDocumento
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    auxTotal = RsDoc!DocTotal: auxIva = RsDoc!DocIva
    
    aTexto = "Doc. " & Trim(RsDoc!DocSerie) & " " & Format(RsDoc!DocNumero, "000000")
    
    If HacerAnulacion Then aTexto = "Servicio " & IdServicio & ": Anulación del " & aTexto
    
    If HacerNota Then
        
        MsgBox "NO SE PUEDE HACER NOTA", vbExclamation, "ATENCIÓN"
        RsAux.Edit
        
        aNumeroNota = NumeroDocumento(paDNDevolucion)
        
        'Inserto la Nota con los datos del documento Original----------------------------------------------------------------------------
        Cons = "Select * from Documento Where DocCodigo = 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.AddNew
        RsAux!DocFecha = Format(gFechaServidor, sqlFormatoFH)
        RsAux!DocTipo = TipoDocumento.NotaDevolucion
        RsAux!DocSerie = Mid(aNumeroNota, 1, 1)
        RsAux!DocNumero = Mid(aNumeroNota, 2, Len(aNumeroNota))
        RsAux!DocCliente = RsDoc!DocCliente
        RsAux!DocMoneda = RsDoc!DocMoneda
        RsAux!DocTotal = RsDoc!DocTotal
        RsAux!DocIva = RsDoc!DocIva
        RsAux!DocAnulado = 0
        RsAux!DocSucursal = paCodigoDeSucursal
        RsAux!DocUsuario = UsuarioSuceso
        RsAux!DocFModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsAux.Update: RsAux.Close
        
        'Saco el id de la Nota---------------------------------------------------------------------------
        Cons = "Select * from Documento Where DocCodigo = (Select MAX(DocCodigo) From Documento)"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        aIdNota = RsAux!DocCodigo
        aTexto = "Servicio " & IdServicio & ": Nota " & Trim(RsAux!DocSerie) & " " & Format(RsAux!DocNumero, "000000") & " a " & aTexto
        RsAux.Close
        ProcesoDocumentoFacturado = aIdNota
        '-------------------------------------------------------------------------------------------------------------------------------------------
        
        Cons = "Select * From Renglon Where RenDocumento = " & idDocumento
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            Cons = "INSERT INTO Renglon (RenDocumento, RenArticulo, RenCantidad, RenPrecio, RenIVA, RenARetirar)" _
                    & " VALUES (" & aIdNota & ", " & RsAux!RenArticulo & ", " & RsAux!RenCantidad & ", " _
                                         & RsAux!RenPrecio & ", " & RsAux!RenIva & ", 0)"
            cBase.Execute Cons
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        Cons = "INSERT INTO Nota (NotFactura, NotNota, NotDevuelve, NotSalidaCaja)" _
            & " Values (" & idDocumento & "," & aIdNota & ", " & auxTotal & ", " & auxIva & ")"
        cBase.Execute Cons
    End If
        
    'Tengo que contrarestar la salida de caja
'    MovimientoDeCaja paMCVtaTelefonica, gFechaServidor, paDisponibilidad, RsDoc!DocMoneda, RsDoc!DocTotal, aTexto, False
    
    'Lo borro de la tabla de documentos pendientes
'    Cons = "Delete DocumentoPendiente Where DPeDocumento = " & idDocumento
'    cBase.Execute Cons

    Cons = "UPDATE DocumentoPendiente SET DPeIDLiquidacion = -1, DPeFLiquidacion = GetDate() WHERE DPeDocumento = " & idDocumento & " AND DPeFLiquidacion IS NULL"
    cBase.Execute Cons
    
    
    clsGeneral.RegistroSuceso cBase, gFechaServidor, 2, paCodigoDeTerminal, UsuarioSuceso, idDocumento, _
                        Descripcion:=aTexto, Defensa:=Trim(DefensaSuceso)
    
    If HacerAnulacion Then
        RsDoc.Edit
        RsDoc!DocAnulado = 1
        RsDoc.Update
    End If
    RsDoc.Close
    
    If idDocumento = DocumentoService Then  'Se trabajó con el del service
        Cons = "Update Servicio Set SerDocumento = Null Where SerCodigo = " & IdServicio
        cBase.Execute Cons
    End If
    
End Function

Public Sub ImprimoNota(idNota As Long, idDocumento As Long, idCliente As Long)

Dim RsCr As rdoResultset
Dim strCliente As String, strRuc As String, strDireccion As String
Dim aDocumento As String, aMonedaN As String, aMonedaS As String

On Error GoTo ErrCrystal

        
    'Saco Datos del cliente y Documento al que le hago la Nota--------------------------------------------------------------------------------------------------------------------
    Cons = "Select * From Cliente" _
                & " LEFT OUTER JOIN CPersona ON CliCodigo = CPeCliente" _
                & " LEFT OUTER JOIN CEmpresa ON CliCodigo = CEmCliente" _
        & " Where CliCodigo = " & idCliente
    Set RsCr = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
           
    If RsCr!CliTipo = TipoCliente.Empresa Then
        If Not IsNull(RsCr!CEmNombre) Then strCliente = Trim(RsCr!CEmNombre) Else strCliente = Trim(RsCr!CEmFantasia)
        If Not IsNull(RsCr!CliCIRuc) Then strRuc = clsGeneral.RetornoFormatoRuc(RsCr!CliCIRuc) Else strRuc = ""
    Else
        strCliente = ArmoNombre(Format(RsCr!CPeApellido1, "#"), Format(RsCr!CPeApellido2, "#"), Format(RsCr!CPeNombre1, "#"), Format(RsCr!CPeNombre2, "#"))
        If Not IsNull(RsCr!CliCIRuc) Then strCliente = strCliente & " (" & clsGeneral.RetornoFormatoCedula(RsCr!CliCIRuc) & ")"
        If Not IsNull(RsCr!CPERuc) Then strRuc = clsGeneral.RetornoFormatoRuc(RsCr!CPERuc) Else strRuc = ""
    End If
    If Not IsNull(RsCr!CliDireccion) Then strDireccion = clsGeneral.ArmoDireccionEnTexto(cBase, RsCr!CliDireccion) Else strDireccion = ""
    RsCr.Close
    
    Cons = "Select * from Documento, Moneda Where DocCodigo = " & idDocumento & " And DocMoneda = MonCodigo"
    Set RsCr = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsCr.EOF Then
        aDocumento = "Sobre Factura " & Trim(RsCr!DocSerie) & " " & Format(RsCr!DocNumero, "000000")
        aMonedaS = Trim(RsCr!MonSigno): aMonedaN = Trim(RsCr!MonNombre)
    End If
    RsCr.Close
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    jobnum = crAbroReporte(gPathListados & "NotaDevolucion.RPT")
    If jobnum = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora -------------------------------------------------------------------------------------------
    If Trim(Printer.DeviceName) <> Trim(paINContadoN) Then SeteoImpresoraPorDefecto paINContadoN
    If Not crSeteoImpresora(jobnum, Printer, paINContadoB) Then GoTo ErrCrystal
    
    'Obtengo la cantidad de formulas que tiene el reporte.
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    'Cargo Propiedades para el reporte -----------------------------------------------------------------------------------------------------
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "nombredocumento": result = crSeteoFormula(jobnum%, NombreFormula, "'" & paDNDevolucion & "'")
                
            Case "cliente":   result = crSeteoFormula(jobnum%, NombreFormula, "'" & strCliente & "'")
            Case "direccion": result = crSeteoFormula(jobnum%, NombreFormula, "'" & strDireccion & "'")
            Case "ruc": result = crSeteoFormula(jobnum%, NombreFormula, "'" & strRuc & "'")
            
            Case "codigobarras": result = crSeteoFormula(jobnum%, NombreFormula, "'" & CodigoDeBarras(TipoDocumento.NotaDevolucion, idNota) & "'")
            
            Case "signomoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'" & aMonedaS & "'")
            Case "nombremoneda": result = crSeteoFormula(jobnum%, NombreFormula, "'(" & aMonedaN & ")'")
            Case "textoretira": result = crSeteoFormula(jobnum%, NombreFormula, "'" & aDocumento & "'")     'Detallamos el documento al cual se le hace la nota.
            
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Documento.DocCodigo , Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor" _
            & " From " & paBD & ".dbo.Documento Documento " _
            & " Where DocCodigo = " & idNota
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
        
    'Subreporte srContado.rpt  y srContado.rpt - 01-----------------------------------------------------------------------------
    JobSRep1 = crAbroSubreporte(jobnum, "srContado.rpt")
    If JobSRep1 = 0 Then GoTo ErrCrystal
    
    Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion," _
            & " From { oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
                                  & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId}"
    If crSeteoSqlQuery(JobSRep1, Cons) = 0 Then GoTo ErrCrystal
    
    JobSRep2 = crAbroSubreporte(jobnum, "srContado.rpt - 01")
    If JobSRep2 = 0 Then GoTo ErrCrystal
    If crSeteoSqlQuery(JobSRep2, Cons) = 0 Then GoTo ErrCrystal
    '-------------------------------------------------------------------------------------------------------------------------------------
    
    If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
    'If crMandoAPantalla(JobNum, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal
    
    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal
    If Not crCierroSubReporte(JobSRep2) Then GoTo ErrCrystal
    crCierroTrabajo jobnum
    Screen.MousePointer = 0
    Exit Sub

ErrCrystal:
    clsGeneral.OcurrioError crMsgErr, Err.Description
    On Error Resume Next
    Screen.MousePointer = 11
    crCierroSubReporte JobSRep1
    crCierroSubReporte JobSRep2
    Screen.MousePointer = 0
End Sub

'------------------------------------------------------------------------------------------------------------------------------------
'   Carga los Nombres de los documentos y las impresoras (por defecto) para cada documento.
'   En la BD se guarda solo el nombre de la impresora (Device Name), y aca cargo los demas datos.
'       DriverName, Port y Bandeja
'------------------------------------------------------------------------------------------------------------------------------------
Public Sub CargoParametrosImpresion(Sucursal As Long)

    On Error GoTo errImp
    'Dado el Código de Sucursal se sacan los nombres de los documentos para cargar los parámetros.
    paDNDevolucion = "": paDNCredito = ""
    
    Cons = "Select * From Sucursal Where SucCodigo = " & Sucursal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        'Nombre de Cada Documento--------------------------------------------------------------------------------
        If Not IsNull(RsAux!SucDNDevolucion) Then paDNDevolucion = Trim(RsAux!SucDNDevolucion)
        If Not IsNull(RsAux!SucDNCredito) Then paDNCredito = Trim(RsAux!SucDNCredito)
    End If
    RsAux.Close
    Exit Sub
    
errImp:
    clsGeneral.OcurrioError "Error al cargar los parámetros de impresión. Informe del error a su administrador de base de datos.", Err.Description
    Screen.MousePointer = 0
End Sub
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
            .GetPrinterDoc 1, paINContadoN, paINContadoB, paPrintCtdoXDef, paPrintCtdoPaperSize
        End If
    End With
    If paPrintConfD = "" Then MsgBox "Por favor verifique la configuración de impresión.", vbInformation, "Atención"
    
errImp:
    Set objP = Nothing
    Screen.MousePointer = 0
End Sub


