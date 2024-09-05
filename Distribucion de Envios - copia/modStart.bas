Attribute VB_Name = "modStart"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public paValorUIUltMes As Currency, paSendWApp As Byte

Public Enum TipoEstadoMercaderia
    Fisico = 1
    Virtual = 2
End Enum

Public Enum TipoMovimientoEstado
    ARetirar = 1
    AEntregar = 2
    Reserva = 3
End Enum

Public Enum TipoLocal
    Camion = 1
    Deposito = 2
End Enum

Public Enum eTiposDeTipoFlete
    Normales = 1
    CostoEspecial = 2
End Enum

Public printBandejaCopiaeTicket As Integer

'DOCUMENTOS
Public paDContado As String
Public paDRemito As String
Public paDNDevolucion As String

Public Enum TipoDocumento
    'Documentos Facturacion
    Contado = 1
    Credito = 2
    NotaDevolucion = 3
    NotaCredito = 4
    ReciboDePago = 5
    remito = 6
    ContadoDomicilio = 7
    CreditoDomicilio = 8
    ServicioDomicilio = 9
    NotaEspecial = 10
    
    'Documentos de Compras
    Compracontado = 11
    CompraCredito = 12
    CompraNotaDevolucion = 13
    CompraNotaCredito = 14
    CompraRemito = 15
    CompraCarta = 16
    CompraCarpeta = 17
    CompraRecibo = 18
    CompraReciboDePago = 19
    CompraSalidaCaja = 30       'Pedidos el 11/12 por carlos y juliana
    CompraEntradaCaja = 31
    
    'Otros
    Traslados = 20
    Envios = 21
    CambioEstadoMercaderia = 22
    
    IngresoMercaderiaEspecial = 24
    ArregloStock = 25
    Servicio = 26
    ServicioCambioEstado = 27
    Devolucion = 28
    RemitoEntrega = 47
    RemitoRetiro = 48
End Enum

Public Enum EstadoEnvio
    AImprimir = 0
    AConfirmar = 1
    Rebotado = 2
    Impreso = 3
    Entregado = 4
    Anulado = 5
    OnLineAConfirmar = 6
End Enum

'Impresora
Public paPrintRemRepB As Integer
Public paPrintRemRepD As String

Public paPrintConfB As Integer
Public paPrintConfD As String
Public paPrintConfXDef As Boolean
Public paPrintConfPaperSize As Integer
Public paIContadoB As Integer
Public paIContadoN As String
Public paPrintCtdoDef As Boolean
Public paPrintCtdoPaperSize As Integer

Public paBandNCtdo As Integer
Public paDevNCtdo As String
Public paXDefNCtdo As Boolean


Public paTComEnvConf As Long            'Comentario Envío a Confirmar.
Public paUIDEnvConf As String
Public paPrimeraHoraEnvio As Long, paUltimaHoraEnvio As Long


Public prmURLFirmaEFactura As String
Public prmEFacturaProductivo As String
Public prmImporteConInfoCliente As Currency
Public TasaBasica As Currency, TasaMinima As Currency
Public EmpresaEmisora As clsClienteCFE
Public oArtDifEnvio As clsProducto
Public oArtPisoAgencia As clsProducto

Public paMCVtaTelefonica As Long
Public paCofis As Currency
Public paArticuloDiferenciaEnvio As Long
Public paArticuloPisoAgencia As Long
Public paTipoArticuloServicio As Integer

Public paHoraTarde As String


Public paBD As String
Public gPathListados As String
'Stock
Public paEstadoArticuloEntrega As Integer

Public miConexion As New clsConexion

Public objGral As New clsorCGSA

Public Sub Main()
On Error GoTo errMain
Dim sAux As String
    
    Screen.MousePointer = 11
    Dim mAppVer As String
    mAppVer = App.Title
    If Not miConexion.AccesoAlMenu(mAppVer) Then
        MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        GoTo evEnd
    Else
    
        If mAppVer <> "" And mAppVer <> App.Title And Not (App.Major & "." & App.Minor & "." & Format(App.Revision, "00")) >= mAppVer Then
            MsgBox "La versión del programa no es la última disponible." & vbCr & _
                        "Ud. debe actualizar el software.", vbExclamation, "Actualizar a Versión " & mAppVer
            End
        End If
    
        paCodigoDeUsuario = miConexion.UsuarioLogueado(True)
        If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then GoTo evEnd
        If Not CargoDatosSucursal(miConexion.NombreTerminal, paDContado, sNameNCtdo:=paDNDevolucion, sNameNRemito:=paDRemito) Then
            GoTo evEnd
        Else
            If paDContado = "" Or paDRemito = "" Or paDNDevolucion = "" Then
                MsgBox "Está aplicación requiere del nombre de los documentos contados y remitos, comuniquese con el administrador.", vbExclamation, "ATENCIÓN"
            End If
        End If
        If Not f_GetParameters Then MsgBox "Imposible continuar sin parámetros de stock.", vbCritical, "Atención": GoTo evEnd
        prj_GetPrinter False
        
        ChDir App.Path
        ChDir ("..")
        ChDir (CurDir & "\REPORTES\")
        gPathListados = CurDir & "\"
        
        paBD = miConexion.RetornoPropiedad(bDB:=True)
        
'        On Error Resume Next
    'ABRO ENGINE DE IMPRESION
'        crAbroEngine
        
        CargoValoresIVA
        Set EmpresaEmisora = New clsClienteCFE
        EmpresaEmisora.CargoClienteCarlosGutierrez paCodigoDeSucursal  '.CargoInformacionCliente cBase, 1, False
    
        frmDistribuirEnvio.Show
        Set miConexion = Nothing
    End If
    
    Screen.MousePointer = 0
    Exit Sub

evEnd:
    Screen.MousePointer = 0
    Set objGral = Nothing
    Set miConexion = Nothing
    End
    Exit Sub
    
errMain:
    objGral.OcurrioError "Error al cargar la aplicación.", Err.Description
    Set objGral = Nothing
    Screen.MousePointer = 0
    End
End Sub

Public Sub prj_GetPrinter(ByVal bShowP As Boolean)
On Error GoTo errImp
    paPrintConfD = ""
    paPrintConfB = -1
    paIContadoN = ""
    Dim objP As New clslPrintConfig
    With objP
        If bShowP Then
            If Not .ShowPrinterSetup("21,1,3,6,47", paCodigoDeTerminal) Then
                GoTo errImp
            End If
        End If
        If .LoadPrinterConfig("21,1,3,6,47", paCodigoDeTerminal) Then
            .GetPrinterDoc TipoDocumento.Envios, paPrintConfD, paPrintConfB, paPrintConfXDef, paPrintConfPaperSize
            .GetPrinterDoc TipoDocumento.Contado, paIContadoN, paIContadoB, paPrintCtdoDef, paPrintCtdoPaperSize
            .GetPrinterDoc NotaDevolucion, paDevNCtdo, paBandNCtdo, paXDefNCtdo
            .GetPrinterDoc RemitoEntrega, paPrintRemRepD, paPrintRemRepB, paXDefNCtdo
            '.GetPrinterDoc 6, "", BandejaRecibos, True, 1
        End If
    End With
    If paIContadoN = "" Or paPrintConfD = "" Or paDevNCtdo = "" Then MsgBox "Por favor verifique la configuración de impresión.", vbInformation, "Atención"
    
errImp:
    Set objP = Nothing
    Screen.MousePointer = 0
End Sub

Private Function f_GetParameters() As Boolean
On Error GoTo errGP
    f_GetParameters = False
    
    Cons = "Select * From Parametro " & _
                " Where ParNombre IN('eFacturaActiva', 'URLFirmaEFactura', 'efactImporteDatosCliente', 'EstadoArticuloEntrega', 'MCVtaTelefonica', 'Cofis'," & _
                            " 'menusuarioenvioconfirma', 'tcomentarioenvaconf', 'ArticuloPisoAgencia'," & _
                            "'ArticuloDiferenciaEnvio', 'TipoArticuloServicio', 'PrimeraHoraEnvio', 'UltimaHoraEnvio', 'ValorUIUltimoMes', 'RepartoMsgWhatsApp')"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "repartomsgwhatsapp": paSendWApp = RsAux("ParValor")
            Case "valoruiultimomes":
                paValorUIUltMes = RsAux("ParValor")
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case "mcvtatelefonica": paMCVtaTelefonica = RsAux!ParValor
            Case "cofis": paCofis = RsAux!ParValor
            Case "articulodiferenciaenvio": paArticuloDiferenciaEnvio = RsAux!ParValor
            Case "articulopisoagencia": paArticuloPisoAgencia = RsAux!ParValor
            Case "tipoarticuloservicio": paTipoArticuloServicio = RsAux!ParValor
            Case "tcomentarioenvaconf": If Not IsNull(RsAux!ParValor) Then paTComEnvConf = RsAux!ParValor
            Case "menusuarioenvioconfirma": If Not IsNull(RsAux!ParTexto) Then paUIDEnvConf = Trim(RsAux!ParTexto)
            Case LCase("PrimeraHoraEnvio"): paPrimeraHoraEnvio = RsAux!ParValor
            Case LCase("UltimaHoraEnvio"): paUltimaHoraEnvio = RsAux!ParValor
            
            Case LCase("efactImporteDatosCliente"): prmImporteConInfoCliente = RsAux("ParValor")
            Case LCase("URLFirmaEFactura"): prmURLFirmaEFactura = Trim(RsAux("ParTexto"))
            Case LCase("eFacturaActiva"): prmEFacturaProductivo = RsAux("ParValor")

        End Select
        RsAux.MoveNext
    Loop
   
    RsAux.Close
    
    
'Cambie esto x leer valor de la registry.

    paHoraTarde = GetSetting(App.Title, "Parametros", "ComienzoHoraTarde", "1200")

'    Cons = "Select IsNull(Clase, 1400) From CodigoTexto " & _
                "Where codigo In(Select ParValor From Parametro Where ParNombre = 'TardeEnvio')"
 '   Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
  '  If Not RsAux.EOF Then paHoraTarde = RsAux(0) Else paHoraTarde = "1400"
   ' RsAux.Close
    
    f_GetParameters = (paEstadoArticuloEntrega > 0)
    Exit Function
errGP:
    objGral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Function
    
Public Function RetornoCliente(ByVal idCliente As Long, ByVal idVtaTelef As Long) As clsClienteCFE
Dim sQy As String
Dim rsC As rdoResultset
Dim nomCliente As String

    Set RetornoCliente = New clsClienteCFE
    RetornoCliente.CargoInformacionCliente cBase, idCliente, True

    If idVtaTelef > 0 Then
        sQy = "SELECT IsNull(VTeDireccionFactura, 0) Dir, IsNull(VTeNombreFactura, '') Nombre, " & _
            " IsNull(CalNombre, '') CalNombre, IsNull(DepNombre, '') DepNombre, IsNull(LocNombre, '') LocNombre, " & _
            " IsNull(DirPuerta, 0) DirPuerta " & _
            " FROM VentaTelefonica " & _
            "LEFT OUTER JOIN Direccion ON VTeDireccionFactura = DirCodigo " & _
            "LEFT OUTER JOIN Calle ON DirCalle = CalCodigo " & _
            "LEFT OUTER JOIN Localidad ON CalLocalidad = LocCodigo " & _
            "LEFT OUTER JOIN Departamento ON LocDepartamento = DepCodigo " & _
            "WHERE VTeCodigo = " & idVtaTelef
        Set rsC = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
        If Not rsC.EOF Then
            If Trim(rsC("Nombre")) <> "" Then RetornoCliente.NombreCliente = Trim(rsC("Nombre"))
            If rsC("Dir") > 0 And Trim(rsC("CalNombre")) <> "" Then
                RetornoCliente.Direccion.Domicilio = Trim(rsC("CalNombre")) & " " & rsC("DirPuerta")
                RetornoCliente.Direccion.Departamento = Trim(rsC("DepNombre"))
                RetornoCliente.Direccion.Localidad = Trim(rsC("LocNombre"))
            End If
        End If
        rsC.Close
    End If

    If RetornoCliente.Direccion.Domicilio = "" Then
        'Busco la dirección del envío.
        sQy = "SELECT Top 1 " & _
            " IsNull(CalNombre, '') CalNombre, IsNull(DepNombre, '') DepNombre, IsNull(LocNombre, '') LocNombre, " & _
            " IsNull(DirPuerta, 0) DirPuerta " & _
            " FROM Envio " & _
            "LEFT OUTER JOIN Direccion ON EnvDireccion = DirCodigo " & _
            "LEFT OUTER JOIN Calle ON DirCalle = CalCodigo " & _
            "LEFT OUTER JOIN Localidad ON CalLocalidad = LocCodigo " & _
            "LEFT OUTER JOIN Departamento ON LocDepartamento = DepCodigo " & _
            "WHERE EnvCliente = " & idCliente
        Set rsC = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
        If Not rsC.EOF Then
            If Trim(rsC("CalNombre")) <> "" Then
                RetornoCliente.Direccion.Domicilio = Trim(rsC("CalNombre")) & " " & rsC("DirPuerta")
                RetornoCliente.Direccion.Departamento = Trim(rsC("DepNombre"))
                RetornoCliente.Direccion.Localidad = Trim(rsC("LocNombre"))
            End If
        End If
        rsC.Close
    End If
End Function

    
Public Function CargoArticulosPrms(ByVal idArt As Long) As clsProducto
    Dim oArt As New clsProducto
    Set CargoArticulosPrms = oArt
    Dim sQy As String
    Dim rsA As rdoResultset
    sQy = "SELECT ArtId, ArtNombre, ArtTipo, IvaPorcentaje  FROM Articulo INNER JOIN ArticuloFacturacion ON ArtID = AFaArticulo INNER JOIN TipoIVA ON IvaCodigo = AFaIva WHERE ArtId = " & idArt
    Set rsA = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    With oArt
        .ID = rsA("ArtID")
        .Nombre = Trim(rsA("ArtNombre"))
        .TipoIVA.Porcentaje = rsA("IvaPorcentaje")
        .TipoArticulo = rsA("ArtTipo")
    End With
    rsA.Close
    
End Function
    
Public Sub RunApp(Path As String, Optional Valor As String = "", Optional Modal As Boolean = False)

    On Error GoTo errApp
    Screen.MousePointer = 11
    Dim plngRet As Long
    
    If Valor = "" Then
        Dim aPos As Integer
        aPos = InStr(Path, ".exe")
        If aPos <> 0 Then
            Valor = Mid(Path, aPos + Len(".exe") + 1)
            Path = Mid(Path, 1, aPos - 1)
        End If
    End If
    plngRet = ShellExecute(0, "open", Path, Valor, 0, 1)
    If plngRet = 0 And Err.Description <> "" Then GoTo errApp
    
    
    Screen.MousePointer = 0
    Exit Sub
    
errApp:
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al ejecutar la aplicación " & Path & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbCritical, "Error de Aplicación"
End Sub
Public Sub CargoCombo(Consulta As String, Combo As Control, Optional Seleccionado As String = "")
Dim RsAuxiliar As rdoResultset
Dim iSel As Integer: iSel = -1     'Guardo el indice del seleccionado
    
On Error GoTo ErrCC
    
    Screen.MousePointer = 11
    Combo.Clear
    Set RsAuxiliar = cBase.OpenResultset(Consulta, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAuxiliar.EOF
        Combo.AddItem Trim(RsAuxiliar(1))
        Combo.ItemData(Combo.NewIndex) = RsAuxiliar(0)
        
        If Trim(RsAuxiliar(1)) = Trim(Seleccionado) Then iSel = Combo.ListCount
        RsAuxiliar.MoveNext
    Loop
    RsAuxiliar.Close
    
    If iSel = -1 Then Combo.ListIndex = iSel Else Combo.ListIndex = iSel - 1
    Screen.MousePointer = 0
    Exit Sub
    
ErrCC:
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al cargar el combo: " & Trim(Combo.Name) & "." & vbCrLf & Err.Description, vbCritical, "ERROR"
End Sub

Public Function NumeroDocumento(Documento As String)

    Dim Auxiliar As String    'Auxiliar para retornar el NRO DOC (Serie + Nro)
    Dim RsDoc As rdoResultset
    
    Cons = "Select * from Contador Where ConDocumento = '" & Trim(Documento) & "'"
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurLock)
    
    If RsDoc.EOF Then
        RsDoc.AddNew
        RsDoc!ConDocumento = Trim(Documento)
        RsDoc!ConValor = 1
        RsDoc!ConSerie = "A"
        Auxiliar = RsDoc!ConSerie & RsDoc!ConValor
        RsDoc.Update
    Else
        If RsDoc!ConValor = 999999 Then
            RsDoc.Edit
            RsDoc!ConValor = 1
            RsDoc!ConSerie = Chr(Asc(RsDoc!ConSerie) + 1)
            Auxiliar = RsDoc!ConSerie & RsDoc!ConValor
            RsDoc.Update
        Else
            Auxiliar = Trim(RsDoc!ConSerie) & RsDoc!ConValor + 1
            RsDoc.Edit
            RsDoc!ConValor = RsDoc!ConValor + 1
            RsDoc.Update
        End If
    End If
    RsDoc.Close
    
    NumeroDocumento = Auxiliar
    
End Function

Public Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer

    For Each X In Printers
        If UCase(Trim(X.DeviceName)) = UCase(Trim(DeviceName)) Then
            Set Printer = X
            Exit For
        End If
    Next
    
End Sub

Public Function CodigoDeBarras(TipoDoc As Integer, CodigoDoc As Long)
    If Len(CodigoDoc) < 6 Then
        CodigoDeBarras = TipoDoc & "D" & Format(CodigoDoc, "000000")
    Else
        CodigoDeBarras = TipoDoc & "D" & CodigoDoc
    End If
    CodigoDeBarras = "*" & CodigoDeBarras & "*"
End Function

Public Function EmitirCFE(ByVal Documento As clsDocumentoCGSA, ByVal CAE As clsCAEDocumento) As String
On Error GoTo errEC
    With New clsCGSAEFactura
        .URLAFirmar = prmURLFirmaEFactura
        .TasaBasica = TasaBasica
        .TasaMinima = TasaMinima
        .ImporteConInfoDeCliente = prmImporteConInfoCliente
        Set .Connect = cBase
        If Not .GenerarEComprobante(CAE, Documento, EmpresaEmisora, paCodigoDGI) Then
            EmitirCFE = .XMLRespuesta
        End If
    End With
    Exit Function
errEC:
    EmitirCFE = "Error en firma: " & Err.Description
End Function

Private Sub CargoValoresIVA()
Dim RsIva As rdoResultset
Dim sQy As String
    sQy = "SELECT IvaCodigo, IvaPorcentaje FROM TipoIva WHERE IvaCodigo IN (1,2)"
    Set RsIva = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsIva.EOF
        Select Case RsIva("IvaCodigo")
            Case 1: TasaBasica = RsIva("IvaPorcentaje")
            Case 2: TasaMinima = RsIva("IvaPorcentaje")
        End Select
        RsIva.MoveNext
    Loop
    RsIva.Close
End Sub

Public Sub BuscoCodigoEnCombo(cCombo As Control, lngCodigo As Long)
Dim I As Integer
    
    If cCombo.ListCount > 0 Then
        For I = 0 To cCombo.ListCount - 1
            If cCombo.ItemData(I) = lngCodigo Then
                cCombo.ListIndex = I
                Exit Sub
            End If
        Next I
        cCombo.ListIndex = -1
    Else
        cCombo.ListIndex = -1
    End If

End Sub

