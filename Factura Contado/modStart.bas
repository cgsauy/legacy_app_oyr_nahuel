Attribute VB_Name = "modStart"
'Cambios
'24-10-2006
'En buscocomentariosalertas agregue función para validar si se le puede vender o no al cliente.
Option Explicit



'Definición del entorno RDO
Public cBase As rdoConnection       'Conexion a la Base de Datos
Public eBase As rdoEnvironment     'Definicion de entorno
Public RsAux As rdoResultset         'Resultset Auxiliar

'String.
Public Cons As String
Public paCodigoDeUsuario As Long
Public paCodigoDeSucursal As Long
'Public paCodigoDGI As Long
Public paCodigoDeTerminal As Long


Public prmURLFirmaEFactura As String
Public prmEFacturaProductivo As String
Public prmImporteConInfoCliente As Currency
Public paNoFletesVta As String

Public paNombreSucursal As String
'IMPRESION
Public paIContadoB As Integer
Public paIContadoN As String
Public paPrintEsXDef As Boolean

'Por aportes agregue recibo y movs de caja.
Public paIReciboB As Integer
Public paIReciboN As String

Public paIRemitoB As Integer
Public paIRemitoN As String

Public paLastUpdate As String
Public paOptPrintSel As String      'El nombre de la opción seleccionada
Public paOptPrintList As String      'Los nombres de las opciones ingresadas están separadas x |
'Public paImpCedula As Currency
'...........................................................

Public paToleranciaAportes As Integer

Public Type tRenglonFact
    IDArticulo As Long
    CodArticulo As Long
    NombreArticulo As String
    IDCombo As Long
'    ArtCombo As Long
    Tipo As Long
    Precio As Currency
    PrecioOriginal As Currency
    PrecioBonificacion As Currency
    EsInhabilitado As Boolean
    DisponibleDesde As Date
    VentaXMayor As Byte
    CantidadAlXMayor As Integer
'    SucesoVtaXMayor As Boolean
End Type

Public Enum TipoArticulo
    Articulo = 1
    Servicio = 2
    Presupuesto = 3
    PagoFlete = 4
    Bonificacion = 5
    Especifico = 6
End Enum

Public Const FormatoCedula = "_.___.___-_"
Public gPathListados As String, paBD As String

Public paCategoriaDistribuidor As String

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public paLocalPuerto As Long, paLocalZF As Long
Public paTipoComCheque As Long
Public paArtsSinCofis As String

'Public tTiposArtsServicio As String     'Trama con todos los tipos que pertenecen a Servicio

Public Sub Main()

    On Error GoTo errMain
    Screen.MousePointer = 11
        
    Dim mAppVer As String
    mAppVer = App.Title
    If miConexion.AccesoAlMenu(mAppVer) Then
        
        If mAppVer <> "" And mAppVer <> App.Title And Not (App.Major & "." & App.Minor & "." & Format(App.Revision, "00")) >= mAppVer Then
            MsgBox "La versión del programa no es la última disponible." & vbCr & _
                        "Ud. debe actualizar el software.", vbExclamation, "Actualizar a Versión " & mAppVer
            End
        End If
        If Not ObtenerConexionBD(cBase, logComercio) Then
            End
            Screen.MousePointer = 0
            Exit Sub
        End If
        If Not loc_GetInfoSucursal Then
            Screen.MousePointer = 0
            End
            Exit Sub
        End If
        paBD = miConexion.RetornoPropiedad(bDB:=True)
                 
        prj_LoadConfigPrint
         
        CargoParametrosComercio
        loc_CargoParametroContado
        SeteoPathReportes
        'CargoTiposDeArticulosServicios
        If Trim(Command()) <> "" Then
            If IsNumeric(Command()) Then
                FacContado.prmIDCliente = CLng(Command())
            End If
        End If
        FacContado.Show vbModeless
    Else
        MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description), vbCritical, "ATENCIÓN"
    
    End
End Sub

'Private Sub CargoTiposDeArticulosServicios()
'Dim sQy As String
'    tTiposArtsServicio = ""
'    sQy = "SELECT TipID FROM dbo.InTipos(" & paTipoArticuloServicio & ")"
'    If ObtenerResultSet(cBase, RsAux, sQy, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
'    Do While Not RsAux.EOF
'        tTiposArtsServicio = tTiposArtsServicio & IIf(tTiposArtsServicio <> "", ",", "") & Trim(RsAux("TipID"))
'        RsAux.MoveNext
'    Loop
'    RsAux.Close
'End Sub

Private Sub SeteoPathReportes()
On Error GoTo errSPR
    ChDir App.Path
    ChDir ("..")
    ChDir (CurDir & "\REPORTES\")
    gPathListados = CurDir & "\"
Exit Sub
errSPR:
    clsGeneral.OcurrioError "Error al buscar el camino a los archivos de impresión.", Err.Description, "Error"
End Sub

Private Sub loc_CargoParametroContado()
    Cons = "Select * from Parametro Where ParNombre IN ('eFacturaActiva', 'URLFirmaEFactura', 'catcliDistribuidor', 'TComentarioOperaCH', 'ImporteConCedula', 'ArtsSinCofis', 'AportesACuentaMesesDisponibles', 'efactImporteDatosCliente', 'fletesnoenviarvtatelef')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux("ParNombre")))
            Case LCase("efactImporteDatosCliente"): prmImporteConInfoCliente = RsAux("ParValor")
            Case LCase("URLFirmaEFactura"): prmURLFirmaEFactura = Trim(RsAux("ParTexto"))
            Case LCase("eFacturaActiva"): prmEFacturaProductivo = RsAux("ParValor")
'            Case "importeconcedula"
'                 paImpCedula = RsAux("ParValor")
            Case LCase("TComentarioOperaCH")
                paTipoComCheque = RsAux!ParValor
            Case LCase("ArtsSinCofis")
                If Not IsNull(RsAux("ParTexto")) Then paArtsSinCofis = Trim(RsAux("ParTexto"))
            Case LCase("catcliDistribuidor")
                If Not IsNull(RsAux("ParTexto")) Then paCategoriaDistribuidor = Trim(RsAux("ParTexto"))
            Case LCase("AportesACuentaMesesDisponibles")
                If Not IsNull(RsAux("ParValor")) Then paToleranciaAportes = RsAux("ParValor")
            Case "fletesnoenviarvtatelef": paNoFletesVta = RsAux("ParTexto")
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If paCategoriaDistribuidor <> "" Then paCategoriaDistribuidor = "," & Replace(paCategoriaDistribuidor, " ", "") & ","
    If paArtsSinCofis <> "" Then paArtsSinCofis = "," & paArtsSinCofis & ","

End Sub

Private Function loc_GetInfoSucursal() As Boolean

    loc_GetInfoSucursal = False
    
    paCodigoDeSucursal = 0
    paCodigoDeTerminal = 0
    paDContado = ""
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select TerSucursal, TerCodigo, SucDisponibilidad, SucDContado, SucAbreviacion, SucDRecibo, SucCodDGI From Terminal, Sucursal" _
            & " Where TerNombre = '" & miConexion.NombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDGI = RsAux("SucCodDGI")
        paCodigoDeTerminal = RsAux!TerCodigo
        If Not IsNull(RsAux!SucDisponibilidad) Then paDisponibilidad = RsAux!SucDisponibilidad Else paDisponibilidad = 0
        paNombreSucursal = Trim(RsAux!SucAbreviacion)
        
        'El documento que necesito solo es el contado.
        If Not IsNull(RsAux!SucDContado) Then paDContado = Trim(RsAux!SucDContado)
        If Not IsNull(RsAux!SucDRecibo) Then paDRecibo = Trim(RsAux!SucDRecibo)
    End If
    RsAux.Close
  
    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(miConexion.NombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
                & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
        Exit Function
    ElseIf paDisponibilidad = 0 Or paDContado = "" Then
        MsgBox "El pc pertenece a una sucursal que no está habilitada para facturar, la ejecución será cancelada.", vbExclamation, "Sucursal"
        Exit Function
    Else
        loc_GetInfoSucursal = True
    End If
    '-------------------------------------------------------------------------------------------------------------------------
    
End Function

Public Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer

    For Each X In Printers
        If Trim(X.DeviceName) = Trim(DeviceName) Then
            Set Printer = X
            Exit For
        End If
    Next
    
End Sub

Public Sub prj_LoadConfigPrint(Optional bShowFrm As Boolean = False)
On Error GoTo errLCP
Dim objPrint As New clsCnfgPrintDocument
Dim sPrint As String, sPrintRecibo As String, sPrintRemito As String
Dim vPrint() As String
    With objPrint
         Set .Connect = cBase
        .Terminal = paCodigoDeTerminal
        If bShowFrm Then
            .CnfgTipoDocumento = "1"
            .ShowConfig
        End If
        If paLastUpdate <> .FechaUltimoCambio Or paLastUpdate = "" Then
            sPrint = .getDocumentoImpresora(Contado)
            paOptPrintSel = .GetOpcionActual
            paOptPrintList = .GetOpcionesPrinter
            paLastUpdate = .FechaUltimoCambio
        End If
        sPrintRecibo = .getDocumentoImpresora(ReciboDePago)
        sPrintRemito = .getDocumentoImpresora(remito)
    End With
    Set objPrint = Nothing
    
    If sPrint <> "" Then
        vPrint = Split(sPrint, "|")
        paIContadoN = vPrint(0)
        paIContadoB = vPrint(1)
        paPrintEsXDef = (Val(vPrint(3)) = 1)
    End If
    
    If sPrintRecibo <> "" Then
        vPrint = Split(sPrintRecibo, "|")
        paIReciboN = vPrint(0)
        paIReciboB = vPrint(1)
    End If
    
    If sPrintRemito <> "" Then
        vPrint = Split(sPrintRemito, "|")
        paIRemitoN = vPrint(0)
        paIRemitoB = vPrint(1)
    End If
    Exit Sub
errLCP:
    MsgBox "Error al leer los datos de configuración: " & Err.Description, vbExclamation, "ATENCIÓN"
End Sub

Public Function ChangeCnfgPrint() As Boolean
    Dim objPrint As New clsCnfgPrintDocument
    ChangeCnfgPrint = (paLastUpdate <> objPrint.FechaUltimoCambio)
    Set objPrint = Nothing
End Function

Public Function SaldoCuentaPersonal(ByVal Cliente As Long, ByVal advertir As Boolean) As Currency
On Error GoTo errSCP
    
    SaldoCuentaPersonal = 0
'    Dim oACta As New clsAporteACuenta
'    SaldoCuentaPersonal = oACta.SaldoCuentaPersonal(cBase, 1, Cliente, False)
'    Set oACta = Nothing
'
'    If advertir And SaldoCuentaPersonal > 0 Then
'        MsgBox "ATENCIÓN!!!" & vbCrLf & vbCrLf & "El cliente posee aportes a cuenta que puede utilizar para pagar esta factura, por favor comuníqueselo.", vbInformation, "IMPORTANTE"
'    End If
    
errSCP:
End Function

Public Function BuscoUsuarioDigito(Digito As Long, Optional Codigo As Boolean = False, Optional Identificacion As Boolean = False, Optional Iniciales As Boolean = False) As Variant
Dim RsUsr As rdoResultset
Dim aRetorno As Variant
On Error GoTo ErrBUD

    Cons = "Select * from Usuario Where UsuDigito = " & Digito
    'Set RsUsr = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsUsr, Cons, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Codigo Then aRetorno = RsUsr!UsuCodigo
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    BuscoUsuarioDigito = aRetorno
    Exit Function
    
ErrBUD:
    MsgBox "Error inesperado al buscar el usuario.", vbCritical, "ATENCIÓN"
End Function
