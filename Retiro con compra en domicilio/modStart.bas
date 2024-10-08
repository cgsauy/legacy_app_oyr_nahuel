Attribute VB_Name = "modStart"
Option Explicit

Public Cons As String
Public cBase As rdoConnection
Public rsAux As rdoResultset

Public paCodigoDeSucursal As Integer
Public paCodigoDeTerminal As Integer
Public paNombreSucursal As String
Public paCodigoDGI As Integer

Public paLocalCompa�ia As Integer
Public paPathReportes As String
Public paTipoComentario As Integer

'Impresora
Public paPrintConfB As Integer
Public paPrintConfD As String
Public paPrintConfXDef As Boolean
Public paPrintConfPaperSize As Integer

Public paArtsNoNotaEsp As String
Public paClienteEmpresa As Long
Public paClienteNoVtoCta As String
Public paDRemito As String

Public paArticuloPisoAgencia As Long, paArticuloDiferenciaEnvio As Long, paTipoArticuloServicio As Long
Public paEstadoARecuperar As Integer, paEstadoArticuloEntrega As Integer, paEstadoRoto As Integer
Public paArticuloARoto As String

Public gFechaServidor As Date

Public clsGeneral As New clsorCGSA
Public ParametrosSist As New clsParametros

Public Sub Main()
Dim aValor As Integer
Dim miConexion As New clsConexion
Dim sError As String
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    sError = "Conexi�n"
    If miConexion.AccesoAlMenu(App.Title) Then
        
        sError = "Conectando"
        Dim objFnc As New clsFncGlobales
        If Not objFnc.GetBDConnect(cBase, "Comercio") Then GoTo evFin
        Set objFnc = Nothing
        
        sError = "Sucursal"
        'C�dgio de Sucursal
        If Not CargoDatosSucursal(miConexion.NombreTerminal) Then GoTo evFin

        sError = "Parametros"
        If Not CargoParametros Then GoTo evFin
        
        Dim colPrms As New Collection
        colPrms.Add NombreDeParametros.efactImporteDatosCliente
        colPrms.Add NombreDeParametros.URLFirmaEFactura
        ParametrosSist.CargoParametrosComercio colPrms
        
        If ParametrosSist.ObtenerValorParametro(URLFirmaEFactura).Texto = "" Then
            MsgBox "Es necesario tener la URL de firma de efactura, IMPOSIBLE SEGUIR", vbExclamation, "ATENCI�N"
            End
            Exit Sub
        End If
        
        Set miConexion = Nothing
        sError = "Impresi�n"
        prj_GetPrinter False
        
        sError = "Abrir Form"
        frmRetiroConCompra.Show
        sError = ""
    Else
        sError = "Sin Permisos"
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicaci�n.", vbExclamation, "ATENCI�N"
        GoTo evFin
    End If
    Exit Sub
    
errMain:
    On Error Resume Next
    MsgBox "Ocurri� un error al inicializar la aplicaci�n " & App.Title & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description) & vbCrLf & " Paso: " & sError
    Screen.MousePointer = 0
    
evFin:
    Screen.MousePointer = 0
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End

    
End Sub

Private Function CargoParametros() As Boolean
'Controlo aquellos que son vitales si no los cargue finalizo la app.
On Error GoTo errCP
    
    'Parametros a cero--------------------------
    paEstadoARecuperar = 0: paEstadoArticuloEntrega = 0
    paArticuloPisoAgencia = 0: paArticuloDiferenciaEnvio = 0: paTipoArticuloServicio = 0

    Cons = "Select * from Parametro Where ParNombre IN('estadoarticuloentrega', 'estadoarecuperar', 'tipoarticuloservicio', 'EstadoRoto', " & _
                                                            "'articulopisoagencia', 'articulodiferenciaenvio', 'clienteempresa', 'ArtsNEspInhabilitado', " & _
                                                            "'ClienteNoCuotaVencida', 'LocalCompa�ia', 'PathReportes', 'tcomentariocambioprod', 'ArticulosIngresanRoto')"
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case LCase(Trim(rsAux!ParNombre))
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = rsAux!ParValor
            Case "estadoarecuperar": paEstadoARecuperar = rsAux!ParValor
            Case "tipoarticuloservicio": paTipoArticuloServicio = rsAux!ParValor
            Case "articulopisoagencia": paArticuloPisoAgencia = rsAux!ParValor
            Case "articulodiferenciaenvio": paArticuloDiferenciaEnvio = rsAux!ParValor
            Case "clienteempresa": paClienteEmpresa = rsAux!ParValor
            Case LCase("ArtsNEspInhabilitado"): paArtsNoNotaEsp = Trim(rsAux!ParTexto)
            Case LCase("ClienteNoCuotaVencida"): paClienteNoVtoCta = Trim(rsAux!ParTexto)
            Case "localcompa�ia": paLocalCompa�ia = rsAux!ParValor
            Case "pathreportes": paPathReportes = Trim(rsAux("ParTexto"))
            Case "tcomentariocambioprod": paTipoComentario = rsAux!ParValor
            Case LCase("ArticulosIngresanRoto"): paArticuloARoto = Trim(rsAux("ParTexto"))
            Case LCase("EstadoRoto"): paEstadoRoto = rsAux("ParValor")
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    paClienteNoVtoCta = Replace(paClienteNoVtoCta, " ", "")
    paArtsNoNotaEsp = Replace(paArtsNoNotaEsp, " ", "")
    
    
    CargoParametros = (paEstadoArticuloEntrega > 0 And paEstadoARecuperar > 0)
    If Not CargoParametros Then MsgBox "Los par�metros de Estado de stock no fueron leidos, no podr� continuar.", vbCritical, "Manejo de Stock"
    Exit Function
errCP:
     clsGeneral.OcurrioError "Error al leer los par�metros.", Err.Description
     CargoParametros = False
End Function

Public Sub prj_GetPrinter(ByVal bShowP As Boolean)
On Error GoTo errImp
    
    paPrintConfD = ""
    paPrintConfB = -1
    Dim objP As New clslPrintConfig
    With objP
        If bShowP Then
            If Not .ShowPrinterSetup("21", paCodigoDeTerminal) Then
                GoTo errImp
            End If
        End If
        If .LoadPrinterConfig("21", paCodigoDeTerminal) Then
            .GetPrinterDoc 21, paPrintConfD, paPrintConfB, paPrintConfXDef, paPrintConfPaperSize
        End If
    End With
    If paPrintConfD = "" Then MsgBox "Por favor verifique la configuraci�n de impresi�n.", vbInformation, "Atenci�n"
    
errImp:
    Set objP = Nothing
    Screen.MousePointer = 0
End Sub


Public Function CargoDatosSucursal(ByVal sNombreTerminal As String) As Boolean
'................................................................................................................................................................
'Dado el nombre de la terminal
'   Cargo el c�digo de la misma, el c�digo de la sucursal y el nombre de los documentos.
'................................................................................................................................................................
On Error GoTo errCDS

    CargoDatosSucursal = False
    
    paCodigoDeSucursal = 0: paCodigoDeTerminal = 0
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & sNombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        paCodigoDeSucursal = rsAux!TerSucursal
        paCodigoDeTerminal = rsAux!TerCodigo
        paNombreSucursal = Trim(rsAux!SucAbreviacion)
        paCodigoDGI = rsAux("SucCodDGI")
    End If
    rsAux.Close
  
    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(sNombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
                & "La ejecuci�n ser� cancelada.", vbCritical, "ATENCI�N"
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------------------------------
    CargoDatosSucursal = (paCodigoDeSucursal > 0)
    Exit Function

errCDS:
    MsgBox "Error al leer la informaci�n de la sucursal." & vbCr & vbCr & "Error: " & Err.Description, vbCritical, "Datos de Sucursal"
End Function

Public Sub FechaDelServidor()

    Dim RsF As rdoResultset
    On Error GoTo errFecha
    Cons = "Select GetDate()"
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    gFechaServidor = RsF(0)
    RsF.Close
    
    On Error Resume Next
    Time = gFechaServidor
    Date = gFechaServidor
    Exit Sub

errFecha:
    gFechaServidor = Now
End Sub

