Attribute VB_Name = "modStart"
Option Explicit

Public Cons As String
Public cBase As rdoConnection
Public rsAux As rdoResultset

Public paCodigoDeSucursal As Integer
Public paCodigoDeTerminal As Integer
Public paNombreSucursal As String

Public paLocalCompañia As Integer
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

Public colLocalesRepara As Collection
Public clsGeneral As New clsorCGSA

Public Sub Main()
Dim aValor As Integer
Dim miConexion As New clsConexion
Dim sError As String
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    sError = "Conexión"
    If miConexion.AccesoAlMenu(App.Title) Then
        
        sError = "Conectando"
        Dim objFnc As New clsFncGlobales
        If Not objFnc.GetBDConnect(cBase, "Comercio") Then GoTo evFin
        Set objFnc = Nothing
        
        sError = "Sucursal"
        'Códgio de Sucursal
        If Not CargoDatosSucursal(miConexion.NombreTerminal, sNameNRemito:=paDRemito) Then GoTo evFin

        sError = "Parametros"
        If Not CargoParametros Then GoTo evFin
        
        Set miConexion = Nothing
        sError = "Impresión"
        prj_GetPrinter False
        
        sError = "Abrir Form"
        frmEntMercaderia.Show
        sError = ""
    Else
        sError = "Sin Permisos"
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        GoTo evFin
    End If
    Exit Sub
    
errMain:
    On Error Resume Next
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description) & vbCrLf & " Paso: " & sError
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
                                                            "'ClienteNoCuotaVencida', 'LocalCompañia', 'PathReportes', 'tcomentariocambioprod', 'ArticulosIngresanRoto')"
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
            Case "localcompañia": paLocalCompañia = rsAux!ParValor
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
    If Not CargoParametros Then MsgBox "Los parámetros de Estado de stock no fueron leidos, no podrá continuar.", vbCritical, "Manejo de Stock"
    Exit Function
errCP:
     clsGeneral.OcurrioError "Error al leer los parámetros.", Err.Description
     CargoParametros = False
End Function

Public Sub prj_GetPrinter(ByVal bShowP As Boolean)
On Error GoTo errImp
    
    paPrintConfD = ""
    paPrintConfB = -1
    Dim objP As New clslPrintConfig
    With objP
        If bShowP Then
            If Not .ShowPrinterSetup("6", paCodigoDeTerminal) Then
                GoTo errImp
            End If
        End If
        If .LoadPrinterConfig("6", paCodigoDeTerminal) Then
            .GetPrinterDoc 6, paPrintConfD, paPrintConfB, paPrintConfXDef, paPrintConfPaperSize
        End If
    End With
    If paPrintConfD = "" Then MsgBox "Por favor verifique la configuración de impresión.", vbInformation, "Atención"
    
errImp:
    Set objP = Nothing
    Screen.MousePointer = 0
End Sub


Public Function CargoDatosSucursal(ByVal sNombreTerminal As String, _
                                        Optional ByRef sNameCtdo As String = "", Optional ByRef sNameCred As String = "", _
                                        Optional ByRef sNameNCtdo As String = "", Optional ByRef sNameNCred As String = "", _
                                        Optional ByRef sNameRecibo As String = "", Optional ByRef sNameNEsp As String = "", Optional ByRef sNameNRemito As String = "") As Boolean
'................................................................................................................................................................
'Dado el nombre de la terminal
'   Cargo el código de la misma, el código de la sucursal y el nombre de los documentos.
'................................................................................................................................................................
On Error GoTo errCDS

    CargoDatosSucursal = False
    
    paCodigoDeSucursal = 0: paCodigoDeTerminal = 0
    sNameCtdo = "": sNameCred = ""
    sNameNCtdo = "": sNameNCred = ""
    sNameRecibo = "": sNameNEsp = ""
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & sNombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
    Set rsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        paCodigoDeSucursal = rsAux!TerSucursal
        paCodigoDeTerminal = rsAux!TerCodigo

        paNombreSucursal = Trim(rsAux!SucAbreviacion)
'        If Not IsNull(rsAux!SucDisponibilidad) Then paDisponibilidad = rsAux!SucDisponibilidad Else paDisponibilidad = 0
        
'        If Not IsNull(rsAux!SucDContado) Then sNameCtdo = Trim(rsAux!SucDContado)
'        If Not IsNull(rsAux!SucDCredito) Then sNameCred = Trim(rsAux!SucDCredito)
'        If Not IsNull(rsAux!SucDNDevolucion) Then sNameNCtdo = Trim(rsAux!SucDNDevolucion)
'        If Not IsNull(rsAux!SucDNCredito) Then sNameNCred = Trim(rsAux!SucDNCredito)
'        If Not IsNull(rsAux!SucDNEspecial) Then sNameNEsp = Trim(rsAux!SucDNEspecial)
'        If Not IsNull(rsAux!SucDRecibo) Then sNameRecibo = Trim(rsAux!SucDRecibo)
        If Not IsNull(rsAux("SucDRemito")) Then sNameNRemito = Trim(rsAux("SucDRemito"))
    End If
    rsAux.Close
  
    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(sNombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
                & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------------------------------
    CargoDatosSucursal = (paCodigoDeSucursal > 0)
    Exit Function

errCDS:
    MsgBox "Error al leer la información de la sucursal." & vbCr & vbCr & "Error: " & Err.Description, vbCritical, "Datos de Sucursal"
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

