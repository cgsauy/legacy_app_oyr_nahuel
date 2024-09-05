Attribute VB_Name = "modStart"
Option Explicit

'MODULO Conección
'Contiene rutinas y variables del entorno RDO.

'Definición del entorno RDO
Public cBase As rdoConnection       'Conexion a la Base de Datos
Public eBase As rdoEnvironment     'Definicion de entorno
Public RsAux As rdoResultset         'Resultset Auxiliar

Public Cons As String

'Usuario y Terminal
Public paCodigoDeUsuario As Long
Public paCodigoDeSucursal As Long
Public paCodigoDeTerminal As Long
Public paNombreSucursal As String
Public paDisponibilidad As Long

'Fecha del Servidor
Public gFechaServidor As Date

Public clsGeneral As New clsorCGSA
Public paEstadoArticuloEntrega As Integer
Public paTipoArticuloServicio As Integer
Public paSonidoTimbre As String
Public paArrimar As Byte    '1= Prendido

Public Enum TipoMovimientoEstado
    ARetirar = 1
    AEntregar = 2
    Reserva = 3
End Enum
Public Enum TipoEstadoMercaderia
    Fisico = 1
    Virtual = 2
End Enum

Public Enum TipoDocumento
    Servicio = 0
    Contado = 1
    Credito = 2
    NotaDevolucion = 3
    NotaCredito = 4
    ReciboDePago = 5
    Remito = 6
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
    
    Traslados = 20
    Envios = 21
    CambioEstadoMercaderia = 22
    IngresoMercaderiaEspecial = 24
    ArregloStock = 25
End Enum

Public Sub Main()
Dim miConexion As clsConexion
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    Set miConexion = New clsConexion
    'Si da error la conexión la misma despliega el msg de error
    If Not miConexion.AccesoAlMenu(App.Title) Then
        Screen.MousePointer = 0
        MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        End
    Else
        If Not ObtenerConexionBD(cBase, "comercio") Then
            Screen.MousePointer = 0
            End: Exit Sub
        End If
        'Guardo el usuario logueado
        CargoDatosSucursal miConexion.NombreTerminal
        Set miConexion = Nothing
        
        CargoParametros
        Screen.MousePointer = 0
        frmMostrador.Show
    End If
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description) & vbCr, vbCritical, "ATENCIÓN"
    End
End Sub

Private Function CargoParametros() As Boolean
'Controlo aquellos que son vitales si no los cargue finalizo la app.
On Error GoTo errCP
    
    'Parametros a cero--------------------------
    paEstadoArticuloEntrega = 0
    paTipoArticuloServicio = 0

    Cons = "Select * from Parametro Where ParNombre IN('estadoarticuloentrega', 'tipoarticuloservicio', 'dep_wav_EntregarTimbre', 'dep_Estado_Arrimar_" & paCodigoDeSucursal & "')"
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case "tipoarticuloservicio": paTipoArticuloServicio = RsAux!ParValor
            Case "dep_wav_entregartimbre": paSonidoTimbre = Trim(RsAux("ParTexto"))
            Case LCase("dep_Estado_Arrimar_") & paCodigoDeSucursal: paArrimar = RsAux("ParValor")
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    CargoParametros = (paEstadoArticuloEntrega > 0)
    If Not CargoParametros Then MsgBox "Los parámetros de Estado de stock no fueron leidos, no podrá continuar.", vbCritical, "Manejo de Stock"
    Exit Function
errCP:
     clsGeneral.OcurrioError "Error al leer los parámetros.", Err.Description
     CargoParametros = False
End Function

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
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
    If Not RsAux.EOF Then
        paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDeTerminal = RsAux!TerCodigo
        paNombreSucursal = Trim(RsAux!SucAbreviacion)
        
        If Not IsNull(RsAux!SucDisponibilidad) Then paDisponibilidad = RsAux!SucDisponibilidad Else paDisponibilidad = 0
        
        If Not IsNull(RsAux!SucDContado) Then sNameCtdo = Trim(RsAux!SucDContado)
        If Not IsNull(RsAux!SucDCredito) Then sNameCred = Trim(RsAux!SucDCredito)
        If Not IsNull(RsAux!SucDNDevolucion) Then sNameNCtdo = Trim(RsAux!SucDNDevolucion)
        If Not IsNull(RsAux!SucDNCredito) Then sNameNCred = Trim(RsAux!SucDNCredito)
        If Not IsNull(RsAux!SucDNEspecial) Then sNameNEsp = Trim(RsAux!SucDNEspecial)
        If Not IsNull(RsAux!SucDRecibo) Then sNameRecibo = Trim(RsAux!SucDRecibo)
        If Not IsNull(RsAux("SucDRemito")) Then sNameNRemito = Trim(RsAux("SucDRemito"))
    End If
    RsAux.Close
  
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
    Date = gFechaServidor
    Time = gFechaServidor
    Exit Sub

errFecha:
    gFechaServidor = Now
End Sub

