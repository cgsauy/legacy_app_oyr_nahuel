Attribute VB_Name = "modStart"
Option Explicit

Public prmURLFirmaEFactura As String
Public prmImporteConInfoCliente As Currency

Public Enum TipoLocal
    Camion = 1
    Deposito = 2
End Enum

Public Enum TipoEstadoMercaderia
    Fisico = 1
    Virtual = 2
End Enum

Public Enum TipoMovimientoEstado
    ARetirar = 1
    AEntregar = 2
    Reserva = 3
End Enum

Public Enum TipoDocumento
    'Documentos Facturacion
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
    
End Enum

Public Enum TipoSuceso
    AnulacionDeDocumentos = 2
    ModificacionDePrecios = 3
    RecepcionDeTraslados = 4
    CambioCostoDeFlete = 6
    ChequesDiferidos = 8
    CambioCategoriaCliente = 9
    Reimpresiones = 10
    DiferenciaDeArticulos = 11
    CederProductoServicio = 12
    FacturaArticuloInhabilitado = 13
    Notas = 14
    FacturaPlanInhabilitado = 15
    FacturaCambioNombre = 16
    CambioTipoArticuloServicio = 17
    ConfiguracionSistema = 18
    EliminarInstalacion = 21
    NotaArticuloSinDocumento = 22
    VariosStock = 98
    Varios = 99
End Enum



Public cbMen As rdoConnection       'Conexion a la Base de Datos de mensajes

Public paEstadoArticuloEntrega As Integer
Public paICartaB As Integer
Public paICartaN As String
Public paCategoriaMensajeStock As Long
Public paUserVerif As String


Public txtConexion As String
Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public Sub Main()
Dim txtConexionMen As String
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    
    If miConexion.AccesoAlMenu(App.Title) Then
                
        'Hago conexion al comercio
        txtConexion = miConexion.TextoConexion("Comercio")
        InicioConexionBD txtConexion
        
        Cons = "Select * from Parametro Where ParNombre IN('CategoriaMensajeStock', 'UsuarioSupervisor')"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            Select Case LCase(Trim(RsAux("ParNombre")))
                Case LCase("CategoriaMensajeStock"): paCategoriaMensajeStock = RsAux!ParValor
                Case LCase("UsuarioSupervisor"): paUserVerif = Trim(RsAux!ParTexto)
            End Select
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        If paCategoriaMensajeStock = 0 Then MsgBox "No existe el parametro 'CategoriaMensajeStock'.", vbCritical, "ATENCIÓN"
        CargoDatosSucursal miConexion.NombreTerminal
        CargoParametrosComercio
        
        'Hago conexion a Mensajeria
        txtConexionMen = miConexion.TextoConexion("login")
        InicioConexionMensaje txtConexionMen
        
        'MeCargoParametrosImpresion (paCodigoDeSucursal)
        
        RecTransferencia.Caption = "Recepción de Traslado (Sucursal: " & Trim(miConexion.NombreTerminal) & ") "
        RecTransferencia.Show vbModeless
    Else
        MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

'Private Sub MeCargoParametrosImpresion(Sucursal As Long)
'On Error GoTo errImp
'    paICartaN = "": paICartaB = -1
'    Cons = "Select * From Local Where LocCodigo = " & Sucursal
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'    If Not RsAux.EOF Then
'        If Not IsNull(RsAux!LocICaNombre) Then          'Carta.
'            paICartaN = Trim(RsAux!LocICaNombre)
'            If Not IsNull(RsAux!LocICaBandejaFlex) Then paICartaB = RsAux!LocICaBandejaFlex
'        End If
'       '------------------------------------------------------------------------------------------------------------------
'    End If
'    RsAux.Close
'    Exit Sub
'errImp:
'    Screen.MousePointer = 0
'    clsGeneral.OcurrioError "Error al cargar los parámetros de impresión. Informe del error a su administrador de base de datos.", Err.Description
'End Sub

Public Sub InicioConexionMensaje(strConexion As String, Optional sqlTimeOut As Integer = 15)
    
    On Error GoTo ErrICBD
    
    'Conexion a la base de datos----------------------------------------
    Set cbMen = eBase.OpenConnection("", , , strConexion)
    cbMen.QueryTimeout = sqlTimeOut
    
    Exit Sub
    
ErrICBD:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al intentar comunicarse con la Base de Datos, se cancelará la ejecución.", vbExclamation, "ATENCIÓN"
    
End Sub


Public Sub CargoParametrosComercio()

    Cons = "Select * from Parametro Where ParNombre IN('estadoarticuloentrega', 'QRenglonCtdo', 'UsuarioSupervisor', 'efactImporteDatosCliente', 'URLFirmaEFactura')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            'Case LCase("QRenglonCtdo"): prmQRenglon = RsAux("ParValor")
            Case LCase("efactImporteDatosCliente"): prmImporteConInfoCliente = RsAux("ParValor")
            Case LCase("URLFirmaEFactura"): prmURLFirmaEFactura = Trim(RsAux("ParTexto"))
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
'    If prmClaseArtStock = "" Then
'        MsgBox "El parámetro de clase de artículos que mueven stock es nulo, sin él los movimientos de stock pueden ocasionar"
'    End If
    
End Sub




'Public Function CargoParametrosSucursal() As String
'
'Dim aNombreTerminal As String
'
'    CargoParametrosSucursal = ""
'    aNombreTerminal = miConexion.NombreTerminal
'
'    paCodigoDeSucursal = 0
'    paCodigoDeTerminal = 0
'
'    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
'    Cons = "Select * From Terminal, Sucursal" _
'            & " Where TerNombre = '" & aNombreTerminal & "'" _
'            & " And TerSucursal = SucCodigo"
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'    If Not RsAux.EOF Then
'        paCodigoDeSucursal = RsAux!TerSucursal
'        paCodigoDeTerminal = RsAux!TerCodigo
'        CargoParametrosSucursal = Trim(RsAux!SucAbreviacion)
'    End If
'    RsAux.Close
'
'    If paCodigoDeSucursal = 0 Then
'        MsgBox "La terminal " & UCase(aNombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
'                & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
'        Exit Function
'    End If
'    '-------------------------------------------------------------------------------------------------------------------------
'
'End Function


