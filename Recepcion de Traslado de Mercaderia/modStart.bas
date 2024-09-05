Attribute VB_Name = "modStart"
Option Explicit

Public paPrintConfD As String
Public paPrintConfB As Integer
Public paPrintConfXDef As Boolean
Public paPrintConfPaperSize As Integer

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
'Public paICartaB As Integer
'Public paICartaN As String
'Public paUserVerif As String

Public txtConexion As String
Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public ParametrosSist As New clsParametros

Public Sub Main()
Dim txtConexionMen As String
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    
    If miConexion.AccesoAlMenu(App.Title) Then
                
        'Hago conexion al comercio
        txtConexion = miConexion.TextoConexion("Comercio")
        InicioConexionBD txtConexion
        
'        Cons = "Select * from Parametro Where ParNombre IN('CategoriaMensajeStock', 'UsuarioSupervisor')"
'        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'        Do While Not RsAux.EOF
'            Select Case LCase(Trim(RsAux("ParNombre")))
'                Case LCase("CategoriaMensajeStock"): paCategoriaMensajeStock = RsAux!ParValor
'                Case LCase("UsuarioSupervisor"): paUserVerif = Trim(RsAux!ParTexto)
'            End Select
'            RsAux.MoveNext
'        Loop
'        RsAux.Close
        'If paCategoriaMensajeStock = 0 Then MsgBox "No existe el parametro 'CategoriaMensajeStock'.", vbCritical, "ATENCIÓN"
        
        CargoDatosSucursal miConexion.NombreTerminal
        
        Dim colPrms As New Collection
        colPrms.Add NombreDeParametros.efactImporteDatosCliente
        colPrms.Add NombreDeParametros.EstadoArticuloEntrega
        colPrms.Add NombreDeParametros.URLFirmaEFactura
        colPrms.Add NombreDeParametros.UsuarioSupervisor
        colPrms.Add NombreDeParametros.CategoriaMensajeStock
        ParametrosSist.CargoParametrosComercio colPrms
        
        paEstadoArticuloEntrega = ParametrosSist.ObtenerValorParametro(EstadoArticuloEntrega).Valor
        
        'CargoParametrosComercio
        
        'Hago conexion a Mensajeria
        txtConexionMen = miConexion.TextoConexion("login")
        InicioConexionMensaje txtConexionMen
        
        'MeCargoParametrosImpresion (paCodigoDeSucursal)
        
        prj_GetPrinter False
        
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

Public Sub prj_GetPrinter(ByVal bShowP As Boolean)
On Error GoTo errImp
    
    paPrintConfD = ""
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
    If paPrintConfD = "" Then MsgBox "Por favor verifique la configuración de impresión.", vbInformation, "Atención"
    
errImp:
    Set objP = Nothing
    Screen.MousePointer = 0
End Sub



