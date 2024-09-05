Attribute VB_Name = "modStart"
Option Explicit

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
    TrasladoServicio = 46
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

Public Enum TipoLocal
    Camion = 1
    Deposito = 2
End Enum

Public paEstadoArticuloEntrega As Integer
'Public paUltimaHoraEnvio As Long

Public paPrintCartaB As Integer
Public paPrintCartaD As String
Public paPrintCartaXDef As Boolean
Public paPrintCartaPaperSize As Integer


Public miConexion As New clsConexion
Public clsGeneral As New clsLibGeneral
'---------------------------------------------------------------------------------------

Public ParametrosSist As New clsParametros

Public Sub Main()
Dim aSucursal As String
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        If Not InicioConexionBD(miConexion.TextoConexion("comercio")) Then
            End
        End If
        
        Dim colPrms As New Collection
        colPrms.Add NombreDeParametros.efactImporteDatosCliente
        colPrms.Add NombreDeParametros.EstadoArticuloEntrega
        colPrms.Add NombreDeParametros.EstadoARecuperar
        colPrms.Add NombreDeParametros.URLFirmaEFactura
        colPrms.Add NombreDeParametros.ClienteEmpresa
        colPrms.Add NombreDeParametros.LocalCompañia
        ParametrosSist.CargoParametrosComercio colPrms
        
        paEstadoArticuloEntrega = ParametrosSist.ObtenerValorParametro(EstadoArticuloEntrega).Valor
        
        CargoDatosSucursal miConexion.NombreTerminal
        
        prj_GetPrinter False
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        frmListado.Show vbModeless
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description)
    End
End Sub
'
'Private Sub CargoParametrosDelComercio()
'    Cons = "Select * from Parametro Where ParNombre IN('clienteempresa', 'estadoarticuloentrega', 'estadoarecuperar', 'localcompañia')"
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'    Do While Not RsAux.EOF
'        Select Case LCase(Trim(RsAux!ParNombre))
'            Case "clienteempresa": paClienteEmpresa = RsAux!ParValor
'            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
'            Case "estadoarecuperar": paEstadoARecuperar = RsAux!ParValor
'            Case "localcompañia": paLocalCompañia = RsAux!ParValor
'        End Select
'        RsAux.MoveNext
'    Loop
'    RsAux.Close
'End Sub

Public Sub prj_GetPrinter(ByVal bShowP As Boolean)
On Error GoTo errImp
    paPrintCartaD = ""
    paPrintCartaB = -1
    Dim objP As New clslPrintConfig
    With objP
        If bShowP Then
            If Not .ShowPrinterSetup("21", paCodigoDeTerminal) Then
                GoTo errImp
            End If
        End If
        If .LoadPrinterConfig("21", paCodigoDeTerminal) Then
            .GetPrinterDoc 21, paPrintCartaD, paPrintCartaB, paPrintCartaXDef, paPrintCartaPaperSize
        End If
    End With
    If paPrintCartaD = "" Then MsgBox "Por favor verifique la configuración de impresión.", vbInformation, "Atención"
    
errImp:
    Set objP = Nothing
    Screen.MousePointer = 0
End Sub


