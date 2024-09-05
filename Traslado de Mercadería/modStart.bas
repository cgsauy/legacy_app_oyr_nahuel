Attribute VB_Name = "modStart"
Option Explicit

'Public prmReportes As String

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

'Definicion de Tipos de Documentos----------------------
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

Public prmURLFirmaEFactura As String
Public prmImporteConInfoCliente As Currency

Public paPrintConfB As Integer
Public paPrintConfD As String
Public paPrintConfXDef As Boolean
Public paPrintConfPaperSize As Integer

'Public paPrintCtdoB As Integer
'Public paPrintCtdoD As String
'Public paPrintCtdoPaperSize As Integer

Public paEstadoArticuloEntrega As Integer
Public prmQRenglon As Byte
'Public prmUserSupervisor As Integer

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public Sub Main()
On Error GoTo errMain
Screen.MousePointer = 11
    
    If miConexion.AccesoAlMenu(App.Title) Then
        
        If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then
            End
        End If
        
        CargoParametrosComercio
        CargoDatosSucursal miConexion.NombreTerminal
        'CargoParametrosSucursal
        
'        ChDir App.Path
'        ChDir ("..")
'        ChDir (CurDir & "\REPORTES\")
'        prmReportes = CurDir & "\"
        
         prj_GetPrinter False
        TraMercaderia.Show vbModeless
        
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
'
'        'Ojo utilizo el prm de contado
'        If Not IsNull(RsAux!SucDTraslado) Then paDContado = Trim(RsAux!SucDTraslado)
'        CargoParametrosSucursal = Trim(RsAux!SucAbreviacion)
'    End If
'    RsAux.Close
'
'    If paCodigoDeSucursal = 0 Then
'        MsgBox "La terminal " & UCase(aNombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
'                & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
'        End
'        Exit Function
'    End If
'
'    If paDContado = "" Then
'        'MsgBox "No existe el documento traslado para su sucursal, avisele al administrador del sistema." & vbCr & "La ejecución será cancelada.", vbCritical, "Traslado de Mercadería"
'        'End
'        'Exit Function
'        MsgBox "No existe el documento traslado para su sucursal, avisele al administrador del sistema." & vbCr & "Sólo prodrá visualizar.", vbCritical, "Traslado de Mercadería"
'
'    End If
'    '-------------------------------------------------------------------------------------------------------------------------
'
'End Function

Public Sub CargoParametrosComercio()

    Cons = "Select * from Parametro Where ParNombre IN('estadoarticuloentrega', 'QRenglonCtdo', 'UsuarioSupervisor', 'efactImporteDatosCliente', 'URLFirmaEFactura')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case LCase("QRenglonCtdo"): prmQRenglon = RsAux("ParValor")
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
            '.GetPrinterDoc 1, paPrintCtdoD, paPrintCtdoB, paPrintConfXDef, paPrintCtdoPaperSize
        End If
    End With
    If paPrintConfD = "" Then MsgBox "Por favor verifique la configuración de impresión.", vbInformation, "Atención"
    
errImp:
    Set objP = Nothing
    Screen.MousePointer = 0
End Sub


