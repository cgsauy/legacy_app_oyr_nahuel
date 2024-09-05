Attribute VB_Name = "ModStart"
Option Explicit

Public Enum TipoMovimientoEstado
    ARetirar = 1
    AEntregar = 2
    Reserva = 3
End Enum

'Definicion de Acciones de Comentario----------------------------------------------------------------------------
Private Enum Accion
    Informacion = 1     'No toma accion es un comentario +
    Alerta = 2             'Activa la pantalla de comentarios Todas
    Cuota = 3              'Activa en Cobranza, Decision, Visualizacion
    Decision = 4            'Activa en Decision
End Enum
'-----------------------------------------------------------------------------------------------------------------------

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
    
    VentaOnLineAConfirmar = 32
    VentaOnLineConfirmada = 33
    RemitoRecepcion = 34
    
    VentaRedPagosTelefonicas = 44
    
End Enum

'Definicion de Tipos de Envios------------------------------------------------------------------------------------
Public Enum TipoEnvio
    Entrega = 1
    Service = 2
    Cobranza = 3
End Enum
'-----------------------------------------------------------------------------------------------------------------------

Public Enum TipoEstadoMercaderia
    Fisico = 1
    Virtual = 2
End Enum

Public paMonedaFacturacion As Long
Public paTipoTelefonoP As Long              'Valor por defecto del tipo de telefono para las personas
Public paTipoTelefonoE As Long              'Valor por defecto del tipo de telefono para las empresas
Public paTipoCuotaContado As Long
Public paArticuloPisoAgencia As Long
Public paArticuloDiferenciaEnvio As Long
'Public paTipoArticuloServicio As Long
Public paEstadoArticuloEntrega As Integer
Public paCofis As Currency
Public paTCComME As Integer
Public paMonedaDolar As Integer
Public paMonedaPesos As Integer
Public paNoFletesVta As String


Public aTexto As String
Public paBD As String



Public clsGeneral As New clsorCGSA
Public miConexion As New clsConexion

Private Sub loc_CargoPrmGlobal()
    
    Cons = "Select * from Parametro " & _
                "Where ParNombre In ('TipoTCCompraME',  'MonedaFacturacion', 'TipoTelefonoP', 'TipoTelefonoE', 'TipoCuotaContado', " & _
                                                "'ArticuloPisoAgencia', 'ArticuloDiferenciaEnvio', 'TipoArticuloServicio', 'Cofis', 'MonedaDolar', 'MonedaPesos', " & _
                                                "'EstadoArticuloEntrega', 'FletesNoEnviarVtaTelef')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux("ParNombre")))
            Case "monedafacturacion": paMonedaFacturacion = RsAux!ParValor
            Case "tipotccomprame": paTCComME = RsAux!ParValor
            Case "tipotelefonop": paTipoTelefonoP = RsAux!ParValor
            Case "tipotelefonoe": paTipoTelefonoE = RsAux!ParValor
            Case "tipocuotacontado": paTipoCuotaContado = RsAux!ParValor
            Case "articulopisoagencia": paArticuloPisoAgencia = RsAux!ParValor
            Case "articulodiferenciaenvio": paArticuloDiferenciaEnvio = RsAux!ParValor
            'Case "tipoarticuloservicio": paTipoArticuloServicio = RsAux!ParValor
            Case "cofis": paCofis = RsAux!ParValor
            Case "monedadolar": paMonedaDolar = RsAux!ParValor
            Case "monedapesos": paMonedaPesos = RsAux!ParValor
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case "fletesnoenviarvtatelef": paNoFletesVta = RsAux("ParTexto")
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close

End Sub

Private Function loc_GetInfoSucursal() As Boolean

    loc_GetInfoSucursal = False
    
    paCodigoDeSucursal = 0
    paCodigoDeTerminal = 0
    paDContado = ""
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & miConexion.NombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDeTerminal = RsAux!TerCodigo
                
        'El documento que necesito solo es el contado.
        If Not IsNull(RsAux!SucDContado) Then paDContado = Trim(RsAux!SucDContado)
    End If
    RsAux.Close
  
    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(miConexion.NombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
                & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
        Exit Function
    Else
        loc_GetInfoSucursal = True
    End If
    '-------------------------------------------------------------------------------------------------------------------------
    
End Function

Public Sub Main()
On Error GoTo ErrMain
Dim m_idCli As Long
    
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        
        If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then
            Set clsGeneral = Nothing
            Set miConexion = Nothing
            End
            Exit Sub
        End If
        
        paCodigoDeUsuario = miConexion.UsuarioLogueado(True)
        paBD = miConexion.RetornoPropiedad(bDB:=True)
        
        loc_CargoPrmGlobal
        loc_GetInfoSucursal
        
        
        frmAnularVentas.Show
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Ocurrio un error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
    End
End Sub

Function BuscoNombreMoneda(Codigo As Long) As String

    On Error GoTo ErrBU
    Dim Rs As rdoResultset
    BuscoNombreMoneda = ""

    Cons = "SELECT * FROM Moneda WHERE MonCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not Rs.EOF Then BuscoNombreMoneda = Trim(Rs!MonNombre)
    Rs.Close
    Exit Function
    
ErrBU:
End Function

'-------------------------------------------------------------------------------------------------------
'   Carga un string con todos los articulos que corresponden a los fletes.
'   Se utiliza en aquellos formularios que no filtren los fletes
'-------------------------------------------------------------------------------------------------------
Public Function CargoArticulosDeFlete() As String
Dim Fletes As String
    On Error GoTo errCargar
    Fletes = ""
    
    'Cargo los articulos a descartar-----------------------------------------------------------
    Cons = "Select Distinct(TFlArticulo) from TipoFlete Where TFlArticulo <> Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        Fletes = Fletes & RsAux!TFlArticulo & ","
        RsAux.MoveNext
    Loop
    RsAux.Close
    Fletes = Fletes & paArticuloPisoAgencia & "," & paArticuloDiferenciaEnvio & ","
    '----------------------------------------------------------------------------------------------
    CargoArticulosDeFlete = Fletes
    Exit Function
    
errCargar:
    CargoArticulosDeFlete = Fletes
End Function

