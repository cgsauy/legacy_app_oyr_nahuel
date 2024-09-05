Attribute VB_Name = "modStart"
Option Explicit

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


'Stock
Public paEstadoArticuloEntrega As Integer

Public objGral As New clsorCGSA

Public Sub Main()
On Error GoTo errMain
Dim sAux As String
Dim objConnect As New clsConexion
    
    Screen.MousePointer = 11
    If Not objConnect.AccesoAlMenu("recepcion envios") Then
        MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        GoTo evEnd
    Else
        If Not InicioConexionBD(objConnect.TextoConexion("Comercio")) Then GoTo evEnd
        If Not CargoDatosSucursal(objConnect.NombreTerminal) Then GoTo evEnd
        If Not f_GetParameters Then MsgBox "Imposible continuar sin parámetros de stock.", vbCritical, "Atención": GoTo evEnd
        Set objConnect = Nothing
        frmRecepcionEnvio.Show
    End If
    
    Screen.MousePointer = 0
    Exit Sub

evEnd:
    Screen.MousePointer = 0
    Set objGral = Nothing
    Set objConnect = Nothing
    End
    Exit Sub
    
errMain:
    objGral.OcurrioError "Error al cargar la aplicación.", Err.Description
    Set objGral = Nothing
    Screen.MousePointer = 0
    End
End Sub

Private Function f_GetParameters() As Boolean
On Error GoTo errGP
    f_GetParameters = False
    Cons = "Select * From Parametro Where ParNombre IN('EstadoArticuloEntrega')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    f_GetParameters = (paEstadoArticuloEntrega > 0)
    Exit Function
errGP:
    objGral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Function
    
