Attribute VB_Name = "modProject"
Option Explicit

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
    
    Traslados = 20
    Envios = 21
    CambioEstadoMercaderia = 22
    IngresoMercaderiaEspecial = 24
    ArregloStock = 25
    Servicio = 26
    ServicioCambioEstado = 27
    Devolucion = 28
End Enum

Public paEstadoArticuloEntrega  As Integer
Public clsGeneral As New clsorCGSA
Public gFechaServidor As Date
'----------------------------------------------------------------------------------------------------
'   Consulta por la fecha del servidor y la carga en la variable global gFechaServidor
'----------------------------------------------------------------------------------------------------
Public Sub FechaDelServidor()

    Dim RsF As rdoResultset
    On Error GoTo errFecha
    Cons = "Select GetDate()"
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    gFechaServidor = RsF(0)
    RsF.Close
    
    Time = gFechaServidor
    Date = gFechaServidor
    Exit Sub

errFecha:
    gFechaServidor = Now
End Sub

Sub Main()
On Error GoTo errMain
Dim objConnect As New clsConexion

    If objConnect.AccesoAlMenu(App.Title) Then
        paCodigoDeUsuario = objConnect.UsuarioLogueado(True)
        If InicioConexionBD(objConnect.TextoConexion("Comercio")) Then
            If CargoParametrosSucursal(objConnect.NombreTerminal) Then
                CargoParametrosComercio
                If paEstadoArticuloEntrega = 0 Then
                    CierroConexion
                    MsgBox "No se obtuvo el código del estado sano, no se podrá continuar.", vbCritical, "Atención"
                Else
                    frmAnulo.Show
                End If
            Else
                CierroConexion
            End If
        End If
    End If
    Set objConnect = Nothing
    Exit Sub
    
errMain:
    clsGeneral.OcurrioError "Error al iniciar la aplicación.", Err.Description
End Sub

Private Function CargoParametrosSucursal(ByVal sTerminal As String) As Boolean
On Error GoTo errCPS

    CargoParametrosSucursal = True
    
    paCodigoDeSucursal = 0
    paCodigoDeTerminal = 0

    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & Trim(sTerminal) & "'" _
            & " And TerSucursal = SucCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDeTerminal = RsAux!TerCodigo
    End If
    RsAux.Close

    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(sTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
                & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
    Else
        CargoParametrosSucursal = True
    End If
    Exit Function

errCPS:
    clsGeneral.OcurrioError "Error al obtener los datos de la terminal.", Err.Description
End Function

Private Sub CargoParametrosComercio()

    'Parametros a cero--------------------------
    paEstadoArticuloEntrega = 0

    Cons = "Select * from Parametro Where ParNombre = 'estadoarticuloentrega'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub

