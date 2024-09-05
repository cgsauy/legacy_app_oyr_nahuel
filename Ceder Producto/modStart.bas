Attribute VB_Name = "modStart"
Option Explicit

Public Enum TipoCliente
    Cliente = 1
    Empresa = 2
End Enum

'REGISTRO DE SUCESOS--------------------------------------
Public Enum TipoSuceso
    ModificacionDeMora = 1
    AnulacionDeDocumentos = 2
    ModificacionDePrecios = 3
    RecepcionDeTraslados = 4
    AnulacionDeEnvios = 5
    CambioCostoDeFlete = 6
    Direcciones = 7
    ChequesDiferidos = 8
    CambioCategoriaCliente = 9
    Reimpresiones = 10
    DiferenciaDeArticulos = 11
    CederProductoServicio = 12
    FacturaArticuloInhabilitado = 13
    Notas = 14
    FacturaPlanInhabilitado = 15
    VariosStock = 98
    Varios = 99
End Enum
'--------------------------------------------------------------------

'Definicion de Tipos de Documentos----------------------
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



Public paPrimeraHoraEnvio As Long
Public paUltimaHoraEnvio As Long

Public paClienteEmpresa As Long, paClienteAnglia As Long

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public txtConexion As String
'---------------------------------------------------------------------------------------

Public Sub Main()
Dim aSucursal As String
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        txtConexion = miConexion.TextoConexion("comercio")
        If InicioConexionBD(txtConexion) Then
            CargoParametrosSucursal
            CargoParametrosServicio
            paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
            frmCedoProducto.Show vbModeless
        End If
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

Public Function CargoParametrosSucursal() As String

Dim aNombreTerminal As String

    CargoParametrosSucursal = ""
    aNombreTerminal = miConexion.NombreTerminal
    
    paCodigoDeSucursal = 0
    paCodigoDeTerminal = 0
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & aNombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDeTerminal = RsAux!TerCodigo
'        If Not IsNull(RsAux!SucDisponibilidad) Then paDisponibilidad = RsAux!SucDisponibilidad Else paDisponibilidad = 0
        CargoParametrosSucursal = Trim(RsAux!SucAbreviacion)
    End If
    RsAux.Close
  
    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(aNombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
                & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
        End
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------------------------------
    
End Function

Private Sub CargoParametrosServicio()

    'Parametros a cero--------------------------

    Cons = "Select * from Parametro Where ParNombre Like 'Cliente%'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "clienteempresa": paClienteEmpresa = RsAux!ParValor
            Case "clienteanglia": paClienteAnglia = RsAux!ParValor
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub

