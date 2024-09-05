Attribute VB_Name = "modStart"
Option Explicit

Public Enum TipoCliente
    Cliente = 1
    Empresa = 2
End Enum

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
    NotaDebito = 40
    
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

Public paLocalesService As String

'Public paIConformeB As Integer
'Public paIConformeN As String
Public paTipoComentario As Long

Public paClienteEmpresa As Long, paEstadoARecuperar As Integer, paEstadoArticuloEntrega As Integer, paEstadoRoto As Integer
Public sNombreEmpresa As String, sDireccion  As String
'Public idSucGallinal As Long

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
            prj_GetPrinter False
            
'            idSucGallinal = 0
'            Cons = "Select * from Sucursal where SucAbreviacion like 'gallinal'"
'            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'            If Not RsAux.EOF Then idSucGallinal = RsAux!SucCodigo
'            RsAux.Close
            
            paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
            frmCambioProducto.Show vbModeless
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
    paLocalesService = "0"
    Cons = "Select * from Parametro Where ParNombre Like 'Cliente%' or ParNombre Like 'Estado%' or ParNombre Like 'tComentarioCambioP%' Or ParNombre = 'LocalesDeService'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "clienteempresa": paClienteEmpresa = RsAux!ParValor
            Case "estadoarecuperar": paEstadoARecuperar = RsAux!ParValor
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case "tcomentariocambioprod": paTipoComentario = RsAux!ParValor
            Case LCase("LocalesDeService"): paLocalesService = RsAux("ParTexto")
            Case "estadoroto": paEstadoRoto = RsAux("ParValor")
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If paClienteEmpresa > 0 Then
        'Cargo el nombre para la impresión.
        Cons = "Select * From Cliente, CEmpresa Where CLiCodigo = " & paClienteEmpresa & " And CliCodigo = CEmCliente"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not IsNull(RsAux!CliCIRuc) Then
            sNombreEmpresa = " Cliente:|(" & clsGeneral.RetornoFormatoRuc(RsAux!CliCIRuc) & ") " & Trim(RsAux!CEmFantasia)
        Else
            sNombreEmpresa = " Cliente:|" & Trim(RsAux!CEmFantasia)
        End If
        If Not IsNull(RsAux!CliDireccion) Then
            sDireccion = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion)
        End If
        RsAux.Close
    Else
        sNombreEmpresa = " Cliente:| Empresa"
    End If
End Sub

'Private Sub CargoParametrosImpresionServicio(Sucursal As Long)
'On Error GoTo errImp
'
'    paIConformeN = ""
'    paIConformeB = -1
'
'    Cons = "Select * From Sucursal Where SucCodigo = " & Sucursal
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'
'    If Not RsAux.EOF Then
'        If Not IsNull(RsAux!SucICnNombre) Then          'CONFORME
'            paIConformeN = Trim(RsAux!SucICnNombre)
'            If Not IsNull(RsAux!SucICnBandeja) Then paIConformeB = RsAux!SucICnBandeja
'        End If
'       '------------------------------------------------------------------------------------------------------------------
'    End If
'    RsAux.Close
'    Exit Sub
'errImp:
'    Screen.MousePointer = 0
'    clsGeneral.OcurrioError "Error al cargar los parámetros de impresión. Informe del error a su administrador de base de datos.", Err.Description
'
'End Sub


