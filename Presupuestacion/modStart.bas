Attribute VB_Name = "modStart"
Option Explicit
Public paEstadoArticuloEntrega As Integer, paMonedaPesos As Long, paEstadoARecuperar As Integer, paEstadoRoto As Integer
Public paTipoArticuloServicio As Integer
Public paTipoCuotaContado As Integer
Public paMonedaDolar As Integer
Public paRepuesto As Long
Public pathApp As String
Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public txtConexion As String
Public paClienteEmpresa As Long

Public paPrimeraHoraEnvio As Long
Public paUltimaHoraEnvio As Long

Public Enum TipoLocal
    Camion = 1
    Deposito = 2
End Enum

Public Enum TipoCliente
    Cliente = 1
    Empresa = 2
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

Public Enum TipoEstadoMercaderia
    Fisico = 1
    Virtual = 2
End Enum
'---------------------------------------------------------------------------------------

Public Sub Main()
Dim aSucursal As String
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        txtConexion = miConexion.TextoConexion("comercio")
        InicioConexionBD txtConexion
        aSucursal = CargoParametrosSucursal
        CargoParametrosComercioYServicio
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        pathApp = App.Path
        frmPresupuestacion.prmSucursal = aSucursal
        If IsNumeric(Command()) Then frmPresupuestacion.prmServicio = Val(Command())
        frmPresupuestacion.Show vbModeless
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

Public Sub CargoParametrosComercioYServicio()

    'Parametros a cero--------------------------
    paEstadoArticuloEntrega = 0: paTipoCuotaContado = 0: paEstadoRoto = 0
    paRepuesto = 0
    
    Cons = "Select * from Parametro"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "estadoarecuperar": paEstadoARecuperar = RsAux!ParValor
            Case "estadoroto": paEstadoRoto = RsAux!ParValor
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case "tipoarticuloservicio": paTipoArticuloServicio = RsAux!ParValor
            Case "tipocuotacontado": paTipoCuotaContado = RsAux!ParValor
            Case "monedapesos": paMonedaPesos = RsAux!ParValor
            Case "monedadolar": paMonedaDolar = RsAux!ParValor
            Case "repuesto": paRepuesto = RsAux!ParValor
            Case "clienteempresa": paClienteEmpresa = RsAux!ParValor
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub

Public Function TasadeCambio(MOriginal As Integer, MDestino As Integer, Fecha As Date, Optional FechaTC As String = "") As Currency

Dim RsTC As rdoResultset

    On Error GoTo errTC
    TasadeCambio = 1
    Cons = "Select * from TasaCambio" _
            & " Where TCaFecha = (Select MAX(TCaFecha) from TasaCambio " _
                                          & " Where TCaFecha < '" & Format(Fecha, "mm/dd/yyyy 23:59") & "'" _
                                          & " And TCaOriginal = " & MOriginal _
                                          & " And TCaDestino = " & MDestino & ")" _
            & " And TCaOriginal = " & MOriginal _
            & " And TCaDestino = " & MDestino
            
    Set RsTC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsTC.EOF Then
        TasadeCambio = CCur(Format(RsTC!TCaComprador, "#.000"))
        FechaTC = Format(RsTC!TCaFecha, "dd/mm/yyyy")
    End If
    RsTC.Close
    Exit Function
    
errTC:
End Function


Public Function f_GetEventos(ByVal sAux As String) As String
On Error Resume Next
    f_GetEventos = ""
    If InStr(1, sAux, "[", vbTextCompare) = 1 And InStr(1, sAux, "/", vbTextCompare) > 1 And InStr(1, sAux, ":", vbTextCompare) > 2 And InStr(1, sAux, "]", vbTextCompare) > 1 Then
        f_GetEventos = Mid(sAux, InStr(1, sAux, "[", vbTextCompare), InStr(InStr(1, sAux, "[", vbTextCompare) + 1, sAux, "]"))
    End If
End Function

Public Function f_QuitarClavesDelComentario(ByVal sComentario As String) As String
Dim sAux As String
    sAux = f_GetEventos(sComentario)
    If sAux <> "" Then
        f_QuitarClavesDelComentario = Replace(sComentario, sAux, "")
    Else
        f_QuitarClavesDelComentario = sComentario
    End If
End Function



