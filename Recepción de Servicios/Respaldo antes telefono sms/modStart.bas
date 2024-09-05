Attribute VB_Name = "modStart"
Option Explicit

'Trato de Stock.-------------------------------------
Public Enum TipoEstadoMercaderia
    Fisico = 1
    Virtual = 2
End Enum

Public Enum TipoMovimientoEstado
    ARetirar = 1
    AEntregar = 2
    Reserva = 3
End Enum
'---------------------------------------------------------

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


'agregue esta variable para unificar el codigo de gallinal con los otros.
Public idSucGallinal As Long

'Impresora
Public paPrintConfB As Integer
Public paPrintConfD As String
Public paPrintConfXDef As Boolean
Public paPrintConfPaperSize As Integer

Public paTipoFleteVentaTelefonica As Long, paCamionRetiroVisita As Long, paEstadoARecuperar As Integer, paEstadoArticuloEntrega As Integer
Public paCobroEnEntrega As Boolean
Public paMonedaDolar As Long, paMonedaPesos As Long
Public paClienteEmpresa As Long
Public paClienteAnglia As Long
Public paCoefFleteRetiro As Currency

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public prmTipoComentario As String
Public paPathReportes As String

Public Enum TipoLocal
    Camion = 1
    Deposito = 2
End Enum

'---------------------------------------------------------------------------------------
Public Sub Main()
Dim Argumento As String
Dim aSucursal As String, Invocacion As String, Articulo As Long, Direccion As Long
Dim aValor As Integer

    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        
        If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then
            Set miConexion = Nothing
            Set clsGeneral = Nothing
            End
            Exit Sub
        End If
        
        aSucursal = CargoParametrosSucursal
        
        prj_GetPrinter False
        
        CargoParametrosComercio

        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        idSucGallinal = 0
        Cons = "Select * from Sucursal where SucAbreviacion like 'gallinal'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then idSucGallinal = RsAux!SucCodigo
        RsAux.Close
        
        Invocacion = ""
        If Trim(Command()) <> "" Then Argumento = Trim(Command()) Else Argumento = ""
        
        'Los argumentos que recibos son de dos tipos:
            '1 =  solo viene el tipo de acción (taller, retiro o visita)
            '2 = L:t::a:d
            '       L = viene letra que dice si es taller, retiro o visita
            '       t = tipo de acción
            '       a = Id de articulo
            '       d = Direccion del cliente. (Si el tipo de llamado es a taller no es necesario, solo se necesita cuando es retiro)
        Articulo = 0: Direccion = 0
        If IsNumeric(Argumento) Then
            aValor = Val(Trim(Argumento))
        Else
            Invocacion = Mid(Trim(Argumento), 1, 1)
            Argumento = Mid(Trim(Argumento), 3, Len(Argumento))
            If InStr(1, Argumento, ":") > 0 Then
                aValor = Val(Mid(Trim(Argumento), 1, InStr(1, Argumento, ":") - 1))
                Argumento = Mid(Trim(Argumento), InStr(1, Argumento, ":") + 1, Len(Argumento))
                If InStr(1, Argumento, ":") > 0 Then
                    'Viene dirección
                    Articulo = Val(Mid(Trim(Argumento), 1, InStr(1, Argumento, ":") - 1))
                    Direccion = Val(Mid(Trim(Argumento), InStr(1, Argumento, ":") + 1, Len(Argumento)))
                Else
                    Articulo = Mid(Trim(Argumento), 1, Len(Argumento))
                End If
            Else
                aValor = Val(Mid(Trim(Argumento), 2, Len(Trim(Argumento))))
            End If
        End If
        frmRecServ.prmDireccion = Direccion
        frmRecServ.prmArticulo = Articulo
        frmRecServ.prmInvocacion = UCase(Invocacion)
        frmRecServ.prmTipoIngreso = aValor
        frmRecServ.prmSucursal = aSucursal
        frmRecServ.Show vbModeless
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Exit Sub
    
errMain:
    On Error Resume Next
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description)
    Screen.MousePointer = 0
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

Public Sub CargoParametrosComercio()

    'Parametros a cero--------------------------
    paTipoFleteVentaTelefonica = 0: paCamionRetiroVisita = 0: paEstadoARecuperar = 0
    paPrimeraHoraEnvio = 0: paUltimaHoraEnvio = 0
    paMonedaDolar = 0: paMonedaPesos = 0: paEstadoArticuloEntrega = 0
    paClienteEmpresa = 0: paClienteAnglia = 0: paCoefFleteRetiro = 1

    Cons = "Select * from Parametro"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case "estadoarecuperar": paEstadoARecuperar = RsAux!ParValor
            
            Case "primerahoraenvio": paPrimeraHoraEnvio = RsAux!ParValor
            Case "ultimahoraenvio": paUltimaHoraEnvio = RsAux!ParValor
            Case "tipofleteventatelefonica": paTipoFleteVentaTelefonica = RsAux!ParValor
            Case "camionretirovisita":
                paCamionRetiroVisita = RsAux!ParValor
                If Not IsNull(RsAux("ParTexto")) Then
                    paCobroEnEntrega = (Val(RsAux("ParTexto")) = 1)
                End If
            Case "serviciocoeffleteretiro": paCoefFleteRetiro = RsAux!ParValor
            
            Case "clienteempresa": paClienteEmpresa = RsAux!ParValor
            Case "clienteanglia": paClienteAnglia = RsAux!ParValor
            'OJO ESTOS SON LOS PARAMETROS COMUNES !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            Case "monedadolar": paMonedaDolar = RsAux!ParValor
            Case "monedapesos": paMonedaPesos = RsAux!ParValor
            
            Case LCase("TipComServRecepcion"): prmTipoComentario = Trim(RsAux!ParTexto)
            Case "pathreportes": paPathReportes = Trim(RsAux("ParTexto"))
        End Select
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub

Public Sub prj_GetPrinter(ByVal bShowP As Boolean)
On Error GoTo errImp
    paPrintConfD = ""
    paPrintConfB = -1
    Dim objP As New clslPrintConfig
    With objP
        If bShowP Then
            If Not .ShowPrinterSetup("6", paCodigoDeTerminal) Then
                GoTo errImp
            End If
        End If
        If .LoadPrinterConfig("6", paCodigoDeTerminal) Then
            .GetPrinterDoc 6, paPrintConfD, paPrintConfB, paPrintConfXDef, paPrintConfPaperSize
        End If
    End With
    If paPrintConfD = "" Then MsgBox "Por favor verifique la configuración de impresión.", vbInformation, "Atención"
    
errImp:
    Set objP = Nothing
    Screen.MousePointer = 0
End Sub

