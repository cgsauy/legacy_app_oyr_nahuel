Attribute VB_Name = "ModStart"
Option Explicit

'IMPRESORA
Public paValorUIUltMes As Currency
Public paIContadoB As Integer
Public paIContadoN As String
Public paPrintEsXDef As Boolean

Public paLastUpdate As String
Public paOptPrintSel As String      'El nombre de la opción seleccionada
Public paOptPrintList As String      'Los nombres de las opciones ingresadas están separadas x |
'...........................................................

Public prmImporteConInfoCliente As Double
Public paCategoriaDistribuidor As String

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
    remito = 6
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
Public gPathListados As String

Public Const FormatoCedula = "_.___.___-_"



Public clsGeneral As New clsorCGSA
Public miConexion As New clsConexion

Private Sub loc_CargoPrmGlobal()
    
    Cons = "Select * from Parametro " & _
                "Where ParNombre In ('efactImporteDatosCliente', 'TipoTCCompraME',  'MonedaFacturacion', 'TipoTelefonoP', 'TipoTelefonoE', 'TipoCuotaContado', " & _
                                                "'ArticuloPisoAgencia', 'ArticuloDiferenciaEnvio', 'TipoArticuloServicio', 'Cofis', 'MonedaDolar', 'MonedaPesos', " & _
                                                "'EstadoArticuloEntrega', 'FletesNoEnviarVtaTelef', 'catcliDistribuidor', 'ValorUIUltimoMes')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux("ParNombre")))
            Case LCase("efactImporteDatosCliente"): prmImporteConInfoCliente = RsAux("ParValor")
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
            Case LCase("catcliDistribuidor")
                If Not IsNull(RsAux("ParTexto")) Then paCategoriaDistribuidor = Trim(RsAux("ParTexto"))
            Case LCase("ValorUIUltimoMes"):
                paValorUIUltMes = RsAux("ParValor")
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    If paCategoriaDistribuidor <> "" Then paCategoriaDistribuidor = "," & Replace(paCategoriaDistribuidor, " ", "") & ","

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
    If miConexion.AccesoAlMenu(App.title) Then
        
        If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then
            Set clsGeneral = Nothing
            Set miConexion = Nothing
            End
            Exit Sub
        End If
        
        ChDir App.Path
        ChDir ("..")
        ChDir (CurDir & "\REPORTES\")
        gPathListados = CurDir & "\"
        
        InicializoEngine
        
        paCodigoDeUsuario = miConexion.UsuarioLogueado(True)
        paBD = miConexion.RetornoPropiedad(bDB:=True)
        
        loc_CargoPrmGlobal
        loc_GetInfoSucursal
        prj_LoadConfigPrint
        
        If Trim(Command()) <> "" Then
            If IsNumeric(Command()) Then
                m_idCli = CLng(Command())
            Else
                If LCase(Mid(Command(), 1, 1)) = "i" Then
                    m_idCli = 0
                    FacVtaTelefonica.prmIDVta = Mid(Command(), 2)
                End If
            End If
            FacVtaTelefonica.prmIDCliente = m_idCli
        End If
        FacVtaTelefonica.Show vbModeless
    End If
    
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Ocurrió un error al activar el ejecutable.", Trim(Err.Description)
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

'Public Function MsgClienteNoVender(ByVal iCliente As Long, ByVal bShowMsg As Boolean) As Boolean
'On Error GoTo errNV
'Dim rsCom As rdoResultset
'    MsgClienteNoVender = False
'    Set rsCom = cBase.OpenResultset("exec gennovender " & iCliente, rdOpenDynamic, rdConcurValues)
'    If Not rsCom.EOF Then
'        If Not IsNull(rsCom(0)) Then
'            If rsCom(0) = 1 Then
'                MsgClienteNoVender = True
'                If bShowMsg Then
'                    Screen.MousePointer = 0
'                    MsgBox "Atención: NO se puede vender sin autorización. Consultar con gerencia!", vbCritical, "ATENCIÓN"
'                End If
'            End If
'        End If
'    End If
'    rsCom.Close
'    Exit Function
'errNV:
'
'End Function


'Public Sub BuscoComentariosAlerta(idCliente As Long, _
'                                                   Optional Alerta As Boolean = False, Optional Cuota As Boolean = False, _
'                                                   Optional Decision As Boolean = False, Optional Informacion As Boolean = False)
'
'Dim rsCom As rdoResultset
'Dim aCom As String
'Dim sHay As Boolean
'
'    On Error GoTo errMenu
'    Screen.MousePointer = 11
'    sHay = False
'    'Armo el str con los comentarios a consultar-------------------------------------------------
'    If Not Alerta And Not Cuota And Not Decision And Not Informacion Then Exit Sub
'    aCom = ""
'    If Alerta Then aCom = aCom & Accion.Alerta & ", "
'    If Cuota Then aCom = aCom & Accion.Cuota & ", "
'    If Decision Then aCom = aCom & Accion.Decision & ", "
'    If Informacion Then aCom = aCom & Accion.Informacion & ", "
'    aCom = Mid(aCom, 1, Len(aCom) - 2)
'    '---------------------------------------------------------------------------------------------------
'
'    Cons = "Select * From Comentario, TipoComentario " _
'            & " Where ComCliente = " & idCliente _
'            & " And ComTipo = TCoCodigo " _
'            & " And TCoAccion IN (" & aCom & ")"
'    Set rsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
'    If Not rsCom.EOF Then sHay = True
'    rsCom.Close
'
'    If sHay Then
'        Dim aObj As New clsCliente
'        aObj.Comentarios idCliente:=idCliente
'        DoEvents
'        Set aObj = Nothing
'    End If
'    MsgClienteNoVender idCliente, True
'    Screen.MousePointer = 0
'    Exit Sub
'
'errMenu:
'    clsGeneral.OcurrioError "Ocurrió un error al acceder al fomulario de comentarios.", Err.Description
'    Screen.MousePointer = 0
'End Sub

Private Sub InicializoEngine()
On Error GoTo errIE
    If crAbroEngine = 0 Then MsgBox Trim(crMsgErr), vbCritical, "ATENCIÓN"
    Exit Sub
errIE:
    MsgBox "Error al inicializar el reporte de impresión.", vbCritical, "ATENCIÓN"
End Sub

Public Sub prj_LoadConfigPrint(Optional bShowFrm As Boolean = False)
On Error GoTo errLCP
Dim objPrint As New clsCnfgPrintDocument
Dim sPrint As String
Dim vPrint() As String
    With objPrint
         Set .Connect = cBase
        .Terminal = paCodigoDeTerminal
        If bShowFrm Then
            .CnfgTipoDocumento = "1"
            .ShowConfig
        End If
        sPrint = .getDocumentoImpresora(eTipoDocumento.Contado)
    End With
    Set objPrint = Nothing
    
    If sPrint <> "" Then
        vPrint = Split(sPrint, "|")
        paIContadoN = Trim(vPrint(0))
        paIContadoB = Trim(vPrint(1))
        paPrintEsXDef = (Val(vPrint(3)) = 1)
    End If
    Exit Sub
errLCP:
    MsgBox "Error al leer los datos de configuración: " & Err.Description, vbExclamation, "ATENCIÓN"
End Sub


Public Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer

    For Each X In Printers
        If Trim(X.DeviceName) = Trim(DeviceName) Then
            Set Printer = X
            Exit For
        End If
    Next
    
End Sub

