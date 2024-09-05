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
Public paTipoArticuloServicio As Long

Public paTComEnvConf As Long            'Comentario Envío a Confirmar.
Public paUIDEnvConf As String

Public paPrimeraHoraEnvio As Long, paUltimaHoraEnvio As Long
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
        paCodigoDeUsuario = objConnect.UsuarioLogueado(True)
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
    Cons = "Select * From Parametro Where ParNombre IN('EstadoArticuloEntrega', 'TipoArticuloServicio', 'tcomentarioenvaconf'" & _
                                    ", 'menusuarioenvioconfirma', 'PrimeraHoraEnvio', 'UltimaHoraEnvio')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case "tipoarticuloservicio": paTipoArticuloServicio = RsAux!ParValor
            Case "tcomentarioenvaconf": paTComEnvConf = RsAux!ParValor
            Case "menusuarioenvioconfirma": paUIDEnvConf = Trim(RsAux!ParTexto)
'            Case "monedapesos": paMonedaPesos = RsAux!ParValor
'            Case LCase("MCVtaTelefonica"): paMCVtaTelefonica = RsAux!ParValor
            Case LCase("PrimeraHoraEnvio"): paPrimeraHoraEnvio = RsAux!ParValor
            Case LCase("UltimaHoraEnvio"): paUltimaHoraEnvio = RsAux!ParValor
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    f_GetParameters = (paEstadoArticuloEntrega > 0)
    Exit Function
errGP:
    objGral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Function

Public Sub CargoCombo(Consulta As String, Combo As Control, Optional Seleccionado As String = "")
Dim RsAuxiliar As rdoResultset
Dim iSel As Integer: iSel = -1     'Guardo el indice del seleccionado
    
On Error GoTo ErrCC
    
    Screen.MousePointer = 11
    Combo.Clear
    Set RsAuxiliar = cBase.OpenResultset(Consulta, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAuxiliar.EOF
        Combo.AddItem Trim(RsAuxiliar(1))
        Combo.ItemData(Combo.NewIndex) = RsAuxiliar(0)
        
        If Trim(RsAuxiliar(1)) = Trim(Seleccionado) Then iSel = Combo.ListCount
        RsAuxiliar.MoveNext
    Loop
    RsAuxiliar.Close
    
    If iSel = -1 Then Combo.ListIndex = iSel Else Combo.ListIndex = iSel - 1
    Screen.MousePointer = 0
    Exit Sub
    
ErrCC:
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al cargar el combo: " & Trim(Combo.Name) & "." & vbCrLf & Err.Description, vbCritical, "ERROR"
End Sub

Public Sub BuscoCodigoEnCombo(cCombo As Control, lngCodigo As Long)
Dim I As Integer
    
    If cCombo.ListCount > 0 Then
        For I = 0 To cCombo.ListCount - 1
            If cCombo.ItemData(I) = lngCodigo Then
                cCombo.ListIndex = I
                Exit Sub
            End If
        Next I
        cCombo.ListIndex = -1
    Else
        cCombo.ListIndex = -1
    End If

End Sub

