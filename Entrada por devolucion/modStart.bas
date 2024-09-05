Attribute VB_Name = "modStart"
Option Explicit

'Impresora
Public paPrintConfB As Integer
Public paPrintConfD As String
Public paPrintConfXDef As Boolean
Public paPrintConfPaperSize As Integer

Public paArtsNoNotaEsp As String
Public paClienteEmpresa As Long
Public paClienteNoVtoCta As String

Public paArticuloPisoAgencia As Long, paArticuloDiferenciaEnvio As Long, paTipoArticuloServicio As Long
Public paEstadoARecuperar As Integer, paEstadoArticuloEntrega As Integer

Private miConexion As New clsConexion

Public clsGeneral As New clsorCGSA

Public Sub Main()
Dim aValor As Integer

    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        'Conexión
        If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then GoTo evFin
        'Códgio de Sucursal
        If Not CargoDatosSucursal(miConexion.NombreTerminal) Then GoTo evFin
                
        If Not CargoParametros Then GoTo evFin

        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        Set miConexion = Nothing
        
        prj_GetPrinter False
        frmIngDe.Show vbModeless
        
        
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        GoTo evFin
    End If
    Exit Sub
    
errMain:
    On Error Resume Next
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description)
    Screen.MousePointer = 0
    
evFin:
    Screen.MousePointer = 0
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End

    
End Sub

Private Function CargoParametros() As Boolean
'Controlo aquellos que son vitales si no los cargue finalizo la app.
On Error GoTo errCP
    
    'Parametros a cero--------------------------
    paEstadoARecuperar = 0: paEstadoArticuloEntrega = 0
    paArticuloPisoAgencia = 0: paArticuloDiferenciaEnvio = 0: paTipoArticuloServicio = 0

    Cons = "Select * from Parametro Where ParNombre IN('estadoarticuloentrega', 'estadoarecuperar', 'tipoarticuloservicio', " & _
                                                            "'articulopisoagencia', 'articulodiferenciaenvio', 'clienteempresa', 'ArtsNEspInhabilitado', " & _
                                                            "'ClienteNoCuotaVencida')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case "estadoarecuperar": paEstadoARecuperar = RsAux!ParValor
            Case "tipoarticuloservicio": paTipoArticuloServicio = RsAux!ParValor
            Case "articulopisoagencia": paArticuloPisoAgencia = RsAux!ParValor
            Case "articulodiferenciaenvio": paArticuloDiferenciaEnvio = RsAux!ParValor
            Case "clienteempresa": paClienteEmpresa = RsAux!ParValor
            Case LCase("ArtsNEspInhabilitado"): paArtsNoNotaEsp = Trim(RsAux!ParTexto)
            Case LCase("ClienteNoCuotaVencida"): paClienteNoVtoCta = Trim(RsAux!ParTexto)
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    paClienteNoVtoCta = Replace(paClienteNoVtoCta, " ", "")
    paArtsNoNotaEsp = Replace(paArtsNoNotaEsp, " ", "")
    
    
    CargoParametros = (paEstadoArticuloEntrega > 0 And paEstadoARecuperar > 0)
    If Not CargoParametros Then MsgBox "Los parámetros de Estado de stock no fueron leidos, no podrá continuar.", vbCritical, "Manejo de Stock"
    Exit Function
errCP:
     clsGeneral.OcurrioError "Error al leer los parámetros.", Err.Description
     CargoParametros = False
End Function

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
