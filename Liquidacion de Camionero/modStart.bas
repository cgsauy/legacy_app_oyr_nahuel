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


Public itmx As ListItem, aTexto As String
Public prmPahtHTML As String
Public prmPlantilla As String

Public gPathListados As String
Public paTC As Integer
Public paMCDeposito As Long, paMCLiquidacionCamionero As Long
Public paDispBuzon As String
Public paxIVA As Currency
Public paMCDifLiqCamion As Long
Public paMCCtaDepositoEfectivo As Long
Public paMCCtaComisionRP As Long
            

Public prmCostoParada As Currency
Public prmCamionesParada As String
Public pathApp As String
Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public rdoCZureo As rdoConnection
Public prmUsuario As Long, prmPWDZureo As String
'---------------------------------------------------------------------------------------

Public Sub Main()
Dim aSucursal As String
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    paxIVA = 1.22
    
    Dim mAppVer As String
    mAppVer = App.Title
    If miConexion.AccesoAlMenu(mAppVer) Then
        
        If mAppVer <> "" And Not (App.Major & "." & App.Minor & "." & Format(App.Revision, "00")) >= mAppVer Then
            MsgBox "La versión del programa no es la última disponible." & vbCr & _
                        "Ud. debe actualizar el software.", vbExclamation, "Actualizar a Versión " & mAppVer
            End
        End If
        
        If Not InicioConexionBD(miConexion.TextoConexion("comercio")) Then
            End
            Exit Sub
        End If
        
        If Not fnc_ConnectZureo Then
            End
            Exit Sub
        End If
        
        If Not CargoDatosSucursal(miConexion.NombreTerminal) Then
            End
            Exit Sub
        End If
        
        CargoParametrosComercioYServicio
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        s_LoadPahtReportes
        frmListado.Show
        
    Else
        MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
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

Private Function fnc_ConnectZureo() As Boolean
    Dim oGeneric As New clsDBFncs
    If Not oGeneric.get_Connection(rdoCZureo, "ORG01", 10) Then
        MsgBox "Error al conectarse a la base de datos de Zureo.", vbExclamation, "Conexión Zureo"
    Else
        fnc_ConnectZureo = True
    End If
End Function

Private Sub s_LoadPahtReportes()
On Error GoTo errLP
    ChDir App.Path
    ChDir ("..")
    ChDir (CurDir & "\EPORTES\")
    gPathListados = CurDir & "\"
    Exit Sub
errLP:
    gPathListados = ""
End Sub
Public Sub CargoParametrosComercioYServicio()
On Error GoTo errCP
    'Parametros a cero--------------------------
    paMonedaPesos = 0
    paMCLiquidacionCamionero = 0
    Cons = "Select * from Parametro Where ParNombre IN ('monedapesos','mcliquidacioncamionero', 'MCCuentaDiferenciaLiqCamion', 'CamionesConParadas'" & _
                                                            ", 'MCCtaDepEfectivoCamioneros', 'MCCuentaComisionRedpagos', 'usuariozureo', 'MCDepositos', 'DisponibilidadBuzoneras',  'tipotccomprame', 'LiqCamPathHtml', 'plLiqCamionHtml')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case LCase("MCCuentaDiferenciaLiqCamion")
                paMCDifLiqCamion = RsAux("ParValor")
            
            Case LCase("MCCuentaComisionRedpagos")
                paMCCtaComisionRP = RsAux("ParValor")
            
            Case LCase("MCCtaDepEfectivoCamioneros")
                paMCCtaDepositoEfectivo = RsAux("ParValor")
            
            Case "mcdepositos": paMCDeposito = RsAux!ParValor
            Case "disponibilidadbuzoneras": paDispBuzon = Trim(RsAux!ParTexto)
            Case "monedapesos": paMonedaPesos = RsAux!ParValor
            Case "mcliquidacioncamionero": paMCLiquidacionCamionero = RsAux!ParValor
            Case "tipotccomprame": paTC = RsAux!ParValor
            Case LCase("LiqCamPathHtml"): prmPahtHTML = Trim(RsAux!ParTexto)
            Case LCase("plLiqCamionHtml"): prmPlantilla = Trim(RsAux!ParTexto)
            Case "usuariozureo"
                prmUsuario = RsAux("ParValor")
                prmPWDZureo = Trim(RsAux("ParTexto"))
            Case "camionesconparadas"
                If Not IsNull(RsAux("ParValor")) Then prmCostoParada = RsAux("ParValor")
                prmCamionesParada = Replace(RTrim(RsAux("ParTexto")), " ", "")
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If prmPahtHTML <> "" Then If VBA.Right(prmPahtHTML, 1) <> "\" Then prmPahtHTML = prmPahtHTML & "\"
    Exit Sub
errCP:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub


