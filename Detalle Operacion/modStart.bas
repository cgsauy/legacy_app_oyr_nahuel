Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public txtConexion As String

Public paToleranciaMora As Integer
Public paCoeficienteMora As Currency

Public paIconoVencimientoN2Dias As Integer
Public paIconoPendienteN2Dias As Integer

Public Sub Main()


    On Error GoTo errMain
    Screen.MousePointer = 11
    
    'If Not miConexion.AccesoAlMenu(App.Title) Then
    miConexion.AccesoAlMenu (App.Title)
    txtConexion = miConexion.TextoConexion(logFacturacion)
    InicioConexionBD txtConexion
    
    CargoParametrosLocales
    dis_CargoArrayMonedas
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    frmDeOperacion.Show vbModeless
    
'    Else
'        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
'        End
'    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocales()

    On Error GoTo errParametro
    Cons = "Select * from Parametro"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            
            Case "toleranciamora": paToleranciaMora = RsAux!ParValor
            Case "coeficientemora": paCoeficienteMora = ((RsAux!ParValor / 100) + 1) ^ (1 / 30)                     'Como es mensual calculo el diario
            
            Case "iconovencimienton2dias": If Not IsNull(RsAux!ParValor) Then paIconoVencimientoN2Dias = RsAux!ParValor
            Case "iconopendienten2dias": If Not IsNull(RsAux!ParValor) Then paIconoPendienteN2Dias = RsAux!ParValor

        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close

    Exit Sub
errParametro:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los parámetros.", Err.Description
End Sub

