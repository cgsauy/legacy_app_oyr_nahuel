Attribute VB_Name = "modStart"
Option Explicit

Public prmPahtHTML As String
Public prmPlantilla As String
Public paMonedaPesos As Long

Public gPathListados As String
Public paTC As Integer

Public pathApp As String
Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
'---------------------------------------------------------------------------------------

Public Sub Main()
Dim aSucursal As String
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion("comercio")
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        CargoParametrosComercioYServicio
        frmLiquidacion.Show vbModeless
        Screen.MousePointer = 0
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
Public Sub CargoParametrosComercioYServicio()

    Cons = "Select * from Parametro Where ParNombre IN ('monedapesos','LiqCamPathHtml', 'plLiqCamionHtml')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "monedapesos": paMonedaPesos = RsAux!ParValor
            Case LCase("LiqCamPathHtml"): prmPahtHTML = Trim(RsAux!ParTexto)
            Case LCase("plLiqCamionHtml"): prmPlantilla = Trim(RsAux!ParTexto)
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    
    If prmPahtHTML <> "" Then If VBA.Right(prmPahtHTML, 1) <> "\" Then prmPahtHTML = prmPahtHTML & "\"
    
End Sub

Public Function BuscoNombreUsuario(Codigo As Long) As String
    BuscoNombreUsuario = BuscoUsuario(Codigo, True)
End Function

Public Function BuscoDigitoUsuario(Codigo As Long) As String
    BuscoDigitoUsuario = BuscoUsuario(Codigo, Digito:=True)
End Function

Public Function FechaDelServidor() As Date

    Dim RsF As rdoResultset
    On Error GoTo errFecha
    Cons = "Select GetDate()"
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    FechaDelServidor = RsF(0)
    RsF.Close
    
    Time = FechaDelServidor
    Date = FechaDelServidor
    Exit Function

errFecha:
    FechaDelServidor = Now
End Function

