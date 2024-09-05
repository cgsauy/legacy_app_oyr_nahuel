Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public prmPathListados As String
Public paBD As String

Public paTipoCuotaContado As Long
Public paMonedaPesos As Long
Public paCuotaMin As Currency
Public paPlanPorDefecto As Long

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    Dim aTexto As String
    
    If Not miConexion.AccesoAlMenu("Listas de Precios") Then
        MsgBox "Acceso denegado. " & vbCrLf & "Consulte a su administrador de Sistemas", vbExclamation, "Acceso Denegado"
        End
    End If
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    
    If Not InicioConexionBD(miConexion.TextoConexion("comercio")) Then End
    
    paBD = miConexion.RetornoPropiedad(bDB:=True)
    
    CargoParametrosLocal
    'prmPathListados = "C:\Proyectos\Precios\Reportes\"
    frmListas.Show vbModeless
    
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocal()
On Error Resume Next
    prmPathListados = ""
    
    Cons = "Select * from Parametro Where ParNombre In ('pathapp', 'TipoCuotaContado', 'MonedaPesos', 'webminimportecuota', 'PlanPorDefecto')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "webminimportecuota": paCuotaMin = RsAux("ParValor")
            Case "pathapp": prmPathListados = Trim(RsAux!ParTexto)
            Case "tipocuotacontado": paTipoCuotaContado = RsAux!ParValor
            Case "monedapesos": paMonedaPesos = RsAux!ParValor
            Case LCase("PlanPorDefecto"): paPlanPorDefecto = RsAux!ParValor
        End Select
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Cons = ""
    Dim aPos As Integer, aT2 As String
    aT2 = prmPathListados
    Do While InStr(aT2, "\") <> 0
        aPos = InStr(aT2, "\")
        Cons = Cons & Mid(aT2, 1, aPos)
        aT2 = Mid(aT2, aPos + 1)
    Loop
    prmPathListados = Cons & "Reportes\"
    
    
    'paCodigoDeSucursal
'    cons = miConexion.NombreTerminal
'    cons = "Select * from Terminal Where TerNombre = '" & cons & "'"
'    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
'    If Not rsAux.EOF Then If Not IsNull(rsAux!TerSucursal) Then paCodigoDeSucursal = rsAux!TerSucursal
'    rsAux.Close
    
End Sub
