Attribute VB_Name = "modStart"
Option Explicit

Public prmPathApp As String
Public prmPathHelps As String
Public prmPathProc As String

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Private txtConexion As String

Public Sub Main()
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    
    miConexion.AccesoAlMenu "orExplorer"
    paCodigoDeUsuario = miConexion.UsuarioLogueado(True)
    If paCodigoDeUsuario = 0 Then End
    
    txtConexion = miConexion.TextoConexion("login")
    If Not InicioConexionBD(txtConexion) Then End
    CargoParametrosLocales
    
    
    If Val(Command()) <> 0 Then
        If miConexion.AccesoAlMenu("Usuarios") Then
            paCodigoDeUsuario = Val(Command())
            
            Dim aTexto As String
            aTexto = BuscoUsuario(paCodigoDeUsuario, True)
            If aTexto = "" Then frmMenu.Caption = frmMenu.Caption & " (NN " & paCodigoDeUsuario & ")"
            If aTexto <> "" Then frmMenu.Caption = frmMenu.Caption & " (" & aTexto & ")"
        End If
    End If
    
    frmMenu.Show
    
    Exit Sub
    
errMain:
    MsgBox "Error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description)
    Screen.MousePointer = 0
    End
End Sub


Private Sub CargoParametrosLocales()
    On Error GoTo errCP
       
    cons = "Select * From cgsa.dbo.parametro where ParNombre like  'Path%'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        
        Select Case (Trim(LCase(rsAux!ParNombre)))
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto)
            Case "pathhelps": prmPathHelps = Trim(rsAux!ParTexto)
            Case "pathhelpsproc": prmPathProc = Trim(rsAux!ParTexto)
        End Select
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    
errCP:
End Sub

