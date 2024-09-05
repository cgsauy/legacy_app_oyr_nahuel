Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsLibGeneral
Public txtConexion As String

Public Sub Main()

    On Error GoTo errMain
    Dim miPrm As Long
    miPrm = Val(Trim(Command))
    If miPrm = 0 Then End
    
    Screen.MousePointer = 11
    miConexion.AccesoAlMenu ("Sucesos")
    
    'If Not miConexion.AccesoAlMenu("Sucesos") Then
    '    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    '    If paCodigoDeUsuario <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
    '    Screen.MousePointer = 0
    '    End
    'End If
    
    'paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    txtConexion = miConexion.TextoConexion("comercio")
    
    If Not InicioConexionBD(txtConexion) Then End
    
    'paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
           
    frmSuceso.prm_IdCliente = miPrm
    frmSuceso.Show vbModeless
    
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & vbCrLf & "Error: " & Trim(Err.Description)
    End
End Sub


