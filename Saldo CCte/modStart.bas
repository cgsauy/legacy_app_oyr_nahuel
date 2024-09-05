Attribute VB_Name = "modStart"
Option Explicit

Public clsGeneral As New clsLibGeneral
Public miConexion As New clsConexion
Public txtConexion As String

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    
    If miConexion.AccesoAlMenu(App.Title) Then
        txtConexion = miConexion.TextoConexion("comercio")
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        InicioConexionBD txtConexion
        
        frmSaldoCC.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        frmSaldoCC.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        frmSaldoCC.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
        
        frmSaldoCC.Show vbModeless
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title
    End
End Sub

