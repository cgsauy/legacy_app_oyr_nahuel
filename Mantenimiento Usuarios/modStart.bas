Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsLibGeneral


Public Sub Main()
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion("login")
        frmMaUsuario.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        frmMaUsuario.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        frmMaUsuario.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        frmMaUsuario.Show
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Exit Sub

errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & _
                Err.Number & " - " & Err.Description
    End
End Sub
