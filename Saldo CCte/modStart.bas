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
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicaci�n.", vbExclamation, "ATENCI�N"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurri� un error al inicializar la aplicaci�n " & App.Title
    End
End Sub

