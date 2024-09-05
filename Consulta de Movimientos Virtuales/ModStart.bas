Attribute VB_Name = "ModStart"
Option Explicit

Public clsGeneral As New clsorCGSA
Public miconexion As New clsConexion
Public txtConexion As String

Public Sub Main()
On Error GoTo ErrMain
        
    If App.StartMode = vbSModeStandalone Then
        Screen.MousePointer = 11
        If miconexion.AccesoAlMenu(App.Title) Then
            txtConexion = miconexion.TextoConexion("comercio")
            InicioConexionBD txtConexion
            paCodigoDeUsuario = miconexion.UsuarioLogueado(True)
            FechaDelServidor
            frmMovVirtuales.Status.Panels("terminal").Text = "Terminal: " & miconexion.NombreTerminal
            frmMovVirtuales.Status.Panels("usuario").Text = "Usuario: " & miconexion.UsuarioLogueado(False, True)
            frmMovVirtuales.Show vbModeless
        Else
            If miconexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
            End
            Screen.MousePointer = 0
        End If
    Else
        'Lo implementamos por si no esta corriendo login pida contraseña.
        miconexion.AccesoAlMenu (App.Title)
        InicioConexionBD miconexion.TextoConexion("comercio")
    End If
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Ocurrio un error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
