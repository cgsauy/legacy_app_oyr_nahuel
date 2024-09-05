Attribute VB_Name = "modStart"
Option Explicit

Public clsGeneral As New clsLibGeneral
Public miConexion As New clsConexion
Public txtConexion As String

Public Sub Main()
On Error GoTo ErrMain
        
    If App.StartMode = vbSModeStandalone Then
        Screen.MousePointer = 11
        If miConexion.AccesoAlMenu(App.Title) Then
            txtConexion = miConexion.TextoConexion(logComercio)
            InicioConexionBD txtConexion
            paCodigoDeUsuario = miConexion.UsuarioLogueado(True)
            
            CargoParametrosComercio
            
            frmPagos.Show vbModeless
            
        Else
            If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicaci�n.", vbExclamation, "ATENCI�N"
            End
            Screen.MousePointer = 0
        End If
    Else
        'Lo implementamos por si no esta corriendo login pida contrase�a.
        miConexion.AccesoAlMenu (App.Title)
        txtConexion = miConexion.TextoConexion(logComercio)
        InicioConexionBD txtConexion
    End If
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Ocurri� un error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
