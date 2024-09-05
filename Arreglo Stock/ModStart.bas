Attribute VB_Name = "ModStart"
Option Explicit
Public Const logComercio = "Comercio"
Public Const TranspTipo = 52
Public clGeneral As New clsLibGeneral
Public miconexion As New clsConexion

Public Sub Main()
On Error GoTo ErrMain
    If App.StartMode = vbSModeStandalone Then
        Screen.MousePointer = 11
        If miconexion.AccesoAlMenu(App.Title) Then
            InicioConexionBD miconexion.TextoConexion(logComercio)
            CargoParametrosComercio
            paCodigoDeUsuario = miconexion.UsuarioLogueado(True)
            frmArrST.Show vbModeless
        Else
            If miconexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
            End
            Screen.MousePointer = 0
        End If
    Else
        'Lo implementamos por si no esta corriendo login pida contraseña.
        miconexion.AccesoAlMenu (App.Title)
        InicioConexionBD miconexion.TextoConexion(logComercio)
    End If
    Exit Sub
ErrMain:
    clGeneral.OcurrioError "Ocurrio un error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
