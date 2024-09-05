Attribute VB_Name = "modStart"
Option Explicit

Public clsGeneral As New clsorCGSA
Public miConexion As New clsConexion
Public txtConexion As String

Public paLocalZF As Long, paLocalPuerto As Long

Public Sub Main()
On Error GoTo ErrMain
        
    If App.StartMode = vbSModeStandalone Then
        Screen.MousePointer = 11
        If miConexion.AccesoAlMenu(App.Title) Then
            txtConexion = miConexion.TextoConexion("comercio")
            InicioConexionBD txtConexion
            paCodigoDeUsuario = miConexion.UsuarioLogueado(True)
            
            'CargoParametrosSucursal
            'CargoParametrosImportaciones
            'CargoParametrosCaja
            CargoParametrosComercio
            
            frmListado.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
            frmListado.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
            frmListado.Status.Panels("bd") = "BD: " & PropiedadesConnect(txtConexion, Database:=True) & " "
            
            frmListado.Show vbModeless
            
        Else
            If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
            End
            Screen.MousePointer = 0
        End If
    Else
        'Lo implementamos por si no esta corriendo login pida contraseña.
        miConexion.AccesoAlMenu (App.Title)
        txtConexion = miConexion.TextoConexion("comercio")
        InicioConexionBD txtConexion
    End If
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
