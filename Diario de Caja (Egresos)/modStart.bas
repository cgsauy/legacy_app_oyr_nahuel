Attribute VB_Name = "modStart"
Option Explicit

Public clsGeneral As New clsorCGSA
Public miConexion As New clsConexion

Public paLocalZF As Integer, paLocalPuerto As Integer

Public Sub Main()
On Error GoTo ErrMain
        
    If App.StartMode = vbSModeStandalone Then
        Screen.MousePointer = 11
        If miConexion.AccesoAlMenu(App.Title) Then
            InicioConexionBD miConexion.TextoConexion("comercio")
            paCodigoDeUsuario = miConexion.UsuarioLogueado(True)
            
            CargoParametrosSucursal
            'CargoParametrosImportaciones
            CargoParametrosComercio
            CargoParametrosCaja
            
            frmDiarioEgresos.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
            frmDiarioEgresos.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
            frmDiarioEgresos.Status.Panels("bd") = "BD: " & PropiedadesConnect(miConexion.TextoConexion("comercio"), Database:=True) & " "
            
            frmDiarioEgresos.Show vbModeless
            
        Else
            If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
            End
            Screen.MousePointer = 0
        End If
    Else
        'Lo implementamos por si no esta corriendo login pida contraseña.
        miConexion.AccesoAlMenu (App.Title)
        InicioConexionBD miConexion.TextoConexion("comercio")
    End If
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
