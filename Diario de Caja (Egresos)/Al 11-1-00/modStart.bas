Attribute VB_Name = "modStart"
Option Explicit

Public clGeneral As New clsLibGeneral
Public miConexion As New clsConexion

Public Sub Main()
On Error GoTo ErrMain
        
    If App.StartMode = vbSModeStandalone Then
        Screen.MousePointer = 11
        If miConexion.AccesoAlMenu(App.Title) Then
            InicioConexionBD miConexion.TextoConexion(logImportaciones)
            paCodigoDeUsuario = miConexion.UsuarioLogueado(True)
            
            CargoParametrosSucursal
            CargoParametrosImportaciones
            CargoParametrosComercio
            CargoParametrosCaja
            
            frmDiarioEgresos.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
            frmDiarioEgresos.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
            frmDiarioEgresos.Status.Panels("bd") = "BD: " & PropiedadesConnect(miConexion.TextoConexion(logImportaciones), Database:=True) & " "
            
            frmDiarioEgresos.Show vbModeless
            
        Else
            If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicaci�n.", vbExclamation, "ATENCI�N"
            End
            Screen.MousePointer = 0
        End If
    Else
        'Lo implementamos por si no esta corriendo login pida contrase�a.
        miConexion.AccesoAlMenu (App.Title)
        InicioConexionBD miConexion.TextoConexion(logImportaciones)
    End If
    Exit Sub
ErrMain:
    clGeneral.OcurrioError "Ocurri� un error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
