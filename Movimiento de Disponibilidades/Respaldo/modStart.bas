Attribute VB_Name = "modStart"
Option Explicit

Public clGeneral As New clsLibGeneral
Public UsuLogueado As Long
Public miConexion As New clsConexion

Public Sub Main()
On Error GoTo ErrMain
        
    If App.StartMode = vbSModeStandalone Then
        Screen.MousePointer = 11
        If miConexion.AccesoAlMenu(App.Title) Then
            InicioConexionBD miConexion.TextoConexion("importaciones")
            UsuLogueado = miConexion.UsuarioLogueado(True)
            
            frmListado.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
            frmListado.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
            frmListado.Status.Panels("bd") = "BD: " & PropiedadesConnect(miConexion.TextoConexion("importaciones"), Database:=True) & " "
            
            frmListado.Show vbModeless
            
        Else
            If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
            End
            Screen.MousePointer = 0
        End If
    Else
        'Lo implementamos por si no esta corriendo login pida contraseña.
        miConexion.AccesoAlMenu (App.Title)
        InicioConexionBD miConexion.TextoConexion("importaciones")
    End If
    Exit Sub
ErrMain:
    clGeneral.OcurrioError "Ocurrií un error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
