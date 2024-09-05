Attribute VB_Name = "ModResLiquidacion"
Option Explicit

Public Const TranspTipo = 52
Public clGeneral As New clsLibGeneral
Public UsuLogueado As Long
Public miconexion As New clsConexion

Public Sub Main()
On Error GoTo ErrMain
        
    If App.StartMode = vbSModeStandalone Then
        Screen.MousePointer = 11
        If miconexion.AccesoAlMenu(App.Title) Then
            InicioConexionBD miconexion.TextoConexion(logImportaciones)
            UsuLogueado = miconexion.UsuarioLogueado(True)
            frmMaTransporte.Show vbModeless
        Else
            If miconexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
            End
            Screen.MousePointer = 0
        End If
    Else
        'Lo implementamos por si no esta corriendo login pida contraseña.
        miconexion.AccesoAlMenu (App.Title)
        InicioConexionBD miconexion.TextoConexion(logImportaciones)
    End If
    Exit Sub
ErrMain:
    clGeneral.OcurrioError "Ocurrio un error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
