Attribute VB_Name = "ModMaCiudad"
Option Explicit

Public Const TranspTipo = 52
'Public clGeneral As New clsLibGeneral
Public clsGeneral As New clsorCGSA
Public UsuLogueado As Long
Public miconexion As New clsConexion

Public Sub Main()
On Error GoTo ErrMain
        
    If App.StartMode = vbSModeStandalone Then
        Screen.MousePointer = 11
        If miconexion.AccesoAlMenu(App.Title) Then
            InicioConexionBD miconexion.TextoConexion(logImportaciones)
            UsuLogueado = miconexion.UsuarioLogueado(True)
            frmMaCiudad.Show vbModeless
        Else
            If miconexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicaci�n.", vbExclamation, "ATENCI�N"
            End
            Screen.MousePointer = 0
        End If
    Else
        'Lo implementamos por si no esta corriendo login pida contrase�a.
        miconexion.AccesoAlMenu (App.Title)
        InicioConexionBD miconexion.TextoConexion(logImportaciones)
    End If
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
