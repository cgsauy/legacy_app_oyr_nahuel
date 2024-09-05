Attribute VB_Name = "modStart"
Option Explicit

'Definicion para Llamados a los formularios----------------------
Public Enum TipoLlamado
    Normal = 0                         'Desde el menú
    IngresoNuevo = 3                    'Para ingresar nuevos datos
    Modificacion = 7                    'Para modificar datos
    Visualizacion = 5                   'Llamado a clietnes
End Enum
Public itmx As ListItem

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public Sub Main()
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then
            Screen.MousePointer = 0
            End: Exit Sub
        End If
        frmTipoArticulo.Show
    Else
        MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description), vbCritical, "ATENCIÓN"
    End
End Sub

