Attribute VB_Name = "modStart"
Option Explicit

'Definicion para Llamados a los formularios----------------------
Public Enum TipoLlamado
    Normal = 0                         'Desde el men�
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
        MsgBox "Ud. no tiene permisos de acceso para la aplicaci�n.", vbExclamation, "ATENCI�N"
        Screen.MousePointer = 0
        End
    End If
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurri� un error al inicializar la aplicaci�n " & App.Title & Chr(13) & "Error: " & Trim(Err.Description), vbCritical, "ATENCI�N"
    End
End Sub

