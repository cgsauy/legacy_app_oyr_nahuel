Attribute VB_Name = "ModPrecioCosto"
Option Explicit
Public txtConexion As String
Public miconexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public Enum Filtros
    Pendientes = 1
    Realizados = 2
End Enum


Sub main()
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miconexion.AccesoAlMenu(App.Title) Then
        txtConexion = miconexion.TextoConexion(logImportaciones)
        InicioConexionBD txtConexion
        
        CargoParametrosImportaciones
                
        MaPrecioCosto.Show vbModeless
    
    Else
        If miconexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title
    End
End Sub

Public Sub RelojA()
    Screen.MousePointer = 11
End Sub
Public Sub RelojD()
    Screen.MousePointer = 0
End Sub
Public Function TextoValido(S As String)
    If InStr(S, "'") > 0 Then TextoValido = False Else TextoValido = True
End Function


