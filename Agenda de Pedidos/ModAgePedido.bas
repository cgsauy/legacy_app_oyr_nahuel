Attribute VB_Name = "ModPedido"
Option Explicit

Public miconexion As New clsConexion

Public Enum Filtros
    Pendientes = 1
    Realizados = 2
End Enum


Sub main()
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miconexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miconexion.TextoConexion(logImportaciones)
        
        CargoParametrosImportaciones
        
        AgePedido.Status.Panels("terminal") = "Terminal: " & miconexion.NombreTerminal
        AgePedido.Status.Panels("usuario") = "Usuario: " & miconexion.UsuarioLogueado(Nombre:=True)
        AgePedido.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
                
        AgePedido.Show vbModeless
    
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


