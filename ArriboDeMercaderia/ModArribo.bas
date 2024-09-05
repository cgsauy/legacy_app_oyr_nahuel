Attribute VB_Name = "ModArribo"
Option Explicit
Public UsuarioLogueado As Long
Public miConexion As New clsConexion

Sub Main()

    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion(logImportaciones)
        UsuarioLogueado = miConexion.UsuarioLogueado(True)
        VerArribo.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        VerArribo.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        VerArribo.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
        VerArribo.Show vbModeless
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
End Sub

Public Sub RelojA()
    Screen.MousePointer = 11
End Sub
Public Sub RelojD()
    Screen.MousePointer = 0
End Sub

