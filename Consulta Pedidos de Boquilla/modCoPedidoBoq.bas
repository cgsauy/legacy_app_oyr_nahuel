Attribute VB_Name = "modCoPedidoBo"

Public miConexion As New clsConexion

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion(logImportaciones)
        
        CargoParametrosImportaciones
        
        CoPedidoBoquilla.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        CoPedidoBoquilla.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        CoPedidoBoquilla.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
                
        CoPedidoBoquilla.Show vbModeless
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicaci�n.", vbExclamation, "ATENCI�N"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurri� un error al inicializar la aplicaci�n " & App.Title
    End
End Sub
