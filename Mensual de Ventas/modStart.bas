Attribute VB_Name = "modStart"

Public miConexion As New clsConexion
Public clsGeneral As New clsLibGeneral
Public txtConexion As String
Public prmPlantilla As Integer

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        txtConexion = miConexion.TextoConexion("comercio")
        InicioConexionBD txtConexion
        
        Cons = "SELECT ParValor From Parametro Where ParNombre = 'PlantillaVentasMesXDia'"
        Set RsAux = cBase.OpenResultset(Cons, rdopnedynamic, rdConcurValues)
        If Not RsAux.EOF Then prmPlantilla = RsAux("ParValor")
        RsAux.Close
        
        
        
        frmListado.Show vbModeless
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description)
    End
End Sub
