Attribute VB_Name = "ModTasaCambio"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsLibGeneral

'Public paMonedaDolar As Long

Sub main()
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        If Not InicioConexionBD(miConexion.TextoConexion("comercio")) Then End
        
        CargoParametrosLocales
        InTasaCambio.Show vbModeless
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title
    End
End Sub

Private Function CargoParametrosLocales()
    On Error GoTo errParam
    
    cons = "Select * from Parametro Where ParNombre like 'monedadolar%'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case LCase(Trim(rsAux!ParNombre))
            Case "monedadolar": paMonedaDolar = rsAux!ParValor
        End Select
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    Exit Function
errParam:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Function
