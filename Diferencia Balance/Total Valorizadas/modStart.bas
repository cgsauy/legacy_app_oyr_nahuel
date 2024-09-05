Attribute VB_Name = "modStart"

Public miConexion As New clsConexion
Public clsGeneral As New clsLibGeneral

Public paEstadoArticuloARecuperar As Long
Public paGrupoRepuesto As Long

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    miConexion.AccesoAlMenu (App.Title)
    InicioConexionBD miConexion.TextoConexion(logComercio)
    
    CargoParametrosLocal
    
    frmListado.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    frmListado.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    frmListado.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
            
    frmListado.Show vbModeless
    
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocal()

    Cons = "Select * from Parametro"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "estadoarecuperar": paEstadoArticuloARecuperar = RsAux!ParValor
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            
            Case "repuesto": paGrupoRepuesto = RsAux!ParValor
            
        End Select
        
        RsAux.MoveNext
    Loop
    RsAux.Close
End Sub
