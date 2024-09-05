Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

'Public paLocalZF As Long
Public prmMenUsuarioImportacion As String
Public prmMenUsuarioSistema As Long

Public Sub Main()

    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion("comercio")
        
        CargoParametrosLocal
        CargoParametrosImportaciones
        CargoParametrosComercio
        CargoParametrosSucursal
        
        frmInMercaderia.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        frmInMercaderia.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        frmInMercaderia.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
                
        frmInMercaderia.Show vbModeless
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocal()

    On Error GoTo errCPM
    cons = "Select * from Parametro Where ParNombre like 'men%'" & _
                                                    " Or ParNombre like '%usuario%'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case Trim(LCase(rsAux!ParNombre))
            Case "menusuarioimportacion": If Not IsNull(rsAux!ParTexto) Then prmMenUsuarioImportacion = Trim(rsAux!ParTexto)
            Case "usuariosistema": prmMenUsuarioSistema = rsAux!ParValor
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
    
errCPM:
End Sub
