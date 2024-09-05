Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public paLocalZF As Long, paLocalPuerto As Long
Public prmTCVentasME As Integer
Public prmPathApp As String

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion(logComercio), sqlTimeOut:=30
        
        CargoParametrosComercio
        CargoParametrosCaja
        CargoParametrosSucursal
        
        CargoParametrosLocal
        
        frmListado.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        frmListado.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        frmListado.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
                
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
    MsgBox "Error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Function CargoParametrosLocal()
On Error GoTo errPrms

    cons = "Select * from Parametro " & _
                " Where ParNombre In ( 'tcasientosvariosventas', 'PathApp' )"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsAux.EOF
        
        Select Case LCase(Trim(rsAux!ParNombre))
            Case "tcasientosvariosventas": prmTCVentasME = rsAux!ParValor
            
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto) & "\"
        End Select
        
        rsAux.MoveNext
    Loop
    
    rsAux.Close
    Exit Function
    
errPrms:
    clsGeneral.OcurrioError "Error al cargar los parámetos locales.", Err.Description
End Function
