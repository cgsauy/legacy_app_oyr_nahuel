Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    
    If Not miConexion.AccesoAlMenu("Impresion de Cartas") Then End
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    If Not InicioConexionBD(miConexion.TextoConexion("comercio")) Then End
    
    'CargoParametrosLocal
            
    frmImCarta.Show vbModeless
    
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocal()

    cons = "Select * from Parametro"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case LCase(Trim(rsAux!ParNombre))
            
            'Case "tipocuotacontado": paTipoCuotaContado = rsAux!ParValor
            'Case "monedapesos": paMonedaPesos = rsAux!ParValor
            
            'Case "tcactualizarprecios": paTipoTC = rsAux!ParValor
        End Select
        
        rsAux.MoveNext
    Loop
    rsAux.Close
End Sub

Public Function EndMain()
On Error Resume Next

    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    
    End
End Function
