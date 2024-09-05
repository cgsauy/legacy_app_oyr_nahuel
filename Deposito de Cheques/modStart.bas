Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public txtConexion As String

Public prmPathApp As String
Public paMDPagoDeCompra As Long

Public Sub Main()

Dim aSucursal As String

    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        txtConexion = miConexion.TextoConexion("comercio")
        InicioConexionBD txtConexion
        
        CargoParametrosLocales
        aSucursal = CargoParametrosSucursal
        
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        frmDepositos.Show vbModeless
    
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

Private Sub CargoParametrosLocales()
On Error GoTo errCP

    cons = "Select * from Parametro " & _
            " Where ParNombre IN ( 'pathapp', 'monedapesos', 'monedadolar', 'mdpagodecompra')"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case (Trim(LCase(rsAux!ParNombre)))
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto) & "\"
            
            Case "monedapesos": paMonedaPesos = rsAux!ParValor
            Case "monedadolar": paMonedaDolar = rsAux!ParValor
            
            Case "mdpagodecompra": paMDPagoDeCompra = rsAux!ParValor
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    Exit Sub
errCP:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub

