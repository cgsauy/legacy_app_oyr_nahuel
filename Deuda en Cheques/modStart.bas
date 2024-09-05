Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public txtConexion As String

Public prmPathApp As String

Public Sub Main()

    On Error GoTo errMain
    Screen.MousePointer = 11
    miConexion.AccesoAlMenu "Deuda en Cheques"
    txtConexion = miConexion.TextoConexion("comercio")
    InicioConexionBD txtConexion
        
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    
    CargoParametrosLocales
    frmDeudaCH.Show vbModeless
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocales()
On Error GoTo errCP

    cons = "Select * from Parametro " & _
            " Where ParNombre IN ( 'pathapp')"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case (Trim(LCase(rsAux!ParNombre)))
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto) & "\"
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    Exit Sub
errCP:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub



