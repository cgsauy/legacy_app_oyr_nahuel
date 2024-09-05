Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public Const prmKeyApp = "Ultimas Cobranzas"
Public Const prmKeyDatos = "Estadísticas"

Public prmPathApp As String
Public prmMonedaPesos As Long

Private txtConexion As String

Public prmAccesoDatos As Boolean

Public Sub Main()

    On Error GoTo errMain
    Screen.MousePointer = 11
    
    If Not miConexion.AccesoAlMenu(prmKeyApp) Then
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        If paCodigoDeUsuario <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "Usuario sin Acceso"
        Screen.MousePointer = 0
        End
    End If
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    txtConexion = miConexion.TextoConexion("comercio")
    
    If Not InicioConexionBD(txtConexion) Then End
    
    prmAccesoDatos = miConexion.AccesoAlMenu(prmKeyDatos)
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
                       
    CargoParametrosSucursal
    CargoParametrosLocales
    
    frmCobranzas.Show vbModeless
    
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & vbCrLf & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocales()
On Error GoTo errCP

    cons = "Select * from Parametro " & _
            " Where ParNombre IN ( 'pathapp', 'monedapesos')"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case (Trim(LCase(rsAux!ParNombre)))
            
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto) & "\"
            
            Case "monedapesos": prmMonedaPesos = rsAux!ParValor
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    Exit Sub
errCP:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub


Public Function EndMain()
    On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Function
