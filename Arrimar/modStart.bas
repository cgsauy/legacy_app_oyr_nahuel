Attribute VB_Name = "modStart"
Option Explicit

Public paArrimar As Byte
Public clsGeneral As New clsorCGSA
Public paEstadoArticuloEntrega As Integer
Public paTipoArticuloServicio As Integer
Public paSonidoTimbre As String
Public paSonidoMal As String
Public paSonidoOK As String

Public Sub Main()
Dim miConexion As clsConexion
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    Set miConexion = New clsConexion
    
    If Not miConexion.AccesoAlMenu(App.Title) Then
        Screen.MousePointer = 0
        MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        End
    Else
        'Si da error la conexión la misma despliega el msg de error
        If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then
            Screen.MousePointer = 0
            End: Exit Sub
        End If
        
        'Guardo el usuario logueado
        CargoDatosSucursal miConexion.NombreTerminal
        CargoParametros
        Screen.MousePointer = 0
        frmArrimar.Show
    End If
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description) & vbCr, vbCritical, "ATENCIÓN"
    End
End Sub

Private Function CargoParametros() As Boolean
'Controlo aquellos que son vitales si no los cargue finalizo la app.
On Error GoTo errCP
    
    'Parametros a cero--------------------------
    paEstadoArticuloEntrega = 0
    paTipoArticuloServicio = 0
    paArrimar = 2
    Cons = "Select * from Parametro Where ParNombre IN('estadoarticuloentrega', 'tipoarticuloservicio', 'dep_wav_ArrimarTimbre', 'dep_wav_arrimarok', 'dep_wav_arrimarmal', 'dep_Estado_Arrimar_" & paCodigoDeSucursal & "')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case "tipoarticuloservicio": paTipoArticuloServicio = RsAux!ParValor
            Case "dep_wav_arrimartimbre": If Not IsNull(RsAux("ParTexto")) Then paSonidoTimbre = Trim(RsAux("ParTexto"))
            Case LCase("dep_Estado_Arrimar_") & paCodigoDeSucursal: paArrimar = RsAux("ParValor")
            Case "dep_wav_arrimarmal": If Not IsNull(RsAux("ParTexto")) Then paSonidoMal = Trim(RsAux("ParTexto"))
            Case "dep_wav_arrimarok": If Not IsNull(RsAux("ParTexto")) Then paSonidoOK = Trim(RsAux("ParTexto"))
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If paArrimar = 2 Then
        MsgBox "Atención no existe el parámetro arrimar de su terminal se considerará como encendida dicha aplicación.", vbExclamation, "Atención"
        paArrimar = 1
    End If
    
    CargoParametros = (paEstadoArticuloEntrega > 0)
    If Not CargoParametros Then MsgBox "Los parámetros de Estado de stock no fueron leidos, no podrá continuar.", vbCritical, "Manejo de Stock"
    Exit Function
errCP:
     clsGeneral.OcurrioError "Error al leer los parámetros.", Err.Description
     CargoParametros = False
End Function

