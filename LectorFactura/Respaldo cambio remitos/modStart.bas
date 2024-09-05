Attribute VB_Name = "modStart"
Option Explicit

Public paWavLista As String
Public paWavOK As String

Public clsGeneral As New clsorCGSA
Public paEstadoArticuloEntrega As Integer
Public paTipoArticuloServicio As Integer
Public paCloseApp As String
Public paEntregaTotal As String
Public paEntregaParcial As String
Public paEntregaCancelar As String
Public paArrimar As Byte
Public paPagHtml As String

Public Sub Main()
Dim miConexion As clsConexion
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    Set miConexion = New clsConexion
    'Si da error la conexión la misma despliega el msg de error
    If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then
        Screen.MousePointer = 0
        End: Exit Sub
    End If
    
    'Guardo el usuario logueado
    CargoDatosSucursal miConexion.NombreTerminal
    CargoParametros
    Screen.MousePointer = 0
    frmFac.Show
    
    Set miConexion = Nothing
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

    Cons = "Select * from Parametro " & _
        "Where ParNombre IN('estadoarticuloentrega', 'tipoarticuloservicio', 'dep_CerrarLector', 'dep_Ent_Total', 'dep_Ent_Parcial', " & _
        "'dep_Ent_Cancelar', 'dep_ent_paghtml', 'dep_wav_ClienteLista', 'dep_wav_ClienteConfirmado', 'dep_Estado_Arrimar_" & paCodigoDeSucursal & "')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case "tipoarticuloservicio": paTipoArticuloServicio = RsAux!ParValor
            Case "dep_cerrarlector": If Not IsNull(RsAux("ParTexto")) Then paCloseApp = Trim(RsAux("ParTexto"))
            Case "dep_ent_total": If Not IsNull(RsAux("ParTexto")) Then paEntregaTotal = Trim(RsAux("ParTexto"))
            Case "dep_ent_parcial": If Not IsNull(RsAux("ParTexto")) Then paEntregaParcial = Trim(RsAux("ParTexto"))
            Case "dep_ent_cancelar": If Not IsNull(RsAux("ParTexto")) Then paEntregaCancelar = Trim(RsAux("ParTexto"))
            Case LCase("dep_Estado_Arrimar_") & paCodigoDeSucursal: paArrimar = RsAux("ParValor")
            Case "dep_ent_paghtml": If Not IsNull(RsAux("ParTexto")) Then paPagHtml = Trim(RsAux("ParTexto"))
            Case "dep_wav_clientelista": If Not IsNull(RsAux("ParTexto")) Then paWavLista = Trim(RsAux("ParTexto"))
            Case "dep_wav_clienteconfirmado": If Not IsNull(RsAux("ParTexto")) Then paWavOK = Trim(RsAux("ParTexto"))
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    If paEntregaTotal = "" Then paEntregaTotal = "S"
    If paEntregaParcial = "" Then paEntregaParcial = "N"
    If paEntregaCancelar = "" Then paEntregaCancelar = "C"
    
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

