Attribute VB_Name = "modStart"
Option Explicit
Public paLocalZF As Long
Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public paEstadoArticuloEntrega As Long

Public Sub Main()
Dim aSucursal As String

    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        InicioConexionBD miConexion.TextoConexion("Comercio")
        CargoDatosSucursal
        CargoParametros
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
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Public Sub CargoDatosSucursal()
Dim aNombreTerminal As String
    
    aNombreTerminal = miConexion.NombreTerminal
    paCodigoDeSucursal = 0
    paCodigoDeTerminal = 0
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & aNombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        paCodigoDeSucursal = rsAux!TerSucursal
        paCodigoDeTerminal = rsAux!TerCodigo
    End If
    rsAux.Close
    'If paCodigoDeSucursal = 0 Then
    '    MsgBox "La terminal " & UCase(aNombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn), vbExclamation, "ATENCIÓN"
    '    Exit Sub
    'End If
    '-------------------------------------------------------------------------------------------------------------------------
    
End Sub

Public Sub CargoParametros()
    paEstadoArticuloEntrega = 0
    cons = "Select * From Parametro Where ParNombre IN ('EstadoArticuloEntrega')"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        paEstadoArticuloEntrega = rsAux!ParValor
    End If
    rsAux.Close
End Sub
