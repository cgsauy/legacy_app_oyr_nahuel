Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public txtConexion As String

Public paArticuloDiferenciaEnvio As Long
Public paArticuloPisoAgencia As Long
Public gPathListados As String

Public prmPathApp As String

Public Sub Main()

    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        txtConexion = miConexion.TextoConexion("comercio")
        InicioConexionBD txtConexion
        
        CargoParametrosLocales
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        frmDeFactura.Show vbModeless
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocales()

    On Error GoTo errParametro
    
    ChDir App.Path
    ChDir ("..")
    ChDir (CurDir & "\REPORTES\")
    gPathListados = CurDir & "\"
    
    cons = "Select * from Parametro Where ParNombre like 'articulo%' OR ParNombre In ('PathApp')"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case LCase(Trim(rsAux!ParNombre))
            
            Case "articulodiferenciaenvio": paArticuloDiferenciaEnvio = rsAux!ParValor
            Case "articulopisoagencia": paArticuloPisoAgencia = rsAux!ParValor
            
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto) & "\"
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close

    Exit Sub
errParametro:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub


'-------------------------------------------------------------------------------------------------------
'   Carga un string con todos los articulos que corresponden a los fletes.
'   Se utiliza en aquellos formularios que no filtren los fletes
'-------------------------------------------------------------------------------------------------------
Public Function CargoArticulosDeFlete() As String

Dim Fletes As String
    On Error GoTo errCargar
    Fletes = ""
    
    'Cargo los articulos a descartar-----------------------------------------------------------
    cons = "Select Distinct(TFlArticulo) from TipoFlete Where TFlArticulo <> Null"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsAux.EOF
        Fletes = Fletes & rsAux!TFlArticulo & ","
        rsAux.MoveNext
    Loop
    rsAux.Close
    Fletes = Fletes & paArticuloPisoAgencia & "," & paArticuloDiferenciaEnvio & ","
    '----------------------------------------------------------------------------------------------
    CargoArticulosDeFlete = Fletes
    Exit Function
    
errCargar:
    CargoArticulosDeFlete = Fletes
End Function

