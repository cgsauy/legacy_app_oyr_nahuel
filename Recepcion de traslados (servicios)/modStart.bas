Attribute VB_Name = "modStart"
Option Explicit


Public Enum TipoLocal
    Camion = 1
    Deposito = 2
End Enum

'Public paPrimeraHoraEnvio As Long
'Public paUltimaHoraEnvio As Long
Public paClienteEmpresa As Long
Public paEstadoARecuperar As Integer

Public paPrintCartaB As Integer
Public paPrintCartaD As String
Public paPrintCartaXDef As Boolean
Public paPrintCartaPaperSize As Integer


Public miConexion As New clsConexion
Public clsGeneral As New clsLibGeneral
Public txtConexion As String
'---------------------------------------------------------------------------------------

Public Sub Main()
Dim aSucursal As String
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        
        If Not InicioConexionBD(miConexion.TextoConexion("comercio")) Then
            End
        End If
        CargoParametrosSucursal
        CargoParametrosDelComercio
        prj_GetPrinter False
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
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
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description)
    End
End Sub

Public Function CargoParametrosSucursal() As String

Dim aNombreTerminal As String

    CargoParametrosSucursal = ""
    aNombreTerminal = miConexion.NombreTerminal
    
    paCodigoDeSucursal = 0
    paCodigoDeTerminal = 0
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & aNombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDeTerminal = RsAux!TerCodigo
        CargoParametrosSucursal = Trim(RsAux!SucAbreviacion)
    End If
    RsAux.Close
  
    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(aNombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
                & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
        End
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------------------------------
    
End Function

Private Sub CargoParametrosDelComercio()
    Cons = "Select * from Parametro Where ParNombre IN('clienteempresa', 'estadoarticuloentrega', 'estadoarecuperar')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "clienteempresa": paClienteEmpresa = RsAux!ParValor
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            Case "estadoarecuperar": paEstadoARecuperar = RsAux!ParValor
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
End Sub

Public Sub prj_GetPrinter(ByVal bShowP As Boolean)
On Error GoTo errImp
    paPrintCartaD = ""
    paPrintCartaB = -1
    Dim objP As New clslPrintConfig
    With objP
        If bShowP Then
            If Not .ShowPrinterSetup("21", paCodigoDeTerminal) Then
                GoTo errImp
            End If
        End If
        If .LoadPrinterConfig("21", paCodigoDeTerminal) Then
            .GetPrinterDoc 21, paPrintCartaD, paPrintCartaB, paPrintCartaXDef, paPrintCartaPaperSize
        End If
    End With
    If paPrintCartaD = "" Then MsgBox "Por favor verifique la configuración de impresión.", vbInformation, "Atención"
    
errImp:
    Set objP = Nothing
    Screen.MousePointer = 0
End Sub



