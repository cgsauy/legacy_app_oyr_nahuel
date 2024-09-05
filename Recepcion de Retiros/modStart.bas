Attribute VB_Name = "modStart"
Option Explicit

Public Enum TipoCliente
    Cliente = 1
    Empresa = 2
End Enum

Public paICartaB As Integer
Public paICartaN As String
Public paPrimeraHoraEnvio As Long
Public paUltimaHoraEnvio As Long

Public miConexion As New clsConexion
Public clsGeneral As New clsLibGeneral
Public txtConexion As String
'---------------------------------------------------------------------------------------

Public Sub Main()
Dim aSucursal As String
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        txtConexion = miConexion.TextoConexion("comercio")
        InicioConexionBD txtConexion
        CargoParametrosSucursal
        CargoParametrosImpresionServicio paCodigoDeSucursal
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
'        If Not IsNull(RsAux!SucDisponibilidad) Then paDisponibilidad = RsAux!SucDisponibilidad Else paDisponibilidad = 0
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

Private Sub CargoParametrosImpresionServicio(Sucursal As Long)
On Error GoTo errImp
    
    paICartaN = "": paICartaB = -1
    
    Cons = "Select * From Sucursal Where SucCodigo = " & Sucursal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!SucICaNombre) Then          'Carta.
            paICartaN = Trim(RsAux!SucICaNombre)
            If Not IsNull(RsAux!SucIRmBandeja) Then paICartaB = RsAux!SucICaBandeja
        End If
       '------------------------------------------------------------------------------------------------------------------
    End If
    RsAux.Close
    Exit Sub
errImp:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar los parámetros de impresión. Informe del error a su administrador de base de datos.", Err.Description

End Sub

