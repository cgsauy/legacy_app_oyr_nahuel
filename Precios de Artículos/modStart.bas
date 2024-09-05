Attribute VB_Name = "modStart"
Option Explicit

'REGISTRO DE SUCESOS--------------------------------------
Public Enum TipoSuceso
    ModificacionDeMora = 1
    AnulacionDeDocumentos = 2
    ModificacionDePrecios = 3
    RecepcionDeTraslados = 4
    AnulacionDeEnvios = 5
    CambioCostoDeFlete = 6
    Direcciones = 7
    ChequesDiferidos = 8
    CambioCategoriaCliente = 9
    Reimpresiones = 10
    DiferenciaDeArticulos = 11
    CederProductoServicio = 12
    FacturaArticuloInhabilitado = 13
    Notas = 14
    FacturaPlanInhabilitado = 15
    FacturaCambioNombre = 16
    CambioTipoArticuloServicio = 17
    VariosStock = 98
    Varios = 99
End Enum
'--------------------------------------------------------------------


Public idArticulo As Long
Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public lArtID As Long
Public paTipoCuotaContado As Long, paMonedaFacturacion As Long

Public Sub Main()
    On Error GoTo errMain
    Screen.MousePointer = 11
    If miConexion.AccesoAlMenu(App.Title) Then
        If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then
            Screen.MousePointer = 0
            End: Exit Sub
        End If
        
        'Cargo parámetro Tipo de cuota Contado
        Cons = "Select * From Parametro Where ParNombre = 'TipoCuotaContado' or ParNombre = 'MonedaFacturacion'"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            Select Case LCase(Trim(RsAux!ParNombre))
                Case "tipocuotacontado": paTipoCuotaContado = RsAux!ParValor
                Case "monedafacturacion": paMonedaFacturacion = RsAux!ParValor
            End Select
            RsAux.MoveNext
        Loop
        RsAux.Close
        CargoParametrosSucursal
        If Trim(Command()) <> "" Then lArtID = Command() Else lArtID = 0
        frmPrecioArticulo.Show
    Else
        MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description), vbCritical, "ATENCIÓN"
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
        MsgBox "La terminal " & UCase(aNombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn), vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------------------------------
    
End Function


