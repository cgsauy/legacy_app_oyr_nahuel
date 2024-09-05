Attribute VB_Name = "modStart"
Option Explicit

Public paLocalZF As Long

Public paICartaB As Integer
Public paICartaN As String

Public aSucursal As String, txtConexion As String
Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public Sub Main()

    On Error GoTo errMain
    Screen.MousePointer = 11
    
    If miConexion.AccesoAlMenu(App.Title) Then
        txtConexion = miConexion.TextoConexion("Comercio")
        InicioConexionBD txtConexion
        CargoParametrosComercio
        aSucursal = CargoParametrosSucursal
        MeCargoParametrosImpresion (paCodigoDeSucursal)
        frmTrasladoEspecial.Show vbModeless
    Else
        MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Screen.MousePointer = 0
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub MeCargoParametrosImpresion(Sucursal As Long)
On Error GoTo errImp
    paICartaN = "": paICartaB = -1
    Cons = "Select * From Local Where LocCodigo = " & Sucursal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!LocICaNombre) Then          'Carta.
            paICartaN = Trim(RsAux!LocICaNombre)
            If Not IsNull(RsAux!LocICaBandejaFlex) Then paICartaB = RsAux!LocICaBandejaFlex
        End If
       '------------------------------------------------------------------------------------------------------------------
    End If
    RsAux.Close
    Exit Sub
errImp:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar los parámetros de impresión. Informe del error a su administrador de base de datos.", Err.Description
End Sub



