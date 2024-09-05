Attribute VB_Name = "modStart"
Option Explicit

Public paIConformeB As Integer
Public paIConformeN As String

Public aSucursal As String, txtConexion As String
Public miConexion As New clsConexion
Public clsGeneral As New clsLibGeneral

Public Sub Main()

    On Error GoTo errMain
    Screen.MousePointer = 11
    
    If miConexion.AccesoAlMenu(App.Title) Then
        txtConexion = miConexion.TextoConexion(logComercio)
        InicioConexionBD txtConexion
        CargoParametrosComercio
         aSucursal = CargoParametrosSucursal
         MeCargoParametrosImpresion (paCodigoDeSucursal)
        FichaAgencia.Show vbModeless
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
    paIConformeN = "": paIConformeB = -1
    Cons = "Select * From Local Where LocCodigo = " & Sucursal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!LocICnNombre) Then          'Conforme.
            paIConformeN = Trim(RsAux!LocICnNombre)
            If Not IsNull(RsAux!LocICnBandeja) Then paIConformeB = RsAux!LocICnBandeja
        End If
       '------------------------------------------------------------------------------------------------------------------
    End If
    RsAux.Close
    Exit Sub
errImp:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar los parámetros de impresión. Informe del error a su administrador de base de datos.", Err.Description
End Sub

