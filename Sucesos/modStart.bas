Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsLibGeneral
Public txtConexion As String

Public prmPathApp As String

Private prmE_FDesde As String
Private prmE_FHasta As String
Private prmE_TSuceso As Long

Public Sub Main()

Dim aSucursal As String

    On Error GoTo errMain
    Screen.MousePointer = 11
    
    If Not miConexion.AccesoAlMenu("Sucesos") Then
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        If paCodigoDeUsuario <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    txtConexion = miConexion.TextoConexion("comercio")
    
    If Not InicioConexionBD(txtConexion) Then End
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    
    CargoParametrosLocales
    CargoParametrosEntrada

    frmSuceso.Show vbModeless
    frmSuceso.gbl_Consulto prmE_FDesde, prmE_FHasta, prmE_TSuceso
    
    
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
            " Where ParNombre IN ( 'pathapp')"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case (Trim(LCase(rsAux!ParNombre)))
            
            Case "pathapp":
                    prmPathApp = Trim(rsAux!ParTexto)
                    prmPathApp = prmPathApp & "\"
            
        End Select
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    Exit Sub
errCP:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub


Private Function CargoParametrosEntrada()
    On Error GoTo errCPE
    'FD F_Desde | FH F_Hasta |S idSuceso
    prmE_FDesde = ""
    prmE_FHasta = ""
    prmE_TSuceso = 0

    Dim mPrms As String
    mPrms = Trim(Command())
    If Trim(mPrms) = "" Then Exit Function
    
    Dim I As Integer
    Dim arrPrms() As String, arrValues() As String
    arrPrms = Split(Trim(mPrms), "|")
    
    For I = LBound(arrPrms) To UBound(arrPrms)
        arrValues = Split(arrPrms(I), " ")
        Select Case UCase(arrValues(0))
            
            Case "FD": prmE_FDesde = Trim(arrValues(1))
            Case "FH": prmE_FHasta = Trim(arrValues(1))
            Case "S": prmE_TSuceso = Val(arrValues(1))
            
        End Select
        
    Next
    
    Exit Function
    
errCPE:
    clsGeneral.OcurrioError "Error al cargar los parámetros de entrada: " & mPrms, Err.Description
End Function



