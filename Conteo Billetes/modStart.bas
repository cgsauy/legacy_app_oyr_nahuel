Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public Const prmKeyApp = "Visualizacion Caja"
Public prmPathApp As String

Private txtConexion As String
Private prmE_IDDisp As Long
Private prmE_Fecha As String

Public Sub Main()

    On Error GoTo errMain
    If App.PrevInstance Then End
    
    Screen.MousePointer = 11
    
    If Not miConexion.AccesoAlMenu(prmKeyApp) Then
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        If paCodigoDeUsuario <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "Usuario sin Acceso"
        Screen.MousePointer = 0
        End
    End If
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    txtConexion = miConexion.TextoConexion("comercio")
    
    If Not InicioConexionBD(txtConexion) Then End
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
                       
    CargoParametrosSucursal
    CargoParametrosLocales
    CargoParametrosEntrada
    
    frmMain.Show vbModeless
    frmMain.prmCargoDatos prmE_IDDisp, prmE_Fecha
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
            " Where ParNombre IN ( 'pathapp', 'monedapesos')"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case (Trim(LCase(rsAux!ParNombre)))
            
            Case "pathapp": prmPathApp = Trim(rsAux!ParTexto) & "\"
            
            Case "monedapesos": paMonedaPesos = rsAux!ParValor
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
    'D Id_Disponibilidad |F fecha
    prmE_IDDisp = 0: prmE_Fecha = ""
    
    Dim mPrms As String
    mPrms = Trim(Command())
    If Trim(mPrms) = "" Then Exit Function
    
    Dim I As Integer
    Dim arrPrms() As String, arrValues() As String
    arrPrms = Split(Trim(mPrms), "|")
    
    For I = LBound(arrPrms) To UBound(arrPrms)
        arrValues = Split(arrPrms(I), " ")
        
        Select Case UCase(arrValues(0))
            Case "D": prmE_IDDisp = Val(arrValues(1))
            Case "F": prmE_Fecha = Trim(arrValues(1))
        End Select
        
    Next
    
    Exit Function
errCPE:
    clsGeneral.OcurrioError "Error al cargar los parámetros de entrada: " & mPrms, Err.Description
End Function

Public Function EndMain()
    On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Function
