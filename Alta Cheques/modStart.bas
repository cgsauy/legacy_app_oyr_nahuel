Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public Const prmKeyApp = "Ingreso de Cheques"
Public Const prmSucesoCheque = 8

Public paMCChequeDiferido As Long

Private txtConexion As String

Dim prmE_TipoCheque As Integer
Dim prmE_IDDoc As Long
Dim prmE_ImporteCheque As Currency
Dim prmE_IDTagLiquidacion As Integer

Public Sub Main()

Dim aSucursal As String

    On Error GoTo errMain
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
    frmCheque.prmTipo = prmE_TipoCheque
    frmCheque.prmIDRecCtdo = prmE_IDDoc
    frmCheque.prmImporte = prmE_ImporteCheque
    frmCheque.prmTAG_LiqCamion = prmE_IDTagLiquidacion
    frmCheque.Show vbModeless
    
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & vbCrLf & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocales()
On Error GoTo errCP

    Cons = "Select * from Parametro " & _
            " Where ParNombre IN ( 'monedapesos', 'MCChequeDiferido')"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case (Trim(LCase(RsAux!ParNombre)))
            
            Case "mcchequediferido": paMCChequeDiferido = RsAux!ParValor
            
            Case "monedapesos": paMonedaPesos = RsAux!ParValor
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Exit Sub
errCP:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub

Private Function CargoParametrosEntrada()
    On Error GoTo errCPE
    'T TipoCheque (0, 1) |D Id_Doc (recibo o ctdo)  |V importe del cheque
    '7/11/2011 Agrego invocación desde liquidación de camioneros.
    'recibo: L IDTag
    prmE_TipoCheque = -1: prmE_IDDoc = 0
    
    Dim mPrms As String
    mPrms = Trim(Command())
    If Trim(mPrms) = "" Then Exit Function
    
    Dim i As Integer
    Dim arrPrms() As String, arrValues() As String
    arrPrms = Split(Trim(mPrms), "|")
    
    For i = LBound(arrPrms) To UBound(arrPrms)
        
        arrValues = Split(arrPrms(i), " ")
        Select Case UCase(arrValues(0))
            
            Case "T": prmE_TipoCheque = Val(arrValues(1))
                        
            Case "D": prmE_IDDoc = Val(arrValues(1))
            
            Case "V":
                If IsNumeric(arrValues(1)) Then prmE_ImporteCheque = arrValues(1)
                
            Case "L"
                If IsNumeric(arrValues(1)) Then prmE_IDTagLiquidacion = arrValues(1)
            
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
