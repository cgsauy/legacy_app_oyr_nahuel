Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public Const prmKeyApp = "CompraME"

Public prmPathApp As String, prmPathListados As String
Public prmTipoTC As Integer
Public prmMCCompraME As Long

Private txtConexion As String
Private prmE_IDDoc As Long
Private prmE_IDMonedaV As Integer
Private prmE_TC As Currency
Private prmE_ImporteV As Currency
Private prmE_Coms As String

'Variables públicas para la impresión de recibos
Public paIReciboB As Integer
Public paIReciboN As String

Public paLastUpdate As String
Public paOptPrintSel As String      'El nombre de la opción seleccionada
Public paOptPrintList As String      'Los nombres de las opciones ingresadas están separadas x |

Public Sub Main()

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
    
    'CargoParametrosImpresion
    prj_LoadConfigPrint
    
    frmMain.prmIDDocumento = prmE_IDDoc
    frmMain.prmIDMonedaV = prmE_IDMonedaV
    frmMain.prmImporteV = prmE_ImporteV
    frmMain.prmTC = prmE_TC
    frmMain.prmComs = prmE_Coms
    
    frmMain.Show vbModeless
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.title & vbCrLf & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocales()
On Error GoTo errCP

    cons = "Select * from Parametro " & _
            " Where ParNombre IN ( 'pathapp', 'MonedaPesos', 'MonedaDolar', 'TipoTCCompraME', 'MCCompraMonedaE')"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        Select Case (Trim(LCase(rsAux!ParNombre)))
            
            Case "pathapp":
                    prmPathApp = Trim(rsAux!ParTexto)
                    prmPathListados = Mid(prmPathApp, 1, InStrRev(prmPathApp, "\"))
                    prmPathListados = prmPathListados & "Reportes\"
                    prmPathApp = prmPathApp & "\"
            
            Case "monedapesos": paMonedaPesos = rsAux!ParValor
            Case "monedadolar": paMonedaDolar = rsAux!ParValor
            
            Case "tipotccomprame": prmTipoTC = rsAux!ParValor
            Case "mccompramonedae": prmMCCompraME = rsAux!ParValor
            
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
    'D Id_Doc (recibo o ctdo)
    'MV id_Moneda (moneda de venta)
    'TC valorTC (Valor de la tasa de cambio)
    'IV importe     (Importe de la venta)
    'CD comentarios de documentos
    prmE_IDDoc = 0
    
    Dim mPrms As String
    mPrms = Trim(Command())
    If Trim(mPrms) = "" Then Exit Function
    
    Dim I As Integer
    Dim arrPrms() As String, arrValues() As String
    arrPrms = Split(Trim(mPrms), "|")
    
    For I = LBound(arrPrms) To UBound(arrPrms)
        arrValues = Split(arrPrms(I), " ")
        Select Case UCase(arrValues(0))
            
            Case "D": prmE_IDDoc = Val(arrValues(1))
            
            Case "MV": prmE_IDMonedaV = Val(arrValues(1))
            Case "TC": prmE_TC = Val(arrValues(1))
            Case "IV": prmE_ImporteV = Val(arrValues(1))
            
            Case "CD": prmE_Coms = Trim(arrValues(1))
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


Public Sub prj_LoadConfigPrint(Optional bShowFrm As Boolean)
On Error GoTo errLCP

Dim objPrint As New clsCnfgPrintDocument
Dim mCReciboPago As String
Dim vPrint() As String

    With objPrint
        Set .Connect = cBase
        .Terminal = paCodigoDeTerminal
        If bShowFrm Then
            .CnfgTipoDocumento = TipoDocumento.ReciboDePago
            .ShowConfig
        End If
        
         If paLastUpdate <> .FechaUltimoCambio Or paLastUpdate = "" Then
            mCReciboPago = .getDocumentoImpresora(ReciboDePago)
            
            paOptPrintSel = .GetOpcionActual
            paOptPrintList = .GetOpcionesPrinter
            paLastUpdate = .FechaUltimoCambio
            
            If mCReciboPago = "" Then
                MsgBox "Falta alguna de las configuraciones de impresoras." & vbCrLf & _
                            "Valide éstos datos antes de imprimir.", vbCritical, "Faltan Valores de Impresión"
            End If
        
        End If
    End With
    Set objPrint = Nothing
    
    If mCReciboPago <> "" Then
        vPrint = Split(mCReciboPago, "|")
        paIReciboN = vPrint(0)
        paIReciboB = vPrint(1)
    End If
    
    Exit Sub
errLCP:
    MsgBox "Error al leer los datos de configuración: " & Err.Description, vbExclamation, "ATENCIÓN"
End Sub

Public Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer
    For Each X In Printers
        If Trim(X.DeviceName) = Trim(DeviceName) Then
            Set Printer = X
            Exit For
        End If
    Next
End Sub

Public Function ChangeCnfgPrint() As Boolean
    Dim objPrint As New clsCnfgPrintDocument
    ChangeCnfgPrint = (paLastUpdate <> objPrint.FechaUltimoCambio)
    Set objPrint = Nothing
End Function

