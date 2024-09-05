Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public prmPathListados As String
Public prmPathApp As String
Public paBD As String

Public prmNombreSucursal As String

Public paArticuloPisoAgencia As Long
Public paArticuloDiferenciaEnvio As Long
Public prmSucesoAnulacionDocs As Integer
Public Const prmSucesoReimpresiones = 10

'Variable Para configuracion de Impresoras  ------------------------------------------------------------------

Public paIRemitoN As String
Public paIRemitoB As Integer
Public paIRemitoP As Integer
Public paPrintEsXDefecto As Boolean

'Definicion de Tipos de Documentos----------------------
Public Enum TipoDocumento
    Contado = 1
    Credito = 2
    NotaDevolucion = 3
    NotaCredito = 4
    ReciboDePago = 5
    Remito = 6
    ContadoDomicilio = 7
    CreditoDomicilio = 8
    ServicioDomicilio = 9
    NotaEspecial = 10
End Enum

Public Sub Main()
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    Dim sPrm As String, vPrm() As String
    Dim iCont As Integer
    
    If Not miConexion.AccesoAlMenu("Remitos de Mercaderia") Then
        MsgBox "Acceso denegado. " & vbCrLf & _
                    "Consulte a su administrador de Sistemas", vbExclamation, "Acceso Denegado"
        End
    End If
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    If Not InicioConexionBD(miConexion.TextoConexion("comercio")) Then End
    paBD = miConexion.RetornoPropiedad(bDB:=True)
    
    'prmNombreSucursal = CargoParametrosSucursal
    CargoParametrosLocal
    
    '30-12-03 agregue esto para que acceda x ctdo.
    'Parámetros doc:#;id:#;etc
    sPrm = Trim(Command())
    If sPrm <> "" Then
        vPrm = Split(sPrm, ";", , vbTextCompare)
        For iCont = 0 To UBound(vPrm)
            'Hago x if ya que en sí serían 2
            If InStr(1, LCase(vPrm(iCont)), "doc:", vbTextCompare) > 0 Then
                vPrm(iCont) = Replace(vPrm(iCont), "doc:", "", , , vbTextCompare)
                If IsNumeric(vPrm(iCont)) Then frmRemito.prm_Documento = vPrm(iCont)
            ElseIf InStr(1, LCase(vPrm(iCont)), "id:", vbTextCompare) > 0 Then
                vPrm(iCont) = Replace(vPrm(iCont), "id:", "", , , vbTextCompare)
                If IsNumeric(vPrm(iCont)) Then frmRemito.prm_Remito = vPrm(iCont)
            End If
        Next
    End If
    frmRemito.Show vbModeless
    
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocal()
On Error Resume Next
    prmSucesoAnulacionDocs = 2
    prmPathListados = ""
    
    Cons = "Select * from Parametro Where ParNombre In ('pathapp', 'ArticuloPisoAgencia', 'ArticuloDiferenciaEnvio' )"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            
            Case "pathapp":
                                    prmPathListados = Trim(RsAux!ParTexto)
                                    prmPathApp = Trim(RsAux!ParTexto)
            
            Case "articulopisoagencia": paArticuloPisoAgencia = RsAux!ParValor
            Case LCase("ArticuloDiferenciaEnvio"): paArticuloDiferenciaEnvio = RsAux!ParValor
            
        End Select
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Cons = ""
    Dim aPos As Integer, aT2 As String
    aT2 = prmPathListados
    Do While InStr(aT2, "\") <> 0
        aPos = InStr(aT2, "\")
        Cons = Cons & Mid(aT2, 1, aPos)
        aT2 = Mid(aT2, aPos + 1)
    Loop
    prmPathListados = Cons & "Reportes\"
    
    
    'paCodigoDeSucursal
    Cons = miConexion.NombreTerminal
    Cons = "Select * from Terminal Left Outer Join Sucursal On TerSucursal = SucCodigo Where TerNombre = '" & Cons & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!TerSucursal) Then
            paCodigoDeSucursal = RsAux!TerSucursal
            prmNombreSucursal = Trim(RsAux!SucAbreviacion)
            paCodigoDeTerminal = RsAux!TerCodigo
        End If
    End If
    RsAux.Close
    
    
    prj_LoadConfigPrint bShowFrm:=False
    
End Sub

Public Sub prj_LoadConfigPrint(Optional bShowFrm As Boolean)
On Error GoTo errLCP

Dim objPrint As New clsCnfgPrintDocument
Dim mCContado As String, mCCredito As String, mCRecibo As String, mCRemito As String
Dim vPrint() As String

    With objPrint
        Set .Connect = cBase
        .Terminal = paCodigoDeTerminal
        If bShowFrm Then
            .CnfgTipoDocumento = TipoDocumento.Remito
            .ShowConfig
        End If
       
        mCRemito = .getDocumentoImpresora(Remito)
    End With
    Set objPrint = Nothing
    
    
    If mCRemito = "" Then
        MsgBox "Falta alguna de las configuraciones de impresoras." & vbCrLf & _
                    "Valide éstos datos antes de imprimir.", vbCritical, "Faltan Valores de Impresión"
    End If
    
    If mCRemito <> "" Then
        vPrint = Split(mCRemito, "|")
        paIRemitoN = vPrint(0)
        paIRemitoB = vPrint(1)
        paIRemitoP = Val(vPrint(2))
        paPrintEsXDefecto = Val(vPrint(3))
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

