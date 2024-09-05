Attribute VB_Name = "modStart"
Option Explicit

'MODULO Conección
'Contiene rutinas y variables del entorno RDO.


'Definición del entorno RDO
Public cBase As rdoConnection       'Conexion a la Base de Datos
Public eBase As rdoEnvironment     'Definicion de entorno
Public RsAux As rdoResultset         'Resultset Auxiliar

'String.
Public Cons As String
Public paCodigoDeUsuario As Long
Public paCodigoDeSucursal As Long
Public paCodigoDeTerminal As Long

Public Const cnfgKeyTicketMovimientoCaja As String = "TickeadoraMovimientosDeCaja"
Public Const cnfgAppNombreMovimientoCaja As String = "MovimientosDeCaja"

Public Const cnfgKeyTicketConformes As String = "TickeadoraConformes"
Public Const cnfgAppNombreConformes As String = "Solicitudes Resueltas"
Public oCnfgPrint As New clsCnfgImpresora

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public prmPathListados As String
Public prmPathApp As String
Public prmLocal3Vias As String
Public paBD As String

Public prmNombreSucursal As String

Public paLocalPuerto As Long, paLocalZF As Long

'Variable Para configuracion de Impresoras  ------------------------------------------------------------------
Public paIContadoN As String
Public paIContadoB As Integer
Public paPrintCtdoPaperSize As Integer

Public paICreditoN As String
Public paICreditoB As Integer

Public paIReciboN As String
Public paIReciboB As Integer

Public paIConformeN As String
Public paIConformeB As Integer
Public paIConformeP As Integer

Public paIRemitoN As String
Public paIRemitoB As Integer

Private paLastUpdate As String
Public paOptPrintSel As String      'El nombre de la opción seleccionada
Public paOptPrintList As String      'Los nombres de las opciones ingresadas están separadas x |


Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    Dim aTexto As String
    
    If Not miConexion.AccesoAlMenu("Reimprimir Documentos") Then
        MsgBox "Acceso denegado. " & vbCrLf & "Consulte a su administrador de Sistemas", vbExclamation, "Acceso Denegado"
        End
    End If
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    If Not ObtenerConexionBD(cBase, "comercio") Then End
    paBD = miConexion.RetornoPropiedad(bDB:=True)
    
    CargoParametrosLocal
    
    If Trim(Command()) <> "" Then
        aTexto = Trim(Command())
        frmReImpresion.prmIDDocumento = Val(aTexto)
    End If
    
    frmReImpresion.Show vbModeless
    
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocal()
On Error Resume Next
    prmPathListados = ""
    
    Cons = "Select * from Parametro Where ParNombre In ('pathapp', 'ArticuloPisoAgencia', 'imp_Local_3Vias')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "imp_local_3vias"
                If Not IsNull(RsAux("ParTexto")) Then prmLocal3Vias = Trim(RsAux("ParTexto"))
            Case "pathapp"
                    prmPathListados = Trim(RsAux!ParTexto)
                    prmPathApp = Trim(RsAux!ParTexto)
            
            Case "articulopisoagencia": paArticuloPisoAgencia = RsAux!ParValor
        End Select
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    'Dejo un string con el formato "1,2,3,..., 50"
    prmLocal3Vias = Replace(prmLocal3Vias, " ", "")
    
    
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
        
         'Nombre de Cada Documento--------------------------------------------------------------------------------
        If Not IsNull(RsAux!SucDContado) Then paDContado = Trim(RsAux!SucDContado)
        If Not IsNull(RsAux!SucDCredito) Then paDCredito = Trim(RsAux!SucDCredito)
        If Not IsNull(RsAux!SucDNDevolucion) Then paDNDevolucion = Trim(RsAux!SucDNDevolucion)
        If Not IsNull(RsAux!SucDNCredito) Then paDNCredito = Trim(RsAux!SucDNCredito)
        If Not IsNull(RsAux!SucDRecibo) Then paDRecibo = Trim(RsAux!SucDRecibo)
        If Not IsNull(RsAux!SucDNEspecial) Then paDNEspecial = Trim(RsAux!SucDNEspecial)
        If Not IsNull(RsAux!SucDNDebito) Then paDNDebito = Trim(RsAux!SucDNDebito)
        If Not IsNull(RsAux!SucDRemito) Then paDRemito = Trim(RsAux!SucDRemito)
        
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
            .CnfgTipoDocumento = TipoDocumento.Contado & "," & TipoDocumento.Credito & "," & TipoDocumento.ReciboDePago & "," & TipoDocumento.Remito
            .ShowConfig
        End If
        
        If paLastUpdate <> .FechaUltimoCambio Or paLastUpdate = "" Then
            mCContado = .getDocumentoImpresora(Contado)
            mCCredito = .getDocumentoImpresora(Credito)
            mCRemito = .getDocumentoImpresora(Remito)
            mCRecibo = .getDocumentoImpresora(ReciboDePago)
            
            paOptPrintSel = .GetOpcionActual
            paOptPrintList = .GetOpcionesPrinter
            paLastUpdate = .FechaUltimoCambio
            
            If mCContado = "" Or mCCredito = "" Or mCRemito = "" Or mCRecibo = "" Then
                MsgBox "Falta alguna de las configuraciones de impresoras." & vbCrLf & _
                            "Valide éstos datos antes de imprimir.", vbCritical, "Faltan Valores de Impresión"
            End If

        End If
    End With
    Set objPrint = Nothing
    
    
    If mCContado <> "" Then
        vPrint = Split(mCContado, "|")
        paIContadoN = vPrint(0)
        paIContadoB = vPrint(1)        'paPrintEsXDefNC = (Val(vPrint(3)) = 1)
        paPrintCtdoPaperSize = vPrint(2)
    End If
    
    If mCCredito <> "" Then
        vPrint = Split(mCCredito, "|")
        paICreditoN = vPrint(0)
        paICreditoB = vPrint(1)
    End If
    
    If mCRecibo <> "" Then
        vPrint = Split(mCRecibo, "|")
        paIReciboN = vPrint(0)
        paIReciboB = vPrint(1)
    End If
    
    If mCRemito <> "" Then
    
        vPrint = Split(mCRemito, "|")
        paIRemitoN = vPrint(0)
        paIRemitoB = vPrint(1)
        paIConformeP = 1 'vPrint(2)
        
        If UBound(vPrint) > 1 Then
            If IsNumeric(vPrint(2)) Then paIConformeP = vPrint(2)
        End If
        
        paIConformeN = paIRemitoN
        paIConformeB = paIRemitoB
        
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
