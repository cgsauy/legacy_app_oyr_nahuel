Attribute VB_Name = "modStart"
Option Explicit

Public Const cnfgKeyTicketCopiaFactura As String = "TickeadoraCopiaFactura"
Public Const cnfgAppNombreCopia As String = "Copia de factura"
Public oCnfgPrint As New clsCnfgImpresora

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public prmPathListados As String
Public paArticuloPisoAgencia As Long
Public paBD As String

'Variable Para configuracion de Impresoras  ------------------------------------------------------------------
Public paIConformeN As String
Public paIConformeB As Integer
Public paIConformePS As Integer 'paper size.

Public paLastUpdate As String
Public paOptPrintSel As String      'El nombre de la opción seleccionada
Public paOptPrintList As String      'Los nombres de las opciones ingresadas están separadas x |

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    Dim aTexto As String
    
    If Not miConexion.AccesoAlMenu("Copia de Facturas") Then
        MsgBox "Acceso denegado. " & vbCrLf & "Consulte a su administrador de Sistemas", vbExclamation, "Acceso Denegado"
        End
    End If
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    
    If Not InicioConexionBD(miConexion.TextoConexion("comercio")) Then End
    
    paBD = miConexion.RetornoPropiedad(bDB:=True)
    
    CargoParametrosLocal
    
    If Trim(Command()) <> "" Then
        aTexto = Trim(Command())
        frmCopia.prmIDDocumento = Val(aTexto)
    End If
    
    frmCopia.Show vbModeless
    
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Private Sub CargoParametrosLocal()
On Error Resume Next
    prmPathListados = ""
    
    Cons = "Select * from Parametro Where ParNombre In ('pathapp', 'ArticuloPisoAgencia')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            
            Case "pathapp": prmPathListados = Trim(RsAux!ParTexto)
            
            Case "articulopisoagencia": paArticuloPisoAgencia = RsAux!ParValor
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
    Cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & Cons & "'" _
            & " And TerSucursal = SucCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!TerSucursal) Then paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDeTerminal = RsAux!TerCodigo
        
         'Nombre de Cada Documento--------------------------------------------------------------------------------
        If Not IsNull(RsAux!SucDContado) Then paDContado = Trim(RsAux!SucDContado)
        If Not IsNull(RsAux!SucDCredito) Then paDCredito = Trim(RsAux!SucDCredito)
        If Not IsNull(RsAux!SucDNDevolucion) Then paDNDevolucion = Trim(RsAux!SucDNDevolucion)
        If Not IsNull(RsAux!SucDNCredito) Then paDNCredito = Trim(RsAux!SucDNCredito)
        If Not IsNull(RsAux!SucDRecibo) Then paDRecibo = Trim(RsAux!SucDRecibo)
        If Not IsNull(RsAux!SucDNEspecial) Then paDNEspecial = Trim(RsAux!SucDNEspecial)
    End If
    RsAux.Close
    
    
    'CargoParametrosImpresion paCodigoDeSucursal, True, True, False, False, False, False, False, False, False
    
    prj_LoadConfigPrint
    
End Sub

Public Sub prj_LoadConfigPrint(Optional bShowFrm As Boolean)
On Error GoTo errLCP

Dim objPrint As New clsCnfgPrintDocument
Dim mCConformes As String
Dim vPrint() As String

    With objPrint
        Set .Connect = cBase
        .Terminal = paCodigoDeTerminal
        If bShowFrm Then
            .CnfgTipoDocumento = TipoDocumento.Remito
            .ShowConfig
        End If
        
         If paLastUpdate <> .FechaUltimoCambio Or paLastUpdate = "" Then
            mCConformes = .getDocumentoImpresora(Remito)
            
            paOptPrintSel = .GetOpcionActual
            paOptPrintList = .GetOpcionesPrinter
            paLastUpdate = .FechaUltimoCambio
            
            If mCConformes = "" Then
                MsgBox "Falta alguna de las configuraciones de impresoras." & vbCrLf & _
                            "Valide éstos datos antes de imprimir.", vbCritical, "Faltan Valores de Impresión"
            End If
        
        End If
    End With
    Set objPrint = Nothing
    
    If mCConformes <> "" Then
        vPrint = Split(mCConformes, "|")
        paIConformeN = vPrint(0)
        paIConformeB = vPrint(1)
        paIConformePS = vPrint(2)
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

