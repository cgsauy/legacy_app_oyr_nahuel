Attribute VB_Name = "modStart"
Option Explicit

Public Enum TipoDocumento
    'Documentos Facturacion
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
    NotaDebito = 40
    
    'Documentos de Compras
    CompraContado = 11
    CompraCredito = 12
    CompraNotaDevolucion = 13
    CompraNotaCredito = 14
    CompraRemito = 15
    CompraCarta = 16
    CompraCarpeta = 17
    CompraRecibo = 18
    CompraReciboDePago = 19
    CompraSalidaCaja = 30       'Pedidos el 11/12 por carlos y juliana
    CompraEntradaCaja = 31
    
    'Otros
    Traslados = 20
    Envios = 21
    CambioEstadoMercaderia = 22
    
    IngresoMercaderiaEspecial = 24
    ArregloStock = 25
    Servicio = 26
    ServicioCambioEstado = 27
    Devolucion = 28
End Enum

Public clsGeneral As New clsorCGSA

Public prmPathListados As String

'Variable Para configuracion de Impresoras  ------------------------------------------------------------------
Public paIContadoN As String
Public paIContadoB As Integer
Public paPrintCtdoPaperSize As Integer

Private paLastUpdate As String


Public Sub Main()
Dim miConexion As clsConexion
    
    On Error GoTo errMain
    Screen.MousePointer = 11
    Set miConexion = New clsConexion
    
    If Not miConexion.AccesoAlMenu(App.Title) Then
        Screen.MousePointer = 0
        MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        End
    Else
        'Si da error la conexión la misma despliega el msg de error
        If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then
            Screen.MousePointer = 0
            End: Exit Sub
        End If
        
        paCodigoDeUsuario = miConexion.UsuarioLogueado(True)
        
        'Guardo el usuario logueado
        CargoDatosSucursal miConexion.NombreTerminal
        CargoParametros
        prj_LoadConfigPrint bShowFrm:=False
        
        
        'CargoParametros
        frmResguardo.Show
        Screen.MousePointer = 0
        
        
    End If
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.Title & Chr(13) & "Error: " & Trim(Err.Description) & vbCr, vbCritical, "ATENCIÓN"
    End
End Sub

Private Function CargoParametros() As Boolean
'Controlo aquellos que son vitales si no los cargue finalizo la app.
On Error GoTo errCP
    
    'Parametros a cero--------------------------
    Cons = "Select * from Parametro Where ParNombre IN('pathapp')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "pathapp"
                    prmPathListados = Trim(RsAux!ParTexto)
                    'prmPathApp = Trim(RsAux!ParTexto)
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    CargoParametros = (prmPathListados <> "")
    If Not CargoParametros Then
        MsgBox "No se cargó el parámetro 'PathApp' para obtener la ubicación del reporte.", vbCritical, "Resguardos"
    Else
        Cons = ""
        Dim aPos As Integer, aT2 As String
        aT2 = prmPathListados
        Do While InStr(aT2, "\") <> 0
            aPos = InStr(aT2, "\")
            Cons = Cons & Mid(aT2, 1, aPos)
            aT2 = Mid(aT2, aPos + 1)
        Loop
        prmPathListados = Cons & "Reportes\"
    End If
    Exit Function
errCP:
     clsGeneral.OcurrioError "Error al leer los parámetros.", Err.Description
     CargoParametros = False
End Function


Public Sub prj_LoadConfigPrint(Optional bShowFrm As Boolean)
On Error GoTo errLCP

Dim objPrint As New clsCnfgPrintDocument
Dim mCContado As String
Dim vPrint() As String

    With objPrint
        Set .Connect = cBase
        .Terminal = paCodigoDeTerminal
        If bShowFrm Then
            .CnfgTipoDocumento = TipoDocumento.Contado
            .ShowConfig
        End If
        
        If paLastUpdate <> .FechaUltimoCambio Or paLastUpdate = "" Then
            mCContado = .getDocumentoImpresora(Contado)
            paLastUpdate = .FechaUltimoCambio
            
            If mCContado = "" Then
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
    Exit Sub
    
errLCP:
    MsgBox "Error al leer los datos de configuración: " & Err.Description, vbExclamation, "ATENCIÓN"
End Sub

Public Function ChangeCnfgPrint() As Boolean
    Dim objPrint As New clsCnfgPrintDocument
    ChangeCnfgPrint = (paLastUpdate <> objPrint.FechaUltimoCambio)
    Set objPrint = Nothing
End Function


