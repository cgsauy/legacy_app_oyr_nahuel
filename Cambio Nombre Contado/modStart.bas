Attribute VB_Name = "modStart"
Option Explicit

'IMPRESION
Public paINContadoB As Integer
Public paINContadoN As String
Public paPrintEsXDefNC As Boolean

Public paIContadoB As Integer
Public paIContadoN As String
Public paPrintEsXDef As Boolean
'...........................................................

Public prmURLFirmaEFactura As String
Public EmpresaEmisora As clsClienteCFE
Public TasaBasica As Currency, TasaMinima As Currency
Public prmImporteConInfoCliente As Double
Public prmEFacturaProductivo As String

Public Type tRenglonFact
    IDArticulo As Long
    CodArticulo As Long
    IDCombo As Long
    ArtCombo As Long
    Tipo As Long
    Precio As Currency
    PrecioOriginal As Currency
    PrecioBonificacion As Currency
    EsInhabilitado As Boolean
End Type

Public Enum TipoArticulo
    Articulo = 1
    Servicio = 2
    Presupuesto = 3
    PagoFlete = 4
    Bonificacion = 5
End Enum


Public gPathListados As String, paBD As String

Public paLocalZF As Long, paLocalPuerto As Long
Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public paTipoComCheque As Long
Public tTiposArtsServicio As String     'Trama con todos los tipos que pertenecen a Servicio

Public Sub Main()

Dim sConnect As String

    On Error GoTo errMain
    Screen.MousePointer = 11
    
    Dim mAppVer As String
    mAppVer = App.title
    If miConexion.AccesoAlMenu(mAppVer) Then
    
        If mAppVer <> "" And mAppVer <> App.title And Not (App.Major & "." & App.Minor & "." & Format(App.Revision, "00")) >= mAppVer Then
            MsgBox "La versión del programa no es la última disponible." & vbCr & _
                        "Ud. debe actualizar el software.", vbExclamation, "Actualizar a Versión " & mAppVer
            End
        End If
    
        sConnect = miConexion.TextoConexion(logComercio)
        If Not InicioConexionBD(sConnect) Then
            End
            Screen.MousePointer = 0
            Exit Sub
        End If
        paBD = miConexion.RetornoPropiedad(bDB:=True)
                        
    
        If Not loc_GetInfoSucursal Then
                Screen.MousePointer = 0
                End: Exit Sub
        End If
        ChDir App.Path
        ChDir ("..")
        ChDir (CurDir & "\REPORTES\")
        gPathListados = CurDir & "\"
            
        CargoParametrosComercio
        loc_CargoParametroContado
        
        prj_LoadConfigPrint
        
        CargoValoresIVA
        CargoTiposDeArticulosServicios
        Set EmpresaEmisora = New clsClienteCFE
        EmpresaEmisora.CargoInformacionCliente cBase, 1, False
    
        frmChangeNameDoc.Show vbModeless
    
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
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.title & Chr(13) & "Error: " & Trim(Err.Description), vbCritical, "ATENCIÓN"
    End
End Sub
Public Sub BuscoComentariosAlerta(idCliente As Long, _
                                                   Optional Alerta As Boolean = False, Optional Cuota As Boolean = False, _
                                                   Optional Decision As Boolean = False, Optional Informacion As Boolean = False)
                                                   
Dim RsCom As rdoResultset
Dim aCom As String
Dim sHay As Boolean

    On Error GoTo errMenu
    Screen.MousePointer = 11
    sHay = False
    'Armo el str con los comentarios a consultar-------------------------------------------------
    If Not Alerta And Not Cuota And Not Decision And Not Informacion Then Exit Sub
    aCom = ""
    If Alerta Then aCom = aCom & Accion.Alerta & ", "
    If Cuota Then aCom = aCom & Accion.Cuota & ", "
    If Decision Then aCom = aCom & Accion.Decision & ", "
    If Informacion Then aCom = aCom & Accion.Informacion & ", "
    aCom = Mid(aCom, 1, Len(aCom) - 2)
    '---------------------------------------------------------------------------------------------------
    
    Cons = "Select * From Comentario, TipoComentario " _
            & " Where ComCliente = " & idCliente _
            & " And ComTipo = TCoCodigo " _
            & " And TCoAccion IN (" & aCom & ")"
    Set RsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsCom.EOF Then sHay = True
    RsCom.Close
    
    If Not sHay Then Screen.MousePointer = 0: Exit Sub
    
    Dim aObj As New clsCliente
    aObj.Comentarios idCliente:=idCliente
    DoEvents
    Set aObj = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
errMenu:
    clsGeneral.OcurrioError "Ocurrió un error al acceder al fomulario de comentarios.", Err.Description
    Screen.MousePointer = 0
End Sub

Public Sub prj_LoadConfigPrint(Optional bShowFrm As Boolean)
On Error GoTo errLCP
Dim objPrint As New clsCnfgPrintDocument
Dim sPCt As String, sPNo As String
Dim vPrint() As String

    With objPrint
        Set .Connect = cBase
        .Terminal = paCodigoDeTerminal
        If bShowFrm Then
            .CnfgTipoDocumento = "1"
            .ShowConfig
        End If
        sPCt = .getDocumentoImpresora(Contado)
        sPNo = .getDocumentoImpresora(NotaDevolucion)
    End With
    Set objPrint = Nothing
    
    If sPNo <> "" Then
        vPrint = Split(sPNo, "|")
        paINContadoN = vPrint(0)
        paINContadoB = vPrint(1)
        paPrintEsXDefNC = (Val(vPrint(3)) = 1)
    Else
        paINContadoN = ""
        MsgBox "No se encontró una configuración de impresión para el tipo de documento Nota Contado." & vbCr & "Comuniquese con el administrador para solucionar este problema.", vbCritical, "Impresora"
        End
    End If
    
    If sPCt <> "" Then
        vPrint = Split(sPCt, "|")
        paIContadoN = vPrint(0)
        paIContadoB = vPrint(1)
        paPrintEsXDef = (Val(vPrint(3)) = 1)
    Else
        paIContadoN = ""
        MsgBox "No existe una configuración de impresión para el documento y su sucursal." & vbCr & "Comuniquese con el administrador para solucionar este problema.", vbCritical, "Impresora"
        End
    End If
    Exit Sub
errLCP:
    MsgBox "Error al leer los datos de configuración: " & Err.Description, vbExclamation, "ATENCIÓN"
End Sub


Private Function loc_GetInfoSucursal() As Boolean

    loc_GetInfoSucursal = False
    
    paCodigoDeSucursal = 0
    paCodigoDeTerminal = 0
    paDContado = ""
    paDNDevolucion = ""
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & miConexion.NombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDeTerminal = RsAux!TerCodigo
        paCodigoDGI = RsAux("SucCodDGI")
        If Not IsNull(RsAux!SucDisponibilidad) Then paDisponibilidad = RsAux!SucDisponibilidad Else paDisponibilidad = 0
        
        'El documento que necesito solo es el contado.
        If Not IsNull(RsAux!SucDContado) Then paDContado = Trim(RsAux!SucDContado)
        If Not IsNull(RsAux!SucDNDevolucion) Then paDNDevolucion = Trim(RsAux!SucDNDevolucion)
    End If
    RsAux.Close
  
    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(miConexion.NombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
                & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
        Exit Function
    Else
        loc_GetInfoSucursal = True
    End If
    '-------------------------------------------------------------------------------------------------------------------------
    
End Function

Private Sub loc_CargoParametroContado()
    Cons = "Select * from Parametro Where ParNombre IN ('eFacturaActiva', 'URLFirmaEFactura', 'efactImporteDatosCliente')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux("ParNombre")))
            Case LCase("efactImporteDatosCliente"): prmImporteConInfoCliente = RsAux("ParValor")
            Case LCase("URLFirmaEFactura"): prmURLFirmaEFactura = Trim(RsAux("ParTexto"))
            Case LCase("eFacturaActiva"): prmEFacturaProductivo = RsAux("ParValor")
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
End Sub

Private Sub CargoValoresIVA()
Dim RsIva As rdoResultset
Dim sQy As String
    sQy = "SELECT IvaCodigo, IvaPorcentaje FROM TipoIva WHERE IvaCodigo IN (1,2)"
    Set RsIva = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsIva.EOF
        Select Case RsIva("IvaCodigo")
            Case 1: TasaBasica = RsIva("IvaPorcentaje")
            Case 2: TasaMinima = RsIva("IvaPorcentaje")
        End Select
        RsIva.MoveNext
    Loop
    RsIva.Close
End Sub

Private Sub CargoTiposDeArticulosServicios()
Dim sQy As String
    tTiposArtsServicio = ""
    sQy = "SELECT TipID FROM dbo.InTipos(" & paTipoArticuloServicio & ")"
    If ObtenerResultSet(cBase, RsAux, sQy, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
    Do While Not RsAux.EOF
        tTiposArtsServicio = tTiposArtsServicio & IIf(tTiposArtsServicio <> "", ",", "") & Trim(RsAux("TipID"))
        RsAux.MoveNext
    Loop
    RsAux.Close
End Sub


Public Function EsTipoDeServicio(ByVal idTipo As Long) As Boolean
     EsTipoDeServicio = (InStr(1, "," & tTiposArtsServicio & ",", "," & idTipo & ",") > 0)
End Function

