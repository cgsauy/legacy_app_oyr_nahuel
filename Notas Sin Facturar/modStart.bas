Attribute VB_Name = "modStart"
Option Explicit

Public prmURLFirmaEFactura As String
Public EmpresaEmisora As clsClienteCFE
Public TasaBasica As Currency, TasaMinima As Currency
Public prmImporteConInfoCliente As Double
Public prmEFacturaProductivo As String
Public prmArtInteresMora As Long

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public mSQL As String
Public prmPathListados As String
Public paBD As String

'Variable Para configuracion de Impresoras  ------------------------------------------------------------------
Public paIReciboB As Integer
Public paIReciboN As String

Private paLastUpdate As String
Public paOptPrintSel As String      'El nombre de la opción seleccionada
Public paOptPrintList As String      'Los nombres de las opciones ingresadas están separadas x |


Public Sub Main()
     
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
    
    
        If Not InicioConexionBD(miConexion.TextoConexion("comercio"), 30) Then End
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        paBD = miConexion.RetornoPropiedad(bDB:=True)
        
        CargoParametrosLocal
        loc_CargoParametrosSucursal
        CargoValoresIVA
        CargoParametroeFactura
        
        Set EmpresaEmisora = New clsClienteCFE
        EmpresaEmisora.CargoInformacionCliente cBase, 1, False
        
        frmControl.Show vbModeless
        
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then
            MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "Usuario sin Autorización"
        End If
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.title & Chr(13) & "Error: " & Trim(Err.Description)
    End
End Sub

Public Function EndMain()
    On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Function

Private Sub CargoParametrosLocal()
On Error Resume Next
    prmPathListados = ""
    
    Cons = "Select * from Parametro Where ParNombre In ('pathapp')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            
            Case "pathapp":
                        'prmPathApp = Trim(rsAux!ParTexto)
                        prmPathListados = Trim(RsAux!ParTexto)
    
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
        
End Sub

Public Function loc_CargoParametrosSucursal() As String

Dim aNombreTerminal As String

    loc_CargoParametrosSucursal = ""
    aNombreTerminal = miConexion.NombreTerminal
    
    paCodigoDeSucursal = 0
    paCodigoDeTerminal = 0
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & aNombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDeTerminal = RsAux!TerCodigo
        If Not IsNull(RsAux!SucDisponibilidad) Then paDisponibilidad = RsAux!SucDisponibilidad Else paDisponibilidad = 0
        loc_CargoParametrosSucursal = Trim(RsAux!SucAbreviacion)
                
        If Not IsNull(RsAux!SucDNDebito) Then paDNDebito = Trim(RsAux!SucDNDebito)
    End If
    RsAux.Close
  
    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(aNombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
                & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
        End
        Exit Function
    End If
    '-------------------------------------------------------------------------------------------------------------------------
    
End Function

Public Sub prj_LoadConfigPrint(Optional bShowFrm As Boolean)
On Error GoTo errLCP

Dim objPrint As New clsCnfgPrintDocument
Dim mCRecibosP As String, mCConforme As String
Dim vPrint() As String

    With objPrint
        Set .Connect = cBase
        .Terminal = paCodigoDeTerminal
        If bShowFrm Then
            .CnfgTipoDocumento = TipoDocumento.ReciboDePago
            .ShowConfig
        End If
        
         If paLastUpdate <> .FechaUltimoCambio Or paLastUpdate = "" Then
            mCRecibosP = .getDocumentoImpresora(TipoDocumento.ReciboDePago)
            
            paOptPrintSel = .GetOpcionActual
            paOptPrintList = .GetOpcionesPrinter
            paLastUpdate = .FechaUltimoCambio
        
            If mCRecibosP = "" Then
                MsgBox "Falta alguna de las configuraciones de impresoras." & vbCrLf & _
                            "Valide éstos datos antes de imprimir.", vbCritical, "Faltan Valores de Impresión"
            End If
            
        End If
    End With
    Set objPrint = Nothing
    
    If mCRecibosP <> "" Then
        vPrint = Split(mCRecibosP, "|")
        paIReciboN = Trim(vPrint(0))
        paIReciboB = vPrint(1)        'paPrintEsXDefNC = (Val(vPrint(3)) = 1)
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

Public Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer
    For Each X In Printers
        If Trim(X.DeviceName) = Trim(DeviceName) Then
            Set Printer = X
            Exit For
        End If
    Next
End Sub

Private Sub CargoParametroeFactura()
    Cons = "Select * from Parametro Where ParNombre IN ('eFacturaActiva', 'URLFirmaEFactura', 'efactImporteDatosCliente', 'ArticuloInteresMora')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux("ParNombre")))
            Case LCase("efactImporteDatosCliente"): prmImporteConInfoCliente = RsAux("ParValor")
            Case LCase("URLFirmaEFactura"): prmURLFirmaEFactura = Trim(RsAux("ParTexto"))
            Case LCase("eFacturaActiva"): prmEFacturaProductivo = RsAux("ParValor")
            Case LCase("ArticuloInteresMora"): prmArtInteresMora = RsAux("ParValor")
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


