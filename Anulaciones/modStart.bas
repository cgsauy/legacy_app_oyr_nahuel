Attribute VB_Name = "modStart"
Option Explicit

Public paFPagoAnulaDocumento As Byte
Public paFPagoAnulaDocumentoNombre As String

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public prmPathListados As String, prmPathApp As String
Public paArticuloPisoAgencia As Long
Public paBD As String
Public paDiasAnulacionRemito As Byte

Public paCodigoDGI As Long
Public TasaBasica As Currency, TasaMinima As Currency
Public prmURLFirmaEFactura As String
Public prmEFacturaProductivo As String
Public prmImporteConInfoCliente As Currency

Public paLocalZF As Long, paLocalPuerto As Long

Public Sub Main()
     
    On Error GoTo errMain
    Screen.MousePointer = 11
    Dim aTexto As String
    
    If Not miConexion.AccesoAlMenu("Anulaciones") Then
        MsgBox "Acceso denegado. " & vbCrLf & "Consulte a su administrador de Sistemas", vbExclamation, "Acceso Denegado"
        End
    End If
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    
    If Not InicioConexionBD(miConexion.TextoConexion("comercio")) Then End
    
    paBD = miConexion.RetornoPropiedad(bDB:=True)
    
    CargoParametrosSucursal
    CargoParametrosLocal
    CargarParametroEFactura
    CargoValoresIVA
        
    If Trim(Command()) <> "" Then
        aTexto = Trim(Command())
        frmAnulacion.prmIDDocumento = Val(aTexto)
    End If
    frmAnulacion.Show vbModeless
    
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
    
    paDiasAnulacionRemito = 0   'solo los del día
    
    Cons = "Select * from Parametro " 'Where ParNombre In ('pathapp', 'ArticuloPisoAgencia')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case LCase("FormaPagoEnvioDocAnulado")
                paFPagoAnulaDocumento = RsAux("ParValor")
                paFPagoAnulaDocumentoNombre = Trim(RsAux("ParTexto"))
            
            Case LCase("AnulacionDiasRemito"): paDiasAnulacionRemito = RsAux("ParValor")
            
            Case "pathapp":
                        prmPathApp = Trim(RsAux!ParTexto)
                        prmPathListados = Trim(RsAux!ParTexto)
            
            Case "articulopisoagencia": paArticuloPisoAgencia = RsAux!ParValor
            
            Case "tipoarticuloservicio": paTipoArticuloServicio = RsAux!ParValor
            
            Case "mcanulacion": paMCAnulacion = RsAux!ParValor
            Case "articulodiferenciaenvio": paArticuloDiferenciaEnvio = RsAux!ParValor
            
            Case "mcingresosoperativos": paMCIngresosOperativos = RsAux!ParValor
            
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
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
    'cons = miConexion.NombreTerminal
    'cons = "Select * from Terminal Where TerNombre = '" & cons & "'"
    'Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    'If Not rsAux.EOF Then If Not IsNull(rsAux!TerSucursal) Then paCodigoDeSucursal = rsAux!TerSucursal
    'rsAux.Close
    
    
    CargoParametrosImpresion paCodigoDeSucursal, False, False, True, False, False, False, False, False, False
    
End Sub


'-------------------------------------------------------------------------------------------------------
'   Carga un string con todos los articulos que corresponden a los fletes.
'   Se utiliza en aquellos formularios que no filtren los fletes
'-------------------------------------------------------------------------------------------------------
Public Function CargoArticulosDeFlete() As String

Dim Fletes As String
    On Error GoTo errCargar
    Fletes = ""
    
    'Cargo los articulos a descartar-----------------------------------------------------------
    Cons = "Select Distinct(TFlArticulo) from TipoFlete Where TFlArticulo <> Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        Fletes = Fletes & RsAux!TFlArticulo & ","
        RsAux.MoveNext
    Loop
    RsAux.Close
    Fletes = "," & Fletes & paArticuloPisoAgencia & "," & paArticuloDiferenciaEnvio & ","
    '----------------------------------------------------------------------------------------------
    CargoArticulosDeFlete = Fletes
    Exit Function
    
errCargar:
    CargoArticulosDeFlete = Fletes
End Function

Public Sub CargoValoresIVA()
Dim RsIva As rdoResultset
Dim sQy As String
    sQy = "SELECT IvaCodigo, IvaPorcentaje FROM TipoIva WHERE IvaCodigo IN (1,2)"
    'Set RsIva = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurValues)
    If ObtenerResultSet(cBase, RsIva, sQy, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
    Do While Not RsIva.EOF
        Select Case RsIva("IvaCodigo")
            Case 1: TasaBasica = RsIva("IvaPorcentaje")
            Case 2: TasaMinima = RsIva("IvaPorcentaje")
        End Select
        RsIva.MoveNext
    Loop
    RsIva.Close
End Sub

Public Sub CargarParametroEFactura()
    Cons = "Select * from Parametro Where ParNombre IN ('eFacturaActiva', 'URLFirmaEFactura', 'efactImporteDatosCliente')"
    If ObtenerResultSet(cBase, RsAux, Cons, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
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

