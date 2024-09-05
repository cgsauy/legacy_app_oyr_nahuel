Attribute VB_Name = "modStart"
Option Explicit

Public Const CGSA_TickeadoraCredito As String = "FacturaCredito"

'Definición del entorno RDO
Public cBase As rdoConnection       'Conexion a la Base de Datos
Public eBase As rdoEnvironment     'Definicion de entorno
Public RsAux As rdoResultset         'Resultset Auxiliar

'String.
Public Cons As String
Public paCodigoDeUsuario As Long
Public paCodigoDeSucursal As Long
Public paCodigoDGI As Long
Public paCodigoDeTerminal As Long

Public prmURLFirmaEFactura As String
Public EmpresaEmisora As clsClienteCFE
Public TasaBasica As Currency, TasaMinima As Currency
Public prmImporteConInfoCliente As Double
Public prmEFacturaProductivo As String
'...........................................................

Public prmHUBURL As String

Public Const cnfgKeyTicketConformes As String = "TickeadoraConformes"
Public Const cnfgAppNombreConformes As String = "Solicitudes Resueltas"

Public oCnfgPrint As New clsCnfgImpresora

Public paLocalesCobraVencidas As String

Public Const FormatoCedula = "_.___.___-_"

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public gPathListados As String
Public paBD As String
Public prmPathApp As String

Public paDisponibilidad As Long

Public paCofis As Long
Public paMCChequeDiferido As Long
Public paMonedaPesos As Integer

Public prmSuc_ModificacionDePrecios As Integer
Public prmSuc_FacturaArticuloNOHabilitado As Integer
Public prmTipoArtSinCofis As String

Public paNombreSucursal As String

'Variable Para configuracion de Impresoras  ------------------------------------------------------------------
Public iJobCre As Integer
Public iJobCon As Integer        'iJobCre= Imp.Credito  - iJobCon= Imp.Conforme

Public paICreditoB As Integer
Public paICreditoN As String

Public paIConformeB As Integer
Public paIConformeN As String
Public paIConformeP As Integer

Private paLastUpdate As String
Public paOptPrintSel As String      'El nombre de la opción seleccionada
Public paOptPrintList As String      'Los nombres de las opciones ingresadas están separadas x |



'Comunicacion con el servidor de Asuntos Pendientes ------------------------------------------------------------------
Public prmIPServer As String
Public prmPortServer As Long

Public Const sc_FIN = vbCrLf

Public Enum Asuntos
    Solicitudes = 1
    Servicios = 2
    GastosAAutorizar = 3
    SucesosAAutorizar = 4
    SolicitudesResueltas = 5
End Enum

Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
'------------------------------------------------------------------------------------------------------------------------------------
Public tTiposArtsServicio As String     'Trama con todos los tipos que pertenecen a Servicio

Public Sub Main()

    On Error GoTo errMain
    
    If Not miConexion.AccesoAlMenu("Solicitudes Resueltas") Then
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then
            MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "Acceso Denegado"
        End If
        End
    End If
    
    Screen.MousePointer = 11
    
    oCnfgPrint.CargarConfiguracion App.title, cnfgKeyTicketConformes

    If Not ObtenerConexionBD(cBase, "comercio") Then End
        
    CargoParametrosLocal
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    
    Dim txtConexion As String
    txtConexion = miConexion.TextoConexion("comercio")
    paBD = PropiedadesConnect(txtConexion, True)
    
    ChDir App.Path: ChDir ("..")
    ChDir (CurDir & "\REPORTES\")
    gPathListados = CurDir & "\"
    
    dis_CargoArrayMonedas
    CargoParametrosSucursal
    CargoPrmsSignalR
    
    'Empresa para CFE
    Set EmpresaEmisora = New clsClienteCFE
    EmpresaEmisora.CargoClienteCarlosGutierrez paCodigoDeSucursal  '.CargoInformacionCliente cBase, 1, False
    CargoValoresIVA
    
    'Abro el Engine del Crystal
    If crAbroEngine = 0 Then MsgBox Trim(crMsgErr), vbCritical, "Engine Error"
    InicializoCrystalEngine
    
    Dim prmValor As Long
    prmValor = 0
    
    '1) Si viene parametro id de solicitud voy directo al formulario para facturarla    ----------------------------
    ' /ID=XXXX
    Dim mTexto As String, mParams() As String
    mTexto = Trim(Command())
    If Trim(mTexto) <> "" Then
        mParams = Split(mTexto, "=")
        Select Case UCase(mParams(0))
            Case "/ID": prmValor = mParams(1)
        End Select
    End If
    '------------------------------------------------------------------------------------------------------------------------
    
    CargoTiposDeArticulosServicios
    
    If prmValor = 0 Then
        frmLista.Show vbModeless
    Else
        fnc_ActivoCredito prmValor
    End If
        
    Screen.MousePointer = 0
    Exit Sub
    
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.title & vbCrLf & "Error: " & Trim(Err.Description)
    End
End Sub

Private Function fnc_ActivoCredito(mid_Solicitud As Long)

    Select Case fnc_BloqueoSolicitud(mid_Solicitud) 'Si es 1 esta todo OK--------------------------
        Case 0  'OTRO USUARIO
            Screen.MousePointer = 0
            MsgBox "La solicitud se está facturando por otro usuario. No podrá visualizarla.", vbExclamation, "Datos Modificados"
            GoTo etqExit
       
        Case -1 'ERROR o FUE RESUELTA
            Screen.MousePointer = 0
            MsgBox "Posiblemente la solicitud ya fue facturada.", vbExclamation, "Datos Modificados"
            GoTo etqExit
    End Select  '----------------------------------------------------------------------------------------
        
    Screen.MousePointer = 11
    
    frmCredito.prmIDSolicitud = mid_Solicitud
    frmCredito.Show vbModal
        
etqExit:
    EndMain
    
End Function


Public Function EndMain()

    On Error Resume Next
    
    crCierroTrabajo (iJobCre)        'Cierro los reportes de credito y conforme
    crCierroTrabajo (iJobCon)
    crCierroEngine                        'Cierro el Engine del Crystal
    
    cBase.Close
    
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Function

Public Sub prj_LoadConfigPrint(Optional bShowFrm As Boolean)
On Error GoTo errLCP

Dim objPrint As New clsCnfgPrintDocument
Dim mCCredito As String, mCConforme As String
Dim vPrint() As String

    With objPrint
        Set .Connect = cBase
        .Terminal = paCodigoDeTerminal
        If bShowFrm Then
            .CnfgTipoDocumento = TipoDocumento.Credito & "," & TipoDocumento.Remito
            .ShowConfig
        End If
        
        If paLastUpdate <> .FechaUltimoCambio Or paLastUpdate = "" Then
            mCCredito = .getDocumentoImpresora(Credito)
            mCConforme = .getDocumentoImpresora(Remito)
            
            paOptPrintSel = .GetOpcionActual
            paOptPrintList = .GetOpcionesPrinter
            paLastUpdate = .FechaUltimoCambio
        
            If mCConforme = "" Or mCCredito = "" Then
                MsgBox "Falta alguna de las configuraciones de impresoras." & vbCrLf & _
                            "Valide éstos datos antes de imprimir.", vbCritical, "Faltan Valores de Impresión"
            End If
        End If
        
    End With
    Set objPrint = Nothing
    
    If mCCredito <> "" Then
        vPrint = Split(mCCredito, "|")
        paICreditoN = Trim(vPrint(0))
        paICreditoB = vPrint(1)
    End If
    
    paIConformeP = 1
    If mCConforme <> "" Then
        vPrint = Split(mCConforme, "|")
        paIConformeN = Trim(vPrint(0))
        paIConformeB = vPrint(1)
        If UBound(vPrint) > 1 Then
            If IsNumeric(vPrint(2)) Then paIConformeP = vPrint(2)
        End If
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

Private Function InicializoCrystalEngine()
    
    'Inicializa el Engine del Crystal y setea la impresora para el JOB
    On Error GoTo ErrCrystal
        
    'Inicializo el Reporte Para el Credito-----------------------------------------------------------------------------------
    iJobCre = crAbroReporte(gPathListados & "Credito.RPT")
    If iJobCre = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
    If Trim(Printer.DeviceName) <> Trim(paICreditoN) Then SeteoImpresoraPorDefecto paICreditoN
    If Not crSeteoImpresora(iJobCre, Printer, paICreditoB) Then GoTo ErrCrystal
    '----------------------------------------------------------------------------------------------------------------------------
    
    'Inicializo el Reporte Para el Conforme---------------------------------------------------------------------------------
    iJobCon = crAbroReporte(gPathListados & "Conforme.RPT")
    If iJobCon = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
    If Trim(Printer.DeviceName) <> Trim(paIConformeN) Then SeteoImpresoraPorDefecto paIConformeN
    If Not crSeteoImpresora(iJobCon, Printer, paIConformeB, paIConformeP) Then GoTo ErrCrystal
    '----------------------------------------------------------------------------------------------------------------------------
    Exit Function

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError Trim(crMsgErr) & " No se podrán imprimir facturas."
End Function

Public Function ArmoNombre(Ape1 As String, Ape2 As String, Nom1 As String, Nom2 As String) As String

    ArmoNombre = Trim(Ape1) & " " & Trim(Ape2)
    ArmoNombre = Trim(ArmoNombre) & ", " & Trim(Nom1) & " " & Trim(Nom2)
    
End Function

Private Sub CargoParametrosSucursal()

Dim aNombreTerminal As String

    aNombreTerminal = miConexion.NombreTerminal
    paCodigoDeSucursal = 0: paCodigoDeTerminal = 0
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & aNombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDeTerminal = RsAux!TerCodigo
        paCodigoDGI = RsAux("SucCodDGI")
        If Not IsNull(RsAux!SucDisponibilidad) Then paDisponibilidad = RsAux!SucDisponibilidad Else paDisponibilidad = 0
        paNombreSucursal = Trim(RsAux!SucAbreviacion)
        
        'Nombre de Cada Documento--------------------------------------------------------------------------------
        If Not IsNull(RsAux!SucDContado) Then paDContado = Trim(RsAux!SucDContado)
        If Not IsNull(RsAux!SucDCredito) Then paDCredito = Trim(RsAux!SucDCredito)
        If Not IsNull(RsAux!SucDNDevolucion) Then paDNDevolucion = Trim(RsAux!SucDNDevolucion)
        If Not IsNull(RsAux!SucDNCredito) Then paDNCredito = Trim(RsAux!SucDNCredito)
        If Not IsNull(RsAux!SucDRecibo) Then paDRecibo = Trim(RsAux!SucDRecibo)
        If Not IsNull(RsAux!SucDNEspecial) Then paDNEspecial = Trim(RsAux!SucDNEspecial)

    End If
    RsAux.Close
  
    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(aNombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & vbCrLf & _
                    "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
        End: Exit Sub
    End If
    '-------------------------------------------------------------------------------------------------------------------------

    prj_LoadConfigPrint
    
End Sub

Public Sub CargoParametrosLocal()

    prmSuc_ModificacionDePrecios = 3
    prmSuc_FacturaArticuloNOHabilitado = 13
    prmPathApp = App.Path
    
    paLocalesCobraVencidas = ""
    
    'Parametros a cero-----------------
    paTipoCuotaContado = 0
    paMonedaFacturacion = 0
    
    Cons = "Select * from Parametro"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "tipocuotacontado": paTipoCuotaContado = RsAux!ParValor
            Case "monedafacturacion": paMonedaFacturacion = RsAux!ParValor
            Case "departamento": paDepartamento = RsAux!ParValor
            Case "localidad": paLocalidad = RsAux!ParValor
            
            Case "articulopisoagencia": paArticuloPisoAgencia = RsAux!ParValor
            Case "articulodiferenciaenvio": paArticuloDiferenciaEnvio = RsAux!ParValor
            
            Case "vatoleranciadiasexh": paVaToleranciaDiasExh = RsAux!ParValor
            Case "vatoleranciamonedaporc": paVaToleranciaMonedaPorc = RsAux!ParValor
            Case "vatoleranciadiasexhtit": paVaToleranciaDiasExhTit = RsAux!ParValor
            Case "categoriacliente": paCategoriaCliente = RsAux!ParValor
            Case "planpordefecto": paPlanPorDefecto = RsAux!ParValor
            Case "estadoarticuloentrega": paEstadoArticuloEntrega = RsAux!ParValor
            
            Case "cofis": paCofis = RsAux!ParValor
            
            Case LCase("ArtsSinCofis"): If Not IsNull(RsAux("ParTexto")) Then prmTipoArtSinCofis = Trim(RsAux("ParTexto"))
            
            Case "mcchequediferido": paMCChequeDiferido = RsAux!ParValor
            Case "monedapesos": paMonedaPesos = RsAux!ParValor
                        
            Case "tipoarticuloservicio": paTipoArticuloServicio = RsAux!ParValor
            
            Case "localescobravencidas": paLocalesCobraVencidas = RsAux("ParTexto")
            
            Case "serverasuntos_port_ip"
                    prmPortServer = RsAux!ParValor
                    prmIPServer = Trim(RsAux!ParTexto)
                    
            Case LCase("URLFirmaEFactura"): prmURLFirmaEFactura = Trim(RsAux("ParTexto"))
            Case LCase("efactImporteDatosCliente"): prmImporteConInfoCliente = RsAux("ParValor")
            Case LCase("eFacturaActiva"): prmEFacturaProductivo = RsAux("ParValor")
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub

'---------------------------------------------------------------------------------------------------------------
'   Valores que Retorna:    -1: Error o No Existe
'                                       0: Facturando o Facturada
'                                       1: Bloqueada OK
Public Function fnc_BloqueoSolicitud(Codigo As Long)

    fnc_BloqueoSolicitud = 0
    Screen.MousePointer = 11
    On Error GoTo errorBT
    
    'Bloqueo la solicitud y Actulizo el SolTipoResolucion a Facturando
    Cons = "Select * from Solicitud Where SolCodigo = " & Codigo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        If RsAux!SolEstado <> EstadoSolicitud.Rechazada Then
            
            If RsAux!SolProceso <> TipoResolucionSolicitud.Facturada And RsAux!SolProceso <> TipoResolucionSolicitud.Facturando _
                And Not IsNull(RsAux!SolUsuarioR) Then
    
                cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
                On Error GoTo errorET
                
                RsAux.Requery
                
                If RsAux!SolProceso = TipoResolucionSolicitud.Facturada Or RsAux!SolProceso = TipoResolucionSolicitud.Facturando Then
                    cBase.RollbackTrans
                    Exit Function
                End If
                
                RsAux.Edit
                RsAux!SolProceso = TipoResolucionSolicitud.Facturando
                RsAux.Update
                
                cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
                
                NotificoCambioSignalR
                
                'RsAux.Requery
                fnc_BloqueoSolicitud = 1    'OK
            
            Else
                fnc_BloqueoSolicitud = -1    'OK
            End If
        
        Else
            fnc_BloqueoSolicitud = 1    'OK
        End If
    End If
    
    RsAux.Close
    Screen.MousePointer = 0
    Exit Function

errorBT:
    fnc_BloqueoSolicitud = -1
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Function
    
errorET:
    Resume ErrorRoll
ErrorRoll:
    fnc_BloqueoSolicitud = -1
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación."
End Function

Public Function fnc_HayArticulosDeInstalaciones(mIDDoc As Long) As Boolean
On Error GoTo errFnc

    fnc_HayArticulosDeInstalaciones = False
    
    Cons = "Select * From Renglon, Articulo " & _
            " Where RenDocumento = " & mIDDoc & _
            " And RenArticulo = ArtID " & _
            " And ArtInstalador > 0"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then fnc_HayArticulosDeInstalaciones = True
    RsAux.Close
    
errFnc:
End Function

Public Function ChangeCnfgPrint() As Boolean
    Dim objPrint As New clsCnfgPrintDocument
    ChangeCnfgPrint = (paLastUpdate <> objPrint.FechaUltimoCambio)
    Set objPrint = Nothing
End Function

Public Function fnc_DevolverSolicitud(idSolADevolver As Long, idEstadoSol As Integer) As Boolean

'09/06/2004 Carlos me dijo q no va mas lo de rechazada
'1) No tiene q estar Rechazada  If aSolicitudEstado = EstadoSolicitud.Rechazada Then
'2) Tiene q estar resuelta pero no facturada    rsAux!SolProceso = TipoResolucionSolicitud.Facturando
fnc_DevolverSolicitud = False

Dim bOK As Boolean

    Cons = "Select * from Solicitud " & _
                " Where SolCodigo = " & idSolADevolver & _
                " And SolProceso Not IN ( " & TipoResolucionSolicitud.Facturada & "," & TipoResolucionSolicitud.Facturando & ")" & _
                " And SolEstado Not IN ( " & EstadoSolicitud.Pendiente & ", " & EstadoSolicitud.ParaRetomar & ")"
                
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    bOK = Not RsAux.EOF
    RsAux.Close
    
    If Not bOK Then
        MsgBox "La solicitud seleccionada no se puede devolver" & vbCrLf & vbCrLf & _
                    "1) Controle que No esté Facturada ni en Proceso de Facturación. " & vbCrLf & _
                    "2) No debe estar: Pendiente  o Para Retomar.", vbExclamation
        Exit Function
    End If
        
    If MsgBox("¿Está seguro que quiere " & IIf(idEstadoSol = 5, "dejar sin efecto", "devolver") & " la solicitud?", vbQuestion + vbYesNo + vbDefaultButton2, "Devolver Solicitud") = vbNo Then Exit Function
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    
    FechaDelServidor    'Saco la fecha del servidor
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    Cons = "Select * from Solicitud " & _
                " Where SolCodigo = " & idSolADevolver & _
                " And SolProceso Not IN ( " & TipoResolucionSolicitud.Facturada & "," & TipoResolucionSolicitud.Facturando & ")" & _
                " And SolEstado Not IN ( " & EstadoSolicitud.Pendiente & ", " & EstadoSolicitud.ParaRetomar & ")"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    RsAux.Edit
    RsAux!SolFecha = Format(gFechaServidor, sqlFormatoFH)
    RsAux!SolProceso = TipoResolucionSolicitud.Manual
    RsAux!SolEstado = idEstadoSol
    
    RsAux!SolUsuarioS = paCodigoDeUsuario
    RsAux!SolFResolucion = Null
    If idEstadoSol <> 5 Then RsAux!SolDevuelta = True
    RsAux!SolVisible = Null
    
    RsAux.Update
    RsAux.Close
    
    Screen.MousePointer = 0
    cBase.CommitTrans    'FIN DE TRANSACCION------------------------------------------
    
    NotificoCambioSignalR
    fnc_DevolverSolicitud = True
    Screen.MousePointer = 0
    Exit Function

errorBT:
    clsGeneral.OcurrioError "Devolver Solicitud: No se ha podido inicializar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Function
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Devolver Solicitud: No se ha podido realizar la transacción.", Err.Description
    Screen.MousePointer = 0
End Function

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
Public Function z_BuscoUsuario(Codigo As Long, Optional Identificacion As Boolean = False, Optional Digito As Boolean = False, Optional Iniciales As Boolean = False)
    z_BuscoUsuario = BuscoUsuario(Codigo, Identificacion, Digito, Iniciales)
End Function

Public Function BuscoUsuario(Codigo As Long, Optional Identificacion As Boolean = False, Optional Digito As Boolean = False, Optional Iniciales As Boolean = False)
Dim RsUsr As rdoResultset
Dim aRetorno As String: aRetorno = ""
    
    On Error Resume Next
    
    Cons = "Select * from Usuario Where UsuCodigo = " & Codigo
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Digito Then aRetorno = Trim(RsUsr!UsuDigito)
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    BuscoUsuario = aRetorno
    
End Function

Public Function z_BuscoUsuarioDigito(Digito As Long, Optional Codigo As Boolean = False, Optional Identificacion As Boolean = False, Optional Iniciales As Boolean = False) As Variant
    z_BuscoUsuarioDigito = BuscoUsuarioDigito(Digito, Codigo, Identificacion, Iniciales)
End Function

Public Function BuscoUsuarioDigito(Digito As Long, Optional Codigo As Boolean = False, Optional Identificacion As Boolean = False, Optional Iniciales As Boolean = False) As Variant
Dim RsUsr As rdoResultset
Dim aRetorno As Variant
On Error GoTo ErrBUD

    Cons = "Select * from Usuario Where UsuDigito = " & Digito
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Codigo Then aRetorno = RsUsr!UsuCodigo
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    BuscoUsuarioDigito = aRetorno
    Exit Function
    
ErrBUD:
    MsgBox "Error inesperado al buscar el usuario.", vbCritical, "ATENCIÓN"
End Function

Public Function PropiedadesConnect(Conexion As String, _
                                                    Optional Database As Boolean = True, Optional DSN As Boolean = False, _
                                                    Optional Server As Boolean = True) As String
Dim aRetorno As String

    On Error GoTo errConnect
    PropiedadesConnect = ""
    Conexion = UCase(Conexion)
    If DSN Then aRetorno = Trim(Mid(Conexion, InStr(Conexion, "DSN=") + 4, Len(Conexion)))
    If Server Then aRetorno = Trim(Mid(Conexion, InStr(Conexion, "SERVER=") + 7, Len(Conexion)))
    If Database Then aRetorno = Trim(Mid(Conexion, InStr(Conexion, "DATABASE=") + 9, Len(Conexion)))
    
    aRetorno = Trim(Mid(aRetorno, 1, InStr(aRetorno, ";") - 1))
    
    PropiedadesConnect = aRetorno
    Exit Function
    
errConnect:
End Function

Private Sub CargoTiposDeArticulosServicios()
Dim sQy As String
    tTiposArtsServicio = ""
    sQy = "SELECT TipID FROM dbo.InTipos(" & paTipoArticuloServicio & ")"
    If ObtenerResultSet(cBase, RsAux, sQy, "Comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
    Do While Not RsAux.EOF
        tTiposArtsServicio = tTiposArtsServicio & IIf(tTiposArtsServicio <> "", ",", "") & Trim(RsAux("TipID"))
        RsAux.MoveNext
    Loop
    RsAux.Close
End Sub

Public Function EsTipoDeServicio(ByVal idTipo As Long) As Boolean
     EsTipoDeServicio = (InStr(1, "," & tTiposArtsServicio & ",", "," & idTipo & ",") > 0)
End Function

'METODOS SIGNALR
Private Sub CargoPrmsSignalR()
On Error GoTo errCPS
Dim sQy As String
Dim rsP As rdoResultset

    sQy = "select ParNombre, ParTexto From Parametro where ParNombre like 'signalr%'"
    Set rsP = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not rsP.EOF
        Select Case LCase(Trim(rsP("ParNombre")))
            Case LCase("signalRURL")
                prmHUBURL = Trim(rsP("ParTexto"))
        End Select
        rsP.MoveNext
    Loop
    rsP.Close
    Exit Sub
errCPS:
clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description, "Cargo Parámetros"
End Sub

Public Sub NotificoCambioSignalR()
    
    Dim hubSol As New ClientHub
    If hubSol.ConnectHub(prmHUBURL, "AsuntosPendientesHub") Then
        hubSol.InvokeMethod "ObtenerSolicitudesResueltas"
        hubSol.InvokeMethod "ObtenerAsuntosPendientes"
    End If
End Sub
