Attribute VB_Name = "modStart"
Option Explicit

'REGISTRO DE SUCESOS--------------------------------------
Public Enum TipoSuceso
    ModificacionDeMora = 1
    AnulacionDeDocumentos = 2
    ModificacionDePrecios = 3
    RecepcionDeTraslados = 4
    AnulacionDeEnvios = 5
    CambioCostoDeFlete = 6
    Direcciones = 7
    ChequesDiferidos = 8
    CambioCategoriaCliente = 9
    Reimpresiones = 10
    DiferenciaDeArticulos = 11
    CederProductoServicio = 12
    FacturaArticuloInhabilitado = 13
    Notas = 14
    FacturaPlanInhabilitado = 15
    FacturaCambioNombre = 16
    CambioTipoArticuloServicio = 17
    ConfiguracionSistema = 18
    EliminacionInstalacion = 21
    VariosStock = 98
    Varios = 99
End Enum
'--------------------------------------------------------------------

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA

Public prmTipoCuotaContado As Long
Public prmPrimeraHoraEnvio As Long
Public prmUltimaHoraEnvio As Long
Public prmArticuloPagaInstalacion As String
Public prmTipoTelefP As Long, prmTipoTelefE As Long

Public Sub Main()
Dim sPrm As String
Dim vPrm() As String
Dim iCont As Integer

Dim sPaso As String


    On Error GoTo errMain
    Screen.MousePointer = 11
    
    Dim bAccesoOK As Boolean, mAppVer As String
    
'    sPaso = "ID Usuario"
'    Dim idUS As Long
'    idUS = miConexion.UsuarioLogueado(True)
'    MsgBox idUS, vbInformation, "ID usuario logueado"
'
    sPaso = "Acceso"
    bAccesoOK = False: mAppVer = App.Title
    bAccesoOK = miConexion.AccesoAlMenu(mAppVer)
    If bAccesoOK Then
        
        sPaso = "Versión"
'        If mAppVer <> "" And mAppVer > App.Major & "." & App.Minor & "." & App.Revision Then
'            MsgBox "La versión del programa no es la última disponible." & vbCr & vbCr & _
'                        "Su versión es la " & App.Major & "." & App.Minor & "." & App.Revision & vbCr & _
'                        "Ud. debe actualizar el software a la versión " & mAppVer, vbExclamation, "Actualizar a Versión " & mAppVer
'
'            Set miConexion = Nothing
'            Set clsGeneral = Nothing
'            End
'            Exit Sub
'
'        End If
    
        sPaso = "Conexión"
        If Not InicioConexionBD(miConexion.TextoConexion("Comercio")) Then
            Set miConexion = Nothing
            Set clsGeneral = Nothing
            End
            Exit Sub
        End If
        sPaso = "Parámetros"
        CargoParametrosLocales
        
        sPaso = "ID Usuario"
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        
        sPaso = "Pars de ingreso"
        'Parámetros de ingreso.
        sPrm = Command()
        frmInstall.prmDocumento = 0
        frmInstall.prmIDInst = 0
        
        If sPrm <> "" Then
            vPrm = Split(sPrm, ";", , vbTextCompare)
            For iCont = 0 To UBound(vPrm)
                If InStr(1, vPrm(iCont), "doc:", vbTextCompare) > 0 Then
                    frmInstall.prmDocumento = Replace(vPrm(iCont), "doc:", "", , , vbTextCompare)
                    frmInstall.prmTipoDoc = 1
                ElseIf InStr(1, vPrm(iCont), "id:", vbTextCompare) Then
                    frmInstall.prmIDInst = Replace(vPrm(iCont), "id:", "", , , vbTextCompare)
                ElseIf InStr(1, vPrm(iCont), "vTe:", vbTextCompare) Then
                    frmInstall.prmDocumento = Replace(vPrm(iCont), "VTe:", "", , , vbTextCompare)
                    frmInstall.prmTipoDoc = 2
                End If
            Next
        End If
        sPaso = "Show form"
        frmInstall.Show vbModeless
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Ocurrió un error al inicializar la aplicación " & App.Title & Chr(vbKeyReturn) & "Error en paso " & sPaso & ", descripción: " & Trim(Err.Description), vbCritical, "Instalaciones"
    End
End Sub

Private Sub CargoParametrosLocales()

    On Error GoTo errParametro
    prmArticuloPagaInstalacion = ""
    Cons = "Select * from Parametro Where ParNombre In ('insArticuloCobro', 'TipoTelefonoP', 'TipoTelefonoE', 'PrimeraHoraEnvio', 'UltimaHoraEnvio', 'tipocuotacontado')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case LCase("TipoTelefonoP"): prmTipoTelefP = RsAux!ParValor
            Case LCase("TipoTelefonoE"): prmTipoTelefE = RsAux!ParValor
            Case LCase("PrimeraHoraEnvio"): prmPrimeraHoraEnvio = RsAux!ParValor
            Case LCase("UltimaHoraEnvio"): prmUltimaHoraEnvio = RsAux!ParValor
            
            Case "tipocuotacontado": prmTipoCuotaContado = RsAux!ParValor
            Case "insarticulocobro": prmArticuloPagaInstalacion = Trim(RsAux!ParTexto)
            
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close

    Exit Sub
errParametro:
    clsGeneral.OcurrioError "Error al cargar los parámetros.", Err.Description
End Sub


