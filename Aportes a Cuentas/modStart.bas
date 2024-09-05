Attribute VB_Name = "modStart"
Option Explicit

Public miConexion As New clsConexion
Public clsGeneral As New clsorCGSA
Public paToleranciaAportes As Byte

Public paLocalPuerto As Long, paLocalZF As Long

Public prmPathApp As String

Public paIReciboN As String
Public paIReciboB As Integer
Public paIRemitoB As Integer
Public paIRemitoN As String

Public paNombreSucursal As String

Private paLastUpdate As String
Public paOptPrintSel As String      'El nombre de la opción seleccionada
Public paOptPrintList As String      'Los nombres de las opciones ingresadas están separadas x |

Public Sub Main()
'Prms: Tipo de Cta, id de cta, id de cliente q aporta (si vacio cargo abajo) separados con |

Dim aSucursal As String

    On Error GoTo errMain
    Screen.MousePointer = 11
    Dim bAccesoOK As Boolean, mAppVer As String
    
    bAccesoOK = False: mAppVer = App.title
    bAccesoOK = miConexion.AccesoAlMenu(mAppVer)
    
    If bAccesoOK Then
        If mAppVer <> "" And Not (App.Major & "." & App.Minor & "." & App.Revision) >= mAppVer Then
            MsgBox "La versión del programa no es la última disponible." & vbCr & _
                        "Ud. debe actualizar el software.", vbExclamation, "Actualizar a Versión " & mAppVer
            End
        End If

        InicioConexionBD miConexion.TextoConexion(logComercio)
        
        
        Cons = "SELECT RTRIM(ParNombre) Nombre, ParValor FROM Parametro WHERE ParNombre IN('AportesACuentaMesesDisponibles')"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            paToleranciaAportes = RsAux("ParValor")
        End If
        RsAux.Close
        
        CargoParametrosComercio
        paNombreSucursal = CargoParametrosSucursalLocal
        
        CargoParametrosLocal
        
        If Trim(Command()) <> "" Then
            Dim sParams() As String
            sParams = Split(Trim(Command()), "|")
            '0- prmTipoCta   1-prmIdCta     2-prmIDAporta
            frmAporte.prmTipoCta = sParams(0)
            If UBound(sParams) >= 1 Then frmAporte.prmIdCta = Trim(sParams(1))
            If UBound(sParams) >= 2 Then frmAporte.prmIDAporta = Trim(sParams(2))
            
        End If
        frmAporte.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
        frmAporte.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
        
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        
        
        frmAporte.Show vbModeless
    
    Else
        If miConexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    
    Exit Sub
errMain:
    On Error Resume Next
    Screen.MousePointer = 0
    MsgBox "Error al inicializar la aplicación " & App.title & Chr(vbKeyReturn) & "Error: " & Trim(Err.Description)
    End
End Sub

Sub CargoParametrosLocal()

    Cons = "Select * from Parametro Where ParNombre = 'pathapp'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case LCase(Trim(RsAux!ParNombre))
            Case "pathapp": prmPathApp = Trim(RsAux!ParTexto)
        End Select
    
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    prj_LoadConfigPrint bShowFrm:=False
    
End Sub


Public Sub prj_LoadConfigPrint(Optional bShowFrm As Boolean)
On Error GoTo errLCP

Dim objPrint As New clsCnfgPrintDocument
Dim mCRecibo As String, sPrintRemito As String
Dim vPrint() As String

    With objPrint
        Set .Connect = cBase
        .Terminal = paCodigoDeTerminal
        If bShowFrm Then
            .CnfgTipoDocumento = TipoDocumento.ReciboDePago
            .ShowConfig
        End If
        
        If paLastUpdate <> .FechaUltimoCambio Or paLastUpdate = "" Then
            mCRecibo = .getDocumentoImpresora(ReciboDePago)
            
            paOptPrintSel = .GetOpcionActual
            paOptPrintList = .GetOpcionesPrinter
            paLastUpdate = .FechaUltimoCambio
            
            sPrintRemito = .getDocumentoImpresora(Remito)
            
            If mCRecibo = "" Then
                MsgBox "Falta alguna de las configuraciones de impresoras." & vbCrLf & _
                            "Valide los datos de configuración antes de imprimir.", vbCritical, "Faltan Valores de Impresión"
            End If

        End If
    End With
    Set objPrint = Nothing
    
    
    If mCRecibo <> "" Then
        vPrint = Split(mCRecibo, "|")
        paIReciboN = vPrint(0)
        paIReciboB = vPrint(1)
    End If
    
    If sPrintRemito <> "" Then
        vPrint = Split(sPrintRemito, "|")
        paIRemitoN = vPrint(0)
        paIRemitoB = vPrint(1)
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

Public Function CargoParametrosSucursalLocal() As String

Dim aNombreTerminal As String

    CargoParametrosSucursalLocal = ""
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
        CargoParametrosSucursalLocal = Trim(RsAux!SucAbreviacion)
        
        If Not IsNull(RsAux!SucDRecibo) Then paDRecibo = Trim(RsAux!SucDRecibo)
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


