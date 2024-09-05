Attribute VB_Name = "ModStart"
Option Explicit

Public Const cnfgKeyTicketMovimientoCaja As String = "TickeadoraMovimientosDeCaja"
Public Const cnfgAppNombreMovimientoCaja As String = "MovimientosDeCaja"



'Impresora
Public paPrintConfB As Integer
Public paPrintConfD As String
Public paPrintConfXDef As Boolean
Public paPrintConfPaperSize As Integer


Public paToleranciaAportes As Integer

Public txtConexion As String
Public palocalzf As Long
Public clsGeneral As New clsorCGSA
Public UsuLogueado As Long
Public miconexion As New clsConexion

Public Sub Main()
On Error GoTo ErrMain
Dim vParam() As String
    
    If App.StartMode = vbSModeStandalone Then
        Screen.MousePointer = 11
        If miconexion.AccesoAlMenu(App.title) Then
            
            txtConexion = miconexion.TextoConexion(logComercio)
            InicioConexionBD txtConexion
            UsuLogueado = miconexion.UsuarioLogueado(True)
            
            CargoParametrosComercio
            CargoParametrosSucursal
            
            Cons = "SELECT RTRIM(ParNombre) Nombre, ParValor FROM Parametro WHERE ParNombre IN('AportesACuentaMesesDisponibles')"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                paToleranciaAportes = RsAux("ParValor")
            End If
            RsAux.Close
            
            prj_GetPrinter False
            
            If Trim(Command()) <> "" Then
                vParam = Split(Command(), "|")
                Select Case UBound(vParam)
                    Case 1
                        frmAsiFacCta.prmTipo = vParam(0)
                        frmAsiFacCta.prmID = vParam(1)
                    Case 2
                        frmAsiFacCta.prmTipo = vParam(0)
                        frmAsiFacCta.prmID = vParam(1)
                        frmAsiFacCta.prmDocumento = vParam(2)
                    Case 3
                        frmAsiFacCta.prmTipo = vParam(0)
                        frmAsiFacCta.prmID = vParam(1)
                        frmAsiFacCta.prmDocumento = vParam(2)
                        frmAsiFacCta.prmImporte = vParam(3)
                End Select
            End If
            frmAsiFacCta.Show vbModeless
        Else
            If miconexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
            End
            Screen.MousePointer = 0
        End If
    Else
        'Lo implementamos por si no esta corriendo login pida contraseña.
        miconexion.AccesoAlMenu (App.title)
        InicioConexionBD miconexion.TextoConexion(logComercio)
    End If
    Exit Sub
ErrMain:
    clsGeneral.OcurrioError "Ocurrio un error al activar el ejecutable.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Public Sub prj_GetPrinter(ByVal bShowP As Boolean)
On Error GoTo errImp
    paPrintConfD = ""
    paPrintConfB = -1
    Dim objP As New clslPrintConfig
    With objP
        If bShowP Then
            If Not .ShowPrinterSetup("6", paCodigoDeTerminal) Then
                GoTo errImp
            End If
        End If
        If .LoadPrinterConfig("6", paCodigoDeTerminal) Then
            .GetPrinterDoc 6, paPrintConfD, paPrintConfB, paPrintConfXDef, paPrintConfPaperSize
        End If
    End With
    If paPrintConfD = "" Then MsgBox "Por favor verifique la configuración de impresión.", vbInformation, "Atención"
    
errImp:
    Set objP = Nothing
    Screen.MousePointer = 0
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
