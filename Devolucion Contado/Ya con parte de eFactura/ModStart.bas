Attribute VB_Name = "ModStart"
Option Explicit
Public paLocalZF As Long
Public paINContadoB As Integer
Public paINContadoN As String

Public paIConformeB As Integer
Public paIConformeN As String
Public paIConformeP As Integer

Public paPrintEsXDefNC As Boolean           'Nota Ctdo.
Public paPrintEsXDefCn As Boolean            'Conforme

Public paLastUpdate As String
Public paOptPrintSel As String      'El nombre de la opción seleccionada
Public paOptPrintList As String      'Los nombres de las opciones ingresadas están separadas x |

Public gPathListados As String, paBD As String

Public clsGeneral As New clsorCGSA
Public miconexion As New clsConexion

Public Sub Main()
On Error GoTo ErrMain
    
    Screen.MousePointer = 11
    If miconexion.AccesoAlMenu(App.title) Then
        If InicioConexionBD(miconexion.TextoConexion("comercio")) Then

            paBD = miconexion.RetornoPropiedad(bDB:=True)
            paCodigoDeUsuario = miconexion.UsuarioLogueado(True)
            
            If Not loc_GetInfoSucursal Then
                Screen.MousePointer = 0
                End: Exit Sub
            End If
            
            PathListados
            CargoParametrosComercio
            prj_LoadConfigPrint
            
            frmNotaCtdo.Show vbModeless
            
        Else
            Screen.MousePointer = 0
            Set clsGeneral = Nothing: Set miconexion = Nothing
            End
        End If
    Else
        If miconexion.UsuarioLogueado(Codigo:=True) <> 0 Then MsgBox "Ud. no tiene permisos de acceso para la aplicación.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        End
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrMain:
    clsGeneral.OcurrioError "Ocurrio un error al activar el ejecutable.", Trim(Err.Description)
    Set clsGeneral = Nothing: Set miconexion = Nothing
    Screen.MousePointer = 0
    End
End Sub

Private Sub PathListados()
On Error GoTo errPL
    ChDir App.Path
    ChDir ("..")
    ChDir (CurDir & "\REPORTES\")
    gPathListados = CurDir & "\"
    Exit Sub
errPL:
    MsgBox "No se encontro la carpeta de reportes.", vbInformation, "ATENCIÓN"
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

Private Function loc_GetInfoSucursal() As Boolean

    loc_GetInfoSucursal = False
    
    paCodigoDeSucursal = 0
    paCodigoDeTerminal = 0
    paDContado = ""
    
    'Saco el codigo de la sucursal por el nombre de la Terminal-----------------------------------------------------
    Cons = "Select * From Terminal, Sucursal" _
            & " Where TerNombre = '" & miconexion.NombreTerminal & "'" _
            & " And TerSucursal = SucCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        paCodigoDeSucursal = RsAux!TerSucursal
        paCodigoDeTerminal = RsAux!TerCodigo
        If Not IsNull(RsAux!SucDisponibilidad) Then paDisponibilidad = RsAux!SucDisponibilidad Else paDisponibilidad = 0
        
        'El documento que necesito solo es el contado.
        'If Not IsNull(RsAux!SucDContado) Then paDContado = Trim(RsAux!SucDContado)
        If Not IsNull(RsAux!SucDNDevolucion) Then paDNDevolucion = Trim(RsAux!SucDNDevolucion)
    End If
    RsAux.Close
  
    If paCodigoDeSucursal = 0 Then
        MsgBox "La terminal " & UCase(miconexion.NombreTerminal) & " no pertenece a ninguna de las sucursales de la empresa." & Chr(vbKeyReturn) _
                & "La ejecución será cancelada.", vbCritical, "ATENCIÓN"
        Exit Function
    Else
        If paDNDevolucion = "" Then
            MsgBox "No se encontró el nombre del documento 'Nota de Devolución', comuniquese con el administrador.", vbCritical, "ATENCIÓN"
        Else
            loc_GetInfoSucursal = True
        End If
    End If
    '-------------------------------------------------------------------------------------------------------------------------
    
End Function

Public Sub prj_LoadConfigPrint(Optional bShowFrm As Boolean)
On Error GoTo errLCP
Dim objPrint As New clsCnfgPrintDocument
Dim sPNC As String, sPCn As String
Dim vPrint() As String

    With objPrint
        Set .Connect = cBase
        .Terminal = paCodigoDeTerminal
        If bShowFrm Then
            .CnfgTipoDocumento = "1,6"
            .ShowConfig
        End If
        If paLastUpdate <> .FechaUltimoCambio Or paLastUpdate = "" Then
            sPNC = .getDocumentoImpresora(Contado)
            sPCn = .getDocumentoImpresora(Remito)
            paOptPrintSel = .GetOpcionActual
            paOptPrintList = .GetOpcionesPrinter
            paLastUpdate = .FechaUltimoCambio
        End If
    End With
    Set objPrint = Nothing
    
    If sPNC <> "" Then
        vPrint = Split(sPNC, "|")
        paINContadoN = vPrint(0)
        paINContadoB = vPrint(1)
        paPrintEsXDefNC = (Val(vPrint(3)) = 1)
    End If
    
    If sPCn <> "" Then
        vPrint = Split(sPCn, "|")
        paIConformeN = vPrint(0)
        paIConformeB = vPrint(1)
        paIConformeP = Val(vPrint(2))
        
        paPrintEsXDefCn = (Val(vPrint(3)) = 1)
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

