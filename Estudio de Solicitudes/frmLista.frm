VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLista 
   Caption         =   "Asuntos Pendientes"
   ClientHeight    =   3615
   ClientLeft      =   2715
   ClientTop       =   4710
   ClientWidth     =   6960
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLista.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3615
   ScaleWidth      =   6960
   Begin VB.Timer tmToBD 
      Interval        =   15000
      Left            =   2040
      Top             =   600
   End
   Begin VB.Timer tmSignalR 
      Left            =   1560
      Top             =   600
   End
   Begin VB.Timer tmCheckLog 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   120
      Top             =   240
   End
   Begin VB.Timer tmAuto 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   600
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsAsunto 
      Height          =   1815
      Left            =   840
      TabIndex        =   0
      Top             =   1140
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3201
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   8421631
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   285
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0   'False
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
   End
   Begin MSComctlLib.ImageList Image1 
      Left            =   4320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   12632256
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":0442
            Key             =   "solicitud"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":075C
            Key             =   "llamara"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":0A76
            Key             =   "gastos"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":0D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":10AA
            Key             =   "s1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":13C4
            Key             =   "s0"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":16DE
            Key             =   "e0"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":19F8
            Key             =   "s9"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":1D12
            Key             =   "s3"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":202C
            Key             =   "s4"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":2346
            Key             =   "s5"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":2660
            Key             =   "s6"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":297A
            Key             =   "s7"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":2C94
            Key             =   "s8"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":2FAE
            Key             =   "s2"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":32C8
            Key             =   "e9"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":35E2
            Key             =   "e2"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":38FC
            Key             =   "e3"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":3C16
            Key             =   "e4"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":3F30
            Key             =   "e5"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":424A
            Key             =   "e6"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":4564
            Key             =   "e7"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":487E
            Key             =   "e8"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":4B98
            Key             =   "e1"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":4EB2
            Key             =   "original"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":51CC
            Key             =   "sucesos"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLista.frx":58FE
            Key             =   "servicio"
         EndProperty
      EndProperty
   End
   Begin VB.Label lTrama 
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   3180
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label lStatus 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   195
      Left            =   5220
      TabIndex        =   1
      Top             =   60
      Width           =   1635
   End
   Begin VB.Menu MnuFiltro 
      Caption         =   "Opciones"
      Visible         =   0   'False
      Begin VB.Menu MnuAutorizar 
         Caption         =   "Autorizar"
      End
      Begin VB.Menu MnuNOAutorizar 
         Caption         =   "No Autorizar"
      End
      Begin VB.Menu MnuAutorizarAll 
         Caption         =   "Autorizar Todo lo Seleccionado"
      End
      Begin VB.Menu MnuAnalizar 
         Caption         =   "Analizar Compra/Gasto"
      End
      Begin VB.Menu MnuAnalizarP 
         Caption         =   "Analizar Proveedor"
      End
      Begin VB.Menu MnuFl1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFiltrarProveedor 
         Caption         =   "Filtrar Proveedor"
      End
      Begin VB.Menu MnuQuitarFiltros 
         Caption         =   "Quitar Filtros"
      End
      Begin VB.Menu MnuFl2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuServicio 
         Caption         =   "Cargar Servicios"
      End
      Begin VB.Menu MnuLineLiberar 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLiberarSolicitud 
         Caption         =   "Liberar solicitud"
      End
      Begin VB.Menu MnuLineaRefrescar 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRefrescar 
         Caption         =   "Refrescar grilla"
      End
   End
   Begin VB.Menu MnuTray 
      Caption         =   "MnuTray"
      Visible         =   0   'False
      Begin VB.Menu MnuTEnd 
         Caption         =   "Finalizar Aplicacion"
      End
      Begin VB.Menu MnuTL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuTsiguiente 
         Caption         =   "Activar Suguiente Asunto"
      End
      Begin VB.Menu MnuTAsuntos 
         Caption         =   "Ver Asuntos Pendientes"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuTOcultar 
         Caption         =   "Ocultar Asuntos Pendientes"
      End
   End
End
Attribute VB_Name = "frmLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim prmHUBURL As String
Dim WithEvents oHub As ClientHub
Attribute oHub.VB_VarHelpID = -1
Private prmIDSel As Long                 'Codigo del Item Seleccionado

Dim rsQry As rdoResultset
Dim mValor As Long
Dim globalData As String

Dim arrTramas() As String

Dim prmCurrentUser As Integer
Dim prmCurrentUserName As String

Dim QSinAnalizar As Integer, QServicios As Integer

Private Sub Form_Activate()
    On Error Resume Next
    vsAsunto.SetFocus
End Sub

Private Sub Form_Load()

    On Error Resume Next
    ReDim arrTramas(0)
    prmCurrentUser = -1
    
    tmToBD.Enabled = False
    TrayNotify Me.hwnd, Me, ""
    Me.Refresh
    
    MnuServicio.Checked = ObtengoSeteoControl(Me.Name & MnuServicio.Name & "Checked")
    
    z_InicializoGrilla
    
    ObtengoSeteoForm Me, Me.Left, Me.Top, Me.Width, Me.Height
    
    FechaDelServidor
    fnc_CargoParametrosSonido
    
    tmSignalR.Enabled = False
    tmSignalR.Interval = 5
    
    CargoPrmsSignalR
    
    If Not ConectarSignalR Then
        MsgBox "No conectó el signalR " & prmHUBURL & ", se procede a lectura por tiempo.", vbCritical, "ATENCIÓN"
        tmToBD.Interval = 15000
        tmToBD.Enabled = True
    Else
        signalR_RefrescoSolicitudes
        tmCheckLog.Enabled = True
        'prmCurrentUser = paCodigoDeUsuario
    End If
    
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim aMM As Integer

    aMM = TrayClick(Button, Shift, X, Y)
    Select Case aMM
        Case 1: acc_SiguienteAsunto
        Case 2: PopupMenu MnuTray
    End Select
    
End Sub

Private Sub MnuAnalizar_Click()

    prmIDSel = vsAsunto.Cell(flexcpValue, vsAsunto.Row, 0)
    Select Case z_TipoDeGasto(prmIDSel)
        Case 2
            EjecutarApp pathApp & "appExploreMsg.exe", CStr(prmPlantillaACompras) & ":" & CStr(prmIDSel)
            
        Case Else
            EjecutarApp pathApp & "appExploreMsg.exe", CStr(prmPlantillaAGastos) & ":" & CStr(prmIDSel)
    End Select
    
End Sub

Private Sub MnuAnalizarP_Click()
On Error GoTo errFncP
    
    prmIDSel = vsAsunto.Cell(flexcpValue, vsAsunto.Row, 0)

    Dim rs2 As rdoResultset
    Dim mIDProveedor As Long: mIDProveedor = 0
    
    Cons = "Select ComProveedor from Compra Where ComCodigo =" & prmIDSel
    Set rs2 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rs2.EOF Then mIDProveedor = rs2!ComProveedor
    rs2.Close
    
    If mIDProveedor = 0 Then Exit Sub
    
    EjecutarApp pathApp & "appExploreMsg.exe", CStr(prmPlantillaAProveedores) & ":" & CStr(prmIDSel)
        
errFncP:
End Sub

Private Sub MnuAutorizarAll_Click()
On Error GoTo errAAll
Dim iRow As Integer, mIDItem As Long, mRowsRemove As String
Dim rs2 As rdoResultset, mSQL As String

    mRowsRemove = ""
    With vsAsunto
        For iRow = .FixedRows To .Rows - 1
            
            Select Case .Cell(flexcpData, iRow, 0)
                Case Asuntos.GastosAAutorizar
                    If .Cell(flexcpChecked, iRow, 4) = flexChecked Then
                        mIDItem = .Cell(flexcpValue, iRow, 0)
                        
                        'mSQL = "Select * from Compra Where ComCodigo = " & mIDItem
                        'Set rs2 = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
                        'If Not rs2.EOF Then
                        '    rs2.Edit
                        '    rs2!ComFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
                        '    rs2!ComVerificado = 1
                        '    rs2.Update
                        'End If
                        'rs2.Close
                        
                        'Estado 0 = Pendiente, 1 = Autorizado, 2 = No Autorizado
                        mSQL = "Select * from ZureoCGSA.dbo.cceComprobantes Where ComID = " & mIDItem
                        Set rs2 = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
                        If Not rs2.EOF Then
                            rs2.Edit
                            rs2("ComEstado").value = 1
                            rs2("ComFechaEstado").value = Now 'Format(Now, prmBDDateFormat)
                            rs2.Update
                        End If
                        rs2.Close
                        
                        
                        If Trim(mRowsRemove) <> "" Then mRowsRemove = mRowsRemove & ","
                        mRowsRemove = mRowsRemove & Asuntos.GastosAAutorizar & "|" & mIDItem
                        
                    End If
                
            End Select
            
        Next
    End With
    
    If mRowsRemove <> "" Then
        Dim arrItm() As String, arrV() As String
        arrItm = Split(mRowsRemove, ",")
        For iRow = LBound(arrItm) To UBound(arrItm)
            arrV = Split(arrItm(iRow), "|")
            
            mValor = itm_FindItem(CInt(arrV(0)), CLng(arrV(1)))
            vsAsunto.RemoveItem mValor
        Next
    End If
    
    Exit Sub
    
errAAll:
    clsGeneral.OcurrioError "Error al autorizar los asuntos seleccionados.", Err.Description
End Sub

Private Sub MnuFiltrarProveedor_Click()

    Dim pstrProveedor As String, iRow As Integer
    
    With vsAsunto
        pstrProveedor = Trim(.Cell(flexcpText, vsAsunto.Row, 1))
        
        For iRow = .FixedRows To .Rows - 1
            If .Cell(flexcpData, iRow, 0) = Asuntos.GastosAAutorizar Then
                If pstrProveedor <> Trim(.Cell(flexcpText, iRow, 1)) Then
                    .RowHidden(iRow) = True
                End If
            End If
        Next
        
    End With
    
End Sub

Private Sub MnuLiberarSolicitud_Click()

    If MsgBox("¿Confirma liberar la solicitud seleccionada?", vbQuestion + vbYesNo, "Liberar solicitud trancada") = vbYes Then
        
        Screen.MousePointer = 11
        On Error GoTo errorBT
        
        cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
        On Error GoTo errorET
    
        'Bloqueo la solicitud y Actulizo el SolUsuarioR (Analizando)
        Cons = "Select * from Solicitud Where SolCodigo = " & vsAsunto.Cell(flexcpValue, vsAsunto.Row, 0) & " AND SolEstado = " & EstadoSolicitud.Pendiente
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If Not RsAux.EOF Then
            
            RsAux.Edit
            RsAux!SolUsuarioR = Null
            RsAux!SolEstado = EstadoSolicitud.Pendiente
            RsAux.Update
        End If
        RsAux.Close
        cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
        
        signalR_RefrescoSolicitudes
        Screen.MousePointer = 0
    End If
    
Exit Sub

errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación."
    
End Sub

Private Sub MnuNOAutorizar_Click()
    If Trim(MnuAutorizar.Tag) <> "" Then acc_AutorizarAsunto bAutorizado:=False
End Sub

Private Sub MnuQuitarFiltros_Click()
    Dim iRow As Integer
    With vsAsunto
        For iRow = .FixedRows To .Rows - 1
            If .RowHidden(iRow) Then .RowHidden(iRow) = False
        Next
    End With

End Sub

Private Sub MnuRefrescar_Click()
'Vuelvo a pedir la info.
    If LCase(tmToBD.Tag) = "conectado" Then
        Me.signalR_RefrescoSolicitudes
    Else
        CargoAsuntosPendientes
    End If
End Sub

Private Sub oHub_CallBack(ByVal methodName As String, ByVal info As String)

    If LCase(methodName) = LCase("AsuntosPendientesServerON") Then
        tmSignalR.Tag = Now
    ElseIf LCase(methodName) = LCase("CargarAsuntosPendientes") Then
        Dim vSolicitudes() As String
        vSolicitudes = Split(info, vbCrLf)
        Dim mIDT As Integer, iT As Integer
        For iT = 0 To UBound(vSolicitudes)
            If Trim(arrTramas(0)) = "" Then mIDT = 0 Else mIDT = UBound(arrTramas) + 1
            ReDim Preserve arrTramas(mIDT)
            arrTramas(mIDT) = vSolicitudes(iT)
        Next
        vsAsunto.Rows = 1
        tmSignalR.Enabled = True
    End If
    
End Sub

Private Sub tmCheckLog_Timer()
On Error GoTo errFnc
    tmCheckLog.Enabled = False
    ws_SendWhoIam
    
    Dim iT As Integer
    With vsAsunto
        For iT = .FixedRows To .Rows - 1
            If .Cell(flexcpData, iT, 0) = Asuntos.solicitudes Then
                If Not (.Cell(flexcpData, iT, 5) <> 0 And .Cell(flexcpData, iT, 1) <> EstadoSolicitud.ParaRetomar) Then
                    If .Cell(flexcpFontBold, iT, 4) = False And .Cell(flexcpForeColor, iT, 4) = .ForeColor Then
                        .Cell(flexcpText, iT, 4) = z_DifTime(CDate(.Cell(flexcpText, iT, 1)))
                    End If
                End If
            End If
        Next
    End With

errFnc:
    tmCheckLog.Enabled = True
End Sub

Private Sub tmSignalR_Timer()
On Error Resume Next
    tmSignalR.Enabled = False
    lStatus.Caption = "Procesando datos ...": lStatus.Refresh
    ws_ProcesoDataArraival
    lStatus.Caption = "": lStatus.Refresh
End Sub

Private Sub tmToBD_Timer()
    
    tmToBD.Enabled = False
    
    Me.Caption = Trim(Me.Caption) & " (Actualizando...)"
    CargoAsuntosPendientes
    Me.Caption = "Asuntos Pendientes"
    tmToBD.Enabled = Not ConectarSignalR

End Sub

Private Sub vsAsunto_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (vsAsunto.Cell(flexcpData, Row, 0) = Asuntos.GastosAAutorizar And Col = 4) Then
        Cancel = True
    End If
End Sub

'Private Sub wsSocket_Close()
'    On Error Resume Next
'    lStatus.Caption = "Desconectado": lStatus.Refresh
'    Timer1.Enabled = True
'End Sub

'Private Sub wsSocket_DataArrival(ByVal bytesTotal As Long)
'
'    Dim strDato As String
'    Dim xPos As Long
'
'    If wsSocket.BytesReceived > 0 Then
'
'        wsSocket.GetData strDato
'        globalData = globalData & strDato
'
'        Do While InStr(globalData, sc_FIN) <> 0
'            xPos = InStr(globalData, sc_FIN)
'
'            strDato = Mid(globalData, 1, xPos - 1)
'            globalData = Mid(globalData, xPos + Len(sc_FIN))
'
'            Select Case UCase(strDato)
'                Case "START":
'                Case "END": ws_ProcesoDataArraival
'                Case Else
'                    lTrama.Caption = strDato
'                    Dim mIDT As Integer
'                    If Trim(arrTramas(0)) = "" Then mIDT = 0 Else mIDT = UBound(arrTramas) + 1
'                    ReDim Preserve arrTramas(mIDT)
'                    arrTramas(mIDT) = strDato
'            End Select
'        Loop
'    End If
'
'End Sub

'Private Function ws_IniciarConexion() As Boolean
'
'    ws_IniciarConexion = True
'    lStatus.Caption = "Conectando ...": lStatus.Refresh
'
''    prmPortServer = 63259
''    prmIPServer = "192.168.1.89"
'    wsSocket.Connect prmIPServer, prmPortServer
'
'    Dim aQIntentos As Integer
'    aQIntentos = 1
'
'    Do While aQIntentos <= 4
'        DoEvents
'        If wsSocket.State = 7 Then Exit Do
'        Sleep 500
'        aQIntentos = aQIntentos + 1
'
'        lStatus.Caption = "Conectando ... (" & aQIntentos & ")"
'        lStatus.Refresh
'    Loop
'
'
'    If wsSocket.State <> 7 Then
'        lStatus.Caption = "Sin Conexión"
'        ws_IniciarConexion = False
'        tmCheckLog.Enabled = False
'    Else
'        lStatus.Caption = "Conectado"
'        ws_SendWhoIam
'        tmCheckLog.Enabled = True
'    End If
'
'    lStatus.Refresh
'
'End Function

Private Function ws_SendWhoIam()
        
    'IDUsr|NombreUsr|0000 (Info Sol Ser Gas Sus)
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    
    If paCodigoDeUsuario <> prmCurrentUser Then
        
        '1) Busco nivel de aprendizaje del usuario  ---------------------------------------------------------------------------------
        prmAutorizaCredHasta = -1
        
        Dim RsUsr As rdoResultset
        Cons = "Select UsuAutorizaCredHasta from Usuario Where UsuCodigo = " & paCodigoDeUsuario
        Set RsUsr = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsUsr.EOF Then
            If Not IsNull(RsUsr!UsuAutorizaCredHasta) Then prmAutorizaCredHasta = RsUsr!UsuAutorizaCredHasta
        End If
        RsUsr.Close
        '----------------------------------------------------------------------------------------------------------------------------------
        
        Dim mDatas As String
        'Cuando el Usuario Cambia ---> Sincronizo a Mano la primera vez
        If LCase(tmToBD.Tag) <> LCase("Conectado") Then
            CargoAsuntosPendientes
        End If
        
        Dim mNameUser As String
        If paCodigoDeUsuario <> 0 Then mNameUser = miConexion.UsuarioLogueado(Nombre:=True)
        
        prmCurrentUser = paCodigoDeUsuario
        prmCurrentUserName = mNameUser
    End If
    
End Function

'Public Function ws_Reconectar() As Boolean
'    On Error Resume Next
'    wsSocket.Close
'    'CargoParametrosLocales SoloServerAP:=True
'    ws_Reconectar = ws_IniciarConexion
'
'End Function

'Public Sub ws_SendData(Trama As String)
'
'    On Error Resume Next
'    If wsSocket.State = 7 Then
'        wsSocket.SendData Trama
'        DoEvents
'    End If
'
'End Sub

Private Function ws_ProcesoDataArraival()
    
    If Trim(arrTramas(0)) = "" Then Exit Function
    
Dim arrValue() As String
Dim mRow As Integer, mRowSel As Integer

    On Error GoTo errCargar
    Screen.MousePointer = 11
    QSinAnalizar = 0: QServicios = 0
    
    mRowSel = 1
    If vsAsunto.Rows > vsAsunto.FixedRows Then mRowSel = vsAsunto.Row
    
    '.IDUserPara = UsuarioPara
    '.IdTipo = TipoT
    '.IDEstadoTrama = Trim(EstadoT)
    '.DatosTrama = DatosT
    
    Dim bChange As Boolean: bChange = False
    Dim iT As Integer
    
    For iT = LBound(arrTramas) To UBound(arrTramas)
        arrValue = Split(arrTramas(iT), "|")
    
        Select Case Trim(arrValue(2))
            Case "E"        'Eliminado          -------------------------------------------------------
                    mRow = itm_FindItem(Val(arrValue(1)), Val(arrValue(3)))
                    If mRow <> -1 Then vsAsunto.RemoveItem mRow: bChange = True
                    
            Case "N"        'Nuevo o Modificada     ------------------------------------------------
            
                    bChange = True
                    Select Case Val(arrValue(1))
                        Case Asuntos.solicitudes
                            'Codigo|Proceso|Devuelta|Fecha|Cliente|CliCategoria|IDUsrR|NameUsrR|Estado|NameUsrS|Importe
                            '3 ..
                            itm_AddSolicitud Val(arrValue(3)), arrValue(6), Trim(arrValue(7)), Val(arrValue(8)), Trim(arrValue(13)), _
                                                    Val(arrValue(4)), Val(arrValue(11)), Val(arrValue(5)), Val(arrValue(9)), _
                                                    Trim(arrValue(10)), Trim(arrValue(12))
                            
                        Case Asuntos.Servicios
                            'Codigo|Producto|Tecnico|CostoT
                            If MnuServicio.Checked Then itm_AddServicio Val(arrValue(3)), Trim((arrValue(4))), Trim(arrValue(5)), Trim(arrValue(6))
                            
                        Case Asuntos.GastosAAutorizar
                            'Codigo|Proveedor|Importe
                            If Val(arrValue(0)) = paCodigoDeUsuario Then
                                itm_AddGasto Val(arrValue(3)), Trim((arrValue(4))), Trim(arrValue(5)), Trim(arrValue(6))
                            End If
                            
                        Case Asuntos.SucesosAAutorizar
                            If Val(arrValue(0)) = paCodigoDeUsuario Or Val(arrValue(0)) = 0 Then
                                'Codigo|NombreSuceso|Descripcion|Valor|Usuario
                                itm_AddSuceso Val(arrValue(3)), Trim((arrValue(4))), Trim(arrValue(5)), Trim(arrValue(6)), Trim(arrValue(7))
                            End If
                            
                    End Select
        End Select
        
        If vsAsunto.Rows > 1 Then
            vsAsunto.Select vsAsunto.FixedRows, vsAsunto.Cols - 1
            vsAsunto.Sort = flexSortGenericAscending
        End If
        
    Next
    
    ReDim arrTramas(0)
    
    If vsAsunto.Rows > 1 Then
        If Not (vsAsunto.Rows > mRowSel) Then mRowSel = vsAsunto.Rows - 1
        vsAsunto.Select mRowSel, 2
    End If
    
    If bChange Then
        With vsAsunto
            For iT = .FixedRows To .Rows - 1
                Select Case .Cell(flexcpData, iT, 0)
                
                    Case Asuntos.solicitudes        '0-TipoTrama    1-EstadoSol     5-UsuarioR
                        If Not (.Cell(flexcpData, iT, 5) <> 0 And .Cell(flexcpData, iT, 1) <> EstadoSolicitud.ParaRetomar) Then
                            QSinAnalizar = QSinAnalizar + 1
                        End If
                
                        
                    Case Asuntos.Servicios: QServicios = QServicios + 1
                End Select
            Next
        End With
    End If
    
    If QSinAnalizar > 0 And Not FormActivo("frmResolucion") Then fnc_ActivoSonido QSinAnalizar
    
    Dim mIcono As String
    mIcono = fnc_IconName(QSinAnalizar, QServicios)
    Me.Icon = Image1.ListImages(mIcono).Picture
    TrayModify Me.hwnd, Me, ""
    Me.Refresh
    
    Screen.MousePointer = 0
    Exit Function
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos en la lista.", Err.Description
    Screen.MousePointer = 0
End Function

Public Function acc_SiguienteAsunto(Optional mTipoAsunto As Integer = -1) As Integer
'   0- No hay
'   1- Hay y se actvo
'   -1- Error

On Error GoTo errFnc
    acc_SiguienteAsunto = 0
    If vsAsunto.Rows = vsAsunto.FixedRows Then Exit Function
    
    'Selecciono la primera q no se este analizando
    Dim I As Integer, bOK As Boolean
    
    bOK = False
    
    With vsAsunto
        For I = 1 To .Rows - 1
            
            If mTipoAsunto <> -1 Then
                If .Cell(flexcpData, I, 0) = mTipoAsunto Then
                    If mTipoAsunto = Asuntos.solicitudes Then
                        
                        If (IsDate(.Cell(flexcpText, I, 4))) Or _
                                    (.Cell(flexcpForeColor, I, 5) = Colores.osGris And UCase(Trim(.Cell(flexcpText, I, 5))) = UCase(prmCurrentUserName)) Then
                                    
                                bOK = True
                        End If
                    Else
                        bOK = True
                    End If
                End If
                
            Else
                If .Cell(flexcpData, I, 0) = Asuntos.solicitudes Then
                    If .Cell(flexcpData, I, 1) = EstadoSolicitud.Pendiente And IsDate(.Cell(flexcpText, I, 4)) Then bOK = True
                Else
                    bOK = True
                End If
            End If
            
            If bOK Then
                .Select I, 2: Exit For
            End If
            
        Next
    End With
    
    If bOK Then
        tmAuto.Tag = vsAsunto.Cell(flexcpData, vsAsunto.Row, 0)
        tmAuto.Enabled = True
        acc_SiguienteAsunto = 1
    End If
    
    Exit Function
    
errFnc:
    acc_SiguienteAsunto = -1
End Function

Private Sub Form_Resize()
    On Error Resume Next
    'lStatus.Top = 0: lStatus.Left = 60
    
    'With vsAsunto
    '    .Top = lStatus.Top + lStatus.Height
    '    .Left = 0
    '    .Width = Me.ScaleWidth - (vsAsunto.Left * 2)
    '    .Height = Me.ScaleHeight - vsAsunto.Top
    'End With
    

    With vsAsunto
        .Top = Me.ScaleTop
        .Left = 0
        .Width = Me.ScaleWidth - (vsAsunto.Left * 2)
        .Height = Me.ScaleHeight - lStatus.Height
    End With
    
    lStatus.Top = vsAsunto.Height
    lStatus.Left = Me.ScaleWidth - lStatus.Width - 60
        
    lTrama.Left = 0
    lTrama.Top = lStatus.Top
    lTrama.Width = lStatus.Left
    Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
'    wsSocket.Close
      
    TrayRemove Me.hwnd
    Me.Refresh
    
    GuardoSeteoForm Me
    GuardoSeteoControl Me.Name & MnuServicio.Name & "Checked", MnuServicio.Checked
        
    cBase.Close
    Set miConexion = Nothing
    Set clsGeneral = Nothing
        
    End
    
End Sub

Private Sub vsAsunto_DblClick()
    
    On Error GoTo errLogin
    If vsAsunto.Rows > 1 Then
        
        Select Case vsAsunto.Cell(flexcpData, vsAsunto.Row, 0)
            Case Asuntos.solicitudes: acc_ResolverSolicitud
            Case Asuntos.Servicios: acc_PresupuestarServicio
            Case Asuntos.GastosAAutorizar: acc_AutorizarGasto
            Case Asuntos.SucesosAAutorizar: acc_AutorizarSuceso
        End Select
        
    End If
    Exit Sub
    
errLogin:
    clsGeneral.OcurrioError "Error al acceder a los datos." & vbCrLf & "Verifique activación del login y vuelva a intentar.", Err.Description
    Set miConexion = New clsConexion
    Screen.MousePointer = 0
End Sub

Public Sub signalR_RefrescoSolicitudes()
    On Error Resume Next
    lStatus.Caption = "Buscando datos ...": lStatus.Refresh
    oHub.InvokeMethod "ObtenerAsuntosPendientes"
    lStatus.Caption = "": lStatus.Refresh
End Sub

Public Sub signalR_RefrescoSolicitudesResueltas()
    On Error Resume Next
    lStatus.Caption = "Buscando datos ...": lStatus.Refresh
    oHub.InvokeMethod "ObtenerSolicitudesResueltas"
    lStatus.Caption = "": lStatus.Refresh
End Sub

Private Sub acc_ResolverSolicitud()

    On Error GoTo errResolver
    
    If Not miConexion.AccesoAlMenu("Estudio de Solicitudes") Then Exit Sub
    ws_SendWhoIam
    prmIDSel = vsAsunto.Cell(flexcpValue, vsAsunto.Row, 0)
    
    Screen.MousePointer = 11
    
    Dim aUsrVisto As Long, retUsuarioR As Long
    aUsrVisto = 0: retUsuarioR = 0
    If vsAsunto.Cell(flexcpForeColor, vsAsunto.Row, 4) = Colores.osGris Then aUsrVisto = vsAsunto.Cell(flexcpData, vsAsunto.Row, 5)
    
    If FormActivo("frmResolucion") Then
        
        If frmResolucion.prmSolicitud <> 0 Then
            Screen.MousePointer = 0
            If MsgBox("La pantalla de Estudio de Solicitudes está activa. " & vbCrLf & _
                            "Si carga la nueva solicitud se perderán los datos anteriores." & vbCrLf & vbCrLf & _
                            "Confirma cargar la nueva solicitud.", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                Exit Sub
            End If
        End If
        
        
        Select Case fnc_BlockearSolicitud(prmIDSel, retUsuarioR) 'Si es 1 esta todo OK--------------------------
            Case 0  'OTRO USUARIO
                If retUsuarioR <> paCodigoDeUsuario And aUsrVisto = 0 Then
                    Screen.MousePointer = 0
                    If MsgBox("La solicitud se está analizando por otro usuario." & vbCrLf & _
                                   "Desea visualizarla.", vbExclamation + vbYesNo, "Solicitud en Proceso") = vbNo Then Exit Sub
                End If
           
            Case -1 'ERROR o FUE RESUELTA
                Screen.MousePointer = 0
                MsgBox "Posiblemente la solicitud ya fue resuelta (o ha sido eliminada).", vbExclamation, "Cambiaron los Datos"
                signalR_RefrescoSolicitudes
                Exit Sub
        End Select  '----------------------------------------------------------------------------------------
        signalR_RefrescoSolicitudes
        Screen.MousePointer = 11
        frmResolucion.prmVistoPor = aUsrVisto
        frmResolucion.prmSolicitud = prmIDSel
        frmResolucion.BuscoSolicitud prmIDSel
        If frmResolucion.WindowState = vbMinimized Then frmResolucion.WindowState = vbNormal
        frmResolucion.SetFocus
                
    Else
        Select Case fnc_BlockearSolicitud(prmIDSel, retUsuarioR) 'Si es 1 esta todo OK--------------------------
            Case 0  'OTRO USUARIO
                If retUsuarioR <> paCodigoDeUsuario And aUsrVisto = 0 Then
                    Screen.MousePointer = 0
                    If MsgBox("La solicitud se está analizando por otro usuario." & _
                                   "Desea visualizarla.", vbExclamation + vbYesNo, "Solicitud en Proceso") = vbNo Then Exit Sub
                End If
                
            Case -1 'ERROR
                MsgBox "Posiblemente la solicitud ya fue resuelta.", vbExclamation, "Solicitud Resuelta"
                signalR_RefrescoSolicitudes
                Exit Sub
        End Select  '----------------------------------------------------------------------------------------
        signalR_RefrescoSolicitudes
        Screen.MousePointer = 11
        frmResolucion.prmVistoPor = aUsrVisto
        frmResolucion.prmSolicitud = prmIDSel
        frmResolucion.Show vbModeless
    End If

errResolver:
End Sub

Private Sub acc_AutorizarAsunto(Optional bAutorizado As Boolean = True)
    
    If vsAsunto.Rows = 1 Then Exit Sub
    On Error GoTo errMnu
    
    Dim arrData() As String
    Dim mIDAsunto As Long, mTipoAsunto As Integer, mCliente As String
    
    arrData = Split(MnuAutorizar.Tag, ":")
    
    mTipoAsunto = Val(arrData(0))
    mIDAsunto = Val(arrData(1))
    mCliente = vsAsunto.Cell(flexcpText, vsAsunto.Row, 1)
    
    Dim mSQL As String, rs2 As rdoResultset
    
    Select Case mTipoAsunto
        Case Asuntos.solicitudes
                If MsgBox("Confirma resolver el crédito con la condición estándar.", vbQuestion + vbYesNo, "Autorizar Crédito") = vbNo Then Exit Sub
                fnc_AutorizarCredito mIDAsunto
                On Error Resume Next
                signalR_RefrescoSolicitudes
        
        Case Asuntos.GastosAAutorizar
                If MsgBox("Confirma " & IIf(bAutorizado, "", "NO ") & "Autorizar el Gasto ID " & mIDAsunto & " ?." & vbCrLf & _
                               "Proveedor " & mCliente, vbQuestion + vbYesNo, "Autorizar Gasto") = vbNo Then Exit Sub
                        
                'mSQL = "Select * from Compra Where ComCodigo = " & mIDAsunto
                'Set rs2 = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
                'If Not rs2.EOF Then
                '    rs2.Edit
                '    rs2!ComFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
                '    rs2!ComVerificado = IIf(bAutorizado, 1, 0)
                '    rs2.Update
                'End If
                'rs2.Close
                
                'Estado 0 = Pendiente, 1 = Autorizado, 2 = No Autorizado
                mSQL = "Select * from ZureoCGSA.dbo.cceComprobantes Where ComID = " & mIDAsunto
                Set rs2 = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
                If Not rs2.EOF Then
                    rs2.Edit
                    rs2("ComEstado").value = IIf(bAutorizado, 1, 2)
                    rs2("ComFechaEstado").value = Now 'Format(Now, prmBDDateFormat)
                    rs2.Update
                End If
                rs2.Close

                
                mValor = itm_FindItem(mTipoAsunto, mIDAsunto)
                vsAsunto.RemoveItem mValor
                
        Case Asuntos.SucesosAAutorizar
                If bAutorizado Then
                    If MsgBox("Confirma Autorizar el Suceso ID " & mIDAsunto & " ?." & vbCrLf & _
                                "Detalle: " & mCliente, vbQuestion + vbYesNo, "Autorizar Suceso") = vbNo Then Exit Sub
                Else
                    If MsgBox("Confirma NO Autorizar el Suceso ID " & mIDAsunto & " ?." & vbCrLf & _
                                "Detalle: " & mCliente, vbQuestion + vbYesNo, "NO Autorizar Suceso") = vbNo Then Exit Sub
                End If
                
                mSQL = "Select * from Suceso Where SucCodigo = " & mIDAsunto
                Set rs2 = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
                If Not rs2.EOF Then
                    rs2.Edit
                    If rs2("SucAutoriza") = 0 Then rs2("SucAutoriza") = paCodigoDeUsuario
                    rs2!SucVerificado = IIf(bAutorizado, 1, 9)
                    rs2.Update
                End If
                rs2.Close
                
                mValor = itm_FindItem(mTipoAsunto, mIDAsunto)
                vsAsunto.RemoveItem mValor
                
    End Select
    
    Exit Sub
errMnu:
    clsGeneral.OcurrioError "Error al autorizar el asunto pendiente.", Err.Description
End Sub

Private Sub acc_PresupuestarServicio()

    If Not miConexion.AccesoAlMenu("Validacion de Presupuesto") Then Exit Sub
    prmIDSel = vsAsunto.Cell(flexcpValue, vsAsunto.Row, 0)
    
    EjecutarApp pathApp & "Validacion de Presupuesto", CStr(prmIDSel)

End Sub

Private Sub acc_AutorizarGasto()
On Error GoTo errFnc

Dim mApp As String

    Screen.MousePointer = 11
    prmIDSel = vsAsunto.Cell(flexcpValue, vsAsunto.Row, 0)
    
    Select Case z_TipoDeGasto(prmIDSel)
        Case 1: mApp = "Ingreso de Gastos.exe"
        Case 2: mApp = "Compra de Mercaderia.exe"
        Case 3: mApp = "Ingreso de Facturas.exe"
    End Select
        
    If mApp <> "" Then EjecutarApp pathApp & mApp, CStr(prmIDSel)
    
    Screen.MousePointer = 0
    Exit Sub
errFnc:
    clsGeneral.OcurrioError "Error al activar el gasto seleccioando.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub acc_AutorizarSuceso()

    If Not miConexion.AccesoAlMenu("Sucesos") Then Exit Sub
    prmIDSel = vsAsunto.Cell(flexcpValue, vsAsunto.Row, 0)
    
    EjecutarApp pathApp & "appExploreMsg.exe", CStr(prmPlantillaASucesos) & ":" & CStr(prmIDSel)

End Sub

Private Sub CargoAsuntosPendientes()
Dim mRowSel As Long

    On Error GoTo errCargar
    Screen.MousePointer = 11
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
    prmCurrentUser = -1     'xq no hay conexion
    QSinAnalizar = 0: QServicios = 0
    mRowSel = 1
    If vsAsunto.Rows > vsAsunto.FixedRows Then mRowSel = vsAsunto.Row
    vsAsunto.Rows = 1
    
    QSinAnalizar = loc_CargoSolicitudes
    If MnuServicio.Checked Then QServicios = loc_CargoServicios
    
    If paCodigoDeUsuario > 0 Then loc_CargoSucesosAAutorizar
    If paCodigoDeUsuario > 0 Then loc_CargoGastosAAutorizar
    
    If vsAsunto.Rows > 1 Then
        vsAsunto.Select vsAsunto.FixedRows, vsAsunto.Cols - 1
        vsAsunto.Sort = flexSortGenericAscending
    
        If Not (vsAsunto.Rows > mRowSel) Then mRowSel = vsAsunto.Rows - 1
        vsAsunto.Select mRowSel, 2
    End If
    
    If QSinAnalizar > 0 And Not FormActivo("frmResolucion") Then fnc_ActivoSonido QSinAnalizar
    
    Dim mIcono As String
    mIcono = fnc_IconName(QSinAnalizar, QServicios)
    Me.Icon = Image1.ListImages(mIcono).Picture
    TrayModify Me.hwnd, Me, ""
    Me.Refresh
    
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos en la lista.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function loc_CargoSolicitudes() As Integer

Dim mQAux As Integer

Dim vImporteSol As String, vDevuelta As Integer, vCliCategoria As Integer, vEstado As Integer
Dim vIDUsrR As Long, vNameUsrR As String, vNameUsrS As String

    On Error GoTo errCargar
    Screen.MousePointer = 11
    mQAux = 0
    
    'Solicitudes Poceso de Resolucion = Manual Y Estado = Pendiente
    Cons = "SELECT SolCodigo, SolFecha, SolDevuelta, SolProceso, SolEstado,  CliCiRuc, CliCategoria, " & _
                 " IsNull(RTrim(CEmNombre), RTrim(CPeApellido1) + RTrim(' ' + IsNull(CPeApellido2,''))+', ' + RTrim(CPeNombre1) + RTrim(' ' + IsNull(CPeNombre2,''))) Nombre, " & _
                 " RTrim(UsrSol.UsuIdentificacion) Solicitante, RTrim(UsrRes.UsuIdentificacion) Resolv, UsrRes.UsuCodigo as IDResolv, " & _
                 " Sum((TCuCantidad + (Convert(bit, TCuVencimientoC)-1)) * RSoValorCuota - (Convert(bit, IsNull(TCuVencimientoE,0)) * IsNull(RSoValorEntrega,0))) Monto " & _
           " FROM Solicitud  " & _
                " INNER JOIN RenglonSolicitud on SolCodigo = RSoSolicitud  " & _
                " INNER JOIN TipoCuota on RSoTipoCuota = TCuCodigo  " & _
                " INNER JOIN Cliente on SolCliente = CliCodigo  " & _
                " INNER JOIN Usuario UsrSol ON SolUsuarioS = UsrSol.UsuCodigo  " & _
                " LEFT OUTER JOIN Usuario UsrRes ON SolUsuarioR = UsrRes.UsuCodigo  " & _
                " LEFT OUTER JOIN CPersona ON CliCodigo = CPeCliente  " & _
                " LEFT OUTER JOIN CEmpresa ON CliCodigo = CEmCliente " & _
           " Where SolFecha Between '" & Format(gFechaServidor, "mm/dd/yyyy 00:00") & "' And  '" & Format(gFechaServidor, "mm/dd/yyyy 23:59") & "'" & _
           " And SolProceso IN (" & TipoResolucionSolicitud.Manual & ", " & TipoResolucionSolicitud.LlamarA & ")" & _
           " And SolEstado IN (" & EstadoSolicitud.Pendiente & ", " & EstadoSolicitud.ParaRetomar & ")" & _
           " GROUP BY SolCodigo, SolFecha, SolDevuelta, SolProceso, SolEstado, CliCiRuc, CliCategoria, CEmNombre, CPeApellido1, CPeApellido2,  CPeNombre1, CPeNombre2, UsrSol.UsuIdentificacion, UsrRes.UsuIdentificacion, UsrRes.UsuCodigo " & _
           " ORDER BY SolCodigo ASC "
 
    Set rsQry = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsQry.EOF
        
        vImporteSol = Format(rsQry!Monto, "#,##0.00")
        
        vDevuelta = 0
        If Not IsNull(rsQry!SolDevuelta) Then If rsQry!SolDevuelta Then vDevuelta = 1
        
        If Not IsNull(rsQry!CliCategoria) Then vCliCategoria = rsQry!CliCategoria Else vCliCategoria = 0
        
        If Not IsNull(rsQry!Solicitante) Then vNameUsrS = rsQry!Solicitante Else vNameUsrS = ""
        If Not IsNull(rsQry!Resolv) Then vNameUsrR = rsQry!Resolv Else vNameUsrR = ""
        If Not IsNull(rsQry!IDResolv) Then vIDUsrR = rsQry!IDResolv Else vIDUsrR = 0
        
        If Not IsNull(rsQry!SolEstado) Then vEstado = rsQry!SolEstado Else vEstado = 0
                    
        If itm_AddSolicitud(rsQry!SolCodigo, Format(rsQry!SolFecha, "dd/mm/yyyy hh:mm:ss"), Trim(rsQry!Nombre), vCliCategoria, vImporteSol, _
                                    rsQry!SolProceso, vEstado, vDevuelta, vIDUsrR, vNameUsrR, vNameUsrS) Then
                                    
                    mQAux = mQAux + 1
        End If
        
        rsQry.MoveNext
    Loop
    rsQry.Close
    
    loc_CargoSolicitudes = mQAux
    
    Screen.MousePointer = 0
    Exit Function
    
errCargar:
    Me.Caption = "(Error al cargar las solicitudes)"
End Function

'Private Function loc_CargoSolicitudes_BKUP() As Integer
'
'Dim mQAux As Integer
'
'Dim vImporteSol As String, vDevuelta As Integer, vCliCategoria As Integer, vEstado As Integer
'Dim vIDUsrR As Long, vNameUsrR As String, vNameUsrS As String
'
'    On Error GoTo errCargar
'    Screen.MousePointer = 11
'    mQAux = 0
'
'    'Solicitudes Poceso de Resolucion = Manual Y Estado = Pendiente
'    Cons = "Select Solicitud.*, CliCategoria, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2), Tipo = 1" _
'           & " From Solicitud, CPersona, Cliente " _
'           & " Where SolFecha Between '" & Format(gFechaServidor, "mm/dd/yyyy 00:00") & "' And  '" & Format(gFechaServidor, "mm/dd/yyyy 23:59") & "'" _
'           & " And SolProceso IN (" & TipoResolucionSolicitud.Manual & ", " & TipoResolucionSolicitud.LlamarA & ")" _
'           & " And SolEstado IN (" & EstadoSolicitud.Pendiente & ", " & EstadoSolicitud.ParaRetomar & ")" _
'           & " And CPeCliente = SolCliente And CPeCliente = CliCodigo" _
'                                                & " UNION " _
'           & " Select Solicitud.*, CliCategoria, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')'), Tipo = 2" _
'           & " From Solicitud, CEmpresa, Cliente " _
'           & " Where SolFecha Between '" & Format(gFechaServidor, "mm/dd/yyyy 00:00") & "' And  '" & Format(gFechaServidor, "mm/dd/yyyy 23:59") & "'" _
'           & " And SolProceso IN (" & TipoResolucionSolicitud.Manual & ", " & TipoResolucionSolicitud.LlamarA & ")" _
'           & " And SolEstado IN (" & EstadoSolicitud.Pendiente & ", " & EstadoSolicitud.ParaRetomar & ")" _
'           & " And CEmCliente = SolCliente And CEmCliente = CliCodigo"
'
'    Set rsQry = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'    Do While Not rsQry.EOF
'
'        vImporteSol = fnc_ItemImportes(rsQry!SolCodigo)
'
'        vDevuelta = 0
'        If Not IsNull(rsQry!SolDevuelta) Then
'            If rsQry!SolDevuelta Then vDevuelta = 1
'        End If
'
'        If Not IsNull(rsQry!CliCategoria) Then vCliCategoria = rsQry!CliCategoria Else vCliCategoria = 0
'
'        vIDUsrR = 0: vNameUsrR = ""
'        If Not IsNull(rsQry!SolUsuarioR) Then
'            vIDUsrR = rsQry!SolUsuarioR
'            vNameUsrR = fnc_ItemUsuarios(vIDUsrR)
'        End If
'
'        If Not IsNull(rsQry!SolEstado) Then vEstado = rsQry!SolEstado Else vEstado = 0
'        If Not IsNull(rsQry!SolUsuarioS) Then vNameUsrS = fnc_ItemUsuarios(rsQry!SolUsuarioS) Else vNameUsrS = ""
'
'        If itm_AddSolicitud(rsQry!SolCodigo, Format(rsQry!SolFecha, "dd/mm/yyyy hh:mm:ss"), Trim(rsQry!Nombre), vCliCategoria, vImporteSol, _
'                                    rsQry!SolProceso, vEstado, vDevuelta, vIDUsrR, vNameUsrR, vNameUsrS) Then
'
'                    mQAux = mQAux + 1
'        End If
'
'        rsQry.MoveNext
'    Loop
'    rsQry.Close
'
'    z_VerificoColImportes
'
'    loc_CargoSolicitudes_BKUP = mQAux
'
'    Screen.MousePointer = 0
'    Exit Function
'
'errCargar:
'    Me.Caption = "(Error al cargar las solicitudes)"
'End Function

'   Carga los servicios pendientes de presupuestacion
Private Function loc_CargoServicios() As Integer

Dim QPendientes As Integer
Dim vTecnico As String

    On Error GoTo errServicio
    QPendientes = 0
    
    Cons = "Select SerCodigo, TalCostoTecnico,  UsuDigito, UsuIdentificacion" & _
                " From Servicio, Taller " & _
                        " Left Outer Join Usuario on TalTecnico = UsuCodigo" & _
                " Where SerCodigo = TalServicio " & _
                " And TalFPresupuesto is Not Null " & _
                " And SerCostoFinal is null"
    
    Set rsQry = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsQry.EOF
        
        vTecnico = ""
        If Not IsNull(rsQry!UsuIdentificacion) Then vTecnico = Trim(rsQry!UsuIdentificacion)
        
        If itm_AddServicio(rsQry!SerCodigo, "", vTecnico, Format(rsQry!TalCostoTecnico, "#,##0.00")) Then
            QPendientes = QPendientes + 1
        End If
        
        rsQry.MoveNext
    Loop
    rsQry.Close
    
    loc_CargoServicios = QPendientes
    Exit Function
    
errServicio:
    Me.Caption = "(Error al cargar servicios)"
End Function

Private Sub vsAsunto_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        'Unload Me
        Unload frmResolucion
        Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn And vsAsunto.Rows > 1 Then
        Call vsAsunto_DblClick
    End If
    
End Sub

Private Sub vsAsunto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errOp

    If Button = vbRightButton Then
        MnuNOAutorizar.Enabled = False
        MnuAutorizarAll.Enabled = False
        MnuAutorizar.Enabled = False
        MnuAutorizar.Tag = ""
        
        MnuAnalizar.Enabled = False
        MnuAnalizarP.Enabled = False
        MnuFiltrarProveedor.Enabled = False
        MnuQuitarFiltros.Enabled = False
        MnuLiberarSolicitud.Enabled = False
        
        If vsAsunto.MouseRow > 0 Then
        
            If vsAsunto.Rows > vsAsunto.FixedRows Then
                With vsAsunto
                    .Select .MouseRow, .MouseCol
                    
                    Dim mIDAsunto As Long, mTipoAsunto As Integer
                    mIDAsunto = .Cell(flexcpText, .Row, 0)
                    mTipoAsunto = .Cell(flexcpData, .Row, 0)
                    
                    MnuAutorizar.Tag = mTipoAsunto & ":" & mIDAsunto
                    
                    Select Case mTipoAsunto
                        Case Asuntos.solicitudes
                            'Si la categotria del cliente es distribuidor (data de la 4) activo el menu
                            If paCodigoDeUsuario <> 0 Then
                                Dim miCat As String
                                miCat = "," & .Cell(flexcpData, .Row, 4) & ","
                                
                                If InStr(paCatsDistribuidor, miCat) <> 0 Then
                                    If miConexion.AccesoAlMenu("Estudio de Solicitudes") Then MnuAutorizar.Enabled = True
                                End If
                                MnuLiberarSolicitud.Enabled = True
                            End If
                    
                        Case Asuntos.GastosAAutorizar, Asuntos.SucesosAAutorizar
                                MnuAutorizar.Enabled = True
                                MnuAnalizar.Enabled = (mTipoAsunto = Asuntos.GastosAAutorizar)
                                MnuAnalizarP.Enabled = MnuAnalizar.Enabled
                                
                                MnuAutorizarAll = (mTipoAsunto = Asuntos.GastosAAutorizar)
                                'MnuNOAutorizar.Enabled = (mTipoAsunto = Asuntos.GastosAAutorizar)
                                MnuNOAutorizar.Enabled = True
                                
                                MnuFiltrarProveedor.Enabled = (mTipoAsunto = Asuntos.GastosAAutorizar)
                                MnuQuitarFiltros.Enabled = MnuFiltrarProveedor.Enabled
                    End Select
                End With
                
            End If
            
        End If
        PopupMenu MnuFiltro
    End If
    
errOp:
End Sub

Private Sub MnuAutorizar_Click()
    If Trim(MnuAutorizar.Tag) <> "" Then acc_AutorizarAsunto
End Sub

Private Sub MnuServicio_Click()
    With MnuServicio
        If .Checked Then .Checked = False Else .Checked = True
        prmCurrentUser = -1
        ws_SendWhoIam
    End With
End Sub

Private Sub MnuTAsuntos_Click()
    On Error Resume Next
    frmLista.Show
    MnuTAsuntos.Enabled = False: MnuTOcultar.Enabled = True
End Sub

Private Sub MnuTEnd_Click()
    
    On Error Resume Next
    Unload Me
    
End Sub

Private Sub MnuTOcultar_Click()
    On Error Resume Next
    frmLista.Hide
    MnuTAsuntos.Enabled = True: MnuTOcultar.Enabled = False
End Sub

Private Sub MnuTsiguiente_Click()

    Screen.MousePointer = 11
    
    If MnuTray.Visible Then MnuTray.Visible = False
    acc_SiguienteAsunto
    
    Screen.MousePointer = 0
End Sub

Private Sub z_InicializoGrilla()

    With vsAsunto
        .Rows = 1: .Cols = 1
        .FormatString = "<Código|Realizada|<Cliente|>Importe|^Analizando|<Solicitada Por|Orden"
        .ColWidth(0) = 900: .ColWidth(2) = 5000: .ColWidth(3) = 1200: .ColWidth(4) = 1000
        .ColWidth(1) = 0
        .ColHidden(.Cols - 1) = True
        .WordWrap = False
        .MergeCells = flexMergeSpill: .ExtendLastCol = True
        
        .Editable = True
    End With
    
End Sub

Private Sub tmAuto_Timer()

    'Timmer que activa el Siguiente Asunto (x el menu)
    tmAuto.Enabled = False
    Select Case Val(tmAuto.Tag)
        Case Asuntos.solicitudes: acc_ResolverSolicitud
        Case Asuntos.Servicios: acc_PresupuestarServicio
        Case Asuntos.GastosAAutorizar: acc_AutorizarGasto
        Case Asuntos.SucesosAAutorizar: acc_AutorizarSuceso
    End Select
    
End Sub


Private Function loc_CargoGastosAAutorizar() As Integer

Dim mQPendientes As Integer

    On Error GoTo errServicio
    mQPendientes = 0
    
    Cons = "Select ComCodigo, ComFecha, ComImporte, IsNull(ComIva, 0) as ComIva, IsNull(ComCofis,0) as ComCofis, MonSigno, PClFantasia" & _
                " From Compra, Moneda, ProveedorCliente " & _
                " Where ComUsrAutoriza = " & paCodigoDeUsuario & _
                " And ComVerificado IS NULL " & _
                " And ComMoneda = MonCodigo And ComProveedor = PClCodigo"
    
    Set rsQry = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsQry.EOF
    
        itm_AddGasto rsQry!ComCodigo, Trim(rsQry!PClFantasia), _
                            Trim(rsQry!MonSigno) & " " & Format(rsQry!ComImporte + rsQry!ComIVA + rsQry!ComCoFIS, "#,##0.00"), ""
                           
        mQPendientes = mQPendientes + 1
        rsQry.MoveNext
    Loop
    rsQry.Close
    
    loc_CargoGastosAAutorizar = mQPendientes
    Exit Function
    
errServicio:
    Me.Caption = "(Error al cargar Gastos)"
End Function

Private Function loc_CargoSucesosAAutorizar() As Integer

Dim mQPendientes As Integer

    On Error GoTo errServicio
    mQPendientes = 0
    
    '2/8/2012 agregué condición Autoriza = 0 para sucesos a autorizar globales.
    Cons = "Select SucCodigo, TSuNombre ,isNull(SucDescripcion, '') as SucDescripcion, SucValor, SucUsuario" & _
                " From Suceso, TipoSuceso " & _
                " Where SucTipo *= TSuCodigoSistema " & _
                " And SucAutoriza IN(0, " & paCodigoDeUsuario & ")" & _
                " And IsNull(SucVerificado, 0) = 0 "
                
    Set rsQry = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsQry.EOF
        
        itm_AddSuceso rsQry!SucCodigo, Trim(rsQry!TSuNombre), Trim(rsQry!SucDescripcion), Format(rsQry!SucValor, "#,##0.00"), ""
            
        mQPendientes = mQPendientes + 1
        
        rsQry.MoveNext
    Loop
    rsQry.Close
    
    loc_CargoSucesosAAutorizar = mQPendientes
    Exit Function
    
errServicio:
    Me.Caption = "(Error al cargar los Sucesos)"
End Function

Private Function z_VerificoColImportes()

    On Error GoTo errVCol
    Dim aItem As String, J As Integer, bEsta As Boolean
    Dim aIdx As Integer
    
    aIdx = 1
    For I = 1 To colImportes.Count
        aItem = Mid(colImportes(aIdx), 1, InStr(colImportes(aIdx), "|") - 1)
        bEsta = False
        
        For J = 1 To vsAsunto.Rows - 1
            If aItem = Trim(vsAsunto.Cell(flexcpText, J, 0)) Then bEsta = True: Exit For
        Next
        
        If Not bEsta Then colImportes.Remove aIdx Else aIdx = aIdx + 1
    Next

errVCol:
End Function

Private Function itm_FindItem(mFTipo As Integer, mFID As Long) As Integer

    Dim iFnd As Integer
    itm_FindItem = -1
    
    With vsAsunto
        For iFnd = .FixedRows To .Rows - 1
            If .Cell(flexcpData, iFnd, 0) = mFTipo And .Cell(flexcpValue, iFnd, 0) = mFID Then
                itm_FindItem = iFnd
                Exit For
            End If
        Next
    End With
    
End Function

Private Function itm_AddSolicitud(mCodigo As Long, mFecha As String, mCliente As String, mCliCategoria As Integer, mImporte As String, _
                                                mProceso As Integer, mEstado As Integer, mDevuelta As Integer, _
                                                mIDUsrR As Long, mNameUsrR As String, mNameUsrS As String) As Boolean

'   Retorna true si la solicitud esta sin analizar
Dim mTexto  As String
Dim mIdRow As Integer

    itm_AddSolicitud = True
    mIdRow = itm_FindItem(Asuntos.solicitudes, mCodigo)
    
'    'Este punto lo agregamos para Supremo no deja visualizar las solicitudes que tiene CGSA por 1 mínuto.
'    If mIdRow = -1 And mIDUsrR = 23 And DateDiff("s", Now, mFecha) > -60 And mDevuelta <> 1 Then
'        Exit Function
'    End If
    
    With vsAsunto
        
        If mIdRow = -1 Then
            .AddItem ""
            mIdRow = .Rows - 1
        End If
        
        .Cell(flexcpText, mIdRow, 0) = mCodigo
        mValor = Asuntos.solicitudes: .Cell(flexcpData, mIdRow, 0) = mValor               'Data del 0- Tipo de Fila

        mTexto = "solicitud"
        If mProceso = TipoResolucionSolicitud.LlamarA Then mTexto = "llamara"
        .Cell(flexcpPicture, mIdRow, 0) = Image1.ListImages(mTexto).ExtractIcon

        If mDevuelta = 1 Then
            .Cell(flexcpBackColor, mIdRow, 0, , .Cols - 1) = Colores.Obligatorio
            .Cell(flexcpFontItalic, mIdRow, 4) = True
        End If

        .Cell(flexcpText, mIdRow, 1) = Format(mFecha, "dd/mm/yyyy hh:mm:ss")
        mValor = Val(mEstado): .Cell(flexcpData, mIdRow, 1) = mValor                                'Data del 1- Estado Solicitud

        .Cell(flexcpText, mIdRow, 2) = Trim(mCliente)

        .Cell(flexcpText, mIdRow, 3) = mImporte
               
         mValor = mCliCategoria                            'Data del 4 Categoria de Cliente
        .Cell(flexcpData, mIdRow, 4) = mValor
        .Cell(flexcpForeColor, mIdRow, 4) = .ForeColor
        
        If mIDUsrR <> 0 Then
            mValor = mIDUsrR
            .Cell(flexcpData, mIdRow, 5) = mValor    'En el data 5 SolUsuarioR
            
            .Cell(flexcpText, mIdRow, 4) = mNameUsrR
            If mEstado = EstadoSolicitud.ParaRetomar Then
                .Cell(flexcpForeColor, mIdRow, 4) = Colores.osGris
            Else
                itm_AddSolicitud = False
            End If
            .Cell(flexcpFontBold, mIdRow, 4) = True
            
        Else
            .Cell(flexcpData, mIdRow, 5) = 0
            .Cell(flexcpText, mIdRow, 4) = "": .Cell(flexcpFontBold, mIdRow, 4) = False
            '.Cell(flexcpForeColor, mIdRow, 4) = .CellForeColor
        End If
        
        .Cell(flexcpText, mIdRow, 5) = Trim(mNameUsrS)
        .Cell(flexcpText, mIdRow, .Cols - 1) = 1
        
        If Trim(.Cell(flexcpText, mIdRow, 4)) = "" Then
            .Cell(flexcpText, mIdRow, 4) = z_DifTime(CDate(mFecha))
            '.Cell(flexcpForeColor, mIdRow, 4) = .CellForeColor
        End If
        
    End With
            
End Function


Private Function itm_AddServicio(mCodigo As Long, mProducto As String, mTecnico As String, mCostoT As String) As Boolean

Dim mIdRow As Integer

    itm_AddServicio = True
    mIdRow = itm_FindItem(Asuntos.Servicios, mCodigo)
    
    With vsAsunto
        If mIdRow = -1 Then
            .AddItem ""
            mIdRow = .Rows - 1
        End If
        
        .Cell(flexcpText, mIdRow, 0) = mCodigo
        mValor = Asuntos.Servicios: .Cell(flexcpData, mIdRow, 0) = mValor               'Data del 0 Tipo de Fila
                
        .Cell(flexcpPicture, mIdRow, 0) = Image1.ListImages("servicio").ExtractIcon
        
        .Cell(flexcpText, mIdRow, 1) = Trim(mProducto)
            
        .Cell(flexcpText, mIdRow, 3) = Trim(mCostoT)
        .Cell(flexcpText, mIdRow, 5) = Trim(mTecnico)
        
        .Cell(flexcpText, mIdRow, .Cols - 1) = 2
    End With
            
End Function

Private Function itm_AddGasto(mCodigo As Long, mProveedor As String, mImporte As String, mUsuarioG As String) As Boolean

Dim mIdRow As Integer

    itm_AddGasto = True
    mIdRow = itm_FindItem(Asuntos.GastosAAutorizar, mCodigo)
    
    With vsAsunto
        If mIdRow = -1 Then
            .AddItem ""
            mIdRow = .Rows - 1
        End If

        .Cell(flexcpText, mIdRow, 0) = mCodigo
        .Cell(flexcpChecked, mIdRow, 4) = flexUnchecked
        mValor = Asuntos.GastosAAutorizar: .Cell(flexcpData, mIdRow, 0) = mValor               'Data del 0 Tipo de Fila
                
        .Cell(flexcpPicture, mIdRow, 0) = Image1.ListImages("gastos").ExtractIcon
        
        .Cell(flexcpText, mIdRow, 1) = Trim(mProveedor)
        
        .Cell(flexcpText, mIdRow, 3) = mImporte
        
        .Cell(flexcpText, mIdRow, 5) = mUsuarioG
        
        .Cell(flexcpText, mIdRow, .Cols - 1) = 4
        .Cell(flexcpForeColor, mIdRow, 0, , .Cols - 1) = Colores.osGris
        
    End With
            
End Function

Private Function itm_AddSuceso(mCodigo As Long, mSNombre As String, mSDescripcion As String, mSValor As String, mUsuario As String) As Boolean

Dim mIdRow As Integer

    itm_AddSuceso = True
    mIdRow = itm_FindItem(Asuntos.SucesosAAutorizar, mCodigo)
    
    With vsAsunto
        If mIdRow = -1 Then
            .AddItem ""
            mIdRow = .Rows - 1
        End If
        
        .Cell(flexcpText, mIdRow, 0) = mCodigo
        mValor = Asuntos.SucesosAAutorizar: .Cell(flexcpData, mIdRow, 0) = mValor               'Data del 0 Tipo de Fila
                
        .Cell(flexcpPicture, mIdRow, 0) = Image1.ListImages("sucesos").ExtractIcon
        
        .Cell(flexcpText, mIdRow, 1) = Trim(mSNombre) & " | " & Trim(mSDescripcion)
        .Cell(flexcpText, mIdRow, 3) = mSValor
        
        .Cell(flexcpText, mIdRow, 5) = mUsuario
        
        .Cell(flexcpText, mIdRow, .Cols - 1) = 3
        .Cell(flexcpForeColor, mIdRow, 0, , .Cols - 1) = &H8080FF
    End With
            
End Function

Private Function z_DifTime(mTDesde As Date) As String

    z_DifTime = Format(Now - mTDesde, "nn:ss")
    
End Function

Private Function z_TipoDeGasto(xIDGasto As Long) As Integer
On Error GoTo errFnc
Dim rs2 As rdoResultset

    z_TipoDeGasto = 0
    
    'Los Gastos pueden venir por compras, importaciones o gastos generales
    'X eso Consulto
    Cons = "Select Top 1 ComCodigo, CReCompra, GImIDCompra " & _
            " From Compra " & _
            " Left Outer Join CompraRenglon On CReCompra = ComCodigo" & _
            " Left Outer Join GastoImportacion On GImIDCompra = ComCodigo" & _
            " Where ComCodigo = " & xIDGasto
    
    Set rs2 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rs2.EOF Then
        If Not IsNull(rs2!GImIDCompra) Then
            z_TipoDeGasto = 1       'Ingreso de Gastos.exe
        ElseIf Not IsNull(rs2!CReCompra) Then
            z_TipoDeGasto = 2       'Compra de Mercaderia.exe
        Else
            z_TipoDeGasto = 3       'Ingreso de Facturas.exe
        End If
    End If
    rs2.Close
    
errFnc:
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

Private Function ConectarSignalR() As Boolean
    tmToBD.Tag = ""
    lStatus.Caption = "Conectando": lStatus.Refresh
    Set oHub = New ClientHub
    oHub.AddMethodCallBack "CargarAsuntosPendientes"
    oHub.AddMethodCallBack "AsuntosPendientesServerON"
'    ConectarSignalR = oHub.ConnectHub("http://192.168.111.11:8082/signalr", "AsuntosPendientesHub")
'    Exit Function
    ConectarSignalR = oHub.ConnectHub(prmHUBURL, "AsuntosPendientesHub")
    tmToBD.Tag = IIf(ConectarSignalR, "conectado", "")
End Function
