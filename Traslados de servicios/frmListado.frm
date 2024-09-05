VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form frmListado 
   Caption         =   "Listado de Traslados de Servicios"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4455
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7858
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
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
      BackColorSel    =   13686989
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      SubtotalPosition=   0
      OutlineBar      =   1
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
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4455
      Left            =   60
      TabIndex        =   22
      Top             =   1200
      Width           =   7335
      _Version        =   196608
      _ExtentX        =   12938
      _ExtentY        =   7858
      _StockProps     =   229
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      PreviewMode     =   1
      Zoom            =   70
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   50
      ScaleHeight     =   435
      ScaleWidth      =   6075
      TabIndex        =   23
      Top             =   5880
      Width           =   6135
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4140
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bGrabar 
         Height          =   310
         Left            =   3420
         Picture         =   "frmListado.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Grabar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2760
         Picture         =   "frmListado.frx":0B78
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2400
         Picture         =   "frmListado.frx":0C62
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmListado.frx":0D4C
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3780
         Picture         =   "frmListado.frx":0F86
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   4800
         Picture         =   "frmListado.frx":1088
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":118A
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmListado.frx":148C
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmListado.frx":17CE
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmListado.frx":1AD0
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   6540
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   13414
            Key             =   "msg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Picture         =   "frmListado.frx":1D0A
            Key             =   "printer"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   50
      TabIndex        =   20
      Top             =   0
      Width           =   9015
      Begin VB.TextBox tServicio 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         MaxLength       =   7
         TabIndex        =   24
         Top             =   660
         Width           =   735
      End
      Begin VB.TextBox tCodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1380
         MaxLength       =   5
         TabIndex        =   7
         Top             =   660
         Width           =   735
      End
      Begin VB.ComboBox cListado 
         Height          =   315
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin AACombo99.AACombo cReparar 
         Height          =   315
         Left            =   7080
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         ListIndex       =   -1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
      End
      Begin AACombo99.AACombo cCamion 
         Height          =   315
         Left            =   3840
         TabIndex        =   3
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         ListIndex       =   -1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
      End
      Begin VB.Label lDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   4620
         TabIndex        =   26
         Top             =   660
         Width           =   4275
      End
      Begin VB.Label Label3 
         Caption         =   "Agregar &Servicio:"
         Height          =   195
         Left            =   2220
         TabIndex        =   25
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "&Código Traslado:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "&Listado:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Local de &Entrega:"
         Height          =   195
         Left            =   5760
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Camión que &Traslada:"
         Height          =   255
         Left            =   2220
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin ComctlLib.ImageList ImgList 
      Left            =   7620
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmListado.frx":1E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmListado.frx":2136
            Key             =   "retiro"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuAcceso 
      Caption         =   "Accesos"
      Visible         =   0   'False
      Begin VB.Menu MnuAccMarcarTodos 
         Caption         =   "Marcar todos"
      End
      Begin VB.Menu MnuAccDesmarcarTodos 
         Caption         =   "Desmarcar todos"
      End
      Begin VB.Menu MnuAccLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAcEstadoPendiente 
         Caption         =   "Volver a Estado Pendiente"
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'1-7-2000 en hagocambiodeestado, tenia error cuando pasaba el tipo de local en camiones(pasaba el del deposito).

Private Enum EstadoS
    Anulado = 0
    Visita = 1
    Retiro = 2
    Taller = 3
    Entrega = 4
    Cumplido = 5
End Enum

Private aTexto As String
Private EmpresaEmisora As clsClienteCFE

Private TasaBasica As Currency, TasaMinima As Currency

Private Sub ImprimoEFactura(ByVal doc As Long)
On Error GoTo errIEF
    
    Dim oPrintEF As New ComPrintEfactura.ImprimoCFE
    oPrintEF.ImprimirCFEPorXML doc, paPrintCartaD, paPrintCartaB, paPrintCartaPaperSize
    Set oPrintEF = Nothing
    Exit Sub
    
errIEF:
    clsGeneral.OcurrioError "Error al imprimir eFactura", Err.Description
End Sub

Private Function EmitirCFE(ByVal Documento As Long) As Boolean
On Error GoTo errEC
    
    EmitirCFE = False
    If (TasaBasica = 0) Then CargoValoresIVA
    
    With New clsCGSAEFactura
        .URLAFirmar = ParametrosSist.ObtenerValorParametro(URLFirmaEFactura).Texto
        .ImporteConInfoDeCliente = ParametrosSist.ObtenerValorParametro(efactImporteDatosCliente).Valor
        .TasaBasica = TasaBasica
        .TasaMinima = TasaMinima
        
        Set .Connect = cBase
        Dim sResult As String
        sResult = .FirmarUnDocumento(Documento)
        If UCase(sResult) <> "TRUE" Then
            MsgBox "Importante!!!!" & vbCrLf & vbCrLf & " NO SE FIRMO EL REMITO " & vbCrLf & vbCrLf & "Por favor comuniquese con un administrador para cumplir esta acción.", vbError, "ATENCIÓN"
        Else
            EmitirCFE = True
        End If
    End With
    Exit Function
errEC:
    MsgBox "Error en firma: " & Err.Description, vbCritical, "ATENCIÓN"
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

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub bImprimir_Click()
    AccionImprimir True
End Sub
Private Sub bPrimero_Click()
    IrAPagina vsListado, 1
End Sub

Private Sub bSiguiente_Click()
    IrAPagina vsListado, vsListado.PreviewPage + 1
End Sub

Private Sub bUltima_Click()
    IrAPagina vsListado, vsListado.PageCount
End Sub

Private Sub bZMas_Click()
    Zoom vsListado, vsListado.Zoom + 5
End Sub

Private Sub bZMenos_Click()
    Zoom vsListado, vsListado.Zoom - 5
End Sub

Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub

Private Sub cCamion_GotFocus()
    With cCamion
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cCamion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cReparar
End Sub
Private Sub cCamion_LostFocus()
    cCamion.SelStart = 0
End Sub

Private Sub cListado_Click()
    InicializoGrillas
    If cListado.ListIndex = 1 Then
        tCodigo.Enabled = False: tCodigo.BackColor = vbButtonFace: tCodigo.Text = ""
        tServicio.Enabled = True: tServicio.BackColor = vbWindowBackground
    Else
        tCodigo.Enabled = True: tCodigo.BackColor = vbWindowBackground
        tServicio.Enabled = False: tServicio.BackColor = vbButtonFace
    End If
End Sub

Private Sub cListado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cCamion
End Sub

Private Sub cReparar_GotFocus()
    With cReparar: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub
Private Sub cReparar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tCodigo.Enabled Then Foco tCodigo Else bConsultar.SetFocus
    End If
End Sub
Private Sub cReparar_LostFocus()
    cReparar.SelStart = 0
End Sub

Private Sub chVista_Click()
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
        Me.Refresh
    Else
        AccionImprimir
        vsListado.ZOrder 0
        Me.Refresh
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad

    Status.Panels("printer").Text = paPrintCartaD
    lDetalle.Visible = False
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    InicializoGrillas
    FechaDelServidor
    
    cListado.AddItem "Impresos": cListado.ItemData(cListado.NewIndex) = 0
    cListado.AddItem "Pendientes": cListado.ItemData(cListado.NewIndex) = 1
    cListado.ListIndex = 1  'Por defecto Pendientes
    
    Cons = "Select CamCodigo, CamNombre From Camion Order by CamNombre"
    CargoCombo Cons, cCamion
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cReparar
    
    vsListado.Orientation = orPortrait: vsListado.PaperSize = 1
    
    If (EmpresaEmisora Is Nothing) Then
        Set EmpresaEmisora = New clsClienteCFE
        EmpresaEmisora.CargoClienteCarlosGutierrez paCodigoDeSucursal
    End If
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub
Private Sub InicializoGrillas()
    On Error Resume Next
    With vsConsulta
        .Redraw = False
        .WordWrap = False
        .Cols = 1: .Rows = 1
        If cListado.ListIndex = 0 Then aTexto = "Traslado" Else aTexto = "Selección"
        .FormatString = aTexto & "|>Servicio|<Producto|<Ingresó|<Origen|<Entregar en|Stock/Cliente|"
        .ColWidth(2) = 3000: .ColWidth(3) = 1150: .ColWidth(4) = 1080: .ColWidth(5) = 1080: .ColWidth(6) = 1650: .ColWidth(7) = 10
'        If cListado.ListIndex = 0 Then .ColDataType(0) = flexDTLong Else .ColDataType(0) = flexDTBoolean
        .Redraw = True
    End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
        
            Case vbKeyE: AccionConsultar
            
            Case vbKeyP: IrAPagina vsListado, 1
            Case vbKeyA: IrAPagina vsListado, vsListado.PreviewPage - 1
            Case vbKeyS: IrAPagina vsListado, vsListado.PreviewPage + 1
            Case vbKeyU: IrAPagina vsListado, vsListado.PageCount
            
            Case vbKeyAdd: Zoom vsListado, vsListado.Zoom + 5
            Case vbKeySubtract: Zoom vsListado, vsListado.Zoom - 5
            
            Case vbKeyI: AccionImprimir True
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            
            Case vbKeyG: AccionGrabar
            
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11

    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = fFiltros.Width
    vsListado.Left = fFiltros.Left
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = vsListado.Height
    vsConsulta.Left = vsListado.Left
    
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Private Sub AccionConsultar()
Dim rs As rdoResultset
    
    On Error GoTo errConsultar
    Screen.MousePointer = 11
    chVista.Value = 0
    vsConsulta.Rows = 1
    vsConsulta.Refresh
    vsConsulta.Redraw = False
    If cListado.ListIndex = 1 Then
        CargoServiciosPendientes
    Else
        If (cCamion.ListIndex = -1 And cReparar.ListIndex = -1) And Not IsNumeric(tCodigo.Text) Then
            MsgBox "Seleccione el camión que traslada o ingrese el código de traslado.", vbExclamation, "ATENCIÓN"
        Else
            CargoServiciosImpresos
        End If
    End If
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True
End Sub

Private Sub Label1_Click()
    Foco cReparar
End Sub
Private Sub Label2_Click()
    Foco cCamion
End Sub

Private Sub Label3_Click()
    tServicio.SetFocus
End Sub

Private Sub Label4_Click()
    Foco cListado
End Sub

Private Sub MnuAccDesmarcarTodos_Click()
        With vsConsulta
            I = 1
            Do While I <= .Rows - 1
                .Cell(flexcpChecked, I, 0) = flexUnchecked
                I = I + 1
            Loop
        End With
End Sub

Private Sub MnuAccMarcarTodos_Click()
    With vsConsulta
        I = 1
        Do While I <= .Rows - 1
            .Cell(flexcpChecked, I, 0) = 1
            I = I + 1
        Loop
    End With
End Sub

Private Sub MnuAcEstadoPendiente_Click()
    If MsgBox("Confirma quitar el servicio seleccionado del traslado impreso?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    AccionVueltaAtras
End Sub

Private Sub Status_PanelClick(ByVal Panel As ComctlLib.Panel)
    If "printer" = Panel.Key Then
        prj_GetPrinter True
        Panel.Text = paPrintCartaD
    End If
End Sub

Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub tCodigo_LostFocus()
    If Not IsNumeric(tCodigo.Text) Then tCodigo.Text = ""
End Sub

Private Sub tServicio_Change()
    lDetalle.Visible = False
    lDetalle.Caption = ""
End Sub

Private Sub tServicio_GotFocus()
On Error Resume Next
    With tServicio
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tServicio_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        AgregoServicioManual tServicio.Text
    End If
End Sub

Private Sub vsConsulta_DblClick()
    If vsConsulta.Row >= 1 And cListado.ListIndex = 1 Then
        If vsConsulta.Cell(flexcpChecked, vsConsulta.Row, 0) = flexChecked Then
            vsConsulta.Cell(flexcpChecked, vsConsulta.Row, 0) = flexUnchecked
        Else
            vsConsulta.Cell(flexcpChecked, vsConsulta.Row, 0) = flexChecked
        End If
    End If
End Sub

Private Sub vsConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 93
            If vsConsulta.Row > 0 And cListado.ListIndex = 0 Then PopupMenu MnuAcceso, , vsConsulta.Left + 800, vsConsulta.Top + 600
        Case vbKeySpace
            If vsConsulta.Row >= 1 And cListado.ListIndex = 1 Then
                If vsConsulta.Cell(flexcpChecked, vsConsulta.Row, 0) = flexChecked Then
                    vsConsulta.Cell(flexcpChecked, vsConsulta.Row, 0) = flexUnchecked
                Else
                    vsConsulta.Cell(flexcpChecked, vsConsulta.Row, 0) = flexChecked
                End If
            End If
    End Select
End Sub

Private Sub vsConsulta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And vsConsulta.Row > 0 Then
       MnuAcEstadoPendiente.Enabled = (cListado.ListIndex = 0)
       MnuAccDesmarcarTodos.Enabled = Not MnuAcEstadoPendiente.Enabled
       MnuAccMarcarTodos.Enabled = Not MnuAcEstadoPendiente.Enabled
       PopupMenu MnuAcceso
    End If
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False, Optional DosCopias As Boolean = False, Optional CodigoTraslado As Long = 0)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub
    SeteoImpresoraPorDefecto paPrintCartaD
    With vsListado
        .Device = paPrintCartaD
        .PaperBin = paPrintCartaB
        .PaperSize = paPrintCartaPaperSize
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    aTexto = "Resumen de Traslados de Servicios"
    If cCamion.ListIndex > -1 Then aTexto = aTexto & ", Camión: " & Trim(cCamion.Text)
    If CodigoTraslado > 0 Then aTexto = aTexto & Space(10) & "Traslado = " & CodigoTraslado
    
    EncabezadoListado vsListado, aTexto, False
    vsListado.FileName = "Traslados de Servicios"
    vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
    vsListado.EndDoc
    
    If Imprimir Then
        With vsListado
            .Device = paPrintCartaD
            .PaperBin = paPrintCartaB
            .PaperSize = paPrintCartaPaperSize
            .PrintDoc
            If DosCopias Then .PrintDoc
        End With
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub

Private Sub Ayuda(strTexto As String)
    Status.Panels("msg").Text = strTexto
End Sub
Private Sub CargoServiciosImpresos()
Dim aModificacion As String, aValor As Long
  
    'Cargo Impresos.
    If IsNumeric(tCodigo.Text) Then
        Cons = "Select * From Servicio, Taller, Producto, TrasladoServicio "
    Else
        Cons = "Select * From Servicio, Taller, Producto " _
            & " Left Outer Join TrasladoServicio ON TSeProducto = ProCodigo "
    End If
    
    Cons = "SELECT * FROM Servicio INNER JOIN Taller ON SerCodigo = TalServicio " & _
                    "INNER JOIN Producto ON SerProducto = ProCodigo " & _
                    "INNER JOIN Articulo ON ProArticulo = ArtID " & _
                    "INNER JOIN Cliente ON SerCliente = CliCodigo " & _
                    "INNER JOIN TrasladoServicio ON SerCodigo = TSeServicio AND SerProducto = TSeProducto " & _
                    "INNER JOIN Sucursal ON SerLocalIngreso = SucCodigo " & _
                    "INNER JOIN Local ON SerLocalReparacion = LocCodigo " & _
            "WHERE SerEstadoServicio = " & EstadoS.Taller & " " & _
            "AND SerLocalIngreso = " & paCodigoDeSucursal
        
    If IsNumeric(tCodigo.Text) Then
        Cons = Cons & " And TSeCodigo = " & CLng(tCodigo.Text)
    Else
        If cCamion.ListIndex > -1 Then Cons = Cons & " And TalIngresoCamion  = " & cCamion.ItemData(cCamion.ListIndex)
        If cReparar.ListIndex > -1 Then Cons = Cons & " And SerLocalReparacion = " & cReparar.ItemData(cReparar.ListIndex)
    End If
    
'    Cons = Cons & " , Articulo, Sucursal, Local " _
'        & " Where SerEstadoServicio = " & EstadoS.Taller _
'        & " And TalFIngresoRealizado <> Null And TalFIngresoRecepcion = Null " _
'        & " And SerLocalIngreso = " & paCodigoDeSucursal
    
'    Cons = Cons & " And SerLocalIngreso <> SerLocalReparacion And SerProducto = ProCodigo And ProArticulo = ArtID " _
'        & " AND TSeServicio = SerCodigo "
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsAux.EOF
    
        With vsConsulta
            .AddItem ""
            
            aModificacion = RsAux!SerModificacion
            .Cell(flexcpData, .Rows - 1, 0) = aModificacion
            If cListado.ListIndex = 0 Then
                aModificacion = RsAux!TalModificacion
                .Cell(flexcpData, .Rows - 1, 1) = aModificacion
            End If
            .Cell(flexcpData, .Rows - 1, 2) = 1 'Me digo que es para Reparar.
            
            If RsAux!SerCliente = EmpresaEmisora.Codigo Then .Cell(flexcpData, .Rows - 1, 3) = EmpresaEmisora.Codigo Else .Cell(flexcpData, .Rows - 1, 3) = 0
            aValor = RsAux!ProCodigo: .Cell(flexcpData, .Rows - 1, 4) = aValor
            aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 5) = aValor
            
            If Not IsNull(RsAux("TSeRemito")) Then
                aValor = RsAux!TSeRemito: .Cell(flexcpData, .Rows - 1, 6) = aValor
            End If
            
            If Not IsNull(RsAux!TSeCodigo) Then .Cell(flexcpText, .Rows - 1, 0) = RsAux!TSeCodigo
            .Cell(flexcpText, .Rows - 1, 1) = RsAux!SerCodigo
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ProCodigo, "(#,000)") & " " & Trim(RsAux!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!SerFecha, FormatoFP)
            .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!SucAbreviacion)
            .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!LocNombre)
            
            If RsAux("CliCodigo") = EmpresaEmisora.Codigo Then
                .Cell(flexcpText, .Rows - 1, 6) = "STOCK"
            Else
                .Cell(flexcpText, .Rows - 1, 6) = IIf(RsAux("CliTipo") = 2, clsGeneral.RetornoFormatoRuc(RsAux("ClICIRUC")), clsGeneral.RetornoFormatoCedula(RsAux("ClICIRUC")))
            End If
            
        End With
        RsAux.MoveNext
        
    Loop
    RsAux.Close
    
    'Ahora Cargo los que estan para entregar.
    'Cargo Impresos.
    If IsNumeric(tCodigo.Text) Then
        Cons = "Select * From Servicio, Taller, Producto, TrasladoServicio "
    Else
        Cons = "Select * From Servicio, Taller, Producto " _
            & " Left Outer Join TrasladoServicio ON TSeProducto = ProCodigo "
    End If
            
    Cons = Cons & " , Articulo, Sucursal, Local " _
        & " Where SerEstadoServicio = " & EstadoS.Taller _
        & " And TalFSalidaRealizado <> Null And TalFSalidaRecepcion = Null " _
        & " And SerLocalReparacion = " & paCodigoDeSucursal
    
    If IsNumeric(tCodigo.Text) Then Cons = Cons & " And TSeCodigo = " & CLng(tCodigo.Text) & " And TSeProducto = ProCodigo "
    If cCamion.ListIndex > -1 Then Cons = Cons & " And TalSalidaCamion  = " & cCamion.ItemData(cCamion.ListIndex)
    If cReparar.ListIndex > -1 Then Cons = Cons & " And TalLocalAlCliente = " & cReparar.ItemData(cReparar.ListIndex)
        
    Cons = Cons & " And SerLocalReparacion <> TalLocalAlCliente And SerCodigo = TalServicio And SerProducto = ProCodigo And ProArticulo = ArtID " _
            & " AND TSeServicio = SerCodigo And SerCostoFinal <> Null And SerLocalIngreso = SucCodigo And SerLocalReparacion = LocCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsAux.EOF
    
        With vsConsulta
            .AddItem ""
            
            aModificacion = RsAux!SerModificacion
            .Cell(flexcpData, .Rows - 1, 0) = aModificacion
            aModificacion = RsAux!TalModificacion: .Cell(flexcpData, .Rows - 1, 1) = aModificacion
            
            .Cell(flexcpData, .Rows - 1, 2) = 2 'Me digo que es para Entregar.
            
            If RsAux!SerCliente = EmpresaEmisora.Codigo Then .Cell(flexcpData, .Rows - 1, 3) = EmpresaEmisora.Codigo Else .Cell(flexcpData, .Rows - 1, 3) = 0
            aValor = RsAux!ProCodigo: .Cell(flexcpData, .Rows - 1, 4) = aValor
            aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 5) = aValor
            
            If Not IsNull(RsAux!TSeCodigo) Then .Cell(flexcpText, .Rows - 1, 0) = RsAux!TSeCodigo
            
            .Cell(flexcpText, .Rows - 1, 1) = RsAux!SerCodigo
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ProCodigo, "(#,000)") & " " & Trim(RsAux!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!SerFecha, FormatoFP)
            .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!SucAbreviacion)
            .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!LocNombre)
            
        End With
        RsAux.MoveNext
        
    Loop
    RsAux.Close
    
    If vsConsulta.Rows > 1 Then
        With vsConsulta
            .Select 1, 1, .Rows - 1, 1
            .Sort = flexSortGenericDescending
            .Subtotal flexSTCount, -1, 2, "#,##0", Obligatorio, Colores.Rojo, True, "Total", , True
        End With
    End If
End Sub
Private Sub CargoServiciosPendientes()
Dim aModificacion As String, aValor As Long

    Cons = "Select * From Servicio INNER JOIN Producto ON SerProducto = ProCodigo" _
        & " INNER JOIN Cliente ON CliCodigo = SerCliente " _
        & " INNER JOIN Articulo ON ProArticulo = ArtID" _
        & " INNER JOIN Sucursal ON SerLocalIngreso = SucCodigo" _
        & " INNER JOIN Local ON SerLocalReparacion = LocCodigo" _
        & " WHERE SerEstadoServicio = " & EstadoS.Taller _
        & " And SerLocalIngreso = " & paCodigoDeSucursal _
        & " AND SerCodigo NOT IN(SELECT TalServicio FROM Taller WHERE TalFIngresoRecepcion IS NOT NULL OR TalIngresoCamion IS NOT NULL) "
        
'OR TalIngresoCamion IS NOT NULL
'& " And SerCodigo Not IN(Select TalServicio From Taller)" _

    If cReparar.ListIndex > -1 Then Cons = Cons & " And SerLocalReparacion = " & cReparar.ItemData(cReparar.ListIndex)
    
    Cons = Cons & " And SerLocalIngreso <> SerLocalReparacion"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsAux.EOF
    
'        If Not EstaCargado(RsAux!SerCodigo) Then
            With vsConsulta
                .AddItem ""
                
                aModificacion = RsAux!SerModificacion
                .Cell(flexcpData, .Rows - 1, 0) = aModificacion
                If cListado.ListIndex = 0 Then
                    aModificacion = RsAux!TalModificacion
                    .Cell(flexcpData, .Rows - 1, 1) = aModificacion
                End If
                .Cell(flexcpData, .Rows - 1, 2) = 1 'Me digo que es para Reparar.
                
                If RsAux!SerCliente = EmpresaEmisora.Codigo Then .Cell(flexcpData, .Rows - 1, 3) = EmpresaEmisora.Codigo Else .Cell(flexcpData, .Rows - 1, 3) = 0
                aValor = RsAux!ProCodigo: .Cell(flexcpData, .Rows - 1, 4) = aValor
                aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 5) = aValor
                
                .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked
                .Cell(flexcpText, .Rows - 1, 1) = RsAux!SerCodigo
                .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ProCodigo, "(#,000)") & " " & Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!SerFecha, FormatoFP)
                .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!SucAbreviacion)
                .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!LocNombre)
                
                If RsAux("CliCodigo") = EmpresaEmisora.Codigo Then
                    .Cell(flexcpText, .Rows - 1, 6) = "STOCK"
                Else
                    .Cell(flexcpText, .Rows - 1, 6) = IIf(RsAux("CliTipo") = 2, clsGeneral.RetornoFormatoRuc(RsAux("ClICIRUC")), clsGeneral.RetornoFormatoCedula(RsAux("ClICIRUC")))
                End If
                
            End With
 '       End If
        RsAux.MoveNext
        
    Loop
    RsAux.Close
    
    'A volver al local de origen
    Cons = "Select * From Servicio, Taller, Producto, Articulo, Sucursal, Local" _
        & " Where SerEstadoServicio = " & EstadoS.Taller _
        & " And TalFIngresoRecepcion Is Not Null And TalSalidaCamion  = Null And TalFSalidaRealizado = Null And TalFSalidaRecepcion = Null " _
        & " And SerLocalReparacion = " & paCodigoDeSucursal & " And TalFReparado Is Not Null"
    
    If cReparar.ListIndex > -1 Then Cons = Cons & " And TalLocalAlCliente = " & cReparar.ItemData(cReparar.ListIndex)
        
    Cons = Cons & " And SerLocalReparacion <> TalLocalAlCliente And SerCodigo = TalServicio And SerProducto = ProCodigo And ProArticulo = ArtID " _
            & " And SerLocalIngreso = SucCodigo And SerLocalReparacion = LocCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsAux.EOF
    
'        If Not EstaCargado(RsAux!SerCodigo) Then
            With vsConsulta
                .AddItem ""
                
                aModificacion = RsAux!SerModificacion
                .Cell(flexcpData, .Rows - 1, 0) = aModificacion
                aModificacion = RsAux!TalModificacion: .Cell(flexcpData, .Rows - 1, 1) = aModificacion
                
                .Cell(flexcpData, .Rows - 1, 2) = 2 'Me digo que es para Entregar.
                
                If RsAux!SerCliente = EmpresaEmisora.Codigo Then .Cell(flexcpData, .Rows - 1, 3) = EmpresaEmisora.Codigo Else .Cell(flexcpData, .Rows - 1, 3) = 0
                aValor = RsAux!ProCodigo: .Cell(flexcpData, .Rows - 1, 4) = aValor
                aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 5) = aValor
                
                .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked
                .Cell(flexcpText, .Rows - 1, 1) = RsAux!SerCodigo
                .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ProCodigo, "(#,000)") & " " & Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!SerFecha, FormatoFP)
                .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!SucAbreviacion)
                .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!LocNombre)
                
                If IsNull(RsAux!TalFReparado) Then
                    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio: .Cell(flexcpChecked, .Rows - 1, 0) = flexUnchecked
                Else
                    If IsNull(RsAux!SerCostoFinal) Or IsNull(RsAux!TalFAceptacion) Then
                        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Inactivo: .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked
                    End If
                End If
                
            End With
 '       End If
        RsAux.MoveNext
        
    Loop
    RsAux.Close
    
'    If vsConsulta.Rows > 1 Then
'        With vsConsulta
'            .Select 1, 1, .Rows - 1, 1
'            .Sort = flexSortGenericDescending
'            .Subtotal flexSTCount, -1, 2, "#,##0", Obligatorio, Colores.Rojo, True, "Total", , True
'        End With
'    End If
    
End Sub

Private Sub AccionGrabar()
Dim Msg As String, Usuario As String, idLocalEntrega As Long
Dim IdTraslado As Long
Dim sPaso As String

    If cListado.ListIndex = 0 Then Exit Sub
    If vsConsulta.Rows = 1 Then MsgBox "No hay datos en la consulta.", vbExclamation, "ATENCIÓN": Exit Sub
    If cCamion.ListIndex = -1 Then MsgBox "Seleccione el camión que trasladara los artículos.", vbInformation, "ATENCIÓN": Foco cCamion: Exit Sub
    
    If MsgBox("¿Confirma grabar el traslado?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    
    Dim beRemito As Boolean
    beRemito = False
    With vsConsulta
        For I = 1 To .Rows - 1
            If .Cell(flexcpChecked, I, 0) = flexChecked And CLng(.Cell(flexcpData, I, 3)) = EmpresaEmisora.Codigo Then
                beRemito = True
                Exit For
            End If
        Next
    End With
    
    'Valido si hay traslado de stock y si puedo emitir el eRemito.
    Dim oInfoCAE As New clsCAEGenerador
    If beRemito Then
        If Not oInfoCAE.SucursalTieneCae(cBase, CGSA_TiposCFE.CFE_eRemito, paCodigoDGI) Then
            MsgBox "No hay un CAE disponible para emitir el eRemito, por favor comuniquese con administración." & vbCrLf & vbCrLf & "No podrá recepcionar", vbCritical, "eFactura"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    
    sPaso = "Usuario"
    Usuario = ""
    Usuario = InputBox("Ingrese su digito de usuario.", "Grabar Traslado")
    If Not IsNumeric(Usuario) Then Exit Sub
    
    Usuario = BuscoUsuarioDigito(CLng(Usuario), True)
    
    If Val(Usuario) = 0 Then MsgBox "Usuario incorrecto.", vbExclamation, "ATENCIÓN": Exit Sub
    
    Dim iEstARecuperar As Integer
    iEstARecuperar = ParametrosSist.ObtenerValorParametro(EstadoARecuperar).Valor

    
    On Error GoTo ErrBT
    Screen.MousePointer = 11
    sPaso = "Inicio Tran"
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo ErrRB
    
    Cons = "Select Max(TSeCodigo) From TrasladoServicio"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not IsNull(RsAux(0)) Then IdTraslado = RsAux(0) + 1 Else IdTraslado = 1
    RsAux.Close
    
    Dim iLocalCpia As Integer
    iLocalCpia = ParametrosSist.ObtenerValorParametro(LocalCompañia).Valor
    
    Dim CAE As clsCAEDocumento
    Dim docRemito As New clsDocumentoCGSA
    'Es cuando llega menos mercadería que la que generó el remito.
    'Inserto traslado con destino --> origen.

    If beRemito Then
        Set CAE = oInfoCAE.ObtenerNumeroCAEDocumento(cBase, CGSA_TiposCFE.CFE_eRemito, paCodigoDGI)
        With docRemito
            Set .Cliente = EmpresaEmisora
            .Emision = gFechaServidor
            .Tipo = TD_TrasladoServicios
            .Numero = CAE.Numero
            .Serie = CAE.Serie
            .Moneda.Codigo = 1
            .Total = 0
            .IVA = 0
            .sucursal = paCodigoDeSucursal
            .Digitador = Usuario
            .Comentario = "Traslado:" & IdTraslado & "<BR/>" & "De: " & Trim(paNombreSucursal) & ", Camión: " & Trim(cCamion.Text) & ", a: " & Trim(cReparar.Text)
            .Vendedor = Usuario
        End With
        Set docRemito.Conexion = cBase
        docRemito.Codigo = docRemito.InsertoCabezalDelDocumento()
    End If
    
    Dim oRengloneRemito As clsDocumentoRenglon
    Dim oRenglonNuevo As clsDocumentoRenglon
   
    With vsConsulta
        For I = 1 To .Rows - 1
            If .Cell(flexcpChecked, I, 0) = flexChecked Then
                Msg = ""
                'TABLA SERVICIO
                sPaso = "Servicio"
                Cons = "Select * From Servicio Where SerCodigo = " & Val(.Cell(flexcpText, I, 1))
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If RsAux.EOF Then
                    Msg = "Otra terminal elimino el servicio = " & Val(.Cell(flexcpText, I, 1))
                    RsAux.Close: RsAux.Edit 'Provoco error.
                Else
                    idLocalEntrega = cReparar.ItemData(cReparar.ListIndex) 'RsAux!SerLocalIngreso
                    If RsAux!SerModificacion = CDate(.Cell(flexcpData, I, 0)) Then
                        RsAux.Edit
                        RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
                        RsAux.Update
                        RsAux.Close
                    Else
                        Msg = "Otra terminal modificó el servicio = " & Val(.Cell(flexcpText, I, 1))
                        RsAux.Close: RsAux.Edit 'Provoco error.
                    End If
                End If
                
                If Val(.Cell(flexcpData, I, 2)) = 1 Then
                    'Si el servicio ya tiene ficha de taller entonces sólo la edito.
                    Cons = "SELECT * FROM Taller WHERE TalServicio = " & Val(.Cell(flexcpText, I, 1))
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If RsAux.EOF Then
                        RsAux.Close
                        sPaso = "INSERT TALLER"
                        'INSERTO En Tabla Taller
                        Cons = "Insert Into Taller (TalServicio, TalFIngresoRealizado, TalIngresoCamion, TalModificacion, TalUsuario, TalLocalAlCliente, TalFIngresoRecepcion) Values (" _
                            & Val(.Cell(flexcpText, I, 1)) & ", '" & Format(gFechaServidor, sqlFormatoFH) & "',  " & cCamion.ItemData(cCamion.ListIndex) _
                            & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', " & Usuario & ", " & idLocalEntrega
                        If cReparar.ItemData(cReparar.ListIndex) = iLocalCpia Or cReparar.ItemData(cReparar.ListIndex) = 17 Then
                            Cons = Cons & ", '" & Format(gFechaServidor, sqlFormatoFH) & "')"
                        Else
                            Cons = Cons & ", Null)"
                        End If
                        cBase.Execute (Cons)
                        
                    Else
                        sPaso = "Edit Taller"
                        RsAux.Edit
                        RsAux("TalIngresoCamion") = cCamion.ItemData(cCamion.ListIndex)
                        RsAux("TalModificacion") = Format(gFechaServidor, sqlFormatoFH)
                        RsAux("TalUsuario") = Usuario
                        If IsNull(RsAux("TalLocalAlCliente")) Then RsAux("TalLocalAlCliente") = idLocalEntrega
                        If cReparar.ItemData(cReparar.ListIndex) = iLocalCpia Or cReparar.ItemData(cReparar.ListIndex) = 17 Then
                            RsAux("TalFIngresoRecepcion") = Format(gFechaServidor, sqlFormatoFH)
                        End If
                        RsAux.Update
                        RsAux.Close
                    End If
                    
                    'Inserto en la tabla trasladoServicio el producto.
                    sPaso = "Traslado servicio"
                    
                    Cons = "Insert Into TrasladoServicio (TSeCodigo, TSeProducto, TSeServicio, TSeRemito) Values (" & IdTraslado & ", " & CLng(.Cell(flexcpData, I, 4)) & ", " & Val(.Cell(flexcpText, I, 1)) _
                        & ", " & IIf(CLng(.Cell(flexcpData, I, 3)) = EmpresaEmisora.Codigo, docRemito.Codigo, "Null") & ")"
                    cBase.Execute (Cons)
                    
                    'Si el artículo es de carlos gutierrez marco el traslado en el stock.
                    If CLng(.Cell(flexcpData, I, 3)) = EmpresaEmisora.Codigo Then
                        
                        Set oRenglonNuevo = Nothing
                        'Primero veo si ya lo tengo en la lista.
                        For Each oRengloneRemito In docRemito.Renglones
                            If oRengloneRemito.Articulo.ID = CLng(.Cell(flexcpData, I, 5)) Then
                                Set oRenglonNuevo = oRengloneRemito
                                Exit For
                            End If
                        Next
                        If oRenglonNuevo Is Nothing Then
                            Set oRenglonNuevo = New clsDocumentoRenglon
                            oRenglonNuevo.Articulo.ID = CLng(.Cell(flexcpData, I, 5))
                            oRenglonNuevo.Cantidad = 1
                            oRenglonNuevo.EstadoMercaderia = iEstARecuperar
                            docRemito.Renglones.Add oRenglonNuevo
                        Else
                            oRenglonNuevo.Cantidad = oRenglonNuevo.Cantidad + 1
                        End If
                    
                        sPaso = "Cambio local": HagoCambioDeLocal CLng(.Cell(flexcpData, I, 5)), Val(.Cell(flexcpText, I, 1)), CLng(Usuario), True, cCamion.ItemData(cCamion.ListIndex), paCodigoDeSucursal
                        'si es compañia o neotron muevo del camión al local.
                        If (cReparar.ItemData(cReparar.ListIndex) = iLocalCpia Or cReparar.ItemData(cReparar.ListIndex) = 17) Then sPaso = "Cambio de local 2": HagoCambioDeLocal CLng(.Cell(flexcpData, I, 5)), Val(.Cell(flexcpText, I, 1)), CLng(Usuario), False, cCamion.ItemData(cCamion.ListIndex), cReparar.ItemData(cReparar.ListIndex)
                    End If
                    
                Else
                
                    Cons = "Select * From Taller Where TalServicio = " & Val(.Cell(flexcpText, I, 1))
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If RsAux.EOF Then
                        Msg = "Otra terminal elimino los datos del taller para el servicio = " & Val(.Cell(flexcpText, I, 1))
                        RsAux.Close: RsAux.Edit 'Provoco error.
                    Else
                        If RsAux!TalModificacion = CDate(.Cell(flexcpData, I, 1)) Then
                            sPaso = "Taller edit 2"
                            RsAux.Edit
                            RsAux!TalFSalidaRealizado = Format(gFechaServidor, sqlFormatoFH)
                            RsAux!TalModificacion = Format(gFechaServidor, sqlFormatoFH)
                            RsAux!TalSalidaCamion = cCamion.ItemData(cCamion.ListIndex)
                            If IsNull(RsAux("TalLocalAlCliente")) Then RsAux!TalLocalAlCliente = idLocalEntrega
                            RsAux!TalUsuario = CLng(Usuario)
                            RsAux.Update
                            RsAux.Close
                            
                            sPaso = "Traslado de servicio 2"
                            'Inserto en la tabla trasladoServicio el producto.
                            Cons = "Insert Into TrasladoServicio (TSeCodigo, TSeProducto, TSeServicio) Values (" & IdTraslado & ", " & CLng(.Cell(flexcpData, I, 4)) & ", " & Val(.Cell(flexcpText, I, 1)) & ")"
                            cBase.Execute (Cons)
                            
                            'Tengo que hacer movimiento físico al camión.
                            If CLng(.Cell(flexcpData, I, 3)) = EmpresaEmisora.Codigo Then
                                
                                Set oRenglonNuevo = Nothing
                                'Primero veo si ya lo tengo en la lista.
                                For Each oRengloneRemito In docRemito.Renglones
                                    If oRengloneRemito.Articulo.ID = CLng(.Cell(flexcpData, I, 5)) Then
                                        Set oRenglonNuevo = oRengloneRemito
                                        Exit For
                                    End If
                                Next
                                If oRenglonNuevo Is Nothing Then
                                    Set oRenglonNuevo = New clsDocumentoRenglon
                                    oRenglonNuevo.Articulo.ID = CLng(.Cell(flexcpData, I, 5))
                                    oRenglonNuevo.Cantidad = 1
                                    oRenglonNuevo.EstadoMercaderia = iEstARecuperar
                                    docRemito.Renglones.Add oRenglonNuevo
                                Else
                                    oRenglonNuevo.Cantidad = oRenglonNuevo.Cantidad + 1
                                End If

                                sPaso = "Cambio local 3 ": HagoCambioDeLocal CLng(.Cell(flexcpData, I, 5)), Val(.Cell(flexcpText, I, 1)), CLng(Usuario), True, cCamion.ItemData(cCamion.ListIndex), paCodigoDeSucursal
                            End If
                        Else
                            Msg = "Otra terminal modifico el servicio = " & Val(.Cell(flexcpText, I, 1))
                            RsAux.Close: RsAux.Edit 'Provoco error.
                        End If
                    End If
                End If
            End If
        Next I
    End With
    
    'Grabo los renglones del eRemito.
    If beRemito Then docRemito.InsertoRenglonDocumentoEstadoBD
    cBase.CommitTrans
    
    If beRemito Then
        If EmitirCFE(docRemito.Codigo) Then ImprimoEFactura docRemito.Codigo
    End If
    
    If MsgBox("¿Desea imprimir los traslados?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
        'Como voy a imprimir eliminó los que no están asignados.
        With vsConsulta
            I = 1
            Do While I <= .Rows - 1
                If .Cell(flexcpChecked, I, 0) = flexUnchecked Then .RemoveItem I: I = I - 1
                I = I + 1
            Loop
        End With
        AccionImprimir True, True, IdTraslado
    End If
    AccionConsultar
    Screen.MousePointer = 0
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrRB:
    Resume ErrVA
ErrVA:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la información." & Chr(13) & Msg & vbCrLf & sPaso, Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub AccionVueltaAtras()
Dim Msg As String, Usuario As String, IdCamion As Long
    
    If CLng(vsConsulta.Cell(flexcpData, vsConsulta.Row, 3)) > 0 Then
        Usuario = ""
        Usuario = InputBox("Ingrese su digito de usuario.", "Grabar Traslado")
        If Not IsNumeric(Usuario) Then Exit Sub
        
        Usuario = BuscoUsuarioDigito(CLng(Usuario), True)
        
        If Val(Usuario) = 0 Then MsgBox "Usuario incorrecto.", vbExclamation, "ATENCIÓN": Exit Sub
    End If
    
    If CLng(vsConsulta.Cell(flexcpData, vsConsulta.Row, 3)) = EmpresaEmisora.Codigo And CLng(vsConsulta.Cell(flexcpData, vsConsulta.Row, 6)) > 0 Then
        
        'Valido si hay traslado de stock y si puedo emitir el eRemito.
        Dim oInfoCAE As New clsCAEGenerador
        
        If Not oInfoCAE.SucursalTieneCae(cBase, CGSA_TiposCFE.CFE_eRemito, paCodigoDGI) Then
            MsgBox "No hay un CAE disponible para emitir el eRemito, por favor comuniquese con administración." & vbCrLf & vbCrLf & "No podrá recepcionar", vbCritical, "eFactura"
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        MsgBox "Se emitirá un eRemito para corregir el traslado original.", vbExclamation, "ATENCIÓN"
    End If
    
    On Error GoTo ErrBT
    Screen.MousePointer = 11
    Dim idLocalRepara As Long
    FechaDelServidor
    
    cBase.BeginTrans
    On Error GoTo ErrRB
    With vsConsulta
        Msg = ""
        'TABLA SERVICIO
        Cons = "Select * From Servicio Where SerCodigo = " & Val(.Cell(flexcpText, .Row, 1))
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            Msg = "Otra terminal elimino el servicio."
            RsAux.Close: RsAux.Edit 'Provoco error.
        Else
            idLocalRepara = RsAux("SerLocalReparacion")
            If RsAux!SerModificacion = CDate(.Cell(flexcpData, .Row, 0)) Then
                RsAux.Edit
                RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
                RsAux.Update
                RsAux.Close
            Else
                Msg = "Otra terminal modifico el servicio."
                RsAux.Close: RsAux.Edit 'Provoco error.
            End If
        End If
        
        'INSERTO En Tabla Taller
        Cons = "Select * From Taller Where TalServicio = " & Val(.Cell(flexcpText, .Row, 1))
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            Msg = "Otra terminal elimino los datos de taller del servicio."
            RsAux.Close: RsAux.Edit 'Provoco error.
        Else
            
            If RsAux!TalModificacion = CDate(.Cell(flexcpData, .Row, 1)) Then
                
                'Elimino el producto de la tabla TrasladoServicio.
                If Val(.Cell(flexcpText, .Row, 0)) > 0 Then
                    Cons = "Delete TrasladoServicio Where TSeCodigo = " & Val(.Cell(flexcpText, .Row, 0)) _
                            & " And TSeProducto =  " & CLng(.Cell(flexcpData, .Row, 4)) & " AND TSeServicio = " & CLng(.Cell(flexcpText, .Row, 1))
                    cBase.Execute (Cons)
                End If
                
                RsAux.Edit
                If IsNull(RsAux!TalSalidaCamion) Then
                    IdCamion = RsAux!TalIngresoCamion
                    RsAux!TalIngresoCamion = Null
                Else
                    IdCamion = RsAux!TalSalidaCamion
                    RsAux!TalSalidaCamion = Null
                End If
                RsAux!TalModificacion = Format(gFechaServidor, sqlFormatoFH)
                RsAux.Update
                RsAux.Close
                
                'Si el artículo es de carlos gutierrez marco el traslado en el stock.
                If CLng(.Cell(flexcpData, .Row, 3)) = EmpresaEmisora.Codigo Then
                    
                    Dim CAE As clsCAEDocumento
                    Dim docRemito As New clsDocumentoCGSA
                    Set CAE = oInfoCAE.ObtenerNumeroCAEDocumento(cBase, CGSA_TiposCFE.CFE_eRemito, paCodigoDGI)
                    With docRemito
                        Set .Cliente = EmpresaEmisora
                        .Emision = gFechaServidor
                        .Tipo = TD_TrasladoServicios
                        .Numero = CAE.Numero
                        .Serie = CAE.Serie
                        .Moneda.Codigo = 1
                        .Total = 0
                        .IVA = 0
                        .sucursal = paCodigoDeSucursal
                        .Digitador = Usuario
                        .Comentario = "Corrección eRemito: " & Val(vsConsulta.Cell(flexcpData, vsConsulta.Row, 6))
                        .Vendedor = Usuario
                    End With
                    Set docRemito.Conexion = cBase
                    docRemito.Codigo = docRemito.InsertoCabezalDelDocumento
                    Dim oRenglon As New clsDocumentoRenglon
                    oRenglon.Articulo.ID = CLng(.Cell(flexcpData, .Row, 5))
                    oRenglon.Cantidad = 1
                    oRenglon.EstadoMercaderia = ParametrosSist.ObtenerValorParametro(EstadoARecuperar).Valor
                    docRemito.Renglones.Add oRenglon
                    docRemito.InsertoRenglonDocumentoEstadoBD
                    
                    If idLocalRepara = ParametrosSist.ObtenerValorParametro(LocalCompañia).Valor Or idLocalRepara = 17 Then HagoCambioDeLocal CLng(.Cell(flexcpData, .Row, 5)), Val(.Cell(flexcpText, .Row, 1)), CLng(Usuario), True, IdCamion, idLocalRepara
                    HagoCambioDeLocal CLng(.Cell(flexcpData, .Row, 5)), Val(.Cell(flexcpText, .Row, 1)), CLng(Usuario), False, IdCamion, paCodigoDeSucursal
                    
                End If
            Else
                Msg = "Otra terminal modifico los datos de taller del servicio."
                RsAux.Close: RsAux.Edit 'Provoco error.
            End If
        End If
    End With
    cBase.CommitTrans
    
    If CLng(vsConsulta.Cell(flexcpData, vsConsulta.Row, 3)) = EmpresaEmisora.Codigo And CLng(vsConsulta.Cell(flexcpData, vsConsulta.Row, 6)) > 0 Then
        EmitirCFE docRemito.Codigo
        ImprimoEFactura docRemito.Codigo
    End If
    
    AccionConsultar
    Screen.MousePointer = 0
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrRB:
    Resume ErrVA
ErrVA:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la vuelta atras." & Chr(13) & Msg, Err.Description
    AccionConsultar
    Screen.MousePointer = 0
End Sub

Private Sub HagoCambioDeLocal(IDArticulo As Long, IdServicio As Long, IdUsuario As Long, LeDoyAlCamion As Boolean, ByVal IdCamion As Long, ByVal idLocal As Long)
'Si Altabajalocal = -1 entonces le doy de baja al local sino le doy de baja al camion .

Dim iEstARecuperar As Integer
iEstARecuperar = ParametrosSist.ObtenerValorParametro(EstadoARecuperar).Valor

    'Cedo el artículo al camión.
    MarcoMovimientoStockFisico IdUsuario, TipoLocal.Deposito, idLocal, IDArticulo, 1, iEstARecuperar, IIf(LeDoyAlCamion, -1, 1), TD_TrasladoServicios, IdServicio
    MarcoMovimientoStockFisico IdUsuario, TipoLocal.Camion, IdCamion, IDArticulo, 1, iEstARecuperar, IIf(LeDoyAlCamion, 1, -1), TD_TrasladoServicios, IdServicio
    
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, idLocal, IDArticulo, 1, iEstARecuperar, IIf(LeDoyAlCamion, -1, 1)
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Camion, IdCamion, IDArticulo, 1, iEstARecuperar, IIf(LeDoyAlCamion, 1, -1)
    
End Sub

Private Sub AgregoServicioManual(ByVal IdServicio As Long)
Dim aModificacion As String, aValor As Long
Dim iUno As Integer

    lDetalle.Visible = False
    'Si hay datos en la lista veo si el mismo esta ingresado.
    If EstaCargado(IdServicio) Then
        MsgBox "El servicio ya esta cargado en la lista.", vbInformation, "ATENCIÓN"
        Exit Sub
    Else
        iUno = 1
        Cons = "Select * From Servicio Where SerCodigo = " & IdServicio
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        If RsAux.EOF Then
            lDetalle.Visible = True
            lDetalle.Caption = "No existe ese servicio"
            tServicio.SelStart = 0: tServicio.SelLength = Len(tServicio.Text): tServicio.SetFocus
            RsAux.Close
            Exit Sub
        Else
            If RsAux!SerEstadoServicio <> EstadoS.Taller Then
                lDetalle.Visible = True
                lDetalle.Caption = "Estado distinto a Taller"
                tServicio.SelStart = 0: tServicio.SelLength = Len(tServicio.Text): tServicio.SetFocus
                RsAux.Close
                Exit Sub
            End If
            If Not EstaEnTaller(RsAux!SerCodigo) Then
                If RsAux!SerLocalIngreso <> paCodigoDeSucursal Then
                    lDetalle.Visible = True
                    lDetalle.Caption = "Sucursal no es donde se dio el ingreso."
                    tServicio.SelStart = 0: tServicio.SelLength = Len(tServicio.Text): tServicio.SetFocus
                    RsAux.Close
                    Exit Sub
                End If
                If RsAux!SerLocalIngreso = RsAux!SerLocalReparacion Then
                    lDetalle.Visible = True
                    lDetalle.Caption = "Se debe reparar EN SU Local."
                    tServicio.SelStart = 0: tServicio.SelLength = Len(tServicio.Text): tServicio.SetFocus
                    RsAux.Close
                    Exit Sub
                End If
                'CUMPLE ESTA CONDICION
                RsAux.Close
                Cons = "Select * From Servicio, Producto, Articulo, Sucursal, Local" _
                    & " Where SerCodigo = " & IdServicio _
                    & " And SerEstadoServicio = " & EstadoS.Taller _
                    & " And SerCodigo Not IN(Select TalServicio From Taller)" _
                    & " And SerLocalIngreso = " & paCodigoDeSucursal _
                    & " And SerLocalIngreso <> SerLocalReparacion And SerProducto = ProCodigo And ProArticulo = ArtID " _
                    & " And SerLocalIngreso = SucCodigo And SerLocalReparacion = LocCodigo"
                
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            Else
                If RsAux!SerLocalReparacion <> paCodigoDeSucursal Then
                    lDetalle.Visible = True
                    lDetalle.Caption = "No esta reparado EN SU Local."
                    tServicio.SelStart = 0: tServicio.SelLength = Len(tServicio.Text): tServicio.SetFocus
                    RsAux.Close
                    Exit Sub
                End If
                
                RsAux.Close
                
                'Lo cargo con la tabla taller
                Cons = "Select * From Servicio, Taller Where SerCodigo = " & IdServicio _
                        & " And SerCodigo = TalServicio"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                
                If IsNull(RsAux!TalFIngresoRecepcion) Then
                    lDetalle.Visible = True
                    lDetalle.Caption = "Servicio sin recepción de traslado IDA."
                    tServicio.SelStart = 0: tServicio.SelLength = Len(tServicio.Text): tServicio.SetFocus
                    RsAux.Close
                    Exit Sub
                End If
                If IsNull(RsAux!TalFReparado) Then
                    lDetalle.Visible = True
                    lDetalle.Caption = "Servicio sin reparar."
                    tServicio.SelStart = 0: tServicio.SelLength = Len(tServicio.Text): tServicio.SetFocus
                    RsAux.Close
                    Exit Sub
                End If
                If Not IsNull(RsAux!TalSalidaCamion) Then
                    lDetalle.Visible = True
                    lDetalle.Caption = "Servicio asignado a traslado Vuelta."
                    tServicio.SelStart = 0: tServicio.SelLength = Len(tServicio.Text): tServicio.SetFocus
                    RsAux.Close
                    Exit Sub
                End If
                
                If RsAux!SerLocalReparacion = RsAux!TalLocalAlCliente Then
                    lDetalle.Visible = True
                    lDetalle.Caption = "Ya esta en el local donde se entregara."
                    tServicio.SelStart = 0: tServicio.SelLength = Len(tServicio.Text): tServicio.SetFocus
                    RsAux.Close
                    Exit Sub
                End If
                RsAux.Close
                'Si llego aca es xq cumple la 2.
                Cons = "Select * from servicio where sercodigo = 0"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            End If
        End If
        
        If RsAux.EOF Then
        
            RsAux.Close
            
            iUno = 2
            'Consulto si esta en el local y reparado.
            Cons = "Select * From Servicio, Taller, Producto, Articulo, Sucursal, Local" _
                & " Where SerCodigo = " & IdServicio _
                & " And SerEstadoServicio = " & EstadoS.Taller _
                & " And TalFIngresoRecepcion Is Not Null And TalSalidaCamion  Is Null And TalFSalidaRealizado Is Null And TalFSalidaRecepcion Is Null " _
                & " And SerLocalReparacion = " & paCodigoDeSucursal & " And TalFReparado Is Not Null"
            
            If cReparar.ListIndex > -1 Then Cons = Cons & " And TalLocalAlCliente = " & cReparar.ItemData(cReparar.ListIndex)
                
            Cons = Cons & " And SerLocalReparacion <> TalLocalAlCliente And SerCodigo = TalServicio And SerProducto = ProCodigo And ProArticulo = ArtID " _
                    & " And SerLocalIngreso = SucCodigo And SerLocalReparacion = LocCodigo"
            
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        End If
        
        If Not RsAux.EOF Then
            
            If RsAux!SerEstadoServicio <> EstadoS.Taller Then
                MsgBox "El estado de este servicio no es taller.", vbInformation, "ATENCIÓN"
                RsAux.Close
                Exit Sub
            End If
        
            With vsConsulta
                If .IsSubtotal(.Rows - 1) Then .RemoveItem .Rows - 1
                .AddItem ""
                
                aModificacion = RsAux!SerModificacion
                .Cell(flexcpData, .Rows - 1, 0) = aModificacion
                If cListado.ListIndex = 0 Then
                    aModificacion = RsAux!TalModificacion
                    .Cell(flexcpData, .Rows - 1, 1) = aModificacion
                End If
                .Cell(flexcpData, .Rows - 1, 2) = 1 'Me digo que es para Reparar.
                
                If RsAux!SerCliente = EmpresaEmisora.Codigo Then .Cell(flexcpData, .Rows - 1, 3) = EmpresaEmisora.Codigo Else .Cell(flexcpData, .Rows - 1, 3) = 0
                aValor = RsAux!ProCodigo: .Cell(flexcpData, .Rows - 1, 4) = aValor
                aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 5) = aValor
                
                .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked
                .Cell(flexcpText, .Rows - 1, 1) = RsAux!SerCodigo
                .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ProCodigo, "(#,000)") & " " & Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!SerFecha, FormatoFP)
                .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!SucAbreviacion)
                .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!LocNombre)
                
                If iUno = 2 Then
                    If IsNull(RsAux!TalFReparado) Then
                        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio: .Cell(flexcpChecked, .Rows - 1, 0) = flexUnchecked
                    Else
                        If IsNull(RsAux!SerCostoFinal) Or IsNull(RsAux!TalFAceptacion) Then
                            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Inactivo: .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked
                        End If
                    End If
                End If
                
                .Sort = flexSortGenericDescending
                .Subtotal flexSTCount, -1, 2, "#,##0", Obligatorio, Colores.Rojo, True, "Total", , True
            End With
            lDetalle.Visible = False: lDetalle.Caption = ""
        End If
        RsAux.Close
        tServicio.SelStart = 0: tServicio.SelLength = Len(tServicio.Text): tServicio.SetFocus
    End If
End Sub

Private Function EstaCargado(ByVal idServ As Long) As Boolean
Dim iCont As Integer
    EstaCargado = False
    For iCont = 1 To vsConsulta.Rows - 1
        If Not vsConsulta.IsSubtotal(iCont) Then
            If vsConsulta.Cell(flexcpText, iCont, 1) = idServ Then
                EstaCargado = True
                Exit Function
            End If
        End If
    Next
End Function

Private Function EstaEnTaller(ByVal idServ As Long) As Boolean
Dim rsT As rdoResultset
    EstaEnTaller = False
    Cons = "Select * From Taller Where TalServicio = " & idServ
    Set rsT = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsT.EOF Then EstaEnTaller = True
    rsT.Close
End Function

Private Function BuscoUsuarioDigito(Digito As Long, Optional Codigo As Boolean = False, Optional Identificacion As Boolean = False, Optional Iniciales As Boolean = False) As Variant
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


