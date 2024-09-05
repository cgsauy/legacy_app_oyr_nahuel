VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACOMBO.OCX"
Begin VB.Form frmListado 
   Caption         =   "Listado de Retiros, Entregas, Traslados y Visitas de Servicios."
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLisHisServicio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4455
      Left            =   300
      TabIndex        =   8
      Top             =   780
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
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
      SelectionMode   =   0
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
      SubtotalPosition=   1
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
      Left            =   120
      TabIndex        =   22
      Top             =   780
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
      Begin VB.CommandButton bGrabar 
         Height          =   310
         Left            =   3240
         Picture         =   "frmLisHisServicio.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Grabar e Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmLisHisServicio.frx":0544
         Height          =   310
         Left            =   3960
         Picture         =   "frmLisHisServicio.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2700
         Picture         =   "frmLisHisServicio.frx":0B78
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2340
         Picture         =   "frmLisHisServicio.frx":0C62
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
         Picture         =   "frmLisHisServicio.frx":0D4C
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3600
         Picture         =   "frmLisHisServicio.frx":0F86
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5400
         Picture         =   "frmLisHisServicio.frx":1088
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
         Picture         =   "frmLisHisServicio.frx":118A
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
         Picture         =   "frmLisHisServicio.frx":148C
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
         Picture         =   "frmLisHisServicio.frx":17CE
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
         Picture         =   "frmLisHisServicio.frx":1AD0
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
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   15028
            Key             =   "msg"
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
      Height          =   675
      Left            =   50
      TabIndex        =   20
      Top             =   0
      Width           =   8655
      Begin VB.ComboBox cListado 
         Height          =   315
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin AACombo99.AACombo cCamion 
         Height          =   315
         Left            =   6720
         TabIndex        =   7
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
      Begin AACombo99.AACombo cTipo 
         Height          =   315
         Left            =   4500
         TabIndex        =   5
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
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
      Begin VB.TextBox tDesde 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   720
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "C&onsulta:"
         Height          =   195
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Camión:"
         Height          =   255
         Left            =   6060
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   4020
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
   End
   Begin ComctlLib.ImageList ImgList 
      Left            =   7860
      Top             =   2700
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
            Picture         =   "frmLisHisServicio.frx":1D0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLisHisServicio.frx":2024
            Key             =   "retiro"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private aTexto As String

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub bGrabar_Click()
    If cCamion.ListIndex = -1 Or vsConsulta.Rows = 1 Or cListado.ListIndex = 0 Then Exit Sub
    AccionGrabar
    AccionImprimir True
End Sub

Private Sub bImprimir_Click()
    If vsConsulta.Rows = 1 Then Exit Sub
    If cCamion.ListIndex > -1 Then
        If MsgBox("¿Confirma hacer una impresión previa?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    End If
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

Private Sub cCamion_Click()
    InicializoGrillas
End Sub

Private Sub cCamion_Change()
    InicializoGrillas
End Sub

Private Sub cCamion_GotFocus()
    With cCamion
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cCamion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub
Private Sub cCamion_LostFocus()
    cCamion.SelStart = 0
End Sub

Private Sub cListado_Click()
    InicializoGrillas
End Sub
Private Sub cListado_Change()
    InicializoGrillas
End Sub

Private Sub cListado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cTipo
End Sub

Private Sub cTipo_GotFocus()
    With cTipo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cCamion
End Sub
Private Sub cTipo_LostFocus()
    cCamion.SelStart = 0
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

    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    
    FechaDelServidor
    tDesde.Text = Format(Date, FormatoFP)
        
    cListado.AddItem "Impresos": cListado.ItemData(cListado.NewIndex) = 0
    cListado.AddItem "Pendientes": cListado.ItemData(cListado.NewIndex) = 1
    cListado.ListIndex = 1  'Por defecto Pendientes
    
    Cons = "Select CamCodigo, CamNombre From Camion"
    CargoCombo Cons, cCamion
    
    cTipo.AddItem "Entrega": cTipo.ItemData(cTipo.NewIndex) = TipoServicio.Entrega
    cTipo.AddItem "Retiro": cTipo.ItemData(cTipo.NewIndex) = TipoServicio.Retiro
    cTipo.AddItem "Visita": cTipo.ItemData(cTipo.NewIndex) = TipoServicio.Visita
    
    'Hoja Carta
    vsListado.Orientation = orPortrait: vsListado.PaperSize = 1
    
    InicializoGrillas
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
    
End Sub
Private Sub InicializoGrillas()
    On Error Resume Next
    With vsConsulta
        .Editable = True
        .Redraw = False
        .WordWrap = False
        .Cols = 1: .Rows = 1
        If cCamion.ListIndex > -1 Then
            .FormatString = "<Zona|<Hora|<Tipo|Flete|>ID|<Dirección|"
            .ColWidth(0) = 1800: .ColWidth(1) = 1100: .ColWidth(2) = 700: .ColWidth(3) = 1100: .ColWidth(4) = 800: .ColWidth(5) = 4000: .ColWidth(6) = 10
        Else
            .FormatString = "<Zona|<Hora|<Tipo|Flete|<Camión|>ID|<Dirección|"
            .ColWidth(0) = 1800: .ColWidth(1) = 1100: .ColWidth(2) = 700: .ColWidth(3) = 800: .ColWidth(4) = 1100: .ColWidth(5) = 800: .ColWidth(6) = 3400: .ColWidth(7) = 10
            .ColComboList(4) = "..."
        End If
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
            
            Case vbKeyI
                If vsConsulta.Rows = 1 Then Exit Sub
                If cCamion.ListIndex > -1 Then
                    If MsgBox("¿Confirma hacer una impresión previa?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
                End If
                AccionImprimir True
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            Case vbKeyC: AccionConfigurar
            
            Case vbKeyG
                If cCamion.ListIndex = -1 Or vsConsulta.Rows = 1 Or cListado.ListIndex = 0 Then Exit Sub
                AccionGrabar
                AccionImprimir True
            
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
    
    If Not IsDate(tDesde.Text) Then MsgBox "Debe ingresar una fecha desde.", vbInformation, "ATENCIÓN": Foco tDesde: Exit Sub
    
    Screen.MousePointer = 11
    chVista.Value = 0
    vsConsulta.Rows = 1
    vsConsulta.Refresh
    vsConsulta.Redraw = False
    CargoServicios
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True
End Sub

Private Sub tDesde_GotFocus()
    With tDesde
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda " Ingrese la fecha desde a consultar."
End Sub
Private Sub tDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cListado
End Sub
Private Sub tDesde_LostFocus()
    Ayuda ""
    If IsDate(tDesde.Text) Then tDesde.Text = Format(tDesde.Text, FormatoFP) Else tDesde.Text = ""
End Sub

Private Sub vsConsulta_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 4 Or cCamion.ListIndex > -1 Then Cancel = True
End Sub

Private Sub vsConsulta_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrBC
    'Presiono para cambiar el camión.
    Cons = "Select CamCodigo, Nombre = CamNombre  From Camion Where CamCodigo <> " & vsConsulta.Cell(flexcpData, Row, 3)
    Dim objLista As New clsListadeAyuda
    objLista.ActivoListaAyuda Cons, False, txtConexion, 4000
    If objLista.ValorSeleccionado > 0 Then
        If cListado.ListIndex = 0 Then If MsgBox("El servicio está impreso, confirma hacer el cambio de camión?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Set objLista = Nothing: Exit Sub
        'Updateo el camión en la tabla.
        Screen.MousePointer = 11
        Cons = "Select * From ServicioVisita Where VisServicio = " & vsConsulta.Cell(flexcpData, Row, 1)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux!VisFModificacion = CDate(vsConsulta.Cell(flexcpData, Row, 4)) Then
            FechaDelServidor
            RsAux.Edit
            RsAux!VisCamion = objLista.ValorSeleccionado
            RsAux!VisFModificacion = Format(gFechaServidor, sqlFormatoFH)
            RsAux.Update: RsAux.Close
            GoTo FinalizoCambio
        Else
            RsAux.Close
            MsgBox "Otra terminal modifico los datos con anterioridad, verifique.", vbExclamation, "ATENCIÓN"
        End If
    End If
    Set objLista = Nothing
    Screen.MousePointer = 0
    Exit Sub
FinalizoCambio:
    AccionConsultar
    Screen.MousePointer = 0
    Exit Sub
ErrBC:
    clsGeneral.OcurrioError "Ocurrio un error al intentar cambiar el camión.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
Dim Consulta As Boolean
    On Error GoTo errImprimir
    
    If cCamion.ListIndex = -1 Then Consulta = True Else Consulta = False
    
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub

    With vsListado
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    
    If Consulta Then
        EncabezadoListado vsListado, "Resumen de Servicios al " & Format(tDesde.Text, FormatoFP), False
        vsListado.filename = "Resumen de Servicios"
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hWnd: vsConsulta.ExtendLastCol = True
    Else
        vsListado.filename = "Impresión de Servicios"
        EncabezadoListado vsListado, "Resumen de Servicios al " & Format(tDesde.Text, FormatoFP) & "  Camión: " & cCamion.Text, False
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hWnd: vsConsulta.ExtendLastCol = True
        ImprimoRetiros
    End If
    vsListado.EndDoc
    
    If Imprimir Then
        frmSetup.pControl = vsListado
        frmSetup.Show vbModal, Me
        Me.Refresh
        If frmSetup.pOK Then vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub

Private Sub AccionConfigurar()
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
End Sub

Private Sub Ayuda(strTexto As String)
    Status.Panels("msg").Text = strTexto
End Sub

Private Sub CargoServicios()
    Dim aCamion As Long
    
    If cCamion.ListIndex > -1 Then aCamion = cCamion.ItemData(cCamion.ListIndex) Else aCamion = 0
    
    If cTipo.ListIndex = -1 Then
        'Cargo Todos.
        CargoRetiro aCamion
        CargoVisita aCamion
    Else
        Select Case cTipo.ItemData(cTipo.ListIndex)
            Case TipoServicio.Entrega
            Case TipoServicio.Retiro: CargoRetiro aCamion
            Case TipoServicio.Visita: CargoVisita aCamion
        End Select
    End If
End Sub

Private Sub CargoRetiro(Camion As Long)
Dim aCod As Long
Dim FModificacion As String

    Cons = "Select * From ServicioVisita, Servicio, Producto, TipoFlete, Zona, Camion" _
        & " Where VisTipo = " & TipoServicio.Retiro _
        & " And VisFecha = '" & Format(tDesde.Text, sqlFormatoF) & "'" _
        & " And VisSinEfecto = 0 "
    
    If cListado.ListIndex = 1 Then Cons = Cons & "And VisFImpresion = Null " Else Cons = Cons & "And VisFImpresion <> Null "
    If Camion > 0 Then Cons = Cons & " And VisCamion = " & Camion
    
    Cons = Cons & "And SerEstadoServicio = " & EstadoS.Retiro _
        & " And VisServicio = SerCodigo And SerProducto = ProCodigo " _
        & " And VisTipoFlete = TFlCodigo And VisZona = ZonCodigo And VisCamion = CamCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsAux.EOF
        With vsConsulta
            .AddItem ""
            
            'DATA
            .Cell(flexcpData, .Rows - 1, 0) = TipoServicio.Retiro
            aCod = RsAux!SerCodigo
            .Cell(flexcpData, .Rows - 1, 1) = aCod
            FModificacion = RsAux!SerModificacion
            .Cell(flexcpData, .Rows - 1, 2) = FModificacion
            aCod = RsAux!CamCodigo
            .Cell(flexcpData, .Rows - 1, 3) = aCod
            FModificacion = RsAux!VisFModificacion
            .Cell(flexcpData, .Rows - 1, 4) = FModificacion
            aCod = RsAux!VisCodigo
            .Cell(flexcpData, .Rows - 1, 5) = aCod
            
            'Imagenes.
            '.Cell(flexcpPicture, .Rows - 1, 0) = ImgList.ListImages("retiro").ExtractIcon
            
            'TEXTO
            .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!ZonNombre)
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!VisHorario)
            .Cell(flexcpText, .Rows - 1, 2) = "Retiro"
            .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!TFlDescripcion)
            
            If cCamion.ListIndex = -1 Then
                .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!CamNombre)
                .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!SerCodigo)
                If Not IsNull(RsAux!ProDireccion) Then
                    .Cell(flexcpText, .Rows - 1, 6) = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!ProDireccion, localidad:=True)
                End If
            Else
                .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!SerCodigo)
                If Not IsNull(RsAux!ProDireccion) Then
                    .Cell(flexcpText, .Rows - 1, 5) = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!ProDireccion, localidad:=True)
                End If
            End If
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub
Private Sub CargoVisita(Camion As Long)
Dim aCod As Long
Dim FModificacion As String

    Cons = "Select * From ServicioVisita, Servicio, Producto, Zona, Camion" _
        & " Where VisTipo = " & TipoServicio.Visita _
        & " And VisFecha = '" & Format(tDesde.Text, sqlFormatoF) & "'" _
        & " And VisSinEfecto = 0 "
    
    If cListado.ListIndex = 1 Then Cons = Cons & "And VisFImpresion = Null " Else Cons = Cons & "And VisFImpresion <> Null "
    If Camion > 0 Then Cons = Cons & " And VisCamion = " & Camion
    
    Cons = Cons & " And SerEstadoServicio = " & EstadoS.Visita _
        & " And VisServicio = SerCodigo And SerProducto = ProCodigo " _
        & " And VisZona = ZonCodigo And VisCamion = CamCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        With vsConsulta
            .AddItem ""
            
            'DATA
            .Cell(flexcpData, .Rows - 1, 0) = TipoServicio.Visita
            
            aCod = RsAux!SerCodigo
            .Cell(flexcpData, .Rows - 1, 1) = aCod
            
            FModificacion = RsAux!SerModificacion
            .Cell(flexcpData, .Rows - 1, 2) = FModificacion
            
            aCod = RsAux!CamCodigo
            .Cell(flexcpData, .Rows - 1, 3) = aCod
            
            FModificacion = RsAux!VisFModificacion
            .Cell(flexcpData, .Rows - 1, 4) = FModificacion
            
            aCod = RsAux!VisCodigo
            .Cell(flexcpData, .Rows - 1, 5) = aCod
            
            'Imagenes.
            '.Cell(flexcpPicture, .Rows - 1, 0) = ImgList.ListImages("retiro").ExtractIcon
            
            'TEXTO
            .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!ZonNombre)
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!VisHorario)
            .Cell(flexcpText, .Rows - 1, 2) = "Visita"
            If cCamion.ListIndex = -1 Then
                .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!CamNombre)
                .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!SerCodigo)
                If Not IsNull(RsAux!ProDireccion) Then
                    .Cell(flexcpText, .Rows - 1, 6) = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!ProDireccion, localidad:=True)
                End If
            Else
                    .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!SerCodigo)
                    If Not IsNull(RsAux!ProDireccion) Then
                    .Cell(flexcpText, .Rows - 1, 5) = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!ProDireccion, localidad:=True)
                End If
            End If
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub

Private Sub ImprimoRetiros()
    With vsConsulta
        For I = 1 To .Rows - 1
            
        Next I
    End With
End Sub

Private Sub AccionGrabar()
        
    If MsgBox("¿Confirma grabar como impresos los sevicios consultados?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrRB
    
    With vsConsulta
        For I = 1 To .Rows - 1
            
            'TABLA SERVICIO
            Cons = "Select * From Servicio Where SerCodigo = " & Val(.Cell(flexcpData, I, 1))
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If RsAux.EOF Then
                RsAux.Close: RsAux.Edit 'Provoco error.
            Else
                If RsAux!SerModificacion = CDate(.Cell(flexcpData, I, 2)) Then
                    RsAux.Edit
                    RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
                    RsAux.Update
                    RsAux.Close
                Else
                    RsAux.Close: RsAux.Edit 'Provoco error.
                End If
            End If
            
            'TABLA SERVICIOVISITA
            Cons = "Select * From ServicioVisita Where VisCodigo = " & Val(.Cell(flexcpData, I, 5))
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If RsAux.EOF Then
                RsAux.Close: RsAux.Edit 'Provoco error.
            Else
                If RsAux!VisFModificacion = CDate(.Cell(flexcpData, I, 4)) Then
                    RsAux.Edit
                    RsAux!VisFModificacion = Format(gFechaServidor, sqlFormatoFH)
                    RsAux!VisFImpresion = Format(gFechaServidor, sqlFormatoFH)
                    RsAux.Update
                    RsAux.Close
                Else
                    RsAux.Close: RsAux.Edit 'Provoco error.
                End If
            End If
            
        Next I
    End With
    cBase.CommitTrans
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrRB:
    Resume ErrResumo
ErrResumo:
    clsGeneral.OcurrioError "Ocurrio un error al intentar grabar los datos, se cargarán los mismos nuevamente.", Err.Description
    AccionConsultar
    Screen.MousePointer = 0
    Exit Sub
End Sub
