VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListado 
   Caption         =   "Estadística de Servicios"
   ClientHeight    =   7125
   ClientLeft      =   765
   ClientTop       =   2325
   ClientWidth     =   10950
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
   ScaleHeight     =   7125
   ScaleWidth      =   10950
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
      Height          =   915
      Left            =   50
      TabIndex        =   27
      Top             =   0
      Width           =   10635
      Begin VB.ComboBox cFecha 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   2295
      End
      Begin VB.CheckBox chVerGrales 
         Caption         =   "Ver Comentarios Grales."
         Height          =   195
         Left            =   8700
         TabIndex        =   6
         Top             =   280
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin VB.ComboBox cEstado 
         Height          =   315
         Left            =   8700
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   540
         Width           =   1755
      End
      Begin AACombo99.AACombo cGrupo 
         Height          =   315
         Left            =   5340
         TabIndex        =   10
         Top             =   540
         Width           =   2415
         _ExtentX        =   4260
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
      End
      Begin VB.TextBox tHasta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   3
         Top             =   220
         Width           =   915
      End
      Begin VB.TextBox tDesde 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   1
         Top             =   220
         Width           =   915
      End
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5340
         TabIndex        =   5
         Top             =   220
         Width           =   3255
      End
      Begin AACombo99.AACombo cLocalR 
         Height          =   315
         Left            =   1020
         TabIndex        =   8
         Top             =   540
         Width           =   2355
         _ExtentX        =   4154
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
      End
      Begin VB.Label Label6 
         Caption         =   "&Estado:"
         Height          =   255
         Left            =   8040
         TabIndex        =   11
         Top             =   585
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "&Reparado:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   580
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "&Grupo:"
         Height          =   195
         Left            =   4680
         TabIndex        =   9
         Top             =   585
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "a&l"
         Height          =   195
         Left            =   3420
         TabIndex        =   2
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label4 
         Caption         =   "&Artículo:"
         Height          =   195
         Left            =   4680
         TabIndex        =   4
         Top             =   240
         Width           =   675
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsServicio 
      Height          =   1635
      Left            =   60
      TabIndex        =   13
      Top             =   1620
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2884
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
      SubtotalPosition=   1
      OutlineBar      =   1
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
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
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   60
      ScaleHeight     =   435
      ScaleWidth      =   6075
      TabIndex        =   26
      Top             =   4380
      Width           =   6135
      Begin VB.CommandButton bExportar 
         Height          =   310
         Left            =   4560
         Picture         =   "frmListado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Exportar"
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":074C
         Height          =   310
         Left            =   4140
         Picture         =   "frmListado.frx":084E
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2760
         Picture         =   "frmListado.frx":0D80
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2400
         Picture         =   "frmListado.frx":0E6A
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmListado.frx":0F54
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3780
         Picture         =   "frmListado.frx":118E
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   4920
         Picture         =   "frmListado.frx":1290
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   180
         Picture         =   "frmListado.frx":1392
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmListado.frx":1694
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmListado.frx":19D6
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmListado.frx":1CD8
         Style           =   1  'Graphical
         TabIndex        =   15
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
      TabIndex        =   24
      Top             =   6870
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   18785
            Key             =   "msg"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   1695
      Left            =   60
      TabIndex        =   25
      Top             =   1620
      Width           =   7335
      _Version        =   196608
      _ExtentX        =   12938
      _ExtentY        =   2990
      _StockProps     =   229
      BorderStyle     =   1
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
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin ComctlLib.TabStrip tabQuery 
      Height          =   1575
      Left            =   120
      TabIndex        =   28
      Top             =   2100
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   2778
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Detalle de Servicios "
            Key             =   "detalle"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Repuestos Utilizados  "
            Key             =   "repuestos"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsRepuesto 
      Height          =   1635
      Left            =   4380
      TabIndex        =   29
      Top             =   1920
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2884
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
      SubtotalPosition=   1
      OutlineBar      =   1
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
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
   Begin MSComDlg.CommonDialog ctlDlg 
      Left            =   8400
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ImgList 
      Left            =   7620
      Top             =   3660
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
            Picture         =   "frmListado.frx":1F12
            Key             =   "cliente"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmListado.frx":222C
            Key             =   "stock"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuBDerecho 
      Caption         =   "mnuBDerecho"
      Visible         =   0   'False
      Begin VB.Menu mnuOpcionX 
         Caption         =   ""
         Index           =   0
      End
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

Private Sub bExportar_Click()
On Error GoTo errCancel
    
    With ctlDlg
        .CancelError = True
        
        .FileName = "EstadisticasServicios"
        
        .Filter = "Libro de Microsoft Exel|*.xls|" & _
                     "Texto (delimitado por tabulaciones)|*.txt|" & "Texto (delimitado por comas)|*.txt"
        
        .ShowSave
        
        'Confirma exportar el contenido de la lista al archivo:
        If MsgBox("Confirma exportar el contenido de la lista al archivo: " & .FileName, vbQuestion + vbYesNo) = vbYes Then
        
            On Error GoTo errSaving
            Screen.MousePointer = 11
            Me.Refresh
            DoEvents
            
            Dim mSSetting As SaveLoadSettings
            
            Select Case .FilterIndex
                Case 1: mSSetting = flexFileTabText
                Case 2: mSSetting = flexFileTabText
                Case 3: mSSetting = flexFileCommaText
            End Select
            
            vsServicio.SaveGrid .FileName, mSSetting, True
                
            Screen.MousePointer = 0
        End If
        
    End With
    
errCancel:
    Screen.MousePointer = 0
    Exit Sub
errSaving:
     clsGeneral.OcurrioError "Error al exportar el contenido de la lista.", Err.Description
    Screen.MousePointer = 0
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

Private Sub cEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub cFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tDesde.SetFocus
End Sub

Private Sub cGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cEstado.SetFocus
End Sub

Private Sub chVerGrales_Click()
    vsServicio.ColHidden(8) = (chVerGrales.Value = vbUnchecked)
End Sub

Private Sub chVista_Click()
    If chVista.Value = 0 Then
        vsListado.ZOrder 1
        Me.Refresh
    Else
        AccionImprimir
        vsListado.ZOrder 0
        Me.Refresh
    End If
End Sub

Private Sub cLocalR_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cGrupo.SetFocus
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub
Private Sub Form_Load()

    On Error GoTo ErrLoad

    ObtengoSeteoForm Me, 100, 100, 10730, 7000
    picBotones.BorderStyle = vbBSNone
    
    Cons = "Select GruCodigo, GruNombre from Grupo Order by GruNombre"
    CargoCombo Cons, cGrupo
    
    Cons = "Select SucCodigo, SucAbreviacion from Sucursal order by SucAbreviacion"
    CargoCombo Cons, cLocalR
    
    cEstado.AddItem "(Todos)": cEstado.ItemData(cEstado.NewIndex) = 0
    cEstado.AddItem "Sin Cargo": cEstado.ItemData(cEstado.NewIndex) = EstadoP.SinCargo
    cEstado.AddItem "Fuera Garantía": cEstado.ItemData(cEstado.NewIndex) = EstadoP.FueraGarantia
    cEstado.ListIndex = 0
    
    cFecha.AddItem "Recepcionado ó Reparado": cFecha.ItemData(cFecha.NewIndex) = 0
    cFecha.AddItem "Recepcionado": cFecha.ItemData(cFecha.NewIndex) = 1
    cFecha.AddItem "Reparado": cFecha.ItemData(cFecha.NewIndex) = 2
    cFecha.ListIndex = 0
    
    InicializoGrillas
    
    With mnuOpcionX
        .Item(0).Caption = "Filtrar datos": .Item(0).Tag = 0
        Load .Item(1): .Item(1).Caption = "-": .Item(1).Tag = 0
        Load .Item(2): .Item(2).Caption = "": .Item(2).Tag = ""
        Load .Item(3): .Item(3).Caption = "": .Item(3).Tag = ""
        Load .Item(4): .Item(4).Caption = "": .Item(4).Tag = ""
        Load .Item(5): .Item(5).Caption = "": .Item(5).Tag = ""
        Load .Item(6): .Item(6).Caption = "": .Item(6).Tag = ""
        Load .Item(7): .Item(7).Caption = "-": .Item(7).Tag = 0
        Load .Item(8): .Item(8).Caption = "-": .Item(8).Tag = 0
        Load .Item(9): .Item(9).Caption = "Cancelar": .Item(9).Tag = 0
    End With

    FechaDelServidor
    tHasta.Text = Format(Date, "dd/mm/yyyy")
    
    With vsListado
        .Orientation = orPortrait
        .PaperSize = 1
        .MarginTop = 750: .MarginBottom = 750
        .MarginLeft = 700: .MarginRight = 500
    End With
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()
    On Error Resume Next
    
    With vsServicio
        .Redraw = False
        .WordWrap = False
        .Cols = 1: .Rows = 1
        .FormatString = "<Servicio|<Artículo|<Obs. Taller|>Dem. p/Reparar|^Recepcionado|^Reparado|^Estado|>Precio|<Nro Serie|<Comentarios Grales.|<Motivos de Ingreso|<Repuestos"
        .ColWidth(0) = 700: .ColWidth(1) = 3200: .ColWidth(2) = 3200: .ColWidth(4) = 1100
        .ColWidth(6) = 800: .ColWidth(7) = 1200:: .ColWidth(8) = 1300: .ColWidth(10) = 3000: .ColWidth(11) = 3000
        
        .Redraw = True
    End With
    
    With vsRepuesto
        .Redraw = False
        .WordWrap = False
        .Cols = 1: .Rows = 1
        .FormatString = "<Repuesto|>Q|"
        .ColWidth(0) = 3700: .ColWidth(1) = 1200
        
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
            
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11
    
    tabQuery.Top = fFiltros.Top + fFiltros.Height + 60
    tabQuery.Left = 60
    tabQuery.Height = Me.ScaleHeight - (tabQuery.Top + Status.Height + picBotones.Height + 30)
    tabQuery.Width = Me.ScaleWidth - (tabQuery.Left * 2)
    
    vsListado.Top = tabQuery.ClientTop
    vsListado.Left = tabQuery.ClientLeft
    vsListado.Width = tabQuery.ClientWidth
    vsListado.Height = tabQuery.ClientHeight
    
    picBotones.Top = tabQuery.Height + tabQuery.Top + 30
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    
    vsServicio.Top = vsListado.Top: vsServicio.Width = vsListado.Width: vsServicio.Height = vsListado.Height: vsServicio.Left = vsListado.Left
    vsRepuesto.Top = vsListado.Top: vsRepuesto.Width = vsListado.Width: vsRepuesto.Height = vsListado.Height: vsRepuesto.Left = vsListado.Left
    
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
    
    On Error GoTo errConsultar
    
    'Valido los datos ingresados-------------------------------------------------------------------------------------
    If Not IsDate(tDesde.Text) Then
        MsgBox "La fecha desde ingresada no es correcta.", vbExclamation, "Faltan datos"
        Foco tDesde: Exit Sub
    End If
    If Not IsDate(tHasta.Text) Then
        MsgBox "La fecha hasta ingresada no es correcta.", vbExclamation, "Faltan datos"
        Foco tHasta: Exit Sub
    End If
    If CDate(tHasta.Text) < CDate(tDesde.Text) Then
        MsgBox "El período de fechas ingresado no es correcto.", vbExclamation, "Error en Fechas"
        Foco tDesde: Exit Sub
    End If
    '--------------------------------------------------------------------------------------------------------------------
    
    Screen.MousePointer = 11
    chVista.Value = 0
    vsServicio.Tag = ""
    
    InicializoGrillas

    Select Case tabQuery.SelectedItem.Key
        Case "repuestos"
                        vsRepuesto.Rows = 1: vsRepuesto.Refresh: vsRepuesto.Redraw = False
                        CargoRepuestos
                        vsRepuesto.Redraw = True
                        
        Case "detalle"
                        vsServicio.Rows = 1: vsServicio.Refresh: vsServicio.Redraw = False
                        CargoServicios
                        vsServicio.Redraw = True
    End Select
    
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    vsServicio.Redraw = True: vsRepuesto.Redraw = True
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Sub mnuOpcionX_Click(Index As Integer)
On Error GoTo errFiltro
Dim idX As Integer, mCol As Integer, mTXT As String

    Select Case Index
        Case 2: mCol = 1: mTXT = mnuOpcionX(Index).Tag
        Case 3: mCol = 3: mTXT = "*" & mnuOpcionX(Index).Tag & "*"
        
        Case 4: mCol = 6: mTXT = mnuOpcionX(Index).Tag
        Case 5: mCol = 7: mTXT = mnuOpcionX(Index).Tag
        Case 6: mCol = 8: mTXT = mnuOpcionX(Index).Tag
        Case 7: mCol = 4: mTXT = mnuOpcionX(Index).Tag
        
        Case Else: Exit Sub
        
    End Select
    
    With vsServicio
        Dim xRow As Integer
        xRow = .FixedRows
        
        .Subtotal flexSTClear
        
        For idX = .FixedRows To .Rows - 1
            If .Cell(flexcpText, xRow, mCol) Like mTXT Then
                xRow = xRow + 1
            Else
                .RemoveItem xRow
            End If
        Next
        
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTCount, -1, 1, "0", Colores.Azul, Colores.Blanco, True, "Totales"
        .Subtotal flexSTSum, -1, 7
        
        Dim mQVtas As Long, mQServicios As Long
        mQVtas = Val(vsServicio.Tag)
        If mQVtas <> 0 Then
            mQServicios = .Cell(flexcpValue, .Rows - 1, 1)
            .Cell(flexcpText, .Rows - 1, 1) = mQServicios & "  " & Format((mQServicios * 100) / mQVtas, "0.00") & "%" & "  Vtas= " & mQVtas
        End If
    
    End With
    
    Exit Sub
errFiltro:
    clsGeneral.OcurrioError "Error al aplicar el filtro seleccionado.", Err.Description
End Sub

Private Sub tabQuery_Click()
    
    Select Case tabQuery.SelectedItem.Key
        Case "repuestos": vsRepuesto.ZOrder 0
        Case "detalle": vsServicio.ZOrder 0
    End Select
    chVista.Value = 0
End Sub

Private Sub tArticulo_Change()
    If Val(tArticulo.Tag) = 0 Then Exit Sub
    tArticulo.Tag = 0
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tArticulo_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrTA
    
    If KeyCode = vbKeyReturn And Trim(tArticulo.Text) <> "" Then
        
        If Val(tArticulo.Tag) <> 0 Then
            cLocalR.SetFocus: Exit Sub
        End If
        
        Screen.MousePointer = 11
        If Not IsNumeric(tArticulo.Text) Then   'Busqueda por nombre
            Cons = "Select ArtID, 'Nombre' = ArtNombre, 'Código' = ArtCodigo From Articulo " _
                    & " Where ArtNombre Like '" & Replace(tArticulo.Text, " ", "%") & "%'" _
                    & " Order by ArtNombre"
            
            Dim LiAyuda  As New clsListadeAyuda
            LiAyuda.ActivoListaAyuda Cons, False, txtConexion, 4800
            Me.Refresh
            If LiAyuda.ItemSeleccionado <> "" Then
                tArticulo.Text = LiAyuda.ItemSeleccionado
                tArticulo.Tag = LiAyuda.ValorSeleccionado
            End If
            Set LiAyuda = Nothing
        
        Else                                            'Busqueda por codigo
            Cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtCodigo = " & Val(tArticulo.Text)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
            If RsAux.EOF Then
                MsgBox "No se encontró un artículo para el código ingresado.", vbInformation, "ATENCIÓN"
            Else
                tArticulo.Text = Trim(RsAux!Nombre)
                tArticulo.Tag = RsAux!ArtId
            End If
            RsAux.Close
        End If
        
        Screen.MousePointer = 0
    Else
        If KeyCode = vbKeyReturn Then cLocalR.SetFocus
    End If
    Exit Sub

ErrTA:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tDesde_GotFocus()
    With tDesde: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tDesde_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If IsDate(tDesde.Text) Then
            tDesde.Text = Format(tDesde.Text, "dd/mm/yyyy")
            'tHasta.Text = Format(UltimoDia(CDate(tDesde.Text)), "dd/mm/yyyy")
            Foco tHasta
        End If
    End If
    
End Sub

Private Sub tDesde_LostFocus()
    If IsDate(tDesde.Text) Then tDesde.Text = Format(tDesde.Text, "dd/mm/yyyy")
End Sub

Private Sub tHasta_GotFocus()
    With tHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tHasta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If IsDate(tHasta.Text) Then
            tHasta.Text = Format(tHasta.Text, "dd/mm/yyyy")
            Foco tArticulo
        End If
    End If
    
End Sub

Private Sub vsServicio_DblClick()

    If vsServicio.Rows = 1 Then Exit Sub
    
    EjecutarApp App.Path & "\Seguimiento de Servicios", vsServicio.Cell(flexcpText, vsServicio.Row, 0)
    
End Sub

Private Sub vsServicio_GotFocus()
    Status.Panels(1).Text = "[Doble Clik] Ir a Seguimiento de Servicios."
End Sub

Private Sub vsServicio_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    Select Case tabQuery.SelectedItem.Key
        Case "repuestos": vsListado.Columns = 2
        Case "detalle": vsListado.Columns = 1
    End Select

    With vsListado
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "Error de Impresión"
            Screen.MousePointer = 0: Exit Sub
        End If
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim aTexto As String
    aTexto = " (" & tDesde.Text & " al " & tHasta.Text & ")"
    If cGrupo.ListIndex <> -1 Then aTexto = aTexto & " GR:" & Trim(cGrupo.Text)
    If cLocalR.ListIndex <> -1 Then aTexto = aTexto & "  REP:" & Trim(cLocalR.Text)
    
    Select Case tabQuery.SelectedItem.Key
        Case "repuestos": aTexto = "Servicios - Repuestos Cambiados" & aTexto
        Case "detalle": aTexto = "Servicios - Detalle" & aTexto
    End Select
    
    EncabezadoListado vsListado, aTexto, False
    
    vsListado.FileName = "Estadística de Servicios"
    Select Case tabQuery.SelectedItem.Key
        Case "repuestos": vsRepuesto.ExtendLastCol = False: vsListado.RenderControl = vsRepuesto.hwnd: vsRepuesto.ExtendLastCol = True
        Case "detalle": vsServicio.ExtendLastCol = False: vsListado.RenderControl = vsServicio.hwnd: vsServicio.ExtendLastCol = True
    End Select
    
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

Private Sub CargoServicios()

    Dim aTexto As String
    Dim rsVis As rdoResultset
    Cons = "Select * From Servicio Left Outer Join Taller On SerCodigo = TalServicio " & _
                            " , Producto, Articulo" & _
                " Where SerProducto = ProCodigo " & _
                " And ProArticulo = ArtID " & _
                " And ((IsNull(TalServicio,0)+SerEstadoServicio =" & EstadoS.Taller & ") OR (TalServicio Is NOT Null) )"
                '" And ( ( TalServicio Is Null And SerEstadoServicio = " & EstadoS.Taller & ") OR (TalServicio Is NOT Null) )"
    
    If Val(tArticulo.Tag) <> 0 Then Cons = Cons & " And ProArticulo = " & Val(tArticulo.Tag)
    If cGrupo.ListIndex <> -1 Then Cons = Cons & " And ProArticulo IN (Select AGrArticulo From ArticuloGrupo Where AGrGrupo = " & Val(cGrupo.ItemData(cGrupo.ListIndex)) & ")"
    
    If cLocalR.ListIndex <> -1 Then Cons = Cons & " And SerLocalReparacion = " & Val(cLocalR.ItemData(cLocalR.ListIndex))
    
    If cEstado.ListIndex <> -1 Then
        If cEstado.ItemData(cEstado.ListIndex) <> 0 Then
            Cons = Cons & " And SerEstadoProducto = " & cEstado.ItemData(cEstado.ListIndex)
        End If
    End If
    
    aTexto = " Between '" & Format(tDesde.Text, sqlFormatoF) & "' AND '" & Format(tHasta.Text, sqlFormatoF) & " 23:59:00'"
    Select Case cFecha.ItemData(cFecha.ListIndex)
        Case 0  '   "Recepcionado ó Reparado"
                    Cons = Cons & " And ( (SerFecha " & aTexto & ") OR (TalFReparado " & aTexto & ") ) "
                    
        Case 1  '   "Recepcionado"
                    Cons = Cons & " And SerFecha " & aTexto
        Case 2  '   "Reparado"
                    Cons = Cons & " And TalFReparado " & aTexto
    End Select
    
    Cons = Cons & " Order by SerFecha ASC"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
    
        With vsServicio
            '<Servicio|<Artículo|<Obs. Taller|>Dem. p/Reparar|^Recepcionado|^Reparado|^Estado|>Precio|<Nro Serie|<Comentarios Grales.|<Motivos de Ingreso
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = RsAux!SerCodigo
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ArtNombre)
            
            .Cell(flexcpText, .Rows - 1, 6) = EstadoProducto(RsAux!SerEstadoProducto, True)
            If RsAux!ProCliente = paClienteEmpresa Then .Cell(flexcpText, .Rows - 1, 6) = .Cell(flexcpText, .Rows - 1, 6) & " (SK)"
            If Not IsNull(RsAux!TalComentario) Then
                .Cell(flexcpText, .Rows - 1, 2) = zfn_QuitoClaves(Replace(Trim(RsAux!TalComentario), vbCrLf, " "))
            End If
            
            If Not IsNull(RsAux!TalFReparado) Then
                .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!TalFReparado, "dd/mm/yy")
            End If

            'Demora en Reparar  -> desde Q trajo el cliente hasta que se lo pusimos en el local p/retirar--------------------------------------------------------- !!!
            If RsAux!ProCliente = paClienteEmpresa Then
                If Not IsNull(RsAux!TalFReparado) Then
                    '.Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!TalFReparado, "dd/mm/yy")
                    aTexto = DateDiff("d", RsAux!SerFecha, RsAux!TalFReparado)
                Else
                    aTexto = DateDiff("d", RsAux!SerFecha, gFechaServidor)
                    .Cell(flexcpForeColor, .Rows - 1, 3) = Colores.Rojo: .Cell(flexcpFontItalic, .Rows - 1, 3) = True
                End If
            Else
                If Not IsNull(RsAux!TalFSalidaRecepcion) Then
                    aTexto = DateDiff("d", RsAux!SerFecha, RsAux!TalFSalidaRecepcion)
                Else
                    aTexto = DateDiff("d", RsAux!SerFecha, gFechaServidor)
                    .Cell(flexcpForeColor, .Rows - 1, 3) = Colores.Rojo: .Cell(flexcpFontItalic, .Rows - 1, 3) = True
                End If
            End If
            .Cell(flexcpText, .Rows - 1, 3) = aTexto
            
            If Not IsNull(RsAux!SerFecha) Then .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!SerFecha, "dd/mm/yy")
            
            If Not IsNull(RsAux!SerCostoFinal) Then .Cell(flexcpText, .Rows - 1, 7) = Format(RsAux!SerCostoFinal, FormatoMonedaP)
            If Not IsNull(RsAux!ProNroSerie) Then .Cell(flexcpText, .Rows - 1, 8) = Trim(RsAux!ProNroSerie) Else .Cell(flexcpText, .Rows - 1, 8) = "N/D"
            
            If Not IsNull(RsAux!SerComentario) Then .Cell(flexcpText, .Rows - 1, 9) = Trim(RsAux!SerComentario)
            
            .Cell(flexcpText, .Rows - 1, 10) = fnc_MotivosServicio(RsAux!SerCodigo)
            .Cell(flexcpText, .Rows - 1, 11) = fnc_RepuestosServicio(RsAux!SerCodigo)
        End With
        RsAux.MoveNext
        
    Loop
    RsAux.Close
    
    If vsServicio.Rows > 1 Then
        vsServicio.ColDataType(4) = flexDTDate
         With vsServicio
            .SubtotalPosition = flexSTBelow
            .Subtotal flexSTCount, -1, 1, "0", Colores.Azul, Colores.Blanco, True, "Totales"
            .Subtotal flexSTSum, -1, 7
                        
            Dim mQVtas As Long, mQServicios As Long
            mQVtas = fnc_ProcesoVentas
            If mQVtas <> 0 Then
                mQServicios = .Cell(flexcpValue, .Rows - 1, 1)
                .Cell(flexcpText, .Rows - 1, 1) = mQServicios & "  " & Format((mQServicios * 100) / mQVtas, "0.00") & "%" & "  Vtas= " & mQVtas

            End If
            vsServicio.Tag = mQVtas
            
         End With
    End If
    
    vsServicio.ColHidden(8) = (chVerGrales.Value = vbUnchecked)
    vsServicio.ColPosition(10) = 2
    vsServicio.ColPosition(11) = 4
    
End Sub

Private Sub CargoRepuestos()

    Dim aTexto As String
    Dim rsVis As rdoResultset
    
    Cons = "Select ArtCodigo, ArtNombre, Sum(SReCantidad) as Q From Servicio Left Outer Join Taller On SerCodigo = TalServicio " & _
                            " , ServicioRenglon, Producto, Articulo" & _
                " Where SerProducto = ProCodigo " & _
                " And SerCodigo = SReServicio " & _
                " And SReMotivo = ArtID" & _
                " And SReTipoRenglon = " & TipoRenglonS.Cumplido & _
                " And ( ( TalServicio Is Null And SerEstadoServicio = " & EstadoS.Taller & ") OR (TalServicio Is NOT Null) )"
    
    If Val(tArticulo.Tag) <> 0 Then Cons = Cons & " And ProArticulo = " & Val(tArticulo.Tag)
    If cGrupo.ListIndex <> -1 Then Cons = Cons & " And ProArticulo IN (Select AGrArticulo From ArticuloGrupo Where AGrGrupo = " & Val(cGrupo.ItemData(cGrupo.ListIndex)) & ")"
    If cLocalR.ListIndex <> -1 Then Cons = Cons & " And SerLocalReparacion = " & Val(cLocalR.ItemData(cLocalR.ListIndex))
    
    If cEstado.ListIndex <> -1 Then
        If cEstado.ItemData(cEstado.ListIndex) <> 0 Then
            Cons = Cons & " And SerEstadoProducto = " & cEstado.ItemData(cEstado.ListIndex)
        End If
    End If
    
    aTexto = " Between '" & Format(tDesde.Text, sqlFormatoF) & "' AND '" & Format(tHasta.Text, sqlFormatoF) & " 23:59:00'"
    Select Case cFecha.ItemData(cFecha.ListIndex)
        Case 0  '   "Recepcionado ó Reparado"
                    Cons = Cons & " And ( (SerFecha " & aTexto & ") OR (TalFReparado " & aTexto & ") ) "
                    
        Case 1  '   "Recepcionado"
                    Cons = Cons & " And SerFecha " & aTexto
        Case 2  '   "Reparado"
                    Cons = Cons & " And TalFReparado " & aTexto
    End Select
    
    Cons = Cons & " Group by ArtCodigo, ArtNombre"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
    
        With vsRepuesto
            If Not IsNull(RsAux!Q) Then
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 1) = RsAux!Q
            End If
        End With
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If vsRepuesto.Rows > 1 Then
         With vsRepuesto
            .Select 1, 0
            .Sort = flexSortGenericAscending
            
            .SubtotalPosition = flexSTBelow
            .Subtotal flexSTSum, -1, 1, "0", Colores.Azul, Colores.Blanco, True, "Totales"
            
            
         End With
    End If
    
End Sub

Private Function fnc_MotivosServicio(xIdServicio As Long) As String
On Error GoTo errMotivos
Dim rsM As rdoResultset
Dim retSTR As String

    Cons = "Select MSeNombre from ServicioRenglon, MotivoServicio" & _
                " Where SReServicio = " & xIdServicio & _
                " And SReTipoRenglon = 1 And MSeID = SReMotivo"
    Set rsM = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsM.EOF
        retSTR = retSTR & IIf(retSTR = "", "", "; ") & Trim(rsM!MSeNombre)
        rsM.MoveNext
    Loop
    rsM.Close

    fnc_MotivosServicio = retSTR
    
errMotivos:
End Function

Private Function fnc_RepuestosServicio(xIdServicio As Long) As String
On Error GoTo errMotivos
Dim rsM As rdoResultset
Dim retSTR As String

    Cons = "Select ArtNombre from ServicioRenglon, Articulos" & _
                " Where SReServicio = " & xIdServicio & _
                " And SReTipoRenglon = 2 And ArtID = SReMotivo"
    Set rsM = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsM.EOF
        retSTR = retSTR & IIf(retSTR = "", "", "; ") & Trim(rsM!ArtNombre)
        rsM.MoveNext
    Loop
    rsM.Close

    fnc_RepuestosServicio = retSTR
errMotivos:
End Function


Private Function fnc_ProcesoVentas() As Long
On Error GoTo errQ

    fnc_ProcesoVentas = 0
    
    Cons = "Select Cantidad = (Sum(AArCantidadNCo) + Sum(AArCantidadNCr) + Sum(AArCantidadECo) + Sum(AArCantidadECr)) " & _
            " From AcumuladoArticulo " & _
            " Where AArFEcha Between " & Format(tDesde.Text, "'mm/dd/yyyy'") & " And " & Format(tHasta.Text, "'mm/dd/yyyy'")
            
    If Val(tArticulo.Tag) <> 0 Then Cons = Cons & " And AArArticulo = " & Val(tArticulo.Tag)
    If cGrupo.ListIndex <> -1 Then Cons = Cons & " And AArArticulo IN (Select AGrArticulo From ArticuloGrupo Where AGrGrupo = " & Val(cGrupo.ItemData(cGrupo.ListIndex)) & ")"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux("Cantidad")) Then fnc_ProcesoVentas = RsAux("Cantidad")
    End If

    RsAux.Close
    
    Exit Function
errQ:
    clsGeneral.OcurrioError "Error al calcular el acumulado de ventas.", Err.Description
End Function

Private Function zfn_QuitoClaves(mTexto As String) As String
On Error GoTo errQK
    zfn_QuitoClaves = mTexto
    
    If Not ((Mid(mTexto, 1, 1) = "[")) And (InStr(mTexto, "]") <> 0) Then Exit Function
    zfn_QuitoClaves = Mid(mTexto, InStr(mTexto, "]") + 1)
    
errQK:
End Function

Private Sub vsServicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If vsServicio.Rows = vsServicio.FixedRows Then Exit Sub
    vsServicio.Row = vsServicio.MouseRow
    
    If Button = vbRightButton Then
    
        mnuOpcionX.Item(2).Caption = vsServicio.Cell(flexcpText, vsServicio.Row, 1)
        mnuOpcionX.Item(2).Tag = vsServicio.Cell(flexcpText, vsServicio.Row, 1)
        
        mnuOpcionX.Item(3).Caption = "donde dice: " & vsServicio.Cell(flexcpText, vsServicio.Row, 3)
        mnuOpcionX.Item(3).Tag = vsServicio.Cell(flexcpText, vsServicio.Row, 3)
        
        
        mnuOpcionX.Item(7).Caption = "repuesto dice: " & vsServicio.Cell(flexcpText, vsServicio.Row, 4)
        mnuOpcionX.Item(7).Tag = vsServicio.Cell(flexcpText, vsServicio.Row, 4)
               
        
        mnuOpcionX.Item(4).Caption = "Recepcionado: " & vsServicio.Cell(flexcpText, vsServicio.Row, 6)
        mnuOpcionX.Item(4).Tag = vsServicio.Cell(flexcpText, vsServicio.Row, 6)

        mnuOpcionX.Item(5).Caption = "Reparado: " & vsServicio.Cell(flexcpText, vsServicio.Row, 7)
        mnuOpcionX.Item(5).Tag = vsServicio.Cell(flexcpText, vsServicio.Row, 7)
        
        mnuOpcionX.Item(6).Caption = "Estado: " & vsServicio.Cell(flexcpText, vsServicio.Row, 8)
        mnuOpcionX.Item(6).Tag = vsServicio.Cell(flexcpText, vsServicio.Row, 8)
        
        
        PopupMenu mnuBDerecho, , , , mnuOpcionX(0)
       
    End If
    
End Sub
