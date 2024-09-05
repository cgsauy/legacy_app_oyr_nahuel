VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmMovVirtuales 
   Caption         =   "Consulta de Movimientos Virtuales"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9585
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
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4455
      Left            =   50
      TabIndex        =   14
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
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
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
      Left            =   50
      TabIndex        =   29
      Top             =   1200
      Width           =   11415
      _Version        =   196608
      _ExtentX        =   20135
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
      TabIndex        =   30
      Top             =   5880
      Width           =   6135
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmListado.frx":0EF0
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmListado.frx":0FDA
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmListado.frx":10C4
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "frmListado.frx":12FE
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Picture         =   "frmListado.frx":1400
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5400
         Picture         =   "frmListado.frx":17C6
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":18C8
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmListado.frx":1BCA
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmListado.frx":1F0C
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmListado.frx":220E
         Style           =   1  'Graphical
         TabIndex        =   16
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
      TabIndex        =   28
      Top             =   6540
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11271
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
      Height          =   1095
      Left            =   50
      TabIndex        =   27
      Top             =   0
      Width           =   9615
      Begin VB.TextBox tHasta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2520
         MaxLength       =   12
         TabIndex        =   3
         Top             =   240
         Width           =   1095
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
      Begin AACombo99.AACombo cTipo 
         Height          =   315
         Left            =   600
         TabIndex        =   9
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
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
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6360
         MaxLength       =   50
         TabIndex        =   13
         Top             =   600
         Width           =   3015
      End
      Begin AACombo99.AACombo cGrupo 
         Height          =   315
         Left            =   3480
         TabIndex        =   11
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
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
      Begin AACombo99.AACombo cLocal 
         Height          =   315
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
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
      Begin AACombo99.AACombo cEstado 
         Height          =   315
         Left            =   6960
         TabIndex        =   7
         Top             =   240
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
         Text            =   ""
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&Estado:"
         Height          =   255
         Left            =   6360
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Local:"
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   5640
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Grupo:"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Menu MnuAccesos 
      Caption         =   "Accesos"
      Visible         =   0   'False
      Begin VB.Menu MnuAccMenues 
         Caption         =   "Menú Accesos"
      End
      Begin VB.Menu MnuAccLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAccDetalleFactura 
         Caption         =   "Detalle de &Factura"
      End
      Begin VB.Menu MnuAccDetalleOperacion 
         Caption         =   "Detalle de &Operación"
      End
   End
End
Attribute VB_Name = "frmMovVirtuales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Rs1 As rdoResultset
Private aTexto As String
Private bCargarImpresion As Boolean

Private Sub AccionLimpiar()
    cLocal.Text = ""
    cEstado.Text = ""
    cTipo.Text = ""
    cGrupo.Text = ""
    tArticulo.Text = "": tArticulo.Tag = ""
    vsConsulta.Rows = 1
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub
Private Sub bImprimir_Click()
    AccionImprimir True
End Sub
Private Sub bNoFiltros_Click()
    AccionLimpiar
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

Private Sub bConfigurar_Click()
    AccionConfigurar
End Sub

Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub


Private Sub cEstado_GotFocus()
    With cEstado
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda " Seleccione el estado físico del artículo a consultar. [Blanco = Todos]"
End Sub
Private Sub cEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cTipo
End Sub
Private Sub cEstado_LostFocus()
    cEstado.SelStart = 0
    Ayuda ""
End Sub

Private Sub cGrupo_GotFocus()
    With cGrupo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda " Seleccione un grupo de artículo a consultar"
End Sub
Private Sub cGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tArticulo
End Sub
Private Sub cGrupo_LostFocus()
    Ayuda ""
    cGrupo.SelStart = 0
End Sub

Private Sub cLocal_GotFocus()
    With cLocal
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda " Seleccione un local donde se dieron los movimientos. [Blanco = Todos] "
End Sub
Private Sub cLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cEstado
End Sub
Private Sub cLocal_LostFocus()
    cLocal.SelStart = 0
    Ayuda ""
End Sub

Private Sub cTipo_GotFocus()
    With cTipo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda " Seleccione el tipo de artículo a consultar"
End Sub
Private Sub cTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cGrupo
End Sub
Private Sub cTipo_LostFocus()
    cTipo.SelStart = 0
    Ayuda ""
End Sub

Private Sub chVista_Click()
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
        Me.Refresh
    Else
        vsListado.ZOrder 0
        Me.Refresh
        AccionImprimir
    End If
End Sub

Private Sub Label1_Click()
    Foco cTipo
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    InicializoGrillas
    AccionLimpiar
    
    'Cargo los combos--------------------------------------------------
    Cons = "Select TipCodigo, TipNombre From Tipo Order by TipNombre"
    CargoCombo Cons, cTipo
    Cons = "Select GruCodigo, GruNombre From Grupo Order by GruNombre"
    CargoCombo Cons, cGrupo
    Cons = "Select LocCodigo, LocNombre From Local Order By LocNombre"
    CargoCombo Cons, cLocal
    
    'Le inserto los estados virtuales
    cEstado.AddItem RetornoEstadoVirtual(TipoMovimientoEstado.AEntregar)
    cEstado.ItemData(cEstado.NewIndex) = TipoMovimientoEstado.AEntregar
    cEstado.AddItem RetornoEstadoVirtual(TipoMovimientoEstado.ARetirar)
    cEstado.ItemData(cEstado.NewIndex) = TipoMovimientoEstado.ARetirar
    cEstado.AddItem RetornoEstadoVirtual(TipoMovimientoEstado.Reserva)
    cEstado.ItemData(cEstado.NewIndex) = TipoMovimientoEstado.Reserva
    '----------------------------------------------------------------------
    
    FechaDelServidor
    tDesde.Text = Format(gFechaServidor, FormatoFP)
    tHasta.Text = tDesde.Text
    bCargarImpresion = True
    vsListado.Orientation = orPortrait
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub
Private Sub InicializoGrillas()
    On Error Resume Next
    With vsConsulta
        .Redraw = False
        .MultiTotals = False
        .WordWrap = False
        .Cols = 1: .Rows = 1:
        .FormatString = "<Fecha|<Artículo|<Estado|>Q|<Local|<Tipo|"
        .ColWidth(0) = 1750: .ColWidth(1) = 3250: .ColWidth(2) = 1000: .ColWidth(3) = 500: .ColWidth(4) = 1600: .ColWidth(5) = 1400
        .ColWidth(6) = 10
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
            
            Case vbKeyQ: AccionLimpiar
            Case vbKeyI: AccionImprimir True
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            Case vbKeyC: AccionConfigurar
            
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
    Set miconexion = Nothing
    End
    
End Sub

Private Sub AccionConsultar()
Dim Rs As rdoResultset
    On Error GoTo errConsultar
    If Not IsDate(tDesde.Text) Then MsgBox "Debe ingresar una fecha desde.", vbInformation, "ATENCIÓN": Foco tDesde: Exit Sub
    If Not IsDate(tHasta.Text) Then MsgBox "Debe ingresar una fecha hasta.", vbInformation, "ATENCIÓN": Foco tHasta: Exit Sub
    If CDate(tDesde.Text) > CDate(tHasta.Text) Then MsgBox "La fecha Hasta es menor que la fecha Desde.", vbInformation, "ATENCIÓN": Foco tDesde: Exit Sub
    Screen.MousePointer = 11
    bCargarImpresion = True
    vsConsulta.Rows = 1
    vsConsulta.Redraw = False
    CargoMovimientos
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True
End Sub

Private Sub Label2_Click()
    Foco cGrupo
End Sub

Private Sub Label3_Click()
    Foco tArticulo
End Sub

Private Sub Label4_Click()
    Foco tDesde
End Sub

Private Sub Label5_Click()
    Foco tHasta
End Sub

Private Sub Label6_Click()
    Foco cLocal
End Sub

Private Sub Label7_Click()
    Foco cEstado
End Sub

Private Sub MnuAccDetalleFactura_Click()
    EjecutarApp App.Path & "\Detalle de Factura.exe", vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)
End Sub

Private Sub MnuAccDetalleOperacion_Click()
    EjecutarApp App.Path & "\Detalle de Operaciones.exe", vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = "0"
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.Panels(3).Text = "Ingrese el mes y año de liquidación a consultar."
End Sub
Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrAP
    If KeyAscii = vbKeyReturn Then
        If Trim(tArticulo.Text) <> "" Then
            If Val(tArticulo.Tag) <> 0 Then Foco bConsultar: Exit Sub
            Screen.MousePointer = 11
            If IsNumeric(tArticulo.Text) Then
                BuscoArticuloPorCodigo tArticulo.Text
            Else
                BuscoArticuloPorNombre tArticulo.Text
            End If
            If Val(tArticulo.Tag) > 0 Then Foco bConsultar
        Else
            Foco bConsultar
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrAP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub tArticulo_LostFocus()
    Ayuda ""
End Sub

Private Sub tDesde_GotFocus()
    With tDesde
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda " Ingrese la fecha desde a consultar."
End Sub
Private Sub tDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tHasta
End Sub
Private Sub tDesde_LostFocus()
    Ayuda ""
    If IsDate(tDesde.Text) Then tDesde.Text = Format(tDesde.Text, FormatoFP) Else tDesde.Text = ""
End Sub

Private Sub tHasta_GotFocus()
    With tHasta
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda " Ingrese la fecha hasta a consultar."
End Sub
Private Sub tHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cLocal
End Sub
Private Sub tHasta_LostFocus()
    Ayuda ""
    If IsDate(tHasta.Text) Then tHasta.Text = Format(tHasta.Text, FormatoFP) Else tHasta.Text = ""
End Sub

Private Sub vsConsulta_DblClick()
    If vsConsulta.Rows = 1 Then Exit Sub
    DeMovimientoStock.pTipoMovimiento = TipoEstadoMercaderia.Virtual
    DeMovimientoStock.pLista = vsConsulta
    DeMovimientoStock.Show vbModal, Me
End Sub

Private Sub vsConsulta_GotFocus()
    Ayuda "Doble Click = Detalle de Movimientos; Botón derecho acceso a Menúes."
End Sub

Private Sub vsConsulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        With vsConsulta
            If .Cell(flexcpData, .Row, 2) <> 0 Then
                MnuAccDetalleOperacion.Enabled = False
                MnuAccDetalleFactura.Enabled = False
                If .Cell(flexcpData, .Row, 2) = TipoDocumento.Credito Or .Cell(flexcpData, .Row, 2) = TipoDocumento.Contado Or .Cell(flexcpData, .Row, 2) = TipoDocumento.NotaCredito _
                    Or .Cell(flexcpData, .Row, 2) = TipoDocumento.NotaDevolucion Or .Cell(flexcpData, .Row, 2) = TipoDocumento.Remito Then
                    MnuAccDetalleOperacion.Enabled = True
                    MnuAccDetalleFactura.Enabled = True
                    PopupMenu MnuAccesos
                End If
            End If
        End With
    End If
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    If bCargarImpresion Then
        If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub
        With vsListado
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        EncabezadoListado vsListado, "Consulta de Movimientos Físicos del " & Format(tDesde.Text, FormatoFP) & " al " & Format(tHasta.Text, FormatoFP), False
        vsListado.FileName = "Consulta de Movimientos Físicos"
        
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        vsListado.EndDoc
        bCargarImpresion = False
    End If
    
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
Private Sub CargoMovimientos()
Dim Rs As rdoResultset
On Error GoTo ErrCS
Dim aCod As Long
    
    Screen.MousePointer = 11
    Cons = "Select * From MovimientoStockEstado "
    If cLocal.ListIndex > -1 Then
        Cons = Cons & ", Local "
    Else
        Cons = Cons & "Left Outer Join Local ON LocCodigo = MSELocal "
    End If
    Cons = Cons & ", Articulo " _
        & " Where MSEFecha >= '" & Format(tDesde.Text & " 00:00:00", sqlFormatoFH) & "'" _
        & " And MSEFecha <= '" & Format(tHasta.Text & " 23:59:59", sqlFormatoFH) & "'" _
    
    'Filtro Estado
    If cEstado.ListIndex > -1 Then Cons = Cons & " And MSEEstado = " & cEstado.ItemData(cEstado.ListIndex)
    
    'Uno Artículo
    If Val(tArticulo.Tag) > 0 Then Cons = Cons & " And MSEArticulo = " & tArticulo.Tag
    Cons = Cons & " And MSEArticulo = ArtID"
    
    'Filtro de Tipo de Artículo
    If cTipo.ListIndex > -1 Then Cons = Cons & " And ArtTipo = " & cTipo.ItemData(cTipo.ListIndex)
    
    'Si hay local lo uno.
    If cLocal.ListIndex > -1 Then
        Cons = Cons & " And MSELocal = " & cLocal.ItemData(cLocal.ListIndex) _
            & " And MSELocal = LocCodigo"
    End If
    
    If cGrupo.ListIndex > -1 Then
        Cons = Cons & " And ArtID IN (" _
                & " Select AGrArticulo from ArticuloGrupo" _
                & " Where AGrGrupo = " & cGrupo.ItemData(cGrupo.ListIndex) & ")"
    End If

    
    Cons = Cons & " Order By MSEFecha "
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCIÓN"
    Else
        Do While Not RsAux.EOF
            With vsConsulta
                .AddItem ""
                aCod = RsAux!MSECodigo
                .Cell(flexcpData, .Rows - 1, 0) = aCod
                If Not IsNull(RsAux!MSEDocumento) Then aCod = RsAux!MSEDocumento Else aCod = 0
                .Cell(flexcpData, .Rows - 1, 1) = aCod
                If Not IsNull(RsAux!MSETipoDocumento) Then aCod = RsAux!MSETipoDocumento Else aCod = 0
                .Cell(flexcpData, .Rows - 1, 2) = aCod
                .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!MSEFecha, FormatoFHP)
                .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "#,###,###") & " " & Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 2) = RetornoEstadoVirtual(RsAux!MSEEstado)
                .Cell(flexcpText, .Rows - 1, 3) = RsAux!MSECantidad
                If Not IsNull(RsAux!LocNombre) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!LocNombre)
                If Not IsNull(RsAux!MSETipoDocumento) Then .Cell(flexcpText, .Rows - 1, 5) = RetornoNombreDocumento(RsAux!MSETipoDocumento)
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrCS:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los movimientos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub BuscoArticuloPorCodigo(CodArticulo As Long)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
    
    Screen.MousePointer = 11
    Cons = "Select * From Articulo Where ArtCodigo = " & CodArticulo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If RsAux.EOF Then
        RsAux.Close
        tArticulo.Tag = "0"
        MsgBox "No existe un artículo que posea ese código.", vbExclamation, "ATENCIÓN"
    Else
        tArticulo.Text = Format(RsAux!ArtCodigo, "#,000,000") & " " & Trim(RsAux!ArtNombre)
        tArticulo.Tag = RsAux!ArtID
        RsAux.Close
    End If
    Screen.MousePointer = 0

End Sub

Private Sub BuscoArticuloPorNombre(NomArticulo As String)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
Dim Resultado As Long

    Screen.MousePointer = 11
    Cons = "Select Código = ArtCodigo, Nombre = ArtNombre from Articulo" _
        & " Where ArtNombre LIKE '" & Replace(NomArticulo, " ", "%") & "%'" _
        & " Order By ArtNombre"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un nombre de artículo con esas características.", vbInformation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            Resultado = RsAux(0)
            RsAux.Close
        Else
            RsAux.Close
            Dim LiAyuda As New clsListadeAyuda
            If LiAyuda.ActivarAyuda(cBase, Cons, Titulo:="Buscar Artículo") > 0 Then
                Resultado = LiAyuda.RetornoDatoSeleccionado(0)
            Else
                Resultado = 0
            End If
            Set LiAyuda = Nothing       'Destruyo la clase.
        End If
        If Resultado > 0 Then BuscoArticuloPorCodigo Resultado
    End If
    Screen.MousePointer = 0
    
End Sub


