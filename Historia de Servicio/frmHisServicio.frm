VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.Form frmHisServicio 
   Caption         =   "Historia de Servicios"
   ClientHeight    =   7380
   ClientLeft      =   4200
   ClientTop       =   2775
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHisServicio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   9120
   Begin VSFlex6DAOCtl.vsFlexGrid vsDetalle 
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2355
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
   Begin VB.Frame fFicha 
      ForeColor       =   &H00000080&
      Height          =   1575
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   7755
      Begin VB.CommandButton bLista 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5820
         TabIndex        =   2
         Top             =   170
         Width           =   375
      End
      Begin VB.TextBox tPFCompra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         MaxLength       =   10
         TabIndex        =   4
         Top             =   540
         Width           =   975
      End
      Begin VB.TextBox tPFacturaN 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3540
         MaxLength       =   6
         TabIndex        =   7
         Top             =   540
         Width           =   675
      End
      Begin VB.TextBox tPFacturaS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "AA"
         Top             =   540
         Width           =   315
      End
      Begin VB.TextBox tPNroMaquina 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5700
         MaxLength       =   15
         TabIndex        =   9
         Top             =   540
         Width           =   1875
      End
      Begin VB.TextBox tPDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox tProducto 
         Height          =   285
         Left            =   960
         MaxLength       =   8
         TabIndex        =   1
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
         Height          =   195
         Left            =   6300
         TabIndex        =   33
         Top             =   180
         Width           =   615
      End
      Begin VB.Label lPEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6900
         TabIndex        =   32
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Garantía:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lPGarantia 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   960
         TabIndex        =   30
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label Label13 
         Caption         =   "F/&Compra:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº Factura: "
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   540
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Seri&e:"
         Height          =   255
         Left            =   4980
         TabIndex        =   8
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "&Producto:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   180
         Width           =   795
      End
      Begin VB.Label LabCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "WALTER ADRIAN OCCHIUZZI MARTINEZ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   26
         Top             =   1200
         Width           =   6615
      End
      Begin VB.Label LabProducto 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(190111) REFRIGERADOR PANAVOX ALTO DE 10 PULGAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1680
         TabIndex        =   25
         Top             =   180
         Width           =   4095
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   1755
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3096
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
      Height          =   3795
      Left            =   120
      TabIndex        =   22
      Top             =   1680
      Width           =   7335
      _Version        =   196608
      _ExtentX        =   12938
      _ExtentY        =   6694
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
      Top             =   5280
      Width           =   6135
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   3120
         Picture         =   "frmHisServicio.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmHisServicio.frx":0D44
         Height          =   310
         Left            =   3480
         Picture         =   "frmHisServicio.frx":0E46
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2040
         Picture         =   "frmHisServicio.frx":1378
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   1680
         Picture         =   "frmHisServicio.frx":1462
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1140
         Picture         =   "frmHisServicio.frx":154C
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   2760
         Picture         =   "frmHisServicio.frx":1786
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmHisServicio.frx":1888
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   420
         Picture         =   "frmHisServicio.frx":198A
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   780
         Picture         =   "frmHisServicio.frx":1CCC
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   60
         Picture         =   "frmHisServicio.frx":1FCE
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
   End
   Begin VB.Label LabDetalle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6300
      TabIndex        =   29
      Top             =   5400
      Width           =   1275
   End
End
Attribute VB_Name = "frmHisServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private aTexto As String, idProducto As Long

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConfigurar_Click()
    AccionConfigurar
End Sub
Private Sub bImprimir_Click()
    AccionImprimir True
End Sub

Private Sub bLista_Click()
Dim aValor As Long
    If idProducto = 0 Or Val(LabCliente.Tag) = 0 Then Exit Sub
    On Error GoTo errLista
    Screen.MousePointer = 11
    Dim objLista As New clsListadeAyuda
    
    Cons = "Select ProCodigo, ProCodigo 'Codigo', ArtNombre 'Producto', ProCompra 'F/Compra', ProNroSerie 'Nº Serie' from Producto, Articulo" _
           & " Where ProCliente = " & Val(LabCliente.Tag) _
           & " And ProArticulo = ArtID"
    
    If objLista.ActivarAyuda(cBase, Cons, 7000, 0, "Productos") > 0 Then
        aValor = objLista.RetornoDatoSeleccionado(0)
    End If
    Set objLista = Nothing
    Me.Refresh
    If aValor <> 0 Then idProducto = aValor: AccionConsultar
    Screen.MousePointer = 0
    Exit Sub

errLista:
    clsGeneral.OcurrioError "Ocurrió un error al procesar la lista de prodcutos.", Err.Description
    Screen.MousePointer = 0
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
    If vsListado.Visible Then Zoom vsListado, vsListado.Zoom + 5
End Sub
Private Sub bZMenos_Click()
    If vsListado.Visible Then Zoom vsListado, vsListado.Zoom - 5
End Sub

Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub

Private Sub chVista_Click()
    If chVista.Value = 0 Then
        'vsConsulta.ZOrder 0
        vsListado.Visible = False
        Me.Refresh
    Else
        AccionImprimir
        vsListado.Visible = True
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

    FechaDelServidor
    ObtengoSeteoForm Me
    picBotones.BorderStyle = vbBSNone
    'Hoja Carta
    vsListado.Orientation = orPortrait: vsListado.PaperSize = 1
    vsListado.Visible = False
    
    LabCliente.Caption = "": LimpioDatosProducto
    InicializoGrillas
    If Trim(Command()) <> "" Then idProducto = CLng(Command()) Else idProducto = 0
    AccionConsultar
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            
            Case vbKeyE: AccionConsultar
            
            Case vbKeyP: IrAPagina vsListado, 1
            Case vbKeyA: IrAPagina vsListado, vsListado.PreviewPage - 1
            Case vbKeyS: IrAPagina vsListado, vsListado.PreviewPage + 1
            Case vbKeyU: IrAPagina vsListado, vsListado.PageCount
            
            Case vbKeyAdd: If vsListado.Visible Then Zoom vsListado, vsListado.Zoom + 5
            Case vbKeySubtract: If vsListado.Visible Then Zoom vsListado, vsListado.Zoom - 5
            
            Case vbKeyI: AccionImprimir True
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            Case vbKeyC: AccionConfigurar
            
            Case vbKeyQ: AccionConfigurar
            
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11
    
    vsListado.Height = Me.ScaleHeight - (vsListado.Top + picBotones.Height + 30)
    picBotones.Top = vsListado.Height + vsListado.Top + 30
    
    LabDetalle.Top = ((picBotones.Top - fFicha.Top + fFicha.Height) - LabDetalle.Height) / 2
    
    fFicha.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = fFicha.Width
    vsListado.Left = fFicha.Left: LabDetalle.Left = fFicha.Left: LabDetalle.Width = fFicha.Width
    
    
    vsConsulta.Height = LabDetalle.Top - (vsListado.Top + 15) 'vsConsulta.Height = vsListado.Height
    vsConsulta.Width = vsListado.Width: vsDetalle.Width = vsListado.Width
    vsConsulta.Top = vsListado.Top
    vsConsulta.ColWidth(1) = vsConsulta.Width - (vsConsulta.ColWidth(0) + vsConsulta.ColWidth(2) + vsConsulta.ColWidth(3) + vsConsulta.ColWidth(4) + 300)
    
    vsConsulta.Left = vsListado.Left: vsDetalle.Left = vsConsulta.Left
    vsDetalle.Top = LabDetalle.Top + LabDetalle.Height + 10
    vsDetalle.Height = vsListado.Top + vsListado.Height - (vsDetalle.Top)
    
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
Private Sub InicializoGrillas()
    On Error Resume Next
    With vsConsulta
        .Editable = False
        .Redraw = False
        .WordWrap = True
        .Rows = 1: .Cols = 1
        .FormatString = "Fecha|Motivos|Estado|>Importe|"
        .ColWidth(0) = 750: .ColWidth(1) = 4800: .ColWidth(4) = 10
        .ColWidth(3) = 1050
        .ColAlignment(0) = flexAlignLeftTop
        .ColAlignment(1) = flexAlignLeftTop
        .ColAlignment(2) = flexAlignLeftTop
        .ColAlignment(3) = flexAlignRightTop
        .ColAlignment(4) = flexAlignLeftTop
        .Redraw = True
    End With
    With vsDetalle
        .Redraw = False
        .WordWrap = True
        .Rows = 1: .Cols = 1
        .FormatString = "Tipo|Fecha|Hecho por|>Importe|Comentario|"
        .ColWidth(0) = 900: .ColWidth(1) = 800: .ColWidth(2) = 900: .ColWidth(3) = 800: .ColWidth(4) = 3000
        .ColAlignment(0) = flexAlignLeftTop
        .ColAlignment(1) = flexAlignLeftTop
        .ColAlignment(2) = flexAlignLeftTop
        .ColAlignment(3) = flexAlignLeftTop
        .ColAlignment(5) = flexAlignLeftTop
        .ColDataType(1) = flexDTDate
        .ColHidden(5) = True
        .Redraw = True
    End With
End Sub

Private Sub AccionConsultar()
    On Error GoTo errConsultar
    Screen.MousePointer = 11
    'Saco los datos del cliente y del producto.
    Cons = "Select * From Producto, Articulo, Cliente " _
            & " Left Outer Join CPersona ON CliCodigo = CPeCliente" _
            & " Left Outer Join CEmpresa ON CliCodigo = CEmCliente" _
        & " Where ProCodigo = " & idProducto _
        & " And ProArticulo = ArtID And ProCliente = CliCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsAux.EOF Then RsAux.Close: Screen.MousePointer = 0: Exit Sub
    
    tProducto.Text = Format(RsAux!ProCodigo, "#,000")
    idProducto = RsAux!ProCodigo
    bLista.Enabled = True
    LabCliente.Tag = RsAux!CliCodigo
    If RsAux!CliTipo = TipoCliente.Cliente Then
        If Not IsNull(RsAux!CliCIRUC) Then LabCliente.Caption = " (" & clsGeneral.RetornoFormatoCedula(RsAux!CliCIRUC) & ")"
        LabCliente.Caption = LabCliente.Caption & " " & Trim(Trim(Format(RsAux!CPeNombre1, "#")) & " " & Trim(Format(RsAux!CPeNombre2, "#"))) & ", " & Trim(Trim(Format(RsAux!CPeApellido1, "#")) & " " & Trim(Format(RsAux!CPeApellido2, "#")))
    Else
        If Not IsNull(RsAux!CliCIRUC) Then LabCliente.Caption = " (" & clsGeneral.RetornoFormatoRuc(RsAux!CliCIRUC) & ")"
        If Not IsNull(RsAux!CEmNombre) Then LabCliente.Caption = LabCliente.Caption & " " & Trim(RsAux!CEmFantasia)
        If Not IsNull(RsAux!CEmFantasia) Then LabCliente.Caption = LabCliente.Caption & " (" & Trim(RsAux!CEmFantasia) & ")"
    End If
    
    LabProducto.Caption = Trim(RsAux!ArtNombre)
    If Not IsNull(RsAux!ProCompra) Then tPFCompra.Text = Format(RsAux!ProCompra, "dd/mm/yyyy")
    If Not IsNull(RsAux!ProFacturaS) Then tPFacturaS.Text = Trim(RsAux!ProFacturaS)
    If Not IsNull(RsAux!ProFacturaN) Then tPFacturaN.Text = RsAux!ProFacturaN
    If Not IsNull(RsAux!ProNroSerie) Then tPNroMaquina.Text = Trim(RsAux!ProNroSerie)
    If Not IsNull(RsAux!ProDireccion) Then
        tPDireccion.Text = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!ProDireccion, True, True, True)
        tPDireccion.Tag = RsAux!ProDireccion
    End If
    
    If Not IsNull(RsAux!ProDocumento) Then
        tPFCompra.Enabled = False: tPFacturaS.Enabled = False: tPFacturaN.Enabled = False
        tPFCompra.BackColor = Inactivo: tPFacturaS.BackColor = Inactivo: tPFacturaN.BackColor = Inactivo
    Else
        tPFCompra.Enabled = True: tPFacturaS.Enabled = True: tPFacturaN.Enabled = True
        tPFCompra.BackColor = Blanco: tPFacturaS.BackColor = Blanco: tPFacturaN.BackColor = Blanco
    End If
    lPGarantia.Caption = " " & RetornoGarantia(RsAux!ArtId)
    lPEstado.Tag = CalculoEstadoProducto(RsAux!ProCodigo)
    lPEstado.Caption = " " & EstadoProducto(Val(lPEstado.Tag))
    
    tPFCompra.Tag = 0: tPFacturaS.Tag = 0: tPFacturaN.Tag = 0: tPNroMaquina.Tag = 0
    RsAux.Close
    
    chVista.Value = 0
    With vsConsulta
        .Rows = 1
        .Refresh
        .Redraw = False
        CargoHistoria
        .Redraw = True
    End With
    If vsConsulta.Rows > 1 Then
        With vsDetalle
            .Rows = 1
            .Refresh
            .Redraw = False
            CargoDetalle CLng(vsConsulta.Cell(flexcpData, 1, 0))
            .Redraw = True
        End With
    End If
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True
    vsDetalle.Redraw = True
End Sub

Private Sub Label13_Click()
    Foco tPFCompra
End Sub

Private Sub Label2_Click()
    Foco tProducto
End Sub

Private Sub Label3_Click()
    Foco tPNroMaquina
End Sub

Private Sub Label4_Click()
    Foco tPFacturaS
End Sub

Private Sub tPFacturaN_Change()
    tPFacturaN.Tag = 1
End Sub

Private Sub tPFacturaN_GotFocus()
    With tPFacturaN: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tPFacturaN_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tPFacturaN.Tag) <> 0 And idProducto <> 0 Then
            ZActualizoCampoProducto idProducto, Trim(tPFacturaN.Text), FacturaN:=True
            tPFacturaN.Tag = 0
        End If
        Foco tPNroMaquina
    End If
End Sub

Private Sub tPFacturaS_Change()
    tPFacturaS.Tag = 1
End Sub

Private Sub tPFacturaS_GotFocus()
    With tPFacturaS: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tPFacturaS_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        If Val(tPFacturaS.Tag) <> 0 And idProducto <> 0 Then
            ZActualizoCampoProducto idProducto, Trim(tPFacturaS.Text), FacturaS:=True
            tPFacturaS.Tag = 0
        End If
        Foco tPFacturaN
    End If
End Sub

Private Sub tPFCompra_Change()
    tPFCompra.Tag = 1
End Sub

Private Sub tPFCompra_GotFocus()
    With tPFCompra: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tPFCompra_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Val(tPFCompra.Tag) <> 0 And idProducto <> 0 Then
            If Not IsDate(tPFCompra.Text) Then MsgBox "La fecha ingresada no es correcta. Verifique", vbExclamation, "ATENCIÓN": Exit Sub
            ZActualizoCampoProducto idProducto, Trim(tPFCompra.Text), FCompra:=True
            tPFCompra.Text = Format(tPFCompra.Text, "dd/mm/yyyy")
            tPFCompra.Tag = 0
        End If
        Foco tPFacturaS
    End If
    
End Sub

Private Sub tPNroMaquina_Change()
    tPNroMaquina.Tag = 1
End Sub

Private Sub tPNroMaquina_GotFocus()
    With tPNroMaquina: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tPNroMaquina_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tPNroMaquina.Tag) <> 0 And idProducto <> 0 Then
            ZActualizoCampoProducto idProducto, Trim(tPNroMaquina.Text), NroMaquina:=True
            tPNroMaquina.Tag = 0
        End If
        vsConsulta.SetFocus
    End If
End Sub

Private Sub tProducto_Change()
    idProducto = 0
    LimpioDatosProducto
    LabCliente.Caption = ""
    vsConsulta.Rows = 1
    vsDetalle.Rows = 1
End Sub

Private Sub tProducto_GotFocus()
    With tProducto
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If IsNumeric(tProducto.Text) Then
        idProducto = tProducto.Text
        AccionConsultar
    End If
End Sub

Private Sub vsConsulta_Click()
    CargoDetalle Val(vsConsulta.Cell(flexcpData, vsConsulta.Row, 0))
End Sub

Private Sub vsConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsConsulta.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            CargoDetalle Val(vsConsulta.Cell(flexcpData, vsConsulta.Row, 0))
        Case vbKeyDown, vbKeyUp: vsDetalle.Rows = 1
    End Select
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
Dim Consulta As Boolean
    On Error GoTo errImprimir
    
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
    
    
    EncabezadoListado vsListado, "Historia de Servicios", False
    vsListado.FileName = "Historia de Servicios"
    
    vsListado.Paragraph = "": vsListado.Paragraph = "Cliente: " & Trim(LabCliente.Caption)
    vsListado.Paragraph = "Producto: " & Trim(LabProducto.Caption): vsListado.Paragraph = ""
    vsListado.Paragraph = "Garantía: " & Trim(lPGarantia.Caption) & Chr(vbKeyTab) & Chr(vbKeyTab) & "Estado: " & Trim(lPEstado.Caption)
    vsListado.Paragraph = "F/Compra: " & Trim(tPFCompra.Text) & Chr(vbKeyTab) & Chr(vbKeyTab) & "Factura: " & Trim(tPFacturaS.Text) & " " & Trim(tPFacturaN.Text) & Chr(vbKeyTab) & Chr(vbKeyTab) & "Nº Serie: " & Trim(tPNroMaquina.Text)
    vsListado.Paragraph = "Dirección: " & Trim(tPDireccion.Text)
     
    vsListado.Paragraph = ""
     
    vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
    vsListado.Paragraph = "": vsListado.Paragraph = "Servicio Realizado el: " & Trim(vsConsulta.Cell(flexcpText, vsConsulta.Row, 0))
    
    Dim aValor As Long
    vsDetalle.ExtendLastCol = False
    aValor = vsDetalle.ColWidth(4): vsDetalle.ColWidth(4) = 6200
    vsListado.RenderControl = vsDetalle.hwnd: vsDetalle.ExtendLastCol = True
    vsDetalle.ColWidth(4) = aValor
    
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

Private Sub CargoHistoria()
Dim IdServicio As Long, aComentario As String
    
    Cons = "Select * From Servicio" _
            & " Left Outer Join ServicioRenglon ON SReTipoRenglon = " & TipoRenglonS.Cumplido _
                                                                    & " And SerCodigo = SReServicio" _
            & " Left Outer Join Articulo ON SReMotivo = ArtID " _
        & " Where SerProducto = " & idProducto _
        & " And SerEstadoServicio In (" & EstadoS.Cumplido & ", " & EstadoS.Anulado & ")" _
        & " Order By SerCodigo DESC"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsAux.EOF
        
        With vsConsulta
            .AddItem ""
            IdServicio = RsAux!SerCodigo
            If RsAux!SerEstadoServicio = EstadoS.Anulado Then
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo ': .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
            End If
            
            .Cell(flexcpData, .Rows - 1, 0) = IdServicio
            
            .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!SerFCumplido, "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, 2) = EstadoProducto(RsAux!SerEstadoProducto)
            If Not IsNull(RsAux!SerMoneda) And Not IsNull(RsAux!SerCostoFinal) Then
                .Cell(flexcpText, .Rows - 1, 3) = BuscoSignoMoneda(RsAux!SerMoneda) & " " & Format(RsAux!SerCostoFinal, FormatoMonedaP)
            End If
            If Not IsNull(RsAux!SerComentarioR) Then aComentario = Trim(RsAux!SerComentarioR) Else aComentario = ""
            Do While Not RsAux.EOF
                If Trim(.Cell(flexcpText, .Rows - 1, 1)) <> "" Then .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1) & ", "
                .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1) & Trim(RsAux!ArtNombre)
                IdServicio = RsAux!SerCodigo
                RsAux.MoveNext
                If RsAux.EOF Then Exit Do
                If IdServicio <> RsAux!SerCodigo Then Exit Do
            Loop
            If Trim(.Cell(flexcpText, .Rows - 1, 1)) <> "" Then .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1) & Chr(13)
            .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1) & aComentario
            
        End With
        'RsAux.MoveNext
    Loop
    RsAux.Close
    With vsConsulta
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 1, , False
        If .Rows = 2 Then .HighLight = flexHighlightNever
    End With
End Sub
Private Sub CargoDetalle(IdServicio As Long)
Dim RsSR As rdoResultset
Dim aTexto As String
    vsDetalle.Rows = 1
    Cons = "Select * From Servicio " _
            & " Left Outer Join ServicioVisita On VisServicio = SerCodigo " _
            & ", Usuario Where SerCodigo = " & IdServicio & " And SerUsuario = UsuCodigo "
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then
        With vsDetalle
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = "Solicitud"
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!SerFecha, "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!UsuIdentificacion)
            Cons = "Select * From ServicioRenglon, MotivoServicio " _
                & " Where SReServicio = " & RsAux!SerCodigo _
                & " And SReTipoRenglon = " & TipoRenglonS.Llamado _
                & " And SReMotivo = MSeID"
            Set RsSR = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
            Do While Not RsSR.EOF
                If Trim(.Cell(flexcpText, .Rows - 1, 4)) <> "" Then .Cell(flexcpText, .Rows - 1, 4) = Trim(.Cell(flexcpText, .Rows - 1, 4)) & ", "
                .Cell(flexcpText, .Rows - 1, 4) = .Cell(flexcpText, .Rows - 1, 4) & Trim(RsSR!MSeNombre)
                RsSR.MoveNext
            Loop
            RsSR.Close
            If Not IsNull(RsAux!SerComentario) Then
                If Trim(.Cell(flexcpText, .Rows - 1, 4)) <> "" Then .Cell(flexcpText, .Rows - 1, 4) = Trim(.Cell(flexcpText, .Rows - 1, 4)) & Chr(13)
                .Cell(flexcpText, .Rows - 1, 4) = Trim(.Cell(flexcpText, .Rows - 1, 4)) & Trim(RsAux!SerComentario)
            End If
            .Cell(flexcpText, .Rows - 1, 5) = 0
        End With
        If Not IsNull(RsAux!VisServicio) Then
            Do While Not RsAux.EOF
                With vsDetalle
                    
                    .AddItem ""
                    'Son Visitas, Retiros y Entregas.
                    Select Case RsAux!VisTipo
                        Case TipoServicio.Entrega: .Cell(flexcpText, .Rows - 1, 0) = "Entrega": .Cell(flexcpText, .Rows - 1, 5) = 50
                        Case TipoServicio.Retiro: .Cell(flexcpText, .Rows - 1, 0) = "Retiro": .Cell(flexcpText, .Rows - 1, 5) = 10
                        Case TipoServicio.Visita: .Cell(flexcpText, .Rows - 1, 0) = "Visita": .Cell(flexcpText, .Rows - 1, 5) = 20
                    End Select
                    If RsAux!VisSinEfecto Then .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Inactivo ': .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Blanco
                    .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!VisFecha, "dd/mm/yy")
                    .Cell(flexcpText, .Rows - 1, 2) = BuscoNombreLocal(RsAux!VisCamion)
                    If Not IsNull(RsAux!VisCosto) Then
                        aTexto = BuscoSignoMoneda(RsAux!VisMoneda)
                        .Cell(flexcpText, .Rows - 1, 3) = Trim(aTexto) & " " & Format(RsAux!VisCosto, FormatoMonedaP)
                    End If
                    If Not IsNull(RsAux!VisComentario) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!VisComentario)
                    If Not IsNull(RsAux!VisTexto) Then
                        If Trim(.Cell(flexcpText, .Rows - 1, 4)) = "" Then
                            .Cell(flexcpText, .Rows - 1, 4) = RetornoTextoVisita(RsAux!VisTexto)
                        Else
                            .Cell(flexcpText, .Rows - 1, 4) = Trim(.Cell(flexcpText, .Rows - 1, 4)) & Chr(13) & RetornoTextoVisita(RsAux!VisTexto)
                        End If
                    End If
                End With
                RsAux.MoveNext
            Loop
        End If
        RsAux.Close
    Else
        RsAux.Close: Exit Sub
    End If
    Cons = "Select * From Servicio, Taller Where SerCodigo = " & IdServicio & " And SerCodigo = TalServicio"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsAux.EOF
        With vsDetalle
            'Son datos del taller y de traslados.
            If Not IsNull(RsAux!TalIngresoCamion) Then
                'Tengo un traslado de ingreso.
                '------------------------------------------------------
                'Ingreso 1ro el traslado.
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = "Traslado"
                .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!TalFIngresoRealizado, "dd/mm/yy")
                .Cell(flexcpText, .Rows - 1, 2) = BuscoNombreLocal(RsAux!TalIngresoCamion)
                'Pongo de donde viene.
                .Cell(flexcpText, .Rows - 1, 4) = "De " & BuscoNombreLocal(RsAux!SerLocalIngreso) & " a " & BuscoNombreLocal(RsAux!SerLocalReparacion)
                '------------------------------------------------------
                .Cell(flexcpText, .Rows - 1, 5) = 41
            End If
            '2do Ingreso los datos de taller.
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = "Taller"
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!TalFIngresoRecepcion, "dd/mm/yy")
            If Not IsNull(RsAux!TalTecnico) Then .Cell(flexcpText, .Rows - 1, 2) = BuscoUsuario(RsAux!TalTecnico, True)
            If Not IsNull(RsAux!TalComentario) Then .Cell(flexcpText, .Rows - 1, 4) = modStart.f_QuitarClavesDelComentario(Trim(RsAux!TalComentario))
            If Not IsNull(RsAux!TalFPresupuesto) Then aTexto = "Presupuestado el " & Format(RsAux!TalFPresupuesto, "dd/mm/yy")
            If Not IsNull(RsAux!TalFAceptacion) Then
                If RsAux!TalAceptado Then
                    aTexto = aTexto & ", Aceptado el " & Format(RsAux!TalFAceptacion, "dd/mm/yy")
                Else
                    aTexto = aTexto & ", NO Aceptado el " & Format(RsAux!TalFAceptacion, "dd/mm/yy")
                End If
            End If
            If Not IsNull(RsAux!TalFReparado) Then aTexto = aTexto & ", Reparado el " & Format(RsAux!TalFReparado, "dd/mm/yy")
            If RsAux!TalSinArreglo Then aTexto = aTexto & ", SIN ARREGLO"
            If Not IsNull(RsAux!TalComentario) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(.Cell(flexcpText, .Rows - 1, 4)) & Chr(13) & Chr(10) & aTexto Else .Cell(flexcpText, .Rows - 1, 4) = aTexto
            .Cell(flexcpText, .Rows - 1, 5) = 42
            '------------------------------------------------------
            '3ro.Verifico si hay traslado para local de entrega.
            If Not IsNull(RsAux!TalSalidaCamion) Then
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = "Traslado"
                .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!TalFSalidaRealizado, "dd/mm/yy")
                .Cell(flexcpText, .Rows - 1, 2) = BuscoNombreLocal(RsAux!TalSalidaCamion)
                .Cell(flexcpText, .Rows - 1, 4) = "De " & BuscoNombreLocal(RsAux!SerLocalReparacion) & " a " & BuscoNombreLocal(RsAux!TalLocalAlCliente)
                If Not IsNull(RsAux!TalFSalidaRecepcion) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(.Cell(flexcpText, .Rows - 1, 4)) & ", Arribó el " & Format(RsAux!TalFSalidaRecepcion, "dd/mm/yy")
                .Cell(flexcpText, .Rows - 1, 5) = 43
            End If
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    With vsDetalle
        If .Rows = 1 Then Exit Sub
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 4, , False
        .Select 1, 5, .Rows - 1, 5
        .Sort = flexSortGenericAscending
        .Select 1, 0, 1, 1
    End With
    
End Sub

Private Function BuscoNombreLocal(IdLocal As Long) As String
Dim RsLocal As rdoResultset
    BuscoNombreLocal = ""
    Cons = "Select * From Local Where LocCodigo = " & IdLocal
    Set RsLocal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsLocal.EOF Then BuscoNombreLocal = Trim(RsLocal!LocNombre)
    RsLocal.Close
End Function
Private Function BuscoSignoMoneda(IdMoneda As Long)
Dim RsMon As rdoResultset
    BuscoSignoMoneda = ""
    Cons = "Select * From Moneda Where MonCodigo = " & IdMoneda
    Set RsMon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsMon.EOF Then BuscoSignoMoneda = Trim(RsMon!MonSigno)
    RsMon.Close
End Function
Private Sub ZActualizoCampoProducto(idProducto As Long, Valor As Variant, _
                                Optional FCompra As Boolean = False, Optional FacturaS As Boolean = False, Optional FacturaN As Boolean = False, _
                                Optional NroMaquina As Boolean = False)
    
    On Error GoTo errActualizar
    If idProducto = 0 Then Exit Sub
    Screen.MousePointer = 11
    
    Cons = "Select * from Producto Where ProCodigo = " & idProducto
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Edit
        
        If FCompra Then If Trim(Valor) = "" Then RsAux!ProCompra = Null Else RsAux!ProCompra = Format(Valor, sqlFormatoF)
        If FacturaS Then If Trim(Valor) = "" Then RsAux!ProFacturaS = Null Else RsAux!ProFacturaS = Trim(Valor)
        If FacturaN Then If Trim(Valor) = "" Then RsAux!ProFacturaN = Null Else RsAux!ProFacturaN = CLng(Valor)
        If NroMaquina Then If Trim(Valor) = "" Then RsAux!ProNroSerie = Null Else RsAux!ProNroSerie = Trim(Valor)
        RsAux.Update
    End If
    RsAux.Close
    
    If FCompra Then
        lPEstado.Tag = CalculoEstadoProducto(idProducto)
        lPEstado.Caption = " " & EstadoProducto(Val(lPEstado.Tag))
    End If
    Screen.MousePointer = 0
    
    Exit Sub
errActualizar:
    clsGeneral.OcurrioError "Ocurrió un error al actualizar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub LimpioDatosProducto()
    
    LabProducto.Caption = ""
    lPEstado.Caption = "": tPFCompra.Text = "": tPFacturaS.Text = "": tPFacturaN.Text = "": tPNroMaquina.Text = ""
    lPGarantia.Caption = "": tPDireccion.Text = ""
    tPFCompra.Tag = 0: tPFacturaS.Tag = 0: tPFacturaN.Tag = 0: tPNroMaquina.Tag = 0
    
    bLista.Enabled = False
    LabCliente.Tag = 0
    idProducto = 0
    
End Sub

Private Function RetornoTextoVisita(IdTexto As Integer) As String
Dim RsTV As rdoResultset
    RetornoTextoVisita = ""
    Cons = "Select * From TextoVisita Where TViCodigo = " & IdTexto
    Set RsTV = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsTV.EOF Then RetornoTextoVisita = Trim(RsTV!TViTexto)
    RsTV.Close
End Function
