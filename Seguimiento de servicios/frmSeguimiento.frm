VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Begin VB.Form frmSeguimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento de Servicios"
   ClientHeight    =   7410
   ClientLeft      =   2340
   ClientTop       =   3030
   ClientWidth     =   10890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSeguimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   10890
   Begin VSPrinter8LibCtl.VSPrinter vsFicha 
      Height          =   2175
      Left            =   9120
      TabIndex        =   74
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
      _cx             =   2566
      _cy             =   3836
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   400
      MarginTop       =   1440
      MarginRight     =   400
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   7.47549019607843
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VB.PictureBox picEntrega 
      Height          =   1035
      Left            =   3720
      ScaleHeight     =   975
      ScaleWidth      =   495
      TabIndex        =   60
      Tag             =   "visita"
      Top             =   5280
      Width           =   555
      Begin VSFlex6DAOCtl.vsFlexGrid vsEntrega 
         Height          =   855
         Left            =   600
         TabIndex        =   65
         Top             =   480
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1508
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
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
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
   End
   Begin VB.PictureBox picTaller 
      Height          =   1335
      Left            =   4200
      ScaleHeight     =   1275
      ScaleWidth      =   4035
      TabIndex        =   51
      Tag             =   "visita"
      Top             =   4440
      Width           =   4095
      Begin VSFlex6DAOCtl.vsFlexGrid vsTaller 
         Height          =   675
         Left            =   480
         TabIndex        =   62
         Top             =   420
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1191
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
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
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
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
   End
   Begin VB.PictureBox picVisita 
      Height          =   1275
      Left            =   2880
      ScaleHeight     =   1215
      ScaleWidth      =   435
      TabIndex        =   44
      Tag             =   "visita"
      Top             =   4980
      Width           =   495
      Begin VSFlex6DAOCtl.vsFlexGrid vsVisita 
         Height          =   855
         Left            =   60
         TabIndex        =   63
         Top             =   360
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1508
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
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
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
   End
   Begin VB.PictureBox picRetiro 
      Height          =   1455
      Left            =   3360
      ScaleHeight     =   1395
      ScaleWidth      =   435
      TabIndex        =   50
      Tag             =   "visita"
      Top             =   4740
      Width           =   495
      Begin VSFlex6DAOCtl.vsFlexGrid vsRetiro 
         Height          =   855
         Left            =   0
         TabIndex        =   64
         Top             =   360
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1508
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
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
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
   End
   Begin VB.PictureBox picBotones 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   8715
      TabIndex        =   45
      Top             =   6480
      Width           =   8775
      Begin VB.CommandButton bEvento 
         Height          =   310
         Left            =   4920
         Picture         =   "frmSeguimiento.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   70
         TabStop         =   0   'False
         ToolTipText     =   "Historia."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bRefresh 
         Height          =   310
         Left            =   6120
         Picture         =   "frmSeguimiento.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         ToolTipText     =   "Refrescar."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bTraslado 
         Height          =   310
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         ToolTipText     =   "Anular Traslado."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   5640
         Picture         =   "frmSeguimiento.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir Copia."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bComentario 
         Height          =   310
         Left            =   780
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "Comentarios."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bTallerAcepta 
         Height          =   310
         Left            =   4320
         Picture         =   "frmSeguimiento.frx":1720
         Style           =   1  'Graphical
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "Acepta presupuesto."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bCumplir 
         Height          =   310
         Left            =   3960
         Picture         =   "frmSeguimiento.frx":1FEA
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "Cumplir."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bHistoria 
         Height          =   310
         Left            =   5280
         Picture         =   "frmSeguimiento.frx":28B4
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Historia."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bTaller 
         Height          =   310
         Left            =   2400
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         ToolTipText     =   "Taller."
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   310
      End
      Begin VB.CommandButton bEntrega 
         Height          =   310
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Entrega."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bRetiro 
         Height          =   310
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         ToolTipText     =   "Retiro."
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   310
      End
      Begin VB.CommandButton bCliente 
         Height          =   310
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Clientes."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bProducto 
         Height          =   310
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Productos."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   8400
         Picture         =   "frmSeguimiento.frx":317E
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bVisita 
         Height          =   310
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Visita a domicilio."
         Top             =   0
         Width           =   310
      End
   End
   Begin ComctlLib.TabStrip tbEstado 
      Height          =   2055
      Left            =   5160
      TabIndex        =   43
      Top             =   4680
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   3625
      TabFixedWidth   =   3350
      TabFixedHeight  =   441
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tMotivo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   60
      MaxLength       =   11
      TabIndex        =   20
      Top             =   4350
      Width           =   2415
   End
   Begin VB.Frame frmProducto 
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   60
      TabIndex        =   27
      Top             =   2880
      Width           =   8715
      Begin VB.CommandButton cbVisualizacion 
         Caption         =   "..."
         Height          =   285
         Left            =   8160
         TabIndex        =   73
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox tPFCompra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   11
         Top             =   540
         Width           =   975
      End
      Begin VB.TextBox tPFacturaN 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5100
         MaxLength       =   7
         TabIndex        =   14
         Top             =   540
         Width           =   795
      End
      Begin VB.TextBox tPFacturaS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4740
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "AA"
         Top             =   540
         Width           =   315
      End
      Begin VB.TextBox tPNroMaquina 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         MaxLength       =   40
         TabIndex        =   16
         Top             =   540
         Width           =   1875
      End
      Begin VB.TextBox tPArticulo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   5595
      End
      Begin VB.TextBox tPDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   840
         Width           =   5055
      End
      Begin VB.CommandButton bPDireccion 
         BackColor       =   &H8000000E&
         Caption         =   "Dirección&..."
         Height          =   320
         Left            =   7620
         Picture         =   "frmSeguimiento.frx":3280
         TabIndex        =   18
         Top             =   825
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   540
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
         Left            =   900
         TabIndex        =   32
         Top             =   540
         Width           =   645
      End
      Begin VB.Label Label4 
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
         Left            =   900
         TabIndex        =   30
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "F/&Compra:"
         Height          =   255
         Left            =   1620
         TabIndex        =   10
         Top             =   540
         Width           =   855
      End
      Begin VB.Label lDocumento 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº Factura: "
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   540
         Width           =   915
      End
      Begin VB.Label Label11 
         Caption         =   "Nº Ser&ie:"
         Height          =   255
         Left            =   6000
         TabIndex        =   15
         Top             =   540
         Width           =   735
      End
      Begin VB.Label lPTipo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   225
         Width           =   675
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Producto:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lPIdProducto 
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
         Left            =   900
         TabIndex        =   28
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Frame frmServicio 
      Caption         =   "Servicio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2715
      Left            =   60
      TabIndex        =   22
      Top             =   60
      Width           =   8715
      Begin VB.CommandButton bFactura 
         Caption         =   "&Ver Factura"
         Height          =   315
         Left            =   1860
         TabIndex        =   61
         Top             =   555
         Width           =   1155
      End
      Begin VB.TextBox tSTelefono 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6060
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2340
         Width           =   2535
      End
      Begin VB.TextBox tSDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2340
         Width           =   4035
      End
      Begin VB.TextBox tSComentario 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   1140
         MaxLength       =   600
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   900
         Width           =   7450
      End
      Begin VB.TextBox tSCodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Width           =   675
      End
      Begin AACombo99.AACombo cSEstado 
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Top             =   555
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         ListIndex       =   -1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
      End
      Begin VB.Label lblCoordinar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Coordinar entrega"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6480
         TabIndex        =   75
         Top             =   2040
         Width           =   2085
      End
      Begin VB.Label lbComentarioInterno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   1140
         TabIndex        =   72
         Top             =   1470
         Width           =   7455
      End
      Begin VB.Label Label3 
         Caption         =   "Comentario Interno:"
         Height          =   495
         Left            =   120
         TabIndex        =   71
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label lReclamo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999999"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   3060
         TabIndex        =   68
         ToolTipText     =   "Sercicio Reclamo"
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Co&mentarios:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Local Ingreso:"
         Height          =   255
         Left            =   3600
         TabIndex        =   42
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lSLocalIngreso 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4680
         TabIndex        =   41
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Esta&do:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Recepción:"
         Height          =   255
         Left            =   3840
         TabIndex        =   40
         Top             =   240
         Width           =   795
      End
      Begin VB.Label lSFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4680
         TabIndex        =   39
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lSProceso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1860
         TabIndex        =   38
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Modificado:"
         Height          =   255
         Left            =   6180
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lSFModificado 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7080
         TabIndex        =   36
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   6300
         TabIndex        =   35
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lSUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7080
         TabIndex        =   34
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label7 
         Caption         =   "Nº &Servicio:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfonos:"
         Height          =   255
         Left            =   5220
         TabIndex        =   26
         Top             =   2340
         Width           =   855
      End
      Begin VB.Label lNDireccion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2340
         Width           =   705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lSCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Casa 9242557; Celular 099405236"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1140
         TabIndex        =   23
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   5295
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsMotivo 
      Height          =   1725
      Left            =   60
      TabIndex        =   21
      Top             =   4680
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3043
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
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   57
      Top             =   7065
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
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
            Object.Width           =   11456
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Picture         =   "frmSeguimiento.frx":3390
            Key             =   "printer"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10500
      Y1              =   0
      Y2              =   0
   End
   Begin ComctlLib.ImageList imImagenes 
      Left            =   9540
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSeguimiento.frx":34A2
            Key             =   "producto"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSeguimiento.frx":37BC
            Key             =   "historia"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSeguimiento.frx":3AD6
            Key             =   "retiro"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSeguimiento.frx":3DF0
            Key             =   "cliente"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSeguimiento.frx":410A
            Key             =   "visita"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSeguimiento.frx":4424
            Key             =   "entrega"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSeguimiento.frx":46B6
            Key             =   "taller"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSeguimiento.frx":4D00
            Key             =   "cumplir"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSeguimiento.frx":5092
            Key             =   "presupuesto"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSeguimiento.frx":53AC
            Key             =   "comentarios"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSeguimiento.frx":56C6
            Key             =   "imprimir"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSeguimiento.frx":57D8
            Key             =   "traslado"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSeguimiento.frx":5AF2
            Key             =   "evento"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Moti&vos de Servicio"
      Height          =   255
      Left            =   60
      TabIndex        =   19
      Top             =   4120
      Width           =   1755
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuClientes 
         Caption         =   "&Clientes"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MnuProductos 
         Caption         =   "&Productos"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuOpVisOpe 
         Caption         =   "&Visualización de Operaciones"
         Shortcut        =   {F12}
      End
      Begin VB.Menu MnuOpL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIrValidacion 
         Caption         =   "Validación de Presupuesto"
      End
      Begin VB.Menu MnuComentarios 
         Caption         =   "Co&mentarios"
         Shortcut        =   {F11}
      End
      Begin VB.Menu MnuOpL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "&Salir del formulario"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuFichas 
      Caption         =   "&Fichas"
      Begin VB.Menu MnuVisita 
         Caption         =   "&Visita a domicilio"
         Shortcut        =   ^V
      End
      Begin VB.Menu MnuRetiro 
         Caption         =   "&Retiro de productos"
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuTaller 
         Caption         =   "Reparaciones en &Taller"
         Shortcut        =   ^T
      End
      Begin VB.Menu MnuEntrega 
         Caption         =   "&Entrega de productos"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu MnuVarios 
      Caption         =   "&Varios"
      Begin VB.Menu MnuCumplir 
         Caption         =   "&Cumplir Servicio"
         Shortcut        =   ^U
      End
      Begin VB.Menu MnuAceptaP 
         Caption         =   "&Acepta Presupuesto"
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuVaL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuHistoria 
         Caption         =   "&Historia de Servicios"
         Shortcut        =   ^H
      End
      Begin VB.Menu MnuImprimir 
         Caption         =   "&Imprimir Copia del Servicio"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuForzarCumplido 
         Caption         =   "Forzar cumplido"
      End
   End
   Begin VB.Menu MnuEventos 
      Caption         =   "&Eventos"
      Begin VB.Menu MnuEvAddMenu 
         Caption         =   "Agregar Eventos"
         Begin VB.Menu MnuEvAdd 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu MnuEvL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEvShow 
         Caption         =   "Ver Eventos"
      End
   End
End
Attribute VB_Name = "frmSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim gCliente As Long, gProducto As Long, localReparacion As Long
Private iUltimoCodigo As Long 'ultimo cod ingresado

Public prmServicio As Long

Private Sub loc_ProcesoCodBarra()
    
    If Val(lSProceso.Tag) = EstadoS.Taller Then
        If gCliente = paClienteEmpresa Then
            MsgBox "El artículo del servicio es de stock. No se puede cumplir este servicio.", vbExclamation, "Artículo de Stock"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        If (Trim(vsTaller.Cell(flexcpText, 0, 3)) = "" Or Trim(vsTaller.Cell(flexcpText, 2, 1)) = "" Or Trim(vsTaller.Cell(flexcpText, 2, 3)) = "") And InStr(1, vsTaller.Cell(flexcpText, 5, 0), "No reparable") = 0 Then
              bCumplir_Click
        Else
            'Si tiene Costo (>0) y  no esta facturado no se puede dejar cumplir
            If vsTaller.Cell(flexcpData, 2, 1) <> 0 And gCliente <> paClienteEmpresa And Val(bFactura.Tag) = 0 And vsTaller.Cell(flexcpForeColor, 2, 3) <> Colores.RojoClaro Then
                MsgBox "Este servicio tiene costo de reparación y no está facturado." & Chr(vbKeyReturn) & _
                            "Para cumplirlo debe facturar el servicio.", vbExclamation, "Servicio No Facturado"
                Screen.MousePointer = 0: Exit Sub
            End If
            loc_validarUsuario
        End If
    Else
        MsgBox "POSBILE ERROR" & vbCrLf & vbCrLf & "El servicio no está en taller", vbCritical, "Atención"
    End If
    
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bCliente_Click()
    
    On Error GoTo errCliente
    If gCliente <> 0 Then
        Screen.MousePointer = 11
        Dim objC As New clsCliente
        
        Select Case Val(lSCliente.Tag)
            Case TipoCliente.Cliente: objC.Personas gCliente
            Case TipoCliente.Empresa: objC.Empresas gCliente
        End Select
        
        Set objC = Nothing
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errCliente:
    clsGeneral.OcurrioError "Ocurrió un error al activar el formulario de clientes.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bComentario_Click()
On Error GoTo errCliente
    If gCliente <> 0 Then
        Screen.MousePointer = 11
        Dim objC As New clsCliente
        objC.Comentarios idCliente:=gCliente
        Set objC = Nothing
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errCliente:
    clsGeneral.OcurrioError "Ocurrió un error al activar el formulario de comentarios.", Err.Description
    Screen.MousePointer = 0
End Sub

Sub CumplirCompañia(ByVal localRepara As Long)
    
    If MsgBox("El servicio es de compañia/Neotrón pasos a seguir:" & vbCrLf & vbTab & "1) Traslado a su local" & vbCrLf & vbTab & "2) Cambio de estado a SANO" & vbCrLf & vbCrLf & "¿Confirma cumplir el servicio?", vbQuestion + vbYesNo, "CUMPLIR COMPAÑIA/NEOTRON") = vbYes Then
        FechaDelServidor
        
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        If paCodigoDeUsuario = 0 Then Exit Sub
        
        Screen.MousePointer = 11
        On Error GoTo errorBT
        cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
        On Error GoTo errorET
    
        Dim idProducto As Long
        'Tabla Servicio-----------------------------------------------------------------------------------------------------
        Cons = "Select * from Servicio Where SerCodigo = " & Val(tSCodigo.Text)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        idProducto = RsAux("SerProducto")
        RsAux.Edit
        RsAux!SerFCumplido = Format(gFechaServidor, sqlFormatoF)
        RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsAux!SerUsuario = paCodigoDeUsuario
        RsAux!SerEstadoServicio = EstadoS.Cumplido
        RsAux.Update
        RsAux.Close

        'Tengo que hacer el traslado al local.
        Dim IDArticulo As Long
        Cons = "SELECT ProArticulo FROM Producto WHERE ProCodigo = " & idProducto
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        IDArticulo = RsAux(0)
        RsAux.Close
        
        HagoTraslado IDArticulo, Val(tSCodigo.Text), localRepara
        
        'Tengo que cambiar el estado a sano.
        MarcoMovimientoStockFisico paCodigoDeUsuario, 2, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, -1, TipoDocumento.ServicioCambioEstado, Val(tSCodigo.Text)
        MarcoMovimientoStockFisico paCodigoDeUsuario, 2, paCodigoDeSucursal, IDArticulo, 1, paEstadoArticuloEntrega, 1, TipoDocumento.ServicioCambioEstado, Val(tSCodigo.Text)
            
        MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, paEstadoARecuperar, 1, -1
        MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, 1, 1
        
        MarcoMovimientoStockFisicoEnLocal 2, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, -1
        MarcoMovimientoStockFisicoEnLocal 2, paCodigoDeSucursal, IDArticulo, 1, paEstadoArticuloEntrega, 1
            
        cBase.CommitTrans
        Screen.MousePointer = 0
'        On Error Resume Next
        CargoDatosServicio Val(tSCodigo.Text)
    End If
    Exit Sub

errorBT:
    MsgBox "No se logró iniciar la transacción." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    MsgBox "No se ha podido grabar el servicio." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description

    
End Sub

Private Sub HagoTraslado(IDArticulo As Long, idServicio As Long, idLocalInicial As Long)
    
    MarcoMovimientoStockFisico paCodigoDeUsuario, 2, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, 1, TipoDocumento.ServicioCambioEstado, idServicio
    MarcoMovimientoStockFisicoEnLocal 2, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, 1
    
    MarcoMovimientoStockFisico paCodigoDeUsuario, 2, idLocalInicial, IDArticulo, 1, paEstadoARecuperar, -1, TipoDocumento.ServicioCambioEstado, idServicio
    MarcoMovimientoStockFisicoEnLocal 2, idLocalInicial, IDArticulo, 1, paEstadoARecuperar, -1

End Sub

Private Sub bCumplir_Click()
    
    'Cumplir Servicio------------------------------------------------------------------------
    
    Select Case Val(lSProceso.Tag)
        Case Is <> EstadoS.Taller
            frmCumplir.prmServicio = Val(tSCodigo.Text)
            frmCumplir.Show vbModal, Me
            Me.Refresh
            
            Screen.MousePointer = 11
            If ServicioModificado(Val(tSCodigo.Text)) Then
                LimpioCampos
                CargoDatosServicio Val(tSCodigo.Text)
            End If
            Screen.MousePointer = 0
            
        Case EstadoS.Taller
            
            If gCliente = paClienteEmpresa Then
                If localReparacion = paLocalCompañia Or localReparacion = 17 Then
                    CumplirCompañia localReparacion
                Else
                    MsgBox "El artículo del servicio es de stock. No se puede cumplir este servicio.", vbExclamation, "Artículo de Stock"
                End If
                Screen.MousePointer = 0: Exit Sub
            End If
            
            'If Trim(lTFReparado.Caption) = "" Or Trim(tTCosto.Text) = "" Or Trim(lTFAceptado.Caption) = "" Or Trim(lTFPresupuesto.Caption) = "" Then
            If vsTaller.Cell(flexcpForeColor, 2, 3) <> Colores.RojoClaro And ((Trim(vsTaller.Cell(flexcpText, 0, 3)) = "" Or Trim(vsTaller.Cell(flexcpText, 2, 1)) = "" Or Trim(vsTaller.Cell(flexcpText, 2, 3)) = "") And InStr(1, vsTaller.Cell(flexcpText, 5, 0), "No reparable") = 0) Then
                'MsgBox "El proceso de taller no está completo. No se podrá dar el servicio como entregado.", vbExclamation, "Proceso Incompleto"
                
                If localReparacion = paLocalCompañia Or localReparacion = 17 Then
                    Dim res As VbMsgBoxResult
                    res = MsgBox("El servicio es de CLIENTE a reparar en compañia." & vbCrLf & vbCrLf & "Presione:" & _
                                                vbCrLf & "SI --> si el producto fue reparado" & _
                                                vbCrLf & "NO --> si el producto NO fue reparado (no tiene arreglo)" & _
                                                vbCrLf & "Cancelar --> para cancelar la acción", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Servicio de cliente")
                    If res = vbCancel Then Exit Sub
                    
                    GraboRecepcionCompañiaAEntrega (res = vbYes)
                    
                    Exit Sub
                    
                End If
                
                If Trim(vsTaller.Cell(flexcpText, 2, 3)) = "" Then
                    If MsgBox("El presupuesto no fue aceptado, si el cliente lo aceptó y se reparó debe confirmar esta acción." & vbCrLf & vbCrLf & "¿Desea continuar anulando el servicio?", vbQuestion + vbYesNo + vbDefaultButton2, "Aceptar presupuesto") = vbNo Then Exit Sub
                    'MsgBox "Debe aceptar el presupuesto, es el botón siguiente.", vbExclamation, "ATENCIÓN"
'                    Exit Sub
                End If
                
                If MsgBox("El proceso de taller no está completo. Si ud. lo cumple va a quedar anulado." & Chr(vbKeyReturn) & _
                                "Desea continuar con la acción.", vbExclamation + vbYesNo + vbDefaultButton2, "Proceso Incompleto") = vbYes Then
                                
                    If MsgBox("UD va a ANULAR EL SERVICIO. ¿ESTÁ SEGURO?", vbQuestion + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbYes Then
                        frmCumplir.prmServicio = Val(tSCodigo.Text)
                        frmCumplir.Show vbModal, Me
                        Me.Refresh
                        
                        Screen.MousePointer = 11
                        If ServicioModificado(Val(tSCodigo.Text)) Then
                            LimpioCampos
                            CargoDatosServicio Val(tSCodigo.Text)
                        End If
                        Screen.MousePointer = 0
                    End If
                End If
                
            Else
                'Si tiene Costo (>0) y  no esta facturado no se puede dejar cumplir
                If vsTaller.Cell(flexcpData, 2, 1) <> 0 And gCliente <> paClienteEmpresa And Val(bFactura.Tag) = 0 And vsTaller.Cell(flexcpForeColor, 2, 3) <> Colores.RojoClaro Then
                    MsgBox "Este servicio tiene costo de reparación y no está facturado." & Chr(vbKeyReturn) & _
                                "Para cumplirlo debe facturar el servicio.", vbExclamation, "Servicio No Facturado"
                    Screen.MousePointer = 0: Exit Sub
                End If

                If MsgBox("Se va a cumplir la reparación en taller y dar como terminado el servicio." & Chr(vbKeyReturn) & _
                              "(*) Confirme que el cliente retira el producto." & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                              "Desea cumplir el servicio.", vbQuestion + vbYesNo + vbDefaultButton2, "Cumplir Reparación") = vbNo Then Exit Sub
                On Error GoTo errGrabar
                
                Screen.MousePointer = 11
                FechaDelServidor
                'Tabla Servicio-----------------------------------------------------------------------------------------------------
                Cons = "Select * from Servicio Where SerCodigo = " & Val(tSCodigo.Text)
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                RsAux.Edit
                
                If IsNull(RsAux!SerComentarioR) Then RsAux!SerComentarioR = "Retira el cliente."
                RsAux!SerFCumplido = Format(gFechaServidor, sqlFormatoF)
                RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
                RsAux!SerUsuario = paCodigoDeUsuario
                RsAux!SerEstadoServicio = EstadoS.Cumplido
                RsAux.Update: RsAux.Close
                
                If ServicioModificado(Val(tSCodigo.Text)) Then
                    LimpioCampos
                    CargoDatosServicio Val(tSCodigo.Text)
                End If
                Screen.MousePointer = 0
            End If
    End Select
    Exit Sub
    
errGrabar:
    clsGeneral.OcurrioError "Ocurrió un error al grabar los datos del cumplido.", Err.Description
    Screen.MousePointer = 0
End Sub

Sub GraboRecepcionCompañiaAEntrega(ByVal bReparado As Boolean)

    On Error GoTo errorBT
    cBase.BeginTrans
    On Error GoTo errorET
    
    Cons = "Select * from Servicio Where SerCodigo = " & Val(tSCodigo.Text)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
    RsAux!SerUsuario = paCodigoDeUsuario
    RsAux("SerEstadoServicio") = 5
    RsAux.Update
    RsAux.Close
    
    Cons = "SELECT * FROM Taller WHERE TalServicio = " & Val(tSCodigo.Text)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.AddNew
        RsAux("TalFIngresoRealizado") = Format(Now, "yyyy/mm/dd hh:nn:ss")
        RsAux("TalServicio") = Val(tSCodigo.Text)
        RsAux("TalUsuario") = paCodigoDeUsuario
    Else
        RsAux.Edit
    End If
    RsAux("TalFSalidaRealizado") = Format(Now, "yyyy/mm/dd hh:nn:ss")
    RsAux("TalFSalidaRecepcion") = Format(Now, "yyyy/mm/dd hh:nn:ss")
    RsAux("TalLocalAlCLiente") = paCodigoDeSucursal
    If Not bReparado Then RsAux("TalSinArreglo") = 1
    RsAux("TalFReparado") = Format(Now, "yyyy/mm/dd hh:nn:ss")
    RsAux("TalModificacion") = Format(Now, "yyyy/mm/dd hh:nn:ss")
    RsAux("TalFAceptacion") = Format(Now, "yyyy/mm/dd hh:nn:ss")
    RsAux.Update
    RsAux.Close
    
    cBase.CommitTrans
    CargoDatosServicio Val(tSCodigo.Text)
    Exit Sub


errorBT:
    MsgBox "No se logró iniciar la transacción." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    MsgBox "No se ha podido grabar el servicio." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description


End Sub

Private Sub bEntrega_Click()
    
    On Error GoTo errServicio
    
    If gCliente = paClienteEmpresa Then
        MsgBox "El artículo del servicio es de stock. No se puede realizar la entrega.", vbExclamation, "Artículo de Stock"
        Screen.MousePointer = 0: Exit Sub
    End If

    If Val(lSProceso.Tag) = EstadoS.Taller Then
    
        If InStr(1, vsTaller.Cell(flexcpText, 5, 0), "No reparable") > 0 Then
            'Paso
        ElseIf Trim(vsTaller.Cell(flexcpText, 0, 3)) = "" Or Trim(vsTaller.Cell(flexcpText, 2, 1)) = "" Or Trim(vsTaller.Cell(flexcpText, 2, 3)) = "" Then
            If MsgBox("El proceso de taller no está completo. Está seguro de lo que quiere hacer ?." & Chr(vbKeyReturn) & _
                            "Realmente desea continuar.", vbExclamation + vbYesNo + vbDefaultButton2, "Proceso Incompleto") = vbNo Then Exit Sub
        End If
        
        Screen.MousePointer = 11
        Dim aEntrega As String, aTexto As String, bSalir As Boolean
        aEntrega = "": aTexto = "": bSalir = False
        'Valido si se puede hacer la entrega
        Cons = "Select * from Taller Left Outer Join Local on TalLocalAlCLiente = LocCodigo" & _
                                                " Left Outer Join Camion on TalSalidaCamion = CamCodigo" & _
                   " Where TalServicio = " & Val(tSCodigo.Text)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If Not IsNull(RsAux!TalLocalAlCliente) Then
                aEntrega = "La mercadería se entregará desde el local " & Trim(RsAux!LocNombre) & Chr(vbKeyReturn) & "Si desea cambiar el local de entrega acceda a la ficha de taller" & Chr(vbKeyReturn) & Chr(vbKeyReturn) & "(*) El camión que realice la entrega deberá pasar por este local a levantarla"
            Else
                aEntrega = "El local de entrega al cliente no se ha ingresado. " & Chr(vbKeyReturn) & _
                                "Si ud. desea inrgesarlo acceda a la ficha de taller (antes de hacer la entrega)." & Chr(vbKeyReturn) & Chr(vbKeyReturn) & _
                                "Este local indica desde dónde se va a recoger el producto para entregarlo."
                bSalir = True
            End If
            If Not IsNull(RsAux!TalFSalidaRealizado) Then
                'Hay Raslado desde el local de reparacion
                If IsNull(RsAux!TalFSalidaRecepcion) Then
                    aTexto = "La mercadería está en traslado desde el local de reparación hacia el local de entrega." & Chr(vbKeyReturn) & _
                                 "Para realizar la entrega, se debe recepcionar el traslado."
                    bSalir = True
                End If
            End If
        End If
        RsAux.Close
        
        If bSalir Then
            If Trim(aTexto) <> "" Then MsgBox aTexto, vbExclamation, "Mercadería en Traslado"
            If Trim(aEntrega) <> "" Then MsgBox aEntrega, vbExclamation, "Falta local de Entrega"
            Screen.MousePointer = 0: Exit Sub
        Else
            MsgBox aEntrega, vbExclamation, "Entrega de Mercadería"
        End If
        Screen.MousePointer = 0
    End If
    
    
    
    frmEntrega.prmServicio = Val(tSCodigo.Text)
    frmEntrega.Show vbModal, Me
    Me.Refresh
    
    
    'Primero veo si cambio el estado del servicio ej: de taller a entrega
    Dim bCambioEstado As Boolean: bCambioEstado = False
    Cons = "Select * from Servicio Where SerCodigo = " & Val(tSCodigo.Text)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux!SerEstadoServicio <> Val(lSProceso.Tag) Then bCambioEstado = True
    RsAux.Close
        
    If Not bCambioEstado Then
        ActualizoDatosFicha Val(tSCodigo.Text), TipoServicio.Entrega
    Else
        CargoDatosServicio Val(tSCodigo.Text)
    End If
    Exit Sub

errServicio:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del servicio.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bEvento_Click()
    MnuEvShow_Click
End Sub

Private Sub bFactura_Click()
    EjecutarApp App.Path & "\Detalle de factura", CStr(bFactura.Tag)
End Sub

Private Sub bHistoria_Click()
    EjecutarApp App.Path & "\Historia Servicio", CStr(gProducto)
End Sub

Private Sub bImprimir_Click()
    
    If Val(tSCodigo.Tag) <> 0 Then
        If MsgBox("Desea imprimir una copia para el servicio Nº " & tSCodigo.Text, vbQuestion + vbYesNo, "Imprimir Copia") = vbNo Then Exit Sub
        ImprimoCopia
    End If
    
End Sub

Private Sub bPDireccion_Click()

Dim aDirAnterior As Long, aRetorno As Long
    
    On Error GoTo errDirecccion
    If Val(lPIdProducto.Caption) = 0 Then Exit Sub
    
    Screen.MousePointer = 11
    aDirAnterior = Val(tPDireccion.Tag)
    
    Dim objDireccion As New clsDireccion
    objDireccion.ActivoFormularioDireccion cBase, aDirAnterior, gCliente, "Producto", "ProDireccion", "ProCodigo", gProducto
    Me.Refresh
    aRetorno = objDireccion.CodigoDeDireccion
    Set objDireccion = Nothing
    
    If aDirAnterior <> aRetorno Then
        If aRetorno <> 0 Then
            Cons = "Update Producto Set ProDireccion = " & aRetorno & " Where ProCodigo = " & gProducto
        Else
            Cons = "Update Producto Set ProDireccion = Null Where ProCodigo = " & gProducto
        End If
        cBase.Execute Cons
    End If
    
    If aRetorno <> 0 Then
        tPDireccion.Text = clsGeneral.ArmoDireccionEnTexto(cBase, aRetorno, True, True, True)
    Else
        tPDireccion.Text = ""
    End If
    tPDireccion.Tag = aRetorno
    
    Screen.MousePointer = 0
    Exit Sub
    
errDirecccion:
    clsGeneral.OcurrioError "Ocurrió un error al cargar la dirección.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bProducto_Click()
    If Val(tSCodigo.Tag) = 0 Then Exit Sub
    EjecutarApp App.Path & "\Productos.exe", CStr(gCliente) & ";p" & Val(lPIdProducto.Caption)
End Sub

Private Sub bRefresh_Click()
On Error Resume Next
    LimpioCampos
    If Val(tSCodigo.Text) > 0 Then CargoDatosServicio Val(tSCodigo.Text)
End Sub

Private Sub bRetiro_Click()
    frmRetiro.prmServicio = Val(tSCodigo.Text)
    frmRetiro.Show vbModal, Me
    Me.Refresh
    ActualizoDatosFicha Val(tSCodigo.Text), TipoServicio.Retiro
End Sub

Private Sub bTaller_Click()
    frmTaller.prmServicio = Val(tSCodigo.Text)
    frmTaller.prmAceptaPto = False
    frmTaller.Show vbModal, Me
    Me.Refresh
    ActualizoDatosFicha Val(tSCodigo.Text), 0, True
End Sub

Private Sub bTallerAcepta_Click()

    frmTaller.prmServicio = Val(tSCodigo.Text)
    frmTaller.prmAceptaPto = True
    frmTaller.Show vbModal, Me
    Me.Refresh
    ActualizoDatosFicha Val(tSCodigo.Text), 0, True
    
End Sub

Private Sub bTraslado_Click()

    If Not miConexion.AccesoAlMenu("Cambio Sucursal") Then
        MsgBox "Ud. no está autorizado a eliminar el traslado del producto.", vbExclamation, "Sin Autorización"
        Exit Sub
    End If
    
    If MsgBox("Confirma eliminar el traslado del producto al local de entrega." & vbCrLf & vbCrLf & _
                    "Si ud. presiona 'SI' el producto volverá a estar en taller.", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Traslado") = vbNo Then Exit Sub
               
    On Error GoTo errGrabar
    Screen.MousePointer = 11
    Cons = "Select * from Taller Where TalServicio = " & Val(tSCodigo.Text)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Edit
        RsAux!TalFSalidaRealizado = Null
        RsAux!TalFSalidaRecepcion = Null
        RsAux!TalSalidaCamion = Null
        RsAux.Update
    Else
        MsgBox "No se encontró el registro para modificarlo.", vbExclamation, "Posible Error"
    End If
    RsAux.Close
    
    ActualizoDatosFicha Val(tSCodigo.Text), 0, True
    
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    clsGeneral.OcurrioError "Ocurrió un error al eliminar el traslado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bVisita_Click()
    
    frmVisita.prmServicio = Val(tSCodigo.Text)
    frmVisita.Show vbModal, Me
    Me.Refresh
    
     ActualizoDatosFicha Val(tSCodigo.Text), TipoServicio.Visita
     
End Sub

Private Sub ActualizoDatosFicha(idServicio As Long, Tipo As Integer, Optional Taller As Boolean = False)
    
    On Error GoTo errCargar
    Screen.MousePointer = 11
    
    'Valido si cambió el estado del Servicio-------------------------------------------------------------
    Dim bCambioEstado As Boolean
    bCambioEstado = True
    Cons = "Select * from Servicio Where SerCodigo = " & idServicio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux!SerEstadoServicio = Val(lSProceso.Tag) Then bCambioEstado = False
    End If
    RsAux.Close
    '--------------------------------------------------------------------------------------------------------------------------
    
    If bCambioEstado Then
        CargoDatosServicio Val(tSCodigo.Text)
        Screen.MousePointer = 0: Exit Sub
    End If
            
    If Not Taller Then
        Cons = "Select * from ServicioVisita " & _
                   " Where VisServicio = " & idServicio & " And VisTipo = " & Tipo & _
                   " Order by VisCodigo DESC"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            Select Case Tipo
                Case TipoServicio.Retiro:
                            LimpioCamposFichas Visita:=False, Retiro:=True, Entrega:=False, Taller:=False
                            CargoCamposTRetiro
                Case TipoServicio.Visita:
                            LimpioCamposFichas Visita:=True, Retiro:=False, Entrega:=False, Taller:=False
                            CargoCamposTVisita
                Case TipoServicio.Entrega:
                            LimpioCamposFichas Visita:=False, Retiro:=False, Entrega:=True, Taller:=False
                            CargoCamposTEntrega
            End Select
        End If
        RsAux.Close
        
    Else
        Cons = "Select * from Servicio Left Outer Join Taller On SerCodigo = TalServicio " & _
                   " Where SerCodigo = " & idServicio & _
                   " And SerLocalReparacion Is Not Null"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

        If Not RsAux.EOF Then
            LimpioCamposFichas Visita:=False, Retiro:=False, Entrega:=False, Taller:=True
            CargoCamposTTaller
        End If
        RsAux.Close
    End If
    
    Screen.MousePointer = 0
    Exit Sub

errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al procesar los datos de la ficha.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cbVisualizacion_Click()
On Error Resume Next
    If Val(cbVisualizacion.Tag) > 0 Then EjecutarApp App.Path & "\visualizacion de operaciones.exe", cbVisualizacion.Tag
End Sub

Private Sub cSEstado_Change()
    cSEstado.Tag = 1
End Sub

Private Sub cSEstado_GotFocus()
    Status.Panels(3).Text = "Estado del producto para el servicio seleccionado."
End Sub

Private Sub cSEstado_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(cSEstado.Tag) <> 0 Then
            If cSEstado.ListIndex = -1 Then MsgBox "El estado seleccionado no es correcto. Verifique", vbInformation, "ATENCIÓN": Exit Sub
            ZActualizoCampoServicio Val(tSCodigo.Tag), cSEstado.ItemData(cSEstado.ListIndex), Estado:=True
            cSEstado.Tag = 0
        End If
        Foco tSComentario
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    
    Status.Panels("printer").Text = paINContadoN
    
    ObtengoSeteoForm Me
    Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    
    vsMotivo.BackColor = Colores.Gris
    Me.Width = 8970: Me.Height = 8010 '7185
    LimpioCampos
    
    FechaDelServidor
    InicializoGrillas
    CargoCombos
    
    imImagenes.UseMaskColor = True
    Set tbEstado.ImageList = imImagenes
    
    bCliente.Picture = imImagenes.ListImages("cliente").ExtractIcon
    bProducto.Picture = imImagenes.ListImages("producto").ExtractIcon
    bComentario.Picture = imImagenes.ListImages("comentarios").ExtractIcon
    bHistoria.Picture = imImagenes.ListImages("historia").ExtractIcon
    bEvento.Picture = imImagenes.ListImages("evento").ExtractIcon
    
    bVisita.Picture = imImagenes.ListImages("visita").ExtractIcon
    bRetiro.Picture = imImagenes.ListImages("retiro").ExtractIcon
    bTaller.Picture = imImagenes.ListImages("taller").ExtractIcon
    bEntrega.Picture = imImagenes.ListImages("entrega").ExtractIcon
    bCumplir.Picture = imImagenes.ListImages("cumplir").ExtractIcon
    bTallerAcepta.Picture = imImagenes.ListImages("presupuesto").ExtractIcon
    bImprimir.Picture = imImagenes.ListImages("imprimir").ExtractIcon
    bTraslado.Picture = imImagenes.ListImages("traslado").ExtractIcon
    
    s_LoadMenuEventos
    
    If prmServicio <> 0 Then
        tSCodigo.Text = prmServicio
        CargoDatosServicio prmServicio
    End If
    lbComentarioInterno.BackColor = Inactivo
End Sub

Private Sub LimpioCampos()
    
    On Error Resume Next
    tSCodigo.Tag = 0
    gCliente = 0: gProducto = 0: cbVisualizacion.Tag = "": localReparacion = 0
    
    If frmEventos.Visible Then frmEventos.s_Clean
    
    bCliente.Enabled = False: bProducto.Enabled = False: bComentario.Enabled = False
    bVisita.Enabled = False: bRetiro.Enabled = False: bTaller.Enabled = False: bEntrega.Enabled = False: bTallerAcepta.Enabled = False
    bTraslado.Enabled = False: bImprimir.Enabled = False
    bCumplir.Enabled = False: bHistoria.Enabled = False
    bRefresh.Enabled = False
    
    lbComentarioInterno.Caption = ""
    LimpioCamposFichas
    ColorCamposFichas Visita:=vbInactiveBorder, Retiro:=vbInactiveBorder, Entrega:=vbInactiveBorder, Taller:=vbInactiveBorder
    
    tbEstado.Tabs.Clear
    tbEstado.Tabs.Add pvcaption:="Visita"
    picVisita.ZOrder 0
    lReclamo.Caption = ""
    lblCoordinar.Visible = False
    lSProceso.Caption = "": lSFecha.Caption = "": lSFModificado.Caption = ""
    cSEstado.Text = "": lSLocalIngreso.Caption = "": lSUsuario.Caption = ""
    tSComentario.Text = ""
    lSCliente.Caption = "": tSDireccion.Text = "": tSTelefono.Text = ""
    cSEstado.Tag = 0: tSComentario.Tag = 0

    lPIdProducto.Caption = "": tPArticulo.Text = ""
    lPEstado.Caption = "": tPFCompra.Text = "": tPFacturaS.Text = "": tPFacturaN.Text = "": tPNroMaquina.Text = ""
    lPGarantia.Caption = "": tPDireccion.Text = "": bPDireccion.Enabled = False: bPDireccion.Tag = 0
    tPFCompra.Tag = 0: tPFacturaS.Tag = 0: tPFacturaN.Tag = 0: tPNroMaquina.Tag = 0
    
    vsMotivo.Rows = 1
    tMotivo.Text = ""
    
    EstadoControles Estado:=False, ColorFondo:=Colores.Gris
    bFactura.Enabled = False: bFactura.Tag = 0
    EstadoMenu
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Panels(3).Text = ""
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    picBotones.BorderStyle = vbBSNone
    picBotones.Top = Me.ScaleHeight - picBotones.Height - Status.Height
    
    tbEstado.Top = frmProducto.Height + frmServicio.Height + 200
    tbEstado.Left = vsMotivo.Left + vsMotivo.Width + 100
    tbEstado.Width = Me.ScaleWidth - tbEstado.Left - 60
    tbEstado.Height = Me.ScaleHeight - tbEstado.Top - picBotones.Height - 120 - Status.Height

    picVisita.Top = tbEstado.ClientTop: picVisita.Left = tbEstado.ClientLeft
    picVisita.Width = tbEstado.ClientWidth: picVisita.Height = tbEstado.ClientHeight
    picVisita.BorderStyle = vbBSNone
    picTaller.Top = tbEstado.ClientTop: picTaller.Left = tbEstado.ClientLeft
    picTaller.Width = tbEstado.ClientWidth: picTaller.Height = tbEstado.ClientHeight
    picTaller.BorderStyle = vbBSNone
    picRetiro.Top = tbEstado.ClientTop: picRetiro.Left = tbEstado.ClientLeft
    picRetiro.Width = tbEstado.ClientWidth: picRetiro.Height = tbEstado.ClientHeight
    picRetiro.BorderStyle = vbBSNone
    picEntrega.Top = tbEstado.ClientTop: picEntrega.Left = tbEstado.ClientLeft
    picEntrega.Width = tbEstado.ClientWidth: picEntrega.Height = tbEstado.ClientHeight
    picEntrega.BorderStyle = vbBSNone
    
    vsTaller.Top = 30: vsTaller.Left = 0: vsTaller.Width = picTaller.Width: vsTaller.Height = picTaller.Height - 30
    vsVisita.Top = 30: vsVisita.Left = 0: vsVisita.Width = vsTaller.Width: vsVisita.Height = vsTaller.Height
    vsRetiro.Top = 30: vsRetiro.Left = 0: vsRetiro.Width = vsTaller.Width: vsRetiro.Height = vsTaller.Height
    vsEntrega.Top = 30: vsEntrega.Left = 0: vsEntrega.Width = vsTaller.Width: vsEntrega.Height = vsTaller.Height
    
    vsRetiro.BorderStyle = flexBorderNone: vsVisita.BorderStyle = flexBorderNone: vsTaller.BorderStyle = flexBorderNone: vsEntrega.BorderStyle = flexBorderNone
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    If frmEventos.Visible Then Unload frmEventos
    CierroConexion
    
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    GuardoSeteoForm Me
    End
    
End Sub

Private Sub InicializoGrillas()

    With vsMotivo
        .Rows = 1: .Cols = 1
        .FormatString = "<Motivos"
        .ColWidth(0) = 1400
        .WordWrap = False
        .ExtendLastCol = True
    End With
    
    With vsTaller
        .Rows = 7: .Cols = 5
        .ColWidth(0) = 1100: .ColWidth(1) = 1800: .ColWidth(2) = 1100: .ColWidth(3) = 1600
        '.WordWrap = False
        .ExtendLastCol = True
        .FocusRect = flexFocusNone: .ScrollBars = flexScrollBarNone: .MergeCells = flexMergeSpill
    End With
    With vsVisita
        .Rows = 7: .Cols = 5
        .ColWidth(0) = 1200: .ColWidth(1) = 2200: .ColWidth(2) = 800: .ColWidth(3) = 1600
        .WordWrap = False: .ExtendLastCol = True
        .FocusRect = flexFocusNone: .ScrollBars = flexScrollBarNone: .MergeCells = flexMergeSpill
    End With
    With vsRetiro
        .Rows = 7: .Cols = 5
        .ColWidth(0) = 1050: .ColWidth(1) = 2300: .ColWidth(2) = 800: .ColWidth(3) = 1600
        .WordWrap = False: .ExtendLastCol = True
        .FocusRect = flexFocusNone: .ScrollBars = flexScrollBarNone: .MergeCells = flexMergeSpill
    End With
    With vsEntrega
        .Rows = 7: .Cols = 5
        .ColWidth(0) = 1050: .ColWidth(1) = 2300: .ColWidth(2) = 800: .ColWidth(3) = 1600
        .WordWrap = False: .ExtendLastCol = True
        .FocusRect = flexFocusNone: .ScrollBars = flexScrollBarNone: .MergeCells = flexMergeSpill
    End With
End Sub

Private Sub CargoInfoDeFulano(ByVal iCli As Long)
    Screen.MousePointer = 11
    'LimpioCamposCliente
    
    If iCli = paClienteEmpresa Then
        cbVisualizacion.Tag = iCli
        tPArticulo.Text = tPArticulo.Text + " (De: STOCK)"
        Exit Sub
    End If
    
    'Ficha del Cliente----------------------------------------------------------------------------------------------------------------
     Cons = "Select Nombre = (RTrim(CPeApellido1)+ ', ' + RTrim(CPeNombre1))  " _
           & " From Cliente, CPersona " _
           & " Where CliCodigo = " & iCli _
           & " And CliCodigo = CPeCliente " _
                                                & " UNION All" _
           & " Select Nombre = RTrim(CEmFantasia)" _
           & " From Cliente, CEmpresa " _
           & " Where CliCodigo = " & iCli _
           & " And CliCodigo = CEmCliente"

    Dim rsC As rdoResultset
    Set rsC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsC.EOF Then
        cbVisualizacion.Tag = iCli
        tPArticulo.Text = tPArticulo.Text + " (De: " + Trim(rsC("Nombre")) + ")"
    End If
    rsC.Close
    

End Sub

Private Sub CargoDatosCliente(idCliente As Long)
    
    On Error GoTo errCliente
    Screen.MousePointer = 11
    'LimpioCamposCliente
    
    'Ficha del Cliente----------------------------------------------------------------------------------------------------------------
     Cons = "Select Cliente.*, Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
           & " From Cliente, CPersona " _
           & " Where CliCodigo = " & idCliente _
           & " And CliCodigo = CPeCliente " _
                                                & " UNION All" _
           & " Select Cliente.*, Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
           & " From Cliente, CEmpresa " _
           & " Where CliCodigo = " & idCliente _
           & " And CliCodigo = CEmCliente"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!CliCIRuc) Then
            If RsAux!CliTipo = TipoCliente.Cliente Then lSCliente.Caption = "  (" & clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc) & ")"
            If RsAux!CliTipo = TipoCliente.Empresa Then lSCliente.Caption = "  (" & clsGeneral.RetornoFormatoRuc(RsAux!CliCIRuc) & ")"
        End If
    End If
    lSCliente.Caption = " " & Trim(RsAux!Nombre) & lSCliente.Caption
    lSCliente.Tag = RsAux!CliTipo
    
    tSDireccion.Tag = 0
    If Not IsNull(RsAux!CliDireccion) Then
        tSDireccion.Text = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion, True, True, True)
        tSDireccion.Tag = RsAux!CliDireccion
    End If
    
    RsAux.Close
    '----------------------------------------------------------------------------------------------------------------------------------
    tSTelefono.Text = TelefonoATexto(idCliente)
    loc_FindComentarios idCliente
    Screen.MousePointer = 0
    Exit Sub
    
errCliente:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatosProducto(idProducto As Long)
    
    On Error GoTo ErrCE
    Screen.MousePointer = 11
    'LimpioCamposProducto
    cbVisualizacion.Tag = ""
    Cons = "Select * from Producto, Articulo " _
            & " Where ProCodigo = " & idProducto _
            & " And ProArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If gCliente = 0 Then gCliente = RsAux!ProCliente
        bPDireccion.Enabled = True
        
        lPIdProducto.Caption = " " & Format(idProducto, "000")
        tPArticulo.Text = Format(RsAux("ArtCodigo"), "###,###") & " " & Trim(RsAux!ArtNombre)
        tPArticulo.Tag = RsAux!ArtId
        lPTipo.Tag = RsAux!ArtTipo      'Tipo del Articulo para ingreso de motivos
        
        If Not IsNull(RsAux!ProCompra) Then tPFCompra.Text = Format(RsAux!ProCompra, "dd/mm/yyyy")
        If Not IsNull(RsAux!ProFacturaS) Then tPFacturaS.Text = Trim(RsAux!ProFacturaS)
        If Not IsNull(RsAux!ProFacturaN) Then tPFacturaN.Text = RsAux!ProFacturaN
        If Not IsNull(RsAux!ProNroSerie) Then tPNroMaquina.Text = Trim(RsAux!ProNroSerie)
        If Not IsNull(RsAux!ProDireccion) Then
            tPDireccion.Text = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!ProDireccion, True, True, True)
            tPDireccion.Tag = RsAux!ProDireccion
        End If
        
        lPGarantia.Caption = " " & RetornoGarantia(RsAux!ArtId)
        lPEstado.Tag = CalculoEstadoProducto(RsAux!ProCodigo)
        lPEstado.Caption = " " & EstadoProducto(Val(lPEstado.Tag))
        
        tPFCompra.Tag = 0: tPFacturaS.Tag = 0: tPFacturaN.Tag = 0: tPNroMaquina.Tag = 0
        
        'Veo si tiene Id de Documento de Compra------------------------------------------------------------------------
        If Not IsNull(RsAux!ProDocumento) Then lDocumento.Tag = RsAux!ProDocumento Else lDocumento.Tag = 0
        '-----------------------------------------------------------------------------------------------------------------------
        
        If gCliente <> RsAux!ProCliente Then
            CargoInfoDeFulano RsAux("ProCliente")
        End If
        
    End If
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub

ErrCE:
    clsGeneral.OcurrioError "Ocurrió un error al cargar la información del producto.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub Label11_Click()
    Foco tPNroMaquina
End Sub

Private Sub Label13_Click()
    Foco tPFCompra
End Sub

Private Sub Label31_Click()
    Foco tMotivo
End Sub

Private Sub Label7_Click()
    Foco tSCodigo
End Sub

Private Sub lDocumento_Click()
    Foco tPFacturaS
End Sub

Private Sub lPTipo_Click()
    If tPArticulo.Enabled Then tPArticulo.SetFocus
End Sub

Private Sub MnuAceptaP_Click()
    Call bTallerAcepta_Click
End Sub

Private Sub MnuClientes_Click()
    Call bCliente_Click
End Sub

Private Sub MnuComentarios_Click()
    Call bComentario_Click
End Sub

Private Sub MnuCumplir_Click()
    Call bCumplir_Click
End Sub

Private Sub MnuEntrega_Click()
    Call bEntrega_Click
End Sub

Private Sub MnuEvAdd_Click(Index As Integer)
    AddEvento MnuEvAdd(Index).Tag, tSCodigo.Tag
    If frmEventos.Visible Then
        frmEventos.s_FillGrid
    End If
End Sub

Private Sub MnuEvShow_Click()
    With frmEventos
        .prmIDServicio = Val(tSCodigo.Tag)
        .s_FillGrid
        .Show , Me
    End With
End Sub

Private Sub MnuForzarCumplido_Click()
    If (miConexion.AccesoAlMenu("Explorador")) Then
        Dim idServicio As String
        idServicio = InputBox("Ingrese el número de servicio", "Forzar cumplido", 0)
        If IsNumeric(idServicio) Then
            'hago esto para complicarla un poco
            If Val(tSCodigo.Tag) = idServicio Then
                If MsgBox("¿Confirma cambiar el estado del servicio " & idServicio & " a cumplido?", vbQuestion + vbYesNo) = vbYes Then
                    Cons = "UPDATE Servicio SET SerEstadoServicio = 5, SerFCumplido = GetDATE() Where Sercodigo = " & idServicio
                    cBase.Execute Cons
                    tSCodigo.Tag = ""
                    Call tSCodigo_KeyPress(13)
                End If
            End If
        End If
    End If
End Sub

Private Sub MnuHistoria_Click()
    Call bHistoria_Click
End Sub

Private Sub MnuImprimir_Click()
    Call bImprimir_Click
End Sub

Private Sub MnuIrValidacion_Click()
    EjecutarApp App.Path & "\validacion de presupuesto.exe", tSCodigo.Text
End Sub

Private Sub MnuOpVisOpe_Click()
    If gCliente <> 0 Then
        EjecutarApp App.Path & "\visualizacion de operaciones.exe", CStr(gCliente)
    End If
End Sub

Private Sub MnuProductos_Click()
    Call bProducto_Click
End Sub

Private Sub MnuRetiro_Click()
    Call bRetiro_Click
End Sub

Private Sub MnuSalir_Click()
    Unload Me
End Sub

Private Sub MnuTaller_Click()
    Call bTaller_Click
End Sub

Private Sub MnuVisita_Click()
    Call bVisita_Click
End Sub

Private Sub Status_PanelClick(ByVal Panel As ComctlLib.Panel)
    If "printer" = Panel.Key Then
        prj_GetPrinter True
        Panel.Text = paINContadoN
    End If
End Sub

Private Sub tbEstado_Click()

    Select Case Val(tbEstado.SelectedItem.Tag)
        Case EstadoS.Taller: picTaller.ZOrder 0
        Case EstadoS.Visita: picVisita.ZOrder 0
        Case EstadoS.Retiro: picRetiro.ZOrder 0
        Case EstadoS.Entrega: picEntrega.ZOrder 0
    End Select
    Me.Refresh

End Sub

Private Sub tMotivo_GotFocus()
    Status.Panels(3).Text = "Motivos de solicitud de servicio. [F1]- Lista de ayuda."
End Sub

Private Sub tMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo errLista
    If KeyCode = vbKeyF1 And Val(tSCodigo.Tag) <> 0 Then
        Dim aIdMotivo As Long
        Screen.MousePointer = 11
        Cons = "Select MSeID, 'Descripción del Motivo' = MSeNombre from MotivoServicio" & _
                " Where MSeTipo = " & Val(lPTipo.Tag) & _
                " Order by MSeNombre"
        Dim objLista As New clsListadeAyuda
        If objLista.ActivarAyuda(cBase, Cons, 5000, 1, "Motivos de Servicio") > 0 Then
            aIdMotivo = objLista.RetornoDatoSeleccionado(0)
        End If
        Set objLista = Nothing
        Me.Refresh
        
        If aIdMotivo <> 0 Then AgregoMotivo aIdMotivo: tMotivo.Text = ""
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errLista:
    clsGeneral.OcurrioError "Ocurrió un error al acceder a la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tMotivo_KeyPress(KeyAscii As Integer)
    
    On Error GoTo errMotivo
    If KeyAscii = vbKeyReturn And Val(tSCodigo.Tag) <> 0 Then
        If Trim(tMotivo.Text) = "" Then vsMotivo.SetFocus: Exit Sub
        
        Screen.MousePointer = 11
        Dim aIdMotivo As Long, aMotivo As String
        aIdMotivo = 0
        
        Cons = "Select MSeID, 'Descripción del Motivo' = MSeNombre from MotivoServicio" & _
                  " Where MSeNombre like '" & Trim(tMotivo.Text) & "%'" & _
                  " And MSeTipo = " & Val(lPTipo.Tag) & _
                  " Order by MSeNombre"
        Dim objLista As New clsListadeAyuda
        If objLista.ActivarAyuda(cBase, Cons, 5000, 1, "Motivos de Servicio") > 0 Then
            aIdMotivo = objLista.RetornoDatoSeleccionado(0)
        End If
        Me.Refresh
        Set objLista = Nothing
        
        If aIdMotivo <> 0 Then AgregoMotivo aIdMotivo: tMotivo.Text = ""
        Screen.MousePointer = 0
    End If
    Exit Sub

errMotivo:
    clsGeneral.OcurrioError "Ocurrió un error al procesar el motivo del servicio.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AgregoMotivo(idMotivo As Long)
    On Error GoTo errAgregar
    
    Screen.MousePointer = 11
    Dim aValor As Long, I As Integer
    
    '1) Valido que el motivo no este ingresado----------------------------------------------------
    With vsMotivo
        For I = 1 To .Rows - 1
            If .Cell(flexcpData, I, 0) = idMotivo Then
                MsgBox "El motivo seleccionado ya está ingresado en la lista.", vbInformation, "Motivo Ingresado"
                Screen.MousePointer = 0: Exit Sub
            End If
        Next
    End With
    '-----------------------------------------------------------------------------------------------------
    
    Cons = "Select * from MotivoServicio Where MSeID = " & idMotivo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        With vsMotivo
            .AddItem Trim(RsAux!MSeNombre)
            aValor = RsAux!MSeID: .Cell(flexcpData, .Rows - 1, 0) = aValor
        End With
    End If
    RsAux.Close
    
    Cons = "Select * from ServicioRenglon Where SReServicio = " & Val(tSCodigo.Tag)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.AddNew
    RsAux!SReServicio = Val(tSCodigo.Tag)
    RsAux!SReMotivo = aValor
    RsAux!SReTipoRenglon = TipoRenglonS.Llamado
    RsAux.Update: RsAux.Close
    
    Screen.MousePointer = 0
    
    Exit Sub
errAgregar:
    clsGeneral.OcurrioError "Ocurrió un error al agregar el motivo de servicio.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub CargoCombos()
    
    cSEstado.AddItem EstadoProducto(EstadoP.Abonado): cSEstado.ItemData(cSEstado.NewIndex) = EstadoP.Abonado
    cSEstado.AddItem EstadoProducto(EstadoP.FueraGarantia): cSEstado.ItemData(cSEstado.NewIndex) = EstadoP.FueraGarantia
    cSEstado.AddItem EstadoProducto(EstadoP.SinCargo): cSEstado.ItemData(cSEstado.NewIndex) = EstadoP.SinCargo
        
End Sub

Private Sub tPArticulo_GotFocus()
    Status.Panels(3).Text = "Presione [F1] para cambiar el tipo del producto."
End Sub

Private Sub tPArticulo_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        If Val(tSCodigo.Tag) = 0 Then Exit Sub
        On Error GoTo errLista
        Screen.MousePointer = 11
        Cons = "Select ArtId, ArtNombre 'Artículo', ArtCodigo 'Código' from Articulo Where ArtTipo = " & Val(lPTipo.Tag) & " Order by ArtNombre"
        Dim miLista As New clsListadeAyuda, aIDSel As Long, aTipoSel As String
        If miLista.ActivarAyuda(cBase, Cons, 5000, 1, "Artículos") > 0 Then
            aIDSel = miLista.RetornoDatoSeleccionado(0)
            aTipoSel = miLista.RetornoDatoSeleccionado(1)
        End If
        Set miLista = Nothing
        If aIDSel <> 0 Then
            tPArticulo.Text = aTipoSel
            tPArticulo.Tag = aIDSel
            ZActualizoCampoProducto CLng(lPIdProducto.Caption), aIDSel, Articulo:=True
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub

errLista:
    clsGeneral.OcurrioError "Ocurrió un error al activar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tPArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tPFCompra
End Sub

Private Sub tPDireccion_GotFocus()
    Status.Panels(3).Text = "Dirección del producto (para realizar servicios)."
End Sub

Private Sub tPFacturaN_Change()
    tPFacturaN.Tag = 1
End Sub

Private Sub tPFacturaN_GotFocus()
    With tPFacturaN: .SelStart = 0: .SelLength = Len(.Text): End With
    Status.Panels(3).Text = "Número de la factura de compra."
End Sub

Private Sub tPFacturaN_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tPFacturaN.Tag) <> 0 And Trim(lPIdProducto.Caption) <> "" Then
            ZActualizoCampoProducto CLng(lPIdProducto.Caption), Trim(tPFacturaN.Text), FacturaN:=True
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
    Status.Panels(3).Text = "Número de serie de la factura de compra."
End Sub

Private Sub tPFacturaS_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        If Val(tPFacturaS.Tag) <> 0 And Trim(lPIdProducto.Caption) <> "" Then
            ZActualizoCampoProducto CLng(lPIdProducto.Caption), Trim(tPFacturaS.Text), FacturaS:=True
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
    Status.Panels(3).Text = "Fecha de compra del producto."
End Sub

Private Sub tPFCompra_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Val(tPFCompra.Tag) <> 0 And Trim(lPIdProducto.Caption) <> "" Then
            If Not IsDate(tPFCompra.Text) Then MsgBox "La fecha ingresada no es correcta. Verifique", vbExclamation, "ATENCIÓN": Exit Sub
            ZActualizoCampoProducto CLng(lPIdProducto.Caption), Trim(tPFCompra.Text), FCompra:=True
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
    Status.Panels(3).Text = "Número de máquina del producto (# serie)."
End Sub

Private Sub tPNroMaquina_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tPNroMaquina.Tag) <> 0 And Trim(lPIdProducto.Caption) <> "" Then
            ZActualizoCampoProducto CLng(lPIdProducto.Caption), Trim(tPNroMaquina.Text), NroMaquina:=True
            tPNroMaquina.Tag = 0
        End If
        Foco tMotivo
    End If
End Sub

Private Sub tSCodigo_Change()
    If Val(tSCodigo.Tag) <> 0 Then LimpioCampos
End Sub

Private Sub tSCodigo_GotFocus()
    Status.Panels(3).Text = "Número de servicio a consultar."
    tSCodigo.SelStart = 0: tSCodigo.SelLength = Len(tSCodigo.Text)
End Sub

Private Sub tSCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyF4: AccionBuscarCliente
    End Select
    
End Sub

Private Sub AccionBuscarCliente()

    On Error GoTo errBuscar
    Screen.MousePointer = 11
    Dim objBuscar As New clsBuscarCliente
    Dim aValor As Long
    objBuscar.ActivoFormularioBuscarClientes cBase, persona:=True
    Me.Refresh
    aValor = objBuscar.BCClienteSeleccionado
    Set objBuscar = Nothing
    
    If aValor <> 0 Then
        Dim aQ As Long: aQ = 0
        Dim aIdServicioS As Long
        Cons = "Select * from Servicio, Producto " & _
                    " Where SerProducto = ProCodigo " & _
                    " And ProCliente = " & aValor & _
                    " And SerEstadoServicio Not In( " & EstadoS.Anulado & ", " & EstadoS.Cumplido & ")"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aQ = 1: aIdServicioS = RsAux!SerCodigo: RsAux.MoveNext
            If Not RsAux.EOF Then aQ = 2
        End If
        RsAux.Close
        
        Select Case aQ
            Case 0: MsgBox "No hay servicios pendientes para el cliente seleccionado.", vbExclamation, "No hay Servicios Pendientes"
            Case 1: CargoDatosServicio aIdServicioS
            Case 2:
                        Cons = "Select SerCodigo, SerCodigo 'Servicio', SerFecha 'Solicitud', ProCodigo 'id_Prod.', ArtNombre 'Producto', ProNroSerie 'Nº Serie', ProCompra 'F/Compra', SerComentario 'Comentarios' " & _
                                " From Servicio, Producto, Articulo " & _
                                " Where SerProducto = ProCodigo and ProArticulo = ArtID" & _
                                " And ProCliente = " & aValor & _
                                " And SerEstadoServicio Not In( " & EstadoS.Anulado & ", " & EstadoS.Cumplido & ")"
                        Dim miLista As New clsListadeAyuda
                        If miLista.ActivarAyuda(cBase, Cons, 5000, 1, "Ayuda") > 0 Then
                            aIdServicioS = miLista.RetornoDatoSeleccionado(0)
                        End If
                        Set miLista = Nothing
                        If aIdServicioS <> 0 Then CargoDatosServicio aIdServicioS
        End Select
        
    End If
    Screen.MousePointer = 0
    Exit Sub

errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al realizar la búsqueda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tSCodigo_KeyPress(KeyAscii As Integer)
Dim bEsCB As Boolean

If KeyAscii = vbKeyReturn And tSCodigo.Text <> "" Then
    
    bEsCB = InStr(1, tSCodigo.Text, "s", vbTextCompare) > 0
    If bEsCB Then tSCodigo.Text = Replace(tSCodigo.Text, "s", "", , , vbTextCompare)
            
    If Not IsNumeric(tSCodigo.Text) Then Exit Sub
    If Val(tSCodigo.Tag) <> 0 Then cSEstado.SetFocus: Exit Sub
    
    If bEsCB Then
        If iUltimoCodigo = Val(tSCodigo.Text) Then
            CargoDatosServicio Val(tSCodigo.Text)
            If Val(tSCodigo.Tag) > 0 Then loc_ProcesoCodBarra
        Else
            iUltimoCodigo = Val(tSCodigo.Text)
            CargoDatosServicio Val(tSCodigo.Text)
        End If
    Else
        iUltimoCodigo = 0
        CargoDatosServicio Val(tSCodigo.Text)
    End If
    
    tSCodigo.SelStart = 0: tSCodigo.SelLength = Len(tSCodigo.Text)
End If
    
End Sub

Private Sub CargoDatosServicio(idServicio As Long)
    
    On Error GoTo errCargar
    Screen.MousePointer = 11
    gProducto = 0: gCliente = 0: cbVisualizacion.Tag = ""
    
    Cons = "Select * from Servicio Left Outer Join Local on SerLocalIngreso = LocCodigo " & _
                " Where SerCodigo = " & idServicio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then
        
        bCliente.Enabled = True: bProducto.Enabled = True: bHistoria.Enabled = True: bComentario.Enabled = True
        bRefresh.Enabled = True: bImprimir.Enabled = True
        
        gCliente = RsAux("SerCliente")
        tSCodigo.Text = idServicio: tSCodigo.Tag = idServicio
        gProducto = RsAux!SerProducto
                
        If Not IsNull(RsAux!SerCliente) Then
            gCliente = RsAux!SerCliente
        Else
            gCliente = 0 'RsAux!SerProducto
        End If
        
        If Not IsNull(RsAux!SerComInterno) Then
            lbComentarioInterno.Caption = RsAux!SerComInterno
        End If
        
        If Not IsNull(RsAux!SerReclamode) Then lReclamo.Caption = RsAux!SerReclamode Else lReclamo.Caption = ""
        If Not IsNull(RsAux("SerCoordinarEntrega")) Then
            If RsAux("SerCoordinarEntrega") Then lblCoordinar.Visible = True
        End If
        lSProceso.Caption = UCase(EstadoServicio(RsAux!SerEstadoServicio))
        lSProceso.Tag = RsAux!SerEstadoServicio
        BuscoCodigoEnCombo cSEstado, RsAux!SerEstadoProducto
                        
        lSFecha.Caption = Format(RsAux!SerFecha, "Ddd d/mm hh:mm")
        lSFModificado.Caption = Format(RsAux!SerModificacion, "dd/mm/yy hh:mm"): lSFModificado.Tag = RsAux!SerModificacion
        lSUsuario.Caption = " " & BuscoUsuario(RsAux!SerUsuario, True)
        
        If Not IsNull(RsAux!SerComentario) Then tSComentario.Text = Trim(RsAux!SerComentario)
        If Not IsNull(RsAux!LocNombre) Then lSLocalIngreso.Caption = " " & Trim(RsAux!LocNombre)
        
        If Not IsNull(RsAux!SerDocumento) Then bFactura.Tag = RsAux!SerDocumento: bFactura.Enabled = True
        
        cSEstado.Tag = 0: tSComentario.Tag = 0
        If frmEventos.Visible Then frmEventos.prmIDServicio = Val(tSCodigo.Tag): frmEventos.s_FillGrid
    Else
        MsgBox "No hay un servicio pendiente con el código: " & idServicio, vbInformation, "ATENCIÓN"
        idServicio = 0
    End If
    RsAux.Close
    
    If idServicio <> 0 Then 'Cargo los motivos de llamado
        Cons = "Select * from ServicioRenglon, MotivoServicio" & _
                   " Where SReServicio = " & idServicio & _
                   " And SReMotivo = MSeID" & _
                   " And SReTipoRenglon = " & TipoRenglonS.Llamado
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        vsMotivo.Rows = 1: Dim aValor As Long
        Do While Not RsAux.EOF
            With vsMotivo
                .AddItem Trim(RsAux!MSeNombre)
                aValor = RsAux!MSeID: .Cell(flexcpData, .Rows - 1, 0) = aValor
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        CargoTabsVisitas idServicio
    End If
    
    If gProducto <> 0 Then CargoDatosProducto gProducto
    If gCliente <> 0 Then CargoDatosCliente gCliente
    
    If idServicio <> 0 Then
        'Ver casos de servicos ya realizados !!!!
        bCumplir.Enabled = True
        If Val(lSProceso.Tag) <> EstadoS.Anulado And Val(lSProceso.Tag) <> EstadoS.Cumplido Then
            If Val(lSProceso.Tag) = EstadoS.Taller Then
                'If Trim(lTFReparado.Caption) = "" Or Trim(tTCosto.Text) = "" Or Trim(lTFAceptado.Caption) = "" Or Trim(lTFPresupuesto.Caption) = "" Then bCumplir.Enabled = False
                'If Trim(vsTaller.Cell(flexcpText, 0, 3)) = "" Or Trim(vsTaller.Cell(flexcpText, 2, 1)) = "" Or Trim(vsTaller.Cell(flexcpText, 2, 3)) = "" Then bCumplir.Enabled = False
                
                '1) Si  es taller y esta reparado dejo entregar y cumplir (04/07)
                'If Trim(vsTaller.Cell(flexcpText, 0, 3)) = "" Then bCumplir.Enabled = False
                bEntrega.Enabled = bCumplir.Enabled
                'al final siempre dejo cumplir
            End If
            
            EstadoControles Estado:=True, ColorFondo:=Colores.Blanco
        ElseIf Val(lSProceso.Tag) = EstadoS.Cumplido Then
'            ColorCamposFichas Visita:=vbInactiveBorder, Retiro:=vbInactiveBorder, Entrega:=vbInactiveBorder, Taller:=vbInactiveBorder
        Else
            ColorCamposFichas Visita:=vbInactiveBorder, Retiro:=vbInactiveBorder, Entrega:=vbInactiveBorder, Taller:=vbInactiveBorder
        End If
    End If
    EstadoMenu
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el servicio.", Err.Description
    EstadoMenu
    Screen.MousePointer = 0
End Sub


Private Sub CargoTabsVisitas(idServicio As Long)

    On Error GoTo errCargar
    Dim bInserteTaller As Boolean
    bInserteTaller = False
    tbEstado.Tabs.Clear
    LimpioCamposFichas
        
    Cons = "Select * from ServicioVisita Where VisServicio = " & idServicio & _
               " Order by VisCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        
        Select Case RsAux!VisTipo
            Case TipoServicio.Visita:               'Visita a Domicilio ------------------------------------------------------------------------------------
                    If Val(picVisita.Tag) = 0 Then
                        tbEstado.Tabs.Add
                        tbEstado.Tabs(tbEstado.Tabs.Count).Caption = "&Visita"
                        tbEstado.Tabs(tbEstado.Tabs.Count).Image = imImagenes.ListImages("visita").Index
                        tbEstado.Tabs(tbEstado.Tabs.Count).Tag = EstadoS.Visita
                    Else
                        LimpioCamposFichas Visita:=True, Retiro:=False, Entrega:=False, Taller:=False
                    End If
                    CargoCamposTVisita
                    
            Case TipoServicio.Retiro:               'Retiro a Domicilio ------------------------------------------------------------------------------------
                    If Val(picRetiro.Tag) = 0 Then
                        tbEstado.Tabs.Add
                        tbEstado.Tabs(tbEstado.Tabs.Count).Caption = "&Retiro"
                        tbEstado.Tabs(tbEstado.Tabs.Count).Image = imImagenes.ListImages("retiro").Index
                        tbEstado.Tabs(tbEstado.Tabs.Count).Tag = EstadoS.Retiro
                    Else
                        LimpioCamposFichas Visita:=False, Retiro:=True, Entrega:=False, Taller:=False
                    End If
                    CargoCamposTRetiro
            
            Case TipoServicio.Entrega:               'Entrega a Domicilio ------------------------------------------------------------------------------------
                    If Val(picEntrega.Tag) = 0 Then
                        'Antes agrego el tab de taller para que quede en orden
                        bInserteTaller = True
                        tbEstado.Tabs.Add
                        tbEstado.Tabs(tbEstado.Tabs.Count).Caption = "&Taller"
                        tbEstado.Tabs(tbEstado.Tabs.Count).Image = imImagenes.ListImages("taller").Index
                        tbEstado.Tabs(tbEstado.Tabs.Count).Tag = EstadoS.Taller
        
                        tbEstado.Tabs.Add
                        tbEstado.Tabs(tbEstado.Tabs.Count).Caption = "&Entrega"
                        tbEstado.Tabs(tbEstado.Tabs.Count).Image = imImagenes.ListImages("entrega").Index
                        tbEstado.Tabs(tbEstado.Tabs.Count).Tag = EstadoS.Entrega
                    Else
                        LimpioCamposFichas Visita:=False, Retiro:=False, Entrega:=True, Taller:=False
                    End If
                    CargoCamposTEntrega
        End Select
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    'Cargo Ficha de Ingreso a Taller -------------------------------------------------------------------------------------------------------------------
    Cons = "Select * from Servicio Left Outer Join Taller On SerCodigo = TalServicio " & _
               " Where SerCodigo = " & idServicio & _
               " And SerLocalReparacion Is Not Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not bInserteTaller Then
            tbEstado.Tabs.Add
            tbEstado.Tabs(tbEstado.Tabs.Count).Caption = "&Taller"
            tbEstado.Tabs(tbEstado.Tabs.Count).Image = imImagenes.ListImages("taller").Index
            tbEstado.Tabs(tbEstado.Tabs.Count).Tag = EstadoS.Taller
        End If
        CargoCamposTTaller
    End If
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo el tab segun el estado del servicio
    If Val(lSProceso.Tag) <> EstadoS.Anulado And Val(lSProceso.Tag) <> EstadoS.Cumplido Then
        For I = 1 To tbEstado.Tabs.Count
            If tbEstado.Tabs(I).Tag = Val(lSProceso.Tag) Then
                tbEstado.Tabs(I).Selected = True
                Select Case Val(lSProceso.Tag)
                    Case EstadoS.Visita:
                                            ColorCamposFichas Visita:=&HC0E0FF, Retiro:=Colores.Gris, Entrega:=Colores.Gris, Taller:=Colores.Gris
                                            TextoAclaracion vsVisita
                    Case EstadoS.Retiro:
                                            ColorCamposFichas Visita:=Colores.Gris, Retiro:=&HC0E0FF, Entrega:=Colores.Gris, Taller:=Colores.Gris
                                            TextoAclaracion vsRetiro
                    Case EstadoS.Taller:
                                            ColorCamposFichas Visita:=Colores.Gris, Retiro:=Colores.Gris, Entrega:=Colores.Gris, Taller:=&HC0E0FF
                                            TextoAclaracion vsTaller
                    Case EstadoS.Entrega:
                                            ColorCamposFichas Visita:=Colores.Gris, Retiro:=Colores.Gris, Entrega:=&HC0E0FF, Taller:=Colores.Gris
                                            TextoAclaracion vsEntrega
                End Select
                Exit For
            End If
        Next
    Else
        If tbEstado.Tabs.Count > 0 Then tbEstado.Tabs(1).Selected = True
    End If
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos de visita.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoCamposTVisita()
    
    Dim aTexto As String
    
    picVisita.Tag = RsAux!VisCodigo     'id de visita
    bVisita.Enabled = True: bTaller.Enabled = True
    
    With vsVisita
        .Cell(flexcpForeColor, 0, 1, .Rows - 1) = Colores.Azul: .Cell(flexcpForeColor, 0, 3, .Rows - 1) = Colores.Azul
        
        'Fila 0 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 0, 0) = "Fecha:"
        aTexto = Format(RsAux!VisFecha, "Ddd d/mm/yy")
        aTexto = aTexto & "   " & Trim(RsAux!VisHorario)
        .Cell(flexcpText, 0, 1) = aTexto
        
        .Cell(flexcpText, 0, 2) = "Impreso:"
        If Not IsNull(RsAux!VisFImpresion) Then aTexto = Format(RsAux!VisFImpresion, "Ddd dd/mm/yy hh:mm") Else aTexto = ""
        .Cell(flexcpText, 0, 3) = aTexto
                
        'Fila 1 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 1, 0) = "Asignado:"
        If Not IsNull(RsAux!VisCamion) Then aTexto = ZBuscoLocal(RsAux!VisCamion)
        If Not IsNull(RsAux!VisLiquidarAlCamion) Then aTexto = aTexto & " (liquidar: " & Format(RsAux!VisLiquidarAlCamion, FormatoMonedaP) & ")"
        .Cell(flexcpText, 1, 1) = aTexto
        
        .Cell(flexcpText, 1, 2) = "Costo:"
        aTexto = ZBuscoMoneda(RsAux!VisMoneda)
        aTexto = aTexto & " " & Format(RsAux!VisCosto, FormatoMonedaP)
        .Cell(flexcpText, 1, 3) = aTexto
        
        'Fila 2 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 2, 0) = "Forma de Pago:"
        .Cell(flexcpText, 2, 1) = TipoFacturaServicio(RsAux!VisFormaPago)
        
        'Fila 4 ------------------------------------------------------------------------------------
        If Not IsNull(RsAux!VisTexto) Then
            Dim rsTxt As rdoResultset
            Cons = "Select * from TextoVisita Where TViCodigo = " & RsAux!VisTexto
            Set rsTxt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsTxt.EOF Then aTexto = Trim(rsTxt!TViTexto) Else aTexto = ""
            rsTxt.Close
            .Cell(flexcpText, 4, 0) = "Sugerencias:"
            .Cell(flexcpText, 4, 1) = aTexto
        End If
        
        'Fila 3 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 3, 0) = "Comentarios:"
        If Not IsNull(RsAux!VisComentario) Then aTexto = Trim(RsAux!VisComentario) Else aTexto = ""
        .Cell(flexcpText, 3, 1) = aTexto
        
        '-----------------------------------------------------------------------------------------------
                
        'Si el estado del servicio conicide cargo descripción
        If Val(lSProceso.Tag) = EstadoS.Visita Then
            If IsNull(RsAux!VisFImpresion) Then aTexto = "A Imprimir" Else aTexto = "En Técnico (en camino a la casa del cliente)"
            If RsAux!VisSinEfecto Then aTexto = "Visita SIN EFECTO"
            .Cell(flexcpText, 5, 0) = aTexto
        End If
        
    End With
    
End Sub

Private Sub CargoCamposTRetiro()

    Dim aTexto As String
    picRetiro.Tag = RsAux!VisCodigo     'id de visita
    bRetiro.Enabled = True: bTaller.Enabled = True
                    
    With vsRetiro
        .Cell(flexcpForeColor, 0, 1, .Rows - 1) = Colores.Azul: .Cell(flexcpForeColor, 0, 3, .Rows - 1) = Colores.Azul
        
        'Fila 0 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 0, 0) = "Fecha:"
        aTexto = Format(RsAux!VisFecha, "Ddd d/mm/yy")
        aTexto = aTexto & "   " & Trim(RsAux!VisHorario)
        .Cell(flexcpText, 0, 1) = aTexto
        
        .Cell(flexcpText, 0, 2) = "Impreso:"
        If Not IsNull(RsAux!VisFImpresion) Then aTexto = Format(RsAux!VisFImpresion, "Ddd dd/mm/yy hh:mm") Else aTexto = ""
        .Cell(flexcpText, 0, 3) = aTexto
        
        'Fila 1 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 1, 0) = "Asignado:"
        If Not IsNull(RsAux!VisCamion) Then aTexto = ZBuscoLocal(RsAux!VisCamion)
        If Not IsNull(RsAux!VisLiquidarAlCamion) Then aTexto = aTexto & " (liquidar: " & Format(RsAux!VisLiquidarAlCamion, FormatoMonedaP) & ")"
        .Cell(flexcpText, 1, 1) = aTexto
        
        .Cell(flexcpText, 1, 2) = "Flete:"
        If Not IsNull(RsAux!VisTipoFlete) Then aTexto = ZBuscoFlete(RsAux!VisTipoFlete)
        .Cell(flexcpText, 1, 3) = aTexto
        
        'Fila 2 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 2, 0) = "Costo:"
        aTexto = ZBuscoMoneda(RsAux!VisMoneda)
        aTexto = aTexto & " " & Format(RsAux!VisCosto, FormatoMonedaP)
        .Cell(flexcpText, 2, 1) = aTexto
        
        .Cell(flexcpText, 2, 2) = "F/ Pago:"
        .Cell(flexcpText, 2, 3) = TipoFacturaServicio(RsAux!VisFormaPago)
        
        'Fila 3 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 3, 0) = "Comentarios:"
        If Not IsNull(RsAux!VisComentario) Then aTexto = Trim(RsAux!VisComentario) Else aTexto = ""
        .Cell(flexcpText, 3, 1) = aTexto
        
        'Si el estado del servicio conicide cargo descripción
        If Val(lSProceso.Tag) = EstadoS.Retiro Then
            If IsNull(RsAux!VisFImpresion) Then aTexto = "A Imprimir" Else aTexto = "En camino de retirarlo."
            If RsAux!VisSinEfecto Then aTexto = "Retiro SIN EFECTO"
            .Cell(flexcpText, 5, 0) = aTexto
        End If
    End With
    
End Sub

Private Sub CargoCamposTEntrega()
    
    Dim aTexto As String
    picEntrega.Tag = RsAux!VisCodigo     'id de visita
    bEntrega.Enabled = True
    
    With vsEntrega
        .Cell(flexcpForeColor, 0, 1, .Rows - 1) = Colores.Azul: .Cell(flexcpForeColor, 0, 3, .Rows - 1) = Colores.Azul
        
        'Fila 0 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 0, 0) = "Fecha:"
        aTexto = Format(RsAux!VisFecha, "Ddd d/mm/yy")
        aTexto = aTexto & "   " & Trim(RsAux!VisHorario)
        .Cell(flexcpText, 0, 1) = aTexto
        
        .Cell(flexcpText, 0, 2) = "Impreso:"
        If Not IsNull(RsAux!VisFImpresion) Then aTexto = Format(RsAux!VisFImpresion, "Ddd dd/mm/yy hh:mm") Else aTexto = ""
        .Cell(flexcpText, 0, 3) = aTexto
        
        'Fila 1 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 1, 0) = "Asignado:"
        If Not IsNull(RsAux!VisCamion) Then aTexto = ZBuscoLocal(RsAux!VisCamion)
        If Not IsNull(RsAux!VisLiquidarAlCamion) Then aTexto = aTexto & " (liquidar: " & Format(RsAux!VisLiquidarAlCamion, FormatoMonedaP) & ")"
        .Cell(flexcpText, 1, 1) = aTexto
        
        .Cell(flexcpText, 1, 2) = "Flete:"
        If Not IsNull(RsAux!VisTipoFlete) Then aTexto = ZBuscoFlete(RsAux!VisTipoFlete)
        .Cell(flexcpText, 1, 3) = aTexto
        
        'Fila 2 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 2, 0) = "Costo:"
        aTexto = ZBuscoMoneda(RsAux!VisMoneda)
        aTexto = aTexto & " " & Format(RsAux!VisCosto, FormatoMonedaP)
        .Cell(flexcpText, 2, 1) = aTexto
        
        .Cell(flexcpText, 2, 2) = "F/ Pago:"
        .Cell(flexcpText, 2, 3) = TipoFacturaServicio(RsAux!VisFormaPago)
        
        'Fila 3 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 3, 0) = "Comentarios:"
        If Not IsNull(RsAux!VisComentario) Then aTexto = Trim(RsAux!VisComentario) Else aTexto = ""
        .Cell(flexcpText, 3, 1) = aTexto
        
        If Val(lSProceso.Tag) = EstadoS.Entrega Then
            If IsNull(RsAux!VisFImpresion) Then aTexto = "A Imprimir" Else aTexto = "En camino hacia la casa del cliente"
            If RsAux!VisSinEfecto Then aTexto = "Entrega SIN EFECTO"
            .Cell(flexcpText, 5, 0) = aTexto
        End If
    End With
   
    
End Sub

Private Function GetPhoneSMSNumber(ByVal idServicio As Long) As String
Dim rsP As rdoResultset
GetPhoneSMSNumber = ""
    Set rsP = cBase.OpenResultset("SELECT MWATelefono FROM MensajeWhatsApp WHERE MWATipo = 2 AND MWADocumento = " & idServicio, rdOpenForwardOnly)
    If Not rsP.EOF Then
        GetPhoneSMSNumber = Trim(rsP(0))
    End If
    rsP.Close
End Function

Private Sub CargoCamposTTaller()
Dim aTexto As String

    picTaller.Tag = RsAux!SerCodigo     'id de visita
    
    Dim sPhoneSMS As String
    sPhoneSMS = GetPhoneSMSNumber(RsAux("SerCodigo"))
    With vsTaller
        .Cell(flexcpForeColor, 0, 1, .Rows - 1) = Colores.Azul: .Cell(flexcpForeColor, 0, 3, .Rows - 1) = Colores.Azul
        
        'Fila 0 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 0, 0) = "Técnico:"
        If Not IsNull(RsAux!TalTecnico) Then aTexto = BuscoUsuario(Codigo:=RsAux!TalTecnico, Identificacion:=True) Else aTexto = ""
        .Cell(flexcpText, 0, 1) = aTexto
        
        .Cell(flexcpText, 0, 2) = "Reparado:"
        If Not IsNull(RsAux!TalFReparado) Then
            aTexto = Format(RsAux!TalFReparado, "Ddd dd/mm/yy hh:mm")
            
            If Val(lSProceso.Tag) = EstadoS.Cumplido Then
                If Not IsNull(RsAux!TalSinArreglo) Then
                    If RsAux!TalSinArreglo Then
                        .Cell(flexcpText, 5, 0) = "Comentario: Se cumplió como 'No reparable'."
                    Else
                        .Cell(flexcpText, 5, 0) = "Comentario: Se cumplió como 'Reparado'."
                    End If
                    TextoAclaracion vsTaller
                End If
            End If
            
        Else
            aTexto = ""
            If Not IsNull(RsAux!TalSinArreglo) Then If RsAux!TalSinArreglo Then aTexto = "No reparable"
        End If
        
        .Cell(flexcpText, 0, 3) = aTexto
            
        'Fila 1 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 1, 0) = "Recepción:"
        If Not IsNull(RsAux!TalFIngresoRecepcion) Then aTexto = Format(RsAux!TalFIngresoRecepcion, "Ddd dd/mm hh:mm") Else aTexto = ""
        .Cell(flexcpText, 1, 1) = aTexto
        
        .Cell(flexcpText, 1, 2) = "A Reparar en:"
        If Not IsNull(RsAux!SerLocalReparacion) Then aTexto = ZBuscoLocal(RsAux!SerLocalReparacion): localReparacion = RsAux!SerLocalReparacion
        .Cell(flexcpText, 1, 3) = aTexto
        
        'Fila 2 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 2, 0) = "Presupuesto:"
        aTexto = ""
        If Not IsNull(RsAux!SerMoneda) Then aTexto = ZBuscoMoneda(RsAux!SerMoneda)
        If Not IsNull(RsAux!SerCostoFinal) Then aTexto = aTexto & " " & Format(RsAux!SerCostoFinal, FormatoMonedaP)
        If Trim(aTexto) <> "" Then
            If Not IsNull(RsAux!TalFPresupuesto) Then aTexto = aTexto & " (" & Format(RsAux!TalFPresupuesto, "Ddd dd/mm") & ")"
        End If
        .Cell(flexcpText, 2, 1) = aTexto
        Dim aCostoFinal As Currency
        If Not IsNull(RsAux!SerCostoFinal) Then aCostoFinal = RsAux!SerCostoFinal Else aCostoFinal = 0
        .Cell(flexcpData, 2, 1) = aCostoFinal
        
        
        .Cell(flexcpText, 2, 2) = "Aceptado:"
         aTexto = ""
        If Not IsNull(RsAux!TalFAceptacion) Then
            aTexto = Format(RsAux!TalFAceptacion, "Ddd dd/mm")
            If Not IsNull(RsAux!TalAceptado) Then
                If Not RsAux!TalAceptado Then
                    aTexto = aTexto & " (NO)"
                    .Cell(flexcpForeColor, 2, 3) = Colores.RojoClaro
                End If
            End If
        End If
        .Cell(flexcpText, 2, 3) = aTexto
        
        'Fila 3 ------------------------------------------------------------------------------------
        .Cell(flexcpText, 3, 0) = "Comentarios:"
        If Not IsNull(RsAux!TalComentario) Then
            aTexto = Trim(RsAux!TalComentario)
            aTexto = f_QuitarClavesDelComentario(aTexto)
        Else
            aTexto = ""
        End If
        .Cell(flexcpText, 3, 1) = aTexto
        
        
        '-----------------------------------------------------------------------------------------------
            
        bTaller.Enabled = True
        If IsNull(RsAux!TalFAceptacion) And Not IsNull(RsAux!SerCostoFinal) Then bTallerAcepta.Enabled = True
        
        'Si el estado del servicio conicide cargo descripción
        If Val(lSProceso.Tag) = EstadoS.Taller Then
            
            Dim aLocalEntrega As Long
            aLocalEntrega = RsAux!SerLocalIngreso
            If Not IsNull(RsAux!TalLocalAlCliente) Then aLocalEntrega = RsAux!TalLocalAlCliente
            
            If Not IsNull(RsAux!TalSinArreglo) And RsAux!TalSinArreglo Then
                If (RsAux("TalAceptado") = 0 And Not IsNull(RsAux("TAlFAceptacion"))) Then
                    aTexto = "Presupuesto rechazado"
                Else
                    aTexto = "No reparable"
                End If
                If Not IsNull(RsAux!TalFSalidaRealizado) Then 'HACIA LOCAL DE ENTREGA
                    
                    If IsNull(RsAux!TalFSalidaRecepcion) Then
                        aTexto = aTexto & " (en camino al local de entrega)"
                    Else
                        aTexto = aTexto & " - Para Entregar en " & ZBuscoLocal(aLocalEntrega)
                    End If
                Else
                    aTexto = aTexto & " (en Taller)"
                End If
            Else
                If Not IsNull(RsAux!TalFReparado) Then          'REPARADO
                    aTexto = "REPARADO"
                    If Not IsNull(RsAux!TalFSalidaRealizado) Then 'HACIA LOCAL DE ENTREGA
                        If IsNull(RsAux!TalFSalidaRecepcion) Then
                            aTexto = aTexto & " (en camino al local de entrega)"
                        Else
                            aTexto = aTexto & " - Para Entregar en " & ZBuscoLocal(aLocalEntrega)
                        End If
                    Else
                        aTexto = aTexto & " (en Taller)"
                    End If
                    
                Else                                                             'NO REPARADO
                    If Not IsNull(RsAux!SerCostoFinal) And Not IsNull(RsAux!TalFPresupuesto) Then
                        aTexto = "PRESUPUESTO"
                        If IsNull(RsAux!TalFAceptacion) Then
                            aTexto = aTexto & " (a aceptar)"
                        Else
                            If RsAux!TalAceptado Then aTexto = aTexto & "(aceptado/a reparar)" Else aTexto = aTexto & " (rechazado)"
                        End If
                        'Especifico donde está !!!!
                        If RsAux!SerLocalIngreso = RsAux!SerLocalReparacion Then
                            aTexto = aTexto & " En " & Trim(lSLocalIngreso.Caption)
                        Else
                            If Not IsNull(RsAux!TalFSalidaRealizado) Then 'HACIA LOCAL DE ENTREGA
                                If IsNull(RsAux!TalFSalidaRecepcion) Then
                                    aTexto = aTexto & " En camino a " & ZBuscoLocal(aLocalEntrega)
                                Else
                                    aTexto = aTexto & " - En " & ZBuscoLocal(aLocalEntrega)
                                End If
                            Else
                                If Not IsNull(RsAux!SerLocalReparacion) Then
                                    aTexto = aTexto & " (en " & ZBuscoLocal(RsAux!SerLocalReparacion) & ")"
                                Else
                                    aTexto = aTexto & " (en Taller)"
                                End If
                            End If
                        End If
                        
                    Else
                        
                        'Puede ser que este en traslado hacia el taller
                        aTexto = "SIN REPARAR" 'aTexto = "A REPARAR"
                        If RsAux!SerLocalIngreso <> RsAux!SerLocalReparacion Then
                            If IsNull(RsAux!TalFIngresoRecepcion) Then
                                If Not IsNull(RsAux!TalFIngresoRealizado) Then aTexto = aTexto & " (en camino al local de reparación)" Else aTexto = aTexto & " en " & Trim(lSLocalIngreso.Caption) & " (a trasladar 'DE IDA')"
                            Else
                                If Not IsNull(RsAux!TalFSalidaRealizado) Then 'HACIA LOCAL DE ENTREGA
                                    If IsNull(RsAux!TalFSalidaRecepcion) Then
                                        aTexto = aTexto & " (en camino al local de entrega)"
                                    Else
                                        'En el local de entrega
                                        aTexto = aTexto & " - Para Entregar en " & ZBuscoLocal(aLocalEntrega)
                                    End If
                                Else
                                    aTexto = aTexto & " (en taller/pte. de presupuesto)"
                                End If
                            End If
                        Else
                            'aTexto = aTexto & " en " & Trim(lSLocalIngreso.Caption)
                            aTexto = "Fase de diagnóstico"
                            If Not IsNull(RsAux("TalFPresupuesto")) And IsNull(RsAux("SerCostoFinal")) Then
                                aTexto = "Presupuesto a validar"
                            End If
                        End If
                    End If
                 End If
            End If
            .Cell(flexcpText, 5, 0) = aTexto
            
        End If
        
        If sPhoneSMS <> "" Then
            .Cell(flexcpText, 6, 0) = "SMS"
            .Cell(flexcpText, 6, 1) = sPhoneSMS
        End If
        
        .AutoSize 1
    
    End With
    
    bTraslado.Enabled = False
    If Val(lSProceso.Tag) = EstadoS.Taller And gCliente <> paClienteEmpresa Then
        If Not IsNull(RsAux!TalFSalidaRealizado) Then bTraslado.Enabled = True
    End If
        
End Sub

Private Sub LimpioCamposFichas(Optional Visita As Boolean = True, Optional Retiro As Boolean = True, Optional Entrega As Boolean = True, Optional Taller As Boolean = True)

    If Visita Then
        picVisita.Tag = 0
        vsVisita.Cell(flexcpText, 0, 0, vsVisita.Rows - 1, vsVisita.Cols - 1) = ""
        'vsVisita.Cell(flexcpBackColor, 0, 0, vsVisita.Rows - 1, vsVisita.Cols - 1) = vbInactiveBorder
    End If
    
    If Taller Then
        picTaller.Tag = 0
        vsTaller.Cell(flexcpText, 0, 0, vsTaller.Rows - 1, vsTaller.Cols - 1) = ""
        'vsTaller.Cell(flexcpBackColor, 0, 0, vsTaller.Rows - 1, vsTaller.Cols - 1) = Colores.Blanco
    End If
    
    If Retiro Then
        picRetiro.Tag = 0
        vsRetiro.Cell(flexcpText, 0, 0, vsRetiro.Rows - 1, vsRetiro.Cols - 1) = ""
        'vsRetiro.Cell(flexcpBackColor, 0, 0, vsRetiro.Rows - 1, vsRetiro.Cols - 1) = Colores.Blanco
    End If
    
    If Entrega Then
        picEntrega.Tag = 0
        vsEntrega.Cell(flexcpText, 0, 0, vsEntrega.Rows - 1, vsEntrega.Cols - 1) = ""
        'vsEntrega.Cell(flexcpBackColor, 0, 0, vsEntrega.Rows - 1, vsEntrega.Cols - 1) = Colores.Blanco
    End If
    
End Sub

Private Sub tSComentario_Change()
    tSComentario.Tag = 1
End Sub

Private Sub tSComentario_GotFocus()
    Status.Panels(3).Text = "Comentarios de solicitud de servicio."
End Sub

Private Sub tSComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tSComentario.Tag) <> 0 Then
            ZActualizoCampoServicio Val(tSCodigo.Tag), Trim(tSComentario.Text), Comentario:=True
            tSComentario.Tag = 0
        End If
        If tPArticulo.Enabled Then tPArticulo.SetFocus
    End If
End Sub

Private Sub ZActualizoCampoServicio(idServicio As Long, Valor As Variant, Optional Comentario As Boolean = False, Optional Estado As Boolean)
    
    On Error GoTo errActualizar
    If idServicio = 0 Then Exit Sub
    Screen.MousePointer = 11
    Cons = "Select * from Servicio Where SerCodigo = " & idServicio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Edit
        If Comentario Then If Trim(Valor) = "" Then RsAux!SerComentario = Null Else RsAux!SerComentario = Trim(Valor)
        If Estado Then If Valor <> 0 Then RsAux!SerEstadoProducto = Valor
        RsAux.Update
    End If
    RsAux.Close
    Screen.MousePointer = 0
    
    Exit Sub
errActualizar:
    clsGeneral.OcurrioError "Ocurrió un error al actualizar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub ZActualizoCampoProducto(idProducto As Long, Valor As Variant, _
                                Optional FCompra As Boolean = False, Optional FacturaS As Boolean = False, Optional FacturaN As Boolean = False, _
                                Optional NroMaquina As Boolean = False, Optional Articulo As Boolean = False)
    
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
        If Articulo Then If Valor <> 0 Then RsAux!ProArticulo = CLng(Valor)
        RsAux.Update
    End If
    RsAux.Close
    
    If Articulo Then lPGarantia.Caption = " " & RetornoGarantia(tPArticulo.Tag)
    If FCompra Or Articulo Then
        lPEstado.Tag = CalculoEstadoProducto(gProducto)
        lPEstado.Caption = " " & EstadoProducto(Val(lPEstado.Tag))
    End If
    Screen.MousePointer = 0
    
    Exit Sub
errActualizar:
    clsGeneral.OcurrioError "Ocurrió un error al actualizar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tSDireccion_LostFocus()
    tSDireccion.SelStart = 0
End Sub

Private Sub tSTelefono_LostFocus()
    tSTelefono.SelStart = 0
End Sub

Private Sub vsMotivo_GotFocus()
    Status.Panels(3).Text = "Lista de motivos de solicitud. [Del]- Eliminar renglón"
End Sub

Private Sub vsMotivo_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyDelete
            If vsMotivo.Rows = 1 Then Exit Sub
            If vsMotivo.Rows = 2 Then MsgBox "No se pueden eliminar todos los motivos de ingreso.", vbExclamation, "ATENCIÓN": Exit Sub
            
            With vsMotivo
                If MsgBox("Confirma eliminar el motivo " & Trim(.Cell(flexcpText, .Row, 0)), vbQuestion + vbYesNo, "Eliminar Motivo") = vbNo Then Exit Sub
                Screen.MousePointer = 11
                On Error GoTo errEliminar
                Cons = "Delete ServicioRenglon Where SReServicio = " & Val(tSCodigo.Tag) & _
                           " And SReMotivo = " & .Cell(flexcpData, .Row, 0) & _
                           " And SReTipoRenglon = " & TipoRenglonS.Llamado
                cBase.Execute Cons
                .RemoveItem .Row
            End With
    End Select
    
    Screen.MousePointer = 0
    Exit Sub
    
errEliminar:
    clsGeneral.OcurrioError "Ocurrió un error al eliminar el motivo.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub EstadoControles(Estado As Boolean, ColorFondo As Long)
    
    'On Error Resume Next
    cSEstado.Enabled = Estado: cSEstado.BackColor = ColorFondo
    tSComentario.Enabled = Estado: tSComentario.BackColor = ColorFondo
    
    tPArticulo.Enabled = Estado: tPArticulo.BackColor = ColorFondo
    cbVisualizacion.Enabled = Estado
    tPFCompra.Enabled = Estado: tPFCompra.BackColor = ColorFondo
    tPFacturaS.Enabled = Estado: tPFacturaS.BackColor = ColorFondo
    tPFacturaN.Enabled = Estado: tPFacturaN.BackColor = ColorFondo
    
    If Estado = True And Val(lDocumento.Tag) <> 0 Then
        tPArticulo.Enabled = False: tPArticulo.BackColor = Colores.Inactivo
        tPFCompra.Enabled = False: tPFCompra.BackColor = Colores.Inactivo
        tPFacturaS.Enabled = False: tPFacturaS.BackColor = Colores.Inactivo
        tPFacturaN.Enabled = False: tPFacturaN.BackColor = Colores.Inactivo
    End If
    tPNroMaquina.Enabled = Estado: tPNroMaquina.BackColor = ColorFondo
    tPDireccion.BackColor = ColorFondo
    
    tMotivo.Enabled = Estado: tMotivo.BackColor = ColorFondo
    vsMotivo.Enabled = Estado: vsMotivo.BackColor = ColorFondo
    
End Sub

Private Sub ColorCamposFichas(Visita As Long, Retiro As Long, Entrega As Long, Taller As Long)

    vsVisita.Cell(flexcpBackColor, 0, 0, vsVisita.Rows - 1, vsVisita.Cols - 1) = Visita
    vsTaller.Cell(flexcpBackColor, 0, 0, vsTaller.Rows - 1, vsTaller.Cols - 1) = Taller
    vsRetiro.Cell(flexcpBackColor, 0, 0, vsRetiro.Rows - 1, vsRetiro.Cols - 1) = Retiro
    vsEntrega.Cell(flexcpBackColor, 0, 0, vsEntrega.Rows - 1, vsEntrega.Cols - 1) = Entrega
    
End Sub

Private Function TextoAclaracion(Grilla As vsFlexGrid)

    Grilla.Cell(flexcpBackColor, 5, 0, , Grilla.Cols - 1) = Colores.Rojo
    Grilla.Cell(flexcpForeColor, 5, 0, , Grilla.Cols - 1) = Colores.Blanco
    Grilla.Cell(flexcpFontBold, 5, 0, , Grilla.Cols - 1) = True
    Grilla.MergeCells = flexMergeSpill
    
End Function

Private Function EstadoMenu()
    
    On Error Resume Next
    MnuClientes.Enabled = bCliente.Enabled
    MnuProductos.Enabled = bProducto.Enabled
    MnuComentarios.Enabled = bComentario.Enabled
    
    MnuVisita.Enabled = bVisita.Enabled
    MnuRetiro.Enabled = bRetiro.Enabled
    MnuTaller.Enabled = bTaller.Enabled
    MnuEntrega.Enabled = bEntrega.Enabled
    
    MnuCumplir.Enabled = bCumplir.Enabled
    MnuAceptaP.Enabled = bTallerAcepta.Enabled
    MnuHistoria.Enabled = bHistoria.Enabled
    
    MnuEvAddMenu.Enabled = Val(tSCodigo.Tag) > 0
    
End Function

Private Function ServicioModificado(idServicio As Long) As Boolean

    On Error GoTo errControl
    ServicioModificado = False
    Cons = "Select * from Servicio Where SerCodigo = " & idServicio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Format(RsAux!SerModificacion, "dd/mm/yyyy hh:mm:ss") <> Format(lSFModificado.Tag, "dd/mm/yyyy hh:mm:ss") Then ServicioModificado = True
    End If
    RsAux.Close
    
errControl:
End Function

Private Function ZBuscoLocal(id As Long) As String
    On Error Resume Next
    Dim rsZ As rdoResultset
    
    ZBuscoLocal = ""
    
    Cons = "Select * from Local Where LocCodigo = " & id
    Set rsZ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsZ.EOF Then ZBuscoLocal = Trim(rsZ!LocNombre)
    rsZ.Close

End Function
Private Function ZBuscoMoneda(id As Long) As String
    On Error Resume Next
    Dim rsZ As rdoResultset
    
    ZBuscoMoneda = ""
    
    Cons = "Select * from Moneda Where MonCodigo = " & id
    Set rsZ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsZ.EOF Then ZBuscoMoneda = Trim(rsZ!MonSigno)
    rsZ.Close

End Function

Private Function ZBuscoFlete(id As Long) As String
    On Error Resume Next
    Dim rsZ As rdoResultset
    
    ZBuscoFlete = ""
    
    Cons = "Select * from TipoFlete Where TFlCodigo = " & id
    Set rsZ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsZ.EOF Then ZBuscoFlete = Trim(rsZ!TFlNombreCorto)
    rsZ.Close

End Function

Private Sub ImprimirFichas()
On Error GoTo errFD
Dim iCont As Integer
Dim oPrint As clsPrintReport
    Dim sPaso As String
    sPaso = "1"
    Set oPrint = New clsPrintReport
    'MsgBox "Impresora " & paPrintConfD & " Bandeja " & paPrintConfB & " y Papel: " & paPrintConfPaperSize
    With oPrint
        .StringConnect = miConexion.TextoConexion("Comercio")
        sPaso = "2"
        .DondeImprimo.Bandeja = paPrintConfB
        .DondeImprimo.Impresora = paPrintConfD
        .DondeImprimo.Papel = paPrintConfPaperSize
        .PathReportes = gPathListados
    End With
    
    On Error GoTo errFD
    Dim aTexto As String
    aTexto = ""
    For I = 1 To vsMotivo.Rows - 1
        If aTexto = "" Then aTexto = Trim(vsMotivo.Cell(flexcpText, I, 0)) Else aTexto = aTexto & ", " & Trim(vsMotivo.Cell(flexcpText, I, 0))
    Next I
    
    Dim sQueryServicio As String
    sQueryServicio = "SELECT SerCodigo infoCodigoServicio, '*S'+ RTRIM(convert(varchar(20), SerCodigo)) + '*' infoCodigoBarras " & _
            ", dbo.FormatCIRuc(CliCIRUC) + ' ' + CASE WHEN CliTipo = 1 THEN rtrim(CPeNombre1) + ' ' + RTRIM(CPeApellido1) ELSE RTRIM(CEmNombre) END infoCliente " & _
            ", dbo.FormatDate(SerFecha, 121) infoFecha " & _
            ",'(' + rtrim(Convert(varchar(10), ProCodigo)) + ') ' + RTRIM(ArtNombre) infoArticulo " & _
            ", RTRIM(usuidentificacion) infoRecibio " & _
            ", dbo.TelefonosCliente(CliCodigo) infoTelefonos, '" & aTexto & "' infoMotivos, IsNull(SerComentario, '') infoMemoIngreso " & _
            ", dbo.ArmoDireccion(CliDireccion) infoDireccion, IsNull(IsNull(ProFacturaS, '') + ' ' + CONVERT(varchar(6), ProFacturaN), '') infoFactura " & _
            ", Rtrim(SucI.SucAbreviacion) infoLocal, Rtrim(IsNull(SucS.SucAbreviacion, '')) infoLocalRepara, RTRIM(ProNroSerie) infoNroSerie, ISNull(dbo.FormatDate(ProCompra, 2), '') infoFCompra " & _
            "FROM Servicio INNER JOIN Cliente ON SerCliente = CliCodigo " & _
            "LEFT OUTER JOIN CPersona ON CliCodigo = CPeCliente " & _
            "LEFT OUTER JOIN CEmpresa ON CliCodigo = CEmCliente " & _
            "INNER JOIN Producto ON SerProducto = ProCodigo " & _
            "INNER JOIN Articulo ON ProArticulo = ArtId " & _
            "INNER JOIN Sucursal SucI ON SerLocalIngreso = SucI.SucCodigo " & _
            "LEFT OUTER JOIN Sucursal SucS ON SerLocalReparacion = SucS.SucCodigo " & _
            "INNER JOIN Usuario ON SerUsuario = UsuCodigo " & _
            "WHERE SerCodigo = " & tSCodigo.Text
    sPaso = "3"
    oPrint.Imprimir_vsReport "FichaServicio.xml", "FichaDeServicio", sQueryServicio, "", ""
    Exit Sub
    
errFD:
    clsGeneral.OcurrioError "Error al imprimir las fichas.", Err.Description, "Fichas de devolución " & sPaso
End Sub


Private Sub ImprimoCopia()
Dim iPY As Single
Dim aTexto As String
    ImprimirFichas
    Exit Sub

    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    SeteoImpresoraPorDefecto paPrintConfD
    
    With vsFicha
        On Error Resume Next
        .paperSize = paPrintConfPaperSize
        .Device = paPrintConfD
        .Orientation = orLandscape
        .PaperBin = paPrintConfB
        
        On Error GoTo errImprimir
        
        .DrawPicture .LoadPicture("C:\Desarrollo\Visual Basic\Aplicaciones\logo.gif"), 200, 200, 8000, 1300
        
        .MarginTop = 700
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
        
        .FileName = "Copia de Servicio"
        .FontSize = 8.25
        .TableBorder = tbNone
        
        .FontSize = 9
        .TextAlign = taRightBaseline: .FontBold = True
        .AddTable ">2000|<1800", "Servicio:|" & tSCodigo.Text, ""
        
        .TextAlign = taLeftTop
        iPY = .CurrentY
        .CurrentY = .CurrentY
        
        .FontName = "3 of 9 Barcode"
        .FontSize = 32
        
        .CurrentX = 8000
        .Paragraph = "*S" & tSCodigo.Text & "*"
        
        .FontName = "Tahoma"
        .FontSize = 8
        
        .CurrentY = iPY
        
        .FontBold = False
        
        .Paragraph = "": .Paragraph = "": .Paragraph = "": .Paragraph = ""
        .AddTable "<900|<1800|>1400|<2000", "Fecha:|" & Format(gFechaServidor, "d-Mmm yyyy hh:mm") & "|Recibido por:|" & lSUsuario.Caption, ""
        
        .Paragraph = ""
        .AddTable "<900|<9000", "Cliente:|" & Trim(lSCliente.Caption), ""
        .AddTable "<900|<9000", "Teléfono:|" & Trim(tSTelefono.Text), ""
        .AddTable "<950|<9000", "Dirección:|" & Trim(tSDireccion.Text), ""
        
        .Paragraph = ""
        .FontBold = True
        aTexto = "(" & Trim(lPIdProducto.Caption) & ") " & Trim(tPArticulo.Text)
        .AddTable "<950|<8000", "Artículo:|" & aTexto, ""
        .FontBold = False
        
        .AddTable "<900|<1500|<1500|<1100|<1200|<2400|<800|<500", _
                        "Factura:|" & Trim(tPFacturaS.Text) & " " & tPFacturaN.Text & _
                        "|Fecha Compra:|" & Trim(tPFCompra.Text) & _
                        "|Nro. Serie:|" & Trim(tPNroMaquina.Text) & _
                        "|Estado:|" & Trim(lPEstado.Caption), ""
              
        .Paragraph = ""
        .AddTable "<900|3000", "Local:|" & Trim(lSLocalIngreso.Caption), ""
        
        aTexto = ""
        For I = 1 To vsMotivo.Rows - 1
            If aTexto = "" Then aTexto = Trim(vsMotivo.Cell(flexcpText, I, 0)) Else aTexto = aTexto & ", " & Trim(vsMotivo.Cell(flexcpText, I, 0))
        Next I
        .AddTable "<900|<10100", "Motivos:|" & aTexto, ""
        
        If Trim(tSComentario.Text) <> "" Then .AddTable "<1050|<10000", "Aclaración:|" & Trim(tSComentario.Text), ""
                
        'Datos de Taller----------------------------------------------------------------------------------------------------------------------
        If vsTaller.Cell(flexcpText, 0, 0) <> "" Then
            Dim aTxt As String: aTxt = ""
            Cons = "Select * From ServicioRenglon, Articulo" & _
                    " Where SReServicio = " & Val(tSCodigo.Text) & _
                    " And SReTipoRenglon = " & TipoRenglonS.Cumplido & _
                    " And SReMotivo = ArtID"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
            Do While Not RsAux.EOF
                aTxt = aTxt & RsAux!SReCantidad & " " & Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre) & ", "
                RsAux.MoveNext
            Loop
            RsAux.Close
            If Len(aTxt) > 1 Then aTxt = Mid(aTxt, 1, Len(aTxt) - 2)
            .AddTable "<1200|<3000|<950|<2000|<950|<2200", "Presupuesto:|" & vsTaller.Cell(flexcpText, 2, 1) & "|Aceptado:|" & vsTaller.Cell(flexcpText, 2, 3) & "|Reparado:|" & vsTaller.Cell(flexcpText, 0, 3), ""
            .AddTable "<1150|<10000", "Repuestos:|" & Trim(aTxt), ""
            .AddTable "<1150|<10000", "Comentario:|" & vsTaller.Cell(flexcpText, 3, 1), ""
        End If
        '--------------------------------------------------------------------------------------------------------------------------------------
        
        .Paragraph = ""
        .FontSize = 7
        aTexto = "1) - Para retirar el aparato es indispensable presentar esta boleta. -"
        .AddTable "900|10100", "Nota:|" & aTexto, ""
        aTexto = "2) - El plazo de retiro del aparato es de 90 días contados a partir de la fecha de esta boleta. Expirado dicho plazo se perderá todo derecho a reclamo " _
            & "sobre el mismo. -"
        .AddTable "900|10100", "|" & aTexto, ""
        .EndDoc
        
        
'        .Device = paPrintConfD
'        .PaperBin = paPrintConfB
'        .paperSize = paPrintConfPaperSize
        
        .PrintDoc   'Cliente
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub

Public Sub loc_FindComentarios(idCliente As Long)
Dim RsCom As rdoResultset
Dim bHay As Boolean

    On Error GoTo errMenu
    Screen.MousePointer = 11
    bHay = False
    
    Cons = "Select * From Comentario " _
            & " Where ComCliente = " & idCliente _
            & " And ComTipo IN (" & prmTipoComentario & ")"
            
    Set RsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsCom.EOF Then bHay = True
    RsCom.Close
    If Not bHay Then Screen.MousePointer = 0: Exit Sub
    
    Dim objC As New clsCliente
    objC.Comentarios idCliente:=idCliente
    Set objC = Nothing
    Me.Refresh
    Screen.MousePointer = 0
    Exit Sub
    
errMenu:
    clsGeneral.OcurrioError "Ocurrió un error al acceder al fomulario de comentarios.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub s_LoadMenuEventos()
Dim rsE As rdoResultset
Dim iQ As Integer
'    Sólo la descripción (en caso de q en el comentario lo ocultes), o la descripción y entre parentesis la clave en caso de q en el comentario NO lo ocultes.
    
    If MnuEvAdd(0).Caption <> "" Then Exit Sub
    iQ = 0
    Cons = "Select rTrim(ESeClave), rtrim(ESeDescripcion) From EventosServicio Order By ESeOrden"
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsE.EOF
        If iQ > 0 Then
            Load MnuEvAdd(iQ)
        End If
        With MnuEvAdd(iQ)
            .Visible = True
            .Enabled = True
            .Caption = rsE(1)
            .Tag = rsE(0)
        End With
        iQ = iQ + 1
        rsE.MoveNext
    Loop
    rsE.Close
End Sub

Private Sub loc_validarUsuario()
Dim sRes As String
Dim idUsuario As Long
Dim id As Integer

    On Error GoTo errGrabar
    
    sRes = InputBox("Ingrese su digito de usuario", "Seguimiento de servicios")
    If IsNumeric(sRes) Then
        idUsuario = sRes
    Else
        Exit Sub
    End If
    
    If BuscoUsuarioDigito(idUsuario, Codigo:=True) = 0 Then
        MsgBox ("El usuario no existe"), vbInformation, "ATENCION"
        Exit Sub
    Else
        idUsuario = BuscoUsuarioDigito(idUsuario, Codigo:=True)
    End If
    
    Screen.MousePointer = 11
    FechaDelServidor
    'Tabla Servicio-----------------------------------------------------------------------------------------------------
    Cons = "Select * from Servicio Where SerCodigo = " & Val(tSCodigo.Text)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    
    If IsNull(RsAux!SerComentarioR) Then RsAux!SerComentarioR = "Retira el cliente."
    RsAux!SerFCumplido = Format(gFechaServidor, sqlFormatoF)
    RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
    RsAux!SerUsuario = idUsuario
    RsAux!SerEstadoServicio = EstadoS.Cumplido
    RsAux.Update: RsAux.Close
    
    If ServicioModificado(Val(tSCodigo.Text)) Then
        LimpioCampos
        CargoDatosServicio Val(tSCodigo.Text)
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errGrabar:
    clsGeneral.OcurrioError "Ocurrió un error al grabar los datos del cumplido.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub VueltaAtrasCumplidoCompañia()
        FechaDelServidor
        
        paCodigoDeUsuario = miConexion.UsuarioLogueado(Codigo:=True)
        If paCodigoDeUsuario = 0 Then Exit Sub
        
        Screen.MousePointer = 11
        
        cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
        
    
        Dim idProducto As Long
        'Tabla Servicio-----------------------------------------------------------------------------------------------------
        Cons = "Select * from Servicio Where SerCodigo = " & Val(tSCodigo.Text)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        idProducto = RsAux("SerProducto")
        RsAux.Edit
        RsAux!SerFCumplido = Null 'Format(gFechaServidor, sqlFormatoF)
        RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsAux!SerUsuario = paCodigoDeUsuario
        RsAux!SerEstadoServicio = EstadoS.Taller
        RsAux.Update
        RsAux.Close

        'Tengo que hacer el traslado al local.
        Dim IDArticulo As Long
        Cons = "SELECT ProArticulo FROM Producto WHERE ProCodigo = " & idProducto
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        IDArticulo = RsAux(0)
        RsAux.Close
        
        HagoTraslado IDArticulo, Val(tSCodigo.Text), paCodigoDeSucursal
        
        'Tengo que cambiar el estado a recuperar.
        MarcoMovimientoStockFisico paCodigoDeUsuario, 2, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, 1, TipoDocumento.ServicioCambioEstado, Val(tSCodigo.Text)
        MarcoMovimientoStockFisico paCodigoDeUsuario, 2, paCodigoDeSucursal, IDArticulo, 1, paEstadoArticuloEntrega, -1, TipoDocumento.ServicioCambioEstado, Val(tSCodigo.Text)
            
        MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, paEstadoARecuperar, 1, 1
        MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, 1, -1
        
        MarcoMovimientoStockFisicoEnLocal 2, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, 1
        MarcoMovimientoStockFisicoEnLocal 2, paCodigoDeSucursal, IDArticulo, 1, paEstadoArticuloEntrega, -1
            
        cBase.CommitTrans

End Sub
