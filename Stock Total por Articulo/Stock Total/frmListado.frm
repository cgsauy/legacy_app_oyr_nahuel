VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{B443E3A5-0B4D-4B43-B11D-47B68DC130D7}#1.6#0"; "orArticulo.ocx"
Begin VB.Form frmListado 
   Caption         =   "Consulta de Stock Total por Artículo"
   ClientHeight    =   6690
   ClientLeft      =   1905
   ClientTop       =   2760
   ClientWidth     =   9990
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   9990
   Begin VB.PictureBox picEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   1800
      ScaleHeight     =   5025
      ScaleWidth      =   8025
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   8055
      Begin VSFlex6DAOCtl.vsFlexGrid vsAcciones 
         Height          =   1215
         Left            =   240
         TabIndex        =   31
         Top             =   3240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2143
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
         BackColorFixed  =   1392133
         ForeColorFixed  =   16777215
         BackColorSel    =   5994576
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   15790320
         GridColor       =   -2147483633
         GridColorFixed  =   4482359
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   285
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
      Begin VB.TextBox txtQMaxXMayor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   20
         Text            =   "888"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtStockMinXMayor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   22
         Text            =   "888"
         Top             =   1560
         Width           =   615
      End
      Begin VB.ComboBox cboAcciones 
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2040
         TabIndex        =   24
         Text            =   "Combo1"
         Top             =   2040
         Width           =   5775
      End
      Begin VB.TextBox tbDemora 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "888"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox tbDisponibleDesde 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton butCancel 
         Caption         =   "&Cancelar"
         Height          =   315
         Left            =   6840
         TabIndex        =   27
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton butModificar 
         Caption         =   "&Grabar"
         Height          =   315
         Left            =   6840
         TabIndex        =   25
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox tSMin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   14
         Text            =   "888"
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Si es 0 no se cierra la venta automáticamente"
         ForeColor       =   &H00A0A0A0&
         Height          =   375
         Left            =   2880
         TabIndex        =   34
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label lblInfoAccion 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Al llegar a esta cantidad se cerrará la venta por mayor"
         ForeColor       =   &H00A0A0A0&
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   2400
         Width           =   7695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Últimas acciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000576C4&
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblInfoQMinXMayor 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sí es 0 no se vende a distribuidor, 1 es ilimitado y > 1 es hasta esa cantidad"
         ForeColor       =   &H00A0A0A0&
         Height          =   375
         Left            =   2880
         TabIndex        =   30
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label lblInfoStckMinXMayor 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Al llegar a esta cantidad se cerrará la venta por mayor (se edita si tiene ingresado un número)."
         ForeColor       =   &H00A0A0A0&
         Height          =   375
         Left            =   2880
         TabIndex        =   29
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Limitar venta:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock mínimo x mayor:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Acción de stock"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "D&emora entrega:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Disponible desde:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Stock mínimo:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox pnlCampos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1815
      ScaleWidth      =   3855
      TabIndex        =   35
      Top             =   1560
      Width           =   3855
      Begin VSFlex6DAOCtl.vsFlexGrid vsCampos 
         Height          =   1815
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   3201
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
         BackColorFixed  =   1392133
         ForeColorFixed  =   16777215
         BackColorSel    =   5994576
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   7570281
         GridColor       =   -2147483633
         GridColorFixed  =   4482359
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   285
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
   Begin VB.PictureBox picHorizontal 
      BackColor       =   &H8000000D&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   2655
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLocal 
      Height          =   4335
      Left            =   5520
      TabIndex        =   3
      Top             =   1560
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7646
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   1
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7646
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   1
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
      TabIndex        =   6
      Top             =   1920
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
      Zoom            =   70
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   6435
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2805
            MinWidth        =   2805
            Key             =   "bd"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14261
            Key             =   "msg"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      BorderStyle     =   0  'None
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
      Height          =   1005
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   10215
      Begin prjFindArticulo.orArticulo tArticulo 
         Height          =   285
         Left            =   600
         TabIndex        =   1
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         FindNombreEnUso =   0   'False
      End
      Begin VB.Image imgDown 
         Height          =   270
         Left            =   3960
         MouseIcon       =   "frmListado.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "frmListado.frx":0614
         ToolTipText     =   "Ver valores de campos"
         Top             =   655
         Width           =   270
      End
      Begin VB.Label lblVtasXDias 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0FFFF&
         ForeColor       =   &H00600000&
         Height          =   285
         Left            =   7080
         MouseIcon       =   "frmListado.frx":0A64
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   655
         Width           =   2295
      End
      Begin VB.Label lbStock 
         BackColor       =   &H00F0FFFF&
         Caption         =   "En uso y en venta"
         ForeColor       =   &H00060000&
         Height          =   285
         Left            =   120
         MouseIcon       =   "frmListado.frx":0D6E
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   655
         Width           =   4695
      End
      Begin VB.Label lbArregloStock 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00DFFFFF&
         ForeColor       =   &H00600000&
         Height          =   285
         Left            =   3510
         MouseIcon       =   "frmListado.frx":1078
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   360
         Width           =   5895
      End
      Begin VB.Label lbPrecioCtdo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00666666&
         Height          =   195
         Left            =   5040
         MouseIcon       =   "frmListado.frx":1382
         MousePointer    =   99  'Custom
         TabIndex        =   10
         ToolTipText     =   "Contado vigente"
         Top             =   0
         Width           =   45
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Art:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   0
         Width           =   495
      End
      Begin VB.Label labUltimaCompra 
         BackColor       =   &H00DFFFFF&
         ForeColor       =   &H00600000&
         Height          =   285
         Left            =   120
         MouseIcon       =   "frmListado.frx":168C
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   360
         Width           =   3355
      End
   End
   Begin MSComctlLib.ImageList imgExplore 
      Left            =   5040
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1996
            Key             =   "next"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1CB2
            Key             =   "back"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1FCE
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":212A
            Key             =   "plantilla"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":2A04
            Key             =   "help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":2D20
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":303C
            Key             =   "print"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":314E
            Key             =   "printcnfg"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":3260
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":357A
            Key             =   "form"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":39CC
            Key             =   "changedb"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":3E1E
            Key             =   "firstpage"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":4270
            Key             =   "previouspage"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":46C2
            Key             =   "nextpage"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":4B14
            Key             =   "lastpage"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":4F66
            Key             =   "exp"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbExplorer 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bback"
            Object.ToolTipText     =   "Anterior (Ctrl+A)."
            Style           =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bnext"
            Object.ToolTipText     =   "Siguiente (Ctrl+S)."
            Style           =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
            Object.ToolTipText     =   "Actualizar (Ctrl+Z)."
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "plantilla"
            Object.ToolTipText     =   "Plantillas interactivas"
            Style           =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   300
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Imprimir [Ctrl+I]"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "preview"
            Object.ToolTipText     =   "Preview [Ctrl+P]"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "firstpage"
            Object.ToolTipText     =   "Primera página."
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "previouspage"
            Object.ToolTipText     =   "Página anterior."
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nextpage"
            Object.ToolTipText     =   "Página Siguiente."
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "lastpage"
            Object.ToolTipText     =   "Última página."
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exp"
            Object.ToolTipText     =   "Expandir locales"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgHorizontal 
      Height          =   45
      Left            =   120
      MousePointer    =   7  'Size N S
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Menu MnuOpcion 
      Caption         =   "&Archivo"
      Begin VB.Menu MnuOpBack 
         Caption         =   "&Anterior"
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuOpNext 
         Caption         =   "&Siguiente"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuOpRefrescar 
         Caption         =   "&Refrescar"
         Shortcut        =   ^Z
      End
      Begin VB.Menu MnuOpLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuOpConfPage 
         Caption         =   "&Configurar Página"
      End
      Begin VB.Menu MnuOpPreview 
         Caption         =   "&Preview"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuOpLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpChangeDB 
         Caption         =   "Cambiar &Base de Datos"
      End
      Begin VB.Menu MnuOpLinea3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpSalir 
         Caption         =   "Sa&lir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuIrA 
      Caption         =   "&Ir a"
      Begin VB.Menu MnuIrAMantenimiento 
         Caption         =   "Mantenimiento de Artículos"
      End
      Begin VB.Menu MnuIrAWizard 
         Caption         =   "Wizard de Artículos"
      End
      Begin VB.Menu MnuIrAStockLocal 
         Caption         =   "&Stock en Locales"
      End
      Begin VB.Menu MnuIrALinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIrAMovimiento 
         Caption         =   "&Movimientos"
         Begin VB.Menu MnuIrAMovFisico 
            Caption         =   "&Movimientos físicos"
         End
         Begin VB.Menu MnuIrAControlMovFis 
            Caption         =   "&Control de Mov. Físicos"
         End
         Begin VB.Menu MnuIrAGenHisStock 
            Caption         =   "Genero Historico Stock"
         End
         Begin VB.Menu MnuIrAMovLinea 
            Caption         =   "-"
         End
         Begin VB.Menu MnuIrAMovVirt 
            Caption         =   "Movimientos &Virtuales"
         End
      End
      Begin VB.Menu MnuIrAArreglos 
         Caption         =   "&Arreglos"
         Begin VB.Menu MnuIrAPendRetiro 
            Caption         =   "&Pendientes de Retiro"
         End
         Begin VB.Menu MnuIrAArrIngEspecial 
            Caption         =   "&Ingreso Especial"
         End
         Begin VB.Menu MnuIrATraslEspecial 
            Caption         =   "&Traslado Especial"
         End
         Begin VB.Menu MnuIrAArrLinea 
            Caption         =   "-"
         End
         Begin VB.Menu MnuIrACorrStockVir 
            Caption         =   "Corrijo Stock &Virtual"
         End
         Begin VB.Menu MnuIrAArrArregloStk 
            Caption         =   "&Arreglo Stock"
         End
         Begin VB.Menu MnuIrAVerifStock 
            Caption         =   "&Verifico Stock"
         End
      End
   End
   Begin VB.Menu MnuPlantillas 
      Caption         =   "&Plantillas"
      Begin VB.Menu MnuPlaIndex 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu MnuPUConteo 
      Caption         =   "PopUpConteo"
      Visible         =   0   'False
      Begin VB.Menu MnuPUConteoInsert 
         Caption         =   "Insertar conteo para el local"
      End
      Begin VB.Menu MnuPUConteoLine 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPUConteoCancel 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type tPlStockTotal
    Estado As Integer
    Plantilla As Long
End Type

Dim ArticuloEditado As clsArticulo
Dim colAccionesStock As Collection
Dim colArtsEspecifico As Collection

Private arrPlStockTotal() As tPlStockTotal
Private lIDArticulo As Long
Private bSizeAjuste As Boolean
Private Rs1 As rdoResultset
Private aTexto As String
Private bCargarImpresion As Boolean

Public Sub CargoAccionesStock()
On Error GoTo errCAS
    cboAcciones.Clear
    Set colAccionesStock = New Collection
    Dim oAccion As clsAccionesStock
    Cons = "SELECT CodID, RTrim(CodTexto) Nombre, RTrim(CodTexto2) Bits FROM Codigos WHERE CodCual = 158 ORDER BY CodTexto"
    Dim rsAS As rdoResultset
    Set rsAS = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAS.EOF
        
        Set oAccion = New clsAccionesStock
        With oAccion
            .ID = rsAS("CodID")
            .Nombre = rsAS("Nombre")
            .Bits = rsAS("Bits")
        End With
        colAccionesStock.Add oAccion
        
        cboAcciones.AddItem oAccion.Nombre
        cboAcciones.ItemData(cboAcciones.NewIndex) = oAccion.ID
        
        Set oAccion = Nothing
        rsAS.MoveNext
    Loop
    rsAS.Close
    Exit Sub
errCAS:
    clsGeneral.OcurrioError "Error al cargar las acciones de stock.", Err.Description, "Error"
End Sub

Private Sub CargarUltimasAccionesGrabadas()
On Error GoTo errGrabar
    Screen.MousePointer = 11
    vsAcciones.Rows = 1
    Cons = "SELECT Top 20 AHSFecha, RTrim(UsuIdentificacion) Usu, RTrim(CodTexto) Nom " & _
        "FROM ArticuloHistorialStock INNER JOIN Usuario ON AHSUsuario = UsuCodigo " & _
        "INNER JOIN Codigos ON CodId = AHSAccionStock AND CodCual = 158 " & _
        "WHERE AHSArticulo = " & tArticulo.prm_ArtID & _
        " ORDER BY AHSFecha Desc"
    Dim rsCU As rdoResultset
    Set rsCU = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsCU.EOF
        With vsAcciones
            .AddItem Format(rsCU("AHSFecha"), "dd/mm/yy hh:nn")
            .Cell(flexcpText, .Rows - 1, 1) = rsCU("Usu")
            .Cell(flexcpText, .Rows - 1, 2) = rsCU("Nom")
        End With
        rsCU.MoveNext
    Loop
    rsCU.Close
    Screen.MousePointer = 0
    Exit Sub
errGrabar:
    clsGeneral.OcurrioError "Error al cargar las últimas acciones de stock.", Err.Description, "Error"
End Sub

Private Sub ps_SetPrecioVigente()
On Error GoTo errSPV
    Cons = "Select PViPrecio From PrecioVigente" _
        & " Where PViArticulo = " & tArticulo.prm_ArtID _
        & " And PViMoneda = 1" _
        & " And PViHabilitado = 1 And PViTipoCuota = (SELECT ParValor FROM Parametro WHERE ParNombre = 'TipoCuotaContado')"
    Dim rsP As rdoResultset
    Set rsP = cBase.OpenResultset(Cons, rdOpenForwardOnly)
    If Not rsP.EOF Then lbPrecioCtdo.Caption = "$ " & Format(rsP(0), "#,##0.00")
    rsP.Close
    
    'Si no hay $ presento USD.
    If lbPrecioCtdo.Caption = "" Then
        Cons = "Select PViPrecio From PrecioVigente" _
            & " Where PViArticulo = " & tArticulo.prm_ArtID _
            & " And PViMoneda = 2" _
            & " And PViHabilitado = 1 And PViTipoCuota = (SELECT ParValor FROM Parametro WHERE ParNombre = 'TipoCuotaContado')"
        Set rsP = cBase.OpenResultset(Cons, rdOpenForwardOnly)
        If Not rsP.EOF Then lbPrecioCtdo.Caption = "USD " & Format(rsP(0), "#,##0.00")
        rsP.Close
    End If
    Exit Sub
errSPV:
    clsGeneral.OcurrioError "Error al leer el contado vigente.", Err.Description
End Sub

Private Function BuscoSituacionArticulo() As String
On Error GoTo errBSA
Dim sCons As String
    sCons = "select dbo.[SituacionArticuloStock](" & ArticuloEditado.ID & ")"
    Dim rsA As rdoResultset
    Set rsA = cBase.OpenResultset(sCons, rdOpenDynamic, rdConcurValues)
    If Not rsA.EOF Then
        BuscoSituacionArticulo = rsA(0)
    End If
    rsA.Close
    Exit Function
errBSA:
End Function

Private Sub ps_SetValoresStock()
    
'    Dim sStock As String
'
'    sStock = "[EnUso]. [EnVta][Disponible] [EnWeb] y [XMayor].  [smin]"
'
'    If ArticuloEditado.EnUso Then    'chEnUso.Value = 1 Then
'        sStock = Replace(sStock, "[EnUso]", "En Uso")
'    Else
'        sStock = Replace(sStock, "[EnUso]", "Fuera de uso")
'    End If
'    If ArticuloEditado.EnVenta Then   'chHabilitado.Value = 1 Then
'        sStock = Replace(sStock, "[EnVta]", " En Venta")
'    Else
'        sStock = Replace(sStock, "[EnVta]", " No en venta")
'    End If
'    If IsDate(tbDisponibleDesde.Text) Then
'        If CDate(tbDisponibleDesde.Text) >= Date Then
'            sStock = Replace(sStock, "[Disponible]", " a partir del " & tbDisponibleDesde.Text)
'        Else
'            sStock = Replace(sStock, "[Disponible]", "")
'        End If
'    Else
'        sStock = Replace(sStock, "[Disponible]", "")
'    End If
'    sStock = Replace(sStock, "[EnWeb]", IIf(ArticuloEditado.EnWeb, "en Web", "NO está en Web"))
'
''    If chEnWeb.Value = 1 Then
''        sStock = Replace(sStock, "[EnWeb]", "en Web")
''    Else
''        sStock = Replace(sStock, "[EnWeb]", "NO está en Web")
''    End If
'    If IsNumeric(tSMin.Text) Then
'        'Stk mín.
'        sStock = Replace(sStock, "[smin]", "Stk mín. " & tSMin.Text)
'    Else
'        sStock = Replace(sStock, "[smin]", "")
'    End If
'
'    sStock = Replace(sStock, "[EnWeb]", IIf(ArticuloEditado.EnWeb, "en Web", "NO está en Web"))
'    sStock = Replace(sStock, "[XMayor]", IIf(ArticuloEditado.EnWeb, "por mayor", "NO por mayor"))
'
    lbStock.Caption = " " & BuscoSituacionArticulo
    
    lbStock.BackColor = &HF0FFFF
    
    If Not ArticuloEditado.EnUso Then
        
        lbStock.ForeColor = &HC0&
    
    ElseIf ArticuloEditado.EnUso And ArticuloEditado.AlPorMayor _
        And ArticuloEditado.EnVenta And ArticuloEditado.EnWeb _
        And ArticuloEditado.EnWebConPrecio Then
        
        lbStock.ForeColor = &H6000&
        
    Else
        lbStock.ForeColor = vbBlack
    End If
    
    
End Sub

Private Sub ps_StatusEdit(ByVal bEdit As Boolean)
    picEdit.Visible = bEdit
    vsConsulta.Visible = Not bEdit
    vsLocal.Visible = Not bEdit
    tArticulo.Enabled = Not bEdit
    MnuOpcion.Enabled = Not bEdit
    tbExplorer.Enabled = Not bEdit
    lbPrecioCtdo.Enabled = Not bEdit
    lbArregloStock.Enabled = Not bEdit
    lbStock.Enabled = Not bEdit
    labUltimaCompra.Enabled = Not bEdit
End Sub

Private Sub ps_LoadArregloStock()
On Error GoTo errLAS
Dim rsS As rdoResultset
    Set rsS = cBase.OpenResultset("SELECT Max(CVaFecha) From MovimientoStockFisico INNER JOIN ComentariosVarios ON CVaOrigen = 25 AND CVaIdOrigen = MSFDocumento " & _
                        "WHERE MSFTipoDocumento = 25 AND MSFArticulo = " & tArticulo.prm_ArtID, rdOpenDynamic, rdConcurValues)
    If Not rsS.EOF Then
        If Not IsNull(rsS(0)) Then lbArregloStock.Caption = "Últ. ajuste: " & Format(rsS(0), "dd/mm/yy") & " "
    End If
    rsS.Close
    Exit Sub
errLAS:
    clsGeneral.OcurrioError "Error al leer la información de ajuste de stock.", Err.Description
End Sub

Public Sub SetArticuloParmetro(ByVal lArtParam As Long)
    If lArtParam > 0 Then BuscoArticuloPorID lArtParam, False, False
End Sub

Private Sub SetPlantillaStockEstado()
Dim vPlantillas() As String
Dim iCont As Integer
    
    If prmPlStockEstado <> "" Then
        vPlantillas = Split(prmPlStockEstado, ",")
        For iCont = 0 To UBound(vPlantillas)
            If InStr(vPlantillas(iCont), ":") > 0 Then
                ReDim Preserve arrPlStockTotal(UBound(arrPlStockTotal) + 1)
                With arrPlStockTotal(UBound(arrPlStockTotal))
                    .Estado = Mid(vPlantillas(iCont), 1, InStr(1, vPlantillas(iCont), ":", vbTextCompare) - 1)
                    .Plantilla = Mid(vPlantillas(iCont), InStr(1, vPlantillas(iCont), ":", vbTextCompare) + 1)
                End With
            End If
        Next
    End If
    
End Sub

Private Sub AccionLimpiar()
    tArticulo.Text = "": tArticulo.Tag = ""
    vsConsulta.Rows = 1: vsLocal.Rows = 1
End Sub

Private Sub butModificar_Click()
On Error GoTo errGrabar
    
    If tSMin.Text <> "" Then
        If Not IsNumeric(tSMin.Text) Then
            MsgBox "El valor ingresado no es numérico.", vbExclamation, "Atención"
            If tSMin.Enabled Then tSMin.SetFocus
            Exit Sub
        Else
            If CInt(tSMin.Text) < 0 Then
                MsgBox "El valor ingresado es incorrecto, sólo mayores a cero.", vbExclamation, "Atención"
                If tSMin.Enabled Then tSMin.SetFocus
                Exit Sub
            ElseIf CInt(tSMin.Text) > 255 Then
                MsgBox "El valor ingresado es incorrecto, sólo se permite entre 0 a 255.", vbExclamation, "Atención"
                If tSMin.Enabled Then tSMin.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    If Trim(tbDemora.Text) <> "" Then
        If Not IsNumeric(tbDemora.Text) Then
            MsgBox "La demora ingresada no es un valor numérico.", vbExclamation, "Atención"
            If tbDemora.Enabled Then tbDemora.SetFocus
            Exit Sub
        Else
            If CInt(tbDemora.Text) < 0 Then
                MsgBox "El valor ingresado es incorrecto, sólo mayores a cero.", vbExclamation, "Atención"
                If tbDemora.Enabled Then tbDemora.SetFocus
                Exit Sub
            ElseIf CInt(tbDemora.Text) > 255 Then
                MsgBox "El valor ingresado es incorrecto, sólo se permite entre 0 a 255.", vbExclamation, "Atención"
                If tbDemora.Enabled Then tbDemora.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    If Not IsDate(tbDisponibleDesde.Text) And Trim(tbDisponibleDesde.Text) <> "" Then
        MsgBox "Dato incorrecto.", vbExclamation, "Validación"
        tbDisponibleDesde.SetFocus
        Exit Sub
    End If
    
    'Si tengo habilitado el x mayor entonces pido que me ingrese los datos.
    If cboAcciones.ListIndex > -1 And BitDadaAccion(6) = 1 Then
        If Not IsNumeric(txtQMaxXMayor.Text) Then
            MsgBox "Debe indicar la cantidad al por mayor.", vbExclamation, "ATENCIÓN"
            txtQMaxXMayor.SetFocus
            Exit Sub
        End If
    End If
    
    Dim rsQ As rdoResultset
    Dim iQ As Integer
    If cboAcciones.ListIndex > -1 And BitDadaAccion(7) = 1 Then
        'Saco a cuantos clientes se les enviara el aviso.
        Cons = "SELECT IsNull(Count(*), 0) as Cantidad" & _
            " FROM cgsa.dbo.AvisoLlegada, cgsa.dbo.Articulo" & _
            " WHERE ArtId = " & tArticulo.prm_ArtID & " And ALlFechaNotificado Is Null" & _
            " And (ArtId = ALlArticulo OR ALLArticulo = ArtSustituyeA)"
        Set rsQ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        iQ = rsQ(0)
        rsQ.Close
        
        If iQ > 0 Then
            If MsgBox("Esta acción implicará enviar un mensaje de aviso de llegada a " & iQ & _
                " personas que están en la lista de aviso." & vbCrLf & vbCrLf & _
                "¿Desea continuar asignando la acción?", vbQuestion + vbYesNo, "AVISO DE LLEGADA") = vbNo Then Exit Sub
            
        End If
        
    ElseIf cboAcciones.ListIndex > -1 And BitDadaAccion(7) = 0 Then
    
        'Saco a cuantos clientes se les enviara el aviso.
        Cons = "SELECT IsNull(Count(*), 0) as Cantidad" & _
            " FROM cgsa.dbo.AvisoLlegada, cgsa.dbo.Articulo" & _
            " WHERE ArtId = " & tArticulo.prm_ArtID & " And ALlFechaNotificado Is Null" & _
            " And (ArtId = ALlArticulo OR ALLArticulo = ArtSustituyeA)"
        
        Set rsQ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        iQ = rsQ(0)
        rsQ.Close
        
        If iQ = 0 Then
            MsgBox "Esta acción no va a tener efecto ya que no hay clientes en la lista de aviso.", vbInformation, "AVISO DE LLEGADA"
            Exit Sub
        End If
    
    End If
    
    
    If MsgBox("¿Confirma modificar la ficha del artículo?", vbYesNo + vbQuestion, "Grabar") = vbYes Then
    
        FechaDelServidor
        
        If butModificar.Tag = "" Then
            butModificar.Tag = InputBox("Ingrese su dígito de Usuario", "Usuario")
            If butModificar.Tag = "" Then Exit Sub
            butModificar.Tag = BuscoUsuarioDigito(Val(butModificar.Tag), True)
            If butModificar.Tag = "0" Then butModificar.Tag = "": Exit Sub
        End If
        
        If cboAcciones.ListIndex > -1 Then
            Cons = "EXEC prg_ArticuloAccionStock " & tArticulo.prm_ArtID & ", " & cboAcciones.ItemData(cboAcciones.ListIndex) & ", " & Val(butModificar.Tag)
            cBase.Execute (Cons)
        End If
        
        Cons = "Select * From Articulo Where ArtID = " & tArticulo.prm_ArtID
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        
        If IsNumeric(tSMin.Text) Then RsAux("ArtStockMinimo") = CInt(tSMin.Text) Else RsAux("ArtStockMinimo") = 0
                
        RsAux("ArtModificado") = Format(gFechaServidor, "mm/dd/yyyy hh:nn:ss")
        RsAux("ArtUsuModificacion") = Val(butModificar.Tag)
        
        If IsNumeric(tbDemora.Text) Then RsAux("ArtDemoraEntrega") = tbDemora.Text Else RsAux("ArtDemoraEntrega") = Null
        If IsDate(tbDisponibleDesde.Text) Then RsAux("ArtDisponibleDesde") = Format(tbDisponibleDesde.Text, "yyyy/mm/dd") Else RsAux("ArtDisponibleDesde") = Null
        
        RsAux("ArtEnVentaXMayor") = Val(txtQMaxXMayor.Text)
        
        'If Val(txtQMaxXMayor.Text) > 0 And IsNumeric(txtStockMinXMayor.Text) Then
        '24/3/2011 Matilde pidió que no toquemos este dato.
        If IsNumeric(txtStockMinXMayor.Text) Then
            RsAux("ArtStkMinXMayor") = Val(txtStockMinXMayor.Text)
            ArticuloEditado.StockMinimoXMayor = Val(txtStockMinXMayor.Text)
'        Else
'            RsAux("ArtStkMinXMayor") = Null
        End If
        
        RsAux.Update
        RsAux.Close
            
        ArticuloEditado.AlPorMayor = Val(txtQMaxXMayor.Text)
        
'            cBase.Execute "INSERT INTO logdb.dbo.Mensaje (MenAsunto, MenDe, MenEnviado, MenFechaHora, MenPublico, MenTexto, MenCategoria) VALUES(" & _
'                "'No vender x mayor " & ArticuloEditado.Nombre & "', " & paCodigoDeUsuario & ", GetDate(), GetDate()+.001, 1, 'Se cierra la venta por mayor de " & Replace(ArticuloEditado.Nombre, "'", "''") & "', 767)"

        
        'Ajusto los valores de stock.
        'ps_SetValoresStock
        
        ps_StatusEdit False
        If tArticulo.LoadArticulo(tArticulo.prm_ArtID) Then s_GetDataArticulo
        
        tArticulo.SetFocus
    End If
    
Exit Sub
errGrabar:
    clsGeneral.OcurrioError "Error al intentar modificar el artículo.", Err.Description
End Sub

Private Sub butCancel_Click()
    ps_StatusEdit False
End Sub

Private Sub cboAcciones_Change()
    lblInfoAccion.Caption = ""
End Sub

Private Sub cboAcciones_Click()
    DespliegoInfoAccion
    'si el stock al por mayor está deshabilitado y con este se prende entonces le pongo 1.
    If BitDadaAccion(6) = 1 Then
        If ArticuloEditado.AlPorMayor = 0 Then
            txtQMaxXMayor.Text = 1
            txtStockMinXMayor.Text = 1
        Else
            txtQMaxXMayor.Text = ArticuloEditado.AlPorMayor
            txtStockMinXMayor.Text = ArticuloEditado.StockMinimoXMayor
        End If
    Else
        txtQMaxXMayor.Text = 0
    End If
End Sub

Private Sub cboAcciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call butModificar_Click
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    
'    chEnWeb.Visible = miConexion.AccesoAlMenu("Datos_Web")
    
    paCodigoDeUsuario = miConexion.UsuarioLogueado(True)
    With vsCampos
        .Rows = 0
        '.FixedRows = 0
        .BackColorAlternate = &HD0D8CD '&HB8C5B4
        '.BackColor = vbWhite
        .BorderStyle = flexBorderNone
        .Cols = 2
        .ColWidth(0) = 2700
        .ExtendLastCol = True
        .HighLight = flexHighlightNever
        .ColAlignment(1) = flexAlignCenterCenter
        .RowHeightMin = 285
        .FocusRect = flexFocusNone
    End With
    
    pnlCampos.BackColor = vbWhite
    pnlCampos.Visible = False
    
    ReDim arrClientes(25)
    ReDim arrPlStockTotal(0)
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    InicializoGrillas
    AccionLimpiar
    bCargarImpresion = True
    vsListado.Orientation = orPortrait
    CargoAccionesStock
    BorroTodo
    CargoMenuPlantillas
    SetPlantillaStockEstado
    Status.Panels("bd").Text = "BD: " & miConexion.RetornoPropiedad(bdb:=True) & "  "
    Status.Tag = miConexion.RetornoPropiedad(bdb:=True)
    MenuExplorer
    InicializoToolbar
    
    With tArticulo
        Set .Connect = cBase
        .KeyQuerySP = "stktotal"
        .DisplayCodigoArticulo = True
    End With
    
    MenuPlantilla False
    
    vsListado.Visible = False
    'AccionPreview
    fFiltros.Top = 480
    imgHorizontal.Top = Me.ScaleHeight / 2
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub
Private Sub InicializoGrillas()
    On Error Resume Next
    
    With vsConsulta
        .Redraw = False
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        .WordWrap = False
        .Cols = 1: .Rows = 1
        .FormatString = "<Estado|>Disponible|>No Disponible|>Total|"
        .ColWidth(0) = 2000: .ColWidth(3) = 1000: .ColWidth(4) = 15
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(0) = True
        .Redraw = True
    End With
    With vsLocal
        .Redraw = False
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        .WordWrap = False
        .Cols = 1: .Rows = 1
        .FormatString = "<Local|<Estado|>Cantidad|"
        .ColWidth(0) = 2300: .ColWidth(1) = 1500: .ColWidth(2) = 1000: .ColWidth(3) = 15
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(0) = True
        .Redraw = True
    End With
    With vsAcciones
        .Redraw = False
        .Cols = 1: .Rows = 1
        .FormatString = "<Fecha|<Usuario|<Acción"
        .ColWidth(0) = 1300: .ColWidth(1) = 1000
        .Redraw = True
    End With
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
      
    
    Screen.MousePointer = 11
    If Me.Height < fFiltros.Top + fFiltros.Height + 1700 Then
        Me.Height = fFiltros.Top + fFiltros.Height + 1700
    End If
    
    fFiltros.Left = 0
    fFiltros.Width = Me.ScaleWidth '- (vsListado.Left * 2)
    
    vsListado.Move fFiltros.Left + 60, fFiltros.Top + fFiltros.Height + 20, fFiltros.Width - 120, _
        Me.ScaleHeight - (vsListado.Top + Status.Height + 20)
    
    With imgHorizontal
        .Left = vsListado.Left
        .Width = vsListado.Width
        If .Top > Me.ScaleHeight Then .Top = Me.ScaleHeight - 800
        If .Top < fFiltros.Height + tbExplorer.Height + 300 Then .Top = fFiltros.Height + tbExplorer.Height + 300
        If Not vsListado.Visible Then
            .ZOrder 0
        End If
    End With
    
    vsConsulta.Move vsListado.Left, vsListado.Top, vsListado.Width, imgHorizontal.Top - vsListado.Top
    
    With vsLocal
        .Left = vsConsulta.Left
        .Width = vsListado.Width
        .Top = imgHorizontal.Top + imgHorizontal.Height
        .Height = vsListado.Top + vsListado.Height - .Top - 20
    End With
    
    lbStock.Width = fFiltros.Width - (lbStock.Left * 2)
    
    lbArregloStock.Width = fFiltros.Width - (lbArregloStock.Left + lbStock.Left)
    picEdit.Move 30, fFiltros.Top + fFiltros.Height, Me.ScaleWidth - 60, vsLocal.Height + vsConsulta.Height + 30
    lblVtasXDias.Left = picEdit.Width - (lblVtasXDias.Width + 45)
    
    lbStock.Width = lbStock.Width - (lblVtasXDias.Width - imgDown.Width)
    
    butCancel.Left = picEdit.ScaleWidth - (120 + butCancel.Width)
    butModificar.Left = butCancel.Left
    vsAcciones.Height = picEdit.ScaleHeight - vsAcciones.Top - 120
    lblInfoAccion.Width = picEdit.ScaleWidth - (lblInfoAccion.Left * 2)
    lblInfoQMinXMayor.Width = picEdit.ScaleWidth - (lblInfoQMinXMayor.Left + 100)
    lblInfoStckMinXMayor.Width = picEdit.ScaleWidth - (lblInfoStckMinXMayor.Left + 100)
    
    vsCampos.Width = pnlCampos.ScaleWidth
    
    imgDown.Left = (lbStock.Width + lbStock.Left) - imgDown.Width - 30
    imgDown.Top = lbStock.Top + 15
    pnlCampos.Top = lbStock.Top + lbStock.Height
    pnlCampos.Left = (imgDown.Left + imgDown.Width) - pnlCampos.Width
    Screen.MousePointer = 0
    
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    Set cBase = Nothing
    Set eBase = Nothing
    End
    
End Sub

Private Sub AccionConsultar()
Dim rs As rdoResultset
    
    On Error GoTo errConsultar
    Screen.MousePointer = 11
    tSMin.Text = tSMin.Tag
    
    bCargarImpresion = True
    Set colArtsEspecifico = Nothing
    Set colArtsEspecifico = New Collection
    vsConsulta.Rows = 1: vsLocal.Rows = 1
    CargoStock
    Me.Refresh
    ButtonRegistros vsListado.Visible
    If tbExplorer.Buttons("plantilla").ButtonMenus.Count > 0 Then MenuPlantilla True
    Foco tArticulo
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Sub imgDown_Click()
    pnlCampos.Top = vsConsulta.Top - 60 '  lbStock.Top + lbStock.Height
    pnlCampos.Left = (imgDown.Left + imgDown.Width) - pnlCampos.Width
    
    pnlCampos.Visible = Not pnlCampos.Visible
    
End Sub

Private Sub imgHorizontal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bSizeAjuste = True
    With picHorizontal
        .Move imgHorizontal.Left, imgHorizontal.Top, imgHorizontal.Width, imgHorizontal.Height
        .Visible = True
        .ZOrder 0
    End With
End Sub

Private Sub imgHorizontal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bSizeAjuste Then
        If Y < picHorizontal.Top Then
            picHorizontal.Move vsConsulta.Left, imgHorizontal.Top + Y, vsConsulta.Width
        Else
            picHorizontal.Move vsConsulta.Left, imgHorizontal.Top - Y, vsConsulta.Width
        End If
    End If
End Sub

Private Sub imgHorizontal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If picHorizontal.Top < fFiltros.Height + 800 Then picHorizontal.Top = fFiltros.Height + 800
    If picHorizontal.Top + 800 > Me.ScaleHeight Then picHorizontal.Top = Me.ScaleHeight - 800
    imgHorizontal.Move 0, picHorizontal.Top
    
    picHorizontal.Visible = False
    bSizeAjuste = False
    Call Form_Resize
End Sub

Private Sub Label3_Click()
    Foco tArticulo
End Sub

Private Sub labUltimaCompra_DblClick()
    If tArticulo.prm_ArtID > 0 And prmPlUltimaCpa <> 0 Then
        EjecutarApp App.Path & "\appExploreMsg.exe ", prmPlUltimaCpa & ":" & tArticulo.prm_ArtID
    ElseIf tArticulo.prm_ArtID = 0 Then
        MsgBox "Seleccione un artículo.", vbInformation, "Validación"
    End If
End Sub

Private Sub lbArregloStock_DblClick()
    If tArticulo.prm_ArtID > 0 And prmPlAjusteStock <> 0 Then
        EjecutarApp App.Path & "\appExploreMsg.exe ", prmPlAjusteStock & ":" & tArticulo.prm_ArtID
    ElseIf tArticulo.prm_ArtID = 0 Then
        MsgBox "Seleccione un artículo.", vbInformation, "Validación"
    End If
End Sub

Private Sub lbPrecioCtdo_DblClick()
On Error Resume Next
    If tArticulo.prm_ArtID > 0 Then
        EjecutarApp App.Path & "\Precio_Articulo.exe", tArticulo.prm_ArtID
    Else
        MsgBox "Seleccione un artículo.", vbInformation, "Validación"
    End If
End Sub

Private Sub lbStock_DblClick()
    pnlCampos.Visible = False
    If tArticulo.prm_ArtID > 0 Then
        ps_StatusEdit True
        tSMin.SetFocus
    Else
        MsgBox "Seleccione un artículo.", vbInformation, "Validación"
    End If
End Sub

Private Sub MnuIrAArrArregloStk_Click()
    EjecutarApp App.Path & "\Arreglo Stock.exe"
End Sub

Private Sub MnuIrAArrIngEspecial_Click()
    EjecutarApp App.Path & "\Ingreso MercaderiaE.exe"
End Sub

Private Sub MnuIrAControlMovFis_Click()
    EjecutarApp App.Path & "\Control Movimiento Fisico.exe"
End Sub

Private Sub MnuIrACorrStockVir_Click()
    EjecutarApp App.Path & "\Corrijo Stock Virtual.exe"
End Sub

Private Sub MnuIrAGenHisStock_Click()
    EjecutarApp App.Path & "\Genero Historico Stock.exe"
End Sub

Private Sub MnuIrAMantenimiento_Click()
    EjecutarApp App.Path & "\articulos.exe", IIf(tArticulo.prm_ArtID > 0, tArticulo.prm_ArtID, "")
End Sub

Private Sub MnuIrAMovFisico_Click()
    EjecutarApp App.Path & "\Movimientos Fisicos.exe"
End Sub

Private Sub MnuIrAMovVirt_Click()
    EjecutarApp App.Path & "\Movimientos Virtuales.exe"
End Sub

Private Sub MnuIrAPendRetiro_Click()
    EjecutarApp App.Path & "\Pendientes de Retiro.exe"
End Sub

Private Sub MnuIrAStockLocal_Click()
    EjecutarApp App.Path & "\Stock de locales.exe"
End Sub

Private Sub MnuIrATraslEspecial_Click()
    EjecutarApp App.Path & "\Traslado_Mercaderia_Especial.exe"
End Sub

Private Sub MnuIrAVerifStock_Click()
    EjecutarApp App.Path & "\Verificacion de Stock.exe"
End Sub

Private Sub MnuIrAWizard_Click()
    EjecutarApp App.Path & "\wiz_articulo.exe", IIf(tArticulo.prm_ArtID > 0, tArticulo.prm_ArtID, "")
End Sub

Private Sub MnuOpBack_Click()
    Call tbExplorer_ButtonMenuClick(tbExplorer.Buttons("bback").ButtonMenus.Item(1))
End Sub

Private Sub MnuOpChangeDB_Click()
Dim newB As String
    
    On Error GoTo errCh
    
    If Not miConexion.AccesoAlMenu("Cambiar_Conexion") Then Exit Sub
    
    newB = InputBox("Ingrese el texto del login para la nueva conexión" & vbCrLf & _
                "Id de aplicación en archivo de conexiones.", "Cambio de Base de Datos")
    
    If Trim(newB) = "" Then Exit Sub
    If MsgBox("Está seguro de cambiar la base de datos al login " & newB, vbQuestion + vbYesNo + vbDefaultButton2, "Cambiar Base") = vbNo Then Exit Sub
    
    'Limpio la ficha
    AccionLimpiar
    
    newB = miConexion.TextoConexion(newB)
    If Trim(newB) = "" Then Exit Sub
    
    If Status.Tag = miConexion.RetornoPropiedad(bdb:=True) Then
        Me.BackColor = vbButtonFace
        fFiltros.BackColor = vbButtonFace
        tSMin.BackColor = vbButtonFace
        txtQMaxXMayor.BackColor = vbButtonFace
        txtStockMinXMayor.BackColor = vbButtonFace
    Else
        Me.BackColor = &HC0C000
        fFiltros.BackColor = &HC0C000
        tSMin.BackColor = &HC0C000
        txtQMaxXMayor.BackColor = &HC0C000
        txtStockMinXMayor.BackColor = &HC0C000
    End If
    Screen.MousePointer = 11
    On Error Resume Next
    cBase.Close
    On Error GoTo errCh
    Set cBase = Nothing
    If InicioConexionBD(newB) Then
        Status.Panels("bd").Text = "BD: " & miConexion.RetornoPropiedad(bdb:=True) & "  "
    Else
        Status.Panels("bd").Text = "BD: EN ERROR  "
    End If
    
    Screen.MousePointer = 0
    MsgBox "Ahora está trabajanbo en la nueva base de datos.", vbExclamation, "Base Cambiada OK"
    Exit Sub
    
errCh:
    Status.Panels("bd").Text = "BD: EN ERROR  "
    clsGeneral.OcurrioError "Error de Conexión." & vbCrLf & " La conexión está en estado de error, conectese a una base de datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuOpConfPage_Click()
    AccionConfigurar
End Sub

Private Sub MnuOpImprimir_Click()
    AccionImprimir True
End Sub

Private Sub MnuOpNext_Click()
On Error Resume Next
    Call tbExplorer_ButtonMenuClick(tbExplorer.Buttons("bnext").ButtonMenus.Item(1))
End Sub

Private Sub MnuOpPreview_Click()
    AccionPreview
End Sub

Private Sub MnuOpRefrescar_Click()
    AccionRefrescar
End Sub

Private Sub MnuOpSalir_Click()
    Unload Me
End Sub

Private Sub MnuPlaIndex_Click(Index As Integer)
    If tArticulo.prm_ArtID > 0 Then
        EjecutarApp App.Path & "\appExploreMsg.exe ", Val(MnuPlaIndex(Index).Tag) & ":" & tArticulo.prm_ArtID
    End If
End Sub

Private Sub MnuPUConteoInsert_Click()
    db_InsertConteo
End Sub


Private Sub picEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pnlCampos.Visible = False
End Sub

Private Sub tArticulo_Change()
    If tArticulo.prm_ArtID > 0 Then
        tArticulo.Tag = ""
        BorroTodo
        MenuPlantilla False
    End If
End Sub

Private Sub tArticulo_GotFocus()
    Status.Panels(2).Text = "Ingrese el artículo a consultar."
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrAP
    
    If KeyAscii = vbKeyReturn Then
        
        If tArticulo.Tag <> "" Then Exit Sub
        
        labUltimaCompra.Caption = ""
        lbStock.Caption = ""
        lbPrecioCtdo.Caption = ""
        lbArregloStock.Caption = ""
        vsConsulta.Rows = 1: vsLocal.Rows = 1
        
        If tArticulo.prm_ArtID > 0 Then s_GetDataArticulo
        tArticulo.SelectAll
        
    End If
    Exit Sub
ErrAP:
    clsGeneral.OcurrioError "Error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub tArticulo_LostFocus()
    Ayuda ""
End Sub

Private Sub tbDemora_GotFocus()
    With tbDemora
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tbDemora_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then txtQMaxXMayor.SetFocus
End Sub

Private Sub tbDisponibleDesde_GotFocus()
    With tbDisponibleDesde
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tbDisponibleDesde_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If Trim(tbDisponibleDesde.Text) <> "" Then
            If Not IsDate(tbDisponibleDesde.Text) Then
                tbDisponibleDesde_GotFocus
                Exit Sub
            Else
                tbDisponibleDesde.Text = Format(tbDisponibleDesde.Text, "dd/mm/yyyy")
            End If
        End If
        tbDemora.SetFocus
    End If
End Sub


Private Sub tSMin_GotFocus()
    With tSMin
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tSMin_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then tbDisponibleDesde.SetFocus

End Sub

Private Sub txtQMaxXMayor_Change()
    
    If Val(txtQMaxXMayor.Text) > 0 Then
        txtStockMinXMayor.Enabled = True
        txtStockMinXMayor.BackColor = vbWindowBackground
        txtStockMinXMayor.ForeColor = vbBlack
        If ArticuloEditado.StockMinimoXMayor > 0 Then txtStockMinXMayor.Text = ArticuloEditado.StockMinimoXMayor
    Else
        txtStockMinXMayor.Enabled = False
        txtStockMinXMayor.BackColor = vbButtonFace
        txtStockMinXMayor.Text = ""
    End If
    
End Sub

Private Sub txtQMaxXMayor_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If (Not IsNumeric(txtQMaxXMayor.Text) And Len(txtQMaxXMayor.Text) > 0) Or Val(txtQMaxXMayor.Text) < 0 Then
            MsgBox "Debe ingresar un número mayor o igual a cero.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        txtStockMinXMayor.SetFocus
    End If
End Sub

Private Sub txtQMaxXMayor_LostFocus()
    If Not IsNumeric(txtQMaxXMayor.Text) Then txtQMaxXMayor.Text = ""
End Sub

Private Sub txtStockMinXMayor_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then cboAcciones.SetFocus
End Sub

Private Sub vsConsulta_Click()
Dim lIDPl As Long
    pnlCampos.Visible = False
    If vsConsulta.Rows > 1 Then
        
        If vsConsulta.Col = 0 Then
            If vsConsulta.Cell(flexcpData, vsConsulta.Row, 0) < 0 Then
                'Plantillas de ventas telefonicas.
                lIDPl = GetIDPlantilla(vsConsulta.Cell(flexcpData, vsConsulta.Row, 0))
                If lIDPl > 0 Then
                    If tArticulo.prm_ArtID > 0 Then EjecutarApp App.Path & "\appExploreMsg.exe ", lIDPl & ":" & tArticulo.prm_ArtID
                End If
            ElseIf vsConsulta.Cell(flexcpData, vsConsulta.Row, 0) = TipoMovimientoEstado.ARetirar Then
                Screen.MousePointer = 11
                If tArticulo.prm_ArtID > 0 Then EjecutarApp App.Path & "\pendientes de retiro.exe", tArticulo.prm_ArtID
                Screen.MousePointer = 0
            ElseIf vsConsulta.Cell(flexcpData, vsConsulta.Row, 0) = TipoMovimientoEstado.AEntregar Then
                lIDPl = GetIDPlantilla(vsConsulta.Cell(flexcpData, vsConsulta.Row, 0))
                If lIDPl > 0 Then
                    If tArticulo.prm_ArtID > 0 Then EjecutarApp App.Path & "\appExploreMsg.exe ", lIDPl & ":" & tArticulo.prm_ArtID
                End If
            End If
        End If
        
    End If
    
End Sub
Private Function GetIDPlantilla(ByVal idEstado As Integer) As Long
Dim iCont As Integer
    GetIDPlantilla = 0
    For iCont = 1 To UBound(arrPlStockTotal)
        If arrPlStockTotal(iCont).Estado = idEstado Then
            GetIDPlantilla = arrPlStockTotal(iCont).Plantilla
            Exit For
        End If
    Next iCont
End Function

Private Sub vsConsulta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pnlCampos.Visible = False
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
        
        EncabezadoListado vsListado, "Consulta de Stock Total al " & Format(Date, FormatoFP), False
        vsListado.FileName = "Consulta de Stock Total"
        vsListado.FontBold = True
        vsListado.Paragraph = "Artículo: " & tArticulo.Text
        vsListado.Paragraph = ""
        vsListado.FontBold = False
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        vsListado.Paragraph = ""
        vsListado.FontBold = True
        vsListado.Paragraph = "Distribución en Locales"
        vsListado.FontBold = False
        vsListado.Paragraph = ""
        vsLocal.ExtendLastCol = False: vsListado.RenderControl = vsLocal.hwnd: vsLocal.ExtendLastCol = True
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
Private Sub CargoStock()
Dim sEstConteo As String
Dim bTotal As Boolean
Dim rs As rdoResultset
Dim QAFacturarR As Long
Dim QAFacturarE As Long, CantExtra As Long
On Error GoTo ErrCS
    
    If tArticulo.prm_ArtID = 0 Then Exit Sub
    Screen.MousePointer = 11
    sEstConteo = ""
    bTotal = True
    If Not miConexion.AccesoAlMenu("contadorverificador") Then sEstConteo = db_ConteoPendiente
    
    ArmoConsultaTotal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If RsAux.EOF Then
        RsAux.Close
'        Cons = "Select LocNombre, StLCantidad, EsMAbreviacion, StLEstado, LocCodigo, LocTipo From Local, StockLocal, EstadoMercaderia " _
            & " Where StlArticulo = " & tArticulo.prm_ArtID _
            & " And StlCantidad <> 0 And StLLocal = LocCodigo And StLEstado = EsMCodigo"
            
        Cons = "SELECT LocCodigo, CASE LocTipo WHEN 0 THEN 1 ELSE LocTipo END LocTipo , LocNombre, StLArticulo, StLLocal , StLEstado, StLCantidad, EsMAbreviacion" & _
                " FROM Local INNER JOIN (SELECT StLArticulo, StLLocal , StLEstado, SUM(StlCantidad) StLCantidad" & _
                " FROM StockLocal WHERE StLArticulo = [prmArticulo] Group By StLArticulo, StLLocal , StLEstado) As Stock ON Stock.StLLocal = LocCodigo" & _
                " INNER JOIN EstadoMercaderia ON StLEstado = EsMCodigo" & _
                " WHERE Stock.StLCantidad <> 0"

        Cons = Replace(Cons, "[prmArticulo]", tArticulo.prm_ArtID, , , vbTextCompare)
'        Cons = Replace(Cons, "[prmEstado]", RsAux!Estado, , , vbTextCompare)
            
            
        Set rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        Do While Not rs.EOF
            If InStr(1, sEstConteo, "," & rs!StLEstado & ",") > 0 And rs!LocCodigo = paCodigoDeSucursal Then
                bTotal = False
                InsertoFilaLocal Trim(rs!LocNombre), Trim(rs!EsMAbreviacion), "AContar", IIf(rs!LocTipo = 2, rs!LocCodigo, rs!LocCodigo * -1), rs!StLEstado
            Else
                InsertoFilaLocal Trim(rs!LocNombre), Trim(rs!EsMAbreviacion), rs!StLCantidad, IIf(rs!LocTipo = 2, rs!LocCodigo, rs!LocCodigo * -1), rs!StLEstado
            End If
            rs.MoveNext
        Loop
        rs.Close
        If vsLocal.Rows = 1 Then
            MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCIÓN"
        End If
        
    Else
    
        StockAFacturarEnviayRetira tArticulo.prm_ArtID, QAFacturarR, QAFacturarE
        If QAFacturarR > 0 Then InsertoFilaConsulta -1, "A Facturar Retirar", QAFacturarR, 0, False
        If QAFacturarE > 0 Then InsertoFilaConsulta -2, "A Facturar Envia", QAFacturarE, 0, False
        
        Dim ArtsEspecificos As String
        ArtsEspecificos = ObtenerCantidadArtsEspecificos(tArticulo.prm_ArtID)
        Dim bArtEspInsert As Boolean
        bArtEspInsert = False
        
        Do While Not RsAux.EOF
            If RsAux!StTTipoEstado = TipoEstadoMercaderia.Fisico Then
                
                CantExtra = CantidadNoDisponible(tArticulo.prm_ArtID, RsAux!Estado)
                
                If InStr(1, sEstConteo, "," & RsAux!Estado & ",") > 0 Then
                    If RsAux!StTCantidad - CantExtra <> 0 Then
                        bTotal = False
                        InsertoFilaConsulta 0, Trim(RsAux!EsMAbreviacion), "AContar", CantExtra
                    End If
                Else
                    If RsAux("Estado") = 274 Then
                        If RsAux!StTCantidad - CantExtra - ArtsEspecificos <> 0 Then
                            InsertoFilaConsulta 0, Trim(RsAux!EsMAbreviacion), RsAux!StTCantidad - CantExtra - ArtsEspecificos, CantExtra
                        End If
                        If ArtsEspecificos > 0 Then InsertoFilaConsulta -3, "Específicos", ArtsEspecificos, 0, False
                        bArtEspInsert = True
                    ElseIf RsAux("EsMBajaStockTotal") = 1 And Not IsNull(RsAux!StTCantidad) Then
                        'Este estado es NO DISPONIBLE.
                        InsertoFilaConsulta 0, Trim(RsAux!EsMAbreviacion), 0, RsAux!StTCantidad
                    Else
                        If RsAux!StTCantidad - CantExtra <> 0 Then
                            InsertoFilaConsulta 0, Trim(RsAux!EsMAbreviacion), RsAux!StTCantidad - CantExtra, CantExtra
                        End If
                    End If
                End If
                
                'Cons = "Select LocNombre, StLCantidad , LocCodigo, LocTipo From Local, StockLocal " _
                    & " Where StlArticulo = " & tArticulo.prm_ArtID _
                    & " And StLEstado = " & RsAux!Estado _
                    & " And StlCantidad <> 0 And StLLocal = LocCodigo"
                
                
                Cons = "SELECT LocCodigo, CASE LocTipo WHEN 0 THEN 1 ELSE LocTipo END LocTipo , LocNombre, StLArticulo, StLLocal , StLEstado, StLCantidad" & _
                    " FROM Local INNER JOIN (SELECT StLArticulo, StLLocal , StLEstado, SUM(StlCantidad) StLCantidad" & _
                    " FROM StockLocal WHERE StLArticulo = [prmArticulo] AND StlEstado = [prmEstado] Group By StLArticulo, StLLocal , StLEstado) As Stock ON Stock.StLLocal = LocCodigo" & _
                    " WHERE Stock.StLCantidad <> 0"

                Cons = Replace(Cons, "[prmArticulo]", tArticulo.prm_ArtID, , , vbTextCompare)
                Cons = Replace(Cons, "[prmEstado]", RsAux!Estado, , , vbTextCompare)
                
                Set rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
                
                Dim QEspLocal As Long
                Do While Not rs.EOF
                    If InStr(1, sEstConteo, "," & RsAux!Estado & ",") > 0 And rs!LocCodigo = paCodigoDeSucursal Then
                        bTotal = False
                        InsertoFilaLocal Trim(rs!LocNombre), Trim(RsAux!EsMAbreviacion), "AContar", IIf(rs!LocTipo = 2, rs!LocCodigo, rs!LocCodigo * -1), RsAux!Estado
                    Else
                        If rs("StLEstado") = 274 Then
                            QEspLocal = CantidadEspecificosEnElLocal(rs("StlLocal"))
                            If QEspLocal > 0 Then
                                InsertoFilaLocal Trim(rs!LocNombre), "Específico", QEspLocal, IIf(rs!LocTipo = 2, rs!LocCodigo, rs!LocCodigo * -1), RsAux!Estado
                            End If
                            InsertoFilaLocal Trim(rs!LocNombre), Trim(RsAux!EsMAbreviacion), rs!StLCantidad - QEspLocal, IIf(rs!LocTipo = 2, rs!LocCodigo, rs!LocCodigo * -1), RsAux!Estado
                        Else
                            InsertoFilaLocal Trim(rs!LocNombre), Trim(RsAux!EsMAbreviacion), rs!StLCantidad, IIf(rs!LocTipo = 2, rs!LocCodigo, rs!LocCodigo * -1), RsAux!Estado
                        End If
                    End If
                    rs.MoveNext
                Loop
                rs.Close
            Else
                'Estados Virtuales
                Select Case RsAux!Estado
                    Case TipoMovimientoEstado.AEntregar
                        If RsAux!StTCantidad - QAFacturarE <> 0 Then InsertoFilaConsulta RsAux!Estado, RetornoEstadoVirtual(RsAux!Estado), RsAux!StTCantidad - QAFacturarE, 0, False
                    Case TipoMovimientoEstado.ARetirar
                        If RsAux!StTCantidad - QAFacturarR <> 0 Then InsertoFilaConsulta RsAux!Estado, RetornoEstadoVirtual(RsAux!Estado), RsAux!StTCantidad - QAFacturarR, 0, False
                    Case TipoMovimientoEstado.Reserva
                        If RsAux!StTCantidad <> 0 Then InsertoFilaConsulta RsAux!Estado, RetornoEstadoVirtual(RsAux!Estado), RsAux!StTCantidad, 0, False
                End Select
            End If
            RsAux.MoveNext
        Loop
        RsAux.Close
        If ArtsEspecificos > 0 And Not bArtEspInsert Then InsertoFilaConsulta -3, "Específicos", ArtsEspecificos, 0, False
    '    CargoFisicosNoDisponibles
    End If
    If vsConsulta.Rows > 1 Then
        With vsLocal
            If .Rows > 1 Then .Select 1, 0, 1, 1
            .Sort = flexSortGenericAscending
            .Subtotal flexSTSum, 0, 2, "#,##0", Inactivo, Rojo, False, "%s"
        End With
        If bTotal Then
            With vsConsulta
                .Subtotal flexSTSum, -1, 1, "#,##0", Obligatorio, Rojo, True, "Total"
                .Subtotal flexSTSum, -1, 2, "#,##0", Obligatorio, Rojo, True, ""
                .Subtotal flexSTSum, -1, 3, "#,##0", Obligatorio, Rojo, True, ""
            End With
            With vsLocal
                .Subtotal flexSTSum, -1, 2, "#,##0", Obligatorio, Rojo, True, "Total"
                For CantExtra = .FixedRows To .Rows - 2
                        If .IsSubtotal(CantExtra) Then
                            If .IsCollapsed(CantExtra) <> flexOutlineCollapsed Then
                                .IsCollapsed(CantExtra) = flexOutlineCollapsed
                            End If
                        End If
                Next
            End With
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrCS:
    clsGeneral.OcurrioError "Ocurrio un error al cargar el stock.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function CantidadEspecificosEnElLocal(ByVal idLocal As Long) As Integer
Dim oAL As clsArticuloLocal
    
    For Each oAL In colArtsEspecifico
        If oAL.Deposito = idLocal Then
            CantidadEspecificosEnElLocal = oAL.Cantidad + CantidadEspecificosEnElLocal
        End If
    Next

End Function

Private Function ObtenerCantidadArtsEspecificos(ByVal idArticulo As Long)
On Error GoTo errOB
Dim sQ As String
Dim oAL As clsArticuloLocal

    ObtenerCantidadArtsEspecificos = 0
'    sQ = "SELECT Count(*) FROM ArticuloEspecifico WHERE AEsEstado = 1  And (AEsDocumento IS NULL or AEsTipoDocumento = 2) AND AEsArticulo = " & idArticulo
    
    sQ = "SELECT AEsLocal, Count(*) Cantidad FROM ArticuloEspecifico WHERE AEsEstado = 1  And (AEsDocumento IS NULL or AEsTipoDocumento = 2) AND AEsArticulo = " & idArticulo _
        & " GROUP BY AEsLocal"
    Dim rs As rdoResultset
    Set rs = cBase.OpenResultset(sQ, rdOpenDynamic, rdConcurValues)
    Do While Not rs.EOF
        
        Set oAL = New clsArticuloLocal
        oAL.Cantidad = rs("Cantidad")
        oAL.Deposito = rs("AEsLocal")
        
        colArtsEspecifico.Add oAL
        
        ObtenerCantidadArtsEspecificos = ObtenerCantidadArtsEspecificos + rs("Cantidad")
        
        rs.MoveNext
    Loop
    
'    If Not rs.EOF Then
'        ObtenerCantidadArtsEspecificos = rs(0)
'    End If
    rs.Close
    
    Exit Function
errOB:
End Function


Private Sub CargoFisicosNoDisponibles()
    
    Cons = "Select ArtID, ArtCodigo, ArtNombre, StTTipoEstado, StTEstado, StTCantidad, EsMAbreviacion From Articulo, StockTotal, EstadoMercaderia " _
        & " Where ArtHabilitado = 'S' And StTTipoEstado = " & TipoEstadoMercaderia.Fisico _
        & " And ArtID = " & tArticulo.prm_ArtID _
        & " And ArtID = StTArticulo And StTEstado = EsMCodigo And EsMBajaStockTotal = 1"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsAux.EOF
        If RsAux!StTCantidad <> 0 Then InsertoFilaConsulta 0, Trim(RsAux!EsMAbreviacion), 0, RsAux!StTCantidad
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub
Private Sub ArmoConsultaTotal()
    
    'Saco todos los artículos y su stock Total.---------------------------
    'Hago una unión por tipo de estado de mercadería. ArtID, ArtCodigo, ArtNombre,
    Cons = "Select  IsNull(StTTipoEstado, 1) as StTTipoEstado, EsMCodigo as Estado, StTCantidad, EsMAbreviacion, EsMBajaStockTotal " _
        & " From EstadoMercaderia " _
                & " Left Outer Join StockTotal On STTEstado = EsMCodigo And StTArticulo = " & tArticulo.prm_ArtID _
                & " And StTTipoEstado = " & TipoEstadoMercaderia.Fisico _
        & " Union" _
        & " Select StTTipoEstado, StTEstado as Estado, StTCantidad, EsMAbreviacion = '', EsMBajaStockTotal = 0 From StockTotal" _
        & " Where StTArticulo = " & tArticulo.prm_ArtID _
        & " And StTTipoEstado = " & TipoEstadoMercaderia.Virtual _
        & " Order by StTTipoEstado DESC"
        
End Sub
Private Function CantidadNoDisponible(idArticulo As Long, idEstado As Integer) As Currency
Dim RsStL As rdoResultset
    Cons = "Select Sum(StLCantidad) From StockLocal " _
        & " Where StlArticulo = " & idArticulo _
        & " And StLEstado = " & idEstado _
        & " And StLLocal IN (Select SucCodigo From Sucursal Where SucExtras = 1)"
    Set RsStL = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not IsNull(RsStL(0)) Then CantidadNoDisponible = RsStL(0) Else CantidadNoDisponible = 0
    RsStL.Close
End Function
Private Sub InsertoFilaConsulta(ByVal idEstado As Integer, ByVal sEstado As String, ByVal sDisponible As String, lNoDisponible As Long, Optional EstFisico As Boolean = True)
    With vsConsulta
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = Trim(sEstado)
        .Cell(flexcpData, .Rows - 1, 0) = idEstado
        If IsNumeric(sDisponible) Then
            .Cell(flexcpText, .Rows - 1, 1) = Format(Val(sDisponible), "#,##0")
            .Cell(flexcpText, .Rows - 1, 2) = Format(lNoDisponible, "#,##0")
            .Cell(flexcpText, .Rows - 1, 3) = Format(lNoDisponible + .Cell(flexcpValue, .Rows - 1, 1), "#,##0")
        Else
            .Cell(flexcpText, .Rows - 1, 1) = "?"
            .Cell(flexcpText, .Rows - 1, 2) = "?"
            .Cell(flexcpText, .Rows - 1, 3) = "?"
        End If
        .Cell(flexcpFontBold, .Rows - 1, 3) = True
        If Not EstFisico Then
            .Cell(flexcpForeColor, .Rows - 1, 0) = vbHighlight
            .Cell(flexcpFontUnderline, .Rows - 1, 0) = True
        End If
    End With
End Sub
Private Sub InsertoFilaLocal(ByVal sLocal As String, ByVal sEstado As String, ByVal sCant As String, ByVal lCodLocal As Long, ByVal lCodEstado As Long)
    With vsLocal
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = sLocal
        .Cell(flexcpData, .Rows - 1, 0) = lCodLocal
        If lCodLocal < 0 Then
            .Cell(flexcpForeColor, .Rows - 1, 0) = vbHighlight
            .Cell(flexcpFontUnderline, .Rows - 1, 0) = True
        End If
        .Cell(flexcpText, .Rows - 1, 1) = Trim(sEstado)
        .Cell(flexcpData, .Rows - 1, 1) = lCodEstado
        If IsNumeric(sCant) Then
            .Cell(flexcpText, .Rows - 1, 2) = Format(Val(sCant), "#,##0")
        Else
            .Cell(flexcpText, .Rows - 1, 2) = "?"
        End If
        .Cell(flexcpFontBold, .Rows - 1, 3) = True
    End With
End Sub

Private Sub BuscoArticuloPorID(ByVal idArticulo As Long, ByVal bNext As Boolean, ByVal bPrevious As Boolean)
'Atención el mapeo de error lo hago antes de entrar al procedimiento

    Screen.MousePointer = 11
    If Not (bNext Or bPrevious) Then
        Cons = "Select * From Articulo Where ArtID = " & idArticulo
    Else
        Cons = "Select Top 1 art2.* From Articulo art1, Articulo art2 Where art1.ArtID = " & idArticulo _
            & " And art2.ArtEnUso = 1 And art1.ArtCodigo "
        If bNext Then
            Cons = Cons & " < art2.ArtCodigo Order By art2.ArtCodigo Asc"
        Else
            Cons = Cons & " > art2.ArtCodigo Order By art2.ArtCodigo Desc"
        End If
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    BorroTodo
    
    If RsAux.EOF Then
        RsAux.Close
        tArticulo.Tag = ""
        If (bNext Or bPrevious) And tArticulo.prm_ArtID > 0 Then AccionConsultar
    Else
        idArticulo = RsAux!ArtId
        RsAux.Close
        If tArticulo.LoadArticulo(idArticulo) Then s_GetDataArticulo
    End If
    Screen.MousePointer = 0
    
End Sub

Private Sub StockAFacturarEnviayRetira(Articulo As Long, ByRef QARetirar As Long, ByRef QAEntregar As Long)
On Error GoTo ErrSAFE
    Cons = "Select IsNull(Sum(RVTARetirar), 0), 1 as T From VentaTelefonica, RenglonVtaTelefonica " _
            & " Where VTeTipo = " & TipoDocumento.ContadoDomicilio _
            & " And VTeDocumento Is Null And VTeAnulado Is Null " _
            & " And RVTArticulo = " & Articulo _
            & "And VTeCodigo = RVTVentaTelefonica"
    Cons = Cons & " Union All " _
        & "Select IsNull(Sum(REvAEntregar), 0) , 2 as T From Envio, RenglonEnvio" _
        & " Where EnvTipo = " & TipoEnvio.Cobranza _
        & " And EnvEstado NOT IN (" & EstadoEnvio.Anulado & " , " & EstadoEnvio.Entregado & " ," & EstadoEnvio.Rebotado & ")" _
        & " AND EnvDocumento IN(SELECT VTeCodigo FROM VentaTelefonica WHERE VTeDocumento IS NULL) And REvArticulo = " & Articulo _
        & " And REvAEntregar > 0" _
        & "And EnvCodigo = REvEnvio"
    Set Rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not Rs1.EOF
        If Rs1("T") = 1 Then
            QARetirar = Rs1(0)
        Else
            QAEntregar = Rs1(0)
        End If
        Rs1.MoveNext
    Loop
    Rs1.Close
    Exit Sub
ErrSAFE:
    clsGeneral.OcurrioError "Error al buscar el stock a facturar que se envía y la Cantidad a Retirar.", Err.Description
End Sub

Private Function RetornoUltimaCompra(lArt As Long) As String
On Error GoTo ErrRUC
    
    RetornoUltimaCompra = ""
    Cons = "Select * From Embarque, ArticuloFolder" _
        & " Where AFoArticulo = " & lArt & " And AFoTipo = 2" _
        & " And AFoCodigo = EmbID And EmbFLocal Is Not Null" _
        & " And EmbFArribo = (" _
            & " Select Max(EmbFArribo) From Embarque, ArticuloFolder" _
            & " Where AFoArticulo = " & lArt & " And AFoTipo = 2" _
            & " And AFoCodigo = EmbID And EmbFLocal Is Not Null)" _

    Set Rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly)
    
    If Not Rs1.EOF Then
        RetornoUltimaCompra = Format(Rs1("EmbFArribo"), "d/mm/yy") & "  Cantidad: " & Rs1("AFoCantidad")
    End If
    Rs1.Close
    
    Exit Function
ErrRUC:
    clsGeneral.OcurrioError "Ocurrio un error al buscar la última fecha de compra.", Err.Description
End Function

Private Function CompraNoImportacion(ByVal lArt As Long) As String
Dim lCant As Long

    CompraNoImportacion = ""
    'Saco la Fecha y el Documento de compra.
    Cons = "Select ComFecha, CReCantidad from Compra, CompraRenglon" _
        & " Where CReArticulo = " & lArt _
        & " And ComTipoDocumento In (" & TipoDocumento.Compracontado & ", " & TipoDocumento.CompraCredito & ")" _
        & " And ComCodigo = CReCompra And ComFecha = (" _
            & " Select Max(ComFecha) from Compra, CompraRenglon" _
            & " Where CReArticulo = " & lArt _
            & " And ComTipoDocumento In (" & TipoDocumento.Compracontado & ", " & TipoDocumento.CompraCredito & ")" _
            & " And ComCodigo = CReCompra)"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        CompraNoImportacion = Format(RsAux("ComFecha"), "d/mm/yy") & "  Cantidad: " & RsAux("CReCantidad")
    End If
    RsAux.Close
        
End Function

Private Sub BorroTodo()

    Set ArticuloEditado = Nothing

    labUltimaCompra.Caption = ""
    lbStock.Caption = ""
    lbPrecioCtdo.Caption = ""
    lbArregloStock.Caption = ""
    lblVtasXDias.Caption = ""
    lblInfoAccion.Caption = ""

    tArticulo.Tag = ""
    vsConsulta.Rows = 1: vsLocal.Rows = 1
    
    tSMin.Text = ""
    tbDisponibleDesde.Text = ""
    tbDemora.Text = ""
    
    tSMin.Tag = "": tSMin.Text = ""
    
    txtQMaxXMayor.Text = ""
    txtStockMinXMayor.Text = ""
    cboAcciones.Text = ""
    
    EstadoObjetos False
    If Not vsListado.Visible Then
        With tbExplorer
            .Buttons("nextpage").Enabled = False
            .Buttons("previouspage").Enabled = False
            .Refresh
        End With
    End If
    
End Sub

Private Sub EstadoObjetos(ByVal bEst As Boolean)
    
    If Not bEst Then
        tSMin.Text = ""
    End If
    
End Sub

Private Sub CargoMenuPlantillas()
On Error GoTo errStart
Dim iCont As Integer
    Screen.MousePointer = 11
    tbExplorer.Buttons("plantilla").ButtonMenus.Clear
    'Cons = "Select PlaCodigo, PlaNombre from Plantilla Where PlaCodigo IN (" & prmPlantillasArtStock & ") Order by PlaNombre"
    Cons = "SELECT RTrim(Texto2) as OpcionMenu, Puntaje as PlaCodigo From CodigoTexto Where Tipo = 105 ORDER BY Valor1"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        MnuPlantillas.Visible = False
    End If
    iCont = 0
    Do While Not RsAux.EOF
        With tbExplorer.Buttons("plantilla").ButtonMenus
            .Add
            If Not IsNull(RsAux!PlaCodigo) Then .Item(.Count).Tag = CStr(RsAux!PlaCodigo)
            .Item(.Count).Text = Trim(RsAux!OpcionMenu)
        End With
        If iCont <> 0 Then
            Load MnuPlaIndex(iCont)
        End If
        MnuPlaIndex(iCont).Caption = Trim(RsAux!OpcionMenu)
        If Not IsNull(RsAux!PlaCodigo) Then MnuPlaIndex(iCont).Tag = RsAux!PlaCodigo
        iCont = iCont + 1
        RsAux.MoveNext
    Loop
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
    
errStart:
    clsGeneral.OcurrioError "Error al cargar las plantillas.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tbExplorer_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    
    Select Case Button.Key
        Case "bback": Call tbExplorer_ButtonMenuClick(tbExplorer.Buttons("bback").ButtonMenus.Item(1))
        Case "bnext": Call tbExplorer_ButtonMenuClick(tbExplorer.Buttons("bnext").ButtonMenus.Item(1))
        
        Case "refresh": AccionRefrescar
        Case "preview": AccionPreview
        Case "print": AccionImprimir True
        
        Case "firstpage": IrAPagina vsListado, 1
        Case "previouspage"
            If vsListado.Visible Then
                IrAPagina vsListado, vsListado.PreviewPage - 1
            Else
                BuscoArticuloPorID lIDArticulo, False, True
            End If
        Case "nextpage"
            If vsListado.Visible Then
                IrAPagina vsListado, vsListado.PreviewPage + 1
            Else
                'Voy al primer código mayor.
                BuscoArticuloPorID lIDArticulo, True, False
            End If
        Case "lastpage": IrAPagina vsListado, vsListado.PageCount
        
        Case "exp": loc_ExpandirLocales
    End Select
    
End Sub

Private Sub tbExplorer_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim aIdx As Long
    DoEvents
    
    Select Case ButtonMenu.Parent.Key
        Case "bback", "bnext"
                    aIdx = Val(ButtonMenu.Tag)
                    BuscoArticuloPorID arrClientes(aIdx).Codigo, False, False
                    
        Case "plantilla"
            If tArticulo.prm_ArtID > 0 Then EjecutarApp App.Path & "\appExploreMsg.exe ", Val(ButtonMenu.Tag) & ":" & tArticulo.prm_ArtID
    End Select
    
End Sub

Private Sub MenuExplorer()
    On Error GoTo errMnu
    Dim aIdx As Integer, miX As Long

    miX = arrcli_Item(lIDArticulo)
    
    tbExplorer.Buttons("bback").ButtonMenus.Clear
    tbExplorer.Buttons("bnext").ButtonMenus.Clear
    
    For aIdx = LBound(arrClientes) To UBound(arrClientes)
        If arrClientes(aIdx).Codigo <> lIDArticulo And arrClientes(aIdx).Codigo <> 0 Then
            If aIdx < miX Then
                With tbExplorer.Buttons("bback").ButtonMenus
                    .Add Index:=1
                    .Item(1).Tag = CStr(aIdx)
                    .Item(1).Text = Trim(arrClientes(aIdx).Nombre)
                End With
                
            Else
                With tbExplorer.Buttons("bnext").ButtonMenus
                    .Add 'Index:=1
                    .Item(.Count).Tag = CStr(aIdx)
                    .Item(.Count).Text = Trim(arrClientes(aIdx).Nombre)
                End With
            End If
        End If
    Next
    
    If tbExplorer.Buttons("bback").ButtonMenus.Count = 0 Then tbExplorer.Buttons("bback").Enabled = False Else tbExplorer.Buttons("bback").Enabled = True
    If tbExplorer.Buttons("bnext").ButtonMenus.Count = 0 Then tbExplorer.Buttons("bnext").Enabled = False Else tbExplorer.Buttons("bnext").Enabled = True
    
    If lIDArticulo = 0 Then tbExplorer.Buttons("refresh").Enabled = False Else tbExplorer.Buttons("refresh").Enabled = True
errMnu:
End Sub

Private Sub InicializoToolbar()
On Error Resume Next
    
    Set tbExplorer.ImageList = imgExplore
    tbExplorer.Buttons("bback").Image = "back"
    tbExplorer.Buttons("bnext").Image = "next"
    tbExplorer.Buttons("refresh").Image = "refresh"
    tbExplorer.Buttons("print").Image = "print"
    tbExplorer.Buttons("preview").Image = "preview"
    tbExplorer.Buttons("plantilla").Image = "plantilla"

    tbExplorer.Buttons("firstpage").Image = "firstpage"
    tbExplorer.Buttons("previouspage").Image = "previouspage"
    tbExplorer.Buttons("nextpage").Image = "nextpage"
    tbExplorer.Buttons("lastpage").Image = "lastpage"
    
    tbExplorer.Buttons("exp").Image = "exp"
        
    With tbExplorer
        .Buttons("firstpage").Visible = False
        .Buttons("lastpage").Visible = False
        
        .Buttons("previouspage").ToolTipText = "Primer código de artículo menor."
        .Buttons("nextpage").ToolTipText = "Primer código de artículo mayor."
        
        .Buttons("nextpage").Enabled = False
        .Buttons("previouspage").Enabled = False
        
    End With
    
End Sub

Private Sub ButtonRegistros(ByVal bPage As Boolean)
    
    With tbExplorer
        If .Buttons("firstpage").Visible <> bPage Then .Buttons("firstpage").Visible = bPage
        If .Buttons("lastpage").Visible <> bPage Then .Buttons("lastpage").Visible = bPage
        
        If bPage Then
            .Buttons("previouspage").ToolTipText = "Página anterior."
            .Buttons("nextpage").ToolTipText = "Página siguiente."
        Else
            .Buttons("previouspage").ToolTipText = "Primer código de artículo menor."
            .Buttons("nextpage").ToolTipText = "Primer código de artículo mayor."
            .Buttons("previouspage").Enabled = (tArticulo.prm_ArtID > 0)
            .Buttons("nextpage").Enabled = (tArticulo.prm_ArtID > 0)
        End If
        .Refresh
    End With
    
End Sub

Private Sub AccionPreview()
    
    If Not vsListado.Visible Then
        ButtonRegistros True
        AccionImprimir
        
        vsConsulta.Visible = False
        vsLocal.Visible = False
        imgHorizontal.Visible = False
        
        vsListado.Visible = True
        vsListado.ZOrder 0
        tbExplorer.Buttons("preview").Value = tbrPressed
        MnuOpPreview.Checked = True
    Else
        ButtonRegistros False
        vsConsulta.ZOrder 0
        vsLocal.ZOrder 0
        vsListado.Visible = False
        picHorizontal.ZOrder 0
        tbExplorer.Buttons("preview").Value = tbrUnpressed
        MnuOpPreview.Checked = False
        vsConsulta.Visible = True
        vsLocal.Visible = True
        imgHorizontal.Visible = True
    End If
    Me.Refresh
    
End Sub

Private Sub AccionRefrescar()
    If lIDArticulo <> 0 Then BuscoArticuloPorID lIDArticulo, False, False
End Sub

Private Sub MenuPlantilla(ByVal bEnabled As Boolean)
    
    tbExplorer.Buttons("plantilla").Enabled = bEnabled
    MnuPlantillas.Enabled = bEnabled
    
End Sub

Private Function db_ConteoPendiente() As String
On Error GoTo errCP
Dim rsCP As rdoResultset
    db_ConteoPendiente = ""
    Cons = "Select CArEstado From ConteoArticulo" & _
             " Where CArArticulo = " & tArticulo.prm_ArtID & _
             " And CArLocal = " & paCodigoDeSucursal & _
             " And CArEstadoConteo = 0"
    Set rsCP = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsCP.EOF
        db_ConteoPendiente = db_ConteoPendiente & "," & Trim(rsCP!CarEstado)
        rsCP.MoveNext
    Loop
    rsCP.Close
    If Trim(db_ConteoPendiente) <> "" Then db_ConteoPendiente = db_ConteoPendiente & ","
    Exit Function
errCP:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al validar si hay conteo.", Err.Description
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
    MsgBox "Ocurrió un error inesperado al buscar el usuario.", vbCritical, "ATENCIÓN"
End Function

Private Sub db_InsertConteo()
On Error GoTo errIC
Dim lUIDCuenta As Long
    'Busco el usuario en la tabla local.
    If Val(vsLocal.Cell(flexcpData, vsLocal.Row, 0)) = 0 Or Val(vsLocal.Cell(flexcpData, vsLocal.Row, 1)) = 0 Then
        MsgBox "No se cargo correctamente el local ó el estado.", vbCritical, "Atención"
        Exit Sub
    End If
    
    Cons = "Select * From Local Where LocCodigo = " & vsLocal.Cell(flexcpData, vsLocal.Row, 0)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!LocUsuarioContador) Then
            lUIDCuenta = RsAux!LocUsuarioContador
        Else
            MsgBox "El local no tiene un usuario asignado a contar.", vbInformation, "Atención"
        End If
    End If
    RsAux.Close
    
    If lUIDCuenta = 0 Then Exit Sub

    If MsgBox("Se insertará un registro para que se cuente el artículo en el local '" & vsLocal.Cell(flexcpText, vsLocal.Row, 0) & "' para el estado '" & vsLocal.Cell(flexcpText, vsLocal.Row, 1) & "'." & vbCr & vbCr & "¿Confirma almacenar la información?", vbQuestion + vbYesNo, "Conteo") = vbYes Then
        With vsLocal
            Cons = "Select * From ConteoArticulo Where CArLocal = " & .Cell(flexcpData, .Row, 0) & " And CArArticulo = " & tArticulo.prm_ArtID & _
                        " And CarEstado = " & .Cell(flexcpData, .Row, 1) & " And CArEstadoConteo = 0"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            If Not RsAux.EOF Then
                MsgBox "Ya existe un conteo pendiente para ese local.", vbInformation, "Atención"
                RsAux.Close
                Exit Sub
            Else
                RsAux.AddNew
                RsAux!CarLocal = .Cell(flexcpData, .Row, 0)
                RsAux!CarArticulo = tArticulo.prm_ArtID
                RsAux!CarEstado = .Cell(flexcpData, .Row, 1)
                RsAux!CarUsuario = lUIDCuenta
                RsAux!CArFecha = Format(Now, "mm/dd/yyyy hh:nn")
                RsAux!CArDifConLocal = 0
                RsAux!CArQ = 0
                RsAux!CarEstadoConteo = 0
                RsAux.Update
            End If
            RsAux.Close
            MsgBox "Registro insertado.", vbInformation, "Grabar"
        End With
    End If
    Exit Sub
    
errIC:
    clsGeneral.OcurrioError "Error al insertar el conteo.", Err.Description
End Sub

Private Sub vsLocal_Click()
    pnlCampos.Visible = False
End Sub

Private Sub vsLocal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 And Shift = 0 And vsConsulta.Rows > 1 Then
        If vsLocal.Cell(flexcpData, vsLocal.Row, 0) < 0 And vsConsulta.Col = 0 And prmPlUbicoStockCamion > 0 Then EjecutarApp App.Path & "\appExploreMsg.exe ", prmPlUbicoStockCamion & ":" & tArticulo.prm_ArtID & ";;" & Abs(Val(vsLocal.Cell(flexcpData, vsLocal.Row, 0)))
        Exit Sub
    End If
    
    If Button <> 2 Or Shift <> 0 Then Exit Sub
    If vsLocal.Row < 1 Then Exit Sub
    If Val(vsLocal.Cell(flexcpData, vsLocal.Row, 0)) <= 0 Then Exit Sub
    If Not miConexion.AccesoAlMenu("contadorverificador") Then Exit Sub
    PopupMenu MnuPUConteo
End Sub

Private Sub CargoDatosArticuloTrasAccion()
    CargarUltimasAccionesGrabadas
    On Error GoTo errCDA
    Screen.MousePointer = 11
    Cons = "SELECT ArtEnUso, ArtHabilitado, IsNull(ArtEnVentaXMayor, 0) ArtEnVentaXMayor, IsNull(ArtStkMinXMayor, 0) ArtStkMinXMayor, ArtModificado " & _
        "FROM Articulo LEFT OUTER JOIN ArticuloWebPage ON ArtId = AWPArticulo WHERE ArtID = " & tArticulo.prm_ArtID
    Dim rsArt As rdoResultset
    Set rsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsArt.EOF Then
        ArticuloEditado.EnVenta = (LCase(rsArt("ArtHabilitado")) = "S")
        ArticuloEditado.EnUso = rsArt("ArtEnUso")
        ArticuloEditado.AlPorMayor = rsArt("ArtEnVentaXMayor")
        ArticuloEditado.StockMinimoXMayor = rsArt("ArtStkMinXMayor")
        ArticuloEditado.Editado = rsArt("ArtModificado")
        txtQMaxXMayor.Text = ArticuloEditado.AlPorMayor
        If txtStockMinXMayor.Enabled Then txtStockMinXMayor.Text = ArticuloEditado.StockMinimoXMayor
    End If
    rsArt.Close
    Screen.MousePointer = 0
    Exit Sub
errCDA:
    clsGeneral.OcurrioError "Error al cargar los datos del artículo.", Err.Description, "Carga"
    Screen.MousePointer = 0
End Sub

Sub CargoInfoDelArticulo()
Dim sCIA As String
    
    Set ArticuloEditado = New clsArticulo
    sCIA = "SELECT ArtEnUso, Case ArtHabilitado WHEN 'S' THEN 1 ELSE 0 END ArtHabilitado " & _
        ", ArtEnWeb, IsNull(AWPSinPrecio, 0) AWPSinPrecio, IsNull(AWPEnVenta, 0) AWPEnVenta " & _
        ", Case WHEN IsNull(Articulo.ArtEnVentaXMayor, 0) > 0 THEN 1 ELSE 0 END XMayor " & _
        ", IsNull(ArtStkMinXMayor, 0) ArtStkMinXMayor, RTRIM(ArtNombre) Nombre " & _
        "FROM Articulo LEFT OUTER JOIN ArticuloWebPage ON ArtId = AWPArticulo " & _
        "WHERE ArtID = " & tArticulo.prm_ArtID
    Dim rsA As rdoResultset
    Set rsA = cBase.OpenResultset(sCIA, rdOpenDynamic, rdConcurValues)
    If Not rsA.EOF Then
        With ArticuloEditado
            .ID = tArticulo.prm_ArtID
            .Nombre = rsA("Nombre")
            .AlPorMayor = rsA("XMayor")
            .StockMinimoXMayor = rsA("ArtStkMinXMayor")
            .EnUso = rsA("ArtEnUso")
            .EnVenta = rsA("ArtHabilitado")
            .EnWeb = rsA("ArtEnWeb")
            .EnWebConPrecio = Not rsA("AWPSinPrecio")
            .EnVentaWeb = rsA("AWPEnVenta")
         End With
    End If
    rsA.Close
    
    With vsCampos
        .Rows = 0
        .AddItem "En uso"
        .Cell(flexcpText, .Rows - 1, 1) = IIf(ArticuloEditado.EnUso, "Si", "No")
        .AddItem "En venta"
        .Cell(flexcpText, .Rows - 1, 1) = IIf(ArticuloEditado.EnVenta, "Si", "No")
        .AddItem "En web"
        .Cell(flexcpText, .Rows - 1, 1) = IIf(ArticuloEditado.EnWeb, "Si", "No")
        .AddItem "Se vende en la web"
        .Cell(flexcpText, .Rows - 1, 1) = IIf(ArticuloEditado.EnVentaWeb, "Si", "No")
        .AddItem "Mostrar precio en web"
        .Cell(flexcpText, .Rows - 1, 1) = IIf(ArticuloEditado.EnWebConPrecio, "Si", "No")
        .AddItem "Se vende por mayor"
        .Cell(flexcpText, .Rows - 1, 1) = IIf(ArticuloEditado.AlPorMayor >= 1, "Si", "No")
        
        .Cell(flexcpFontBold, 0, 1, .Rows - 1) = True
    End With

End Sub


Private Sub s_GetDataArticulo()

    CargarUltimasAccionesGrabadas
    
    Screen.MousePointer = 11
    
    Set ArticuloEditado = New clsArticulo
    
    ArticuloEditado.EnUso = tArticulo.GetField("ArtEnUso")
    If Not IsNull(tArticulo.GetField("ArtHabilitado")) Then
        If UCase(tArticulo.GetField("ArtHabilitado")) = "S" Then
            ArticuloEditado.EnVenta = True
        End If
    End If
    ArticuloEditado.EnWeb = tArticulo.GetField("ArtEnWeb")
    
    If Not IsNull(tArticulo.GetField("ArtEnVentaXMayor")) Then
        ArticuloEditado.AlPorMayor = tArticulo.GetField("ArtEnVentaXMayor")
        txtQMaxXMayor.Text = ArticuloEditado.AlPorMayor
    End If
    
    If Not IsNull(tArticulo.GetField("ArtStkMinXMayor")) Then
        ArticuloEditado.StockMinimoXMayor = tArticulo.GetField("ArtStkMinXMayor")
        txtStockMinXMayor.Text = ArticuloEditado.StockMinimoXMayor
    End If
    ArticuloEditado.Editado = tArticulo.GetField("ArtModificado")
    
    If Not IsNull(tArticulo.GetField("ArtStockMinimo")) Then tSMin.Text = tArticulo.GetField("ArtStockMinimo"): tSMin.Tag = tSMin.Text
    
    If Not IsNull(tArticulo.GetField("ArtDisponibleDesde")) Then tbDisponibleDesde.Text = Format(tArticulo.GetField("ArtDisponibleDesde"), "dd/mm/yyyy")
    If Not IsNull(tArticulo.GetField("ArtDemoraEntrega")) Then tbDemora.Text = tArticulo.GetField("ArtDemoraEntrega")
    
    CargoInfoDelArticulo
    
    ps_SetValoresStock
    
    EstadoObjetos True
    
    If tArticulo.GetField("ArtSeImporta") Then
        labUltimaCompra.Caption = " Última Compra: " & RetornoUltimaCompra(tArticulo.prm_ArtID)
    Else
        labUltimaCompra.Caption = " Última Compra: " & CompraNoImportacion(tArticulo.prm_ArtID)
    End If
    
    ObtenerDiasVentas
    
    lIDArticulo = tArticulo.prm_ArtID
    tArticulo.Tag = lIDArticulo
    
    ps_LoadArregloStock
    ps_SetPrecioVigente
    
    arr_AddItem tArticulo.prm_ArtID, Trim(tArticulo.Text)
    MenuExplorer
    AccionConsultar
    Screen.MousePointer = 0
    
End Sub

Function ObtenerDiasVentas()
On Error GoTo errDV
    Dim rsDV As rdoResultset
    Set rsDV = cBase.OpenResultset("select dbo.[HayStockParaXDias](" & tArticulo.prm_ArtID & ", 0)", rdOpenDynamic, rdConcurValues)
    If Not rsDV.EOF Then
        If Not IsNull(rsDV(0)) Then lblVtasXDias.Caption = "Días: " & rsDV(0)
    End If
    rsDV.Close
    lblVtasXDias.Refresh
    Exit Function
errDV:
    clsGeneral.OcurrioError "Error al obtener los días de ventas.", Err.Description, "Dias Ventas"
End Function

Private Sub loc_ExpandirLocales()
On Error GoTo errEL
Dim iQ As Integer
    With vsLocal
        For iQ = .FixedRows To .Rows - 2
            If .IsSubtotal(iQ) Then
                If .IsCollapsed(iQ) = flexOutlineCollapsed Then
                    .IsCollapsed(iQ) = flexOutlineExpanded
                End If
            End If
        Next
    End With
errEL:
End Sub

Private Function BuscarAccionDeStock(ByVal idAccion As Integer) As clsAccionesStock
    Set BuscarAccionDeStock = Nothing
    Dim oAccion As clsAccionesStock
    For Each oAccion In colAccionesStock
        If oAccion.ID = cboAcciones.ItemData(cboAcciones.ListIndex) Then
            Set BuscarAccionDeStock = oAccion
            Exit Function
        End If
    Next
End Function

Public Function BitDadaAccion(ByVal idBit As Byte) As Byte
On Error Resume Next
    
    Dim oAccion As clsAccionesStock
    Set oAccion = BuscarAccionDeStock(cboAcciones.ItemData(cboAcciones.ListIndex))
    If oAccion Is Nothing Then Exit Function
    
    If Mid(oAccion.Bits, idBit, 1) = "1" Then
        BitDadaAccion = 1
    ElseIf Mid(oAccion.Bits, idBit, 1) = "0" Then
        BitDadaAccion = 0
    Else
        BitDadaAccion = 2
    End If

End Function

Private Sub DespliegoInfoAccion()
    On Error Resume Next
    Dim sAccion As String, sBits As String
    lblInfoAccion.Caption = ""
    If cboAcciones.ListIndex = -1 Then Exit Sub
    
    '-- Bit 1=En uso, 2=En.Vta Locales, 3=Mostrar en Web, 4=VerPrecio Web, 5=EnVta Web, 6=EnVta XMayor, 8=Aviso Mail, Del 8 al 10 están libres.
    Dim oAccion As clsAccionesStock
    Set oAccion = BuscarAccionDeStock(cboAcciones.ItemData(cboAcciones.ListIndex))
    If oAccion Is Nothing Then Exit Sub
    
    If Mid(oAccion.Bits, 1, 1) = "1" Then
        sAccion = "'En Uso'  "
    ElseIf Mid(oAccion.Bits, 1, 1) = "0" Then
        sAccion = "'Fuera de uso'  "
    End If
    
    If Mid(oAccion.Bits, 2, 1) = "1" Then
        sAccion = sAccion & "'En Venta'  "
    ElseIf Mid(oAccion.Bits, 2, 1) = "0" Then
        sAccion = sAccion & "'No se vende'  "
    End If
    
    If Mid(oAccion.Bits, 3, 1) = "1" Then
        sAccion = sAccion & "'En Web' "
    ElseIf Mid(oAccion.Bits, 3, 1) = "0" Then
        sAccion = sAccion & "'No está en la web'  "
    End If
    
    If Mid(oAccion.Bits, 4, 1) = "1" Then
        sAccion = sAccion & "'Precio en Web'  "
    ElseIf Mid(oAccion.Bits, 4, 1) = "0" Then
        sAccion = sAccion & "'Sin precio en Web'  "
    End If
    
    If Mid(oAccion.Bits, 5, 1) = "1" Then
        sAccion = sAccion & "'Venta en Web' "
    ElseIf Mid(oAccion.Bits, 5, 1) = "0" Then
        sAccion = sAccion & "'No se vende en Web'  "
    End If
    
    If Mid(oAccion.Bits, 6, 1) = "1" Then
        sAccion = sAccion & "'Venta por mayor' "
    ElseIf Mid(oAccion.Bits, 6, 1) = "0" Then
        sAccion = sAccion & "'No se vende al por mayor'    "
    End If
    If Mid(oAccion.Bits, 7, 1) = "1" Then
        sAccion = sAccion & "'Avisa llegada por mail'   "
    ElseIf Mid(oAccion.Bits, 7, 1) = "0" Then
        sAccion = sAccion & "'Anula aviso llegada por mail'   "
    End If
    lblInfoAccion.Caption = sAccion
    
End Sub
