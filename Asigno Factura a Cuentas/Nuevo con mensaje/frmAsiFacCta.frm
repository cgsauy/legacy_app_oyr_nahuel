VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form frmAsiFacCta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Facturas a Cuentas"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAsiFacCta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Datos de la Cuenta"
      ForeColor       =   &H00153E05&
      Height          =   5055
      Left            =   120
      TabIndex        =   27
      Top             =   1920
      Width           =   10575
      Begin VB.CommandButton butRehabilitar 
         Caption         =   "&Rehabilitar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton butGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox tAsigna 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5400
         MaxLength       =   16
         TabIndex        =   19
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox tNumero 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         MaxLength       =   6
         TabIndex        =   17
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox tSerie 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   16
         Top             =   3000
         Width           =   375
      End
      Begin AACombo99.AACombo cSucursal 
         Height          =   315
         Left            =   840
         TabIndex        =   14
         Top             =   3000
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
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsRecibo 
         Height          =   2295
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   4048
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
         BackColorSel    =   7572329
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   4210752
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsFactura 
         Height          =   975
         Left            =   120
         TabIndex        =   20
         Top             =   3360
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   1720
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
         BackColorSel    =   7572329
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
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
      Begin VB.Label lblTitSaldoCVenc 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo con aportes vencidos:"
         Height          =   255
         Left            =   6960
         TabIndex        =   40
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblSaldoConVencido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "1,125,250.00"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   9120
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Salida aprobado:"
         Height          =   255
         Left            =   7680
         TabIndex        =   38
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H002C511E&
         Caption         =   "1,125,250.00"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   9120
         TabIndex        =   37
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblAporteVencido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "1,125,250.00"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   5280
         TabIndex        =   36
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblTituloAporte 
         BackStyle       =   0  'Transparent
         Caption         =   "Aportes vencidos:"
         Height          =   255
         Left            =   3840
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label labSalida 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00738B69&
         Caption         =   "1,125,250.00"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   7080
         TabIndex        =   31
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Salida por asignación:"
         Height          =   255
         Left            =   5280
         TabIndex        =   30
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "&Asigna:"
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         Height          =   255
         Left            =   8520
         TabIndex        =   29
         Top             =   4440
         Width           =   495
      End
      Begin VB.Label labSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H002C511E&
         Caption         =   "1,125,250.00"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   9120
         TabIndex        =   28
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "&Numero:"
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "&Sucursal:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Moneda:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificación de Cuenta"
      ForeColor       =   &H00153E05&
      Height          =   1815
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   10575
      Begin VB.TextBox tNombreColectivo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         MaxLength       =   60
         TabIndex        =   5
         Top             =   240
         Width           =   5655
      End
      Begin AACombo99.AACombo cTipoCuenta 
         Height          =   315
         Left            =   600
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
      Begin VB.TextBox tCodigo 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3360
         MaxLength       =   15
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin MSMask.MaskEdBox tCiCliente1 
         Height          =   285
         Left            =   600
         TabIndex        =   7
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#.###.###-#"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tRUC 
         Height          =   285
         Left            =   2640
         TabIndex        =   9
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###.###.####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentario:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lComentario 
         BackColor       =   &H00D0D8CD&
         Caption         =   "1.939.325-9"
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   1080
         TabIndex        =   32
         Top             =   1320
         Width           =   9375
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&R.U.C.:"
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.Label labCI2 
         BackColor       =   &H00D0D8CD&
         Caption         =   "1.939.325-9"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "C.I.:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   735
      End
      Begin VB.Label labCliente2 
         BackColor       =   &H00D0D8CD&
         Caption         =   "Alberta Justiniana"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   960
         Width           =   8415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&C.I.:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.Label labCliente1 
         BackColor       =   &H00D0D8CD&
         Caption         =   "Alberto Mapache"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   4560
         TabIndex        =   22
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Título:"
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Ti&po:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Có&digo:"
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   7065
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12965
            Key             =   "Msg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6068
            MinWidth        =   6068
            Picture         =   "frmAsiFacCta.frx":0442
            Key             =   "printer"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSPrinter8LibCtl.VSPrinter vspPrinter 
      Height          =   2295
      Left            =   0
      TabIndex        =   34
      Top             =   2400
      Visible         =   0   'False
      Width           =   3495
      _cx             =   6165
      _cy             =   4048
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
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
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
      Zoom            =   8.82352941176471
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
   Begin VSReport8LibCtl.VSReport vsrReport 
      Left            =   -120
      Top             =   6000
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
   Begin VB.Menu MnuPrinter 
      Caption         =   "MnuPrinter"
      Visible         =   0   'False
      Begin VB.Menu MnuPriConfiguracion 
         Caption         =   "Configuración"
      End
      Begin VB.Menu MnuPriDondeImprimo 
         Caption         =   "¿Dónde imprimo?"
      End
   End
End
Attribute VB_Name = "frmAsiFacCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eColGrillaCuenta
    Fecha
    Sucursal
    Documento
    Cliente
    Articulo
    Aporte
    Asignado
End Enum

Private oCnfgPrint As New clsCnfgImpresora

Private vUsrPermisos As Integer
Private vPermisoSaldoVencido As Boolean
Private bSolicitarAutorizacion As Boolean

Private m_Disponibilidad As Long
Private m_Tipo As Integer
Private m_Id As Long
Private m_Documento As Long
Private m_Importe As Currency

Private posMonedaPesos As Integer
Private paBD As String
Private gPathListados As String
Private aMovimientoCaja  As Long

Public Property Let prmTipo(ByVal iTipo As Integer)
On Error Resume Next
    m_Tipo = iTipo
End Property

Public Property Let prmID(ByVal lID As Long)
On Error Resume Next
    m_Id = lID
End Property

Public Property Let prmDocumento(ByVal lDoc As Long)
On Error Resume Next
    m_Documento = lDoc
End Property

Public Property Let prmImporte(ByVal cImp As Currency)
On Error Resume Next
    m_Importe = cImp
End Property

Private Sub butGrabar_Click()
    AccionGrabar
End Sub

Private Sub butRehabilitar_Click()
    RehabilitarRecibo
End Sub

Private Sub cMoneda_Click()
    LimpioCamposCuenta
End Sub

Private Sub cMoneda_Change()
    LimpioCamposCuenta
End Sub

Private Sub cMoneda_GotFocus()
    With cMoneda
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione una moneda."
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cMoneda.ListIndex > -1 Then CargoDatosCuenta cMoneda.ItemData(cMoneda.ListIndex)
    End If
End Sub

Private Sub cMoneda_LostFocus()
    cMoneda.SelStart = 0
    Ayuda ""
End Sub

Private Sub cSucursal_GotFocus()
    With cSucursal
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda " Seleccione la sucursal en donde se facturo el contado."
End Sub

Private Sub cSucursal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cSucursal.ListIndex > -1 Then
            Foco tSerie
        Else
            If butGrabar.Enabled Then AccionGrabar
        End If
    End If
End Sub

Private Sub cSucursal_LostFocus()
    Ayuda ""
    cSucursal.SelStart = 0
End Sub

Private Sub cTipoCuenta_Click()
    LimpioCamposInformacion
    OcultoCamposInformacion
    HabilitoCamposInformacion
    LimpioCamposCuenta
    OcultoCamposCuenta
End Sub

Private Sub cTipoCuenta_Change()
    FormularioACero
End Sub

Private Sub cTipoCuenta_GotFocus()
On Error Resume Next
    With cTipoCuenta
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda " Seleccione el tipo de cuenta."
End Sub

Private Sub cTipoCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        HabilitoCamposInformacion
        If cTipoCuenta.ListIndex > -1 Then
            If cTipoCuenta.ItemData(cTipoCuenta.ListIndex) = Cuenta.Colectivo Then
                Foco tCodigo
            Else
                Foco tCiCliente1
            End If
        End If
    End If
End Sub

Private Sub cTipoCuenta_LostFocus()
On Error Resume Next
    Ayuda ""
    With cTipoCuenta: .SelStart = 0: End With
End Sub

Private Sub Form_Activate()
On Error Resume Next
    
    If m_Tipo > 0 Then
        
        If m_Tipo = Cuenta.Colectivo Then
            
            cTipoCuenta.ListIndex = 0
            tCodigo.Text = m_Id
            Call tCodigo_KeyPress(13)
            
            If m_Documento > 0 Then BuscoFacturaPorID m_Documento
            'Si me paso el documento lo cargo.
            If m_Importe > 0 Then tAsigna.Text = Format(m_Importe, "#,##0.00")
                
        ElseIf m_Tipo = Cuenta.Personal Then
            
            cTipoCuenta.ListIndex = 1
            BuscoClientePorID m_Id
            HabilitoCamposCuenta
            BuscoSiHayParaPesos
            
            If m_Documento > 0 Then BuscoFacturaPorID m_Documento
            'Si me paso el documento lo cargo.
            If m_Importe > 0 Then tAsigna.Text = Format(m_Importe, "#,##0.00")
            
        End If
        m_Tipo = 0
    End If
    Screen.MousePointer = 0
    Me.Refresh
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift <> vbCtrlMask Then Exit Sub
    Select Case KeyCode
        Case vbKeyX: Unload Me
        Case vbKeyG: If butGrabar.Enabled Then AccionGrabar
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    
    lComentario.FontBold = True
    
    ChDir App.Path
    ChDir "..": gPathListados = CurDir & "\Reportes\"
    paBD = PropiedadesConnect(txtConexion, Database:=True)
    
    'Cargo los tipos de cuentas
    cTipoCuenta.Clear
    cTipoCuenta.AddItem "Colectivo"
    cTipoCuenta.ItemData(cTipoCuenta.NewIndex) = Cuenta.Colectivo
    cTipoCuenta.AddItem "Personal"
    cTipoCuenta.ItemData(cTipoCuenta.NewIndex) = Cuenta.Personal
    
    Cons = "Select MonCodigo, MonSigno From Moneda Where MonFactura = 1 Order by MonSigno"
    CargoCombo Cons, cMoneda
        
    oCnfgPrint.CargarConfiguracion cnfgAppNombreMovimientoCaja, cnfgKeyTicketMovimientoCaja
    
    'Me quedo con la posición de la moneda pesos.
    posMonedaPesos = -1
    For I = 0 To cMoneda.ListCount - 1
        If paMonedaPesos = cMoneda.ItemData(I) Then posMonedaPesos = I: Exit For
    Next I
    
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Where SucDContado <> Null Or SucDRecibo <> Null Order by SucAbreviacion"
    CargoCombo Cons, cSucursal
    
    OcultoCamposInformacion
    LimpioCamposInformacion
    LimpioCamposCuenta
    OcultoCamposCuenta
    InicializoGrillas
   
'    tooMenu.Buttons("inhabilitados").Enabled = False
'    tooMenu.Buttons("autorizar").Enabled = False
   
    
    If oCnfgPrint.Opcion = 0 Then
        Status.Panels("printer") = paPrintConfD
    Else
        Status.Panels("printer") = oCnfgPrint.ImpresoraTickets
    End If
    
    'Abro el Engine del Crystal
    If crAbroEngine = 0 Then MsgBox Trim(crMsgErr), vbCritical, "ATENCIÓN"
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Ayuda ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    crCierroEngine
    CierroConexion
    Set clsGeneral = Nothing
    End
End Sub

Private Sub Label1_Click()
    Foco tCiCliente1
End Sub
Private Sub Label10_Click()
    Foco cSucursal
End Sub
Private Sub Label13_Click()
    Foco tSerie
End Sub

Private Sub Label3_Click()
    Foco tNombreColectivo
End Sub

Private Sub Label5_DblClick()
On Error Resume Next
    If oCnfgPrint.Opcion = 0 Then
        ImprimoSalidaCaja 385702
    Else
        ImprimoSalidaCajaTicket 385702
    End If
End Sub

Private Sub Label6_Click()
    Foco tCodigo
End Sub

'------------------------------------------------------------
'Limpio los objetos de pantalla.
'------------------------------------------------------------
Private Sub LimpioCamposInformacion()
    tCodigo.Text = ""
    tCiCliente1.Text = "": labCliente1.Caption = "": tCiCliente1.Tag = ""
    tRUC.Text = ""
    labCI2.Caption = "": labCliente2.Caption = ""
    tNombreColectivo.Text = ""
    lComentario.Caption = ""
End Sub

'------------------------------------------------------------
'No dejo acceder a los datos.
'------------------------------------------------------------
Sub OcultoCamposInformacion()
    tCodigo.Enabled = False: tCodigo.BackColor = Inactivo
    tCiCliente1.Enabled = False
    tRUC.Enabled = False
    tNombreColectivo.Enabled = False: tNombreColectivo.BackColor = Inactivo
End Sub
'------------------------------------------------------------
'Dejo libre los objetos.
'------------------------------------------------------------
Sub HabilitoCamposInformacion()
    
    If cTipoCuenta.ListIndex = -1 Then Exit Sub
    
    If cTipoCuenta.ItemData(cTipoCuenta.ListIndex) = Cuenta.Colectivo Then
        tCodigo.Enabled = True: tCodigo.BackColor = vbWhite
        tNombreColectivo.Enabled = True: tNombreColectivo.BackColor = vbWhite
    Else
        tRUC.Enabled = True
    End If
    tCiCliente1.Enabled = True: tCiCliente1.BackColor = vbWhite
    
End Sub

'------------------------------------------------------------
'Presiono boton grabar, la misma puede ser un alta o una modificación.
'------------------------------------------------------------
Sub AccionGrabar()
    
    Ayuda "Grabando........."
    If ControlesGrabar Then
        If MsgBox("¿Confirma la asignación de facturas?", vbQuestion + vbYesNo, "GRABAR") = vbYes Then
            If bSolicitarAutorizacion Then
                Dim frmAut As New frmSucesoAutoriza
                With frmAut
                    .ImporteAAutorizar = CCur(labSalida.Caption) - CCur(labSaldo.Caption)
                    .ImporteTotalAAsignar = CCur(labSalida.Caption)
                    If cTipoCuenta.ItemData(cTipoCuenta.ListIndex) = Cuenta.Colectivo Then
                        .Cliente = tCodigo.Tag
                    Else
                        .Cliente = tCiCliente1.Tag
                    End If
                    .Moneda = cMoneda.Text
                    .Show vbModal, Me
                    If .CodigoSuceso = 0 Then
                        Set frmAut = Nothing
                        Exit Sub
                    End If
                End With
                Set frmAut = Nothing
            End If
            
            'Valido los saldos no sea que me los tomaron en otro pc.
            Dim SaldoTotal As Currency
            SaldoTotal = SaldoCuentaPersonal(False)
            If SaldoTotal <> CCur(labSaldo.Caption) Then
                MsgBox "Otro usuario utilizó aportes y altero el saldo, debe iniciar el proceso.", vbExclamation, "ATENCIÓN"
                Screen.MousePointer = 0
                Exit Sub
            End If
            GraboAsignacion
        End If
    End If
    Ayuda ""
    
End Sub
Private Sub GraboAsignacion()
Dim strDocumento As String

    Screen.MousePointer = 11
    FechaDelServidor
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrResumir
    strDocumento = ""
    
    With vsFactura
    
        Dim oMovCaja As New clsCobrarConQuePaga
        Set oMovCaja.Conexion = cBase
        
    
        For I = 1 To .Rows - 1
            If .Cell(flexcpData, I, 1) = 2 Then     'Es para asignar.
            
                Dim oAporte As New clsAporteACuenta
                oAporte.MovimientoAporte = TipoMovimientosAportes.Asignado
                oAporte.TipoCuenta = cTipoCuenta.ItemData(cTipoCuenta.ListIndex)
                If cTipoCuenta.ItemData(cTipoCuenta.ListIndex) = Cuenta.Colectivo Then
                    oAporte.IDCuenta = tCodigo.Tag
                Else
                    oAporte.IDCuenta = tCiCliente1.Tag
                End If
                oAporte.Documento = .Cell(flexcpData, I, 0)
                oAporte.Importe = CCur(.Cell(flexcpText, I, 5))
                If I > 1 Then
                    'Si inserto a nivel de proceso salta por clave duplicada
                    'Hago pausa de unas milesimas.
                    PausaClaveDuplicada
                End If
                oAporte.InsertarAsignacion cBase

                If strDocumento = "" Then
                    strDocumento = Trim(.Cell(flexcpText, I, 1))
                Else
                    strDocumento = strDocumento & ", " & Trim(.Cell(flexcpText, I, 1))
                End If
                
                Dim oCQP As New clsDesicionConQuePaga
                oCQP.InsertoRelacionConQueCobra ContadooCuota, oAporte.Documento, AsignoAporteCta, oAporte.IDCuenta
                Set oCQP = Nothing
                
            End If
        Next I
        
        Dim oMC As New clsDesicionConQuePaga
        oMC.Sucursal = paCodigoDeSucursal
        aMovimientoCaja = oMC.GrabarMovimientoDeCaja(CCur(labSalida.Caption), strDocumento)
        Set oMovCaja = Nothing
        
        cBase.CommitTrans
        
        If oCnfgPrint.Opcion = 0 Then
            ImprimoSalidaCaja aMovimientoCaja
        Else
            ImprimoSalidaCajaTicket aMovimientoCaja
        End If
        
    End With
    CargoDatosCuenta cMoneda.ItemData(cMoneda.ListIndex)
    Screen.MousePointer = 0
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "Error al iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrResumir:
    Resume ErrFT
ErrFT:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar almacenar la información.", Err.Description
End Sub

Sub PausaClaveDuplicada()
Dim Inicio As Date
    
    Inicio = Now
    Do 'While DateDiff("s", Inicio, Now) > 1
        DoEvents
    Loop Until DateDiff("s", Inicio, Now) > 0.5
    
End Sub


Private Sub MnuPriConfiguracion_Click()
    prj_GetPrinter True
    If oCnfgPrint.Opcion = 0 Then
        Status.Panels("printer") = paPrintConfD
    Else
        Status.Panels("printer") = oCnfgPrint.ImpresoraTickets
    End If
End Sub

Private Sub MnuPriDondeImprimo_Click()
On Error Resume Next
    frmDondeImprimoSC.Show vbModal
    oCnfgPrint.CargarConfiguracion cnfgAppNombreMovimientoCaja, cnfgKeyTicketMovimientoCaja
End Sub

Private Sub Status_PanelClick(ByVal Panel As ComctlLib.Panel)
    If "printer" = Panel.Key Then
        PopupMenu MnuPrinter
    End If
End Sub

Private Sub Status_PanelDblClick(ByVal Panel As ComctlLib.Panel)
    If "printer" = Panel.Key Then
        prj_GetPrinter True
        If oCnfgPrint.Opcion = 0 Then
            Status.Panels("printer") = paPrintConfD
        Else
            Status.Panels("printer") = oCnfgPrint.ImpresoraTickets
        End If
    End If
End Sub

Private Sub tAsigna_GotFocus()
    With tAsigna
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda " Ingrese el importe a asignar del contado."
End Sub

Private Sub tAsigna_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Val(tNumero.Tag) > 0 And IsNumeric(tAsigna.Text) Then
            If CCur(tAsigna.Text) > CCur(tAsigna.Tag) Then
                MsgBox "No puede asignar un valor mayor al total de la factura.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            If CCur(tAsigna.Text) <= 0 Then
                MsgBox "Ingrese un importe mayor que cero.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If

            'ccur(labSaldo.Caption) Tengo el total habilitado
            'CCur(labSaldo.Tag) Tengo el total de la cuenta.
            Dim SaldoPermitido As Currency
            
            vPermisoSaldoVencido = miconexion.AccesoAlMenu("Asigno Factura Saldo Vencido")
            SaldoPermitido = IIf(vPermisoSaldoVencido, CCur(lblSaldoConVencido.Caption), CCur(labSaldo.Caption))
            
            If CCur(labSalida.Caption) + CCur(tAsigna.Text) > SaldoPermitido Then
                
                '(And Not vPermisoSaldoVencido)  no lo agrego ya que no es necesario.
                If CCur(labSalida.Caption) + CCur(tAsigna.Text) <= CCur(lblSaldoConVencido.Caption) And Not bSolicitarAutorizacion Then
                    Dim ret As VbMsgBoxResult
                    Do
                        ret = MsgBox("Debe solicitar autorización para utilizar aportes vencidos." & vbCrLf & vbCrLf & "¿Al grabar desea solicitar la autorización?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Utilizar aportes vencidos")
                    Loop Until ret <> vbCancel
                    
                    'Si es un usuario que no posee permisos entonces debe pedir autorización.
                    If ret = vbYes Then
                        bSolicitarAutorizacion = True
                    Else
                        Exit Sub
                    End If
                    
                Else
                    
                    MsgBox "Está asignando un importe superior al saldo disponible, verifique.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                    
                End If
                
            Else
            
                'Si es usuario admin le aviso si excede el saldo permitido.
                If CCur(labSalida.Caption) + CCur(tAsigna.Text) > CCur(labSaldo.Caption) Then
                    MsgBox "Está asignando un importe que utiliza aportes vencidos.", vbInformation, "ATENCIÓN"
                End If
                
            End If
            
            CargoFacturaEnGrilla tNumero.Tag
        End If
    End If
    
End Sub

Private Sub tCiCliente1_Change()
    
    If Val(tCiCliente1.Tag) > 0 Then tCiCliente1.Tag = "0": FormularioACero
    tCiCliente1.Tag = "0": labCliente1.Caption = "": labCliente1.Tag = ""
    On Error Resume Next
    If Me.ActiveControl.Name = tCiCliente1.Name Then tCiCliente1.SetFocus
    
End Sub
Private Sub tCiCliente1_GotFocus()
    tCiCliente1.SelStart = 0: tCiCliente1.SelLength = 11
    Ayuda "Ingrese la cédula de uno de los integrantes del colectivo."
End Sub
Private Sub tCiCliente1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            
            If tCiCliente1.Tag <> "0" Then cMoneda.SetFocus: Exit Sub
            
            tCiCliente1.Tag = "0": labCliente1.Caption = "": labCliente1.Tag = ""
            If tCiCliente1.Text <> "" Then
                If clsGeneral.CedulaValida(tCiCliente1.Text) Then
                    BuscoClientePorCedula tCiCliente1
                    If Val(tCiCliente1.Tag) > 0 Then
                        If cTipoCuenta.ItemData(cTipoCuenta.ListIndex) = Cuenta.Colectivo Then
                            BuscoColectivoPorCliente tCiCliente1.Tag
                        Else
                            HabilitoCamposCuenta
                            BuscoSiHayParaPesos
                        End If
                    End If
                Else
                    MsgBox "La cédula ingresada no es válida.", vbExclamation, "ATENCIÓN"
                End If
            Else
                If tRUC.Enabled Then Foco tRUC
            End If
            
        Case vbKeyF4: AccionBuscarCliente True, False
    End Select
End Sub
Private Sub tCiCliente1_LostFocus()
    tCiCliente1.SelStart = 0
    Ayuda ""
End Sub

Private Sub tCodigo_Change()
    If Val(tCiCliente1.Tag) > 0 Or Val(tCodigo.Tag) > 0 Then FormularioACero
End Sub

Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese el código del colectivo."
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
On Error GoTo ErrBC
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCodigo.Text) Then
            BuscoColectivoPorCodigo tCodigo.Text
        ElseIf Trim(tCodigo.Text) <> "" Then
            LimpioCamposInformacion
            MsgBox "El código ingresado no es válido.", vbExclamation, "ATENCIÓN"
        Else
            Foco tNombreColectivo
        End If
    End If
    Exit Sub
ErrBC:
    clsGeneral.OcurrioError "Error al buscar el colectivo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tCodigo_LostFocus()
    Ayuda ""
    tCodigo.SelStart = 0
End Sub


Private Sub tNombreColectivo_GotFocus()
    With tNombreColectivo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese el título del colectivo a buscar."
End Sub
Private Sub tNombreColectivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(tNombreColectivo.Text) <> "" Then BuscoColectivoPorTitulo tNombreColectivo Else Foco tCiCliente1
    End If
End Sub
Private Sub tNombreColectivo_LostFocus()
    Ayuda ""
    tNombreColectivo.SelStart = 0
End Sub

Private Sub tNumero_GotFocus()
    With tNumero
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda " Ingrese el número del contado."
End Sub
Private Sub tNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cSucursal.ListIndex > -1 And Trim(tSerie.Text) <> "" And IsNumeric(tNumero.Text) Then
            If vsRecibo.Rows > 1 Then BuscoFactura Else MsgBox "No hay recibos con aportes.", vbExclamation, "ATENCIÖN"
        Else
            MsgBox "Los datos de búsqueda no son válidos.", vbExclamation, "ATENCIÓN"
        End If
    End If
End Sub
Private Sub tNumero_LostFocus()
    tNumero.SelStart = 0
    Ayuda ""
End Sub
Private Sub BuscoFactura()
On Error GoTo ErrBF
    
    Screen.MousePointer = 11
    
    Cons = "Select * From Documento, Cliente" _
        & " Where DocTipo IN(" & TipoDocumento.Contado & ", " & TipoDocumento.ReciboDePago & ")" _
        & " And DocSerie = '" & tSerie.Text & "' And DocNumero = " & tNumero.Text _
        & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex) _
        & " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
        & " And DocCliente = CliCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        
        RsAux.Close
        MsgBox "No se encontró un contado con esas características, o la moneda no corresponde.", vbExclamation, "ATENCIÓN"
        
    Else
        
        'Veo si pertenece al cliente.
        If cTipoCuenta.ItemData(cTipoCuenta.ListIndex) = Cuenta.Colectivo Then
            If Not (RsAux!DocCliente = Val(labCI2.Tag) Or RsAux!DocCliente = Val(tCiCliente1.Tag)) Then
                If MsgBox("Ese documento no pertenece a los clientes del colectivo." & Chr(13) & "¿Desea ingresarlo de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                    RsAux.Close
                    Screen.MousePointer = 0
                    cSucursal.Text = "": tSerie.Text = "": tNumero.Text = "": tNumero.Tag = "0": tAsigna.Tag = "0": cSucursal.Tag = 0
                    Foco cSucursal
                    Exit Sub
                End If
            End If
        Else
            If RsAux!DocCliente <> Val(tCiCliente1.Tag) Then
                If MsgBox("Ese documento no pertenece a la cuenta personal seleccionada." & Chr(13) & "¿Desea ingresarlo de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                    RsAux.Close
                    Screen.MousePointer = 0
                    cSucursal.Text = "": tSerie.Text = "": tNumero.Text = "": tNumero.Tag = "0": tAsigna.Tag = "0": cSucursal.Tag = 0
                    Foco cSucursal
                    Exit Sub
                End If
            End If
        End If
        
        If RsAux!DocAnulado = 1 Then
            RsAux.Close
            MsgBox "El documento seleccionado está anulado, verifique.", vbExclamation, "ATENCIÓN"
            cSucursal.Text = "": tSerie.Text = "": tNumero.Text = "": tNumero.Tag = "0": tAsigna.Tag = "0": cSucursal.Tag = 0
            Foco cSucursal
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        cSucursal.Tag = RsAux!DocTipo
        tNumero.Tag = RsAux!DocCodigo
        tAsigna.Tag = RsAux!DocTotal
        tAsigna.Text = Format(RsAux!DocTotal, FormatoMonedaP)
        Foco tAsigna
        RsAux.Close
        
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrBF:
    clsGeneral.OcurrioError "Error al buscar el contado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub BuscoFacturaPorID(ByVal lID As Long)
On Error GoTo ErrBF
    
    Screen.MousePointer = 11
    
    Cons = "Select * From Documento, Cliente" _
        & " Where DocCodigo = " & lID _
        & " And DocTipo IN(" & TipoDocumento.Contado & ", " & TipoDocumento.ReciboDePago & ")" _
        & " And DocCliente = CliCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        
        RsAux.Close
        MsgBox "No se encontró un contado con esas características, o la moneda no corresponde.", vbExclamation, "ATENCIÓN"
        
    Else
        
        'Veo si pertenece al cliente.
        If cTipoCuenta.ItemData(cTipoCuenta.ListIndex) = Cuenta.Colectivo Then
            If Not (RsAux!DocCliente = Val(labCI2.Tag) Or RsAux!DocCliente = Val(tCiCliente1.Tag)) Then
                If MsgBox("Ese documento no pertenece a los clientes del colectivo." & Chr(13) & "¿Desea ingresarlo de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                    RsAux.Close
                    Screen.MousePointer = 0
                    cSucursal.Text = "": tSerie.Text = "": tNumero.Text = "": tNumero.Tag = "0": tAsigna.Tag = "0": cSucursal.Tag = 0
                    Foco cSucursal
                    Exit Sub
                End If
            End If
        Else
            If RsAux!DocCliente <> Val(tCiCliente1.Tag) Then
                If MsgBox("Ese documento no pertenece a la cuenta personal seleccionada." & Chr(13) & "¿Desea ingresarlo de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                    RsAux.Close
                    Screen.MousePointer = 0
                    cSucursal.Text = "": tSerie.Text = "": tNumero.Text = "": tNumero.Tag = "0": tAsigna.Tag = "0": cSucursal.Tag = 0
                    Foco cSucursal
                    Exit Sub
                End If
            End If
        End If
        
        If RsAux!DocAnulado = 1 Then
            RsAux.Close
            MsgBox "El documento seleccionado está anulado, verifique.", vbExclamation, "ATENCIÓN"
            cSucursal.Text = "": tSerie.Text = "": tNumero.Text = "": tNumero.Tag = "0": tAsigna.Tag = "0": cSucursal.Tag = 0
            Foco cSucursal
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        'Busco en el combo la sucursal.
        BuscoCodigoEnCombo cSucursal, RsAux!DocSucursal
        tSerie.Text = Trim(RsAux!DocSerie)
        tNumero.Text = Trim(RsAux!DocNumero)
        
        cSucursal.Tag = RsAux!DocTipo
        tNumero.Tag = RsAux!DocCodigo
        tAsigna.Tag = RsAux!DocTotal
        tAsigna.Text = Format(RsAux!DocTotal, FormatoMonedaP)
        RsAux.Close
        Foco tAsigna
        
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrBF:
    clsGeneral.OcurrioError "Error al buscar el contado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoFacturaEnGrilla(DocCodigo As Long)
Dim Codigo As Long
Dim Rs As rdoResultset
On Error GoTo ErrCF
    
    Screen.MousePointer = 11
    
    With vsFactura
        For I = 1 To .Rows - 1
            If .Cell(flexcpData, I, 0) = DocCodigo Then
                MsgBox "Este documento ya está en la lista, verifique.", vbExclamation, "ATENCIÓN"
                Screen.MousePointer = 0:  Exit Sub
            End If
        Next I
    End With
    
    If Val(cSucursal.Tag) = TipoDocumento.Contado Then
        Cons = "Select * From Documento, Renglon, Articulo, Cliente" _
            & " Where DocCodigo = " & DocCodigo _
            & " And DocCodigo = RenDocumento And RenArticulo = ArtID And DocCliente = CliCodigo"
    Else
        Cons = "Select * From Documento, Cliente" _
            & " Where DocCodigo = " & DocCodigo & " And DocCliente = CliCodigo"
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Codigo = 0
    
    Cons = "Select SUM(AaCImporte) From AporteACuenta WHERE AaCTipo = " & TipoMovimientosAportes.Asignado & " AND AaCDocumento = " & DocCodigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not Rs.EOF Then
        If Not IsNull(Rs(0)) Then
            If CCur(tAsigna.Tag) - CCur(Rs(0)) < CCur(tAsigna.Text) Then
                Screen.MousePointer = 0
                MsgBox "Este documento está asignado por el importe de " & Format(Rs(0), FormatoMonedaP) & ", verifique.", vbExclamation, "ATENCIÓN"
                Rs.Close
                Exit Sub
            End If
        End If
    End If
    Rs.Close
    
    Cons = "Select SUM(AaCImporte) From AporteACuenta WHERE AaCTipo = " & TipoMovimientosAportes.Aporte & " AND AaCDocumento = " & DocCodigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not Rs.EOF Then
        If Not IsNull(Rs(0)) Then
            Rs.Close: Screen.MousePointer = 0
            MsgBox "Este recibo es de seña, no podrá asignarlo.", vbExclamation, "ATENCIÓN": Exit Sub
        End If
    End If
    Rs.Close
    
    Do While Not RsAux.EOF
        With vsFactura
            If Codigo <> RsAux!DocCodigo Then
                Codigo = RsAux!DocCodigo
                .AddItem ""
                .Cell(flexcpData, .Rows - 1, 0) = Codigo
                .Cell(flexcpData, .Rows - 1, 1) = 2     'Es Nueva, si es uno la cargo por ya asignadas
                .Cell(flexcpText, .Rows - 1, 0) = Trim(cSucursal.Text)
                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!DocSerie) & " " & Trim(RsAux!DocNumero)
                If Not IsNull(RsAux!CLiCIRUC) Then
                    If RsAux!CliTipo = TipoCliente.Cliente Then
                        .Cell(flexcpText, .Rows - 1, 2) = clsGeneral.RetornoFormatoCedula(RsAux!CLiCIRUC)
                    Else
                        .Cell(flexcpText, .Rows - 1, 2) = clsGeneral.RetornoFormatoRuc(RsAux!CLiCIRUC)
                    End If
                End If
                If Val(cSucursal.Tag) = TipoDocumento.Contado Then .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!DocTotal, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 5) = Format(tAsigna.Text, FormatoMonedaP)
                .Cell(flexcpData, .Rows - 1, 5) = CCur(tAsigna.Tag)
'                labAsignado.Caption = Format(CCur(labAsignado.Caption) + tAsigna.Text, FormatoMonedaP)
'                labSaldo.Caption = Format(CCur(labAporte.Caption) - CCur(labAsignado.Caption), FormatoMonedaP)
                labSalida.Caption = Format(CCur(labSalida.Caption) + CCur(tAsigna.Text), FormatoMonedaP)
            Else
                .Cell(flexcpText, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 3) & ", " & Trim(RsAux!ArtNombre)
            End If
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If vsFactura.Rows > 1 Then
        butGrabar.Enabled = True
        If CCur(labSaldo.Caption) < 0 Then MsgBox "El saldo da negativo, verifique.", vbExclamation, "ATENCIÓN"
    End If
    cSucursal.Text = "": tSerie.Text = "": tNumero.Text = "": tNumero.Tag = "0": tAsigna.Text = "": tAsigna.Tag = "0"
    Foco cSucursal
    Screen.MousePointer = 0
    Exit Sub
ErrCF:
    clsGeneral.OcurrioError "Error al buscar el contado.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoClientePorCedula(strCedula As String)
On Error GoTo ErrBC
    Screen.MousePointer = 11
    
    Cons = "Select CliCodigo, CliDireccion, CPeApellido1, Nombre = RTRIM(RTRIM(CPeApellido1) + ' ' + RTRIM(CPeApellido2)) + ', ' +  RTRIM(RTRIM(CPeNombre1) + ' ' + RTRIM(CPeNombre2)) " _
        & " From Cliente, CPersona " _
        & " Where CLiCIRUC = '" & strCedula & "'" _
        & " And CliCodigo = CPeCliente "
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un cliente ingresado con esa cédula.", vbInformation, "ATENCIÓN"
        tCiCliente1.Tag = "0"
        labCliente1.Caption = " "
        labCliente1.Tag = ""
    Else
        tCiCliente1.Tag = RsAux!CliCodigo
        labCliente1.Caption = " " & Trim(RsAux!Nombre)
        labCliente1.Tag = Trim(RsAux!CPeApellido1)
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBC:
    clsGeneral.OcurrioError "Error al buscar al cliente por cédula.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub BuscoClientePorID(ByVal lID As String)
On Error GoTo ErrBC
    Screen.MousePointer = 11
    
    Cons = "Select * " _
        & " From Cliente " _
            & " Left Outer Join CPersona On CliCodigo = CPeCliente " _
            & " Left Outer Join CEmpresa On CliCodigo = CEmCliente " _
        & " Where CLiCodigo = " & lID
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If RsAux.EOF Then
        RsAux.Close
        tCiCliente1.Tag = "0"
        labCliente1.Caption = " "
        labCliente1.Tag = ""
    Else
        If RsAux!CliTipo = TipoCliente.Cliente Then
            If Not IsNull(RsAux!CLiCIRUC) Then tCiCliente1.Text = clsGeneral.RetornoFormatoCedula(RsAux!CLiCIRUC)
            labCliente1.Caption = " " & RTrim(RsAux!CPeApellido1)
            If Not IsNull(RsAux!CPeApellido2) Then
                labCliente1.Caption = labCliente1.Caption & " " & RTrim(RsAux!CPeApellido2)
            End If
            labCliente1.Caption = labCliente1.Caption & ", " + RTrim(RsAux!CPeNombre1)
            If Not IsNull(RsAux!CPeNombre2) Then
                labCliente1.Caption = labCliente1.Caption & " " & RTrim(RsAux!CPeNombre2)
            End If
            labCliente1.Tag = Trim(RsAux!CPeApellido1)
        Else
            If Not IsNull(RsAux!CLiCIRUC) Then tRUC.Text = clsGeneral.RetornoFormatoRuc(RsAux!CLiCIRUC)
            labCliente1.Caption = " " & Trim(RsAux!CEmFantasia)
        End If
        tCiCliente1.Tag = RsAux!CliCodigo
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBC:
    clsGeneral.OcurrioError "Error al buscar al cliente por cédula.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ControlesGrabar() As Boolean
    ControlesGrabar = False
    If vsRecibo.Rows = 1 Then Exit Function
    If vsFactura.Rows = 1 Then
        MsgBox "No hay facturas ingresadas.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    If CCur(labSalida.Caption) = 0 Then
        MsgBox "No hay facturas ingresadas.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    If CCur(labSaldo.Caption) < 0 Then
        If MsgBox("El saldo da negativo. ¿Desea almacenar la información ingresada de todas formas?", vbQuestion + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbNo Then Exit Function
    End If
    'Busco si la moneda tiene disponibilidad para esta sucursal.
    m_Disponibilidad = modMaeDisponibilidad.dis_DisponibilidadPara(paCodigoDeSucursal, cMoneda.ItemData(cMoneda.ListIndex))
    If m_Disponibilidad = 0 Then
        MsgBox "No hay una disponibilidad para hacer los movimientos de caja en la moneda seleccionada." & vbCrLf & "Consulte al Administrador.", vbQuestion, "Falta Disponibilidad"
        Exit Function
    End If
    
    
    ControlesGrabar = True
End Function
Private Sub BuscoColectivoPorCodigo(IDColectivo As Long)
    
    Cons = "Select Colectivo.*,  Dir1 = C1.CliDireccion , Dir2 = C2.CliDireccion, Ced1 = C1.CliCIRUC, Ced2 = C2.CliCIRUC, Nom1 = RTRIM(RTRIM(CP1.CPeApellido1) + ' ' + RTRIM(CP1.CPeApellido2)) + ', ' +  RTRIM(RTRIM(CP1.CPeNombre1) + ' ' + RTRIM(CP1.CPeNombre2)) " _
                        & " , Nom2 = RTRIM(RTRIM(CP2.CPeApellido1) + ' ' + RTRIM(CP2.CPeApellido2)) + ', ' +  RTRIM(RTRIM(CP2.CPeNombre1) + ' ' + RTRIM(CP2.CPeNombre2)) " _
        & " From Colectivo, Cliente C1, CPersona CP1, Cliente C2, CPersona CP2 " _
        & " Where ColCodigo = " & IDColectivo _
        & " And ColCliente1 = C1.CliCodigo And C1.CliCodigo = CP1.CPeCliente " _
        & " And ColCliente2 = C2.CliCodigo And C2.CliCodigo = CP2.CPeCliente  And ColCerrado = 0"
    
    CargoCamposColectivo
End Sub
Private Sub CargoCamposColectivo()
On Error GoTo ErrBC
    Screen.MousePointer = 11
    LimpioCamposInformacion
    LimpioCamposCuenta
    OcultoCamposCuenta
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un colectivo con esas características, o el mismo esta cerrado.", vbExclamation, "ATENCIÓN"
    Else
        tCodigo.Text = RsAux!ColCodigo
        tCodigo.Tag = RsAux!ColCodigo
        If Not IsNull(RsAux!Ced1) Then tCiCliente1.Text = RsAux!Ced1
        tCiCliente1.Tag = RsAux!ColCliente1
        labCliente1.Caption = " " & Trim(RsAux!Nom1)
        If Not IsNull(RsAux!Ced2) Then labCI2.Caption = clsGeneral.RetornoFormatoCedula(RsAux!Ced2)
        labCI2.Tag = RsAux!ColCliente2
        labCliente2.Caption = " " & Trim(RsAux!Nom2)
        If Not IsNull(RsAux!ColNombre) Then tNombreColectivo.Text = Trim(RsAux!ColNombre)
        If Not IsNull(RsAux!ColComentario) Then lComentario.Caption = Trim(RsAux!ColComentario)
        RsAux.Close
        HabilitoCamposCuenta
        BuscoSiHayParaPesos
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBC:
    clsGeneral.OcurrioError "Error al buscar el colectivo.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoColectivoPorTitulo(strNombre As String)
On Error GoTo ErrBCPT
Dim lID As Long
Dim LiAyuda As New clsListadeAyuda

    Cons = "Select ColCodigo, Colectivo = ColCodigo,  'Cédula Cliente 1' = C1.CliCIRUC,  'Nombre Cliente 1' = RTRIM(RTRIM(CP1.CPeApellido1) + ' ' + RTRIM(CP1.CPeApellido2)) + ', ' +  RTRIM(RTRIM(CP1.CPeNombre1) + ' ' + RTRIM(CP1.CPeNombre2)) " _
                    & ", 'Cédula Cliente 2' = C2.CliCIRUC, 'Nombre Cliente 2' = RTRIM(RTRIM(CP2.CPeApellido1) + ' ' + RTRIM(CP2.CPeApellido2)) + ', ' +  RTRIM(RTRIM(CP2.CPeNombre1) + ' ' + RTRIM(CP2.CPeNombre2)) " _
        & " From Colectivo, Cliente C1, CPersona CP1, Cliente C2, CPersona CP2 " _
        & " Where ColNombre Like '" & Replace(strNombre, " ", "%") & "%'" _
        & " And ColCliente1 = C1.CliCodigo And C1.CliCodigo = CP1.CPeCliente " _
        & " And ColCliente2 = C2.CliCodigo And C2.CliCodigo = CP2.CPeCliente  And ColCerrado = 0"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un colectivo con esas características, o el mismo esta cerrado.", vbExclamation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            lID = RsAux(0)
            RsAux.Close
        Else
            RsAux.Close
            If LiAyuda.ActivarAyuda(cBase, Cons, 8500, 1, "Buscar") > 0 Then
                lID = LiAyuda.RetornoDatoSeleccionado(0)
            End If
        End If
        If lID > 0 Then BuscoColectivoPorCodigo lID
    End If
    Set LiAyuda = Nothing
    Exit Sub
ErrBCPT:
    clsGeneral.OcurrioError "Error al buscar el colectivo por título.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoColectivoPorCliente(idCliente As Long)
    Cons = "Select Colectivo.*,  Dir1 = C1.CliDireccion , Dir2 = C2.CliDireccion, Ced1 = C1.CliCIRUC, Ced2 = C2.CliCIRUC, Nom1 = RTRIM(RTRIM(CP1.CPeApellido1) + ' ' + RTRIM(CP1.CPeApellido2)) + ', ' +  RTRIM(RTRIM(CP1.CPeNombre1) + ' ' + RTRIM(CP1.CPeNombre2)) " _
                        & " , Nom2 = RTRIM(RTRIM(CP2.CPeApellido1) + ' ' + RTRIM(CP2.CPeApellido2)) + ', ' +  RTRIM(RTRIM(CP2.CPeNombre1) + ' ' + RTRIM(CP2.CPeNombre2)) " _
        & " From Colectivo, Cliente C1, CPersona CP1, Cliente C2, CPersona CP2 " _
        & " Where (ColCliente1 = " & idCliente & " Or ColCliente2 = " & idCliente & ") " _
        & " And ColCliente1 = C1.CliCodigo And C1.CliCodigo = CP1.CPeCliente " _
        & " And ColCliente2 = C2.CliCodigo And C2.CliCodigo = CP2.CPeCliente  And ColCerrado = 0"
    
    CargoCamposColectivo
    
End Sub
Private Sub Ayuda(strTexto As String)
    Status.Panels("Msg").Text = strTexto
End Sub
Private Sub AccionBuscarCliente(Persona As Boolean, Empresa As Boolean)
On Error GoTo ErrBC
    
    Dim frmBusco As New clsBuscarCliente
    Screen.MousePointer = 11
    frmBusco.ActivoFormularioBuscarClientes cBase, Persona, Empresa
    Screen.MousePointer = 11
    
    If frmBusco.BCClienteSeleccionado > 0 Then
        If frmBusco.BCTipoClienteSeleccionado = TipoCliente.Cliente Then
        
            Cons = "Select CliCodigo, CliCIRUC, CliDireccion, CPeApellido1, Nombre = RTRIM(RTRIM(CPeApellido1) + ' ' + RTRIM(CPeApellido2)) + ', ' +  RTRIM(RTRIM(CPeNombre1) + ' ' + RTRIM(CPeNombre2)) " _
                & " From Cliente, CPersona " _
                & " Where CLiCodigo = " & frmBusco.BCClienteSeleccionado _
                & " And CliCodigo = CPeCliente "
            
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
            
            If RsAux.EOF Then
                RsAux.Close
                MsgBox "No existe un cliente ingresado con esa cédula.", vbInformation, "ATENCIÓN"
            Else
                If Not IsNull(RsAux!CLiCIRUC) Then tCiCliente1.Text = RsAux!CLiCIRUC Else tCiCliente1.Text = ""
                tCiCliente1.Tag = RsAux!CliCodigo
                labCliente1.Caption = " " & Trim(RsAux!Nombre)
                labCliente1.Tag = Trim(RsAux!CPeApellido1)
                RsAux.Close
                If cTipoCuenta.ItemData(cTipoCuenta.ListIndex) = Cuenta.Colectivo Then
                    BuscoColectivoPorCliente tCiCliente1.Tag
                Else
                    HabilitoCamposCuenta
                    BuscoSiHayParaPesos
                End If
            End If
        Else
            Cons = "Select * " _
                & " From Cliente, CEmpresa " _
                & " Where CLiCodigo = " & frmBusco.BCClienteSeleccionado _
                & " And CliCodigo = CEmCliente "
            
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
            
            If RsAux.EOF Then
                RsAux.Close
            Else
                If Not IsNull(RsAux!CLiCIRUC) Then tRUC.Text = RsAux!CLiCIRUC Else tRUC.Text = ""
                tCiCliente1.Tag = RsAux!CliCodigo
                labCliente1.Caption = " " & Trim(RsAux!CEmFantasia)
                labCliente1.Tag = ""
                RsAux.Close
                HabilitoCamposCuenta
                BuscoSiHayParaPesos
            End If
        
        End If
    End If
    Set frmBusco = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrBC:
    clsGeneral.OcurrioError "Error al buscar el cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub InicializoGrillas()
    With vsRecibo
        .Cols = 1
        .Rows = 1
        .ExtendLastCol = True
        '.FormatString = "<Sucursal|<Recibo|<Cliente|<Comentario|>Importe|"
        '.ColWidth(0) = 850: .ColWidth(1) = 800: .ColWidth(2) = 2000: .ColWidth(3) = 3350: .ColWidth(4) = 1350:: .ColWidth(5) = 14:
        .FormatString = "Fecha|<Sucursal|<Documento|<Cliente|<Artículo/Comentario|>Aporte|>Asignado"
        .ColWidth(eColGrillaCuenta.Fecha) = 850
        .ColWidth(eColGrillaCuenta.Sucursal) = 900: .ColWidth(eColGrillaCuenta.Documento) = 900: .ColWidth(eColGrillaCuenta.Cliente) = 2000
        .ColWidth(eColGrillaCuenta.Articulo) = 2650: .ColWidth(eColGrillaCuenta.Aporte) = 1350: .ColWidth(eColGrillaCuenta.Asignado) = 1350
        .BackColorBkg = vbWindowBackground
    End With
    With vsFactura
        .Cols = 1
        .Rows = 1
        .Editable = True
        .FormatString = "<Sucursal|<Documento|<Cliente|<Artículos|>Importe|>Asignado|"
        .ColWidth(0) = 850: .ColWidth(2) = 1200: .ColWidth(3) = 3300: .ColWidth(4) = 1100:: .ColWidth(5) = 1100: .ColWidth(6) = 14:
        .BackColorBkg = vbWindowBackground
    End With
End Sub

Private Sub LimpioCamposCuenta()
    labSaldo.Caption = "0.00"
    labSalida.Caption = "0.00"
    lblAporteVencido.Caption = "0.00"
    lblAporteVencido.Visible = False
    lblTituloAporte.Visible = False
    lblTitSaldoCVenc.Visible = False
    lblSaldoConVencido.Visible = False
    cSucursal.Text = ""
    tSerie.Text = ""
    tNumero.Text = ""
    InicializoGrillas
    bSolicitarAutorizacion = False
End Sub
Private Sub OcultoCamposCuenta()
    cMoneda.Text = ""
    cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    vsRecibo.Enabled = False
    cSucursal.Enabled = False: cSucursal.BackColor = Inactivo
    tSerie.Enabled = False: tSerie.BackColor = Inactivo
    tNumero.Enabled = False: tNumero.BackColor = Inactivo
    tAsigna.Enabled = False: tAsigna.BackColor = Inactivo
    vsFactura.Enabled = False
    butGrabar.Enabled = False
    butRehabilitar.Enabled = False
    butRehabilitar.Visible = False
End Sub
Private Sub HabilitoCamposCuenta()
    cMoneda.Enabled = True: cMoneda.BackColor = vbWhite
    vsRecibo.Enabled = True
    cSucursal.Enabled = True: cSucursal.BackColor = vbWhite
    tSerie.Enabled = True: tSerie.BackColor = vbWhite
    tNumero.Enabled = True: tNumero.BackColor = vbWhite
    tAsigna.Enabled = True: tAsigna.BackColor = vbWhite
    vsFactura.Enabled = True
End Sub
Private Sub CargoDatosCuenta(IdMoneda As Integer)
On Error GoTo errCDC
    Screen.MousePointer = 11
    LimpioCamposCuenta
    
    CargoCuentaPersonal IdMoneda, cTipoCuenta.ItemData(cTipoCuenta.ListIndex)
    
    If vsRecibo.Rows = 1 Then
        MsgBox "No hay recibos para esa moneda.", vbExclamation, "ATENCIÓN"
        Foco cMoneda
        butGrabar.Enabled = False
    Else
    
        If vUsrPermisos <> miconexion.UsuarioLogueado(Codigo:=True) Then
            vUsrPermisos = miconexion.UsuarioLogueado(Codigo:=True)
            vPermisoSaldoVencido = miconexion.AccesoAlMenu("Asigno Factura Saldo Vencido")
        End If
        
        butRehabilitar.Visible = vPermisoSaldoVencido
        butRehabilitar.Enabled = False
        
        'Cargo sólo el saldo que puede utilizar
        labSaldo.Caption = Format(SaldoCuentaPersonal(False), FormatoMonedaP)
        
        'Cargo lo vencido.
        labSaldo.Tag = SaldoCuentaPersonal(True)
        
        lblSaldoConVencido.Caption = Format(CCur(labSaldo.Caption) + CCur(labSaldo.Tag), FormatoMonedaP)
        
        If CCur(labSaldo.Caption) <> CCur(labSaldo.Tag) And CCur(labSaldo.Tag) <> 0 And vPermisoSaldoVencido Then
            lblAporteVencido.Caption = Format(CCur(labSaldo.Tag), FormatoMonedaP)
            lblAporteVencido.Visible = True
            lblTituloAporte.Visible = True
            lblTitSaldoCVenc.Visible = True
            lblSaldoConVencido.Visible = True
        Else
            lblAporteVencido.Visible = False
            lblTituloAporte.Visible = False
            lblTitSaldoCVenc.Visible = False
            lblSaldoConVencido.Visible = False
        End If
        
        If vsFactura.Rows > 1 Then butGrabar.Enabled = True
        Foco cSucursal
    End If
    Screen.MousePointer = 0
    Exit Sub
errCDC:
    clsGeneral.OcurrioError "Error al buscar los datos de cuenta.", Err.Description
    Screen.MousePointer = 0
End Sub

'Private Sub CargoRecibos(IdMoneda As Integer, TipoCta As Integer)
'Dim idCli As Long
'
'    Cons = "SELECT SucAbreviacion, DocSerie, DocNumero, CPeApellido1, CPeNombre1, CEmFantasia, ArtNombre, DocTotal From Documento, Sucursal, Cliente " _
'                    & " Left Outer Join CPersona On CliCodigo = CPeCliente " _
'                    & " Left Outer Join CEmpresa On CliCodigo = CEmCliente " _
'            & ", AporteACuenta " _
'                & " Left Outer Join Articulo On ArtID = AaCArticulo " _
'        & " Where AaCTipoCuenta = " & TipoCta & " AND AaCTipo = " & TipoMovimientosAportes.Aporte _
'        & " And DocMoneda = " & IdMoneda
'
'    If TipoCta = Cuenta.Colectivo Then
'        Cons = Cons & " And AaCIDCuenta = " & tCodigo.Tag
'    Else
'        Cons = Cons & " And AaCIDCuenta = " & tCiCliente1.Tag
'    End If
'    Cons = Cons & " And AaCImporte Is Null And AaCDocumento = DocCodigo And DocAnulado = 0 " _
'        & " And DocTipo = " & TipoDocumento.ReciboDePago _
'        & " And DocCliente = CliCodigo And DocSucursal = SucCodigo "
'
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
'
'    Do While Not RsAux.EOF
'        With vsRecibo
'            .AddItem ""
'            .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!SucAbreviacion)
'            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!DocSerie) & " " & Trim(RsAux!DocNumero)
'            If Not IsNull(RsAux!CPeApellido1) Then
'                'Es persona
'                .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!CPeApellido1) & ", " & Trim(RsAux!CPeNombre1)
'            Else
'                'Es empresa.
'                .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!CEmFantasia)
'            End If
'            If Not IsNull(RsAux!ArtNombre) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!ArtNombre)
'            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!DocTotal, FormatoMonedaP)
'            labAporte.Caption = Format(CCur(labAporte.Caption) + RsAux!DocTotal, FormatoMonedaP)
'        End With6
'        RsAux.MoveNext
'    Loop
'    RsAux.Close
'
'End Sub

Private Sub CargoCuentaPersonal(IdMoneda As Integer, TipoCta As Integer)

    Cons = "SELECT AaCFecha, AaCTipo, AaCDocumento, SucAbreviacion, DocSerie, DocNumero" & _
        ", CASE WHEN CPeApellido1 IS NULL THEN CEmFantasia ELSE RTRIM(CPeNombre1) + ' ' + RTRIM(CPeApellido1) END Cliente " & _
        ", CASE WHEN AaCImporte IS Null THEN DocTotal ELSE AaCImporte END Importe " & _
        ", CASE WHEN DocTipo = 5 THEN '' ELSE dbo.ListaArticulosDelDocumento(DocCodigo) END Articulos " & _
        "FROM AporteACuenta INNER JOIN Documento ON AaCDocumento = DocCodigo AND DocAnulado = 0 AND DocMoneda = " & IdMoneda & _
        " INNER JOIN Sucursal ON DocSucursal = SucCodigo " & _
        "LEFT OUTER JOIN CPersona ON DocCliente = CPeCliente " & _
        "LEFT OUTER JOIN CEmpresa ON DocCliente = CEmCliente " & _
        "WHERE AaCTipoCuenta = " & TipoCta
            
    Dim idCta As Long
    If TipoCta = Cuenta.Colectivo Then
        idCta = tCodigo.Tag
    Else
        idCta = tCiCliente1.Tag
    End If
    
    Cons = Cons & " And AaCIDCuenta = " & idCta & " Order by AaCFecha "

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)

    Do While Not RsAux.EOF
        
        With vsRecibo
            .AddItem Format(RsAux("AaCFecha"), "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, eColGrillaCuenta.Sucursal) = Trim(RsAux!SucAbreviacion)
            .Cell(flexcpText, .Rows - 1, eColGrillaCuenta.Documento) = Trim(RsAux!DocSerie) & " " & Trim(RsAux!DocNumero)
            .Cell(flexcpText, .Rows - 1, eColGrillaCuenta.Cliente) = Trim(RsAux!Cliente)
            .Cell(flexcpText, .Rows - 1, eColGrillaCuenta.Articulo) = Trim(RsAux!Articulos)
            
            Select Case RsAux("AaCTipo")
                Case TipoMovimientosAportes.Aporte, TipoMovimientosAportes.Rehabilitado
                    .Cell(flexcpText, .Rows - 1, eColGrillaCuenta.Aporte) = Format(RsAux!Importe, FormatoMonedaP)
                Case Else
                    .Cell(flexcpText, .Rows - 1, eColGrillaCuenta.Asignado) = Format(RsAux!Importe, FormatoMonedaP)
                    If RsAux("AacTipo") = TipoMovimientosAportes.Inhabilitado Then
                        .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = &H808080
                    ElseIf RsAux("AacTipo") = TipoMovimientosAportes.Eliminado Then
                        .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = &H808080
                        .Cell(flexcpFontStrikethru, .Rows - 1, 0, , .Cols - 1) = True
                        .Cell(flexcpText, .Rows - 1, eColGrillaCuenta.Asignado) = ""
                    End If
            End Select
            
            'Datos de info.
            .Cell(flexcpData, .Rows - 1, eColGrillaCuenta.Documento) = CStr(RsAux("AaCDocumento"))
            .Cell(flexcpData, .Rows - 1, eColGrillaCuenta.Fecha) = CStr(RsAux("AaCFecha"))
            .Cell(flexcpData, .Rows - 1, eColGrillaCuenta.Sucursal) = CStr(RsAux("AaCTipo"))
            
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    

    Dim bAccADM As Boolean
    bAccADM = miconexion.AccesoAlMenu("Asigno Factura Saldo Vencido")
    
    'Consulto si alguno de estos está vencido y aún no se inhabilito.
    Cons = "SELECT DocFecha AaCFecha, " & TipoMovimientosAportes.Inhabilitado & " AaCTipo, DocCodigo AaCDocumento, SucAbreviacion, DocSerie, DocNumero" & _
        ", CASE WHEN CPeApellido1 IS NULL THEN CEmFantasia ELSE RTRIM(CPeNombre1) + ' ' + RTRIM(CPeApellido1) END Cliente " & _
        ", Saldo Importe " & _
        ", '' Articulos " & _
        "FROM SaldoCtaPersonalRecibos(" & cTipoCuenta.ItemData(cTipoCuenta.ListIndex) & ", " & idCta & ", 2) " & _
        "INNER JOIN Documento ON Recibo = DocCodigo AND DocAnulado = 0 AND DocMoneda = " & IdMoneda & _
        " INNER JOIN Sucursal ON DocSucursal = SucCodigo " & _
        "LEFT OUTER JOIN CPersona ON DocCliente = CPeCliente " & _
        "LEFT OUTER JOIN CEmpresa ON DocCliente = CEmCliente "

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsAux.EOF
        With vsRecibo
            .AddItem Format(RsAux("AaCFecha"), "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, eColGrillaCuenta.Sucursal) = Trim(RsAux!SucAbreviacion)
            .Cell(flexcpText, .Rows - 1, eColGrillaCuenta.Documento) = Trim(RsAux!DocSerie) & " " & Trim(RsAux!DocNumero)
            .Cell(flexcpText, .Rows - 1, eColGrillaCuenta.Cliente) = Trim(RsAux!Cliente)
            .Cell(flexcpText, .Rows - 1, eColGrillaCuenta.Articulo) = Trim(RsAux!Articulos)
            
            .Cell(flexcpText, .Rows - 1, eColGrillaCuenta.Asignado) = Format(RsAux!Importe, FormatoMonedaP)
            If RsAux("AacTipo") = TipoMovimientosAportes.Inhabilitado Then
                .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = IIf(bAccADM, &H80&, &H808080)
            End If
            
            'Datos de info.
            .Cell(flexcpData, .Rows - 1, eColGrillaCuenta.Documento) = CStr(RsAux("AaCDocumento"))
            .Cell(flexcpData, .Rows - 1, eColGrillaCuenta.Fecha) = CStr(RsAux("AaCFecha"))
            .Cell(flexcpData, .Rows - 1, eColGrillaCuenta.Sucursal) = CStr(RsAux("AaCTipo"))
            
            'Me marco que este está inhabilitado por función.
            .Cell(flexcpData, .Rows - 1, eColGrillaCuenta.Cliente) = 1
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If vsRecibo.Rows > vsRecibo.FixedRows Then
        With vsRecibo
            .SubtotalPosition = flexSTAbove
            .Subtotal flexSTSum, -1, eColGrillaCuenta.Aporte, FormatoMonedaP, &HD0D8CD, &H2C511E, True, "Total"
            .Subtotal flexSTSum, -1, eColGrillaCuenta.Asignado, FormatoMonedaP
            .Cell(flexcpForeColor, 1, eColGrillaCuenta.Asignado) = &H80&
            .Refresh
        End With
    End If
    
End Sub

Private Function SaldoCuentaPersonal(ByVal IncluyoInhabilitado As Boolean) As Currency
On Error GoTo errSCP
    
    Screen.MousePointer = 11
    SaldoCuentaPersonal = 0
    Dim rsS As rdoResultset
    
    Dim idCta As Long
    If cTipoCuenta.ItemData(cTipoCuenta.ListIndex) = Cuenta.Colectivo Then
        idCta = tCodigo.Tag
    Else
        idCta = tCiCliente1.Tag
    End If
    
    Dim oACta As New clsAporteACuenta
    SaldoCuentaPersonal = oACta.SaldoCuentaPersonal(cBase, cTipoCuenta.ItemData(cTipoCuenta.ListIndex), idCta, IncluyoInhabilitado)
    Set oACta = Nothing
    
    Screen.MousePointer = 0
    Exit Function
    
errSCP:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar el saldo de la cuenta.", Err.Description, "Saldo cuenta personal"
End Function

Private Sub CargoFacturas(IdMoneda As Integer, TipoCta As Integer)
Dim CodDocumento As Long

    Cons = "Select DocCodigo, SucAbreviacion, DocSerie, DocNumero, CLiCIRUC, CliTipo, ArtNombre, DocTotal, AaCImporte From AporteACuenta " _
            & "INNER JOIN Documento ON AaCDocumento = DocCodigo And DocAnulado = 0 " _
            & "LEFT OUTER JOIN Renglon ON DocCodigo = RenDocumento " _
            & "LEFT OUTER JOIN Articulo ON RenArticulo = ArtID " _
        & " , Sucursal, Cliente " _
        & "WHERE AaCTipo = " & TipoMovimientosAportes.Asignado _
        & "AND AaCTipoCuenta = " & TipoCta _
        & " And DocMoneda = " & IdMoneda
        
    If TipoCta = Cuenta.Colectivo Then
        Cons = Cons & " And AaCIDCuenta = " & tCodigo.Tag
    Else
        Cons = Cons & " And AaCIDCuenta = " & tCiCliente1.Tag
    End If
    Cons = Cons _
        & " And DocTipo IN (" & TipoDocumento.Contado & ", " & TipoDocumento.ReciboDePago & ")" _
        & " And DocCliente = CliCodigo And DocSucursal = SucCodigo" _
        & " Order by DocCodigo "

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)

    CodDocumento = 0
    Do While Not RsAux.EOF
        With vsFactura
            If CodDocumento <> RsAux!DocCodigo Then
                CodDocumento = RsAux!DocCodigo
                .AddItem ""
                .Cell(flexcpData, .Rows - 1, 0) = CodDocumento
                .Cell(flexcpData, .Rows - 1, 1) = 1
                .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!SucAbreviacion)
                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!DocSerie) & " " & Trim(RsAux!DocNumero)
                If Not IsNull(RsAux!CLiCIRUC) Then
                    If RsAux!CliTipo = TipoCliente.Cliente Then
                        .Cell(flexcpText, .Rows - 1, 2) = clsGeneral.RetornoFormatoCedula(RsAux!CLiCIRUC)
                    Else
                        .Cell(flexcpText, .Rows - 1, 2) = clsGeneral.RetornoFormatoRuc(RsAux!CLiCIRUC)
                    End If
                End If
                If Not IsNull(RsAux!ArtNombre) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!DocTotal, FormatoMonedaP)
                If Not IsNull(RsAux!AaCImporte) Then
                    .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!AaCImporte, FormatoMonedaP)
'                    labAsignado.Caption = Format(CCur(labAsignado.Caption) + RsAux!AaCImporte, FormatoMonedaP)
                Else
                    .Cell(flexcpText, .Rows - 1, 5) = "0.00"
                End If
            Else
                'No tiene que ocurrir que un recibo devuelva + de una fila.
                If Not IsNull(RsAux!ArtNombre) Then .Cell(flexcpText, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 3) & ", " & Trim(RsAux!ArtNombre)
            End If
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub

Private Sub tRUC_Change()
    tCiCliente1.Tag = "0": labCliente1.Caption = "": labCliente1.Tag = ""
End Sub

Private Sub tRUC_GotFocus()
    With tRUC
        .SelStart = 0: .SelLength = 15
    End With
    Ayuda "Ingrese el R.U.C. de la empresa."
End Sub

Private Sub tRUC_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            tCiCliente1.Tag = "0": labCliente1.Caption = "": labCliente1.Tag = ""
            BuscoClientePorRUC tRUC.Text
            HabilitoCamposCuenta
            BuscoSiHayParaPesos
        Case vbKeyF4: AccionBuscarCliente False, True
    End Select
End Sub

Private Sub tRUC_LostFocus()
    Ayuda ""
    tRUC.SelStart = 0
End Sub

Private Sub BuscoClientePorRUC(Ruc As String)
On Error GoTo ErrBR
    Screen.MousePointer = 11
    Cons = "Select CliCodigo, CEmFantasia " _
        & " From Cliente, CEmpresa " _
        & " Where CLiCIRUC = '" & Ruc & "'" _
        & " And CliCodigo = CEmCliente "
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un cliente ingresado con ese R.U.C.", vbInformation, "ATENCIÓN"
        tCiCliente1.Tag = "0"
        labCliente1.Caption = " "
        labCliente1.Tag = ""
    Else
        tCiCliente1.Tag = RsAux!CliCodigo
        labCliente1.Caption = " " & Trim(RsAux!CEmFantasia)
        labCliente1.Tag = ""
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBR:
    clsGeneral.OcurrioError "Error al buscar al cliente por R.U.C.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tSerie_Change()
    tSerie.Text = UCase(tSerie.Text)
End Sub

Private Sub tSerie_GotFocus()
    With tSerie
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese la serie del contado."
End Sub
Private Sub tSerie_KeyPress(KeyAscii As Integer)
    'tSerie.Text = UCase(tSerie.Text)
    If KeyAscii = vbKeyReturn Then
        If Trim(tSerie.Text) <> "" Then Foco tNumero
    End If
End Sub
Private Sub tSerie_LostFocus()
    tSerie.SelStart = 0: tSerie.Text = UCase(tSerie.Text)
    Ayuda ""
End Sub

Private Sub vsFactura_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 5 Then Cancel = True: Exit Sub
    If Val(vsFactura.Cell(flexcpData, vsFactura.Row, 1)) <> 2 Then Cancel = True
End Sub

Private Sub vsFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrBT
    
    Select Case KeyCode
        Case vbKeyDelete
            
            If vsFactura.Row > 0 Then
            
                If vsFactura.Cell(flexcpData, vsFactura.Row, 1) = 2 Then
                
                    labSalida.Caption = Format(CCur(labSalida.Caption) - CCur(vsFactura.Cell(flexcpText, vsFactura.Row, 5)), FormatoMonedaP)
                    vsFactura.RemoveItem vsFactura.Row
                    butGrabar.Enabled = (vsFactura.Rows > 1)
                    'Si elimina todo entonces saco condición de pedir suceso.
                    If vsFactura.Rows = vsFactura.FixedRows Then bSolicitarAutorizacion = False
                    
                End If
            End If
    End Select
    Screen.MousePointer = 0
    Exit Sub

ErrBT:
    clsGeneral.OcurrioError "Error al iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrResumir:
    Resume ErrFT
ErrFT:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar almacenar la información.", Err.Description
End Sub

Private Sub vsFactura_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not IsNumeric(vsFactura.EditText) Then Cancel = True: Exit Sub
    If CCur(vsFactura.EditText) <= 0 Then Cancel = True: Exit Sub
    If CCur(vsFactura.EditText) > CCur(vsFactura.Cell(flexcpData, Row, 5)) Then Cancel = True: Exit Sub
    labSalida.Caption = Format((CCur(labSalida.Caption) - CCur(vsFactura.Cell(flexcpText, Row, Col))) + CCur(vsFactura.EditText), FormatoMonedaP)
End Sub

Private Sub ImprimoSalidaCajaTicket(ByVal Codigo As Long)
On Error GoTo errPrint
    Screen.MousePointer = 11
    Dim sSucursal As String
    sSucursal = CargoParametrosSucursal
    Dim oImp As New clsImpresionDeDocumentos
    oImp.PathReportes = gPathListados
    'Set oImp.DondeImprimo = New clsConfigImpresora
    oImp.NombreBaseDatos = miconexion.RetornoPropiedad(False, False, False, True)
    oImp.DondeImprimo.Impresora = oCnfgPrint.ImpresoraTickets
    oImp.ImprimoSalidaCajaTicket Codigo, sSucursal, BuscoUsuario(miconexion.UsuarioLogueado(True), False, False, True), "Señas recibidas"
    Set oImp = Nothing
    Screen.MousePointer = 0
    Exit Sub
errPrint:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar imprimir la salida de caja.", Err.Description, "Impresión"

'    With vsrReport
'        .Clear                  ' clear any existing fields
'        .FontName = "Tahoma"    ' set default font for all controls
'        .FontSize = 8
'
'        .Load gPathListados & "MovimientoDeCaja.xml", "MovimientoDeCaja"
'
'        .DataSource.ConnectionString = cBase.Connect
'        .DataSource.RecordSource = "SELECT MDiID Numero, '" & CargoParametrosSucursal & "' Sucursal, '" & BuscoUsuario(miconexion.UsuarioLogueado(True), False, False, True) & _
'                    "' Usuario, 'Señas Recibas' TipoMovimiento " & _
'                    ", MDiFecha + MDiHora Fecha, DisNombre Disponibilidad, Moneda.MonSigno  + ' ' + CONVERT(varchar(20), MDrImporteCompra) Importe " & _
'                    ", RTRIM(MDiComentario) Memo, CASE WHEN MDRDebe IS NULL THEN 'Salida de caja' ELSE 'Entrada de caja' END EntradaSalida " & _
'                    "FROM CGSA.dbo.MovimientoDisponibilidad MovimientoDisponibilidad " & _
'                    "INNER JOIN CGSA.dbo.MovimientoDisponibilidadRenglon MovimientoDisponibilidadRenglon ON MDiID = MDRIDMovimiento " & _
'                    "INNER JOIN CGSA.dbo.Disponibilidad Disponibilidad  ON MDRIdDisponibilidad = DisID " & _
'                    "INNER JOIN CGSA.dbo.Moneda ON DisMoneda = MonCodigo WHERE MDiID = " & Codigo
'
'        vspPrinter.Device = oCnfgPrint.ImpresoraTickets
'        .Render vspPrinter
'
'    End With
'
'    vspPrinter.PrintDoc False
End Sub

Private Sub ImprimoSalidaCaja(Codigo As Long)

Dim NombreFormula As String, Result As Integer
Dim JobNumMC As Integer, CantFormMC As Integer

    'Inicializa el Engine del Crystal y setea la impresora para el JOB
    On Error GoTo ErrCrystal
    
    'Inicializo el Reporte y SubReportes
    JobNumMC = crAbroReporte(gPathListados & "MovimientoCaja.RPT")
    If JobNumMC = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
    If Trim(Printer.DeviceName) <> Trim(paPrintConfD) Then SeteoImpresoraPorDefecto paPrintConfD
    If Not crSeteoImpresora(JobNumMC, Printer, paPrintConfB) Then GoTo ErrCrystal

    'Obtengo la cantidad de formulas que tiene el reporte.
    CantFormMC = crObtengoCantidadFormulasEnReporte(JobNumMC)
    If CantFormMC = -1 Then GoTo ErrCrystal

    For I = 0 To CantFormMC - 1
        NombreFormula = crObtengoNombreFormula(JobNumMC, I)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "sucursal": Result = crSeteoFormula(JobNumMC%, NombreFormula, "'Sucursal: " & CargoParametrosSucursal & "'")
            Case "tipo": Result = crSeteoFormula(JobNumMC%, NombreFormula, "'Señas Recibas'")
            Case "moneda": Result = crSeteoFormula(JobNumMC%, NombreFormula, "'" & cMoneda.Text & "'")
            Case "usuario": Result = crSeteoFormula(JobNumMC%, NombreFormula, "'" & BuscoUsuario(miconexion.UsuarioLogueado(True), False, False, True) & "'")
            Case Else: Result = 1
        End Select
        If Result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT * " _
            & " From " & paBD & ".dbo.MovimientoDisponibilidad MovimientoDisponibilidad, " _
                            & paBD & ".dbo.MovimientoDisponibilidadRenglon MovimientoDisponibilidadRenglon, " _
                            & paBD & ".dbo.Disponibilidad Disponibilidad " _
            & " Where MDiID = " & Codigo _
            & " And MDiID = MDRIdMovimiento And MDRIdDisponibilidad = DisID"
    
    If crSeteoSqlQuery(JobNumMC%, Cons) = 0 Then GoTo ErrCrystal
            
    'If crMandoAPantalla(JobNumMC, "Movimiento de Caja") = 0 Then GoTo ErrCrystal
    If crMandoAImpresora(JobNumMC, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(JobNumMC, True, False) Then GoTo ErrCrystal
    'crEsperoCierreReportePantalla

    crCierroTrabajo JobNumMC
    Screen.MousePointer = 0
    Exit Sub

ErrCrystal:
    On Error Resume Next
    crCierroTrabajo JobNumMC
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr, Err.Description
End Sub

Private Sub BuscoSiHayParaPesos()
    cMoneda.ListIndex = posMonedaPesos
    If cMoneda.ListIndex > -1 Then CargoDatosCuenta cMoneda.ItemData(cMoneda.ListIndex)
    If vsRecibo.Rows > 1 Then Foco cSucursal Else Foco cMoneda
End Sub

Private Sub vsRecibo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errDel
    
    If KeyCode <> vbKeyDelete Then Exit Sub
    
    If vsRecibo.Cell(flexcpData, vsRecibo.Row, eColGrillaCuenta.Sucursal) = TipoMovimientosAportes.Asignado Then
        
        'Quedamos en que si no es del día la asignación entonces no dejo eliminarlo.
        If CDate(vsRecibo.Cell(flexcpText, vsRecibo.Row, eColGrillaCuenta.Fecha)) <> Date Then
            MsgBox "Sólo se pueden eliminar asignaciones del día.", vbExclamation, "Posible error"
            Exit Sub
        End If
        
        m_Disponibilidad = modMaeDisponibilidad.dis_DisponibilidadPara(paCodigoDeSucursal, cMoneda.ItemData(cMoneda.ListIndex))
        If m_Disponibilidad = 0 Then
            MsgBox "No hay una disponibilidad para hacer los movimientos de caja en la moneda seleccionada." & vbCrLf & "Consulte al Administrador.", vbQuestion, "Falta Disponibilidad"
            Exit Sub
        End If
        
        
        If MsgBox("Este documento está asignado a la cuenta, si lo elimina se hará una entrada de caja por el importe asignado." & Chr(13) _
                            & "¿Confirma quitar esta factura asignada a la cuenta?", vbQuestion + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbYes Then
                    Screen.MousePointer = 11
            
            FechaDelServidor
            On Error GoTo ErrBT
            cBase.BeginTrans
            On Error GoTo ErrResumir
            
            'ELIMINO EL APORTE.
            Dim oApo As New clsAporteACuenta
            With vsRecibo
                oApo.Documento = .Cell(flexcpData, .Row, eColGrillaCuenta.Documento)
                oApo.Fecha = .Cell(flexcpData, .Row, eColGrillaCuenta.Fecha)
                oApo.MovimientoAporte = .Cell(flexcpData, .Row, eColGrillaCuenta.Sucursal)
                oApo.Importe = CCur(.Cell(flexcpText, .Row, eColGrillaCuenta.Asignado))
            End With
            oApo.TipoCuenta = cTipoCuenta.ItemData(cTipoCuenta.ListIndex)
            If cTipoCuenta.ItemData(cTipoCuenta.ListIndex) = Cuenta.Colectivo Then
                oApo.IDCuenta = tCodigo.Tag
            Else
                oApo.IDCuenta = tCiCliente1.Tag
            End If
            'oApo.EliminarAporte cBase
            oApo.CambioTipoAporte cBase, TipoMovimientosAportes.Eliminado
            '...............................................
            
            Dim oCx As New clsCobrarConQuePaga
            Set oCx.Conexion = cBase
            
            'Tengo que quitarle la relación conquecobra.
            Dim oMC As New clsDesicionConQuePaga
            oMC.Sucursal = paCodigoDeSucursal
            aMovimientoCaja = oMC.GrabarMovimientoDeCaja(oApo.Importe, "Ctdo.: " & Trim(vsRecibo.Cell(flexcpText, vsRecibo.Row, eColGrillaCuenta.Sucursal)) & " " & Trim(vsRecibo.Cell(flexcpText, vsRecibo.Row, eColGrillaCuenta.Documento)))
            Set oMC = Nothing
            
            Set oCx = Nothing
            cBase.CommitTrans

            On Error Resume Next
            
            If oCnfgPrint.Opcion = 0 Then
                ImprimoSalidaCaja aMovimientoCaja
            Else
                ImprimoSalidaCajaTicket aMovimientoCaja
            End If
                
            'Cargo de nuevo la información.
            CargoDatosCuenta cMoneda.ItemData(cMoneda.ListIndex)
            Screen.MousePointer = 0
        End If
        
    End If
    Exit Sub
errDel:
    clsGeneral.OcurrioError "Error al eliminar la asignación.", Err.Description, "Eliminar asignación"
    Exit Sub

ErrBT:
    clsGeneral.OcurrioError "Error al iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrResumir:
    Resume ErrFT
ErrFT:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar almacenar la información.", Err.Description
End Sub

Private Sub vsRecibo_RowColChange()
On Error Resume Next
    If butRehabilitar.Visible Then
        If vsRecibo.Rows > vsRecibo.FixedRows Then
            butRehabilitar.Enabled = (vsRecibo.Cell(flexcpData, vsRecibo.Row, eColGrillaCuenta.Sucursal) = TipoMovimientosAportes.Inhabilitado And vsRecibo.Cell(flexcpData, vsRecibo.Row, eColGrillaCuenta.Cliente) = 0)
        End If
    End If
End Sub

Private Sub RehabilitarRecibo()
    If MsgBox("¿Desea rehabilitar el recibo seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2, "Rehabilitar aportes") = vbYes Then
        MsgBox "TODO:", vbExclamation, "ATENCIÓN"
    End If
End Sub

Private Sub FormularioACero()
'    LimpioCamposInformacion
    OcultoCamposInformacion
    HabilitoCamposInformacion
'    LimpioCamposCuenta
    OcultoCamposCuenta
    
End Sub
