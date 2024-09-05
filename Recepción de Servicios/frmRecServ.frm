VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form frmRecServ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Servicios"
   ClientHeight    =   7830
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRecServ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   9210
   Begin VB.CheckBox chCoordinarEntrega 
      Alignment       =   1  'Right Justify
      Caption         =   "Al finalizar se debe coordinar entrega"
      Height          =   255
      Left            =   3840
      TabIndex        =   85
      Top             =   5100
      Width           =   3255
   End
   Begin VB.CheckBox chkNoDeseaSMS 
      Caption         =   "No desea recibir SMS"
      Height          =   255
      Left            =   2880
      TabIndex        =   84
      Top             =   7200
      Width           =   2535
   End
   Begin VB.PictureBox PicTaller 
      Height          =   1575
      Left            =   2040
      ScaleHeight     =   1515
      ScaleWidth      =   5475
      TabIndex        =   60
      Top             =   5760
      Width           =   5535
      Begin VB.CheckBox chkFueraGarantia 
         Caption         =   "Ingresar servicio Fuera de Garantía"
         Height          =   375
         Left            =   3720
         TabIndex        =   83
         Top             =   840
         Width           =   3495
      End
      Begin AACombo99.AACombo cTalDeposito 
         Height          =   315
         Left            =   4920
         TabIndex        =   27
         Top             =   420
         Width           =   2115
         _ExtentX        =   3731
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
      Begin VB.Label ltMotivo 
         BackStyle       =   0  'Transparent
         Caption         =   "&Motivos:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   60
         Width           =   855
      End
      Begin VB.Label lTalFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12/12/99 "
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   4920
         TabIndex        =   62
         Top             =   60
         Width           =   1155
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Re&paración en:"
         Height          =   195
         Left            =   3600
         TabIndex        =   26
         Top             =   420
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ingreso:"
         Height          =   195
         Left            =   3600
         TabIndex        =   61
         Top             =   60
         Width           =   1155
      End
   End
   Begin VB.PictureBox PicRetiro 
      Height          =   1635
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   9375
      TabIndex        =   63
      Top             =   5400
      Width           =   9435
      Begin AACombo99.AACombo cRDeposito 
         Height          =   315
         Left            =   7500
         TabIndex        =   43
         Top             =   900
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
      Begin VB.TextBox tRLiquidar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6300
         MaxLength       =   14
         TabIndex        =   41
         Top             =   900
         Width           =   735
      End
      Begin AACombo99.AACombo cRMoneda 
         Height          =   315
         Left            =   7080
         TabIndex        =   36
         Top             =   480
         Width           =   750
         _ExtentX        =   1323
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
      Begin VB.TextBox tRImporte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7860
         TabIndex        =   37
         Text            =   "1,114,500.00"
         Top             =   480
         Width           =   1035
      End
      Begin AACombo99.AACombo cRCamion 
         Height          =   315
         Left            =   7080
         TabIndex        =   31
         Top             =   60
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
      Begin AACombo99.AACombo cRHora 
         Height          =   315
         Left            =   5340
         TabIndex        =   34
         Top             =   480
         Width           =   1155
         _ExtentX        =   2037
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
      Begin VB.TextBox tRFecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         TabIndex        =   33
         Top             =   495
         Width           =   1155
      End
      Begin AACombo99.AACombo cTipoFlete 
         Height          =   315
         Left            =   4140
         TabIndex        =   29
         Top             =   60
         Width           =   2115
         _ExtentX        =   3731
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
      Begin AACombo99.AACombo cRFactura 
         Height          =   315
         Left            =   4200
         TabIndex        =   39
         Top             =   900
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label25 
         Caption         =   "&Va a:"
         Height          =   255
         Left            =   7080
         TabIndex        =   42
         Top             =   900
         Width           =   435
      End
      Begin VB.Label Label23 
         Caption         =   "&Liquidar:"
         Height          =   195
         Left            =   5640
         TabIndex        =   40
         Top             =   900
         Width           =   735
      End
      Begin VB.Label lRtMotivo 
         BackStyle       =   0  'Transparent
         Caption         =   "&Motivos:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         Height          =   195
         Left            =   3540
         TabIndex        =   32
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Ca&mión:"
         Height          =   195
         Left            =   6420
         TabIndex        =   30
         Top             =   60
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor:"
         Height          =   195
         Left            =   6600
         TabIndex        =   35
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&Factura:"
         Height          =   195
         Left            =   3540
         TabIndex        =   38
         Top             =   900
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "F&lete:"
         Height          =   195
         Left            =   3540
         TabIndex        =   28
         Top             =   60
         Width           =   435
      End
   End
   Begin VB.TextBox tComentarioInterno 
      Appearance      =   0  'Flat
      Height          =   675
      Left            =   1020
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   67
      Top             =   6480
      Width           =   7935
   End
   Begin VB.TextBox tReclamo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2340
      MaxLength       =   10
      TabIndex        =   82
      Top             =   5100
      Width           =   1155
   End
   Begin VB.CheckBox chEsReclamo 
      Alignment       =   1  'Right Justify
      Caption         =   "Es reclamo"
      Height          =   255
      Left            =   120
      TabIndex        =   80
      Top             =   5100
      Width           =   1095
   End
   Begin vsViewLib.vsPrinter vsFicha 
      Height          =   555
      Left            =   4200
      TabIndex        =   79
      Top             =   3300
      Visible         =   0   'False
      Width           =   1635
      _Version        =   196608
      _ExtentX        =   2884
      _ExtentY        =   979
      _StockProps     =   229
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
      PageBorder      =   0
   End
   Begin VB.TextBox tTelCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox tDirCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   77
      Top             =   780
      Width           =   7695
   End
   Begin VB.TextBox tTelefonoServicio 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1020
      MaxLength       =   15
      TabIndex        =   69
      Top             =   7200
      Width           =   1575
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   75
      Top             =   7560
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   13626
            Key             =   "msg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Picture         =   "frmRecServ.frx":030A
            Key             =   "printer"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8340
      MaxLength       =   3
      TabIndex        =   71
      Top             =   7200
      Width           =   615
   End
   Begin VB.PictureBox picHistoria 
      Height          =   735
      Left            =   600
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   73
      Top             =   3600
      Width           =   1095
      Begin VSFlex6DAOCtl.vsFlexGrid vsHistoria 
         Height          =   795
         Left            =   60
         TabIndex        =   76
         Top             =   60
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1402
         _ConvInfo       =   1
         Appearance      =   1
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "P&roductos"
      Height          =   2055
      Left            =   60
      TabIndex        =   7
      Top             =   1200
      Width           =   9075
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   70
         TabIndex        =   10
         Top             =   1320
         Width           =   4095
      End
      Begin VB.TextBox tFCompra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   12
         Text            =   "12/10/2000"
         Top             =   1320
         Width           =   1035
      End
      Begin VB.TextBox tfSerie 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7860
         MaxLength       =   14
         TabIndex        =   14
         Text            =   "QQ"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox tfNumero 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8220
         MaxLength       =   14
         TabIndex        =   15
         Text            =   "888888"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox tPSerie 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7860
         MaxLength       =   40
         TabIndex        =   19
         Text            =   "888888"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox tDireccion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   100
         TabIndex        =   58
         Top             =   1680
         Width           =   5775
      End
      Begin VB.CommandButton bDireccionP 
         Height          =   315
         Left            =   6660
         Picture         =   "frmRecServ.frx":041C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1680
         Width           =   315
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsProducto 
         Height          =   1035
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   1826
         _ConvInfo       =   1
         Appearance      =   1
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
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
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
         MultiTotals     =   0   'False
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
      Begin VB.Label lTipoProducto 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "&F. Compra:"
         Height          =   255
         Left            =   5100
         TabIndex        =   11
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "F&actura:"
         Height          =   255
         Left            =   7140
         TabIndex        =   13
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "&Dirección:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "&N° Serie:"
         Height          =   255
         Left            =   7140
         TabIndex        =   18
         Top             =   1680
         Width           =   675
      End
   End
   Begin VB.TextBox tAclaracion 
      Appearance      =   0  'Flat
      Height          =   1035
      Left            =   1020
      MaxLength       =   600
      MultiLine       =   -1  'True
      TabIndex        =   65
      Top             =   5400
      Width           =   7935
   End
   Begin VB.TextBox tMotivo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      MaxLength       =   35
      TabIndex        =   24
      Top             =   3900
      Width           =   1695
   End
   Begin VB.ComboBox cDatoIngreso 
      Height          =   315
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton bDireccionT 
      Height          =   315
      Left            =   8700
      Picture         =   "frmRecServ.frx":05A6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   780
      Width           =   315
   End
   Begin ComctlLib.TabStrip TabRecepcion 
      Height          =   1695
      Left            =   60
      TabIndex        =   20
      Top             =   3300
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2990
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Historia"
            Key             =   "historia"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "S&olicitud de Servicio"
            Key             =   "comodin"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox tCi 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   165
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   12582912
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
   Begin MSMask.MaskEdBox tRuc 
      Height          =   285
      Left            =   2340
      TabIndex        =   2
      Top             =   165
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   12582912
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
      Mask            =   "99 999 999 9999"
      PromptChar      =   "_"
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsMotivos 
      Height          =   795
      Left            =   6600
      TabIndex        =   25
      Top             =   4260
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1402
      _ConvInfo       =   1
      Appearance      =   1
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
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
   Begin VB.PictureBox PicVisita 
      Height          =   3495
      Left            =   -120
      ScaleHeight     =   3435
      ScaleWidth      =   8655
      TabIndex        =   72
      Top             =   3840
      Width           =   8715
      Begin VB.TextBox tVLiquidar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         MaxLength       =   14
         TabIndex        =   55
         Top             =   900
         Width           =   915
      End
      Begin AACombo99.AACombo cVMoneda 
         Height          =   315
         Left            =   4200
         TabIndex        =   50
         Top             =   480
         Width           =   750
         _ExtentX        =   1323
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
      Begin VB.TextBox tVFecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6600
         TabIndex        =   47
         Top             =   75
         Width           =   1155
      End
      Begin VB.TextBox tVImporte 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4980
         TabIndex        =   51
         Top             =   480
         Width           =   1035
      End
      Begin AACombo99.AACombo cVCamion 
         Height          =   315
         Left            =   4200
         TabIndex        =   45
         Top             =   60
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
      Begin AACombo99.AACombo cVHora 
         Height          =   315
         Left            =   7740
         TabIndex        =   48
         Top             =   60
         Width           =   1155
         _ExtentX        =   2037
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
      Begin AACombo99.AACombo cVFactura 
         Height          =   315
         Left            =   6720
         TabIndex        =   53
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
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
      Begin AACombo99.AACombo cVComentario 
         Height          =   315
         Left            =   6240
         TabIndex        =   57
         Top             =   900
         Width           =   2595
         _ExtentX        =   4577
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Co&mentario:"
         Height          =   195
         Left            =   5220
         TabIndex        =   56
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "&Liquidar:"
         Height          =   195
         Left            =   3540
         TabIndex        =   54
         Top             =   900
         Width           =   795
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "&Factura:"
         Height          =   195
         Left            =   6060
         TabIndex        =   52
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "&Valor:"
         Height          =   195
         Left            =   3540
         TabIndex        =   49
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Ca&mión:"
         Height          =   195
         Left            =   3540
         TabIndex        =   44
         Top             =   60
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Fec&ha:"
         Height          =   195
         Left            =   6060
         TabIndex        =   46
         Top             =   60
         Width           =   615
      End
      Begin VB.Label lVtMotivo 
         BackStyle       =   0  'Transparent
         Caption         =   "&Motivos:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario Interno:"
      Height          =   615
      Left            =   120
      TabIndex        =   66
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "del servicio:"
      Height          =   255
      Left            =   1380
      TabIndex        =   81
      Top             =   5100
      Width           =   1155
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   6000
      Picture         =   "frmRecServ.frx":0730
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "T&eléfono:"
      Height          =   255
      Left            =   120
      TabIndex        =   68
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "&Aclaración para cliente:"
      Height          =   615
      Left            =   120
      TabIndex        =   74
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label ltUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   255
      Left            =   7620
      TabIndex        =   70
      Top             =   6840
      Width           =   675
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de &Recepción:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6240
      TabIndex        =   3
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   195
      Left            =   180
      TabIndex        =   64
      Top             =   480
      Width           =   615
   End
   Begin ComctlLib.ImageList Image1 
      Left            =   8340
      Top             =   3540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   12632256
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecServ.frx":0A3A
            Key             =   "taller"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecServ.frx":1084
            Key             =   "retiro"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecServ.frx":139E
            Key             =   "visita"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecServ.frx":16B8
            Key             =   "historia"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmRecServ.frx":19D2
            Key             =   "servicio"
         EndProperty
      EndProperty
   End
   Begin VB.Label lTitular 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      TabIndex        =   59
      Top             =   480
      UseMnemonic     =   0   'False
      Width           =   4935
   End
   Begin VB.Label lNDireccion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Dirección:"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   795
      Width           =   705
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&C.I./RUC:"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   855
   End
   Begin VB.Shape shpTitular 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1095
      Left            =   105
      Top             =   60
      Width           =   9015
   End
   Begin VB.Menu MnuBuscar 
      Caption         =   "Buscar"
      Visible         =   0   'False
      Begin VB.Menu MnuBuFichaCliente 
         Caption         =   "Ficha del Cliente"
         Shortcut        =   {F2}
      End
      Begin VB.Menu MnuBuCGSA 
         Caption         =   "Empresa C.G.S.A."
         Shortcut        =   {F11}
      End
      Begin VB.Menu MnuBuFichaLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBuNuevo 
         Caption         =   "Nuevo"
         Checked         =   -1  'True
         Shortcut        =   {F3}
      End
      Begin VB.Menu MnuBuNuevoCliente 
         Caption         =   "Nuevo Cliente"
      End
      Begin VB.Menu MnuBuNuevaEmpresa 
         Caption         =   "Nueva Empresa"
      End
      Begin VB.Menu MnuBuFiLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBusquedas 
         Caption         =   "Búsquedas"
         Checked         =   -1  'True
         Shortcut        =   {F4}
      End
      Begin VB.Menu MnuBPersonas 
         Caption         =   "Buscar &Personas"
      End
      Begin VB.Menu MnuBEmpresas 
         Caption         =   "Buscar &Empresas"
      End
   End
   Begin VB.Menu MnuOpDireccion 
      Caption         =   "OpDireccion"
      Visible         =   0   'False
      Begin VB.Menu MnuOpDiTitulo 
         Caption         =   "Menú Dirección"
      End
      Begin VB.Menu MnuOD1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpDiConfirmar 
         Caption         =   "&Confirmar Dirección"
      End
      Begin VB.Menu MnuOpDiModificar 
         Caption         =   "&Modificar Dirección"
      End
   End
   Begin VB.Menu MnuProducto 
      Caption         =   "Producto&s"
      Begin VB.Menu MnuProFicha 
         Caption         =   "&Ficha del Producto"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuProCederProducto 
         Caption         =   "&Ceder el Producto a Otro Cliente"
      End
      Begin VB.Menu MnuProFindSerie 
         Caption         =   "Buscar producto por Nº de Serie"
      End
      Begin VB.Menu MnuProCumplirServicio 
         Caption         =   "Cumplir Servicio  (-)"
      End
      Begin VB.Menu MnuProLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuProSeguimientoReporte 
         Caption         =   "&Seguimiento de Servicio"
      End
      Begin VB.Menu MnuProHistoria 
         Caption         =   "&Historia de Servicio"
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalDel 
         Caption         =   "Cerrar"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuAyuda 
      Caption         =   "?"
      Begin VB.Menu MnuAyuHelp 
         Caption         =   "&Ayuda"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmRecServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Cambios:
    '6-7-2000: se agregaron los argumentos para realizar notas.
    '                 si una direccion es <> a la del cliente cambio el color del label.
    '13-11-00 Agregue Agenda de tipos de flete
    '11-12-00 Ordeno por tipo de articulo los productos de un cliente. (agregue array cargando los productos del cliente.)
    '                Despliego de a 10 productos de un cliente.
    '12/12/00 junte exe de gallinal cargando una variable al inicio que la llame idSucGallinal
    '5-11-02: agregue porcentaje a retiros y pongo siempre cero a liq. a camión.
    '18-01-05   Elimine AAGeneral y agregue opción de ingresar un nuevo motivo.
    '27-01-05   Deje mal impresión apuntaba a celda 1 en vsmotivo, cambie a modConnect y elimine cargadeimpresión vieja.
    '29-9-06  campos nuevos SerCliente..
    
    '18-3-07    Agregamos Campos Comentario Interno.
    '14-04-07   Agregamos código de barras.
    '12/10/2007 eliminé consulta de valorflete y puse las nuevas tablas.
    '11-3-08    Si ingresa un producto que no tiene el servicio cumplido lo cumplo.
    
Option Explicit
'Parametros para seleccionar tipo de DatosIngreso
'prmTipoIngreso    Taller = 1,    Retiro = 2,    Visita = 3
'prmInvocacion Ingreso de servicio para nota con cliente CGSA
'                       Si es T va directo a taller, si es R va a retiro.
'prmArticulo trae el ID de artículo que se le asignará al servicio
'prmDireccion id de direccion del cliente que devuelve el artículo.

Private Enum TipoCliente
    Cliente = 1
    Empresa = 2
End Enum


Public prmTipoIngreso As Integer, prmInvocacion As String, prmArticulo As Long, prmDireccion As Long, prmSucursal As String
Private strCierre As String, douHabilitado As Double, douAgenda As Double
Private gCliente As Long, gTipoCliente As Integer
Private sEsProducto As Boolean

Private bolEOF As Boolean
Private Const CantTuplas = 10
Private Function fnc_GetTamañoArticulo() As Long
On Error Resume Next
Dim rsT As rdoResultset
        Set rsT = cBase.OpenResultset("Select ArtTamaño From Articulo Where ArtID = " & vsProducto.Cell(flexcpData, vsProducto.Row, 0) & " And ArtTamaño Is Not Null", _
                                                    rdOpenDynamic, rdConcurValues)
        If Not rsT.EOF Then fnc_GetTamañoArticulo = rsT(0)
        rsT.Close
End Function

Private Sub loc_DefinoPrecioFlete(ByVal iTipoFlete As Long, ByVal iZona As Long, ByRef cValorFlete As Currency, ByRef cLiquidar As Currency)
On Error GoTo errDPF
Dim iTam As Long
    
    iTam = fnc_GetTamañoArticulo
    
    Cons = "SELECT Top 1 PFLPrecioPpal, PFLCostoPpal FROM PrecioFlete, GrupoZonaZona" & _
        " WHERE PFlTipoFlete = " & iTipoFlete & " AND GZZZona = " & iZona & _
        " AND  PFlGrupoZona = GZZGrupo"
    
    If iTam > 0 Then Cons = Cons & " And PFlTamañoArt = " & iTam
    
    Cons = Cons & " Order by PFLPrecioPpal"
    
    'PFlTamañoArt
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux("PFLPrecioPpal")) Then cValorFlete = RsAux("PFLPrecioPpal")
        If Not IsNull(RsAux("PFLCostoPpal")) Then cValorFlete = RsAux("PFLCostoPpal")
    Else
        If iTam > 0 Then
            'busco el menor precio.
            RsAux.Close
            Cons = "SELECT Top 1 PFLPrecioPpal, PFLCostoPpal FROM PrecioFlete, GrupoZonaZona" & _
                " WHERE PFlTipoFlete = " & iTipoFlete & " AND GZZZona = " & iZona & _
                " AND  PFlGrupoZona = GZZGrupo Order by PFLPrecioPpal"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                If Not IsNull(RsAux("PFLPrecioPpal")) Then cValorFlete = RsAux("PFLPrecioPpal")
                If Not IsNull(RsAux("PFLCostoPpal")) Then cValorFlete = RsAux("PFLCostoPpal")
            End If
        End If
    End If
    RsAux.Close
    Exit Sub
errDPF:
    clsGeneral.OcurrioError "Error al buscar el costo del flete.", Err.Description, "Definir precio del flete"
End Sub

Private Sub bDireccionP_KeyDown(KeyCode As Integer, Shift As Integer)
    If Val(tDireccion.Tag) <> 0 And Shift = 0 And KeyCode = 93 Then sEsProducto = True: PopupMenu MnuOpDireccion, , bDireccionP.Left, bDireccionP.Top, MnuOpDiTitulo
End Sub

Private Sub bDireccionP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Val(tDireccion.Tag) <> 0 Then sEsProducto = True: PopupMenu MnuOpDireccion, , bDireccionP.Left, bDireccionP.Top, MnuOpDiTitulo
End Sub

Private Sub bDireccionT_KeyDown(KeyCode As Integer, Shift As Integer)
    If gCliente <> 0 And Shift = 0 And KeyCode = 93 Then sEsProducto = False: PopupMenu MnuOpDireccion, , bDireccionT.Left, bDireccionT.Top, MnuOpDiTitulo
End Sub

Private Sub bDireccionT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If gCliente <> 0 Then sEsProducto = False: PopupMenu MnuOpDireccion, , bDireccionT.Left, bDireccionT.Top, MnuOpDiTitulo
End Sub

Private Sub cDatoIngreso_Click()
    
    If TabRecepcion.SelectedItem.Index = 1 Then Exit Sub
    Select Case cDatoIngreso.ItemData(cDatoIngreso.ListIndex)
        Case 1: AjustoFichaTaller
        Case 2: AjustoFichaRetiro
        Case 3: AjustoFichaVisita
    End Select
    
End Sub

Private Sub cDatoIngreso_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tCi
End Sub

Private Sub chCoordinarEntrega_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tAclaracion
End Sub

Private Sub chEsReclamo_Click()
    If chEsReclamo.Value = 0 Then
        tReclamo.Enabled = False: tReclamo.BackColor = Inactivo: tReclamo.Text = ""
    Else
        tReclamo.Enabled = True: tReclamo.BackColor = Blanco
    End If
End Sub

Private Sub chEsReclamo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tReclamo.Enabled Then tReclamo.SetFocus Else Foco chCoordinarEntrega
    End If
End Sub

Private Sub cRCamion_GotFocus()
    With cRCamion
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cRCamion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cRCamion.ListIndex > -1 Then Foco tRFecha
End Sub
Private Sub cRCamion_LostFocus()
    cRCamion.SelStart = 0
End Sub

Private Sub cRDeposito_GotFocus()
    With cRDeposito
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cRDeposito_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cRDeposito.ListIndex > -1 Then
        If chEsReclamo.Value = 1 And tReclamo.Enabled Then tReclamo.SetFocus Else Foco tAclaracion
    End If
End Sub

Private Sub cRFactura_Click()
    If cRFactura.ListIndex > -1 Then
        If cRFactura.ItemData(cRFactura.ListIndex) = FacturaServicio.Camion Then tRLiquidar.Text = "0.00" 'Else tRLiquidar.Text = tRImporte.Text
    End If
End Sub

Private Sub cRFactura_GotFocus()
    With cRFactura
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub cRFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cRFactura.ListIndex > -1 Then tRLiquidar.SetFocus
End Sub
Private Sub cRFactura_LostFocus()
    cRFactura.SelStart = 0
End Sub

Private Sub cRHora_GotFocus()
    With cRHora
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cRHora_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cRHora.ListIndex = -1 Then cRHora.Text = ValidoRangoHorario(cRHora.Text)
        Foco cRMoneda
    End If
End Sub
Private Sub cRHora_LostFocus()
    cRHora.SelStart = 0
End Sub

Private Sub cRMoneda_GotFocus()
    With cRMoneda
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub cRMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cRMoneda.ListIndex > -1 Then Foco tRImporte
End Sub
Private Sub cRMoneda_LostFocus()
    cRMoneda.SelStart = 0
End Sub

Private Sub cTalDeposito_GotFocus()
    With cTalDeposito
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cTalDeposito_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cTalDeposito.ListIndex > -1 Then
        If chEsReclamo.Value = 1 And tReclamo.Enabled Then tReclamo.SetFocus Else Foco tAclaracion
    End If
End Sub
Private Sub cTalDeposito_LostFocus()
    cTalDeposito.SelStart = 0
End Sub

Private Sub cTipoFlete_Click()
    If cTipoFlete.ListIndex > -1 Then BuscoValorFlete
End Sub
Private Sub cTipoFlete_Change()
    'Cuando cambia el tipo de flete borro las variables que tenía cargada para el mismo
    douAgenda = 0
    douHabilitado = 0
    strCierre = ""
'    If cTipoFlete.ListIndex > -1 Then BuscoDatosTipoDeFlete cTipoFlete.ItemData(cTipoFlete.ListIndex)
End Sub
Private Sub cTipoFlete_GotFocus()
    With cTipoFlete
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    ArmoXDefectoRetiro
End Sub
Private Sub cTipoFlete_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cTipoFlete.ListIndex > -1 Then
        If cTipoFlete.ListIndex > -1 Then BuscoDatosTipoDeFlete cTipoFlete.ItemData(cTipoFlete.ListIndex)
        Foco cRCamion
    End If
End Sub
Private Sub cTipoFlete_LostFocus()
    cTipoFlete.SelStart = 0
End Sub

'Private Sub cTipoTelefono_GotFocus()
'    With cTipoTelefono
'        .SelStart = 0: .SelLength = Len(.Text)
'    End With
'End Sub
'
'Private Sub cTipoTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        If cTipoTelefono.ListIndex > -1 Then
'            'Cargo el telefono para este tipo
'            Cons = "Select * From Telefono Where TelCliente = " & gCliente _
'                & " And TelTipo = " & cTipoTelefono.ItemData(cTipoTelefono.ListIndex)
'            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'            If Not RsAux.EOF Then
'                tTelefono.Text = clsGeneral.RetornoFormatoTelefono(cBase, RsAux!TelNumero, Val(tDirCliente.Tag))
'                If Not IsNull(RsAux!TelInterno) Then tInterno.Text = Trim(RsAux!TelInterno)
'            End If
'            RsAux.Close
'        End If
'        Foco tTelefono
'    End If
'End Sub
'
'Private Sub cTipoTelefono_LostFocus()
'    cTipoTelefono.SelStart = 0
'End Sub

Private Sub cVCamion_GotFocus()
    With cVCamion
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cVCamion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cVCamion.ListIndex > -1 Then
        ArmoTipoFleteVisita cVCamion.ItemData(cVCamion.ListIndex)
        Foco tVFecha
    End If
End Sub
Private Sub cVCamion_LostFocus()
    cVCamion.SelStart = 0
End Sub

Private Sub cVComentario_GotFocus()
    With cVComentario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cVComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If chEsReclamo.Value = 1 And tReclamo.Enabled Then tReclamo.SetFocus Else Foco tAclaracion
    End If
End Sub

Private Sub cVComentario_LostFocus()
    cVComentario.SelStart = 0
End Sub

Private Sub cVFactura_Click()
    If cVFactura.ListIndex > -1 Then
        If cVFactura.ItemData(cVFactura.ListIndex) = FacturaServicio.Camion Then tVLiquidar.Text = "0.00" Else tVLiquidar.Text = tVImporte.Text
    End If
End Sub
Private Sub cVFactura_GotFocus()
    With cVFactura
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cVFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cVFactura.ListIndex > -1 Then Foco tVLiquidar
End Sub
Private Sub cVFactura_LostFocus()
    cVFactura.SelStart = 0
End Sub

Private Sub cVHora_GotFocus()
    With cVHora
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cVHora_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cVHora.ListIndex = -1 Then cVHora.Text = ValidoRangoHorario(cVHora.Text)
        Foco cVMoneda
    End If
End Sub

Private Sub cVMoneda_GotFocus()
    With cVMoneda
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cVMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cVMoneda.ListIndex > -1 Then Foco tVImporte
End Sub
Private Sub cVMoneda_LostFocus()
    cVMoneda.SelStart = 0
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = 0: Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    
    Status.Panels("printer").Text = paPrintConfD
    Me.Height = 8550
    
    FechaDelServidor
    LimpioTodo
    InicializoGrillaMotivos
    InicializoGrillaProducto
    InicializoGrillaHistoria
    
    PicTaller.BorderStyle = 0: PicRetiro.BorderStyle = 0: PicVisita.BorderStyle = 0: picHistoria.BorderStyle = 0
    
    With TabRecepcion
        Set .ImageList = Image1
        .Tabs("historia").Image = Image1.ListImages("historia").Index
        .Tabs("comodin").Image = Image1.ListImages("taller").Index
    End With
    
    With cDatoIngreso
        .Clear
        .AddItem "Taller": .ItemData(.NewIndex) = 1
        .AddItem "Retiro": .ItemData(.NewIndex) = 2
        .AddItem "Visita": .ItemData(.NewIndex) = 3
    End With
'
'    Cons = "Select TTeCodigo, TTeNombre From TipoTelefono Order by TTeNombre"
'    CargoCombo Cons, cTipoTelefono, ""
'
    'Defino la acción que toma el formulario.
    If prmInvocacion <> "" Then     'Acción para realizar una Nota.
        If prmInvocacion = "T" Then
            prmTipoIngreso = 1 'Por las dudas que nos equivoquemos.
        ElseIf prmInvocacion = "R" Then
            prmTipoIngreso = 2
        Else
            prmTipoIngreso = 3
        End If
    End If
    
    If prmTipoIngreso <> 0 Then BuscoCodigoEnCombo cDatoIngreso, CLng(prmTipoIngreso) Else cDatoIngreso.ListIndex = 0
    
    If prmInvocacion <> "" And prmArticulo > 0 Then
    
        gCliente = paClienteEmpresa
        LimpioFichaCliente
        CargoDatosCliente paClienteEmpresa       'Cargo Datos del Cliente Seleccionado
        
        'Hago copia de la dirección del cliente.
        If prmDireccion > 0 Then prmDireccion = CopioDireccion(prmDireccion)
        
        'Inserto un artículo a la empresa.
        Cons = "Insert into Producto (ProArticulo, ProCliente, ProDireccion, ProFModificacion) Values (" _
            & prmArticulo & ", " & paClienteEmpresa & ", "
        If prmDireccion > 0 Then Cons = Cons & prmDireccion & ", " Else Cons = Cons & " Null, "
        Cons = Cons & "'" & Format(gFechaServidor, sqlFormatoFH) & "')"
        cBase.Execute (Cons)
        CargoDatosProducto gCliente
    End If
    
    picHistoria.ZOrder 0
    PrueboBandejaImpresora
    
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Error al iniciar el formulario.", Trim(Err.Description)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    TabRecepcion.Width = Me.ScaleWidth - (TabRecepcion.Left * 2)
    TabRecepcion.Height = tReclamo.Top - (TabRecepcion.Top + 40)
    
    PicTaller.Height = TabRecepcion.ClientHeight
    PicRetiro.Height = PicTaller.Height
    PicVisita.Height = PicRetiro.Height
    picHistoria.Height = PicRetiro.Height
    
    PicTaller.Width = TabRecepcion.Width - 100
    PicRetiro.Width = PicTaller.Width: PicVisita.Width = PicTaller.Width: picHistoria.Width = PicTaller.Width
    
    PicTaller.Top = TabRecepcion.ClientTop: PicTaller.Left = TabRecepcion.ClientLeft
    PicRetiro.Top = PicTaller.Top: PicRetiro.Left = TabRecepcion.ClientLeft
    PicVisita.Top = PicTaller.Top: PicVisita.Left = TabRecepcion.ClientLeft
    picHistoria.Top = PicTaller.Top: picHistoria.Left = PicVisita.Left
    
    vsHistoria.Height = picHistoria.Height - 120
    vsHistoria.Width = picHistoria.Width - (vsHistoria.Left * 2)
    AjustoFichaInicial  'Coloca la grilla de motivos y textos comunes en el tab.
    vsMotivos.Height = PicTaller.Height + PicTaller.Top - vsMotivos.Top
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    'GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
End Sub

Private Sub Label10_Click()
    Foco cRCamion
End Sub

Private Sub Label12_Click()
    Foco tRFecha
End Sub
Private Sub Label13_Click()
    Foco cVCamion
End Sub
Private Sub Label15_Click()
    Foco tFCompra
End Sub
Private Sub Label16_Click()
    Foco tfSerie
End Sub
Private Sub Label17_Click()
    Foco bDireccionP
End Sub
Private Sub Label18_Click()
    Foco tPSerie
End Sub
Private Sub Label2_Click()
    Foco cTalDeposito
End Sub
Private Sub Label20_Click()
    Foco cVMoneda
End Sub
Private Sub Label21_Click()
    Foco cVFactura
End Sub
Private Sub Label22_Click()
    Foco tAclaracion
End Sub
Private Sub Label11_Click()
    Foco tComentarioInterno
End Sub

Private Sub Label23_Click()
    Foco tRLiquidar
End Sub
Private Sub Label24_Click()
    Foco tVLiquidar
End Sub

Private Sub Label25_Click()
    Foco cRDeposito
End Sub

Private Sub Label3_Click()
    Foco cVComentario
End Sub

Private Sub Label4_Click()
    tCi.SetFocus
End Sub
Private Sub Label5_Click()
    Foco cTipoFlete
End Sub
Private Sub Label6_Click()
    Foco tVFecha
End Sub
Private Sub Label7_Click()
    Foco cRFactura
End Sub
Private Sub Label8_Click()
    Foco cRMoneda
End Sub
Private Sub lNDireccion_Click()
    Foco bDireccionT
End Sub
Private Sub lRtMotivo_Click()
    Foco tMotivo
End Sub
Private Sub lTipoProducto_Click()
    Foco tArticulo
End Sub
Private Sub ltMotivo_Click()
    Foco tMotivo
End Sub
Private Sub ltUsuario_Click()
    Foco tUsuario
End Sub
Private Sub lVtMotivo_Click()
    Foco tMotivo
End Sub

Private Sub MnuAyuHelp_Click()
On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    Cons = "Select * from Aplicacion Where AplNombre = '" & Trim(App.Title) & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux!AplHelp) Then aFile = Trim(RsAux!AplHelp)
    RsAux.Close
    
    If aFile <> "" Then EjecutarApp aFile
    Screen.MousePointer = 0
    Exit Sub
    
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuBuCGSA_Click()
    CargoClienteEmpresa
End Sub

Private Sub MnuBuFichaCliente_Click()
    FichaCliente gTipoCliente
End Sub

Private Sub MnuBuNuevaEmpresa_Click()
    NuevoCliente TipoCliente.Empresa
End Sub

Private Sub MnuBuNuevoCliente_Click()
    NuevoCliente TipoCliente.Cliente
End Sub

Private Sub MnuOpDiConfirmar_Click()
    ConfirmarDireccion sEsProducto
End Sub

Private Sub MnuOpDiModificar_Click()
On Error GoTo errDir
    Screen.MousePointer = 11
    
    Dim aCodDir As Long, idDireccion As Long
    Dim objDir As New clsDireccion
    
    If sEsProducto Then idDireccion = Val(bDireccionP.Tag) Else idDireccion = Val(tDirCliente.Tag)
    
    If sEsProducto Then
        objDir.ActivoFormularioDireccion cBase, idDireccion, gCliente, "Producto", "ProDireccion", "ProCodigo", CLng(tDireccion.Tag)
    Else
        objDir.ActivoFormularioDireccion cBase, idDireccion, gCliente, "Cliente", "CliDireccion", "CliCodigo", gCliente
    End If
    
    Me.Refresh
    aCodDir = objDir.CodigoDeDireccion
    Set objDir = Nothing
    
    If idDireccion = 0 And aCodDir <> 0 Then
        If Not sEsProducto Then
            Cons = "Update Cliente Set CliDireccion = " & aCodDir & " Where CliCodigo = " & gCliente
        Else
            Cons = "Update Producto Set ProDireccion = " & aCodDir & " Where ProCodigo = " & CLng(tDireccion.Tag)
        End If
        cBase.Execute Cons
    End If
    If Not sEsProducto Then
        tDirCliente.Tag = aCodDir: tDirCliente.Text = ""
        If aCodDir <> 0 Then tDirCliente.Text = clsGeneral.ArmoDireccionEnTexto(cBase, aCodDir, Departamento:=True, Localidad:=True, Zona:=True, ConfYVD:=True)
    Else
        bDireccionP.Tag = aCodDir: tDireccion.Text = ""
        If aCodDir <> 0 Then tDireccion.Text = clsGeneral.ArmoDireccionEnTexto(cBase, aCodDir, Departamento:=True, Localidad:=True, Zona:=True, ConfYVD:=True)
        vsProducto.Cell(flexcpData, vsProducto.Row, 2) = BuscoZonaDireccion(aCodDir)
    End If
    Screen.MousePointer = 0
    Exit Sub
errDir:
    clsGeneral.OcurrioError "Ocurrió un error al editar la dirección", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuProCederProducto_Click()
    EjecutarApp App.Path & "\Ceder Producto", CStr(gCliente), True
    Me.Refresh
    DeshabilitoIngreso
    CargoDatosProducto gCliente
End Sub

Private Sub MnuProCumplirServicio_Click()
    If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) > 0 Then
        CumploServicio Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4))
    End If
End Sub

Private Sub MnuProFicha_Click()
Dim sPrm As String
    If vsProducto.Rows = 1 Then
        EjecutarApp App.Path & "\Productos", "N" & CStr(gCliente), True
    Else
        If vsProducto.Row >= 1 Then
            sPrm = ";a" & vsProducto.Cell(flexcpData, vsProducto.Row, 0) & ";p" & vsProducto.Cell(flexcpValue, vsProducto.Row, 0)
        End If
        EjecutarApp App.Path & "\Productos", CStr(gCliente) & sPrm, True
    End If
    Me.Refresh
    DeshabilitoIngreso
    'inicializo el tag para que cargue todos de nuevo.
    vsProducto.Tag = "0"
    CargoDatosProducto gCliente
End Sub

Private Sub MnuProFindSerie_Click()
    'Busco por Nº de serie algún producto del cliente.
On Error GoTo errFind
Dim sSerie As String, idPro As Long, idCli As Long

    sSerie = InputBox("Ingrese el nº de serie del producto.", "Buscar por nº de serie")
    If sSerie <> "" Then
        Screen.MousePointer = 11
        Cons = "Select ProCodigo, ArtNombre, ProNroSerie, ProCliente From Producto, Articulo " _
            & " Where ProNroSerie = '" & sSerie & "' And ProArticulo = ArtID"
        
        
        'ProCliente = " & gCliente
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            RsAux.Close
            
            Cons = "SELECT DocCliente FROM ProductosVendidos INNER JOIN Documento on PVeDocumento = DocCodigo" & _
                    " WHERE PVeNSerie = '" & sSerie & "' ORDER BY PveDocumento desc"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                MsgBox "PRODUCTO SIN SERVICIO, se cargará el cliente, debe validar la información", vbExclamation, "ATENCIÓN"
                idCli = RsAux(0)
                RsAux.Close
                BuscoClienteSeleccionado idCli
                Call MnuProFicha_Click
                Exit Sub
            End If
            RsAux.Close
            MsgBox "No se encontró un producto para el cliente y serie ingresada.", vbInformation, "ATENCIÓN"
        Else
            RsAux.MoveNext
            If RsAux.EOF Then
                RsAux.MoveFirst
                idPro = RsAux!ProCodigo
                idCli = RsAux("ProCliente")
                RsAux.Close
                vsProducto.Tag = 0
                
                If gCliente = 0 Or gCliente <> idCli Then
                    gCliente = idCli
                    BuscoClienteSeleccionado idCli, idPro
                Else
                    CargoDatosProducto gCliente, idPro
                End If
            Else
                RsAux.Close
                Dim objHelp As New clsListadeAyuda
                If objHelp.ActivarAyuda(cBase, Cons, 4800, 0, "Productos") > 0 Then
                    vsProducto.Tag = 0
                    If gCliente = 0 Then gCliente = objHelp.RetornoDatoSeleccionado(3)
                    CargoDatosProducto gCliente, objHelp.RetornoDatoSeleccionado(0)
                End If
                Set objHelp = Nothing
            End If
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
errFind:
    clsGeneral.OcurrioError "Ocurrió el siguiente error al buscar por nº de serie.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuProHistoria_Click()
    If vsProducto.Row > 0 Then EjecutarApp App.Path & "\Historia Servicio", vsProducto.Cell(flexcpValue, vsProducto.Row, 0)
    Me.Refresh
End Sub

Private Sub MnuProSeguimientoReporte_Click()
    If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) = 0 Then
        EjecutarApp App.Path & "\Seguimiento de Servicios", vsProducto.Cell(flexcpData, vsProducto.Row, 4)
    Else
        EjecutarApp App.Path & "\Seguimiento de Servicios", ""
    End If
End Sub

Private Sub MnuSalDel_Click()
    Unload Me
End Sub

Private Sub Status_PanelClick(ByVal Panel As ComctlLib.Panel)
    If "printer" = Panel.Key Then
        prj_GetPrinter True
        Panel.Text = paPrintConfD
    End If
End Sub

Private Sub TabRecepcion_Click()
    If TabRecepcion.SelectedItem.Index = 1 Then
        'HISTORIA
        picHistoria.ZOrder 0
    Else
        If cDatoIngreso.ListIndex = -1 Then Exit Sub
        'Ficha de Recepción
        Select Case cDatoIngreso.ItemData(cDatoIngreso.ListIndex)
            Case 1: AjustoFichaTaller
            Case 2: AjustoFichaRetiro: ArmoXDefectoRetiro
            Case 3: AjustoFichaVisita: ArmoXDefectoVisita
        End Select
    End If
End Sub

Private Sub TabRecepcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If TabRecepcion.SelectedItem.Index = 2 Then
            Foco tMotivo
            If cDatoIngreso.ListIndex = -1 Then Exit Sub
            'Ficha de Recepción
            Select Case cDatoIngreso.ItemData(cDatoIngreso.ListIndex)
                Case 1: AjustoFichaTaller
                Case 2: AjustoFichaRetiro: ArmoXDefectoRetiro
                Case 3: AjustoFichaVisita: ArmoXDefectoVisita
            End Select
        End If
    End If
End Sub

Private Sub tAclaracion_GotFocus()
    With tAclaracion
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tAclaracion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentarioInterno
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = ""
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrTA
Dim fModificacion As Date

    If KeyCode = vbKeyReturn And Trim(tArticulo.Text) <> "" And Val(tDireccion.Tag) > 0 Then
        
        If Val(tArticulo.Tag) = Val(lTipoProducto.Tag) Then Foco tFCompra: Exit Sub
        
        Screen.MousePointer = 11
        FechaDelServidor
        
        If Not IsNumeric(tArticulo.Text) Then   'Busqueda por nombre
            
            Cons = "Select ArtID, 'Nombre' = ArtNombre, 'Código' = ArtCodigo From Articulo " _
                    & " Where ArtNombre Like '" & tArticulo.Text & "%'" _
                    & " Order by ArtNombre"
            
            Dim LiAyuda  As New clsListadeAyuda
            If LiAyuda.ActivarAyuda(cBase, Cons, 4800, 1, "Artículos") > 0 Then
                Me.Refresh
                
                If Val(lTipoProducto.Tag) <> Val(LiAyuda.RetornoDatoSeleccionado(0)) Then
                    If MsgBox("¿Confirma modificar el tipo del artículo?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Set LiAyuda = Nothing: GoTo Abandono
                End If
                            
                tArticulo.Text = LiAyuda.RetornoDatoSeleccionado(1)
                tArticulo.Tag = LiAyuda.RetornoDatoSeleccionado(0)
                
                
                On Error GoTo ErrMA
                Cons = "Select * From Producto Where ProCodigo = " & Val(tDireccion.Tag)
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If RsAux!ProFModificacion <> CDate(vsProducto.Cell(flexcpData, vsProducto.Row, 1)) Then
                    RsAux.Close
                    MsgBox "El artículo fue modificado por otra terminal, verifique.", vbExclamation, "ATENCIÓN"
                    Set LiAyuda = Nothing: GoTo Abandono
                End If
                RsAux.Edit
                RsAux!ProArticulo = Val(tArticulo.Tag)
                fModificacion = gFechaServidor
                RsAux!ProFModificacion = Format(fModificacion, sqlFormatoFH)
                RsAux.Update
                RsAux.Close
                
                'Updateo en la grilla el tipo y el id de artículo.
                With vsProducto
                    .Cell(flexcpText, .Row, 1) = tArticulo.Text
                    .Cell(flexcpData, .Row, 0) = tArticulo.Tag
                    .Cell(flexcpData, .Row, 1) = fModificacion
                    .Cell(flexcpText, .Rows - 1, 2) = EstadoProducto(CInt(CalculoEstadoProducto(.Cell(flexcpText, .Rows - 1, 0))), True)
                    .Cell(flexcpText, .Rows - 1, 4) = RetornoGarantia(tArticulo.Tag)
                End With
                lTipoProducto.Tag = tArticulo.Tag   'Me quedo con el tipo de artículo.
                Foco tFCompra
            Else
                Me.Refresh
            End If
            Set LiAyuda = Nothing
        Else                                            'Busqueda por codigo
            Cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtCodigo = " & Val(tArticulo.Text)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
            If RsAux.EOF Then
                RsAux.Close
                MsgBox "No se encontró un artículo para el código ingresado.", vbInformation, "ATENCIÓN"
            Else
                If Val(lTipoProducto.Tag) <> RsAux("ArtID") Then
                    If MsgBox("¿Confirma modificar el tipo del artículo a '" & Trim(RsAux("Nombre")) & "'?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                        RsAux.Close
                        GoTo Abandono
                    End If
                End If
                tArticulo.Text = Trim(RsAux!Nombre)
                tArticulo.Tag = RsAux!ArtID
                RsAux.Close
                
                On Error GoTo ErrMA
                Cons = "Select * From Producto Where ProCodigo = " & Val(tDireccion.Tag)
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If RsAux!ProFModificacion <> CDate(vsProducto.Cell(flexcpData, vsProducto.Row, 1)) Then
                    RsAux.Close
                    MsgBox "El artículo fue modificado por otra terminal, verifique.", vbExclamation, "ATENCIÓN"
                    Set LiAyuda = Nothing: GoTo Abandono
                End If
                RsAux.Edit
                RsAux!ProArticulo = Val(tArticulo.Tag)
                fModificacion = gFechaServidor
                RsAux!ProFModificacion = Format(fModificacion, sqlFormatoFH)
                RsAux.Update
                RsAux.Close
                
                'Updateo en la grilla el tipo y el id de artículo.
                With vsProducto
                    .Cell(flexcpText, .Row, 1) = tArticulo.Text
                    .Cell(flexcpData, .Row, 0) = tArticulo.Tag
                    .Cell(flexcpData, .Row, 1) = fModificacion
                    .Cell(flexcpText, .Rows - 1, 2) = EstadoProducto(CInt(CalculoEstadoProducto(.Cell(flexcpText, .Rows - 1, 0))), True)
                    .Cell(flexcpText, .Rows - 1, 4) = RetornoGarantia(tArticulo.Tag)
                End With
                lTipoProducto.Tag = tArticulo.Tag   'Me quedo con el tipo de artículo.
                Foco tFCompra
            End If
        End If
        
        Screen.MousePointer = 0
    Else
        If KeyCode = vbKeyReturn Then Foco tFCompra
    End If
    Exit Sub

Abandono:
    CargoProductoParaEditar (tDireccion.Tag): Exit Sub
    
ErrTA:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0: Exit Sub
ErrMA:
    Set LiAyuda = Nothing
    clsGeneral.OcurrioError "Ocurrió un error al intentar modificar el tipo de artículo.", Err.Description
    CargoProductoParaEditar (tDireccion.Tag): Screen.MousePointer = 0: Exit Sub

End Sub

Private Sub tCi_GotFocus()
    tCi.SelStart = 0: tCi.SelLength = 11
End Sub

Private Sub tCi_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case 93: PopupMenu MnuBuscar, , tCi.Left + (tCi.Width / 2), (tCi.Top + tCi.Height) - (tCi.Height / 2)
        Case vbKeyF2: If Shift = 0 Then FichaCliente TipoCliente.Cliente
        Case vbKeyF3: If Shift = 0 Then NuevoCliente TipoCliente.Cliente
        Case vbKeyF4: If Shift = 0 Then BuscarClientes TipoCliente.Cliente
        Case vbKeyF11: If Shift = 0 Then CargoClienteEmpresa
    End Select
    
End Sub

Private Sub TCI_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        
        Dim aCi As String
        Screen.MousePointer = 11
        'If Len(tCi.Text) = 7 Then tCi.Text = clsGeneral.AgregoDigitoControlCI(tCi.Text)
                
        'Valido la Cédula ingresada----------
        If Trim(tCi.Text) <> "" Then
            If Len(tCi.Text) <> 8 Then
                Screen.MousePointer = 0
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            If Not clsGeneral.CedulaValida(tCi.Text) Then
                Screen.MousePointer = 0
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        End If
        
        'Busco el Cliente -----------------------
        If Trim(tCi.Text) <> "" Then
            LimpioTodo
            gCliente = BuscoClienteCIRUC(tCi.Text)
            If gCliente = 0 Then
                LimpioTodo
                Screen.MousePointer = 0
                MsgBox "No existe un cliente para la cédula ingresada.", vbExclamation, "ATENCIÓN"
            Else
                 BuscoClienteSeleccionado gCliente
            End If
        Else
            tRuc.SetFocus
        End If
        Screen.MousePointer = 0
    End If

End Sub

Private Function BuscoClienteCIRUC(CiRuc As String)

    On Error GoTo errBuscar
    BuscoClienteCIRUC = 0
    Cons = "Select * from Cliente Where CliCiRuc = '" & Trim(CiRuc) & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then BuscoClienteCIRUC = RsAux!CliCodigo
    RsAux.Close
    Exit Function

errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el cliente."
    Screen.MousePointer = 0
End Function

Private Sub tComentarioInterno_GotFocus()
    With tComentarioInterno
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tComentarioInterno_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tTelefonoServicio
    
End Sub

Private Sub tFCompra_Change()
    tFCompra.Tag = "-1"
End Sub

Private Sub tFCompra_GotFocus()
    With tFCompra
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tFCompra_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tFCompra.Text) = Trim(tFCompra.Tag) Then Foco tfSerie: Exit Sub
        
        If Trim(tFCompra.Text) <> "" Then
            If Not IsDate(tFCompra.Text) Then
                MsgBox "No se ingresó un formato de fecha válido.", vbExclamation, "ATENCIÓN": Exit Sub
            Else
                tFCompra.Text = Format(tFCompra.Text, "dd/mm/yyyy")
            End If
        End If
        
        If MsgBox("¿Confirma modificar la fecha de compra del producto?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
            tFCompra.Text = Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 3))
            tFCompra.Tag = tFCompra.Text
        Else
            Screen.MousePointer = 11
            FechaDelServidor
            Dim fModificacion  As Date
            On Error GoTo ErrMA
            Cons = "Select * From Producto Where ProCodigo = " & Val(tDireccion.Tag)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If RsAux!ProFModificacion <> CDate(vsProducto.Cell(flexcpData, vsProducto.Row, 1)) Then
                RsAux.Close
                MsgBox "El artículo fue modificado por otra terminal, verifique.", vbExclamation, "ATENCIÓN"
                CargoProductoParaEditar Val(tDireccion.Tag): Screen.MousePointer = 0: Exit Sub
            End If
            RsAux.Edit
            If IsDate(tFCompra.Text) Then RsAux!ProCompra = Format(tFCompra.Text, sqlFormatoF) Else RsAux!ProCompra = Null
            fModificacion = gFechaServidor
            RsAux!ProFModificacion = Format(fModificacion, sqlFormatoFH)
            RsAux.Update
            RsAux.Close
            'Updateo en la grilla el tipo y el id de artículo.
            With vsProducto
                .Cell(flexcpData, .Row, 1) = fModificacion
                .Cell(flexcpText, .Row, 3) = tFCompra.Text
                .Cell(flexcpText, .Rows - 1, 2) = EstadoProducto(CInt(CalculoEstadoProducto(.Cell(flexcpText, .Rows - 1, 0))), True)
                .Cell(flexcpText, .Rows - 1, 4) = RetornoGarantia(tArticulo.Tag)
            End With
            
            Foco tfSerie
            Screen.MousePointer = 0
        End If
    End If
    Exit Sub
ErrMA:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al modificar la fecha de compra del producto.", Trim(Err.Description)
End Sub

Private Sub tfNumero_Change()
    tfNumero.Tag = "-11"
End Sub

Private Sub tfNumero_GotFocus()
    With tfNumero
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tfNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
        tfSerie.Text = UCase(tfSerie.Text)
        
        If Trim(tfSerie.Tag) = Trim(tfSerie.Text) And Trim(tfNumero.Tag) = Trim(tfNumero.Text) Then bDireccionP.SetFocus: Exit Sub
        
        If Trim(tfNumero.Text) <> "" Then If Not IsNumeric(tfNumero.Text) Then MsgBox "El formato ingresado no es numérico.", vbExclamation, "ATENCIÓN": Exit Sub
        
        If MsgBox("¿Confirma modificar los datos de la factura?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
            CargoProductoParaEditar Val(tDireccion.Tag): Screen.MousePointer = 0: Exit Sub
        Else
            Screen.MousePointer = 11
            FechaDelServidor
            Dim fModificacion  As Date
            On Error GoTo ErrMA
            Cons = "Select * From Producto Where ProCodigo = " & Val(tDireccion.Tag)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If RsAux!ProFModificacion <> CDate(vsProducto.Cell(flexcpData, vsProducto.Row, 1)) Then
                RsAux.Close
                MsgBox "El artículo fue modificado por otra terminal, verifique.", vbExclamation, "ATENCIÓN"
                CargoProductoParaEditar Val(tDireccion.Tag): Screen.MousePointer = 0: Exit Sub
            End If
            RsAux.Edit
            If Trim(tfSerie.Text) <> "" Then RsAux!ProFacturaS = Trim(tfSerie.Text) Else RsAux!ProFacturaS = Null
            If Trim(tfNumero.Text) Then RsAux!ProFacturaN = tfNumero.Text Else RsAux!ProFacturaN = Null
            fModificacion = gFechaServidor
            RsAux!ProFModificacion = Format(fModificacion, sqlFormatoFH)
            RsAux.Update
            RsAux.Close
            'Updateo en la grilla el tipo y el id de artículo.
            With vsProducto
                .Cell(flexcpData, .Row, 1) = fModificacion
                .Cell(flexcpText, .Row, 3) = tFCompra.Text
                If Trim(tfSerie.Text) <> "" Then .Cell(flexcpText, .Row, 6) = Trim(tfSerie.Text) & " " Else .Cell(flexcpText, .Row, 6) = ""
                If Trim(tfNumero.Text) <> "" Then .Cell(flexcpText, .Row, 6) = .Cell(flexcpText, .Row, 6) & Trim(tfNumero.Text)
            End With
            tfSerie.Tag = UCase(tfSerie.Text)
            tfNumero.Tag = tfNumero.Text
            bDireccionP.SetFocus
            Screen.MousePointer = 0
        End If
    End If
    Exit Sub
    
ErrMA:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al modificar los datos de la factura del producto .", Trim(Err.Description)
    
End Sub

Private Sub tfSerie_Change()
    tfSerie.Tag = "-1"
End Sub

Private Sub tfSerie_GotFocus()
    With tfSerie
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tfSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tfNumero
End Sub

'Private Sub tInterno_GotFocus()
'    With tInterno
'        .SelStart = 0: .SelLength = Len(.Text)
'    End With
'End Sub
'
'Private Sub tInterno_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then Foco tUsuario
'End Sub

Private Sub tMotivo_GotFocus()
    With tMotivo
        If .Text = "" Then .Text = "%"
        If .Text = "%" Then .SelStart = Len(.Text): Exit Sub
        .SelStart = 0
        .SelLength = Len(tMotivo.Text)
    End With
    Status.Panels(1).Text = "Ingrese parte o el nombre de un mótivo. [F3] Nuevo"
End Sub

Private Sub tMotivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyF3 Then
        Dim idTipo As Long
        idTipo = fnc_GetTipoArticulo
        EjecutarApp App.Path & "\Motivos de Servicio.exe", CStr(idTipo), True
    End If
    Me.Refresh
End Sub

Private Sub tMotivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(tMotivo.Text) = "" Then
            On Error Resume Next
            Select Case cDatoIngreso.ItemData(cDatoIngreso.ListIndex)
                Case 1: Foco cTalDeposito
                Case 2: Foco cTipoFlete
                Case 3: Foco cVCamion
            End Select
        Else
            On Error GoTo ErrBM
            Screen.MousePointer = 11
            Cons = "Select MSeID, Nombre = MSeNombre From MotivoServicio " _
                & "Where MSeTipo = (Select ArtTipo From Articulo Where ArtID = " & vsProducto.Cell(flexcpData, vsProducto.Row, 0) & ")" _
                & " And MSeNombre Like '" & Replace(tMotivo.Text, " ", "%") & "%'"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
            If RsAux.EOF Then
                Screen.MousePointer = 0: RsAux.Close
                MsgBox "No existen coincidencias para el tipo de artículo seleccionado y el dato ingresado .[F1] Ayuda", vbInformation, "ATENCIÓN"
                Exit Sub
            Else
                RsAux.MoveNext
                If RsAux.EOF Then
                    RsAux.MoveFirst
                    InsertoMotivoEnGrilla RsAux!MSeID, RsAux!Nombre
                Else
                    Dim objLista As New clsListadeAyuda
                    If objLista.ActivarAyuda(cBase, Cons, 4800, 1, "Motivos de Servicio") > 0 Then
                        InsertoMotivoEnGrilla objLista.RetornoDatoSeleccionado(0), objLista.RetornoDatoSeleccionado(1)
                    End If
                    Set objLista = Nothing
                    Me.Refresh
                End If
                RsAux.Close
            End If
            Screen.MousePointer = 0
            tMotivo.Text = ""
        End If
    End If
    Exit Sub
ErrBM:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar los motivos.", Trim(Err.Description)
End Sub

Private Sub tMotivo_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub tPSerie_Change()
    tPSerie.Tag = "-11"
End Sub

Private Sub tPSerie_GotFocus()
    With tPSerie
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tPSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If UCase(Trim(tPSerie.Tag)) = UCase(Trim(tPSerie.Text)) Then TabRecepcion.SetFocus: Exit Sub
        tPSerie.Text = UCase(tPSerie.Text)
        If MsgBox("¿Confirma modificar el N° de serie del producto?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
            CargoProductoParaEditar Val(tDireccion.Tag): Screen.MousePointer = 0: Exit Sub
        Else
            Screen.MousePointer = 11
            FechaDelServidor
            Dim fModificacion  As Date
            On Error GoTo ErrMA
            Cons = "Select * From Producto Where ProCodigo = " & Val(tDireccion.Tag)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If RsAux!ProFModificacion <> CDate(vsProducto.Cell(flexcpData, vsProducto.Row, 1)) Then
                RsAux.Close
                MsgBox "El artículo fue modificado por otra terminal, verifique.", vbExclamation, "ATENCIÓN"
                CargoProductoParaEditar Val(tDireccion.Tag): Screen.MousePointer = 0: Exit Sub
            End If
            RsAux.Edit
            If Trim(tPSerie.Text) <> "" Then RsAux!ProNroSerie = Trim(tPSerie.Text) Else RsAux!ProNroSerie = Null
            fModificacion = gFechaServidor
            RsAux!ProFModificacion = Format(fModificacion, sqlFormatoFH)
            RsAux.Update
            RsAux.Close
            'Updateo en la grilla el tipo y el id de artículo.
            With vsProducto
                .Cell(flexcpData, .Row, 1) = fModificacion
                .Cell(flexcpText, .Row, 3) = tFCompra.Text
                .Cell(flexcpText, .Row, 5) = Trim(tPSerie.Text)
            End With
            tPSerie.Tag = UCase(tPSerie.Text)
            bDireccionP.SetFocus
            Screen.MousePointer = 0
        End If
    End If
    Exit Sub
ErrMA:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al modificar los datos de la factura del producto .", Trim(Err.Description)
End Sub

Private Sub tReclamo_GotFocus()
    With tReclamo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tReclamo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tReclamo.Text <> "" Then
            If IsNumeric(tReclamo.Text) Then
                If tReclamo.Text <> tReclamo.Tag Then
                    'Busco si el id esta en la historia.
                    If Not ReclamoValido(tReclamo.Text) Then
                        MsgBox "El servicio que ud. ingreso no es un servicio del producto o el mismo fue anulado." & vbCrLf & "Verifique en la historia del producto.", vbInformation, "ATENCIÓN"
                        Exit Sub
                    End If
                End If
            Else
                tReclamo.Text = ""
                Exit Sub
            End If
        End If
        Foco tAclaracion
    End If
End Sub

Private Sub tRFecha_GotFocus()
    With tRFecha
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tRFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        
        Case vbKeyReturn
            If Not IsDate(tRFecha.Text) Then
                If Trim(tRFecha.Text) <> vbNullString Then
                    If IsDate(tRFecha.Tag) Then
                        CargoHoraEntregaParaDia cRHora, tRFecha.Tag
                        If cRHora.ListCount > 0 Then Foco cRHora
                    Else
                        MsgBox "La fecha ingresada no es correcta.", vbExclamation, "ATENCIÓN"
                    End If
                Else
                    tRFecha.Tag = tRFecha.Text
                    CargoHoraEntregaParaDia cRHora, tRFecha.Tag
                    If cRHora.ListCount > 0 Then Foco cRHora
                End If
            Else
                If IsDate(tRFecha.Text) Then tRFecha.Tag = tRFecha.Text: tRFecha.Text = Format(tRFecha.Tag, "ddd d/mm/yy")
                CargoHoraEntregaParaDia cRHora, tRFecha.Tag
                If cRHora.ListCount > 0 Then Foco cRHora
            End If
        
        Case vbKeyUp
                If IsDate(tRFecha.Tag) Then tRFecha.Tag = Format(CDate(tRFecha.Tag) + 1, "d-mm-yy")
                If IsDate(tRFecha.Tag) Then tRFecha.Text = Format(tRFecha.Tag, "ddd d/mm/yy")
        
        Case vbKeyDown
                If IsDate(tRFecha.Tag) Then tRFecha.Tag = Format(CDate(tRFecha.Tag) - 1, "d-mm-yy")
                If IsDate(tRFecha.Tag) Then tRFecha.Text = Format(tRFecha.Tag, "ddd d/mm/yy")
    
    End Select

End Sub

Private Sub tRFecha_LostFocus()
    If IsDate(tRFecha.Tag) Then tRFecha.Text = Format(tRFecha.Tag, "ddd d/mm/yy") Else tRFecha.Text = ""
End Sub

Private Sub tRImporte_GotFocus()
    With tRImporte
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tRImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsNumeric(tRImporte.Text) Then Foco cRFactura
End Sub

Private Sub tRImporte_LostFocus()
    If IsNumeric(tRImporte.Text) Then tRImporte.Text = Format(tRImporte.Text, FormatoMonedaP) Else tRImporte.Text = ""
End Sub

Private Sub tRLiquidar_GotFocus()
    With tRLiquidar
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tRLiquidar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cRDeposito
End Sub

Private Sub tRLiquidar_LostFocus()
    If IsNumeric(tRLiquidar.Text) Then tRLiquidar.Text = Format(tRLiquidar.Text, FormatoMonedaP) Else tRLiquidar.Text = "0.00"
End Sub

Private Sub tRuc_GotFocus()
    tRuc.SelStart = 0: tRuc.SelLength = 15
End Sub

Private Sub tRuc_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case 93: PopupMenu MnuBuscar, , tRuc.Left + (tRuc.Width / 2), (tRuc.Top + tCi.Height) - (tRuc.Height / 2), MnuBusquedas
        Case vbKeyF2: If Shift = 0 Then FichaCliente TipoCliente.Empresa
        Case vbKeyF3: If Shift = 0 Then NuevoCliente TipoCliente.Empresa
        Case vbKeyF4: If Shift = 0 Then BuscarClientes TipoCliente.Empresa
        Case vbKeyF11: If Shift = 0 Then CargoClienteEmpresa
    End Select
    
End Sub

Private Sub tRuc_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        
        If Trim(tRuc.Text) <> "" Then
            Screen.MousePointer = 11
            gCliente = BuscoClienteCIRUC(Trim(tRuc.Text))
            If gCliente = 0 Then
                Screen.MousePointer = 0
                MsgBox "No existe un cliente para el número de RUC ingresado.", vbExclamation, "ATENCIÓN"
            Else
                'Cargo Datos del Cliente Seleccionado------------------------------------------------
                 BuscoClienteSeleccionado gCliente
            End If
        Else
            tCi.SetFocus
        End If
        Screen.MousePointer = 0
    End If
    
End Sub

Private Sub BuscarClientes(aTipoCliente As Integer)
    
    Screen.MousePointer = 11
    Dim objBuscar As New clsBuscarCliente
    Dim aTipo As Integer, aCliente As Long
    
    If aTipoCliente = TipoCliente.Cliente Then objBuscar.ActivoFormularioBuscarClientes cBase, Persona:=True
    If aTipoCliente = TipoCliente.Empresa Then objBuscar.ActivoFormularioBuscarClientes cBase, Empresa:=True
    Me.Refresh
    aTipo = objBuscar.BCTipoClienteSeleccionado
    aCliente = objBuscar.BCClienteSeleccionado
    Set objBuscar = Nothing
    
    On Error GoTo errCargar
    If aCliente <> 0 Then
        gCliente = aCliente
        BuscoClienteSeleccionado gCliente
    End If
    
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub
Public Sub BuscoClienteSeleccionado(ByVal Codigo As Long, Optional selectPro As Long)

Dim aCliente As Long
    Screen.MousePointer = 11
    
    gCliente = Codigo
    gTipoCliente = 0
    LimpioFichaCliente
    If gCliente > 0 Then
        CargoDatosCliente Codigo         'Cargo Datos del Cliente Seleccionado
        CargoDatosProducto Codigo, selectPro    'Cargo los productos asociados al cliente.
        MsgClienteNoVender Codigo, True
        'Accedo a la grilla de productos.
        vsProducto.SetFocus
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errSolicitud:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos de la solicitud.", Err.Description
End Sub
Private Sub CargoDatosCliente(idCliente As Long)

    Cons = "Select * from Cliente " _
                & " Left Outer Join CPersona ON CliCodigo = CPeCliente " _
                & " Left Outer Join CEmpresa ON CliCodigo = CEmCliente " _
           & " Where CliCodigo = " & idCliente
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If Not RsAux.EOF Then       'CI o RUC
        Select Case RsAux!CliTipo
            Case TipoCliente.Cliente
                gTipoCliente = TipoCliente.Cliente
                If Not IsNull(RsAux!CliCiRuc) Then tCi.Text = clsGeneral.RetornoFormatoCedula(RsAux!CliCiRuc) Else tCi.Text = ""
                tCi.Tag = Trim(tCi.Text)
                tRuc.Text = "": tRuc.Tag = ""
                lTitular.Caption = Trim(Trim(Format(RsAux!CPeNombre1, "#")) & " " & Trim(Format(RsAux!CPeNombre2, "#"))) & ", " & Trim(Trim(Format(RsAux!CPeApellido1, "#")) & " " & Trim(Format(RsAux!CPeApellido2, "#")))
            Case TipoCliente.Empresa
                gTipoCliente = TipoCliente.Empresa
                If Not IsNull(RsAux!CliCiRuc) Then tRuc.Text = Trim(RsAux!CliCiRuc)
                tRuc.Tag = Trim(tRuc.Text)
                tCi.Text = "": tCi.Tag = ""
                If Not IsNull(RsAux!CEmNombre) Then lTitular.Caption = Trim(RsAux!CEmFantasia)
                If Not IsNull(RsAux!CEmFantasia) Then lTitular.Caption = lTitular.Caption & " (" & Trim(RsAux!CEmFantasia) & ")"
        End Select
        loc_FindComentarios idCliente
    Else
        tRuc.Text = "": tRuc.Tag = ""
        tCi.Text = "": tCi.Tag = ""
    End If
    
    'Direccion
    If Not IsNull(RsAux!CliDireccion) Then
        tDirCliente.Text = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion, Departamento:=True, Localidad:=True, Zona:=True, ConfYVD:=True)
        tDirCliente.Tag = RsAux!CliDireccion
    End If
    
    tTelCliente.Text = TelefonoATexto(idCliente)     'Telefonos
    
    RsAux.Close
    Exit Sub
    
errCliente:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente."
End Sub

Private Sub LimpioFichaCliente()
    'Datos del cliente.----------------------------
    lTitular.Caption = ""
    tDirCliente.Text = ""
    tTelCliente.Text = ""
    tRuc.Tag = ""
    tCi.Tag = ""
    gTipoCliente = 0
End Sub
Private Sub LimpioCamposProducto()
    tArticulo.Text = "": tArticulo.Tag = "": lTipoProducto.Tag = ""
    tFCompra.Text = "": tFCompra.Tag = ""
    tfSerie.Text = "": tfSerie.Tag = ""
    tfNumero.Text = "": tfNumero.Tag = ""
    tPSerie.Text = "": tPSerie.Tag = ""
    tDireccion.Text = "": tDireccion.Tag = ""
    bDireccionP.Tag = ""
End Sub

Private Sub MnuBEmpresas_Click()
    BuscarClientes TipoCliente.Empresa
End Sub

Private Sub MnuBPersonas_Click()
    BuscarClientes TipoCliente.Cliente
End Sub


Private Sub ConfirmarDireccion(sEsProducto As Boolean)
Dim idDireccion As Long

    If Not sEsProducto Then idDireccion = Val(tDirCliente.Tag) Else idDireccion = Val(bDireccionP.Tag)
    If idDireccion = 0 Then Exit Sub
    
    Dim aResp As Integer
    If sEsProducto Then
        aResp = MsgBox("Que desea realizar con la dirección del Producto." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
            & "Si - Confirmar dirección." & Chr(vbKeyReturn) _
            & "No - Eliminar confirmación de dirección." & Chr(vbKeyReturn) _
            & "Cancelar - Cancela la operación", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Confirmar Dirección")
    Else
        aResp = MsgBox("Que desea realizar con la dirección del cliente." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
            & "Si - Confirmar dirección." & Chr(vbKeyReturn) _
            & "No - Eliminar confirmación de dirección." & Chr(vbKeyReturn) _
            & "Cancelar - Cancela la operación", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Confirmar Dirección")
        End If
    
    If aResp = vbCancel Then Exit Sub
    
    On Error GoTo errConfirmar
    Screen.MousePointer = 11
    If aResp = vbYes Then
        Cons = "Update Direccion Set DirConfirmada = 1 Where DirCodigo = " & idDireccion
    Else
        Cons = "Update Direccion Set DirConfirmada = 0 Where DirCodigo = " & idDireccion
    End If
    cBase.Execute Cons
    If sEsProducto Then
        tDireccion.Text = clsGeneral.ArmoDireccionEnTexto(cBase, idDireccion, Departamento:=True, Localidad:=True, Zona:=True, ConfYVD:=True)
    Else
        tDirCliente.Text = clsGeneral.ArmoDireccionEnTexto(cBase, idDireccion, Departamento:=True, Localidad:=True, Zona:=True, ConfYVD:=True)
    End If
    Screen.MousePointer = 0
    
    Exit Sub
errConfirmar:
    clsGeneral.OcurrioError "Error al confirmar la dirección del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoProductoParaEditar(idProducto As Long)
        
    On Error GoTo ErrCE
    Screen.MousePointer = 11
    LimpioCamposProducto
    
    Cons = "Select * from Producto, Articulo " _
            & " Where ProCodigo = " & idProducto _
            & " And ProArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsAux.EOF Then
        
        tArticulo.Text = Trim(RsAux!ArtNombre)
        tArticulo.Tag = RsAux!ArtID
        lTipoProducto.Tag = RsAux!ArtID
        tDireccion.Tag = RsAux!ProCodigo
        
        If Not IsNull(RsAux!ProCompra) Then tFCompra.Text = Format(RsAux!ProCompra, "dd/mm/yyyy")
        If Not IsNull(RsAux!ProFacturaS) Then tfSerie.Text = UCase(Trim(RsAux!ProFacturaS))
        If Not IsNull(RsAux!ProFacturaN) Then tfNumero.Text = RsAux!ProFacturaN
        If Not IsNull(RsAux!ProNroSerie) Then tPSerie.Text = UCase(Trim(RsAux!ProNroSerie))
        If Not IsNull(RsAux!ProDireccion) Then CargoCamposDesdeBDDireccion RsAux!ProDireccion
        
        'Verifico si me modificaron los datos
        If RsAux!ProFModificacion <> CDate(vsProducto.Cell(flexcpData, vsProducto.Row, 1)) Then
            With vsProducto
                .Cell(flexcpData, .Row, 0) = lTipoProducto.Tag
                tFCompra.Tag = RsAux!ProFModificacion
                .Cell(flexcpData, .Row, 1) = tFCompra.Tag
                .Cell(flexcpText, .Row, 0) = Format(RsAux!ProCodigo, "#,000")
                .Cell(flexcpText, .Row, 1) = Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Row, 2) = EstadoProducto(CalculoEstadoProducto(RsAux!ProCodigo), True)
                .Cell(flexcpText, .Row, 3) = Trim(tFCompra.Text)
                .Cell(flexcpText, .Row, 4) = RetornoGarantia(RsAux!ArtID)
                .Cell(flexcpText, .Rows - 1, 5) = tPSerie.Text
                .Cell(flexcpText, .Rows - 1, 6) = ""
                If Not IsNull(RsAux!ProFacturaS) Then .Cell(flexcpText, .Row, 6) = Trim(RsAux!ProFacturaS) & " "
                If Not IsNull(RsAux!ProFacturaN) Then .Cell(flexcpText, .Row, 6) = .Cell(flexcpText, .Row, 6) & Trim(RsAux!ProFacturaN)
            End With
        End If
        
        'Me copio en los tag los valores que cargo.
        tFCompra.Tag = tFCompra.Text
        tPSerie.Tag = tPSerie.Text
        tfSerie.Tag = UCase(tfSerie.Text)
        tfNumero.Tag = tfNumero.Text
        
    End If
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub

ErrCE:
    clsGeneral.OcurrioError "Ocurrió un error al cargar la información del producto.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub CargoCamposDesdeBDDireccion(idDireccion As Long)
    If idDireccion <> 0 Then
        tDireccion.Text = clsGeneral.ArmoDireccionEnTexto(cBase, idDireccion, Departamento:=True, Localidad:=True, Zona:=True, EntreCalles:=True, Ampliacion:=True, ConfYVD:=True, ConEnter:=False)
    Else
        tDireccion.Text = ""
    End If
    tDireccion.Refresh
    bDireccionP.Tag = idDireccion
End Sub

Private Sub CargoDatosProducto(idCliente As Long, Optional idProducto As Long = 0)
On Error GoTo ErrCDP
Dim aValor As Integer, fModificado As String
Dim intVoy As Integer
    
    
    Screen.MousePointer = 11
    LimpioCamposProducto
    vsProducto.Rows = 1
    bolEOF = False
    Cons = "Select * From Producto, Articulo " _
        & " Where ProCliente = " & idCliente & " And ProArticulo = ArtID"
        
    If idProducto > 0 Then
        Cons = Cons & " And ProCodigo = " & idProducto
    Else
        If idCliente = paClienteEmpresa Or idCliente = paClienteAnglia Then Cons = Cons & " And ProFModificacion >= '" & Format(Date, "mm/dd/yyyy 00:00:00") & "'"
        Cons = Cons & " Order by ProFModificacion Desc, ArtNombre"
    End If
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If Val(vsProducto.Tag) > 0 Then
        For intVoy = 0 To Val(vsProducto.Tag) - 1
            If RsAux.EOF Then Exit For Else RsAux.MoveNext
        Next intVoy
    End If
    intVoy = 0
    
    Do While Not RsAux.EOF And intVoy < CantTuplas
        With vsProducto
            .AddItem ""
            
            aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor                                      'ID ARTICULO
            fModificado = RsAux!ProFModificacion: .Cell(flexcpData, .Rows - 1, 1) = fModificado  'F MODIFICACION
            
            aValor = 0
            If Not IsNull(RsAux!ProDireccion) Then
                aValor = BuscoZonaDireccion(RsAux!ProDireccion)
            Else
                If Val(tDirCliente.Tag) > 0 Then aValor = BuscoZonaDireccion(CLng(tDirCliente.Tag))
            End If
            .Cell(flexcpData, .Rows - 1, 2) = aValor    'Guardo la ZONA
            
            aValor = CalculoEstadoProducto(RsAux!ProCodigo): .Cell(flexcpData, .Rows - 1, 3) = aValor   'Estado del producto para cargar combo automático.
            .Cell(flexcpText, .Rows - 1, 2) = EstadoProducto(CInt(aValor), True)
            
            'Veo si tengo algún reporte abierto.
            .Cell(flexcpData, .Rows - 1, 4) = TieneReporteAbierto(RsAux!ProCodigo)
            
            aValor = RsAux!ArtTipo: .Cell(flexcpData, .Rows - 1, 5) = aValor
            
            If Not IsNull(RsAux!ProDocumento) Then .Cell(flexcpData, .Rows - 1, 6) = 1 Else .Cell(flexcpData, .Rows - 1, 6) = 0
            
            If Val(.Cell(flexcpData, .Rows - 1, 4)) > 0 Then .Cell(flexcpPicture, .Rows - 1, 0) = Image1.ListImages("servicio").ExtractIcon
            
            .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!ProCodigo, "#,000")
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ArtNombre)
            
            If Not IsNull(RsAux!ProCompra) Then .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!ProCompra, "dd/mm/yyyy")
            'Saco garantia-----------------------------------------------------------------------------------------------------------------------
            .Cell(flexcpText, .Rows - 1, 4) = RetornoGarantia(RsAux!ArtID)
            '--------------------------------------------------------------------------------------------------------------------------------------
            If Not IsNull(RsAux!ProNroSerie) Then .Cell(flexcpText, .Rows - 1, 5) = RsAux!ProNroSerie
            If Not IsNull(RsAux!ProFacturaS) Then .Cell(flexcpText, .Rows - 1, 6) = Trim(RsAux!ProFacturaS) & " "
            If Not IsNull(RsAux!ProFacturaN) Then .Cell(flexcpText, .Rows - 1, 6) = .Cell(flexcpText, .Rows - 1, 6) & Trim(RsAux!ProFacturaN)
            intVoy = intVoy + 1
            .Tag = Val(.Tag) + 1
        End With
        RsAux.MoveNext
    Loop
    If RsAux.EOF Then bolEOF = True
    'si llego al final o no tiene pongo en el tag el fina.
    If RsAux.EOF And vsProducto.Rows = 1 Then vsProducto.Tag = -2
    RsAux.Close
    
    If vsProducto.Rows > 1 And Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) = 0 Then HabilitoParaIngreso: MuestroCamposProducto Else DeshabilitoIngreso
    If vsProducto.Rows > 1 Then CargoHistoria vsProducto.Cell(flexcpText, 1, 0)
    Screen.MousePointer = 0
    Exit Sub
    
ErrCDP:
    clsGeneral.OcurrioError "Error al cargar los datos del producto.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub InicializoGrillaProducto()
    With vsProducto
        .Rows = 1
        .Cols = 1
        .ExtendLastCol = True
        .FormatString = "ID|Tipo de Artículo|Estado|>F.Compra|Garantía|N° Serie|Factura|"
        .ColWidth(0) = 650: .ColWidth(1) = 3000: .ColWidth(3) = 1000: .ColWidth(5) = 1100: .ColWidth(6) = 1100
    End With
End Sub
Private Sub LimpioObjetosComunes()
    tMotivo.Text = ""
    tAclaracion.Text = ""
    tComentarioInterno.Text = ""
    chEsReclamo.Value = 0
    chCoordinarEntrega.Value = 0
    tReclamo.Text = ""
    tUsuario.Text = ""
    'cTipoTelefono.ListIndex = -1
    tTelefonoServicio.Text = ""
    'tInterno.Text = ""
    vsMotivos.Rows = 1
    vsHistoria.Rows = 1
End Sub
Private Sub LimpioFichaTaller()
    LimpioObjetosComunes
    lTalFecha.Caption = Format(gFechaServidor, FormatoFP)
    cTalDeposito.Text = ""
    chkFueraGarantia.Value = 0
End Sub
Private Sub LimpioFichaRetiro()
    LimpioObjetosComunes
    cTipoFlete.Text = ""
    cRCamion.Text = ""
    tRFecha.Text = "": tRFecha.Tag = ""
    cRHora.Text = ""
    cRMoneda.Text = ""
    tRImporte.Text = ""
    cRFactura.Text = ""
    tRLiquidar.Text = ""
    cRDeposito.Text = ""
End Sub
Private Sub LimpioFichaVisita()
    LimpioObjetosComunes
    cVCamion.Text = ""
    tVFecha.Text = "": tVFecha.Tag = ""
    cVHora.Text = ""
    cVMoneda.Text = ""
    tVImporte.Text = ""
    cVFactura.Text = ""
    tVLiquidar.Text = ""
    cVComentario.Text = ""
End Sub
Private Sub FichaCliente(aTipoCliente As Integer)
On Error GoTo ErrFC
    
    If gCliente = 0 Then Exit Sub
    
    Screen.MousePointer = 11
    
    Dim objCliente As New clsCliente
    
    If aTipoCliente <> gTipoCliente Then gCliente = 0
    
    If aTipoCliente = TipoCliente.Cliente Then
        objCliente.Personas gCliente, 0, 0
    Else
        objCliente.Empresas gCliente, False
    End If
    Me.Refresh
    
    gCliente = objCliente.IDIngresado
    Set objCliente = Nothing
    If gCliente <> 0 Then
        BuscoClienteSeleccionado gCliente
    Else
        LimpioTodo
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrFC:
    clsGeneral.OcurrioError "Error al ir a ficha de cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub NuevoCliente(aTipoCliente As Integer)
On Error GoTo ErrFC
    Screen.MousePointer = 11
    Dim objCliente As New clsCliente
    If aTipoCliente <> gTipoCliente Then gCliente = 0
    
    If aTipoCliente = TipoCliente.Cliente Then
        objCliente.Personas gCliente, 0, 1
    Else
        objCliente.Empresas gCliente, True
    End If
    Me.Refresh
    gCliente = objCliente.IDIngresado
    Set objCliente = Nothing
    BuscoClienteSeleccionado gCliente
    Screen.MousePointer = 0
    Exit Sub
ErrFC:
    clsGeneral.OcurrioError "Error al ir a ficha de cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub LimpioTodo()
    sEsProducto = False
    gCliente = 0
    vsMotivos.Rows = 1
    LimpioFichaCliente
    LimpioCamposProducto
    OcultoCamposProducto
    DeshabilitoIngreso
    vsProducto.Rows = 1: vsProducto.Tag = "0"
End Sub

Private Sub tTelefonoServicio_GotFocus()
    With tTelefonoServicio
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tTelefonoServicio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Foco tUsuario
    End If
End Sub

Private Sub tTelefonoServicio_LostFocus()
    If Trim(tTelefonoServicio.Text) <> "" Then
        'If cTipoTelefono.ListIndex = -1 Then MsgBox "Debe seleccionar un tipo de teléfono.", vbExclamation, "ATENCIÓN": Foco cTipoTelefono: Exit Sub
        tTelefonoServicio.Tag = clsGeneral.RetornoFormatoTelefono(cBase, tTelefonoServicio.Text, Val(tDirCliente.Tag))
        If tTelefonoServicio.Tag <> "" Then
            tTelefonoServicio.Text = tTelefonoServicio.Tag
            If Not tTelefonoServicio.Text Like "09*" And chkNoDeseaSMS.Value = 0 Then
                MsgBox "El teléfono debe ser un celular para poder enviar SMS", vbInformation, "ATENCIÓN"
            End If
        Else
            MsgBox "El teléfono ingresado no coincide con los formatos establecidos.", vbExclamation, "ATENCIÓN"
            Foco tTelefonoServicio
        End If
    End If
End Sub

Private Sub tUsuario_GotFocus()
    With tUsuario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tUsuario.Text) Then
            tUsuario.Tag = 0
            tUsuario.Tag = BuscoUsuarioDigito(Val(tUsuario.Text), True)
            If Val(tUsuario.Tag) > 0 Then AccionGrabar
        Else
            MsgBox "Ingrese su dígito de usuario.", vbExclamation, "ATENCIÓN"
        End If
    End If
End Sub

Private Sub tVFecha_GotFocus()
    With tVFecha
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tVFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Not IsDate(tVFecha.Text) Then
                If Trim(tVFecha.Text) <> vbNullString Then
                    If IsDate(tVFecha.Tag) Then
                        CargoHoraEntregaParaDia cVHora, tVFecha.Tag
                        If cVHora.ListCount > 0 Then Foco cVHora
                    Else
                        MsgBox "La fecha ingresada no es correcta.", vbExclamation, "ATENCIÓN"
                    End If
                Else
                    tVFecha.Tag = tVFecha.Text
                    Foco cVHora
                End If
            Else
                If IsDate(tVFecha.Text) Then tVFecha.Tag = tVFecha.Text: tVFecha.Text = Format(tVFecha.Tag, "ddd d/mm/yy")
                CargoHoraEntregaParaDia cVHora, tVFecha.Tag
                If cVHora.ListCount > 0 Then Foco cVHora
            End If
        Case vbKeyUp
                If IsDate(tVFecha.Tag) Then tVFecha.Tag = Format(CDate(tVFecha.Tag) + 1, "d-mm-yy")
                If IsDate(tVFecha.Tag) Then tVFecha.Text = Format(tVFecha.Tag, "ddd d/mm/yy")
            Case vbKeyDown
                If IsDate(tVFecha.Tag) Then tVFecha.Tag = Format(CDate(tVFecha.Tag) - 1, "d-mm-yy")
                If IsDate(tVFecha.Tag) Then tVFecha.Text = Format(tVFecha.Tag, "ddd d/mm/yy")
        End Select
End Sub
Private Sub tVFecha_LostFocus()
    If IsDate(tVFecha.Tag) Then tVFecha.Text = Format(tVFecha.Tag, "ddd d/mm/yy") Else tVFecha.Text = ""
End Sub

Private Sub tVImporte_GotFocus()
    With tVImporte
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tVImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsNumeric(tVImporte.Text) Then Foco cVFactura
End Sub

Private Sub tVImporte_LostFocus()
    If IsNumeric(tVImporte.Text) Then tVImporte.Text = Format(tVImporte.Text, FormatoMonedaP) Else tVImporte.Text = ""
End Sub

Private Sub tVLiquidar_GotFocus()
    With tVLiquidar
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tVLiquidar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cVComentario
End Sub

Private Sub vsMotivos_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case vbKeyDelete: If vsMotivos.Row > 0 Then vsMotivos.RemoveItem vsMotivos.Row
        End Select
End Sub

Private Sub vsProducto_GotFocus()
    Status.Panels(1).Text = "Oprima [Av. Pag., Re. Pag.] avanza y retrocede; [+] accede al servicio pendiente."
End Sub

Private Sub vsProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsProducto.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            CargoProductoParaEditar Val(vsProducto.Cell(flexcpValue, vsProducto.Row, 0))
            If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) = 0 Then MuestroCamposProducto
            
            If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 6)) = 1 Then
                'Tiene Documento.
                tArticulo.Enabled = False: tArticulo.BackColor = Inactivo
                tFCompra.Enabled = False: tFCompra.BackColor = Inactivo
                tfSerie.Enabled = False: tfSerie.BackColor = Inactivo
                tfNumero.Enabled = False: tfNumero.BackColor = Inactivo
            End If
            If gCliente = paClienteEmpresa And UCase(prmInvocacion) = "R" Then
                If Val(tDirCliente.Tag) > 0 Then If Val(tDirCliente.Tag) <> Val(bDireccionP.Tag) Then tDireccion.BackColor = Obligatorio
            End If
            If vsHistoria.Rows = 1 Then TabRecepcion.Tabs(2).Selected = True Else TabRecepcion.Tabs(1).Selected = True
            TabRecepcion.SetFocus
        
        Case vbKeyAdd
            If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) > 0 Then
                EjecutarApp App.Path & "\Seguimiento de Servicios", vsProducto.Cell(flexcpData, vsProducto.Row, 4)
            Else
                EjecutarApp App.Path & "\Seguimiento de Servicios", ""
            End If
            
        Case vbKeySubtract
            If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) > 0 Then
                CumploServicio Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4))
            End If
            
        Case 93: If gCliente > 0 Then PopupMenu MnuProducto
        
        Case vbKeyPageUp
            If Val(vsProducto.Tag) = CantTuplas Then Exit Sub
            If Val(vsProducto.Tag) > 0 Then
                vsProducto.Tag = Val(vsProducto.Tag) - (CantTuplas * 2)
                If Val(vsProducto.Tag) < 0 Then vsProducto.Tag = "0"
            Else
                vsProducto.Tag = "0"
            End If
            CargoDatosProducto gCliente
        
        Case vbKeyPageDown
            If vsProducto.Tag = "-2" Then Exit Sub  'Esta en el final
            If bolEOF Then Exit Sub  'Esta en el final
            CargoDatosProducto gCliente
        
    End Select
End Sub

Private Sub vsProducto_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub vsProducto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And gCliente > 0 Then
        MnuProCumplirServicio.Enabled = (Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) > 0)
        PopupMenu MnuProducto
    End If
End Sub

Private Sub vsProducto_RowColChange()
On Error Resume Next
    OcultoCamposProducto
    If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) = 0 Then HabilitoParaIngreso Else DeshabilitoIngreso
    CargoHistoria vsProducto.Cell(flexcpValue, vsProducto.Row, 0)
    If Val(tDireccion.Tag) > 0 Then If Val(vsProducto.Cell(flexcpValue, vsProducto.Row, 0)) <> Val(tDireccion.Tag) Then LimpioCamposProducto
End Sub

Private Sub AjustoFichaTaller()
    On Error Resume Next
    TabRecepcion.Tabs("comodin").Image = Image1.ListImages("taller").Index
    PicTaller.ZOrder 0
    AjustoObjetosComunes
    
    If Not cTalDeposito.Enabled Then Exit Sub
    If cTalDeposito.ListCount = 0 Then CargoComboDeposito
    Cons = "Select * From Tipo Where TipCodigo = " & vsProducto.Cell(flexcpData, vsProducto.Row, 5)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!TipLocalRep) Then BuscoCodigoEnCombo cTalDeposito, RsAux!TipLocalRep
    End If
    RsAux.Close
    EsPosibleReclamo
End Sub
Private Sub AjustoObjetosComunes()
    tMotivo.ZOrder 0
    vsMotivos.ZOrder 0
End Sub
Private Sub AjustoFichaRetiro()
    On Error Resume Next
    
    If cRCamion.ListCount = 0 Then CargoCombosCamion
    If cTipoFlete.ListCount = 0 Then CargoComboTipoFlete
    If cRFactura.ListCount = 0 Then CargoComboFactura
    If cRMoneda.ListCount = 0 Then CargoComboMoneda
'    If cRHora.ListCount = 0 Then CargoComboHorario
    If cRDeposito.ListCount = 0 Then CargoComboDeposito
    TabRecepcion.Tabs("comodin").Image = Image1.ListImages("retiro").Index
    PicRetiro.ZOrder 0
    AjustoObjetosComunes
    EsPosibleReclamo
End Sub

Private Sub AjustoFichaVisita()
    On Error Resume Next
    
    If cVCamion.ListCount = 0 Then CargoCombosCamion
    If cVFactura.ListCount = 0 Then CargoComboFactura
    If cVMoneda.ListCount = 0 Then CargoComboMoneda
'    If cRHora.ListCount = 0 Then CargoComboHorario
    If cVComentario.ListCount = 0 Then CargoTextoVisita
    TabRecepcion.Tabs("comodin").Image = Image1.ListImages("visita").Index
    PicVisita.ZOrder 0
    AjustoObjetosComunes
    EsPosibleReclamo
End Sub

Private Sub AjustoFichaInicial()

    lVtMotivo.Left = ltMotivo.Left
    lVtMotivo.Top = ltMotivo.Top: lRtMotivo.Top = ltMotivo.Top
    
    tMotivo.Top = lVtMotivo.Top + PicTaller.Top - 20
    tMotivo.Left = ltMotivo.Left + ltMotivo.Width
    tMotivo.Width = 2600
    
    'Ajusto Grilla de Motivos.
    vsMotivos.Width = 3330
    vsMotivos.Left = TabRecepcion.Left + 150
    vsMotivos.Top = tMotivo.Top + tMotivo.Height + 40
    
End Sub
Private Sub InicializoGrillaMotivos()
    With vsMotivos
        .Rows = 1: .Cols = 1
        .FormatString = "Motivo"
    End With
End Sub
Private Sub InicializoGrillaHistoria()
    With vsHistoria
        .Rows = 1
        .WordWrap = True
        .FormatString = "Código|Fecha|Motivos|Estado|Llamado|>Importe"
        .ColWidth(1) = 750: .ColWidth(2) = 5200: .ColWidth(4) = 750 ': .ColWidth(5) = 1300
        .ColAlignment(1) = flexAlignLeftTop
        .ColAlignment(3) = flexAlignLeftTop
        .ColAlignment(4) = flexAlignLeftTop
        .ColAlignment(5) = flexAlignRightTop
    End With
End Sub

Private Sub OcultoCamposProducto()
    tArticulo.Enabled = False: tArticulo.BackColor = Inactivo
    tFCompra.Enabled = False: tFCompra.BackColor = Inactivo
    tfSerie.Enabled = False: tfSerie.BackColor = Inactivo
    tfNumero.Enabled = False: tfNumero.BackColor = Inactivo
    tPSerie.Enabled = False: tPSerie.BackColor = Inactivo
    tDireccion.Enabled = False: tDireccion.BackColor = Inactivo
    bDireccionP.Enabled = False
End Sub
Private Sub OcultoCamposMotivos()
    vsMotivos.Enabled = False
    tAclaracion.Enabled = False: tAclaracion.BackColor = Inactivo
    tComentarioInterno.Enabled = False: tComentarioInterno.BackColor = Inactivo
    chEsReclamo.Enabled = False
    chCoordinarEntrega.Enabled = False
    
    tReclamo.Enabled = False: tReclamo.BackColor = Inactivo
    tMotivo.Enabled = False: tMotivo.BackColor = Inactivo
    tUsuario.Enabled = False: tUsuario.BackColor = Inactivo
    'cTipoTelefono.Enabled = False: cTipoTelefono.BackColor = Inactivo
    tTelefonoServicio.Enabled = False: tTelefonoServicio.BackColor = Inactivo
    'tInterno.Enabled = False: tInterno.BackColor = Inactivo
End Sub
Private Sub OcultoCamposTaller()
    OcultoCamposMotivos
    cTalDeposito.Enabled = False: cTalDeposito.BackColor = Inactivo
End Sub
Private Sub OcultoCamposRetiro()
    OcultoCamposMotivos
    cTipoFlete.Enabled = False: cTipoFlete.BackColor = Inactivo
    cRCamion.Enabled = False: cRCamion.BackColor = Inactivo
    tRFecha.Enabled = False: tRFecha.BackColor = Inactivo
    cRHora.Enabled = False: cRHora.BackColor = Inactivo
    cRMoneda.Enabled = False: cRMoneda.BackColor = Inactivo
    tRImporte.Enabled = False: tRImporte.BackColor = Inactivo
    cRFactura.Enabled = False: cRFactura.BackColor = Inactivo
    tRLiquidar.Enabled = False: tRLiquidar.BackColor = Inactivo
    cRDeposito.Enabled = False: cRDeposito.BackColor = Inactivo
End Sub
Private Sub OcultoCamposVisita()
    OcultoCamposMotivos
    cVCamion.Enabled = False: cVCamion.BackColor = Inactivo
    tVFecha.Enabled = False: tVFecha.BackColor = Inactivo
    cVHora.Enabled = False: cVHora.BackColor = Inactivo
    cVMoneda.Enabled = False: cVMoneda.BackColor = Inactivo
    tVImporte.Enabled = False: tVImporte.BackColor = Inactivo
    cVFactura.Enabled = False: cVFactura.BackColor = Inactivo
    tVLiquidar.Enabled = False: tVLiquidar.BackColor = Inactivo
    cVComentario.Enabled = False: cVComentario.BackColor = Inactivo
End Sub
Private Sub MuestroCamposComunes()
    vsMotivos.Enabled = True
    tAclaracion.Enabled = True: tAclaracion.BackColor = Blanco
    tComentarioInterno.Enabled = True: tComentarioInterno.BackColor = Blanco
    chEsReclamo.Enabled = True
    chCoordinarEntrega.Enabled = True
    tMotivo.Enabled = True: tMotivo.BackColor = Blanco
    tUsuario.Enabled = True: tUsuario.BackColor = Obligatorio
    'cTipoTelefono.Enabled = True: cTipoTelefono.BackColor = Blanco
    tTelefonoServicio.Enabled = True: tTelefonoServicio.BackColor = Blanco
    'tInterno.Enabled = True: tInterno.BackColor = Blanco
End Sub
Private Sub MuestroCamposTaller()
    MuestroCamposComunes
    cTalDeposito.Enabled = True: cTalDeposito.BackColor = Obligatorio
End Sub
Private Sub MuestroCamposRetiro()
    MuestroCamposComunes
    cTipoFlete.Enabled = True: cTipoFlete.BackColor = Obligatorio
    cRCamion.Enabled = True: cRCamion.BackColor = Obligatorio
    tRFecha.Enabled = True: tRFecha.BackColor = Obligatorio
    cRHora.Enabled = True: cRHora.BackColor = Blanco
    cRMoneda.Enabled = True: cRMoneda.BackColor = Blanco
    tRImporte.Enabled = True: tRImporte.BackColor = Blanco
    cRFactura.Enabled = True: cRFactura.BackColor = Obligatorio
    tRLiquidar.Enabled = True: tRLiquidar.BackColor = vbWhite
    cRDeposito.Enabled = True: cRDeposito.BackColor = Obligatorio
End Sub
Private Sub MuestroCamposVisita()
    MuestroCamposComunes
    cVCamion.Enabled = True: cVCamion.BackColor = Obligatorio
    tVFecha.Enabled = True: tVFecha.BackColor = Obligatorio
    cVHora.Enabled = True: cVHora.BackColor = Blanco
    cVMoneda.Enabled = True: cVMoneda.BackColor = Blanco
    tVImporte.Enabled = True: tVImporte.BackColor = Blanco
    cVFactura.Enabled = True: cVFactura.BackColor = Obligatorio
    tVLiquidar.Enabled = True: tVLiquidar.BackColor = Blanco
    cVComentario.Enabled = True: cVComentario.BackColor = Blanco
End Sub
Private Sub MuestroCamposProducto()
    tArticulo.Enabled = True: tArticulo.BackColor = Blanco
    tFCompra.Enabled = True: tFCompra.BackColor = Blanco
    tfSerie.Enabled = True: tfSerie.BackColor = Blanco
    tfNumero.Enabled = True: tfNumero.BackColor = Blanco
    tPSerie.Enabled = True: tPSerie.BackColor = Blanco
    tDireccion.Enabled = True: tDireccion.BackColor = Blanco
    bDireccionP.Enabled = True
End Sub
Private Sub CargoCombosCamion()
On Error GoTo ErrCC
    Screen.MousePointer = 11
    cRCamion.Clear
    cVCamion.Clear
    Cons = "Select CamCodigo, CamNombre From Camion Order by CamNombre"
    CargoCombo Cons, cRCamion
    CargoCombo Cons, cVCamion
    Screen.MousePointer = 0
    Exit Sub
ErrCC:
    clsGeneral.OcurrioError "Error al cargar los camiones."
    Screen.MousePointer = 0
End Sub

Private Sub CargoComboTipoFlete()
On Error GoTo ErrCC
    Screen.MousePointer = 11
    cTipoFlete.Clear
    Cons = "Select TFlCodigo, TFlNombreCorto From TipoFlete Order by TFlNombreCorto"
    CargoCombo Cons, cTipoFlete
    Screen.MousePointer = 0
    Exit Sub
ErrCC:
    clsGeneral.OcurrioError "Error al cargar los tipos de fletes."
    Screen.MousePointer = 0
End Sub

Private Sub CargoComboDeposito()
On Error GoTo ErrCC
    Screen.MousePointer = 11
    cTalDeposito.Clear
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Where SucHabilitada = 1 Order by SucAbreviacion"
    CargoCombo Cons, cTalDeposito
    CargoCombo Cons, cRDeposito
    Screen.MousePointer = 0
    Exit Sub
ErrCC:
    clsGeneral.OcurrioError "Error al cargar los Depósitos."
    Screen.MousePointer = 0
End Sub

Private Sub CargoComboMoneda()
On Error GoTo ErrCC
    Screen.MousePointer = 11
    cRMoneda.Clear
    cVMoneda.Clear
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cRMoneda
    CargoCombo Cons, cVMoneda
    Screen.MousePointer = 0
    Exit Sub
ErrCC:
    clsGeneral.OcurrioError "Error al cargar las monedas."
    Screen.MousePointer = 0
End Sub
Private Sub CargoComboFactura()
On Error GoTo ErrCC
    
    Screen.MousePointer = 11
    cRFactura.Clear
    cVFactura.Clear
    
    cRFactura.AddItem TipoFacturaServicio(FacturaServicio.Camion): cRFactura.ItemData(cRFactura.NewIndex) = FacturaServicio.Camion
    cRFactura.AddItem TipoFacturaServicio(FacturaServicio.CGSA): cRFactura.ItemData(cRFactura.NewIndex) = FacturaServicio.CGSA
    
    cVFactura.AddItem TipoFacturaServicio(FacturaServicio.Camion): cVFactura.ItemData(cVFactura.NewIndex) = FacturaServicio.Camion
    cVFactura.AddItem TipoFacturaServicio(FacturaServicio.CGSA): cVFactura.ItemData(cVFactura.NewIndex) = FacturaServicio.CGSA
    
    Screen.MousePointer = 0
    Exit Sub
ErrCC:
    clsGeneral.OcurrioError "Error al cargar los tipos de facturación."
    Screen.MousePointer = 0
End Sub

Private Sub HabilitoParaIngreso()
    MuestroCamposTaller
    MuestroCamposRetiro
    MuestroCamposVisita
End Sub

Private Sub DeshabilitoIngreso()
    OcultoCamposTaller
    LimpioFichaTaller
    OcultoCamposRetiro
    LimpioFichaRetiro
    OcultoCamposVisita
    LimpioFichaVisita
End Sub

Private Sub AccionGrabar()

    If Not ValidoIngreso Then Exit Sub
    Dim strMsg As String
    strMsg = "Producto ID = " & vsProducto.Cell(flexcpText, vsProducto.Row, 0)
    strMsg = strMsg & Chr(13) & "Tipo = " & vsProducto.Cell(flexcpText, vsProducto.Row, 1)
    
    
    '1 Taller, 2  Retiro, 3  Visita
    Select Case cDatoIngreso.ItemData(cDatoIngreso.ListIndex)
        Case 1
            If MsgBox("¿Confirma grabar el ingreso de taller?" & Chr(13) & strMsg, vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then AccionGrabarTaller
        Case 2
            If MsgBox("¿Confirma grabar la solicitud de retiro?" & Chr(13) & strMsg, vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then AccionGrabarRetiro
        Case 3
            
            Dim bCambio As Boolean
            bCambio = ElComentarioEsCambio
            If bCambio Then
                strMsg = strMsg & vbCrLf & vbCrLf & _
                        "Luego de grabar se activará la aplicación para ingresar el Cambio de Artículo en Garantía"
            End If
            
            If MsgBox("¿Confirma grabar la solicitud de visita?" & Chr(13) & strMsg, vbQuestion + vbYesNo, "Grabar Visita") = vbYes Then
                AccionGrabarVisita bCambio
            End If
    End Select
End Sub
Private Function ElComentarioEsCambio() As Boolean

'Valido Si el comentario es cambio de artículo 11/9/2002
'Valor1 AS TViEsCambioArticulo
On Error GoTo errBusco
    ElComentarioEsCambio = False
    
    If cVComentario.ListIndex <> -1 Then
        Cons = "Select * from TextoVisita " & _
            " Where TViCodigo = " & cVComentario.ItemData(cVComentario.ListIndex)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If Not IsNull(RsAux!TViEsCambioArticulo) Then
                If RsAux!TViEsCambioArticulo = 1 Then ElComentarioEsCambio = True
            End If
        End If
        RsAux.Close
    End If

errBusco:
End Function

Private Sub AccionGrabarTaller()
On Error GoTo ErrBT
Dim idServicio As Long
    idServicio = 0
    Screen.MousePointer = 11
    FechaDelServidor
    
    Dim iUsu As Integer, idAut As Long
    Dim sDef As String
    If MsgClienteNoVender(gCliente, False) Then
        Dim objSuceso As New clsSuceso
        objSuceso.TipoSuceso = 23 'TipoSuceso.ClienteNoVender
        objSuceso.ActivoFormulario CLng(tUsuario.Tag), "No vender a cliente", cBase
        Me.Refresh
        iUsu = objSuceso.Usuario
        sDef = objSuceso.Defensa
        idAut = objSuceso.Autoriza
        Set objSuceso = Nothing
        If iUsu = 0 Then Screen.MousePointer = 0: Exit Sub
    End If
    
    
    cBase.BeginTrans
    
    If iUsu > 0 Then
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, 23, paCodigoDeTerminal, iUsu, 0, , "Cliente No vender", Trim(sDef), 1, gCliente, idAut
    End If
    
    On Error GoTo ErrResumir
    Cons = "Select * From Servicio Where SerProducto = " & vsProducto.Cell(flexcpValue, vsProducto.Row, 0) _
        & " And SerEstadoServicio Not IN (" & EstadoS.Anulado & ", " & EstadoS.Cumplido & ")"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        'If Trim(tTelefonoServicio.Text) <> "" Then GraboTelefono
        
        Dim iEstProd As Integer
        iEstProd = vsProducto.Cell(flexcpData, vsProducto.Row, 3)
        If (chkFueraGarantia.Value = 1) Then iEstProd = EstadoP.FueraGarantia
        
        idServicio = InsertoServicio(vsProducto.Cell(flexcpValue, vsProducto.Row, 0), iEstProd, EstadoS.Taller, cTalDeposito.ItemData(cTalDeposito.ListIndex), tTelefonoServicio.Text, IIf(chkNoDeseaSMS.Value = 1, 0, 1), Trim(tAclaracion.Text), -1, tUsuario.Tag, Trim(tComentarioInterno.Text))
        If vsMotivos.Rows > 1 Then InsertoMotivos idServicio
        
        'Si ingresa directo al local inserto la tabla taller.
        If cTalDeposito.ItemData(cTalDeposito.ListIndex) = paCodigoDeSucursal Then InsertoServicioTaller idServicio, tUsuario.Tag
        If gCliente = paClienteEmpresa Then
            HagoCambioDeEstado vsProducto.Cell(flexcpData, vsProducto.Row, 0), paEstadoARecuperar, idServicio
        ElseIf cTalDeposito.ItemData(cTalDeposito.ListIndex) = 41 Then
            Cons = "UPDATE Taller SET TalUbicacionSucursal = 1161 WHERE TalServicio = " & idServicio
            cBase.Execute Cons
        End If
        cBase.CommitTrans
        'Imprimo fichas.
        'ANULE
        ImprimoFichaTaller idServicio, iEstProd
        ImprimirFichaServicio idServicio, IIf(chkNoDeseaSMS.Value = 0, tTelefonoServicio.Text, "")
        'ImprimirFichaServicio idServicio
        Foco tCi
    Else
        idServicio = RsAux!SerCodigo
        RsAux.Close
        cBase.RollbackTrans
        MsgBox "El producto ya tiene un servicio abierto, el código de servicio es el " & idServicio, vbInformation, "ATENCIÓN"
    End If
    LimpioTodo
    Screen.MousePointer = 0
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Error al iniciar la transacción.", Trim(Err.Description)
    Screen.MousePointer = 0
    Exit Sub
ErrResumir:
    Resume ErrRB
ErrRB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al intentar almacenar la información de taller.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
Private Sub ImprimoFichaTaller(idServicio As Long, ByVal iEstProd As Integer)
Dim aTexto As String
Dim iPY As Single


    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    'Seteo por defecto la impresora

    SeteoImpresoraPorDefecto paPrintConfD

    With vsFicha

        On Error Resume Next
        .Orientation = orLandscape
        .Device = paPrintConfD
        .PaperBin = paPrintConfB
        .PaperSize = paPrintConfPaperSize

        On Error GoTo errImprimir

        .AbortWindow = False
        .FileName = "Ficha de Ingreso a Taller"

        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If

        .FontSize = 10
        .TableBorder = tbNone
        .TextAlign = taLeftTop
        .FontBold = True
'        .Paragraph = ""
        .CurrentX = 7450
'        .Paragraph = "Servicio: " & idServicio & IIf(paClienteEmpresa = gCliente, Space(4) & "STOCK", "")
        .FontBold = False

        .FontSize = 8
        iPY = .CurrentY
        .CurrentX = 8000

        .Font = "3 of 9 Barcode"
        .FontSize = 20
        .Paragraph = "*S" & idServicio & "*"

        .Font = "Tahoma"
        .FontSize = 10
        .CurrentY = iPY

        .Paragraph = ""
        .CurrentX = 6000
        .Paragraph = "Fecha:" & Format(gFechaServidor, "d-Mmm yyyy hh:mm") & Space(5) & "Recibido por: " & tUsuario.Text
        .Paragraph = ""
        .CurrentX = 6000
        .Paragraph = "Local: " & Trim(cTalDeposito.Text)

        .CurrentY = iPY
        .FontBold = False
        .FontSize = 58
        .Paragraph = Space(1) & idServicio
        .FontSize = 9.25

'ARTICULO
        .FontBold = True
        aTexto = "(" & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 0)) & ") " & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 1))
        .Paragraph = aTexto
        .FontBold = False

        .AddTable "<900|<1500|<900|<1100|<900|<1800|<900|<500", "Factura:|" & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 6)) & "|Compra:|" & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 3)) & "|# Serie:|" & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 5)) & "|Estado:|" & EstadoProducto(iEstProd, True), ""

        aTexto = ""
        For I = 1 To vsMotivos.Rows - 1
            If aTexto = "" Then aTexto = Trim(vsMotivos.Cell(flexcpText, I, 0)) Else aTexto = aTexto & ", " & Trim(vsMotivos.Cell(flexcpText, I, 0))
        Next I
        .AddTable "<900|<10100", "Motivos:|" & aTexto, ""
        If Trim(tAclaracion.Text) <> "" Then .AddTable "<1000|<10000", "Aclaración:|" & Trim(tAclaracion.Text), ""

        If IsNumeric(tReclamo.Text) Then .Paragraph = "Reclamo Servicio: " & tReclamo.Text Else .Paragraph = ""
        .Paragraph = CargoStringUltServicio(idServicio, vsProducto.Cell(flexcpText, vsProducto.Row, 0))
        
        
        
        If gTipoCliente = TipoCliente.Cliente Then
            If Trim(tCi.Text) <> "" Then
                aTexto = "Cliente:|(" & clsGeneral.RetornoFormatoCedula(tCi.Text) & ")"
            Else
                aTexto = "Cliente:|"
            End If
        Else
            If Trim(tRuc.Text) <> "" Then
                aTexto = "Cliente:|(" & clsGeneral.RetornoFormatoRuc(tRuc.Text) & ")"
            Else
                aTexto = "Cliente:|"
            End If
        End If
        aTexto = aTexto & lTitular.Caption

        .Paragraph = aTexto
        .Paragraph = "Teléfono: " & Trim(tTelCliente.Text)
        .Paragraph = "Dirección: " & Trim(tDirCliente.Text)
        .Paragraph = ""
        .Paragraph = "Nota: 1) - Para retirar el aparato es indispensable presentar esta boleta. -"
        .Paragraph = "      2) - El plazo de retiro del aparato es de 30 días contados a partir de la fecha de esta boleta. Expirado dicho plazo se perderá todo derecho a reclamo sobre el mismo. -"
        .Paragraph = "Carlos Gutiérrez no será responsable por el contenido almacenado en los dispositivos de terceros.  Todas las medidas para la preservación de la información, por ejemplo respaldar y eliminar todo aquello confidencial será responsabilidad del cliente ."
        .Paragraph = ""
        .Paragraph = " Reparado:"
        .Paragraph = " Vía Archivo"
        

        .EndDoc

        .Device = paPrintConfD
        .PaperBin = paPrintConfB
        .PaperSize = paPrintConfPaperSize

        .PrintDoc   'Archivo
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
    Exit Sub

errImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub
Private Function CargoStringUltServicio(ByVal idServ As Long, ByVal idProd As Long) As String
On Error GoTo errCS
Dim rsUS As rdoResultset
Dim sText As String, sFCump As String
    
    CargoStringUltServicio = ""
    Cons = "Select Count(*), Max(SerFCumplido) From Servicio Where SerProducto = " & idProd & " And SerCodigo <> " & idServ
    Set rsUS = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsUS(0) = 0 Then
        rsUS.Close: Exit Function
    Else
        sText = "Ver ficha técnica."
'        sFCump = Format(rsUS(1), "dd/mm/yyyy")
'        sText = "Este artículo vino " & rsUS(0) & " veces."
    End If
    rsUS.Close
    
    CargoStringUltServicio = sText
    Exit Function
    
    'Veo si para la fecha del servicio hay alguna entrega
    Cons = "Select * From ServicioVisita, Servicio Where VisTipo = 3 And VisSinEfecto = 0 " _
        & " And VisFImpresion Is Not Null And VisServicio = SerCodigo " _
        & " And SerProducto = " & idProd & " And SerFCumplido = '" & Format(sFCump, "mm/dd/yyyy hh:nn:ss") & "'"
    Set rsUS = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsUS.EOF Then
        sText = sText & " Última reparación fue entregada el " & Format(rsUS!VisFecha, "dd/mm/yyyy")
    Else
        sText = sText & " Última reparación fue entregada el " & Format(sFCump, "dd/mm/yyyy")
    End If
    rsUS.Close
    CargoStringUltServicio = sText
    Exit Function
errCS:
End Function
'Private Sub GraboTelefono()
'    Cons = "Select * From Telefono Where TelCliente = " & gCliente _
'        & " And TelTipo = " & cTipoTelefono.ItemData(cTipoTelefono.ListIndex)
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'    If RsAux.EOF Then
'        RsAux.AddNew
'        RsAux!TelCliente = gCliente
'        RsAux!TelTipo = cTipoTelefono.ItemData(cTipoTelefono.ListIndex)
'        RsAux!TelNumero = tTelefono.Text
'        If Trim(tInterno.Text) <> "" Then RsAux!TelInterno = tInterno.Text
'        RsAux.Update
'    Else
'        RsAux.Edit
'        RsAux!TelNumero = tTelefono.Text
'        If Trim(tInterno.Text) <> "" Then RsAux!TelInterno = tInterno.Text Else RsAux!TelInterno = Null
'        RsAux.Update
'    End If
'    RsAux.Close
'    tTelCliente.Text = TelefonoATexto(gCliente)     'Telefonos
'End Sub
Private Function ValidoIngreso() As Boolean
    
    ValidoIngreso = False
    
    If Val(tUsuario.Tag) = 0 Then MsgBox "Ingrese su digito de usuario.", vbExclamation, "ATENCIÓN": Foco tUsuario: Exit Function
    If Not clsGeneral.TextoValido(tAclaracion.Text) Then MsgBox "Se ingresó alguna comilla simple en la aclaración, debe eliminarla.", vbExclamation, "ATENCIÓN": Foco tAclaracion: Exit Function
    If Not clsGeneral.TextoValido(tComentarioInterno.Text) Then MsgBox "Se ingresó alguna comilla simple en el comentario, debe eliminarla.", vbExclamation, "ATENCIÓN": Foco tComentarioInterno: Exit Function
    If Trim(tTelefonoServicio.Text) <> "" Then
        'If cTipoTelefono.ListIndex = -1 Then MsgBox "Debe seleccionar un tipo de teléfono.", vbExclamation, "ATENCIÓN": Foco cTipoTelefono: Exit Function
        tTelefonoServicio.Tag = clsGeneral.RetornoFormatoTelefono(cBase, tTelefonoServicio.Text, 0)
        'If tTelefono.Tag = "" Then MsgBox "El formato del teléfono no es válido.", vbExclamation, "ATENCIÓN": Foco tTelefono: Exit Function
        If gCliente <> paClienteEmpresa Then
            If Not (tTelefonoServicio.Text Like "09*") And chkNoDeseaSMS.Value = 0 Then
                MsgBox "El teléfono debe ser un celular para poder envíar SMS.", vbExclamation, "ATENCIÓN"
                tTelefonoServicio.SetFocus
                Exit Function
            End If
        End If
        
    ElseIf gCliente <> paClienteEmpresa Then
        MsgBox "Debe ingresar un número de teléfono.", vbExclamation, "Validación"
        tTelefonoServicio.SetFocus
        Exit Function
    End If
    
    If tReclamo.Text <> "" Then
        If IsNumeric(tReclamo.Text) Then
            If Not ReclamoValido(tReclamo.Text) Then
                MsgBox "No es un servicio válido de reclamo el que ingreso.", vbExclamation, "ATENCIÓN"
                tReclamo.SetFocus: Exit Function
            End If
        Else
            tReclamo.SetFocus
            MsgBox "No es un código de servicio.", vbExclamation, "ATENCIÓN"
            Exit Function
        End If
    End If
    
    '1 Taller, 2  Retiro, 3  Visita
    Select Case cDatoIngreso.ItemData(cDatoIngreso.ListIndex)
        Case 1: If cTalDeposito.ListIndex = -1 Then MsgBox "Debe seleccionar un local de reparación para el producto.", vbExclamation, "ATENCIÓN": Foco cTalDeposito: Exit Function
        Case 2
            If cTipoFlete.ListIndex = -1 Then MsgBox "Debe seleccionar un tipo de flete.", vbExclamation, "ATENCIÓN": Foco cTipoFlete: Exit Function
            If cRCamion.ListIndex = -1 Then MsgBox "Debe seleccionar un camión para el retiro.", vbExclamation, "ATENCIÓN": Foco cRCamion: Exit Function
            If Not IsDate(tRFecha.Tag) Then MsgBox "Debe ingresar una fecha de retiro.", vbExclamation, "ATENCIÓN": Foco tRFecha: Exit Function
            If cRMoneda.ListIndex = -1 And Trim(cRMoneda.Text) <> "" Then MsgBox "La moneda no es correcta.", vbExclamation, "ATENCIÓN": Foco cRMoneda: Exit Function
            If Not IsNumeric(tRImporte.Text) And Trim(tRImporte.Text) <> "" Then MsgBox "El importe ingresado no es válido.", vbExclamation, "ATENCIÓN": Foco tRImporte: Exit Function
            If IsNumeric(tRImporte.Text) And cRMoneda.ListIndex = -1 Then MsgBox "Debe ingresar una moneda para el importe.", vbExclamation, "ATENCIÓN": Foco cRMoneda: Exit Function
            If cRFactura.ListIndex = -1 Then MsgBox "Debe seleccionar un tipo de facturación.", vbExclamation, "ATENCIÓN": Foco cRFactura: Exit Function
            If Not clsGeneral.TextoValido(cRHora.Text) Then MsgBox "Se ingresó alguna comilla simple en el horario de retiro, debe eliminarla.", vbExclamation, "ATENCIÓN": Foco cRHora: Exit Function
            If cRDeposito.ListIndex = -1 Then MsgBox "No se ingresó un local de reparación correcto.", vbExclamation, "ATENCIÓN": Foco cRDeposito: Exit Function
            If tRLiquidar.Text = "" Then tRLiquidar.Text = "0"
            If Not IsNumeric(tRLiquidar.Text) Then MsgBox "El formato no es numérico para la liquidación del camionero.", vbExclamation, "ATENCIÓN": Foco tRLiquidar: Exit Function
            If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 2)) = 0 Then MsgBox "La dirección del producto seleccionado no tiene asignada una zona.", vbInformation, "ATENCIÓN": Exit Function
        Case 3
            If cVCamion.ListIndex = -1 Then MsgBox "Debe seleccionar un camión para la visita.", vbExclamation, "ATENCIÓN": Foco cVCamion: Exit Function
            If Not IsDate(tVFecha.Tag) Then MsgBox "Debe ingresar una fecha de visita.", vbExclamation, "ATENCIÓN": Foco tVFecha: Exit Function
            If cVMoneda.ListIndex = -1 And Trim(cVMoneda.Text) <> "" Then MsgBox "La moneda no es correcta.", vbExclamation, "ATENCIÓN": Foco cVMoneda: Exit Function
            If Not IsNumeric(tVImporte.Text) And Trim(tVImporte.Text) <> "" Then MsgBox "El importe ingresado no es válido.", vbExclamation, "ATENCIÓN": Foco tVImporte: Exit Function
            If IsNumeric(tVImporte.Text) And cVMoneda.ListIndex = -1 Then MsgBox "Debe ingresar una moneda para el importe.", vbExclamation, "ATENCIÓN": Foco cVMoneda: Exit Function
            If cVFactura.ListIndex = -1 Then MsgBox "Debe seleccionar un tipo de facturación.", vbExclamation, "ATENCIÓN": Foco cVFactura: Exit Function
            If Not clsGeneral.TextoValido(cVHora.Text) Then MsgBox "Se ingresó alguna comilla simple en el horario de visita, debe eliminarla.", vbExclamation, "ATENCIÓN": Foco cVHora: Exit Function
            If tVLiquidar.Text = "" Then tVLiquidar.Text = "0"
            If Not IsNumeric(tVLiquidar.Text) Then MsgBox "El formato no es numérico para la liquidación del camionero.", vbExclamation, "ATENCIÓN": Foco tVLiquidar: Exit Function
            If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 2)) = 0 Then MsgBox "La dirección del producto seleccionado no tiene asignada una zona.", vbInformation, "ATENCIÓN": Exit Function
    End Select
    ValidoIngreso = True
End Function

Private Sub InsertoMotivoEnGrilla(ByVal lIDMot As Long, ByVal sName As String)
    'Tengo el RSAUX con el motivo
    Dim aValor As Long
    'Verifico que no este insertado.
    With vsMotivos
        For I = 1 To .Rows - 1
            If Val(.Cell(flexcpData, I, 0)) = lIDMot Then MsgBox "El motivo ya fue ingresado, verifique.", vbInformation, "ATENCIÓN": Exit Sub
        Next I
        .AddItem Trim(sName)
        .Cell(flexcpData, .Rows - 1, 0) = lIDMot
    End With
End Sub

Private Function InsertoServicio(idProducto As Long, EstadoProducto As Integer, EstadoServicio As Integer, _
LocalReparacion As Long, ByVal telefono As String, ByVal SendSMS As Boolean, _
Optional Comentario As String = "", Optional LocalRecepcion As Long = -1, Optional Usuario As Long = -1, _
Optional ComentarioInterno As String = "") As Long
    
    If LocalRecepcion = -1 Then LocalRecepcion = paCodigoDeSucursal
    If Usuario = -1 Then Usuario = paCodigoDeUsuario
    
    '---------------------------------------------
    'Inserto
    Cons = "INSERT INTO Servicio (SerProducto, SerCliente, SerFecha, SerEstadoProducto, SerLocalIngreso, " _
        & " SerLocalReparacion, SerEstadoServicio, SerUsuario, SerModificacion, SerReclamoDe, SerComentario, SerComInterno, SerTelefono, SerSendSMS, SerCoordinarEntrega) Values (" _
        & idProducto & ", " & gCliente & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', " & EstadoProducto & ", " & LocalRecepcion
    
    If LocalReparacion = 0 Then Cons = Cons & ", Null " Else Cons = Cons & ", " & LocalReparacion
    
    Cons = Cons & ", " & EstadoServicio & ", " & Usuario & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', "
    If IsNumeric(tReclamo.Text) Then Cons = Cons & tReclamo.Text & ", " Else Cons = Cons & " Null, "
    If Comentario = "" Then Cons = Cons & "Null, " Else Cons = Cons & "'" & Comentario & "', "
    If ComentarioInterno = "" Then Cons = Cons & "Null, " Else Cons = Cons & "'" & ComentarioInterno & "',"
    
    Cons = Cons & "'" & telefono & "', " & IIf(SendSMS, 1, 0)
    
    Cons = Cons & ", " & IIf(chCoordinarEntrega.Value = 0, " Null ", "1")
    
    Cons = Cons & ")"
    
    cBase.Execute (Cons)
    
    '---------------------------------------------
    'Saco el mayor código de servicio.
    Cons = "Select Max(SerCodigo) From Servicio"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    InsertoServicio = RsAux(0)
    RsAux.Close
    '---------------------------------------------
    
    If (SendSMS) Then
        telefono = "598" + Trim(Mid(telefono, 2))
        Cons = "INSERT INTO MensajeWhatsApp (MWACliente,MWATipo,MWADocumento,MWATelefono,MWAVencimiento,MWAEstado,MWADesde) Values (" & _
            gCliente & ", 2, " & InsertoServicio & ", '" + Replace(Trim(telefono), " ", "") + "', '" + _
            Format(CDate(Now) + 60, "yyyyMMdd hh:mm:ss") & "', 2, '" + Format(CDate(Now), "yyyyMMdd hh:mm:ss") & "')"
        cBase.Execute Cons
    End If
    
End Function

Private Sub InsertoMotivos(idServicio As Long)
    With vsMotivos
        For I = 1 To .Rows - 1
            Cons = "Insert Into ServicioRenglon (SReServicio, SReTipoRenglon,  " _
                & " SReMotivo, SReCantidad) Values (" & idServicio & ", " & TipoRenglonS.Llamado & ",  " & Val(.Cell(flexcpData, I, 0)) & ", 1)"
            cBase.Execute (Cons)
        Next I
    End With
End Sub

Private Sub InsertoServicioTaller(idServicio As Long, Optional Usuario As Integer = -1)

    If Usuario = -1 Then Usuario = paCodigoDeUsuario
    If cTalDeposito.ItemData(cTalDeposito.ListIndex) <> paCodigoDeSucursal Then
        'Inserto también el local para el traslado.
        Cons = "Insert Into Taller(TalServicio, TalFIngresoRealizado, TalFIngresoRecepcion, TalModificacion, TalUsuario, TalLocalAlCliente) Values (" _
            & idServicio & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
            & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', " & Usuario & ", " & cTalDeposito.ItemData(cTalDeposito.ListIndex) & ")"
    Else
        Cons = "Insert Into Taller(TalServicio, TalFIngresoRealizado, TalFIngresoRecepcion, TalModificacion, TalUsuario) Values (" _
            & idServicio & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
            & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', " & Usuario & ")"
    End If
    cBase.Execute (Cons)
    
End Sub

Private Sub ArmoXDefectoRetiro()
    If cTipoFlete.ListIndex = -1 And paTipoFleteVentaTelefonica > 0 And cTipoFlete.Enabled Then
        'Cargo primero la moneda porque sino tengo la misma no me busca el valor del flete.
        If paMonedaPesos > 0 Then BuscoCodigoEnCombo cRMoneda, paMonedaPesos
        BuscoDatosTipoDeFlete paTipoFleteVentaTelefonica
        BuscoCodigoEnCombo cTipoFlete, paTipoFleteVentaTelefonica
        Cons = "Select * From Tipo Where TipCodigo = " & vsProducto.Cell(flexcpData, vsProducto.Row, 5)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
        If Not RsAux.EOF Then If Not IsNull(RsAux!TipLocalRep) Then BuscoCodigoEnCombo cRDeposito, RsAux!TipLocalRep
        RsAux.Close
    End If
End Sub

Private Sub BuscoDatosTipoDeFlete(IDTFlete As Long)
    
    If IDTFlete = 0 Then
        If cTipoFlete.ListIndex = -1 Then Exit Sub
        IDTFlete = cTipoFlete.ItemData(cTipoFlete.ListIndex)
    End If
    
    
    'Por defecto pongo que factura camión
    BuscoCodigoEnCombo cRFactura, FacturaServicio.Camion: tRLiquidar.Text = "0.00"
    
    
    douAgenda = 0
    douHabilitado = 0
    strCierre = ""
    
    On Error GoTo ErrBDTF
    Screen.MousePointer = 11
    'Cons = "Select * From TipoFlete Where TFlCodigo = " & IDTFlete
    
    Cons = "Select * From TipoFlete LEFT OUTER JOIN FleteAgendaZona " & _
            " ON FAZZona = " & Val(vsProducto.Cell(flexcpData, vsProducto.Row, 2)) & " And FAZTipoFlete = TFLCodigo Where TFlCodigo = " & IDTFlete
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!TFlFormaPago) Then
            'Factura Camión = 3
            If RsAux!TFlFormaPago = 3 Then
                BuscoCodigoEnCombo cRFactura, FacturaServicio.Camion
            Else
                BuscoCodigoEnCombo cRFactura, RsAux!TFlFormaPago
            End If
        End If
        
        If IsNull(RsAux!TFlAgenda) And IsNull(RsAux("FAZAgenda")) Then
            tRFecha.Text = "": tRFecha.Tag = "": cRHora.ListIndex = -1
        Else
            tRFecha.Tag = BuscoPrimerDiaAEnviar
            If tRFecha.Tag = "" Then
                tRFecha.Text = ""
                cRHora.Clear
            Else
                tRFecha.Text = Format(CDate(tRFecha.Tag), "ddd d/mm/yy")
                CargoHoraEntregaParaDia cRHora, tRFecha.Tag
            End If
        End If
    End If
    RsAux.Close
    
    If paCamionRetiroVisita > 0 Then
        BuscoCodigoEnCombo cRCamion, paCamionRetiroVisita
    Else
        Cons = "Select * From CamionFlete, CamionZona " _
            & " Where CTFTipoFlete = " & IDTFlete _
            & " And CZoZona = " & Val(vsProducto.Cell(flexcpData, vsProducto.Row, 2)) _
            & " And CTFCamion = CZoCamion Order by CZoPrioridad "
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then BuscoCodigoEnCombo cRCamion, RsAux!CZoCamion
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBDTF:
    clsGeneral.OcurrioError "Error al cargar los datos para el tipo de flete.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoValorFlete()
    
    If cRMoneda.ListIndex = -1 Or cTipoFlete.ListIndex = -1 Then Exit Sub
    
    tRImporte.Text = 0
    
    If paCobroEnEntrega Then Exit Sub
    
    Dim cVF As Currency, cLi As Currency
    loc_DefinoPrecioFlete cTipoFlete.ItemData(cTipoFlete.ListIndex), vsProducto.Cell(flexcpData, vsProducto.Row, 2), cVF, cLi
    
    'Le sumo el coeficiente.
    tRImporte.Text = Format(CCur(cVF) * paCoefFleteRetiro, "#,##0.00")
    

End Sub

Private Sub AccionGrabarRetiro()
On Error GoTo ErrBR
Dim idServicio As Long, strHora As String
    
    idServicio = 0
    Screen.MousePointer = 11
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo ErrResumir
    idServicio = TieneReporteAbierto(vsProducto.Cell(flexcpValue, vsProducto.Row, 0))
    If idServicio = 0 Then
        'If Trim(tTelefono.Text) <> "" Then GraboTelefono
        'NO paso la aclaración x la impresión.
        idServicio = InsertoServicio(vsProducto.Cell(flexcpValue, vsProducto.Row, 0), CInt(vsProducto.Cell(flexcpData, vsProducto.Row, 3)), EstadoS.Retiro, cRDeposito.ItemData(cRDeposito.ListIndex), tTelefonoServicio.Text, IIf(chkNoDeseaSMS.Value = 0, 1, 0), "", -1, tUsuario.Tag, tComentarioInterno.Text)
        If vsMotivos.Rows > 1 Then InsertoMotivos idServicio
        If cRHora.ListIndex > -1 Then
            strHora = RetornoHoraEnString(cRHora.ItemData(cRHora.ListIndex))
        Else
            strHora = Trim(cRHora.Text)
        End If
        InsertoServicioVisita idServicio, TipoServicio.Retiro, cRCamion.ItemData(cRCamion.ListIndex), tRFecha.Tag, strHora, vsProducto.Cell(flexcpData, vsProducto.Row, 2), cRMoneda.ItemData(cRMoneda.ListIndex), _
            tRImporte.Text, cRFactura.ItemData(cRFactura.ListIndex), tRLiquidar.Text, cTipoFlete.ItemData(cTipoFlete.ListIndex)
        cBase.CommitTrans
        Foco tCi
    Else
        cBase.RollbackTrans
        MsgBox "El producto ya tiene un servicio abierto, el código de servicio es el " & idServicio, vbInformation, "ATENCIÓN"
    End If
    LimpioTodo
    Screen.MousePointer = 0: Exit Sub
ErrBR:
    clsGeneral.OcurrioError "Error al iniciar la transacción.", Trim(Err.Description)
    Screen.MousePointer = 0: Exit Sub
ErrResumir:
    Resume ErrRB
ErrRB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al intentar almacenar la información del retiro.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
Private Sub InsertoServicioVisita(idServicio As Long, TipoServ As Integer, IDCamion As Long, Fecha As String, Hora As String, Zona As Long, _
                                                        Moneda As Integer, Importe As Currency, FormaPago As Integer, LiquidarC As Currency, Optional TipoFlete As Integer = -1, Optional SinEfecto As Boolean = False, Optional TextoVisita As Integer = 0)

    Cons = "Select * From ServicioVisita Where VisServicio = " & idServicio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    RsAux.AddNew
    RsAux!VisServicio = idServicio
    RsAux!VisTipo = TipoServ
    RsAux!VisCamion = IDCamion
    RsAux!VisFecha = Format(Fecha, sqlFormatoF)
    If Trim(Hora) <> "" Then RsAux!VisHorario = Hora
    RsAux!VisZona = Zona
    If TipoFlete <> -1 Then RsAux!VisTipoFlete = TipoFlete
    RsAux!VisMoneda = Moneda
    RsAux!VisCosto = Importe
    RsAux!VisFormaPago = FormaPago
    If Trim(tAclaracion.Text) <> "" Then RsAux!VisComentario = Trim(tAclaracion.Text)
    RsAux!VisLiquidarAlCamion = LiquidarC
    RsAux!VisSinEfecto = SinEfecto
    If TextoVisita > 0 Then RsAux!VisTexto = TextoVisita
    RsAux!VisFModificacion = Format(gFechaServidor, sqlFormatoFH)
    RsAux!VisUsuario = tUsuario.Tag
    RsAux.Update
    RsAux.Close
    
End Sub

Private Function TieneReporteAbierto(idProducto As Long) As Long
On Error GoTo ErrTRA
Dim RsSA As rdoResultset
    
    Screen.MousePointer = 11
    TieneReporteAbierto = 0
    Cons = "Select * From Servicio Where SerProducto = " & idProducto _
        & " And SerEstadoServicio Not IN (" & EstadoS.Anulado & ", " & EstadoS.Cumplido & ")"
    Set RsSA = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsSA.EOF Then TieneReporteAbierto = RsSA!SerCodigo
    RsSA.Close
    Screen.MousePointer = 0
    Exit Function
ErrTRA:
    clsGeneral.OcurrioError "Error al verificar si existe algún servicio abierto.", Trim(Err.Description)
    Screen.MousePointer = 0
End Function
Private Function RetornoHoraEnString(IDHora As Integer) As String
    
    RetornoHoraEnString = ""
    Cons = "Select * From CodigoTexto Where Codigo = " & IDHora
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Clase) Then RetornoHoraEnString = Format(RsAux!Clase, "0000") & "-"
        If Not IsNull(RsAux!Puntaje) Then RetornoHoraEnString = RetornoHoraEnString & Format(RsAux!Puntaje, "0000")
    End If
    RsAux.Close
    
End Function

Private Sub AccionGrabarVisita(bEsCambio As Boolean)
On Error GoTo ErrBV
Dim idServicio As Long, strHora As String
    idServicio = 0
    Screen.MousePointer = 11
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo ErrResumir
    idServicio = TieneReporteAbierto(vsProducto.Cell(flexcpValue, vsProducto.Row, 0))
    If idServicio = 0 Then
        'If Trim(tTelefono.Text) <> "" Then GraboTelefono
        idServicio = InsertoServicio(vsProducto.Cell(flexcpValue, vsProducto.Row, 0), CInt(vsProducto.Cell(flexcpData, vsProducto.Row, 3)), EstadoS.Visita, 0, tTelefonoServicio.Text, IIf(chkNoDeseaSMS.Value = 1, 0, 1), Trim(tAclaracion.Text), -1, tUsuario.Tag, Trim(tComentarioInterno.Text))
        If vsMotivos.Rows > 1 Then InsertoMotivos idServicio
        If cVHora.ListIndex > -1 Then
            strHora = RetornoHoraEnString(cVHora.ItemData(cVHora.ListIndex))
        Else
            strHora = Trim(cVHora.Text)
        End If
        Dim aValor As Integer
        If cVComentario.ListIndex > -1 Then aValor = cVComentario.ItemData(cVComentario.ListIndex) Else aValor = 0
          InsertoServicioVisita idServicio, TipoServicio.Visita, cVCamion.ItemData(cVCamion.ListIndex), tVFecha.Tag, strHora, vsProducto.Cell(flexcpData, vsProducto.Row, 2), cVMoneda.ItemData(cVMoneda.ListIndex), _
            tVImporte.Text, cVFactura.ItemData(cVFactura.ListIndex), tVLiquidar.Text, , , aValor
        cBase.CommitTrans
        Foco tCi
    Else
        cBase.RollbackTrans
        MsgBox "El producto ya tiene un servicio abierto, el código de servicio es el " & idServicio, vbInformation, "ATENCIÓN"
    End If
    LimpioTodo
    
    If bEsCambio And idServicio <> 0 Then
        EjecutarApp App.Path & "\CambioArticulo.exe", "S " & CStr(idServicio)
    End If
    
    Screen.MousePointer = 0
    Exit Sub
ErrBV:
    clsGeneral.OcurrioError "Error al iniciar la transacción.", Trim(Err.Description)
    Screen.MousePointer = 0
    Exit Sub
ErrResumir:
    Resume ErrRB
ErrRB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al intentar almacenar la información de la visita.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
Private Sub ArmoXDefectoVisita()
    
    If cVCamion.Enabled Then
        If paMonedaPesos > 0 Then BuscoCodigoEnCombo cVMoneda, paMonedaPesos
        tVFecha.Tag = Format(Date, FormatoFP): tVFecha.Text = Format(tVFecha.Tag, "ddd d/mm/yy")
        If paCamionRetiroVisita > 0 Then
            BuscoCodigoEnCombo cVCamion, paCamionRetiroVisita
        Else
            Cons = "Select * From CamionZona " _
                & " Where CZoZona = " & Val(vsProducto.Cell(flexcpData, vsProducto.Row, 2)) _
                & " Order by CZoPrioridad "
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then BuscoCodigoEnCombo cVCamion, RsAux!CZoCamion
            RsAux.Close
        End If
        'Para seguir armando por defecto busco el primer tipo de flete que tiene asignado el camión.
        If cVCamion.ListIndex > -1 Then ArmoTipoFleteVisita cVCamion.ItemData(cVCamion.ListIndex)
    End If
    
End Sub

Private Sub CargoClienteEmpresa()
    If paClienteEmpresa > 0 Then
        LimpioTodo
        gTipoCliente = 0: gCliente = paClienteEmpresa
        LimpioFichaCliente
        CargoDatosCliente paClienteEmpresa       'Cargo Datos del Cliente Seleccionado
        'Por defecto doy nuevo producto.
        EjecutarApp App.Path & "\Productos", "N" & CStr(gCliente), True
        Me.Refresh
        DeshabilitoIngreso
        CargoDatosProducto gCliente
    Else
        MsgBox "No se cargó el parámetro cliente empresa.", vbExclamation, "ATENCIÓN"
    End If
End Sub

Private Sub HagoCambioDeEstado(IDArticulo As Long, EstadoNuevo As Integer, idServicio As Long)
    
    'Cambio el estado del artículo como Sano a Recuperar.
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, EstadoNuevo, 1, TipoDocumento.ServicioCambioEstado, idServicio
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoArticuloEntrega, -1, TipoDocumento.ServicioCambioEstado, idServicio
        
    MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, EstadoNuevo, 1, 1
    MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, 1, -1
    
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, EstadoNuevo, 1
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoArticuloEntrega, -1
    
End Sub

Private Sub CargoHistoria(idProducto As Long)
Dim idServicio As Long: idServicio = 0
Dim aComentario As String

    On Error GoTo ErrCH
    Screen.MousePointer = 11
    vsHistoria.Rows = 1
    Cons = "Select * From Servicio" _
            & " Left Outer Join ServicioRenglon ON SReTipoRenglon = " & TipoRenglonS.Cumplido _
                                                                    & " And SerCodigo = SReServicio" _
            & " Left Outer Join Articulo ON SReMotivo = ArtID " _
        & " Where SerProducto = " & idProducto _
        & " And SerEstadoServicio In (" & EstadoS.Cumplido & ", " & EstadoS.Anulado & ")" _
        & " Order By SerCodigo DESC"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsAux.EOF
        With vsHistoria
            .AddItem RsAux!SerCodigo
            If RsAux!SerEstadoServicio = EstadoS.Anulado Then .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!SerFCumplido, "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, 3) = EstadoProducto(RsAux!SerEstadoProducto)
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!SerFecha, "dd/mm/yy")
            If Not IsNull(RsAux!SerMoneda) And Not IsNull(RsAux!SerCostoFinal) Then .Cell(flexcpText, .Rows - 1, 5) = BuscoSignoMoneda(RsAux!SerMoneda) & " " & Format(RsAux!SerCostoFinal, FormatoMonedaP)
            If Not IsNull(RsAux!SerComentarioR) Then aComentario = Trim(RsAux!SerComentarioR) Else aComentario = ""
            idServicio = RsAux!SerCodigo
            Do While Not RsAux.EOF
                If Not IsNull(RsAux!ArtNombre) Then
                    If Trim(.Cell(flexcpText, .Rows - 1, 2)) <> "" Then .Cell(flexcpText, .Rows - 1, 2) = .Cell(flexcpText, .Rows - 1, 2) & ", "
                    .Cell(flexcpText, .Rows - 1, 2) = .Cell(flexcpText, .Rows - 1, 2) & Trim(RsAux!ArtNombre)
                End If
                idServicio = RsAux!SerCodigo
                RsAux.MoveNext
                If RsAux.EOF Then Exit Do
                If idServicio <> RsAux!SerCodigo Then Exit Do
            Loop
            If Trim(.Cell(flexcpText, .Rows - 1, 2)) <> "" And Trim(aComentario) <> "" Then .Cell(flexcpText, .Rows - 1, 2) = .Cell(flexcpText, .Rows - 1, 2) & Chr(13)
            .Cell(flexcpText, .Rows - 1, 2) = .Cell(flexcpText, .Rows - 1, 2) & aComentario
        End With
    Loop
    RsAux.Close
    With vsHistoria
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 2, , False
    End With
    Screen.MousePointer = 0
    Exit Sub
ErrCH:
    clsGeneral.OcurrioError "Error al buscar la historia del producto.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Function BuscoSignoMoneda(IdMoneda As Long)
Dim RsMon As rdoResultset
    BuscoSignoMoneda = ""
    Cons = "Select * From Moneda Where MonCodigo = " & IdMoneda
    Set RsMon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsMon.EOF Then BuscoSignoMoneda = Trim(RsMon!MonSigno)
    RsMon.Close
End Function
Private Sub ArmoTipoFleteVisita(IDCamion As Integer)
    
    'Por defecto pongo que factura camión
    BuscoCodigoEnCombo cVFactura, FacturaServicio.Camion: tVLiquidar.Text = "0.00"
    
    douAgenda = 0: douHabilitado = 0: strCierre = ""
    
    Cons = "Select * From CamionFlete INNER JOIN TipoFlete ON CTFTipoFlete = TFlCodigo LEFT OUTER JOIN FleteAgendaZona " & _
        " ON FAZZona = " & Val(vsProducto.Cell(flexcpData, vsProducto.Row, 2)) & " And FAZTipoFlete = TFLCodigo " & _
        " Where CTFCamion = " & IDCamion
       
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then
    
        If Not IsNull(RsAux!TFlFormaPago) Then
            'Factura Camión = 3
            If RsAux!TFlFormaPago = 3 Then
                BuscoCodigoEnCombo cVFactura, FacturaServicio.Camion
            Else
                BuscoCodigoEnCombo cVFactura, RsAux!TFlFormaPago
            End If
        End If
        
        If IsNull(RsAux!TFlAgenda) Then
            tVFecha.Tag = "": tVFecha.Text = "": cVHora.ListIndex = -1
        Else
            tVFecha.Tag = BuscoPrimerDiaAEnviar
            If tVFecha.Tag = "" Then
                tVFecha.Text = ""
                cVHora.Clear
            Else
                tVFecha.Text = Format(CDate(tVFecha.Tag), "ddd d/mm/yy")
                CargoHoraEntregaParaDia cVHora, tVFecha.Tag
            End If
        End If
        
        Dim IdTipoFlete As Integer
        IdTipoFlete = RsAux!TFlCodigo
        RsAux.Close
        
        If cVMoneda.ListIndex = -1 Then BuscoCodigoEnCombo cVMoneda, paMonedaPesos
        
        Dim cVF As Currency, cLi As Currency
        tVImporte.Text = ""
        loc_DefinoPrecioFlete IdTipoFlete, vsProducto.Cell(flexcpData, vsProducto.Row, 2), cVF, cLi
        tVImporte.Text = Format(cVF, FormatoMonedaP)
        If cVF > 0 And cVFactura.ItemData(cVFactura.ListIndex) <> FacturaServicio.Camion Then tVLiquidar.Text = Format(cLi, FormatoMonedaP) Else tVLiquidar.Text = "0.00"
    End If
    
End Sub

Private Sub CargoTextoVisita()
On Error GoTo ErrCC
    Cons = "Select TViCodigo, TViTexto From TextoVisita Order By TViTexto"
    CargoCombo Cons, cVComentario
    Exit Sub
ErrCC:
    clsGeneral.OcurrioError "Error al cargar los textos de visita."
    Screen.MousePointer = 0
End Sub

'------------------------------------------------------------------------------------------------------------------------------------
'   Setea la impresora pasada como parámetro como: por defecto
'------------------------------------------------------------------------------------------------------------------------------------
Private Sub SeteoImpresoraPorDefecto(DeviceName As String)
Dim X As Printer

    For Each X In Printers
        If Trim(X.DeviceName) = Trim(DeviceName) Then
            Set Printer = X
            Exit For
        End If
    Next
    
End Sub

Private Function CopioDireccion(lnCodDireccion As Long)

    'Copio la Direccion
    Screen.MousePointer = 11
    On Error GoTo errorBT
    Dim RsDO As rdoResultset
    Dim RsDC As rdoResultset
    
    CopioDireccion = 0
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    'Direccion ORIGINAL
    Cons = "Select * from Direccion Where DirCodigo = " & lnCodDireccion
    Set RsDO = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Direccion COPIA
    Cons = "Select * from Direccion Where DirCodigo = " & lnCodDireccion
    Set RsDC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    RsDC.AddNew
    If Not IsNull(RsDO!DirComplejo) Then RsDC!DirComplejo = RsDO!DirComplejo
    RsDC!DirCalle = RsDO!DirCalle
    RsDC!DirPuerta = RsDO!DirPuerta
    RsDC!DirBis = RsDO!DirBis
    If Not IsNull(RsDO!DirLetra) Then RsDC!DirLetra = RsDO!DirLetra
    If Not IsNull(RsDO!DirApartamento) Then RsDC!DirApartamento = RsDO!DirApartamento
    
    If Not IsNull(RsDO!DirCampo1) Then RsDC!DirCampo1 = RsDO!DirCampo1
    If Not IsNull(RsDO!DirSenda) Then RsDC!DirSenda = RsDO!DirSenda
    If Not IsNull(RsDO!DirCampo2) Then RsDC!DirCampo2 = RsDO!DirCampo2
    If Not IsNull(RsDO!DirBloque) Then RsDC!DirBloque = RsDO!DirBloque
    
    If Not IsNull(RsDO!DirEntre1) Then RsDC!DirEntre1 = RsDO!DirEntre1
    If Not IsNull(RsDO!DirEntre2) Then RsDC!DirEntre2 = RsDO!DirEntre2
    If Not IsNull(RsDO!DirAmpliacion) Then RsDC!DirAmpliacion = RsDO!DirAmpliacion
    RsDC!DirConfirmada = RsDO!DirConfirmada
    If Not IsNull(RsDO!DirVive) Then RsDC!DirVive = RsDO!DirVive
    
    RsDC.Update
    RsDC.Close: RsDO.Close
                
    Cons = "Select Max(DirCodigo) from Direccion"
    Set RsDC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    CopioDireccion = RsDC(0)
    RsDC.Close
    
    cBase.CommitTrans       'FIN TRANSACCION------------------------------------------
    Screen.MousePointer = vbDefault
    Exit Function
    
errorBT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al intentar iniciar la transacción."
    Exit Function

errorET:
    Resume ErrTransaccion

ErrTransaccion:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al intentar copiar la dirección."

End Function

Private Sub PrueboBandejaImpresora()
On Error GoTo ErrPBI
    On Error Resume Next
    With vsFicha
        .MarginTop = 300
        .MarginLeft = 550
        .PageBorder = pbNone
        .Device = paPrintConfD
        If .Device <> paPrintConfD Then MsgBox "Ud no tiene instalada la impresora para imprimir Conformes. Avise al administrador.", vbExclamation, "ATENCIÒN"
        .PaperBin = paPrintConfB
        If .PaperBin <> paPrintConfB Then
            MsgBox "Está mal definida la bandeja para imprimir fichas, comuniquele al administrador.", vbInformation, "ATENCIÓN": paPrintConfB = .PaperBin
        End If
        .PaperSize = paPrintConfPaperSize
    End With
    Exit Sub
ErrPBI:
    clsGeneral.OcurrioError "Error al setear la impresora, consulte con el administrador de impresión este problema.", Err.Description
End Sub
'Private Function s_SetAgenda(ByVal lZona As Long) As Boolean
'Dim RsF As rdoResultset
'
'    Cons = "Select * From FleteAgendaZona " & _
'            " Where FAZZona = " & lZona & " And FAZTipoFlete = " & cTipoFlete.ItemData(cTipoFlete.ListIndex)
'
'    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'    If Not RsF.EOF Then
'        With arrDatosFlete(iIndex)
'            If Not IsNull(RsF!FAZAgenda) Then
'                .Agenda = RsF!FAZAgenda
'                If Not IsNull(RsF!FAZAgendaHabilitada) Then .AgendaAbierta = RsF!FAZAgendaHabilitada Else .AgendaAbierta = .Agenda
'                If Not IsNull(RsF!FAZFechaAgeHab) Then .AgendaCierre = RsF("FAZFechaAgeHab")
'            End If
'            If Not IsNull(RsF!FAZRangoHS) Then .HorarioRango = RsF!FAZRangoHS
'            If Not IsNull(RsF!FAZHoraEnvio) Then .HoraEnvio = Trim(RsF!FAZHoraEnvio)
'        End With
'    End If
'    RsF.Close
'
'
'
'End Function

Private Function BuscoPrimerDiaAEnviar() As String
Dim strMat As String
Dim intSumaDia As Integer

    BuscoPrimerDiaAEnviar = ""
        
    strMat = ""
    If Not IsNull(RsAux!FAZFechaAgeHab) Then
        strCierre = RsAux("FAZFechaAgeHab")
    
        douAgenda = RsAux!FAZAgenda
        If IsNull(RsAux!TFlAgendaHabilitada) Then douHabilitado = douAgenda Else douHabilitado = RsAux!FAZAgendaHabilitada
        
        If IsDate(strCierre) Then
            If DateDiff("d", CDate(strCierre), Date) >= 7 Then
                'Como cerro hace una semana tomo la agenda normal.
                strMat = superp_MatrizSuperposicion(douAgenda)
            Else
                strMat = superp_MatrizSuperposicion(douHabilitado)
            End If
            If CDate(strCierre) < Date Then strCierre = Date
        Else
            strMat = superp_MatrizSuperposicion(douAgenda)
            strCierre = Date
        End If
    
    Else
        If IsNull(RsAux!TFlFechaAgeHab) Then strCierre = Date Else strCierre = RsAux!TFlFechaAgeHab
        douAgenda = RsAux!TFlAgenda
        If IsNull(RsAux!TFlAgendaHabilitada) Then douHabilitado = douAgenda Else douHabilitado = RsAux!TFlAgendaHabilitada
        
        If IsDate(strCierre) Then
            If DateDiff("d", CDate(strCierre), Date) >= 7 Then
                'Como cerro hace una semana tomo la agenda normal.
                strMat = superp_MatrizSuperposicion(douAgenda)
            Else
                strMat = superp_MatrizSuperposicion(douHabilitado)
            End If
            If CDate(strCierre) < Date Then strCierre = Date
        Else
            strMat = superp_MatrizSuperposicion(douAgenda)
            strCierre = Date
        End If
    End If
    
    If strMat <> "" Then
        intSumaDia = BuscoProximoDia(strCierre, strMat, douAgenda)
        If intSumaDia <> -1 Then BuscoPrimerDiaAEnviar = Format(CDate(strCierre) + intSumaDia, "d-mm-yy")
    Else
        MsgBox "No existe una agenda de reparto ingresada para el tipo de flete seleccionado.", vbInformation, "ATENCIÓN"
    End If
    
End Function
Private Function BuscoProximoDia(strFecha As String, strMat As String, douAgenda As Double)
Dim rsHora As rdoResultset
Dim intDia As Integer, intSuma As Integer
    
    'Por las dudas que no cumpla en la semana paso la agenda normal.
    
    On Error GoTo errBDER
    
    BuscoProximoDia = -1
    
    'Consulto en base a la matriz devuelta.
    Cons = "Select * From HorarioFlete Where HFlIndice IN (" & strMat & ")"
    Set rsHora = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsHora.EOF Then
        
        'Busco el valor que coincida con el dia de hoy y ahí busco para arriba.
        intSuma = 0
        Do While intSuma < 7
            rsHora.MoveFirst
            intDia = Weekday(CDate(strFecha) + intSuma)
            Do While Not rsHora.EOF
                If rsHora!HFlDiaSemana = intDia Then
                    If douHabilitado = 0 Then
                        'Esta toda la semana cerrada para el dia de cierre y entró con la agenda normal entonces veo el cierre.
                        If Abs(DateDiff("d", strCierre, CDate(strFecha))) < 7 Then
                            intSuma = intSuma + 7
                        End If
                    End If
                    BuscoProximoDia = intSuma
                    GoTo Encontre
                End If
                rsHora.MoveNext
            Loop
            intSuma = intSuma + 1
        Loop
        rsHora.Close
        
    End If
    
    'Si llega a este punto por las siguientes posibilidades.
    '1 _ dio eof la consulta. Lo peor sería que la matriz este formada por la agenda normal
    '2_ no encontre ningún día habilitado, busco en la normal para la semana siguiente.
    '3_ hay error en la formula almacenada en la tabla y la vista no retorna nada.
    
    strMat = superp_MatrizSuperposicion(douAgenda)
    
    Cons = "Select * From HorarioFlete Where HFlIndice IN (" & strMat & ")"
    Set rsHora = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    If Not rsHora.EOF Then
        intSuma = 7
        Do While intSuma < 14
            rsHora.MoveFirst
            intDia = Weekday(CDate(strFecha) + intSuma)
            Do While Not rsHora.EOF
                If rsHora!HFlDiaSemana = intDia Then
                    BuscoProximoDia = intSuma
                    GoTo Encontre
                End If
                rsHora.MoveNext
            Loop
            intSuma = intSuma + 1
        Loop
    End If

Encontre:
    rsHora.Close
    Exit Function
    
errBDER:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el primer día disponible para el tipo de flete.", Trim(Err.Description)
End Function

Private Sub CargoHoraEntregaParaDia(Combo As Control, strFecha As String)
On Error GoTo errCHEPD
Dim strMat As String
    
    Combo.Clear
    If IsDate(strFecha) And IsDate(strCierre) Then
        strMat = ""
        If Abs(DateDiff("d", strCierre, CDate(strFecha))) >= 7 Or douHabilitado = 0 Then
            strMat = superp_MatrizSuperposicion(douAgenda)
        Else
            strMat = superp_MatrizSuperposicion(douHabilitado)
        End If
        If strMat <> "" Then
            Cons = "Select HFlCodigo, HFlNombre From HorarioFlete Where HFlIndice IN (" & strMat & ")" _
                & " And HFlDiaSemana = " & Weekday(CDate(strFecha)) & " Order by HFlInicio"
            CargoCombo Cons, Combo
            If Combo.ListCount > 0 Then
                Combo.ListIndex = 0
            Else
                MsgBox "No está abierto el envío para el tipo de flete seleccionado, verifique.", vbExclamation, "ATENCIÓN"
            End If
        End If
    Else
        If Not IsDate(strCierre) Then MsgBox "No existe una agenda de reparto para el tipo de flete seleccionado.", vbInformation, "ATENCIÓN"
    End If
    Exit Sub
    
errCHEPD:
    clsGeneral.OcurrioError "Ocurrió un error al buscar los horarios para el día de semana.", Trim(Err.Description)
End Sub

Private Sub EsPosibleReclamo()
Dim iCont As Integer
    'Veo si es posible reclamo, recorro lista historia.
    If vsHistoria.Rows > 1 Then
        For iCont = 1 To vsHistoria.Rows - 1
            If vsHistoria.Cell(flexcpBackColor, iCont, 0) <> Colores.Inactivo Then
                'Primera válida
                If Abs(DateDiff("d", CDate(vsHistoria.Cell(flexcpText, iCont, 1)), Now)) <= 30 Then
                    If chEsReclamo.Value = 0 Then
                        chEsReclamo.Value = 1
                        tReclamo.Text = vsHistoria.Cell(flexcpText, iCont, 0)
                        tReclamo.Tag = tReclamo.Text
                    End If
                ElseIf Abs(DateDiff("d", CDate(vsHistoria.Cell(flexcpText, iCont, 1)), Now)) < 45 Then
                    MsgBox "Existe un servicio con menos de 45 días." & vbCrLf & "Verifique en la historia que este servicio no sea un reclamo del mismo.", vbInformation, "ATENCIÓN"
                End If
                Exit For
            End If
        Next
    Else
        chEsReclamo.Value = 0
        chEsReclamo.Enabled = False
        tReclamo.Text = ""
    End If
End Sub

Private Function ReclamoValido(ByVal idServ As Long) As Boolean
Dim iCont As Integer
    ReclamoValido = False
    'Veo si es posible reclamo, recorro lista historia.
    If vsHistoria.Rows > 1 Then
        For iCont = 1 To vsHistoria.Rows - 1
            If CLng(vsHistoria.Cell(flexcpText, iCont, 0)) = idServ And vsHistoria.Cell(flexcpBackColor, iCont, 0) <> Colores.Inactivo Then
                ReclamoValido = True
                Exit For
            End If
        Next
    End If
End Function

Public Sub loc_FindComentarios(idCliente As Long)
Dim rsCom As rdoResultset
Dim bHay As Boolean

    On Error GoTo errMenu
    Screen.MousePointer = 11
    bHay = False
    
    Cons = "Select * From Comentario " _
            & " Where ComCliente = " & idCliente _
            & " And ComTipo IN (" & prmTipoComentario & ")"
            
    Set rsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not rsCom.EOF Then bHay = True
    rsCom.Close
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

Private Function fnc_GetTipoArticulo() As Long
On Error Resume Next
Dim rsT As rdoResultset
        Set rsT = cBase.OpenResultset("Select ArtTipo From Articulo Where ArtID = " & vsProducto.Cell(flexcpData, vsProducto.Row, 0), _
                                                    rdOpenDynamic, rdConcurValues)
        If Not rsT.EOF Then fnc_GetTipoArticulo = rsT(0)
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
Private Sub CumploServicio(ByVal iServicio As Long)
Dim rsS As rdoResultset
Dim iCosto As Currency, iDoc As Long, iProEdit As Long
On Error GoTo errCumplir

    
    Cons = "SELECT * FROM Servicio WHERE SerCodigo = " & iServicio & " AND SerEstadoServicio <> " & EstadoS.Cumplido
    Set rsS = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsS.EOF Then
        iCosto = 0
        If Not IsNull(rsS("SerCostoFinal")) Then iCosto = rsS("SerCostoFinal")
        If Not IsNull(rsS("SerDocumento")) Then iDoc = rsS("SerDocumento")
    
        If (iCosto > 0 And iDoc = 0) And rsS("SerCliente") <> paClienteEmpresa Then
            MsgBox "No se puede cumplir este servicio ya que el mismo tiene pendiente el pago.", vbExclamation, "Atención"
            rsS.Close
            Exit Sub
        End If
        
        If MsgBox("¿Confirma cumplir el servicio?", vbQuestion + vbYesNo, "CUMPLIR SERVICIO") = vbYes Then
            
            FechaDelServidor
            iProEdit = rsS("SerProducto")
            
            rsS.Edit
            rsS!SerFCumplido = Format(gFechaServidor, sqlFormatoF)
            rsS!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
            rsS!SerUsuario = paCodigoDeUsuario
            rsS!SerEstadoServicio = EstadoS.Cumplido
            rsS.Update
            
        End If
    End If
    rsS.Close
    
    If iProEdit > 0 Then
        vsProducto.Tag = Val(vsProducto.Tag) - CantTuplas
        CargoDatosProducto gCliente
        On Error Resume Next
        Dim iR As Integer
        For iR = 1 To vsProducto.Rows - 1
            If Val(vsProducto.Cell(flexcpValue, iR, 0)) = iProEdit Then
                vsProducto.Select iR, 0
                If Not vsProducto.RowIsVisible(iR) Then
                    vsProducto.Row = iR
                    vsProducto.TopRow = iR
                End If
                'iR = vsProducto.RowPos(iR)
                Exit For
            End If
        Next
    End If
    
    Exit Sub
errCumplir:
    clsGeneral.OcurrioError "Error al dar cumplir servicio.", Err.Description, "Cumplir servicio"
End Sub

Private Function MsgClienteNoVender(ByVal iCliente As Long, ByVal bShowMsg As Boolean) As Boolean
Dim rsCom As rdoResultset
    MsgClienteNoVender = False
    Set rsCom = cBase.OpenResultset("exec gennovender " & iCliente, rdOpenDynamic, rdConcurValues)
    If Not rsCom.EOF Then
        If Not IsNull(rsCom(0)) Then
            If rsCom(0) = 1 Then
                MsgClienteNoVender = True
                If bShowMsg Then
                    Screen.MousePointer = 0
                    MsgBox "Atención: NO se puede agregar un servicio sin autorización. Consultar con gerencia!", vbCritical, "ATENCIÓN"
                End If
            End If
        End If
    End If
    rsCom.Close
End Function

Private Sub ImprimirFichaServicio(ByVal idServicio As Long, ByVal telefono As String)
On Error GoTo errFD
Dim iCont As Integer
Dim oPrint As clsPrintReport
    
    Set oPrint = New clsPrintReport
    With oPrint
        .StringConnect = miConexion.TextoConexion("comercio")
        .DondeImprimo.Bandeja = paPrintConfB
        .DondeImprimo.Impresora = paPrintConfD
        .DondeImprimo.Papel = paPrintConfPaperSize
        .PathReportes = paPathReportes
    End With
    
    If (telefono <> "") Then
        telefono = "Acepto recibir toda comunicación y/o aviso (inclusive de naturaleza confidencial, crediticia y/o personal) al teléfono " & telefono & " por cualquier vía (sms, whatsapp, etc.)"
    End If
    
    Dim sQueryServicio As String
    sQueryServicio = "SELECT SerCodigo infoCodigoServicio, '*S'+ RTRIM(convert(varchar(20), SerCodigo)) + '*' infoCodigoBarras " & _
            ", dbo.FormatCIRuc(CliCIRUC) + ' ' + CASE WHEN CliTipo = 1 THEN rtrim(CPeNombre1) + ' ' + RTRIM(CPeApellido1) ELSE RTRIM(CEmNombre) END infoCliente " & _
            ", dbo.FormatDate(SerFecha, 121) infoFecha " & _
            ",'(' + rtrim(Convert(varchar(10), ProCodigo)) + ') ' + RTRIM(ArtNombre) + ' (' + Convert(varchar(15), ArtCodigo) + ')' infoArticulo " & _
            ", RTRIM(usuidentificacion) infoRecibio " & _
            ", dbo.TelefonosCliente(CliCodigo) infoTelefonos, '{1}' infoMotivos, IsNull(SerComentario, '') infoMemoIngreso " & _
            ", dbo.ArmoDireccion(CliDireccion) infoDireccion, IsNull(IsNull(ProFacturaS, '') + ' ' + CONVERT(varchar(6), ProFacturaN), '') infoFactura " & _
            ", Rtrim(SucI.SucAbreviacion) infoLocal, Rtrim(IsNull(SucS.SucAbreviacion, '')) infoLocalRepara, RTRIM(ProNroSerie) infoNroSerie, ISNull(dbo.FormatDate(ProCompra, 2), '') infoFCompra, '" & telefono & "' infoAceptoSMS " & _
            "FROM Servicio INNER JOIN Cliente ON SerCliente = CliCodigo " & _
            "LEFT OUTER JOIN CPersona ON CliCodigo = CPeCliente " & _
            "LEFT OUTER JOIN CEmpresa ON CliCodigo = CEmCliente " & _
            "INNER JOIN Producto ON SerProducto = ProCodigo " & _
            "INNER JOIN Articulo ON ProArticulo = ArtId " & _
            "INNER JOIN Sucursal SucI ON SerLocalIngreso = SucI.SucCodigo " & _
            "LEFT OUTER JOIN Sucursal SucS ON SerLocalReparacion = SucS.SucCodigo " & _
            "INNER JOIN Usuario ON SerUsuario = UsuCodigo " & _
            "WHERE SerCodigo = {0}"
            
    
    Dim sMotivos As String
    For I = 1 To vsMotivos.Rows - 1
        If sMotivos = "" Then sMotivos = Trim(vsMotivos.Cell(flexcpText, I, 0)) Else sMotivos = sMotivos & ", " & Trim(vsMotivos.Cell(flexcpText, I, 0))
    Next I
    
    Dim query As String
    query = Replace(sQueryServicio, "{0}", idServicio)
    query = Replace(query, "{1}", sMotivos)
    oPrint.Imprimir_vsReport "FichaServicio.xml", "FichaDeServicio", query, "", ""
    Exit Sub
    
errFD:
    clsGeneral.OcurrioError "Error al imprimir las fichas.", Err.Description, "Fichas de devolución"
End Sub


