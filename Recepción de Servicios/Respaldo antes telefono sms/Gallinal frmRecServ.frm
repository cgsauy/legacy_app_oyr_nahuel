VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACOMBO.OCX"
Begin VB.Form frmRecServ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Servicios"
   ClientHeight    =   6075
   ClientLeft      =   2190
   ClientTop       =   3270
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
   Icon            =   "Gallinal frmRecServ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
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
   Begin VB.TextBox tInterno 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   69
      Top             =   5400
      Width           =   1155
   End
   Begin AACombo99.AACombo cTipoTelefono 
      Height          =   315
      Left            =   1020
      TabIndex        =   67
      Top             =   5400
      Width           =   1635
      _ExtentX        =   2884
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
   Begin VB.TextBox tTelefono 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2700
      MaxLength       =   15
      TabIndex        =   68
      Top             =   5400
      Width           =   1575
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   75
      Top             =   5805
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   476
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
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
      Top             =   5400
      Width           =   615
   End
   Begin VB.PictureBox picHistoria 
      Height          =   735
      Left            =   1740
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   73
      Top             =   4020
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
         MaxLength       =   14
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
         Picture         =   "Gallinal frmRecServ.frx":030A
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
   Begin VB.PictureBox PicVisita 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   72
      Top             =   4620
      Width           =   1335
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
   Begin VB.TextBox tAclaracion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1020
      MaxLength       =   75
      TabIndex        =   65
      Top             =   5100
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
      Picture         =   "Gallinal frmRecServ.frx":0494
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   780
      Width           =   315
   End
   Begin VB.PictureBox PicRetiro 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   63
      Top             =   4200
      Width           =   1575
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
   Begin VB.PictureBox PicTaller 
      Height          =   435
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1035
      TabIndex        =   60
      Top             =   3720
      Width           =   1095
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
      _Version        =   327681
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
      _Version        =   327681
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
   Begin VB.Image Image2 
      Height          =   300
      Left            =   6000
      Picture         =   "Gallinal frmRecServ.frx":061E
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "T&eléfono:"
      Height          =   255
      Left            =   120
      TabIndex        =   66
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "&Aclaración:"
      Height          =   255
      Left            =   120
      TabIndex        =   74
      Top             =   5100
      Width           =   855
   End
   Begin VB.Label ltUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   255
      Left            =   7620
      TabIndex        =   70
      Top             =   5400
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
            Picture         =   "Gallinal frmRecServ.frx":0928
            Key             =   "taller"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Gallinal frmRecServ.frx":0F72
            Key             =   "retiro"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Gallinal frmRecServ.frx":128C
            Key             =   "visita"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Gallinal frmRecServ.frx":15A6
            Key             =   "historia"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Gallinal frmRecServ.frx":18C0
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
      Caption         =   "&Productos"
      Begin VB.Menu MnuProFicha 
         Caption         =   "&Ficha del Producto"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuProCederProducto 
         Caption         =   "&Ceder el Producto a Otro Cliente"
      End
      Begin VB.Menu MnuProLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuProSeguimientoReporte 
         Caption         =   "&Seguimiento de Servicio"
      End
      Begin VB.Menu MnuProHistoria 
         Caption         =   "&Historia de Servicio"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuSalDel 
         Caption         =   "Del Formulario"
         Shortcut        =   ^X
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
    
Option Explicit
'Parametros para seleccionar tipo de DatosIngreso
'prmTipoIngreso    Taller = 1,    Retiro = 2,    Visita = 3
'prmInvocacion Ingreso de servicio para nota con cliente CGSA
'                       Si es T va directo a taller, si es R va a retiro.
'prmArticulo trae el ID de artículo que se le asignará al servicio
'prmDireccion id de direccion del cliente que devuelve el artículo.

Public prmTipoIngreso As Integer, prmInvocacion As String, prmArticulo As Long, prmDireccion As Long
Private strCierre As String, douHabilitado As Double, douAgenda As Double
Private gCliente As Long, gTipoCliente As Integer
Private sEsProducto As Boolean

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
    If KeyAscii = vbKeyReturn And cRDeposito.ListIndex > -1 Then Foco tAclaracion
End Sub

Private Sub cRFactura_Click()
    If cRFactura.ListIndex > -1 Then
        If cRFactura.ItemData(cRFactura.ListIndex) = FacturaServicio.Camion Then tRLiquidar.Text = "0.00" Else tRLiquidar.Text = tRImporte.Text
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
    If KeyAscii = vbKeyReturn And cTalDeposito.ListIndex > -1 Then Foco tAclaracion
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

Private Sub cTipoTelefono_GotFocus()
    With cTipoTelefono
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cTipoTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cTipoTelefono.ListIndex > -1 Then
            'Cargo el telefono para este tipo
            Cons = "Select * From Telefono Where TelCliente = " & gCliente _
                & " And TelTipo = " & cTipoTelefono.ItemData(cTipoTelefono.ListIndex)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                tTelefono.Text = clsGeneral.RetornoFormatoTelefono(RsAux!TelNumero, Val(tDirCliente.Tag), txtConexion)
                If Not IsNull(RsAux!TelInterno) Then tInterno.Text = Trim(RsAux!TelInterno)
            End If
            RsAux.Close
        End If
        Foco tTelefono
    End If
End Sub

Private Sub cTipoTelefono_LostFocus()
    cTipoTelefono.SelStart = 0
End Sub

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
    If KeyAscii = vbKeyReturn Then Foco tAclaracion
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
    If KeyAscii = vbKeyReturn Then Foco cVMoneda
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

    FechaDelServidor
    LimpioTodo
    InicializoGrillaMotivos
    InicializoGrillaProducto
    InicializoGrillaHistoria
    
    PicTaller.BorderStyle = 0: PicRetiro.BorderStyle = 0: PicVisita.BorderStyle = 0: picHistoria.BorderStyle = 0
    
    Set TabRecepcion.ImageList = Image1
    TabRecepcion.Tabs("historia").Image = Image1.ListImages("historia").Index
    TabRecepcion.Tabs("comodin").Image = Image1.ListImages("taller").Index
    
    cDatoIngreso.Clear
    cDatoIngreso.AddItem "Taller": cDatoIngreso.ItemData(cDatoIngreso.NewIndex) = 1
    cDatoIngreso.AddItem "Retiro": cDatoIngreso.ItemData(cDatoIngreso.NewIndex) = 2
    cDatoIngreso.AddItem "Visita": cDatoIngreso.ItemData(cDatoIngreso.NewIndex) = 3

    Cons = "Select TTeCodigo, TTeNombre From TipoTelefono Order by TTeNombre"
    CargoCombo Cons, cTipoTelefono, ""
    
    'Defino la acción que toma el formulario.
    If prmInvocacion <> "" Then     'Acción para realizar una Nota.
        If prmInvocacion = "T" Then
            prmTipoIngreso = 1 'Por las dudas que nos equivoquemos.
        ElseIf prmInvocacion = "R" Then
            prmTipoIngreso = 2
        End If
        cDatoIngreso.Enabled = False
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
    With vsFicha
    '    .Orientation = orPortrait: .PaperSize = 1
    '    .PaperHeight = .PaperHeight / 2
        .MarginTop = 300
        .MarginLeft = 500
    End With
    
    PrueboBandejaImpresora
    
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar el formulario.", Trim(Err.Description)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    TabRecepcion.Width = Me.ScaleWidth - (TabRecepcion.Left * 2)
    TabRecepcion.Height = tAclaracion.Top - (TabRecepcion.Top + 40)
    
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
        objDir.ActivoFormularioDireccion idDireccion, gCliente, txtConexion, "Producto", "ProDireccion", "ProCodigo", CLng(tDireccion.Tag)
    Else
        objDir.ActivoFormularioDireccion idDireccion, gCliente, txtConexion, "Cliente", "CliDireccion", "CliCodigo", gCliente
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
        If aCodDir <> 0 Then tDirCliente.Text = clsGeneral.DireccionATexto(aCodDir, Departamento:=True, Localidad:=True, Zona:=True, ConfYVD:=True, strConeccion:=txtConexion)
    Else
        bDireccionP.Tag = aCodDir: tDireccion.Text = ""
        If aCodDir <> 0 Then tDireccion.Text = clsGeneral.DireccionATexto(aCodDir, Departamento:=True, Localidad:=True, Zona:=True, ConfYVD:=True, strConeccion:=txtConexion)
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

Private Sub MnuProFicha_Click()
    If vsProducto.Rows = 1 Then EjecutarApp App.Path & "\Productos", "N" & CStr(gCliente), True Else EjecutarApp App.Path & "\Productos", CStr(gCliente), True
    Me.Refresh
    DeshabilitoIngreso
    CargoDatosProducto gCliente
End Sub

Private Sub MnuProHistoria_Click()
    If vsProducto.Row > 0 Then EjecutarApp App.Path & "\Historia Servicio", vsProducto.Cell(flexcpValue, vsProducto.Row, 0)
    Me.Refresh
End Sub

Private Sub MnuProSeguimientoReporte_Click()
    If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) = 0 Then Exit Sub
    If vsProducto.Row > 0 Then EjecutarApp App.Path & "\Seguimiento de Servicios", vsProducto.Cell(flexcpData, vsProducto.Row, 4)
End Sub

Private Sub MnuSalDel_Click()
    Unload Me
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
    If KeyAscii = vbKeyReturn Then Foco cTipoTelefono
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
            LiAyuda.ActivoListaAyuda Cons, False, txtConexion, 4800
            Me.Refresh
            If LiAyuda.ItemSeleccionado <> "" Then
                tArticulo.Text = LiAyuda.ItemSeleccionado
                If Val(lTipoProducto.Tag) <> Val(LiAyuda.ValorSeleccionado) Then
                    If MsgBox("¿Confirma modificar el tipo del artículo?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Set LiAyuda = Nothing: GoTo Abandono
                End If
                tArticulo.Tag = LiAyuda.ValorSeleccionado
                
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
            Set LiAyuda = Nothing
        
        Else                                            'Busqueda por codigo
            Cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtCodigo = " & Val(tArticulo.Text)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
            If RsAux.EOF Then
                RsAux.Close
                MsgBox "No se encontró un artículo para el código ingresado.", vbInformation, "ATENCIÓN"
            Else
                tArticulo.Text = Trim(RsAux!Nombre)
                If Val(lTipoProducto.Tag) <> Val(LiAyuda.ValorSeleccionado) Then
                    If MsgBox("¿Confirma modificar el tipo del artículo?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then GoTo Abandono
                End If
                tArticulo.Tag = RsAux!ArtID
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
    clsGeneral.OcurrioError "Ocurrio un error al modificar la fecha de compra del producto.", Trim(Err.Description)
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
    clsGeneral.OcurrioError "Ocurrio un error al modificar los datos de la factura del producto .", Trim(Err.Description)
    
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

Private Sub tInterno_GotFocus()
    With tInterno
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tInterno_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub

Private Sub tMotivo_GotFocus()
    With tMotivo
        If .Text = "" Then .Text = "%"
        If .Text = "%" Then .SelStart = Len(.Text): Exit Sub
        .SelStart = 0
        .SelLength = Len(tMotivo.Text)
    End With
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
            If IsNumeric(tMotivo.Text) Then
                Cons = "Select * From MotivoServicio " _
                    & " Where MSeTipo = (Select ArtTipo From Articulo Where ArtID = " & vsProducto.Cell(flexcpData, vsProducto.Row, 0) & ")" _
                    & " And MSeCodigo = " & tMotivo.Text
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
                If RsAux.EOF Then
                    RsAux.Close
                    MsgBox "No existe un motivo con el código ingresado para el tipo de artículo seleccionado.[F1] Ayuda", vbInformation, "ATENCIÓN"
                    Screen.MousePointer = 0: Exit Sub
                Else
                    InsertoMotivoEnGrilla
                    RsAux.Close
                End If
            Else
                Cons = "Select MSeID, Código = MSeCodigo, Nombre = MSeNombre From MotivoServicio " _
                    & " Where MSeTipo = (Select ArtTipo From Articulo Where ArtID = " & vsProducto.Cell(flexcpData, vsProducto.Row, 0) & ")" _
                    & " And MSeNombre Like '" & tMotivo.Text & "%'"
                Dim objLista As New clsListadeAyuda
                objLista.ActivoListaAyuda Cons, False, txtConexion, 4500
                If objLista.ValorSeleccionado > 0 Then
                    Cons = "Select * From MotivoServicio Where MSeID = " & objLista.ValorSeleccionado
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
                    If Not RsAux.EOF Then InsertoMotivoEnGrilla
                    RsAux.Close
                End If
                Set objLista = Nothing
            End If
            Screen.MousePointer = 0
            tMotivo.Text = ""
        End If
    End If
    Exit Sub
ErrBM:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al buscar los motivos.", Trim(Err.Description)
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
    clsGeneral.OcurrioError "Ocurrio un error al modificar los datos de la factura del producto .", Trim(Err.Description)
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
    
    If aTipoCliente = TipoCliente.Cliente Then objBuscar.ActivoFormularioBuscarClientes txtConexion, Persona:=True
    If aTipoCliente = TipoCliente.Empresa Then objBuscar.ActivoFormularioBuscarClientes txtConexion, Empresa:=True
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
Public Sub BuscoClienteSeleccionado(Codigo As Long)

Dim aCliente As Long
    Screen.MousePointer = 11
    
    gCliente = Codigo
    gTipoCliente = 0
    LimpioFichaCliente
    If gCliente > 0 Then
        CargoDatosCliente Codigo        'Cargo Datos del Cliente Seleccionado
        CargoDatosProducto Codigo    'Cargo los productos asociados al cliente.
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
    Else
        tRuc.Text = "": tRuc.Tag = ""
        tCi.Text = "": tCi.Tag = ""
    End If
    
    'Direccion
    If Not IsNull(RsAux!CliDireccion) Then
        tDirCliente.Text = clsGeneral.DireccionATexto(RsAux!CliDireccion, Departamento:=True, Localidad:=True, Zona:=True, ConfYVD:=True, strConeccion:=txtConexion)
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
        tDireccion.Text = clsGeneral.DireccionATexto(idDireccion, Departamento:=True, Localidad:=True, Zona:=True, ConfYVD:=True, strConeccion:=txtConexion)
    Else
        tDirCliente.Text = clsGeneral.DireccionATexto(idDireccion, Departamento:=True, Localidad:=True, Zona:=True, ConfYVD:=True, strConeccion:=txtConexion)
    End If
    Screen.MousePointer = 0
    
    Exit Sub
errConfirmar:
    clsGeneral.OcurrioError "Ocurrió un error al confirmar la dirección del cliente.", Err.Description
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

Private Sub CargoDatosProducto(idCliente As Long)
On Error GoTo ErrCDP
Dim aValor As Integer, fModificado As String
    
    Screen.MousePointer = 11
    LimpioCamposProducto
    vsProducto.Rows = 1
    
    Cons = "Select * From Producto, Articulo " _
        & " Where ProCliente = " & idCliente & " And ProArticulo = ArtID"
    
    If idCliente = paClienteEmpresa Or idCliente = paClienteAnglia Then Cons = Cons & " And ProFModificacion >= '" & Format(Date, "mm/dd/yyyy 00:00:00") & "'"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsAux.EOF
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
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    If vsProducto.Rows > 1 And Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) = 0 Then HabilitoParaIngreso: MuestroCamposProducto Else DeshabilitoIngreso
    If vsProducto.Rows > 1 Then CargoHistoria vsProducto.Cell(flexcpText, 1, 0)
    Screen.MousePointer = 0
    Exit Sub
ErrCDP:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los datos del producto.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub InicializoGrillaProducto()
    With vsProducto
        .Rows = 1
        .Cols = 1
        .ExtendLastCol = True
        .FormatString = "ID|Tipo de Artículo|Estado|>F.Compra|Garantía|N° Serie|Factura|"
        .ColWidth(0) = 650: .ColWidth(1) = 3000: .ColWidth(3) = 1000: .ColWidth(5) = 1100
    End With
End Sub
Private Sub LimpioObjetosComunes()
    tMotivo.Text = ""
    tAclaracion.Text = ""
    tUsuario.Text = ""
    cTipoTelefono.ListIndex = -1
    tTelefono.Text = ""
    tInterno.Text = ""
    vsMotivos.Rows = 1
    vsHistoria.Rows = 1
End Sub
Private Sub LimpioFichaTaller()
    LimpioObjetosComunes
    lTalFecha.Caption = Format(gFechaServidor, FormatoFP)
    cTalDeposito.Text = ""
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
    clsGeneral.OcurrioError "Ocurrio un error al ir a ficha de cliente.", Err.Description
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
    clsGeneral.OcurrioError "Ocurrio un error al ir a ficha de cliente.", Err.Description
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
    vsProducto.Rows = 1
End Sub

Private Sub tTelefono_GotFocus()
    With tTelefono
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tTelefono_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tInterno
End Sub

Private Sub tTelefono_LostFocus()
    If Trim(tTelefono.Text) <> "" Then
        If cTipoTelefono.ListIndex = -1 Then MsgBox "Debe seleccionar un tipo de teléfono.", vbExclamation, "ATENCIÓN": Foco cTipoTelefono: Exit Sub
        tTelefono.Tag = clsGeneral.RetornoFormatoTelefono(tTelefono.Text, Val(tDirCliente.Tag), txtConexion)
        If tTelefono.Tag <> "" Then
            tTelefono.Text = tTelefono.Tag
        Else
            MsgBox "El teléfono ingresado no coincide con los formatos establecidos.", vbExclamation, "ATENCIÓN"
            Foco tTelefono
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
            If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) > 0 Then EjecutarApp App.Path & "\Seguimiento de Servicios", vsProducto.Cell(flexcpData, vsProducto.Row, 4)
        Case 93
            If gCliente > 0 Then PopupMenu MnuProducto
    End Select
End Sub

Private Sub vsProducto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And gCliente > 0 Then PopupMenu MnuProducto
End Sub

Private Sub vsProducto_RowColChange()
On Error Resume Next
    CargoHistoria vsProducto.Cell(flexcpValue, vsProducto.Row, 0)
    OcultoCamposProducto
    If Val(vsProducto.Cell(flexcpData, vsProducto.Row, 4)) = 0 Then HabilitoParaIngreso Else DeshabilitoIngreso
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
        .Rows = 1
        .FormatString = "ID|Motivo"
        .ColWidth(0) = 550
    End With
End Sub
Private Sub InicializoGrillaHistoria()
    With vsHistoria
        .Rows = 1
        .WordWrap = True
        .FormatString = "Fecha|Motivos|Estado|Llamado|>Importe"
        .ColWidth(0) = 750: .ColWidth(1) = 5200: .ColWidth(3) = 750 ': .ColWidth(4) = 1300
        .ColAlignment(0) = flexAlignLeftTop
        .ColAlignment(2) = flexAlignLeftTop
        .ColAlignment(3) = flexAlignLeftTop
        .ColAlignment(4) = flexAlignRightTop
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
    tMotivo.Enabled = False: tMotivo.BackColor = Inactivo
    tUsuario.Enabled = False: tUsuario.BackColor = Inactivo
    cTipoTelefono.Enabled = False: cTipoTelefono.BackColor = Inactivo
    tTelefono.Enabled = False: tTelefono.BackColor = Inactivo
    tInterno.Enabled = False: tInterno.BackColor = Inactivo
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
    tMotivo.Enabled = True: tMotivo.BackColor = Blanco
    tUsuario.Enabled = True: tUsuario.BackColor = Obligatorio
    cTipoTelefono.Enabled = True: cTipoTelefono.BackColor = Blanco
    tTelefono.Enabled = True: tTelefono.BackColor = Blanco
    tInterno.Enabled = True: tInterno.BackColor = Blanco
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
    clsGeneral.OcurrioError "Ocurrio un error al cargar los camiones."
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
    clsGeneral.OcurrioError "Ocurrio un error al cargar los tipos de fletes."
    Screen.MousePointer = 0
End Sub

Private Sub CargoComboDeposito()
On Error GoTo ErrCC
    Screen.MousePointer = 11
    cTalDeposito.Clear
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cTalDeposito
    CargoCombo Cons, cRDeposito
    Screen.MousePointer = 0
    Exit Sub
ErrCC:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los Depósitos."
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
    clsGeneral.OcurrioError "Ocurrio un error al cargar las monedas."
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
    clsGeneral.OcurrioError "Ocurrio un error al cargar los tipos de facturación."
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
            If MsgBox("¿Confirma grabar la solicitud de visita?" & Chr(13) & strMsg, vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then AccionGrabarVisita
    End Select
End Sub
Private Sub AccionGrabarTaller()
On Error GoTo ErrBT
Dim IdServicio As Long
    IdServicio = 0
    Screen.MousePointer = 11
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo ErrResumir
    Cons = "Select * From Servicio Where SerProducto = " & vsProducto.Cell(flexcpValue, vsProducto.Row, 0) _
        & " And SerEstadoServicio Not IN (" & EstadoS.Anulado & ", " & EstadoS.Cumplido & ")"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        If Trim(tTelefono.Text) <> "" Then GraboTelefono
        IdServicio = InsertoServicio(vsProducto.Cell(flexcpValue, vsProducto.Row, 0), vsProducto.Cell(flexcpData, vsProducto.Row, 3), EstadoS.Taller, cTalDeposito.ItemData(cTalDeposito.ListIndex), Trim(tAclaracion.Text), Usuario:=tUsuario.Tag)
        If vsMotivos.Rows > 1 Then InsertoMotivos IdServicio
        
        'Si ingresa directo al local inserto la tabla taller.
        If cTalDeposito.ItemData(cTalDeposito.ListIndex) = paCodigoDeSucursal Then InsertoServicioTaller IdServicio, tUsuario.Tag
        If gCliente = paClienteEmpresa Then HagoCambioDeEstado vsProducto.Cell(flexcpData, vsProducto.Row, 0), paEstadoARecuperar, IdServicio
        
        cBase.CommitTrans
        'Imprimo fichas.
        ImprimoFichaTaller IdServicio
        Foco tCi
    Else
        IdServicio = RsAux!SerCodigo
        RsAux.Close
        cBase.RollbackTrans
        MsgBox "El producto ya tiene un servicio abierto, el código de servicio es el " & IdServicio, vbInformation, "ATENCIÓN"
    End If
    LimpioTodo
    Screen.MousePointer = 0
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar la transacción.", Trim(Err.Description)
    Screen.MousePointer = 0
    Exit Sub
ErrResumir:
    Resume ErrRB
ErrRB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la información de taller.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
Private Sub ImprimoFichaTaller(IdServicio As Long)
Dim aTexto As String
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    'Seteo por defecto la impresora
    SeteoImpresoraPorDefecto paIConformeN
    
    With vsFicha
        .PaperSize = 1
'        .PaperHeight = .PaperHeight / 2
'        .Orientation = orLandscape
        
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
        .filename = "Ficha de Ingreso a Taller"
        .FontSize = 10
        .TableBorder = tbNone
        
        .TextAlign = taRightBaseline
        .FontBold = True
        .AddTable ">2000|<1500", "Servicio Número:|" & IdServicio, ""
        If paClienteEmpresa = gCliente Then
            .Paragraph = "": .AddTable ">2000|<1500", "|STOCK", ""
        Else
            .Paragraph = "": .Paragraph = "":
        End If
        .FontBold = False
        .TextAlign = taLeftBaseline
        .FontSize = 8.25
        .Paragraph = "": .Paragraph = "": .Paragraph = ""
        .AddTable "<900|<1800|>1400|<1000", "Fecha:|" & Format(gFechaServidor, "d-Mmm yyyy hh:mm") & "|Recibido por:|" & tUsuario.Text, ""
        
        .Paragraph = ""
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
        aTexto = aTexto & lTitular.Caption & "|Teléfono:|" & Trim(tTelCliente.Text)
            
        .AddTable "<900|<4500|<950|4600", aTexto, ""
        .AddTable "<900|<9000", "Dirección:|" & Trim(tDirCliente.Text), ""
        
        .Paragraph = ""
        .FontBold = True
        aTexto = "(" & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 0)) & ") " & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 1))
        .AddTable "<900|<8000", "Artículo:|" & aTexto, ""
        .FontBold = False
        
        .AddTable "<900|<1500|<1500|<1100|<1200|<1800|<900|<500", "Factura:|" & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 6)) & "|Fecha Compra:|" & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 3)) & "|Nro. Serie:|" & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 5)) & "|Estado:|" & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 2)), ""
        
        .Paragraph = ""
        .AddTable "<900|3000", "Local:|" & Trim(cTalDeposito.Text), ""
        
        .Paragraph = ""
        aTexto = ""
        For I = 1 To vsMotivos.Rows - 1
            If aTexto = "" Then aTexto = Trim(vsMotivos.Cell(flexcpText, I, 1)) Else aTexto = aTexto & ", " & Trim(vsMotivos.Cell(flexcpText, I, 1))
        Next I
        .AddTable "<900|<10100", "Motivos:|" & aTexto, ""
        If Trim(tAclaracion.Text) <> "" Then .AddTable "<1000|<10000", "Aclaración:|" & Trim(tAclaracion.Text), ""
        .Paragraph = "": .Paragraph = "": .Paragraph = "": .Paragraph = "": .Paragraph = ""
        .FontSize = 7
        aTexto = "1) - Para retirar el aparato es indispensable presentar esta boleta. -"
        .AddTable "900|10100", "Nota:|" & aTexto, ""
        aTexto = "2) - El plazo de retiro del aparato es de 90 días contados a partir de la fecha de esta boleta. Expirado dicho plazo se perderá todo derecho a reclamo " _
            & "sobre el mismo. -"
        .AddTable "900|10100", "|" & aTexto, ""
        .Paragraph = ""
        .FontSize = 9.25
        .Paragraph = "Vía Cliente"
'        .EndDoc
'        If .Error Then
'            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
'            Screen.MousePointer = 0: Exit Sub
'        End If
'        .PrintDoc   'Cliente
'        If .Error Then
'            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
'            Screen.MousePointer = 0: Exit Sub
'        End If
        
'        .StartDoc
'        If .Error Then
'            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
'            Screen.MousePointer = 0: Exit Sub
'        End If
'        .filename = "Ficha de Ingreso a Taller"
        .Paragraph = "": .Paragraph = ""
        .Paragraph = "----------------------------------------------------------------------------------------------------------------"
        .Paragraph = "": .Paragraph = ""
        .FontSize = 10
        .TableBorder = tbNone
        
        .TextAlign = taRightBaseline
        .FontBold = True
        .AddTable ">2000|<1500", "Servicio Número:|" & IdServicio, ""
        If paClienteEmpresa = gCliente Then
            .Paragraph = "": .AddTable ">2000|<1500", "|STOCK", ""
        Else
            .Paragraph = "": .Paragraph = "":
        End If
        .FontBold = False
        .TextAlign = taLeftBaseline
        .FontSize = 8.25
        .Paragraph = "": .Paragraph = "": .Paragraph = ""
        .AddTable "<900|<1800|>1400|<1000", "Fecha:|" & Format(gFechaServidor, "d-Mmm yyyy hh:mm") & "|Recibido por:|" & tUsuario.Text, ""
        
        .Paragraph = ""
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
        aTexto = aTexto & lTitular.Caption & "|Teléfono:|" & Trim(tTelCliente.Text)
            
        .AddTable "<900|<4500|<950|4600", aTexto, ""
        .AddTable "<900|<9000", "Dirección:|" & Trim(tDirCliente.Text), ""
        
        .Paragraph = ""
        .FontBold = True
        aTexto = "(" & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 0)) & ") " & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 1))
        .AddTable "<900|<8000", "Artículo:|" & aTexto, ""
        .FontBold = False
        
        .AddTable "<900|<1500|<1500|<1100|<1200|<1800|<900|<500", "Factura:|" & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 6)) & "|Fecha Compra:|" & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 3)) & "|Nro. Serie:|" & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 5)) & "|Estado:|" & Trim(vsProducto.Cell(flexcpText, vsProducto.Row, 2)), ""
        
        .Paragraph = ""
        .AddTable "<900|3000", "Local:|" & Trim(cTalDeposito.Text), ""
        
        .Paragraph = ""
        aTexto = ""
        For I = 1 To vsMotivos.Rows - 1
            If aTexto = "" Then aTexto = Trim(vsMotivos.Cell(flexcpText, I, 1)) Else aTexto = aTexto & ", " & Trim(vsMotivos.Cell(flexcpText, I, 1))
        Next I
        .AddTable "<900|<10100", "Motivos:|" & aTexto, ""
        If Trim(tAclaracion.Text) <> "" Then .AddTable "<1000|<10000", "Aclaración:|" & Trim(tAclaracion.Text), ""
        .Paragraph = "": .Paragraph = "": .Paragraph = "": .Paragraph = "": .Paragraph = ""
        .Paragraph = "Recibido:"
        .Paragraph = ""
        .Paragraph = "Reparado:"
        .Paragraph = ""
        .FontSize = 9.25
        .Paragraph = "Vía Archivo"
        .EndDoc
'        .PaperBin = paIConformeB
'        .Device = paIConformeN
        .PrintDoc   'Archivo
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub
Private Sub GraboTelefono()
    Cons = "Select * From Telefono Where TelCliente = " & gCliente _
        & " And TelTipo = " & cTipoTelefono.ItemData(cTipoTelefono.ListIndex)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.AddNew
        RsAux!TelCliente = gCliente
        RsAux!TelTipo = cTipoTelefono.ItemData(cTipoTelefono.ListIndex)
        RsAux!TelNumero = tTelefono.Text
        If Trim(tInterno.Text) <> "" Then RsAux!TelInterno = tInterno.Text
        RsAux.Update
    Else
        RsAux.Edit
        RsAux!TelNumero = tTelefono.Text
        If Trim(tInterno.Text) <> "" Then RsAux!TelInterno = tInterno.Text Else RsAux!TelInterno = Null
        RsAux.Update
    End If
    RsAux.Close
    tTelCliente.Text = TelefonoATexto(gCliente)     'Telefonos
End Sub
Private Function ValidoIngreso() As Boolean
    
    ValidoIngreso = False
    
    If Val(tUsuario.Tag) = 0 Then MsgBox "Ingrese su digito de usuario.", vbExclamation, "ATENCIÓN": Foco tUsuario: Exit Function
    If Not clsGeneral.TextoValido(tAclaracion.Text) Then MsgBox "Se ingresó alguna comilla simple en la aclaración, debe eliminarla.", vbExclamation, "ATENCIÓN": Foco tAclaracion: Exit Function
    
    If Trim(tTelefono.Text) <> "" Then
        If cTipoTelefono.ListIndex = -1 Then MsgBox "Debe seleccionar un tipo de teléfono.", vbExclamation, "ATENCIÓN": Foco cTipoTelefono: Exit Function
        Cons = "Select * From Producto Where ProCodigo = " & vsProducto.Cell(flexcpValue, vsProducto.Row, 0)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        If Not IsNull(RsAux!ProDireccion) Then tTelefono.Tag = RsAux!ProDireccion Else tTelefono.Tag = 0
        RsAux.Close
        tTelefono.Tag = clsGeneral.RetornoFormatoTelefono(tTelefono.Text, Val(tTelefono.Tag), txtConexion)
        If tTelefono.Tag = "" Then MsgBox "El formato del teléfono no es válido.", vbExclamation, "ATENCIÓN": Foco tTelefono: Exit Function
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

Private Sub InsertoMotivoEnGrilla()
    'Tengo el RSAUX con el motivo
    Dim aValor As Long
    'Verifico que no este insertado.
    With vsMotivos
        For I = 1 To .Rows - 1
            If Val(.Cell(flexcpData, I, 0)) = RsAux!MSeID Then MsgBox "El motivo ya fue ingresado, verifique.", vbInformation, "ATENCIÓN": Exit Sub
        Next I
        .AddItem ""
        aValor = RsAux!MSeID
        .Cell(flexcpData, .Rows - 1, 0) = aValor
        .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!MSeCodigo, "#,000")
        .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!MSeNombre)
    End With
End Sub

Private Function InsertoServicio(idProducto As Long, EstadoProducto As Integer, EstadoServicio As Integer, LocalReparacion As Long, Optional Comentario As String = "", Optional LocalRecepcion As Long = -1, Optional Usuario As Long = -1) As Long
    
    If LocalRecepcion = -1 Then LocalRecepcion = paCodigoDeSucursal
    If Usuario = -1 Then Usuario = paCodigoDeUsuario
    
    '---------------------------------------------
    'Inserto
    Cons = "INSERT INTO Servicio (SerProducto, SerFecha, SerEstadoProducto, SerLocalIngreso, " _
        & " SerLocalReparacion, SerEstadoServicio, SerUsuario, SerModificacion, SerComentario) Values (" _
        & idProducto & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', " & EstadoProducto & ", " & LocalRecepcion
    
    If LocalReparacion = 0 Then Cons = Cons & ", Null " Else Cons = Cons & ", " & LocalReparacion
    
    Cons = Cons & ", " & EstadoServicio & ", " & Usuario & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', "
    If Comentario = "" Then Cons = Cons & "Null)" Else Cons = Cons & "'" & Comentario & "')"
    cBase.Execute (Cons)
    
    '---------------------------------------------
    'Saco el mayor código de servicio.
    Cons = "Select Max(SerCodigo) From Servicio"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    InsertoServicio = RsAux(0)
    RsAux.Close
    '---------------------------------------------
    
End Function

Private Sub InsertoMotivos(IdServicio As Long)
    With vsMotivos
        For I = 1 To .Rows - 1
            Cons = "Insert Into ServicioRenglon (SReServicio, SReTipoRenglon,  " _
                & " SReMotivo, SReCantidad) Values (" & IdServicio & ", " & TipoRenglonS.Llamado & ",  " & Val(.Cell(flexcpData, I, 0)) & ", 1)"
            cBase.Execute (Cons)
        Next I
    End With
End Sub

Private Sub InsertoServicioTaller(IdServicio As Long, Optional Usuario As Integer = -1)

    If Usuario = -1 Then Usuario = paCodigoDeUsuario
    If cTalDeposito.ItemData(cTalDeposito.ListIndex) <> paCodigoDeSucursal Then
        'Inserto también el local para el traslado.
        Cons = "Insert Into Taller(TalServicio, TalFIngresoRealizado, TalFIngresoRecepcion, TalModificacion, TalUsuario, TalLiquidadoATecnico, TalLocalAlCliente) Values (" _
            & IdServicio & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
            & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', " & Usuario & ", 0, " & cTalDeposito.ItemData(cTalDeposito.ListIndex) & ")"
    Else
        Cons = "Insert Into Taller(TalServicio, TalFIngresoRealizado, TalFIngresoRecepcion, TalModificacion, TalUsuario, TalLiquidadoATecnico) Values (" _
            & IdServicio & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
            & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', " & Usuario & ", 0)"
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
    Cons = "Select * From TipoFlete Where TFlCodigo = " & IDTFlete
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
        
        If IsNull(RsAux!TFlAgenda) Then
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
    clsGeneral.OcurrioError "Ocurrio un error al cargar los datos para el tipo de flete.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub BuscoValorFlete()

    If cRMoneda.ListIndex = -1 Or cTipoFlete.ListIndex = -1 Then Exit Sub
    Cons = "Select * From ValorFlete Where VFlTipoFlete = " & cTipoFlete.ItemData(cTipoFlete.ListIndex)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!VFlValor) Then tRImporte.Text = Format(RsAux!VFlValor, FormatoMonedaP) Else tRImporte.Text = ""
        If cRFactura.ListIndex > -1 Then
            If Not IsNull(RsAux!VFlCosto) And cRFactura.ItemData(cRFactura.ListIndex) <> FacturaServicio.Camion Then tRLiquidar.Text = Format(RsAux!VFlCosto, FormatoMonedaP) Else tRLiquidar.Text = "0.00"
        End If
    End If
    RsAux.Close

End Sub

Private Sub AccionGrabarRetiro()
On Error GoTo ErrBR
Dim IdServicio As Long, strHora As String
    
    IdServicio = 0
    Screen.MousePointer = 11
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo ErrResumir
    IdServicio = TieneReporteAbierto(vsProducto.Cell(flexcpValue, vsProducto.Row, 0))
    If IdServicio = 0 Then
        If Trim(tTelefono.Text) <> "" Then GraboTelefono
        IdServicio = InsertoServicio(vsProducto.Cell(flexcpValue, vsProducto.Row, 0), CInt(vsProducto.Cell(flexcpData, vsProducto.Row, 3)), EstadoS.Retiro, cRDeposito.ItemData(cRDeposito.ListIndex), Trim(tAclaracion.Text), Usuario:=tUsuario.Tag)
        If vsMotivos.Rows > 1 Then InsertoMotivos IdServicio
        If cRHora.ListIndex > -1 Then
            strHora = RetornoHoraEnString(cRHora.ItemData(cRHora.ListIndex))
        Else
            strHora = Trim(cRHora.Text)
        End If
        InsertoServicioVisita IdServicio, TipoServicio.Retiro, cRCamion.ItemData(cRCamion.ListIndex), tRFecha.Tag, strHora, vsProducto.Cell(flexcpData, vsProducto.Row, 2), cRMoneda.ItemData(cRMoneda.ListIndex), _
            tRImporte.Text, cRFactura.ItemData(cRFactura.ListIndex), tRLiquidar.Text, cTipoFlete.ItemData(cTipoFlete.ListIndex)
        cBase.CommitTrans
        Foco tCi
    Else
        cBase.RollbackTrans
        MsgBox "El producto ya tiene un servicio abierto, el código de servicio es el " & IdServicio, vbInformation, "ATENCIÓN"
    End If
    LimpioTodo
    Screen.MousePointer = 0: Exit Sub
ErrBR:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar la transacción.", Trim(Err.Description)
    Screen.MousePointer = 0: Exit Sub
ErrResumir:
    Resume ErrRB
ErrRB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la información del retiro.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
Private Sub InsertoServicioVisita(IdServicio As Long, TipoServ As Integer, IDCamion As Long, fecha As String, Hora As String, Zona As Long, _
                                                        Moneda As Integer, Importe As Currency, FormaPago As Integer, LiquidarC As Currency, Optional TipoFlete As Integer = -1, Optional SinEfecto As Boolean = False, Optional TextoVisita As Integer = 0)

    Cons = "Select * From ServicioVisita Where VisServicio = " & IdServicio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    RsAux.AddNew
    RsAux!VisServicio = IdServicio
    RsAux!VisTipo = TipoServ
    RsAux!VisCamion = IDCamion
    RsAux!VisFecha = Format(fecha, sqlFormatoF)
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
    clsGeneral.OcurrioError "Ocurrio un error al verificar si existe algún servicio abierto.", Trim(Err.Description)
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

Private Sub AccionGrabarVisita()
On Error GoTo ErrBV
Dim IdServicio As Long, strHora As String
    IdServicio = 0
    Screen.MousePointer = 11
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo ErrResumir
    IdServicio = TieneReporteAbierto(vsProducto.Cell(flexcpValue, vsProducto.Row, 0))
    If IdServicio = 0 Then
        If Trim(tTelefono.Text) <> "" Then GraboTelefono
        IdServicio = InsertoServicio(vsProducto.Cell(flexcpValue, vsProducto.Row, 0), CInt(vsProducto.Cell(flexcpData, vsProducto.Row, 3)), EstadoS.Visita, 0, Trim(tAclaracion.Text), Usuario:=tUsuario.Tag)
        If vsMotivos.Rows > 1 Then InsertoMotivos IdServicio
        If cVHora.ListIndex > -1 Then
            strHora = RetornoHoraEnString(cVHora.ItemData(cVHora.ListIndex))
        Else
            strHora = Trim(cVHora.Text)
        End If
        Dim aValor As Integer
        If cVComentario.ListIndex > -1 Then aValor = cVComentario.ItemData(cVComentario.ListIndex) Else aValor = 0
          InsertoServicioVisita IdServicio, TipoServicio.Visita, cVCamion.ItemData(cVCamion.ListIndex), tVFecha.Tag, strHora, vsProducto.Cell(flexcpData, vsProducto.Row, 2), cVMoneda.ItemData(cVMoneda.ListIndex), _
            tVImporte.Text, cVFactura.ItemData(cVFactura.ListIndex), tVLiquidar.Text, , , aValor
        cBase.CommitTrans
        Foco tCi
    Else
        cBase.RollbackTrans
        MsgBox "El producto ya tiene un servicio abierto, el código de servicio es el " & IdServicio, vbInformation, "ATENCIÓN"
    End If
    LimpioTodo
    Screen.MousePointer = 0
    Exit Sub
ErrBV:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar la transacción.", Trim(Err.Description)
    Screen.MousePointer = 0
    Exit Sub
ErrResumir:
    Resume ErrRB
ErrRB:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la información de la visita.", Trim(Err.Description)
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

Private Sub HagoCambioDeEstado(IDArticulo As Long, EstadoNuevo As Integer, IdServicio As Long)
    'Cambio el estado del artículo como Sano a Recuperar.
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, EstadoNuevo, 1, TipoDocumento.ServicioCambioEstado, IdServicio
    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoArticuloEntrega, -1, TipoDocumento.ServicioCambioEstado, IdServicio
        
    MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, EstadoNuevo, 1, 1
    MarcoMovimientoStockTotal IDArticulo, TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, 1, -1
    
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, EstadoNuevo, 1
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoArticuloEntrega, -1
    
End Sub

Private Sub CargoHistoria(idProducto As Long)
Dim IdServicio As Long: IdServicio = 0
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
            .AddItem ""
            If RsAux!SerEstadoServicio = EstadoS.Anulado Then .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!SerFCumplido, "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, 2) = EstadoProducto(RsAux!SerEstadoProducto)
            .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!SerFecha, "dd/mm/yy")
            If Not IsNull(RsAux!SerMoneda) And Not IsNull(RsAux!SerCostoFinal) Then .Cell(flexcpText, .Rows - 1, 4) = BuscoSignoMoneda(RsAux!SerMoneda) & " " & Format(RsAux!SerCostoFinal, FormatoMonedaP)
            If Not IsNull(RsAux!SerComentarioR) Then aComentario = Trim(RsAux!SerComentarioR) Else aComentario = ""
            IdServicio = RsAux!SerCodigo
            Do While Not RsAux.EOF
                If Not IsNull(RsAux!ArtNombre) Then
                    If Trim(.Cell(flexcpText, .Rows - 1, 1)) <> "" Then .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1) & ", "
                    .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1) & Trim(RsAux!ArtNombre)
                End If
                IdServicio = RsAux!SerCodigo
                RsAux.MoveNext
                If RsAux.EOF Then Exit Do
                If IdServicio <> RsAux!SerCodigo Then Exit Do
            Loop
            If Trim(.Cell(flexcpText, .Rows - 1, 1)) <> "" Then .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1) & Chr(13)
            .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 1) & aComentario
        End With
    Loop
    RsAux.Close
    With vsHistoria
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 1, , False
    End With
    Screen.MousePointer = 0
    Exit Sub
ErrCH:
    clsGeneral.OcurrioError "Ocurrio un error al buscar la historia del producto.", Err.Description
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
    
    Cons = "Select * From CamionFlete, TipoFlete Where CTFCamion = " & IDCamion _
        & " And CTFTipoFlete = TFlCodigo"
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
                tRFecha.Text = Format(CDate(tVFecha.Tag), "ddd d/mm/yy")
                CargoHoraEntregaParaDia cVHora, tVFecha.Tag
            End If
        End If
        
        Dim IdTipoFlete As Integer
        IdTipoFlete = RsAux!TFlCodigo
        RsAux.Close
        Cons = "Select * From ValorFlete Where VFlTipoFlete = " & IdTipoFlete
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then
            If cVMoneda.ListIndex = -1 Then BuscoCodigoEnCombo cVMoneda, paMonedaPesos
            If Not IsNull(RsAux!VFlValor) Then tVImporte.Text = Format(RsAux!VFlValor, FormatoMonedaP) Else tVImporte.Text = ""
            If cVFactura.ListIndex > -1 Then
                If Not IsNull(RsAux!VFlCosto) And cVFactura.ItemData(cVFactura.ListIndex) <> FacturaServicio.Camion Then tVLiquidar.Text = Format(RsAux!VFlCosto, FormatoMonedaP) Else tVLiquidar.Text = "0.00"
            End If
        End If
        RsAux.Close
    End If
    
End Sub

Private Sub CargoTextoVisita()
On Error GoTo ErrCC
    Cons = "Select TViCodigo, TViTexto From TextoVisita Order By TViTexto"
    CargoCombo Cons, cVComentario
    Exit Sub
ErrCC:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los textos de visita."
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
    
    With vsFicha
        .PageBorder = pbNone
        .Device = paIConformeN
        If .Device <> paIConformeN Then MsgBox "Ud no tiene instalada la impresora para imprimir Conformes. Avise al administrador.", vbExclamation, "ATENCIÒN"
        If .PaperBins(paIConformeB) Then .PaperBin = paIConformeB Else MsgBox "Esta mal definida la bandeja de conformes en su sucursal, comuniquele al administrador.", vbInformation, "ATENCIÓN": paIConformeB = .PaperBin
        .PaperSize = 256 'Hoja carta
        .Orientation = orPortrait
       ' .PaperHeight = .PaperHeight / 2
        .MarginTop = 300
        .MarginLeft = 500
    End With

    Exit Sub
ErrPBI:
    clsGeneral.OcurrioError "Ocurrio un error al setear la impresora, consulte con el administrador de impresión este problema.", Err.Description
End Sub

Private Function BuscoPrimerDiaAEnviar() As String
Dim strMat As String

Dim intSumaDia As Integer

    BuscoPrimerDiaAEnviar = ""
    
    strMat = ""
    
    If IsNull(RsAux!TFlFechaAgeHab) Then strCierre = Date Else strCierre = RsAux!TFlFechaAgeHab
    douAgenda = RsAux!TFlAgenda
    If IsNull(RsAux!TFlAgendaHabilitada) Then douHabilitado = -1 Else douHabilitado = RsAux!TFlAgendaHabilitada
    
    If IsDate(strCierre) Then
        If Abs(DateDiff("d", strCierre, Date)) >= 7 Then
            'Como cerro hace una semana tomo la agenda normal.
            strMat = superp_MatrizSuperposicion(douAgenda)
        Else
            If douHabilitado > 0 Then
                strMat = superp_MatrizSuperposicion(douHabilitado)
            Else
                strMat = superp_MatrizSuperposicion(douAgenda)
            End If
        End If
    Else
        strMat = superp_MatrizSuperposicion(douAgenda)
    End If
    
    If strMat <> "" Then
        intSumaDia = BuscoProximoDia(Date, strMat, douAgenda)
        If intSumaDia <> -1 Then BuscoPrimerDiaAEnviar = Format(Date + intSumaDia, "d-mm-yy")
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
            intDia = WeekDay(CDate(strFecha) + intSuma)
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
            intDia = WeekDay(CDate(strFecha) + intSuma)
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
                & " And HFlDiaSemana = " & WeekDay(CDate(strFecha)) & " Order by HFlInicio"
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



