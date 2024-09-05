VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form frmInFactura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compra de Mercadería"
   ClientHeight    =   6375
   ClientLeft      =   4335
   ClientTop       =   2355
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInFactura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8625
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   300
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.TextBox tDtoRecibo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   1080
      MaxLength       =   14
      TabIndex        =   22
      Top             =   5100
      Width           =   975
   End
   Begin VB.TextBox tAutoriza 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MaxLength       =   80
      TabIndex        =   26
      Top             =   5760
      Width           =   2055
   End
   Begin VB.CheckBox chVerificado 
      Appearance      =   0  'Flat
      Caption         =   "Autoriza&do"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3540
      TabIndex        =   27
      Top             =   5805
      Width           =   1395
   End
   Begin VB.TextBox tCofis 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   5640
      MaxLength       =   14
      TabIndex        =   42
      Text            =   "0.00"
      Top             =   4590
      Width           =   1455
   End
   Begin VB.TextBox tIva 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   5640
      MaxLength       =   14
      TabIndex        =   20
      Text            =   "0.00"
      Top             =   4860
      Width           =   1455
   End
   Begin VB.TextBox tNeto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   5640
      MaxLength       =   14
      TabIndex        =   19
      Text            =   "0.00"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalle de Factura"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   35
      Top             =   480
      Width           =   8415
      Begin VB.TextBox tImporte 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         MaxLength       =   7
         TabIndex        =   47
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox tCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   600
         Width           =   4455
      End
      Begin AACombo99.AACombo cTipo 
         Height          =   315
         Left            =   4800
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
      End
      Begin VB.CommandButton bMercaderia 
         Caption         =   "I&ngreso de Mercadería"
         Height          =   315
         Left            =   5580
         TabIndex        =   9
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   3
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox tFactura 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   6900
         MaxLength       =   7
         TabIndex        =   6
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox tArbitraje 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox tDescuento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   7080
         MaxLength       =   14
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
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
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8100
         TabIndex        =   40
         Top             =   990
         Width           =   255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "P&roveedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lFactura 
         BackStyle       =   0  'Transparent
         Caption         =   "&Documento:"
         Height          =   255
         Left            =   3840
         TabIndex        =   4
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento &General:"
         Height          =   255
         Left            =   5595
         TabIndex        =   14
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Imp. Total:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "&Arbitraje:"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   990
         Width           =   735
      End
   End
   Begin VB.CheckBox cIva 
      BackColor       =   &H00808080&
      Caption         =   "&Precios sin I.V.A."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   16
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VSFlex6DAOCtl.vsFlexGrid FGrid 
      Height          =   2175
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3836
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
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   7800
      MaxLength       =   2
      TabIndex        =   29
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MaxLength       =   80
      TabIndex        =   24
      Top             =   5460
      Width           =   7455
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   6120
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "sucursal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4895
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "&Dto. Recibo:"
      Height          =   255
      Left            =   60
      TabIndex        =   21
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2085
      TabIndex        =   48
      Top             =   5130
      Width           =   255
   End
   Begin VB.Label lAutoriza 
      BackStyle       =   0  'Transparent
      Caption         =   "Autori&za:"
      Height          =   255
      Left            =   60
      TabIndex        =   25
      Top             =   5805
      Width           =   975
   End
   Begin VB.Label lTTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Documento:"
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
      Left            =   7080
      TabIndex        =   31
      Top             =   5130
      Width           =   1455
   End
   Begin VB.Label lNetoCofis 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   7080
      TabIndex        =   45
      Top             =   4600
      Width           =   1395
   End
   Begin VB.Label lIva 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   7080
      TabIndex        =   44
      Top             =   4860
      Width           =   1395
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&fis:"
      Height          =   255
      Left            =   5100
      TabIndex        =   43
      Top             =   4590
      Width           =   615
   End
   Begin VB.Label lTDescuento 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Proveedor:"
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
      Height          =   255
      Left            =   2880
      TabIndex        =   39
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lTSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Fecha:"
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
      Height          =   255
      Left            =   2880
      TabIndex        =   38
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Descuento:"
      Height          =   255
      Left            =   1920
      TabIndex        =   37
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "SubTotal:"
      Height          =   255
      Left            =   1920
      TabIndex        =   36
      Top             =   4320
      Width           =   975
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":0A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInFactura.frx":0D86
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   ARTÍCU&LOS DE LA FACTURA"
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   120
      TabIndex        =   17
      Top             =   1905
      Width           =   8415
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "&Neto:"
      Height          =   255
      Left            =   5100
      TabIndex        =   34
      Top             =   4320
      Width           =   555
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&I.V.A.:"
      Height          =   255
      Left            =   5100
      TabIndex        =   33
      Top             =   4860
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Total de la Factura........:"
      Height          =   255
      Left            =   5100
      TabIndex        =   32
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   255
      Left            =   7140
      TabIndex        =   28
      Top             =   5820
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Comentarios:"
      Height          =   255
      Left            =   60
      TabIndex        =   23
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
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
      Height          =   885
      Left            =   7080
      TabIndex        =   46
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuOpL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario "
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmInFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------
'   --> Los remitos que el Local <> Compañia: La cantidad en compañia es la que queda por
'         sacar del remito (la que no está asignada a las facturas).
'
'   --> Los remitos que el Local = Compañia: La cantidad en compañia es la que queda para
'         pasar a los locales.
'
'   En este formulario solo se manejan los remitos con local <> de compañia
'----------------------------------------------------------------------------------------------------------

Option Explicit
Public prmIDCompra As Long

Dim I As Long, gCompra As Long
Dim sNuevo As Boolean, sModificar As Boolean
Dim Msg As String, aTexto As String

Private Sub bMercaderia_Click()

    If Val(tProveedor.Tag) = 0 Then
        MsgBox "Seleccione el proveedor de mercadería para proceder al ingreso.", vbExclamation, "ATENCIÓN"
        Foco tProveedor: Exit Sub
    End If
    
    If FGrid.Rows > 1 Then
        If MsgBox("Hay artículos ingresados, si va al ingreso de mercadería se perderán los costos ya actualizados.", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    End If
    
    On Error GoTo errCargar
    Screen.MousePointer = 11
    
    EjecutarApp App.Path & "\Ingreso de Mercaderia.exe ", CStr(tProveedor.Tag)
    Foco tProveedor
    Exit Sub
    
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al procesar la información.", Err.Description
End Sub

Private Sub cIva_Click()

    If cIva.Value = vbChecked Then
        FGrid.ColHidden(5) = False
        FGrid.ColHidden(6) = True
    Else
        FGrid.ColHidden(5) = True
        FGrid.ColHidden(6) = False
    End If
        
End Sub

Private Sub Form_DblClick()

    For I = 1 To FGrid.Cols - 1
        If FGrid.ColHidden(I) Then FGrid.ColHidden(I) = False
    Next
End Sub

Private Sub Label6_Click()
    Foco tIva
End Sub

Private Sub Label9_Click()
    Foco tNeto
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub tAutoriza_Change()
    tAutoriza.Tag = 0
End Sub

Private Sub tCodigo_Change()
    gCompra = 0
End Sub

Private Sub tCodigo_GotFocus()
    tCodigo.SelStart = 0: tCodigo.SelLength = Len(tCodigo.Text)
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCodigo.Text) Then CargoCompra Val(tCodigo.Text)
    End If
End Sub

Private Sub tCofis_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        On Error GoTo errR
        If IsNumeric(tCofis.Text) Then
        
            tCofis.Text = Format(tCofis.Text, FormatoMonedaP)
            lNetoCofis.Caption = Format(CCur(tNeto.Text) + CCur(tCofis.Text), FormatoMonedaP)
            lTTotal.Caption = Format(CCur(lNetoCofis.Caption) + CCur(lIva.Caption), FormatoMonedaP)
            
            Foco tIva
        End If
    End If
    Exit Sub
    
errR:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al procesar los importes.", Err.Description
End Sub

Private Sub tDtoRecibo_GotFocus()
    tDtoRecibo.SelStart = 0
    tDtoRecibo.SelLength = Len(tDtoRecibo.Text)
    Status.SimpleText = " Ingrese el porcentaje de descuento del Recibo."
End Sub

Private Sub tDtoRecibo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        On Error Resume Next
        tDtoRecibo.Text = CalculoPorcentajeDescuento(Trim(tDtoRecibo.Text))
        If tComentario.Enabled Then tComentario.SetFocus Else AccionGrabar
    End If
    
End Sub

Private Sub tFecha_Change()
    If Not sNuevo And Not sModificar Then gCompra = 0
End Sub

Private Sub tFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And Not sNuevo And Not sModificar Then BuscoCompras
End Sub

Private Sub tIva_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        On Error GoTo errR
        If IsNumeric(tIva.Text) Then
            tIva.Text = Format(tIva.Text, FormatoMonedaP)
            lIva.Caption = tIva.Text
            lTTotal.Caption = Format(CCur(lNetoCofis.Caption) + CCur(tIva.Text), FormatoMonedaP)
            Foco tDtoRecibo
        End If
    End If
    
    Exit Sub
errR:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al procesar al redondeo."
End Sub

Private Sub tProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then BuscoCompras
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        If Val(tProveedor.Tag) <> 0 Or Trim(tProveedor.Text) = "" Then
            If cMoneda.Enabled Then Foco cMoneda Else Foco tFecha
            Exit Sub
        End If
        
        'Busco el proveedor         ----------------------------------------------------------------------------------------
        Screen.MousePointer = 11
        Dim aNameSel As String
        Dim aIdSel As Long, aQ As Integer
        aQ = 0
        
        aNameSel = Replace(Trim(tProveedor.Text), " ", "%")
        Cons = "Select PMeCodigo, PMeFantasia as 'Nombre', PMeNombre 'Razón Social' " & _
                   " from ProveedorMercaderia " & _
                   " Where PMeNombre like '" & aNameSel & "%' Or PMeFantasia like '" & aNameSel & "%'"
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aIdSel = RsAux!PMeCodigo: aNameSel = Trim(RsAux(1))
            aQ = 1
            RsAux.MoveNext
            If Not RsAux.EOF Then
                aQ = 2: aIdSel = 0
            End If
        End If
        RsAux.Close
        '-------------------------------------------------------------------------------------------------------------------
        
        Select Case aQ
            Case 0: MsgBox "No hay datos para el texto ingresado.", vbExclamation, "No Hay Datos"
            
            Case 2:
                        Dim aLista As New clsListadeAyuda
                        aIdSel = aLista.ActivarAyuda(cBase, Cons, 5500, 1, "Lista de Proveedores")
                        If aIdSel <> 0 Then
                            aNameSel = aLista.RetornoDatoSeleccionado(1)
                            aIdSel = aLista.RetornoDatoSeleccionado(0)
                        End If
                        Set aLista = Nothing
                        Me.Refresh
        End Select
        
        If aIdSel <> 0 Then
            tProveedor.Text = aNameSel
            tProveedor.Tag = aIdSel
        End If
        Screen.MousePointer = 0
    End If
    
End Sub


Private Sub FGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    '1- Codigo
    '2- Articulo
    '3- Cantidad c/Cargo
    '4- Cantidad sin Cargo
    '5- Unitario Sin IVA
    '6- Unitario Con IVA
    '7- % Descuento
    '8- SubTotal (sin IVA)
    '9- % IVA del artículo
    '10- Total Subtotal + Iva
    '11- Numero de Remito (si ya está en algun local)
    '12- ID interno de Remito (si ya está en algun local)
    '13- Q en remito
    If Row = 0 Then Exit Sub
    Select Case Col
        Case 3, 5, 6, 7
            Dim miUnitario As Currency
            'Si me Ingresaron un precio ---> Ajusto el otro-------------------------------------------------------------
            If Col = 6 And FGrid.TextMatrix(Row, 6) <> "" Then     'Unitario CON IVA ---> Pongo el Unitario sin IVA
                ' Ingresan el unitario con IVA --> es con cofis
                miUnitario = (FGrid.TextMatrix(Row, 6) * 100) / (100 + FGrid.TextMatrix(Row, 9))
                FGrid.Cell(flexcpText, Row, 5) = Format(miUnitario, "#,##0.00")
            End If
            
            If Col = 5 And FGrid.TextMatrix(Row, 5) <> "" Then     'Unitario Sin IVA ---> Pongo el Unitario Con IVA
                miUnitario = FGrid.TextMatrix(Row, 5)
                FGrid.TextMatrix(Row, 6) = (miUnitario + ((miUnitario * Val(FGrid.TextMatrix(Row, 9))) / 100))   'Unitario sin IVA
            End If
            '------------------------------------------------------------------------------------------------------------------
            
            'Si Hay Cantidad y Precio
            If FGrid.TextMatrix(Row, 5) <> "" And FGrid.TextMatrix(Row, 3) <> "" Then
                'Calculo Total del Renglón
                
                'Saco el Precio Neto Total sin IVA  y sin Cofis
                Dim aNeto As Currency
                aNeto = FGrid.TextMatrix(Row, 5)
                aNeto = FGrid.TextMatrix(Row, 3) * aNeto      'Cant. P.Unitario
                               
                'Veo si Hay otros artículos con el mismo codigo para igualar precios y descuentos----------------
                For I = 1 To FGrid.Rows - 1
                    If FGrid.TextMatrix(I, 1) = FGrid.TextMatrix(Row, 1) Then
                        FGrid.TextMatrix(I, 5) = FGrid.ValueMatrix(Row, 5)  'Precio sIVa
                        FGrid.TextMatrix(I, 6) = FGrid.ValueMatrix(Row, 6)  'Precio cIva
                        FGrid.TextMatrix(I, 7) = FGrid.ValueMatrix(Row, 7)  '% Descuento
                        
                        aNeto = FGrid.TextMatrix(I, 5)
                        aNeto = FGrid.TextMatrix(I, 3) * aNeto      'Cant. P.Unitario
                        
                        FGrid.TextMatrix(I, 8) = aNeto - ((aNeto * FGrid.TextMatrix(I, 7)) / 100)       'Subtotal (sin IVA sin Cofis)
                                   
                                            
                        '10- Total IVA incluido
                        aNeto = FGrid.TextMatrix(I, 8)
                        FGrid.TextMatrix(I, 10) = aNeto + ((aNeto * Val(FGrid.TextMatrix(I, 9)) / 100))        'Agrego IVa
                        'FGrid.TextMatrix(I, 10) = FGrid.TextMatrix(I, 8) + ((FGrid.TextMatrix(I, 8) * (FGrid.TextMatrix(I, 9)) / 100))
                        
                        'Si Hay descuento general, aplico el descuento general (Pero al Total,  No al Sub)
                        If IsNumeric(tDescuento.Text) Then
                            FGrid.TextMatrix(I, 10) = FGrid.ValueMatrix(I, 10) - ((FGrid.ValueMatrix(I, 10) * CCur(tDescuento.Text)) / 100)
                        End If
                        
                    End If
                Next
                '------------------------------------------------------------------------------------------------------------
                CargoTotales    'Cargo las labels con los totales
            End If
            
    End Select
    
End Sub

Private Sub CargoTotales()
    
    'Cargo las labels con los totales------------------------------------------
    Dim aTotal As Currency
    Dim aNeto As Currency
    On Error Resume Next
    
    aNeto = 0: aTotal = 0
    For I = 1 To FGrid.Rows - 1
        aNeto = aNeto + FGrid.ValueMatrix(I, 8)     'Total s/Iva
        aTotal = aTotal + FGrid.ValueMatrix(I, 10)   'Total c/Iva
    Next
    aTotal = Format(aTotal, FormatoMonedaP)
    
    lTSubTotal.Caption = Format(aNeto, FormatoMonedaP)
    If IsNumeric(tDescuento.Text) Then
        lTDescuento.Caption = Format((aNeto * CCur(tDescuento.Text)) / 100, FormatoMonedaP)
    Else
        lTDescuento.Caption = "0.00"
    End If
    
    tNeto.Text = Format(aNeto - CCur(lTDescuento.Caption), FormatoMonedaP)
    tCofis.Text = Format(CCur(tNeto.Text) - CCur(tNeto.Text), FormatoMonedaP)
    lNetoCofis.Caption = Format(CCur(tNeto.Text) + CCur(tCofis.Text), FormatoMonedaP)
    
    tIva.Text = Format(aTotal - CCur(tNeto.Text) - CCur(tCofis.Text), FormatoMonedaP)
    lIva.Caption = tIva.Text
    
    lTTotal.Caption = Format(CCur(lNetoCofis.Caption) + CCur(lIva.Caption), FormatoMonedaP)
    '------------------------------------------------------------------------------
    
End Sub


Private Sub FGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not FGrid.Editable Then Cancel = True: Exit Sub
    If Row = 0 Then Cancel = True
End Sub

Private Sub FGrid_Click()
    EditoCelda
End Sub

Private Sub FGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        SiguienteCelda FGrid.Row, FGrid.Col
    Else
        EditoCelda
    End If
    
End Sub


Private Sub FGrid_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)

    If KeyCode = vbKeyReturn Then SiguienteCelda Row, Col
    
End Sub

Private Sub SiguienteCelda(aRow As Long, aCol As Long)

    On Error GoTo errMov
    Select Case aCol
        Case 3: FGrid.Select aRow, aCol + 1
        Case 4: If Not FGrid.ColHidden(5) Then FGrid.Select aRow, aCol + 1 Else: FGrid.Select aRow, aCol + 2
        Case 5: FGrid.Select aRow, aCol + 2
        Case 6: FGrid.Select aRow, aCol + 1
        Case 7: If aRow = FGrid.Rows - 1 Then Foco tNeto Else: FGrid.Select aRow + 1, 3
    End Select

errMov:
End Sub

Private Sub FGrid_LostFocus()
    On Error Resume Next
    FGrid.Select 0, 0
End Sub

Private Sub FGrid_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '1- Codigo
    '2- Articulo
    '3- Cantidad c/Cargo
    '4- Cantidad sin Cargo
    '5- Unitario Sin IVA
    '6- Unitario Con IVA
    '7- % Descuento
    '8- SubTotal (sin IVA)
    '9- % IVA del artículo
    '10- Total Subtotal + Iva
    '11- Numero de Remito (si ya está en algun local)
    '12- ID interno de Remito (si ya está en algun local)
    '13- Q original en Remito
    
    Select Case Col
        Case 3
                If Not IsNumeric(FGrid.EditText) Then FGrid.EditText = FGrid.TextMatrix(Row, Col): Exit Sub
                'Q$ + Q sc > QRemito
                If CCur(FGrid.EditText) + FGrid.ValueMatrix(Row, 4) > FGrid.ValueMatrix(Row, 13) Then
                    MsgBox "La cantidad de artículos no puede superar la ingresada en el remito.", vbExclamation, "ATENCIÓN"
                    FGrid.EditText = FGrid.TextMatrix(Row, Col)
                Else
                    FGrid.EditText = CCur(FGrid.EditText)
                End If
        
        Case 4: 'Cantidad sin cargo
                If Not IsNumeric(FGrid.EditText) Then FGrid.EditText = FGrid.TextMatrix(Row, Col): Exit Sub
                'Q$ + Q sc > QRemito
                If CCur(FGrid.EditText) + FGrid.ValueMatrix(Row, 3) > FGrid.ValueMatrix(Row, 13) Then
                    MsgBox "La cantidad de artículos no puede superar la ingresada en el remito.", vbExclamation, "ATENCIÓN"
                    FGrid.EditText = FGrid.TextMatrix(Row, Col)
                Else
                    FGrid.EditText = CCur(FGrid.EditText)
                End If
        
        Case 5, 6: If Not IsNumeric(FGrid.EditText) Then FGrid.EditText = FGrid.TextMatrix(Row, Col)
        
        Case 7: FGrid.EditText = CalculoPorcentajeDescuento(FGrid.EditText)
            
    End Select
    
End Sub

Private Function CalculoPorcentajeDescuento(Texto As String) As String

    If Trim(Texto) = "" Then CalculoPorcentajeDescuento = Texto: Exit Function
    If IsNumeric(Texto) Then CalculoPorcentajeDescuento = Texto: Exit Function
    
Dim aValor As Currency
Dim aDesc As Currency

    'Si no es numérico
    aDesc = 0
    aTexto = ""
    'Formateo el string con 2 puntos para despues sacar los numeros----------
    For I = 1 To Len(Texto)
        If Mid(Texto, I, 1) = "+" Or Mid(Texto, I, 1) = "-" Then
            aTexto = aTexto & ":" & Mid(Texto, I, 1)
        Else
            aTexto = aTexto & Mid(Texto, I, 1)
        End If
    Next
    If Mid(aTexto, 1, 1) = ":" Then aTexto = Mid(aTexto, 2, Len(aTexto))
    '--------------------------------------------------------------------------------------
    
    On Error GoTo errCalculo
    Do While aTexto <> ""
        'El texto queda 2:-3:+6.5
        If Not IsNumeric(aTexto) Then
            aValor = CCur(Mid(aTexto, 1, InStr(aTexto, ":") - 1))
            aTexto = Trim(Mid(aTexto, InStr(aTexto, ":") + 1, Len(aTexto)))
            
        Else
            aValor = CCur(aTexto)
            aTexto = ""
        End If
        
        If aDesc = 0 Then
            aDesc = 1 - (aValor / 100)
        Else
            aDesc = aDesc * (1 - (aValor / 100))
        End If
    Loop
    
    CalculoPorcentajeDescuento = Str(100 - aDesc * 100)
    Exit Function

errCalculo:
    clsGeneral.OcurrioError "Error al procesar el descuento. " & Trim(Err.Description)
    CalculoPorcentajeDescuento = ""
End Function

Private Sub EditoCelda()
    
    '1- Codigo
    '2- Articulo
    '3- Cantidad c/Cargo
    '4- Cantidad sin Cargo
    '5- Unitario Sin IVA
    '6- Unitario Con IVA
    '7- % Descuento
    '8- SubTotal (sin IVA)
    '9- % IVA del artículo
    '10- Total Subtotal + Iva
    '11- Numero de Remito (si ya está en algun local)
    '12- ID interno de Remito (si ya está en algun local)
    
    Select Case FGrid.Col
        Case 3, 4, 5, 6, 7: If FGrid.TextMatrix(FGrid.Row, 1) <> "" Then FGrid.EditCell
    End Select
    
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    
    LimpioFicha
    sNuevo = False: sModificar = False
    
    CargoDocumentos         'Tipos de Documentos de Compra de Mercaderia
    
    'Cargo las monedas
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda, ""

    DeshabilitoIngreso
        
    InicializoGrilla
    
    If prmIDCompra <> 0 Then
        Me.Show
        CargoCompra prmIDCompra
    End If
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Err.Description
End Sub

Private Sub InicializoGrilla()

    'Encabezado----------------------------------------
    FGrid.Clear
    FGrid.Rows = 1

    FGrid.Row = 0
    FGrid.FormatString = "|Código|Artículo|>Q $|>Q s/c|>Unitario|>Unitario|> % Desc|>SubTotal|>I.V.A.|>Total|<Remito|ID Rem.|Q Rem."
        
    FGrid.WordWrap = True
    FGrid.ColWidth(0) = 0
    FGrid.ColWidth(1) = 700             'Codigo
    FGrid.ColWidth(2) = 2600           'Articulo
    FGrid.ColWidth(3) = 600             'Cantidad Paga
    FGrid.ColWidth(4) = 600             'Cantidad Sin Cargo
    FGrid.ColWidth(5) = 850           'Unitario Sin IVA
    FGrid.ColWidth(6) = 850           'Unitario Con IVA
    FGrid.ColWidth(7) = 700             '% Descuento
    FGrid.ColWidth(8) = 1100             'SubTotal (sin IVA)
    FGrid.ColWidth(9) = 600             '% IVA del artículo
    FGrid.ColWidth(10) = 1050           'Total Subtotal + Iva
    FGrid.ColWidth(11) = 800            'Numero de Remito (si ya está en algun local)
    FGrid.ColWidth(12) = 700            'ID interno de Remito (si ya está en algun local)
    FGrid.ColWidth(13) = 700            'Q original en Remito (si ya está en algun local)
    
    'Formatos-------------------------------------------
    FGrid.ColDataType(2) = flexDTString
    FGrid.ColDataType(3) = flexDTCurrency
    FGrid.ColDataType(4) = flexDTCurrency
    
    FGrid.ColDataType(5) = flexDTCurrency
    FGrid.ColFormat(5) = "#,##0.00"
    FGrid.ColDataType(6) = flexDTCurrency
    FGrid.ColFormat(6) = "#,##0.00"
    
    FGrid.ColDataType(7) = flexDTCurrency
    FGrid.ColFormat(7) = "#.00"
    
    FGrid.ColDataType(8) = flexDTCurrency
    FGrid.ColFormat(8) = "#,##0.00"
    
    FGrid.ColDataType(9) = flexDTCurrency
    FGrid.ColDataType(10) = flexDTCurrency
    FGrid.ColFormat(10) = "#,##0.00"
    FGrid.ColDataType(11) = flexDTLong
    FGrid.ColDataType(12) = flexDTLong
    FGrid.ColDataType(12) = flexDTCurrency
    
    'Oculto las columnas
    FGrid.ColHidden(6) = True       'Precio c/Iva
    FGrid.ColHidden(9) = True       '% Iva
    FGrid.ColHidden(10) = True       'Total + Iva
    FGrid.ColHidden(12) = True     'ID Remito
    FGrid.ColHidden(13) = True     'Q Remito
    
    'Seteo las columnas que se pueden agrupar
    'FGrid.MergeCells = flexMergeRestrictAll
    FGrid.MergeCol(1) = True        'Codigo
    FGrid.MergeCol(2) = True        'Nombre
    'FGrid.MergeCol(5) = True        'Precio sIva
    'FGrid.MergeCol(6) = True        'Precio cIVA
    'FGrid.MergeCol(7) = True        '% Descuento

    FGrid.FixedCols = 3
    FGrid.MergeCells = flexMergeFixedOnly

End Sub

Private Sub CargoArticulos(Proveedor As Long)

    On Error GoTo errCargar
    Screen.MousePointer = 11
    FGrid.Rows = 1
    I = 0
    Cons = "Select * from RemitoCompra, RemitoCompraRenglon, Articulo " _
            & " Where RCoProveedor = " & Proveedor _
            & " And RCoCodigo = RCRRemito " _
            & " And RCRRemanente > 0 " _
            & " And RCRArticulo = ArtId" _
            & " And RCoFecha > '" & Format(DateAdd("yyyy", -2, Date), "mm/dd/yyyy") & "'"
            '& " Order by RCRArticulo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    With FGrid
    Do While Not RsAux.EOF
        I = I + 1
        .AddItem "", I
        .Cell(flexcpText, I, 0) = Trim(RsAux!ArtId)
        .Cell(flexcpText, I, 1) = Trim(RsAux!ArtCodigo)
        .Cell(flexcpBackColor, I, 1) = Inactivo
        .Cell(flexcpText, I, 2) = Trim(RsAux!ArtNombre)
        .Cell(flexcpBackColor, I, 2) = Inactivo
        .Cell(flexcpText, I, 3) = RsAux!RCRRemanente
        .Cell(flexcpText, I, 13) = RsAux!RCRRemanente
        
        .Cell(flexcpText, I, 4) = "0"
        
        .Cell(flexcpBackColor, I, 8) = Inactivo
        .Cell(flexcpFontBold, I, 8) = True
        
        .Cell(flexcpText, I, 9) = IVAArticulo(RsAux!ArtId)
        
        '11 Serie y Nro de Remito - 12 ID Remito
        If Not IsNull(RsAux!ArtId) Then aTexto = Trim(RsAux!RCoSerie) & RsAux!RCoNumero Else aTexto = RsAux!RCoNumero
        .Cell(flexcpText, I, 11) = aTexto
        .Cell(flexcpBackColor, I, 11) = Inactivo
        .Cell(flexcpText, I, 12) = RsAux!RCoCodigo
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If .Rows > 1 Then
        .FixedCols = 0
        .ColSel = 1
        .ColSort(1) = flexSortNumericAscending
        .Sort = flexSortUseColSort

        .FixedCols = 3
        .MergeCells = flexMergeFixedOnly
    End If
    End With
    
    Screen.MousePointer = 0
    tNeto.Text = "0.00": tIva.Text = "0.00": lTTotal.Caption = "0.00"
    lTSubTotal.Caption = "0.00": lTDescuento.Caption = "0.00"
    tCofis.Text = "0.00"
    lNetoCofis.Caption = "0.00": lIva.Caption = "0.00"
    Exit Sub

errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos del proveedor.", Err.Description
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.SimpleText = vbNullString
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    
End Sub

Private Sub Label12_Click()
    Foco tProveedor
End Sub

Private Sub Label2_Click()
    Foco tFecha
End Sub

Private Sub Label5_Click()
    Foco tDescuento
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub
Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub AccionModificar()
    
    If gCompra = 0 Then
        MsgBox "No hay una compra selecconada para modificarla.", vbExclamation, "Posible Error"
        Exit Sub
    End If
    
     If cTipo.ListIndex = -1 Then
        MsgBox "El tipo de comprobante no es correcto." & vbCrLf & _
               "No se puede modificar el registro.", vbExclamation, "Posible Error"
        Exit Sub
    End If
    Screen.MousePointer = 11
    If Val(tProveedor.Tag) <> 0 And FGrid.Rows = FGrid.FixedRows Then CargoArticulos Val(tProveedor.Tag)
        
            
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoIngreso
        
    sModificar = True
    
    On Error Resume Next
    If tDescuento.Enabled Then Foco tDescuento Else FGrid.SetFocus
    
    Screen.MousePointer = 0
End Sub

Private Sub AccionGrabar()

    If Not ValidoDatos Then Exit Sub
    If MsgBox("Confirma almacenar los datos ingresados.", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    
    GraboDatos
    
End Sub

Private Sub AccionCancelar()
    
Dim mAuxiliar As Long

    Screen.MousePointer = 11
    mAuxiliar = 0
    If sModificar And gCompra <> 0 Then mAuxiliar = gCompra
    
    LimpioFicha
    DeshabilitoIngreso
    Botones True, (mAuxiliar <> 0), False, False, False, Toolbar1, Me
    
    If mAuxiliar <> 0 Then CargoCompra mAuxiliar
    
    Screen.MousePointer = 0
    sNuevo = False: sModificar = False
    
End Sub

Private Sub tDescuento_GotFocus()
    tDescuento.SelStart = 0
    tDescuento.SelLength = Len(tDescuento.Text)
    Status.SimpleText = " Ingrese el porcentaje de descuento general de la factura."
End Sub

Private Sub tDescuento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        tDescuento.Text = CalculoPorcentajeDescuento(Trim(tDescuento.Text))
        
        If IsNumeric(tDescuento.Text) And FGrid.Rows > 1 Then
            CargoTotales
        Else
            If Trim(tDescuento.Text) = "" And FGrid.Rows > 1 Then CargoTotales
        End If
        
        FGrid.SetFocus
    End If
    
End Sub

Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0: tFecha.SelLength = Len(tFecha.Text)
    Status.SimpleText = " Ingrese una fecha."
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Not IsDate(tFecha.Text) And (sNuevo Or sModificar) Then
            MsgBox " La fecha ingresada no es correcta.", vbExclamation, "ATENCIÓN"
            Foco tFecha
        Else
            If cTipo.Enabled Then Foco cTipo Else Foco tProveedor
        End If
    End If
    
End Sub

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "d-Mmm-yyyy")
End Sub

Private Sub tFecha_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.SimpleText = " Ingrese una fecha."
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        
        Case "modificar": AccionModificar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "eliminar": EliminarAsignacion
        Case "salir": Unload Me
    End Select

End Sub
Private Sub DeshabilitoIngreso()
    
    tFecha.Enabled = True: tFecha.BackColor = Colores.Blanco
    tCodigo.Enabled = True: tCodigo.BackColor = Colores.Blanco
    tProveedor.Enabled = True: tProveedor.BackColor = Colores.Blanco
    
    cTipo.Enabled = False: cTipo.BackColor = Inactivo
    tFactura.Enabled = False: tFactura.BackColor = Inactivo
    
    bMercaderia.Enabled = False
    cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    tImporte.Enabled = False: tImporte.BackColor = Inactivo
    tArbitraje.Enabled = False: tArbitraje.BackColor = Inactivo
    tDescuento.BackColor = Inactivo: tDescuento.Enabled = False
    
    tIva.Enabled = False: tIva.BackColor = Inactivo
    tNeto.Enabled = False: tNeto.BackColor = Inactivo
    tCofis.Enabled = False: tCofis.BackColor = Inactivo
    
    tUsuario.Enabled = False: tUsuario.BackColor = Inactivo
    tComentario.BackColor = Inactivo: tComentario.Enabled = False
    tDtoRecibo.BackColor = Inactivo: tDtoRecibo.Enabled = False
    
    FGrid.Editable = False
    cIva.Enabled = False
    
    tAutoriza.Enabled = False: tAutoriza.BackColor = Inactivo
    chVerificado.Enabled = False
            
End Sub

Private Sub HabilitoIngreso()

Dim bModificar As Boolean

    bModificar = (Val(bMercaderia.Tag) = 1)
    
    tFecha.Enabled = False: tFecha.BackColor = Colores.Inactivo
    tCodigo.Enabled = False: tCodigo.BackColor = Colores.Inactivo
    tProveedor.Enabled = False: tProveedor.BackColor = Colores.Inactivo
    
    If Not bModificar Then
        tDescuento.BackColor = Blanco: tDescuento.Enabled = True
    End If
    tDtoRecibo.BackColor = Blanco: tDtoRecibo.Enabled = True
    
    'tNeto.Enabled = True: tNeto.BackColor = Blanco
    'tIva.Enabled = True: tIva.BackColor = Blanco
    'tCofis.Enabled = True: tCofis.BackColor = Blanco

    'tUsuario.Enabled = True: tUsuario.BackColor = Obligatorio
    'tComentario.BackColor = Blanco: tComentario.Enabled = True
    
    If Not bModificar Then
        FGrid.Editable = True
        cIva.Enabled = True
    End If
    
    If cTipo.ListIndex <> -1 Then
        Select Case cTipo.ItemData(cTipo.ListIndex)
            Case TipoDocumento.CompraContado, TipoDocumento.CompraCredito: bMercaderia.Enabled = (Not bModificar)
            Case Else: bMercaderia.Enabled = False
        End Select
    Else
        bMercaderia.Enabled = False
    End If
   
End Sub

Private Sub LimpioFicha()
    
    tFecha.Text = ""
    cTipo.Text = ""
    cMoneda.Text = ""
    tProveedor.Text = ""
    tCodigo.Text = ""
    tFactura.Text = ""
    tArbitraje.Text = ""
    tDescuento.Text = ""
    tDtoRecibo.Text = ""
    
    FGrid.Rows = 1
    tNeto.Text = "0.00"
    tIva.Text = "0.00"
    tCofis.Text = "0.00"
    lTTotal.Caption = "0.00"
    lTSubTotal.Caption = "0.00"
    lTDescuento.Caption = "0.00"
    lNetoCofis.Caption = "0.00": lIva.Caption = "0.00"
    
    tUsuario.Text = "": tComentario.Text = ""
    tAutoriza.Text = "": chVerificado.Value = vbUnchecked
    
End Sub

Private Sub tProveedor_Change()
    tProveedor.Tag = 0
    If Not sNuevo And Not sModificar Then gCompra = 0
End Sub

Private Sub tProveedor_GotFocus()
    tProveedor.SelStart = 0: tProveedor.SelLength = Len(tProveedor.Text)
End Sub

Private Sub tNeto_GotFocus()
    tNeto.SelStart = 0: tNeto.SelLength = Len(tNeto.Text)
    Status.SimpleText = " Ingrese el importe de redondeo para la factura."
End Sub

Private Sub tNeto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        On Error GoTo errR
        If IsNumeric(tNeto.Text) Then
            tNeto.Text = Format(tNeto.Text, FormatoMonedaP)
            lNetoCofis.Caption = Format(CCur(tNeto.Text) + CCur(tCofis.Text), FormatoMonedaP)
            lTTotal.Caption = Format(CCur(lNetoCofis.Caption) + CCur(lIva.Caption), FormatoMonedaP)
            
            Foco tCofis
        Else
            If Mid(tNeto.Text, 1, 1) = "*" Then     'Ingresa el Total
                Dim aTotal As Currency
                aTotal = Mid(tNeto.Text, 2, Len(tNeto.Text) - 1)
                lTTotal.Caption = Format(aTotal, FormatoMonedaP)
                lNetoCofis.Caption = Format(aTotal - CCur(lIva.Caption), FormatoMonedaP)
                tNeto.Text = Format(CCur(lNetoCofis.Caption) - CCur(tCofis.Text), FormatoMonedaP)
                Foco tCofis
            End If
        End If
    End If
    
    Exit Sub
errR:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al procesar los importes.", Err.Description
End Sub


Private Function ValidoDatos() As Boolean

    ValidoDatos = False
    
    If FGrid.Rows < 2 Then
        MsgBox "No se han ingresado los artículos del documento.", vbExclamation, "ATENCIÓN"
        Foco tProveedor: Exit Function
    End If
    
'    If tUsuario.Tag = "" Or Val(tUsuario.Tag) = 0 Then
'        MsgBox "Debe ingresar el dígito  de usuario.", vbExclamation, "ATENCIÓN"
'        Foco tUsuario: Exit Function
'    End If
    
    If Val(tProveedor.Tag) = 0 Then
        MsgBox "Seleccione el proveedor de la mercadería.", vbExclamation, "ATENCIÓN"
        Foco tProveedor: Exit Function
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Seleccione la moneda del documento.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
    
    If cTipo.ListIndex = -1 Then
        MsgBox "Seleccione el tipo de documento ingresado.", vbExclamation, "ATENCIÓN"
        Foco cTipo: Exit Function
    End If
               
    If Not IsNumeric(tCofis.Text) Then
        MsgBox "El cofis del comprobante no es correcto.", vbExclamation, "Posible Error": Foco tCofis: Exit Function
    End If
    If Not IsNumeric(tIva.Text) Then
        MsgBox "El IVA del comprobante no es correcto.", vbExclamation, "Posible Error": Foco tIva: Exit Function
    End If
    
    If CCur(tNeto.Text) = 0 Or CCur(tIva.Text) = 0 Or CCur(tCofis.Text) = 0 Then
        If MsgBox("Los subtotales ingresados están en cero. ¿Desea continuar con el ingreso?.", vbExclamation + vbOKCancel + vbDefaultButton2, "ATENCIÓN") = vbCancel Then
            FGrid.SetFocus: Exit Function
        End If
    End If
    
    'Valido si todos los precios unitarios fueron ingresados
    Dim aPrecios As Currency: aPrecios = 0
    For I = 1 To FGrid.Rows - 1
        aPrecios = aPrecios + FGrid.ValueMatrix(I, 3) + FGrid.ValueMatrix(I, 4)
    Next I
    If aPrecios = 0 Then
        MsgBox "No se ingresaron las cantidades asignadas a la factura.", vbExclamation, "ATENCIÓN"
        FGrid.SetFocus: Exit Function
    End If
    
    'Neto + Iva  = Total
    If CCur(tNeto.Text) + CCur(tIva.Text) + CCur(tCofis.Text) <> CCur(lTTotal.Caption) And (CCur(tNeto.Text) + CCur(tIva.Text) + CCur(tCofis.Text) <> 0) Then
        MsgBox "Error en los totales ingresados (Neto + Cofis + I.V.A.  <> Total).", vbExclamation, "ATENCIÓN"
        Foco tNeto: Exit Function
    End If
    
    If Abs(CCur(tImporte.Text) - CCur(lTTotal.Caption)) >= 1 Then
        If MsgBox("Los importes ingresados NO COINCIDEN !! Total de la Boleta <> Total de los Artículos." & vbCr & _
                "¿Está seguro que quiere continuar?", vbExclamation + vbYesNo + vbDefaultButton2, "Posible Error") = vbNo Then
            Foco tNeto: Exit Function
        End If
    End If
    
    'Verifico que los precios no esten en blanco si las cantidades son mayores a cero (0)
    For I = 1 To FGrid.Rows - 1
        If FGrid.Cell(flexcpValue, I, 3) + FGrid.Cell(flexcpValue, I, 4) > 0 Then
            If FGrid.Cell(flexcpText, I, 5) = "" Then
                MsgBox "No se han ingresado los precios para el artículo <<" & FGrid.Cell(flexcpText, I, 2) & ">> " & " del Remito " & FGrid.Cell(flexcpText, I, 11), vbExclamation, "Faltan Precios"
                Exit Function
            End If
            
            If FGrid.Cell(flexcpValue, I, 5) = 0 Then
                If MsgBox("Los precios para el artículo <<" & FGrid.Cell(flexcpText, I, 2) & ">> " & " del Remito " & FGrid.Cell(flexcpText, I, 11) & " están en cero." & Chr(vbKeyReturn) & "Desea continuar.", vbExclamation + vbYesNo + vbDefaultButton2, "Precios en cero") = vbNo Then Exit Function
            End If
        End If
        
    Next I
    
    If Val(tAutoriza.Tag) = 0 And tAutoriza.Enabled Then
        MsgBox "Debe ingresar el usuario que autoriza el ingreso del gasto.", vbExclamation, "Falta Usuario que Autoriza el Gasto"
        Foco tAutoriza: Exit Function
    End If
    
    ValidoDatos = True
    
End Function

Private Sub GraboDatos()
'   --> Trabajo con la mercadería que figura en compañia.
    '0- Id Articulo
    '1- Codigo
    '2- Articulo
    '3- Cantidad c/Cargo
    '4- Cantidad sin Cargo
    '5- Unitario Sin IVA
    '6- Unitario Con IVA
    '7- % Descuento
    '8- SubTotal (sin IVA)
    '9- % IVA del artículo
    '10- Total Subtotal + Iva
    '11- Numero de Remito (si ya está en algun local)
    '12- ID interno de Remito (si ya está en algun local)
    '13- Q original en remito

Dim aCodigoCompra As Long
Dim aUnitario As Currency, aDescU As Currency
Dim aCantidad As Currency, aCantidadF As Currency, aSubTotal As Currency
Dim J As Integer, SumoAI As Integer
Dim pcurPrecioReal As Currency, pcurDescuento As Currency

    FechaDelServidor
    On Error GoTo ErrGD
    cBase.BeginTrans            '--------------------------------------------------------------------------------
    On Error GoTo ErrResumo
    
    aCodigoCompra = gCompra
    If Val(bMercaderia.Tag) = 0 Then
        'Actualizo los datos de cada cada Remito de Compra e inserto renglones de Compra
        'El precio del articulo y los descuentos son SIN IVA
        For I = 1 To FGrid.Rows - 1
        
            aCantidad = FGrid.ValueMatrix(I, 3) + FGrid.ValueMatrix(I, 4)
            aCantidadF = FGrid.ValueMatrix(I, 3)
            aSubTotal = FGrid.ValueMatrix(I, 8)
            SumoAI = 0
            
            'Hay que hacer un loop para ver si hay otro artículo (el siguiente)
            For J = I + 1 To FGrid.Rows - 1
                If FGrid.ValueMatrix(I, 0) = FGrid.ValueMatrix(J, 0) Then
                    aCantidad = aCantidad + FGrid.ValueMatrix(J, 3) + FGrid.ValueMatrix(J, 4)
                    aCantidadF = aCantidadF + FGrid.ValueMatrix(J, 3)
                    aSubTotal = aSubTotal + FGrid.ValueMatrix(J, 8)
                    SumoAI = SumoAI + 1
                Else
                    Exit For
                End If
            Next J
            
            If aCantidad > 0 Then       'Si hay cantidades mayores a 0 grabo sino NOOO
                'Inserto el Renglon en la Tabla CompraRenglon
                Cons = "Insert into CompraRenglon (CReCompra, CReArticulo, CReCantidad, CRePrecioU, CRePrecioReal, CReACostear) Values(" _
                       & aCodigoCompra & ", " _
                       & FGrid.ValueMatrix(I, 0) & ", " _
                       & aCantidad & ", "
                
                '8- SubTotal (sin IVA c/descuento de renglon)     ---> Unitario = STotal / (Q$ + Qsc) - %desc general
                aUnitario = aSubTotal / aCantidad
                'Si hay descuento global, lo aplico
                If IsNumeric(tDescuento.Text) Then aUnitario = aUnitario - (aUnitario * CCur(tDescuento.Text) / 100)
                
                pcurPrecioReal = aUnitario
                If IsNumeric(Trim(tDtoRecibo.Text)) Then
                    pcurDescuento = CCur(Trim(tDtoRecibo.Text))
                    pcurPrecioReal = pcurPrecioReal - (pcurPrecioReal * pcurDescuento / 100)
                End If
                
                Cons = Cons & aUnitario & ", " & pcurPrecioReal & ", "
                Cons = Cons & aCantidad & ")"       'Remanente a Costear
                cBase.Execute Cons
            End If
            
            I = I + SumoAI      'Incremento la I segun el loop
        Next
        
        'Actualizo las cantidades en los remitos de compra
        GraboBDRemitoRenglon
    Else
        If IsNumeric(Trim(tDtoRecibo.Text)) Then
            Dim prsAux As rdoResultset, pstrSQL As String
            pstrSQL = "Select * from CompraRenglon Where CReCompra = " & aCodigoCompra
            Set prsAux = cBase.OpenResultset(pstrSQL, rdOpenDynamic, rdConcurValues)
            Do While Not prsAux.EOF
                prsAux.Edit
                
                pcurPrecioReal = prsAux!CRePrecioU
                pcurDescuento = CCur(Trim(tDtoRecibo.Text))
                pcurPrecioReal = pcurPrecioReal - (pcurPrecioReal * pcurDescuento / 100)
                prsAux!CRePrecioReal = pcurPrecioReal
                prsAux.Update
                prsAux.MoveNext
            Loop
            prsAux.Close
        End If
    End If
    
    cBase.CommitTrans           '------------------------------------------------------------------------------
    
    AccionCancelar
    Exit Sub

ErrGD:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción.", Err.Description
    Exit Sub
ErrResumo:
    Resume ErrRollback
ErrRollback:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    clsGeneral.OcurrioError Msg, Err.Description
    Exit Sub
End Sub

Private Sub GraboBDRemitoRenglon()

    For I = 1 To FGrid.Rows - 1
        If FGrid.ValueMatrix(I, 3) + FGrid.ValueMatrix(I, 4) <> 0 Then      'Si la cantidad <> 0 --> GRABO
            
            Cons = "Select * from RemitoCompraRenglon " _
                  & " Where RCRRemito = " & FGrid.ValueMatrix(I, 12) _
                  & " And RCRArticulo = " & FGrid.ValueMatrix(I, 0)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            RsAux.Edit
            RsAux!RCRRemanente = RsAux!RCRRemanente - (FGrid.ValueMatrix(I, 3) + FGrid.ValueMatrix(I, 4))
            RsAux.Update
            RsAux.Close
            '-------------------------------------------------------------------------------------------
        End If
    Next
    
End Sub

Private Sub CargoDocumentos()

    cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraContado)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraContado
    
    cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraCredito)
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraCredito
    
    'cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaCredito)
    'cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraNotaCredito
    
    'cTipo.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaDevolucion)
    'cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraNotaDevolucion
    
End Sub

Private Sub BuscoCompras()

    On Error GoTo errBuscar
    If sNuevo Or sModificar Then Exit Sub
    
    If Val(tProveedor.Tag) = 0 And Not IsDate(tFecha.Text) Then
        MsgBox "Debe ingresar algunos de los filtros para ver las compras (fecha o proveedor).", vbExclamation, "Fltan Filtros"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Cons = " SELECT ComCodigo, ComCodigo Compra, ComFecha Fecha, PMeNombre Proveedor, ComSerie Serie, ComNumero 'Número', MonSigno Moneda, ComImporte Importe, ComComentario Comentarios" _
            & " From Compra, ProveedorMercaderia, Moneda " _
            & " WHERE ComProveedor = PMECodigo " _
            & " And ComMoneda = MonCodigo" _
            & " And ComCodigo In (Select CReCompra From CompraRenglon) "
            
    If Val(tProveedor.Tag) <> 0 Then Cons = Cons & " And ComProveedor = " & Val(tProveedor.Tag)
    If IsDate(tFecha.Text) Then Cons = Cons & " And ComFecha >= '" & Format(tFecha.Text, sqlFormatoF) & "'"
                
    Cons = Cons & " ORDER BY ComFecha DESC"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsAux.EOF Then
        MsgBox "No hay facturas ingresadas para el proveedor " & Trim(tProveedor.Text), vbExclamation, "ATENCIÓN"
        RsAux.Close: Screen.MousePointer = 0: Exit Sub
    End If
    RsAux.Close
        
    Dim aLista As New clsListadeAyuda
    Dim aIdSel As Long
    aIdSel = aLista.ActivarAyuda(cBase, Cons, 8000, 1)
    If aIdSel <> 0 Then aIdSel = aLista.RetornoDatoSeleccionado(0)
    Set aLista = Nothing
    Me.Refresh
    
    If aIdSel <> 0 Then CargoCompra aIdSel
    
    Screen.MousePointer = 0
    Exit Sub
        
errBuscar:
    clsGeneral.OcurrioError "Error al cargar los datos de la compra.", Err.Description
    Botones True, False, False, False, False, Toolbar1, Me
    Screen.MousePointer = 0
End Sub

Private Sub CargoCompra(Codigo As Long)
    
    Botones True, False, False, False, False, Toolbar1, Me
    LimpioFicha
    gCompra = 0
    
    Dim pcurTotal As Currency
    
    'Cargo los datos de la tabla compra-------------------------------------------------------------------------------
    Cons = "Select Compra.*, ProveedorMercaderia.*, " & _
                         " UsrA.UsuCodigo as UsuACodigo, UsrA.UsuIdentificacion as UsuAIdentificacion, " & _
                         " UsrC.UsuCodigo as UsuCCodigo, UsrC.UsuDigito as UsuCDigito " & _
                " From Compra " & _
                    " Left Outer Join Usuario UsrA On ComUsrAutoriza = UsrA.UsuCodigo " & _
                    " Left Outer Join Usuario UsrC On ComUsuario = UsrC.UsuCodigo, " & _
                " ProveedorMercaderia " & _
                " Where ComCodigo = " & Codigo & _
                " And ComProveedor = PMeCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)

    If Not RsAux.EOF Then
        tCodigo.Text = RsAux!ComCodigo
        tFecha.Text = Format(RsAux!ComFecha, "d-Mmm-yyyy")
        BuscoCodigoEnCombo cTipo, RsAux!ComTipoDocumento
        If Not IsNull(RsAux!ComNumero) Then tFactura.Text = Trim(RsAux!ComNumero)
        
        If Not IsNull(RsAux!PMeNombre) Then tProveedor.Text = Trim(RsAux!PMeNombre)
        tProveedor.Tag = RsAux!ComProveedor
        
        pcurTotal = RsAux!ComImporte
        If Not IsNull(RsAux!ComIVA) Then pcurTotal = pcurTotal + RsAux!ComIVA
        
        tImporte.Text = Format(pcurTotal, "#,##0.00")

        BuscoCodigoEnCombo cMoneda, RsAux!ComMoneda
        If Not IsNull(RsAux!ComTC) Then tArbitraje.Text = Format(RsAux!ComTC, "#.000") Else tArbitraje.Text = "1.000"
        
        If Not IsNull(RsAux!ComDescuento) Then lTDescuento.Caption = Format(RsAux!ComDescuento, FormatoMonedaP) Else lTDescuento.Caption = "0.00"
        
        'Aca hay que poner lo asignado al rubro CompraMercadería y listo !! -------------------------
        tNeto.Text = Format(RsAux!ComImporte, FormatoMonedaP)
        
        If Not IsNull(RsAux!ComIVA) Then tIva.Text = Format(RsAux!ComIVA, FormatoMonedaP) Else tIva.Text = "0.00"
        'If Not IsNull(rsAux!ComCofis) Then tCofis.Text = Format(rsAux!ComCofis, FormatoMonedaP) Else tCofis.Text = "0.00"
        
        lNetoCofis.Caption = Format(CCur(tNeto.Text) + CCur(tCofis.Text), FormatoMonedaP)
        lIva.Caption = tIva.Text
        lTTotal.Caption = Format(CCur(lNetoCofis.Caption) + CCur(lIva.Caption), FormatoMonedaP)
        
        lTTotal.Caption = Format(pcurTotal, FormatoMonedaP)
        
        If Not IsNull(RsAux!ComComentario) Then tComentario.Text = Trim(RsAux!ComComentario)
        
        If Not IsNull(RsAux!UsuCDigito) Then
            tUsuario.Text = Trim(RsAux!UsuCDigito)
            tUsuario.Tag = RsAux!UsuCCodigo
        End If
        
        If Not IsNull(RsAux!UsuACodigo) Then
            tAutoriza.Text = Trim(RsAux!UsuAIdentificacion)
            tAutoriza.Tag = RsAux!UsuACodigo
        End If
        lAutoriza.Tag = tAutoriza.Tag
        
        If Not IsNull(RsAux!ComVerificado) Then
            chVerificado.Value = IIf(RsAux!ComVerificado = 1, vbChecked, vbUnchecked)
        Else
            chVerificado.Value = vbGrayed
        End If
        gCompra = Codigo
    End If
    RsAux.Close
    '------------------------------------------------------------------------------------------------------------------------
        
    If gCompra <> 0 Then
        Dim bHayArticulos As Boolean, bEsCompraM As Boolean
        'Cargo los datos de las articulos------------------------------------------------------------------------------------
        Cons = "Select * from CompraRenglon, Articulo" _
                & " Where CReCompra = " & Codigo _
                & " And CReArticulo = ArtID"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        bHayArticulos = Not RsAux.EOF
        If Not RsAux.EOF Then
            If RsAux!CRePrecioU <> RsAux!CRePrecioReal Then
                tDtoRecibo.Text = Format(100 - (RsAux!CRePrecioReal * 100) / RsAux!CRePrecioU, "#,##0.00")
            End If
        End If
        
        Do While Not RsAux.EOF
            With FGrid
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = RsAux!ArtId
                .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "000,000"): .Cell(flexcpBackColor, .Rows - 1, 1) = Inactivo
                .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!ArtNombre): .Cell(flexcpBackColor, .Rows - 1, 2) = Inactivo
                
                .Cell(flexcpText, .Rows - 1, 3) = RsAux!CReCantidad
                .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!CRePrecioU, FormatoMonedaP)
                
                .Cell(flexcpText, .Rows - 1, 8) = Format(RsAux!CRePrecioU * RsAux!CReCantidad, FormatoMonedaP)
                .Cell(flexcpBackColor, .Rows - 1, 8) = Inactivo: .Cell(flexcpFontBold, .Rows - 1, 8) = True
                .Cell(flexcpBackColor, .Rows - 1, 11) = Inactivo
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
        '------------------------------------------------------------------------------------------------------------------------
        
        bMercaderia.Tag = IIf(bHayArticulos, 1, 0)
        bEsCompraM = bHayArticulos
        
        If Not bHayArticulos Then
            Cons = "Select * from GastoSubRubro " & _
                    " Where GSrIDCompra = " & gCompra & _
                    " And GSrIDSubRubro = " & paSubrubroCompraMercaderia
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            bEsCompraM = Not RsAux.EOF
            RsAux.Close
        End If
    
        'If bEsCompraM And Not bHayArticulos Then
        '    If Val(tProveedor.Tag) <> 0 Then CargoArticulos Val(tProveedor.Tag)
        'End If
       
        If bEsCompraM Then 'If FGrid.Rows > 1 Then
            Botones True, True, True, False, False, Toolbar1, Me
        Else
            MsgBox "Posiblemente la compra seleccionada no es una Compra de Mercadería." & vbCrLf & _
                   "Causa: No existen artículos para la compra.", vbInformation, "Es compra de Mercadería ?"
        End If
    End If
    
End Sub

Private Function MesUltimoCosteo() As Date
On Error GoTo errUC
Dim rsMU As rdoResultset

    Screen.MousePointer = 11
    MesUltimoCosteo = "01/01/2000"
    
    Cons = "SELECT Max(CabMesCosteo) FROM CMCabezal"
    '"SELECT CabMesCosteo FROM CMCabezal WHERE CabId IN (SELECT Max(CabID) FROM CMCabezal)"
    Set rsMU = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsMU.EOF Then
        If Not IsNull(rsMU(0)) Then MesUltimoCosteo = rsMU(0) 'rsMU("CabMesCosteo")
    End If
    rsMU.Close
    Screen.MousePointer = 0
    Exit Function
errUC:
    clsGeneral.OcurrioError "Error al buscar la fecha de último costeo.", Err.Description, "Fecha de último costeo"
    Screen.MousePointer = 0
End Function

Private Sub EliminarAsignacion()

'Validación
'Veo si ya está costeado.
    If CDate(tFecha.Text) < MesUltimoCosteo() Then
        MsgBox "La fecha de compra ya fue costeada para eliminar la asignación primero debe eliminar el costeo.", vbInformation, "ATENCIÓN"
        Exit Sub
    End If
    
    If MsgBox("¿Confirma eliminar la asignación?", vbQuestion + vbYesNo, "Eliminar asignación") = vbNo Then Exit Sub

    Dim CantArt As Long
    On Error GoTo errEA
    Screen.MousePointer = 11
    cBase.BeginTrans
    
    Dim rsC As rdoResultset
    'Elimino la tabla comprarenglon.
    Cons = "SELECT * FROM CompraRenglon WHERE CReCompra = " & Val(tCodigo.Text)
    Set rsC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsC.EOF
        CantArt = rsC("CReCantidad")
        
        'Busco en tabla relación de asignación para el artículo y la compra.
        'TODO
        
        'si no está en la tabla relación busco los remitos con fecha menor a la compra.
        Cons = "SELECT * FROM RemitoCompraRenglon " & _
                "WHERE RCRRemito IN (SELECT RCoCodigo FROM RemitoCompra WHERE RCoProveedor IN( SELECT ComProveedor FROM Compra WHERE ComCodigo = " & Val(tCodigo.Text) & ")) " & _
                "AND RCRArticulo = " & rsC("CReArticulo") & " AND RCRRemanente < RCRCantidad order by RCRRemito desc"
                
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        Dim aPosibles As Integer
        Dim aDescontar As Integer
        Do While Not RsAux.EOF And CantArt > 0
            aPosibles = RsAux("RCRCantidad") - RsAux("RCRRemanente")
            If aPosibles >= CantArt Then
                aDescontar = CantArt
                CantArt = 0
            Else
                aDescontar = aPosibles
                CantArt = CantArt - aPosibles
            End If
            
            RsAux.Edit
            RsAux("RCRRemanente") = RsAux("RCRRemanente") + aDescontar
            RsAux.Update
            
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        'Borro el registro.
        rsC.Delete
        
        rsC.MoveNext
    Loop
    rsC.Close
    
    cBase.CommitTrans
    Screen.MousePointer = 0
    tCodigo_KeyPress 13
    Exit Sub

errEA:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción.", Err.Description
    Exit Sub
ErrResumo:
    Resume ErrRollback
ErrRollback:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al intentar eliminar la asignación.", Err.Description
    Exit Sub

End Sub
