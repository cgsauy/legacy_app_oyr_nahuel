VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{190700F0-8894-461B-B9F5-5E731283F4E1}#1.1#0"; "orHiperlink.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{D851F632-A4E6-4F61-863C-9480B5EC86D9}#1.2#0"; "ORGDAT~1.OCX"
Object = "{ED77B5E2-D033-48D4-A500-5B5C27404B54}#1.1#0"; "orgMenu.ocx"
Begin VB.Form frmDistribuirEnvio 
   Appearance      =   0  'Flat
   Caption         =   "Distribución de Envíos"
   ClientHeight    =   6960
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   13470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDistEnvio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   13470
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmModal 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7680
      Top             =   4080
   End
   Begin VB.PictureBox picCodImpresion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   10305
      TabIndex        =   9
      Top             =   1320
      Width           =   10335
      Begin VB.TextBox txtEnvioEntregado 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         MaxLength       =   11
         TabIndex        =   14
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox tCodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   9
         TabIndex        =   11
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Envío entregado:"
         Height          =   255
         Left            =   4680
         TabIndex        =   15
         Top             =   120
         Width           =   1455
      End
      Begin VB.Image imgMercaderia 
         Height          =   480
         Left            =   7200
         MouseIcon       =   "frmDistEnvio.frx":0442
         MousePointer    =   99  'Custom
         Picture         =   "frmDistEnvio.frx":074C
         Top             =   75
         Width           =   480
      End
      Begin VB.Label lbMercaderia 
         BackStyle       =   0  'Transparent
         Caption         =   "Hay mercadería para reclamar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   7800
         MouseIcon       =   "frmDistEnvio.frx":0DD1
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   195
         Width           =   2535
      End
      Begin VB.Label lCamion 
         BackColor       =   &H007280FA&
         BackStyle       =   0  'Transparent
         Caption         =   "Camión: Martín"
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
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Top             =   150
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código de impresión:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   150
         Width           =   1695
      End
   End
   Begin prjorgMenu.orgMenu oMenu 
      Left            =   8280
      Top             =   2280
      _ExtentX        =   820
      _ExtentY        =   820
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
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   6705
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3246
            MinWidth        =   3246
            Text            =   "Envíos"
            TextSave        =   "Envíos"
            Key             =   "envio"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3254
            MinWidth        =   3246
            Text            =   "Tipo"
            TextSave        =   "Tipo"
            Key             =   "tipo"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3254
            MinWidth        =   3246
            Text            =   "Artículo"
            TextSave        =   "Artículo"
            Key             =   "articulo"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Bultos"
            TextSave        =   "Bultos"
            Key             =   "bultos"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Tarde a partir:"
            TextSave        =   "Tarde a partir:"
            Key             =   "tarde"
            Object.ToolTipText     =   "Doble clic cambia el valor"
         EndProperty
      EndProperty
   End
   Begin vsViewLib.vsPrinter vsPrint 
      Height          =   4995
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   6495
      _Version        =   196608
      _ExtentX        =   11456
      _ExtentY        =   8811
      _StockProps     =   229
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   End
   Begin MSComctlLib.ImageList imgMini 
      Left            =   240
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483648
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistEnvio.frx":10DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistEnvio.frx":143E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tooMenu 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   1058
      ButtonWidth     =   2355
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Asignar"
            Key             =   "save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sugerir Camión"
            Key             =   "sugerircamion"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ver Envío"
            Key             =   "envio"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Impresión"
            Key             =   "print"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Otros"
            Key             =   "otros"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mercadería"
            Key             =   "mercaderia"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   13470
      TabIndex        =   4
      Top             =   600
      Width           =   13470
      Begin VB.Label lTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Envíos a distribuir para el 10/10/2004"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   60
         Width           =   9375
      End
   End
   Begin VB.PictureBox picLink 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H0000A0A0&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   13440
      TabIndex        =   3
      Top             =   1005
      Width           =   13470
      Begin orgDateCtrl.orgDate tFecha 
         Height          =   315
         Left            =   510
         TabIndex        =   1
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   39300
      End
      Begin prjHiperLink.orHiperLink hlTab 
         Height          =   495
         Index           =   0
         Left            =   1920
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BackColor       =   540847
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorOver   =   16777215
         Caption         =   "Para Asignar"
         MouseIcon       =   "frmDistEnvio.frx":17A5
         MousePointer    =   99
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjHiperLink.orHiperLink hlTab 
         Height          =   495
         Index           =   1
         Left            =   3240
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BackColor       =   9437184
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorOver   =   16777215
         Caption         =   "A Imprimir"
         MouseIcon       =   "frmDistEnvio.frx":1ABF
         MousePointer    =   99
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjHiperLink.orHiperLink hlTab 
         Height          =   495
         Index           =   2
         Left            =   4440
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BackColor       =   9437184
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorOver   =   16777215
         Caption         =   "Por Entregar"
         MouseIcon       =   "frmDistEnvio.frx":1DD9
         MousePointer    =   99
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjHiperLink.orHiperLink hlTab 
         Height          =   495
         Index           =   3
         Left            =   5760
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BackColor       =   9437184
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorOver   =   16777215
         Caption         =   "A Confirmar"
         MouseIcon       =   "frmDistEnvio.frx":20F3
         MousePointer    =   99
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjHiperLink.orHiperLink hlTab 
         Height          =   495
         Index           =   4
         Left            =   8400
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BackColor       =   6325760
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorOver   =   16777215
         Caption         =   "Estadísticas"
         MouseIcon       =   "frmDistEnvio.frx":240D
         MousePointer    =   99
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjHiperLink.orHiperLink hlTab 
         Height          =   495
         Index           =   5
         Left            =   7080
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BackColor       =   9437184
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorOver   =   16777215
         Caption         =   "Recepcionar"
         MouseIcon       =   "frmDistEnvio.frx":2727
         MousePointer    =   99
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjHiperLink.orHiperLink hlTab 
         Height          =   495
         Index           =   6
         Left            =   9720
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BackColor       =   16761024
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorOver   =   16777215
         Caption         =   "Arts a Reclamar"
         MouseIcon       =   "frmDistEnvio.frx":2A41
         MousePointer    =   99
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjHiperLink.orHiperLink hlTab 
         Height          =   495
         Index           =   7
         Left            =   11280
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BackColor       =   9084546
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColorOver   =   16777215
         Caption         =   "Duplicados"
         MouseIcon       =   "frmDistEnvio.frx":2D5B
         MousePointer    =   99
         BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   1800
         X2              =   1800
         Y1              =   0
         Y2              =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Día:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   375
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrid 
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4260
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   7368816
      ForeColorFixed  =   -2147483634
      BackColorSel    =   13891065
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   15724527
      GridColor       =   -2147483636
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
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
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   960
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistEnvio.frx":3075
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistEnvio.frx":3187
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistEnvio.frx":35D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistEnvio.frx":36EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistEnvio.frx":3B3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistEnvio.frx":3E57
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistEnvio.frx":45A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistEnvio.frx":46BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistEnvio.frx":696D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistEnvio.frx":8C1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistEnvio.frx":AED1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuOtros 
      Caption         =   "otros"
      Visible         =   0   'False
      Begin VB.Menu MnuOtrEnvio 
         Caption         =   "Envíos"
      End
      Begin VB.Menu MnuOtrFichaAgencia 
         Caption         =   "Fichas de Agencia"
      End
      Begin VB.Menu MnuOtrCambioEstadoEnvio 
         Caption         =   "Cambio de estado"
      End
      Begin VB.Menu MnuOtrLine1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOtrAgendaFlete 
         Caption         =   "Agenda de fletes"
      End
      Begin VB.Menu MnuOtrTiposFletes 
         Caption         =   "Tipos de fletes"
      End
      Begin VB.Menu MnuOtrLine2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOtrDeshabilitoCamion 
         Caption         =   "Deshabilitar camión"
      End
      Begin VB.Menu MnuOtrHabilitarCamion 
         Caption         =   "Habilitar camión"
      End
      Begin VB.Menu MnuOtrLine3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOtrSendWhatsapp 
         Caption         =   "Enviar whatsapp"
      End
      Begin VB.Menu MnuOtrEnviarWhatsApp 
         Caption         =   "Enviar whatsapp a código"
      End
   End
   Begin VB.Menu MnuGrid 
      Caption         =   "BotonD"
      Visible         =   0   'False
      Begin VB.Menu MnuGridEditEnvio 
         Caption         =   "Editar Envío"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuRecepEntregado 
         Caption         =   "Dar por entregado"
      End
      Begin VB.Menu MnuCambiarLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEliminoVaCon 
         Caption         =   "Eliminar Va Con"
      End
      Begin VB.Menu MnuCambiarDividir 
         Caption         =   "Dividir envío"
      End
      Begin VB.Menu MnuCambiarHora 
         Caption         =   "Cambiar fecha y horario"
      End
      Begin VB.Menu MnuCambiarCamion 
         Caption         =   "Cambiar camionero"
      End
      Begin VB.Menu MnuCambiarAnular 
         Caption         =   "Anular el envío"
      End
      Begin VB.Menu MnuGridL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGridTodosA 
         Caption         =   "Asignar cliqueados a ..."
         Begin VB.Menu MnuGridAsiTodosCamion 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu MnuGridACamion 
         Caption         =   "Asignar envío(s) a ..."
         Begin VB.Menu MnuGridCamion 
            Caption         =   "Automáticamente"
            Index           =   0
         End
         Begin VB.Menu MnuGridCamion 
            Caption         =   "-"
            Index           =   1
         End
      End
      Begin VB.Menu MnuGridSugerirC 
         Caption         =   "Sugerir Camión"
         Begin VB.Menu MnuGSCSelect 
            Caption         =   "A Seleccionados"
         End
         Begin VB.Menu MnuGSCResto 
            Caption         =   "A todos los que no tengan"
         End
      End
      Begin VB.Menu MnuGridL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGridSelectAll 
         Caption         =   "Marcar todos"
      End
      Begin VB.Menu MnuGridUnSelectAll 
         Caption         =   "Desmarcar todos"
      End
      Begin VB.Menu MnuGridL3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGridFiltroSel 
         Caption         =   "Filtrar Por Selección"
      End
      Begin VB.Menu MnuGridFiltroExcSel 
         Caption         =   "Filtrar excluyendo la selección"
      End
      Begin VB.Menu MnuGridL4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGridAConfirmar 
         Caption         =   "Cambiar Estado A Confirmar"
      End
      Begin VB.Menu MnuGridAAsignar 
         Caption         =   "Cambiar Estado Para Asignar"
      End
      Begin VB.Menu MnuGridL5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrint 
         Caption         =   "Imprimir"
         Begin VB.Menu MnuPrintReimprimir 
            Caption         =   "Reimprimir planilla para un código"
         End
         Begin VB.Menu MnuPrintAux 
            Caption         =   "Planilla Auxiliar"
         End
         Begin VB.Menu MnuPrintDocumento 
            Caption         =   "Reimprimir Remitos"
         End
         Begin VB.Menu MnuPrintBultos 
            Caption         =   "Bultos en reparto"
         End
         Begin VB.Menu MnuPrintLine 
            Caption         =   "-"
         End
         Begin VB.Menu MnuPrintPapelRosa 
            Caption         =   "Bandeja copia de eTicket "
         End
         Begin VB.Menu MnuPrintConfig 
            Caption         =   "Configurar"
         End
      End
   End
   Begin VB.Menu MnuCamion 
      Caption         =   "Camión"
      Visible         =   0   'False
      Begin VB.Menu MnuAIECamion 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu MnuMercaderia 
      Caption         =   "Mercadería"
      Visible         =   0   'False
      Begin VB.Menu MnuMerReclamar 
         Caption         =   "Reclamar mercadería al camión"
      End
      Begin VB.Menu MnuMerLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMerDar 
         Caption         =   "Entregar mercadería al camión"
      End
      Begin VB.Menu MnuMerDevolucion 
         Caption         =   "Devolución de mercadería"
      End
      Begin VB.Menu MnuMerStockTotal 
         Caption         =   "Stock total"
      End
   End
   Begin VB.Menu MnuBultos 
      Caption         =   "Bultos"
      Visible         =   0   'False
      Begin VB.Menu MnuBulIdx 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmDistribuirEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modificaciones
'8/1/2007      s_VerficoEnvioEditado encontre que concatenaba en .rows -1 el tipo de artículo, limpio el texto de los artículos y el data donde guardo la q y el tipo.
'9/1/2007      le agregue la cantidad de envíos que posee c/camión en envíos asignados.
'              Si el envío tiene agencia muestro el nombre de la misma en lugar de la zona.
'              Agregé en impresión la carga del valor flete y corregí el valor a cobrar en vacon no cargaba el de todos y pase a leer tabla docpendientes.
'3/6/2008
'               Si un remito o ctdo vuelve a salir en un nuevo código de impresión lo imprimo en hoja en blanco.
'7/6/2008
'               Puse condición para imprimir o no en la sección del camión los artículos (es campo de bd fijo devuelto en el sp).
Option Explicit
Dim oArtFlete As clsProducto
Dim oLog As New clsVBLog

Private sNTipo As String, iIDTipo As Long
Private sCodArt As String, iIDArt As Long
Private bPrintColEsArt As Boolean, bPrintPlanilla As Boolean

Private iCamSelect As Integer               'Camión seleccionado en tag = 1 o tag = 2
Private Const cte_FormatFH As String = "mm/dd/yyyy hh:nn:ss"
Private colBultos As Collection
Private docToPrint As Collection
Private docCopia As Collection

Private Type tArticulo
    ID As Long
    Nombre As String
End Type

Private Enum TiposDeImpresionPlanillas
    ETDIP_ConAgencia = 1
    ETDIP_SinAgenciaComun = 2
    ETDIP_SinAgenciaEspecial = 3
    ETDIP_EnviosRemitos = 4
End Enum

Private arrNomArt() As tArticulo

Private Type tCamion
    Codigo As Long
    Nombre As String
    Habilitado As Boolean
End Type
Private arrCamion() As tCamion

Private Const settingCopiaeTicket = "Print_RepartoCopiaeTicket"

Private Function whatsappMsgEnvioConAgencia() As String

    whatsappMsgEnvioConAgencia = "¡Gracias por comprar en *Carlos Gutiérrez!*" & vbCrLf & _
                "Estamos prontos para *despachar [diaEnvio] en Agencia [nombreagencia] entre las [horaEnvioInit] y [horaEnvioEnd] hs*" & vbCrLf & " " & vbCrLf & _
                "[tablaArticulosEnvio]" & vbCrLf & " " & vbCrLf & _
                "a *[callePuertaDireccion]*[entreesquinasDireccion]" & vbCrLf & _
                "*[localidaddeptoDireccion]*" & vbCrLf & " " & vbCrLf & _
                "Solo si encontrás alguna discrepancia, comunícate al tel. 29027737." & vbCrLf & _
                "¡Muchas gracias!"
                
End Function

Private Function whatsappMsgEnvioSinAgencia() As String

    whatsappMsgEnvioSinAgencia = "¡Gracias por comprar en *Carlos Gutiérrez!*" & vbCrLf & _
                "*Estamos prontos para entregar [diaEnvio]*:" & vbCrLf & " " & vbCrLf & _
                "[tablaArticulosEnvio]" & vbCrLf & " " & vbCrLf & _
                "Por favor, *asegurate de que en [callePuertaDireccion]* [entreesquinasDireccion], [localidaddeptoDireccion]  *entre las [horaEnvioInit] y [horaEnvioEnd] hs* haya una persona *responsable todo el horario*[desembalarmercaderia] y revise su estado exterior." & vbCrLf & " " & vbCrLf & _
                "Solo si encontrás alguna discrepancia, comunícate al tel. 29027737." & vbCrLf & _
                "¡Muchas gracias!"

End Function

Private Function whatsappMsgEnvioSinAgenciaConCobranza() As String
    whatsappMsgEnvioSinAgenciaConCobranza = "¡Gracias por comprar en *Carlos Gutiérrez!*" & vbCrLf & _
                "*Estamos prontos para entregar [diaEnvio]*:" & vbCrLf & " " & vbCrLf & _
                "[tablaArticulosEnvio]" & vbCrLf & " " & vbCrLf & _
                "Por favor, *asegurate de que en [callePuertaDireccion]* [entreesquinasDireccion], [localidaddeptoDireccion]  *entre las [horaEnvioInit] y [horaEnvioEnd] hs* esté:" & vbCrLf & " " & vbCrLf & _
                "- Una persona *responsable todo el horario* que desembale la mercadería y revise su estado exterior. " & vbCrLf & _
                "- El pago de *$ [cobranzaenvio] en efectivo* o cheque (si fue autorizado previamente)." & vbCrLf & " " & vbCrLf & _
                "Solo si encontrás alguna discrepancia, comunícate al tel. 29027737." & vbCrLf & _
                "¡Muchas gracias!"
End Function

Private Function whatsappMsgEnvioRetiroDomicilio(ByVal pagaAlgo As Boolean) As String
'RETIRA UN ARTíCULO  y no deja nada.
    whatsappMsgEnvioRetiroDomicilio = "*Carlos Gutiérrez* te informa que *[diaEnvio]* retiraremos:" & vbCrLf & _
                "*[tablaArticulosEnvio]*" & vbCrLf & " " & vbCrLf & _
                "Por favor, *asegurate de que en [callePuertaDireccion]* [entreesquinasDireccion], [localidaddeptoDireccion]  *entre las [horaEnvioInit] y [horaEnvioEnd] hs* esté:" & vbCrLf & " " & vbCrLf & _
                "- Una persona *responsable todo el horario*. " & vbCrLf & _
                "- *El artículo pronto para ser retirado* (desenchufado, desinstalado y fuera del soporte)." & vbCrLf & _
                IIf(pagaAlgo, "- El pago de *$ [cobranzaenvio] en efectivo* o cheque (si fue autorizado previamente).", "Si hay un saldo a favor,  lo acreditaremos en la cuenta del cliente en Carlos Gutiérrez y dispondrá de 6 meses para utilizarlo.") & vbCrLf & " " & vbCrLf & _
                "Solo si encontrás alguna discrepancia, comunícate al tel. 29027737." & vbCrLf & _
                "¡Muchas gracias!"
End Function

Private Function whatsappMsgEnvioCambioXOtroIgual(ByVal bPagaAlgo As Boolean) As String
'2) RETIRA UN Artículo   por OTRO ARTICULO IGUAL
    whatsappMsgEnvioCambioXOtroIgual = "*Carlos Gutiérrez!* te informa que *[diaEnvio]* cambiaremos:" & vbCrLf & _
                "*[tablaArticulosEnvio]*" & vbCrLf & " " & vbCrLf & _
                "Por favor, *asegurate de que en [callePuertaDireccion]* [entreesquinasDireccion], [localidaddeptoDireccion]  *entre las [horaEnvioInit] y [horaEnvioEnd] hs* esté:" & vbCrLf & " " & vbCrLf & _
                "- Una persona *responsable todo el horario*. " & vbCrLf & _
                "- *El artículo pronto para ser retirado* (desenchufado, desinstalado y fuera del soporte)." & vbCrLf & _
                IIf(bPagaAlgo, "- El pago de *$ [cobranzaenvio] en efectivo* o cheque (si fue autorizado previamente)." & vbCrLf, "") & _
                "Solo si encontrás alguna discrepancia, comunícate al tel. 29027737." & vbCrLf & _
                "¡Muchas gracias!"
End Function

Private Function whatsappMsgEnvioConNota(ByVal pagaAlgo As Boolean) As String
'3) RETIRA UN Artículo   y entrega OTRO ARTICULO MAS CARO. Hay que cobrar
    whatsappMsgEnvioConNota = "*Carlos Gutiérrez!* te informa que *[diaEnvio] retiraremos [tablaARetirar] y entregaremos [tablaArticulosEnvio]*" & vbCrLf & _
                "Por favor, *asegurate de que en [callePuertaDireccion]* [entreesquinasDireccion], [localidaddeptoDireccion]  *entre las [horaEnvioInit] y [horaEnvioEnd] hs* esté:" & vbCrLf & " " & vbCrLf & _
                "- Una persona *responsable todo el horario*. " & vbCrLf & _
                "- *El artículo pronto para ser retirado* (desenchufado, desinstalado y fuera del soporte)." & vbCrLf & _
                IIf(pagaAlgo, "- El pago de *$ [cobranzaenvio] en efectivo* o cheque (si fue autorizado previamente).", "Si hay un saldo a favor,  lo acreditaremos en la cuenta del cliente en Carlos Gutiérrez y dispondrá de 6 meses para utilizarlo.") & vbCrLf & " " & vbCrLf & _
                "Solo si encontrás alguna discrepancia, comunícate al tel. 29027737." & vbCrLf & _
                "¡Muchas gracias!"
End Function

Private Sub whatsappSaveMsg(ByVal idEnvio As Long, ByVal msg As String, ByVal HoraEnvio As Date, ByVal horaFin As Date, ByVal cliente As Long, ByVal celular As String)
    
    celular = "598" + Mid(celular, 2)
    Cons = "INSERT INTO MensajeWhatsApp (MWACliente,MWATipo,MWADocumento,MWATelefono,MWATexto,MWAVencimiento,MWAEstado,MWADesde) Values (" & _
        cliente & ", 1, " & idEnvio & ", '" + Trim(celular) + "', '" + msg + "', '" + Format(horaFin, "yyyyMMdd hh:mm:ss") & "',0, '" + Format(HoraEnvio, "yyyyMMdd hh:mm:ss") & "')"
    cBase.Execute Cons
    
End Sub

Private Sub whatsappEnviarMensajeCodigo(ByVal codImpresion As Long, Optional celTesting As String = "")
'CASE WHEN Es1.CalNombre IS NOT NULL AND Es2.CalNombre IS NOT NULL THEN ' entre ' + rtrim(es1.CalNombre) + ' y ' + RTRIM(Es2.CalNombre) ELSE " & _
'"CASE WHEN Es1.CalNombre IS NOT NULL THEN ' y ' + RTrim(es1.CalNombre) ELSE CASE WHEN Es2.CalNombre IS NOT NULL THEN ' y ' + RTrim(es2.CalNombre) END END END Esquinas, " & _
'RTRIM(ISNULL(ZonNombre, '')) + ', ' +

On Error GoTo ErrBT
    cBase.BeginTrans
On Error GoTo errRB

    Cons = "SELECT ISNULL(EVCID, EnvCodigo) IDentificador, EnvCodImpresion, EnvCliente, AgeNombre, EnvCodigo, EnvFechaPrometida, EnvRangoHora, EnvTelefono, EnvTelefono2 " & _
        ", EVCID, CASE WHEN DepCodigo = 1 THEN RTRIM(DepNombre) ELSE RTRIM(LocNombre) + ', ' + rtrim(DepNombre) END Depto " & _
        ", RTRIM(ISNULL(CP.CalNomFAmiliar, CP.CalNombre)) + ' ' + RTRIM(CAST(DirPuerta as varchar(7))) + CASE WHEN DirBis = 1 THEN ' bis ' ELSE '' END Domicilio " & _
        ", DirLetra, DirApartamento, Isnull(Es1.CalNomFAmiliar, Es1.CalNombre) Esq1, ISNULL(Es2.CalNomFAmiliar, Es2.CalNombre) Esq2 " & _
        ", REvCantidad, ISNULL(AEsNombre, ArtNombre) articulo, ZonNombre, DepCodigo, EnvTipoFlete, EnvReclamoCobro, EReRemito, EnvDocumento, EReNota " & _
        "FROM ENVIO LEFT OUTER JOIN Agencia ON EnvAgencia = AgeCodigo LEFT OUTER JOIN EnvioVaCon ON EVCEnvio = EnvCodigo " & _
        " INNER JOIN Direccion ON EnvDireccion = DirCodigo INNER JOIN Calle CP ON DirCalle = CP.CalCodigo INNER JOIN Localidad ON LocCodigo = CP.CalLocalidad INNER JOIN Departamento ON LocDepartamento = DepCodigo " & _
        " LEFT OUTER JOIN Zona ON EnvZona = ZonCodigo " & _
        " LEFT OUTER JOIN Calle Es1 ON DirEntre1 = Es1.CalCodigo LEFT OUTER JOIN Calle Es2 ON DirEntre2 = Es2.CalCodigo " & _
        " INNER JOIN RenglonEnvio ON EnvCodigo = REvEnvio INNER JOIN Articulos ON ArtID = REvArticulo LEFT OUTER JOIN ArticuloEspecifico ON EnvDocumento = AEsDocumento AND REvArticulo = AEsArticulo " & _
        " LEFT OUTER JOIN EnviosRemitos ON EReEnvio = EnvCodigo " & _
        " WHERE EnvCodImpresion = " & codImpresion & _
        "  AND (EnvTelefono like '09%' OR EnvTelefono2 like '09%') AND EnvTipo = 1 ORDER BY EVCID"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Dim idCli As Long
    Dim idEnv As Long
    Dim idVC As Long
    Dim idIx As Long
    Dim sArts As String
    Dim sMsg As String
    Dim horaI As Integer
    Dim sHorario As String
    
    Dim horaF As Date
    Dim Fecha As Date
    Dim sArticulos As String
    Dim celular As String
    Dim qArts As Integer
    
    Dim cCobro As Currency
    Dim rsC As rdoResultset
    
    Dim bRetiroCambio As Boolean
    Dim bRetiro As Boolean
    Dim sArtsNota As String
    
    
    Do While Not RsAux.EOF
        If (idIx <> RsAux("Identificador")) Then
            
            If (idEnv > 0 And celular <> "") Then
                sMsg = Replace(sMsg, "[tablaArticulosEnvio]", sArticulos)
                sMsg = Replace(sMsg, "[tablaARetirar]", sArtsNota)
                whatsappSaveMsg idEnv, sMsg, Fecha, horaF, idCli, celular
            End If
            
            sArtsNota = ""
            qArts = 0
            idEnv = RsAux("EnvCodigo")
            idCli = RsAux("EnvCliente")
            idIx = RsAux("IDentificador")
            cCobro = 0
            bRetiroCambio = Not IsNull(RsAux("EReRemito"))
            bRetiro = False
            
            If (idIx <> idEnv) Then
                Cons = "SELECT SUM(DPeImporte) from DocumentoPendiente where DPeTipo = 1 AND DPeIDTipo IN (SELECT EVcEnvio FROM EnvioVaCon WHERE EVCID = " & idIx & ")"
            Else
                Cons = "SELECT SUM(DPeImporte) from DocumentoPendiente where DPeTipo = 1 AND DPeIDTipo = " & idEnv
            End If
            Set rsC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not rsC.EOF Then
                If Not IsNull(rsC(0)) Then cCobro = rsC(0)
            End If
            rsC.Close
            
            If cCobro > 0 Then 'Not IsNull(RsAux("EnvReclamoCobro")) Then
                If (bRetiroCambio) Then
                    'Puede tener flete
                    If (RsAux("EReRemito") = RsAux("EnvDocumento")) Then
                        bRetiro = True
                        sMsg = whatsappMsgEnvioRetiroDomicilio(True)
                    ElseIf Not IsNull(RsAux("EReNota")) Then
                        sMsg = whatsappMsgEnvioConNota(True)
                    Else
                        sMsg = whatsappMsgEnvioCambioXOtroIgual(True)
                    End If
                Else
                    sMsg = whatsappMsgEnvioSinAgenciaConCobranza
                End If
                sMsg = Replace(sMsg, "[cobranzaenvio]", Format(cCobro, "#,##0"))
            ElseIf IsNull(RsAux("AgeNombre")) Then
                If (bRetiroCambio) Then
                    If (RsAux("EReRemito") = RsAux("EnvDocumento")) Then
                        sMsg = whatsappMsgEnvioRetiroDomicilio(False)
                        bRetiro = True
                    ElseIf Not IsNull(RsAux("EReNota")) Then
                        sMsg = whatsappMsgEnvioConNota(False)
                    Else
                        sMsg = whatsappMsgEnvioCambioXOtroIgual(False)
                    End If
                Else
                    sMsg = whatsappMsgEnvioSinAgencia
                End If
            Else
                sMsg = whatsappMsgEnvioConAgencia
                sMsg = Replace(sMsg, "[nombreagencia]", Trim(RsAux("AgeNombre")))
            End If
            celular = ""
            If (celTesting = "") Then
                If Not IsNull(RsAux("EnvTelefono")) Then
                    If Left(Trim(RsAux("EnvTelefono")), 2) = "09" Then
                        celular = Trim(RsAux("EnvTelefono"))
                    End If
                End If
                If (celular = "") Then
                    If Not IsNull(RsAux("EnvTelefono2")) Then
                        If Left(Trim(RsAux("EnvTelefono2")), 2) = "09" Then
                            celular = Trim(RsAux("EnvTelefono2"))
                        End If
                    End If
                End If
            Else
                celular = celTesting
            End If
            celular = Replace(celular, " ", "")
            
            sMsg = Replace(sMsg, "[diaEnvio]", "hoy " & Format(RsAux("EnvFechaPrometida"), "dddd d"))
            
            sArticulos = Trim(RsAux("Domicilio"))
            If Not IsNull(RsAux("DirLetra")) Then sArticulos = sArticulos & " " + Trim(RsAux("DirLetra"))
            If Not IsNull(RsAux("DirApartamento")) Then sArticulos = sArticulos & " / " + Trim(RsAux("DirApartamento"))
            sMsg = Replace(sMsg, "[callePuertaDireccion]", sArticulos)
            If (RsAux("DepCodigo") = 1) Then
                sMsg = Replace(sMsg, "[localidaddeptoDireccion]", Trim(RsAux("ZonNombre")) & ", " & Trim(RsAux("Depto")))
            Else
                sMsg = Replace(sMsg, "[localidaddeptoDireccion]", Trim(RsAux("Depto")))
            End If
            
            sArticulos = ""
            If (Not IsNull(RsAux("Esq1")) And Not IsNull(RsAux("Esq2"))) Then
                sArticulos = " entre " & Trim(RsAux("Esq1")) & " y " & Trim(RsAux("Esq2"))
            ElseIf (Not IsNull(RsAux("Esq1"))) Then
                sArticulos = " y " & Trim(RsAux("Esq1"))
            ElseIf (Not IsNull(RsAux("Esq2"))) Then
                sArticulos = " y " & Trim(RsAux("Esq2"))
            End If
            sMsg = Replace(sMsg, "[entreesquinasDireccion]", sArticulos)
            
            If RsAux("EnvTipoFlete") = 4 Or RsAux("EnvTipoFlete") = 12 Or RsAux("EnvTipoFlete") = 17 Or RsAux("EnvTipoFlete") = 18 Or RsAux("EnvTipoFlete") = 24 Then
                sMsg = Replace(sMsg, "[desembalarmercaderia]", "")
            Else
                sMsg = Replace(sMsg, "[desembalarmercaderia]", " que desembale la mercadería")
            End If
            
        'horario.
            sArticulos = Trim(Mid(RsAux("EnvRangoHora"), 1, InStr(1, RsAux("EnvRangoHora"), "-") - 1))
            If IsNumeric(sArticulos) Then
                horaI = CInt(sArticulos)
            Else
                horaI = 800
            End If
            
            If horaI < 1000 Then
                sMsg = Replace(sMsg, "[horaEnvioInit]", Format(horaI, "0:00"))
            Else
                sMsg = Replace(sMsg, "[horaEnvioInit]", Format(horaI, "00:00"))
            End If
            
            If (horaI >= 800) Then
                If (horaI >= 1200) Then
                    Fecha = Format(RsAux("EnvFechaPrometida"), "dd/MM/yyyy") & " " & Format(1200, "00:00")
                    horaF = Format(RsAux("EnvFechaPrometida"), "dd/MM/yyyy") & " " & Format(1500, "00:00")
                Else
                    Fecha = Format(RsAux("EnvFechaPrometida"), "dd/MM/yyyy") & " " & Format(800, "00:00")
                    horaF = Format(RsAux("EnvFechaPrometida"), "dd/MM/yyyy") & " " & Format(1000, "00:00")
                End If
            Else
                Fecha = Format(RsAux("EnvFechaPrometida"), "dd/MM/yyyy") & " " & Format(800, "00:00")
                horaF = Format(RsAux("EnvFechaPrometida"), "dd/MM/yyyy") & " " & Format(1000, "00:00")
            End If
            
            sArticulos = Trim(Mid(RsAux("EnvRangoHora"), InStr(1, RsAux("EnvRangoHora"), "-") + 1))
            If IsNumeric(sArticulos) Then
                horaI = CInt(sArticulos)
            Else
                horaI = 2000
            End If
            If horaI < 1000 Then
                sMsg = Replace(sMsg, "[horaEnvioEnd]", Format(horaI, "0:00"))
            Else
                sMsg = Replace(sMsg, "[horaEnvioEnd]", Format(horaI, "00:00"))
            End If
            'End horario
            sArticulos = ""
            
        End If
        
        If Not IsNull(RsAux("EReNota")) Then
            If sArtsNota = "" Then sArtsNota = CargarArticulosRetira(RsAux("EReRemito"))
            If sArticulos <> "" Then sArticulos = sArticulos & ", "
                sArticulos = sArticulos & RsAux("REvCantidad") & " " + Trim(RsAux("Articulo"))
        Else
            If qArts < 3 Then
                If sArticulos <> "" Then sArticulos = sArticulos & vbCrLf
                sArticulos = sArticulos & RsAux("REvCantidad") & " " + Trim(RsAux("Articulo"))
            ElseIf qArts >= 3 And qArts < 4 Then
                sArticulos = sArticulos & vbCrLf & " y otros ..."
            End If
            qArts = qArts + 1
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If (idIx <> 0 And celular <> "") Then
        sMsg = Replace(sMsg, "[tablaArticulosEnvio]", sArticulos)
        sMsg = Replace(sMsg, "[tablaARetirar]", sArtsNota)
        whatsappSaveMsg idEnv, sMsg, Fecha, horaF, idCli, celular
    End If
    cBase.CommitTrans
    Exit Sub
    
ErrBT:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al bloquear los msgs de whatsapp.", Err.Description
    Exit Sub

errRB:
    Resume errVT
    Exit Sub
    
errVT:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al grabar los msgs de whatsapp.", Err.Description
End Sub

Private Function CargarArticulosRetira(ByVal idDoc As Long) As String
Dim sQy As String
Dim rsN As rdoResultset
    CargarArticulosRetira = ""
    sQy = "SELECT * FROM Renglon INNER JOIN Articulos ON ArtID = RenArticulo WHERE RenDocumento = " & idDoc
    Set rsN = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsN.EOF
        If (CargarArticulosRetira <> "") Then CargarArticulosRetira = CargarArticulosRetira & ", "
        CargarArticulosRetira = CargarArticulosRetira & rsN("RenCantidad") & " " & Trim(rsN("ArtNombre"))
        rsN.MoveNext
    Loop
    rsN.Close
End Function

Private Sub GraboSeteoBandejaCopiaeTicket(ByVal nroB As Integer)
    SaveSetting App.Title, "Settings", settingCopiaeTicket, nroB
    printBandejaCopiaeTicket = LeoSeteoBandejaCopiaeTicket()
End Sub

Private Function LeoSeteoBandejaCopiaeTicket() As Integer
    LeoSeteoBandejaCopiaeTicket = GetSetting(App.Title, "Settings", settingCopiaeTicket, printBandejaCopiaeTicket)
End Function

Private Function ValidarCopiaInsertada(ByVal idDoc As Long, ByVal idEnvio As Long) As Boolean
    Dim oCopia As clsDocToPrint
    For Each oCopia In docCopia
        If oCopia.Documento = idDoc And oCopia.idEnvio = idEnvio Then
            ValidarCopiaInsertada = True
            Exit Function
        End If
    Next
End Function

Private Sub HabilitarCamion(ByVal habilitar As Boolean)
On Error GoTo errHC
    Cons = "SELECT CamCodigo, CamNombre, CamTelefono FROM Camion WHERE CamHabilitado = " & IIf(habilitar, 0, 1) & " ORDER BY CamNombre"
    Dim objH As New clsListadeAyuda
    If objH.ActivarAyuda(cBase, Cons, 6000, 1, "Seleccione el camión a " & IIf(habilitar, "habilitar", "deshabilitar")) > 0 Then
        Screen.MousePointer = 11
        If MsgBox("¿Confirma " & IIf(habilitar, "habilitar", "deshabiltar") & " el camión '" & Trim(objH.RetornoDatoSeleccionado(1)) & "'?", vbQuestion + vbYesNo + vbDefaultButton2, "Habilitar/Deshabilitar") = vbYes Then
            Cons = "Select * From Local Where LocCodigo = " & objH.RetornoDatoSeleccionado(0)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            RsAux.Edit
            RsAux("LocHabilitado") = habilitar
            RsAux.Update
            RsAux.Close
        End If
    End If
    Set objH = Nothing
    Screen.MousePointer = 0
    Exit Sub
errHC:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al habilitar/deshabilitar el camión.", Err.Description, "Habilitar/deshabilitar camión"
End Sub

Public Function fnc_PrintDocumento(ByVal iDoc As Long, Optional ByVal codigoImpresion As Long = 0) As Boolean
On Error GoTo errMPD
    vsPrint.Header = ""
    vsPrint.Footer = ""
    vsPrint.Orientation = orPortrait
    Dim oPrint As New clsPrintManager
    With oPrint
        If (codigoImpresion > 0) Then
            .SetDevice paPrintConfD, paPrintConfB, paPrintConfPaperSize
        Else
            .SetDevice paIContadoN, paIContadoB, paPrintCtdoPaperSize
        End If
        If .LoadFileData(gPathListados & "\rptRemitoEnvio.txt") Then
            Dim sQy As String
            If (codigoImpresion > 0) Then
                sQy = "Exec prg_DistribuirEnvio_PrintRemitoCtdo " & iDoc & ", " & codigoImpresion
            Else
                sQy = "Exec prg_DistribuirEnvio_PrintRemitoCtdo " & iDoc
            End If
            fnc_PrintDocumento = .PrintDocumento(sQy, vsPrint, (codigoImpresion > 0))
        End If
    End With
    Set oPrint = Nothing
    Exit Function
errMPD:
    objGral.OcurrioError "Error al imprimir el documento de código: " & iDoc, Err.Description, "Impresión de documentos"
End Function
Private Function db_CambioEstadoEnvioVaCon(ByVal iEstado As EstadoEnvio, ByVal lCamion As Long, ByVal iIDVaCon As Long, Optional sFecha As String) As Boolean
Dim rsE As rdoResultset
'Esta rutina la utilizo para pasar un envío a Asignar ó a Imprimir, le puedo asignar o no el camión
    db_CambioEstadoEnvioVaCon = False
    FechaDelServidor
    On Error GoTo ErrBT
    Cons = "Update Envio Set EnvCamion = " & IIf(lCamion > 0, lCamion, "Null") & _
                    ", EnvEstado = " & iEstado & ", EnvFModificacion = '" & Format(gFechaServidor, cte_FormatFH) & "'"
    If IsDate(sFecha) Then
        Cons = Cons & ", EnvFechaPrometida = '" & Format(sFecha, "yyyy/mm/dd") & "'"
    End If
    Cons = Cons & " FROM Envio, EnvioVaCon WHERE EVCID = " & iIDVaCon & " And EVCEnvio = EnvCodigo"
    cBase.Execute Cons
    db_CambioEstadoEnvioVaCon = True
    Screen.MousePointer = 0
    Exit Function
ErrBT:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al grabar el estado de un va con.", Err.Description
    Exit Function
End Function

Private Sub loc_AppReclamar(Optional iCamion As Integer = 0)
On Error Resume Next
    RunApp App.Path & "\MReclamarCamion.exe", IIf(iCamion > 0, iCamion, "")
End Sub

Private Sub loc_ShowDividirEnvio()
On Error GoTo errDE
    Dim oFrm As New frmDividoEnvio
    oFrm.prmEnvio = vsGrid.Cell(flexcpValue, vsGrid.Row, 0)
    oFrm.Show vbModal, Me
    Set oFrm = Nothing
    db_FillGridDatosCodImpresion True
    Exit Sub
errDE:
    objGral.OcurrioError "Error al acceder a dividir el envío.", Err.Description, "Dividir envío"
End Sub

Private Sub db_RecepcionGraboEntregado(ByVal bEsTodo As Boolean, ByVal idEnvio As Long, ByVal CargarGrilla As Boolean)
On Error GoTo errRGE
    
    Screen.MousePointer = 11
    If bEsTodo Then
        Cons = "EXEC prg_RecepcionEnvio_DoyPorEntregadoEnvio " & Val(tCodigo.Tag) & ", 0, " & paTipoArticuloServicio & ", " & paCodigoDeUsuario & ", " & paCodigoDeSucursal & ", " & paCodigoDeTerminal
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux(0) = -1 Then
            MsgBox "Error al dar los envíos como entregados, refresque el código de impresión y reintente." & vbCrLf & vbCrLf & RsAux(1), vbExclamation, "Atención"
        End If
        RsAux.Close
        tCodigo.Text = ""
        txtEnvioEntregado.Text = ""
        tCodigo.SetFocus
    Else
    
        Cons = "SELECT EnvCodigo FROM Envio " _
            & "INNER JOIN EnviosRemitos ON EReEnvio = EnvCodigo " _
            & "INNER JOIN Renglon ON EReRemito = RenDocumento AND RenARetirar > 0 " _
            & "WHERE EnvCodigo = " & idEnvio
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then
            MsgBox "Envío de retiro de mercadería, la mercadería no ingresó al local debe controlar en mercadería a reclamar.", vbInformation, "ATENCIÓN"
        End If
        RsAux.Close
        
        'Ejecuto storeprocedure para que de cumplido sólo el envío.
        Cons = "EXEC prg_RecepcionEnvio_DoyPorEntregadoEnvio " & Val(tCodigo.Tag) & ", " & idEnvio & ", " & paTipoArticuloServicio & ", " & paCodigoDeUsuario & ", " & paCodigoDeSucursal & ", " & paCodigoDeTerminal
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux(0) = -1 Then
            MsgBox "Error al dar como entregado el envío, refresque el código de impresión y reintente." & vbCrLf & vbCrLf & "Detalle: " & RsAux(1), vbCritical, "Atención"
            CargarGrilla = True
        End If
        RsAux.Close
        
        txtEnvioEntregado.Text = ""
        If CargarGrilla Then
            s_FillGrid
        Else
            'Recorro la grilla y lo elimino a mano.
            Dim iR As Integer
            For iR = 1 To vsGrid.Rows - 1
                If vsGrid.Cell(flexcpValue, iR, 0) = idEnvio Then
                    vsGrid.RemoveItem iR
                    Exit For
                End If
            Next
        End If
        
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errRGE:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al intentar grabar como entregado.", Err.Description, "Dar por entregado"
    On Error Resume Next
    txtEnvioEntregado.Enabled = True
End Sub

Private Sub loc_EliminoVaCon()
On Error GoTo errDV
    
    If vsGrid.Row < 1 Then Exit Sub
       
    
    If MsgBox("Atención un va con está en un remito si ud. elimina el va con el remito será anulado automáticamente." & vbCrLf & vbCrLf & "¿Confirma desvincular todos los envíos del va con?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar va con") = vbYes Then
        
        Dim rsEnvio As rdoResultset
        Dim lOld As Long, iQ As Integer
    
        Screen.MousePointer = 11
        Cons = "Select EnvFModificacion " & _
            " From Envio Where EnvCodigo = " & Val(vsGrid.Cell(flexcpValue, vsGrid.Row, 0))
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux("EnvFModificacion") <> CDate(vsGrid.Cell(flexcpData, vsGrid.Row, 1)) Then
            RsAux.Close
            MsgBox "Atención el envío fue modificado por otra terminal, se refrescará la información.", vbExclamation, "Atención"
            s_FillGrid
            Screen.MousePointer = 0
            Exit Sub
        End If
        RsAux.Close
        
        Cons = "EXEC prg_Envio_VaCon 2, " & Val(vsGrid.Cell(flexcpData, vsGrid.Row, 2)) & ", 0, " & paCodigoDeUsuario
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux(0) = -1 Then
            'Error
            MsgBox "No se logró independizar el Va Con." & vbCrLf & vbCrLf & "Detalle:" & RsAux(1), vbCritical, "Error"
        End If
        RsAux.Close
        s_FillGrid
    End If
    Screen.MousePointer = 0
Exit Sub
errDV:
    objGral.OcurrioError "Error al intentar desvincular el envío.", Err.Description
End Sub

Private Sub db_FillGridDatosCodImpresion(Optional ByVal bReload As Boolean = False)
On Error GoTo errGDR
Dim QTotal As Integer, QCamion As Integer
Dim lLastID As Long, lLastVC As Long
Dim sFM As String
Dim totalBultos As Integer

    'Busco los datos de la tabla repartoimpresión.
    Screen.MousePointer = 11
    loc_CleanRecepcion
    
    'El dato puede ser un envìo o un código de impresión.
    
    Cons = "Select IsNull(EnvCodImpresion, 0) EnvCodImpresion, EnvCodigo, EnvFModificacion, EnvTipo, IsNull(EVCID, 0) as VC, " _
                & " EnvCamion, LocCodigo, LocNombre, CamNombre, CalNombre, DirPuerta, DirLetra, DirApartamento, " _
                & " IsNull(EnvComentario, '') as Memo, REvCantidad, IsNull(AEsID, ArtCodigo) as ArtCodigo, IsNull(AEsNombre, ArtNombre) as AD, ArtTipo, IsNull(AEsID, ArtID) ArtID, DocTipo " _
        & " From ((((((((Envio LEFT OUTER JOIN EnvioVaCon ON EnvCodigo = EVCEnvio) INNER JOIN Direccion ON EnvDireccion = DirCodigo) " _
        & " INNER JOIN Calle ON DirCalle = CalCodigo) INNER JOIN Localidad ON CalLocalidad = LocCodigo) INNER JOIN RenglonEnvio ON EnvCodigo = REvEnvio) " _
        & " INNER JOIN Articulo ON REvArticulo = ArtID) INNER JOIN Camion ON EnvCamion = CamCodigo)" _
        & " LEFT OUTER JOIN ArticuloEspecifico ON AEsTipoDocumento IN (1, 6) AND AEsDocumento = EnvDocumento AND AEsArticulo = REvArticulo) LEFT OUTER JOIN Documento ON DocCodigo = EnvDocumento AND EnvTipo = 1 " _
        & " WHERE ((EnvCodImpresion = " & Val(tCodigo.Text) & " Or RevCodImpresion = " & Val(tCodigo.Text) & ")" _
            & " OR EnvCodImpresion = (SELECT EnvCodImpresion FROM Envio WHERE EnvCodigo = " & Val(tCodigo.Text) & " AND EnvEstado = 3 AND EnvTipo In (1, 2) AND EnvDocumento > 0))" _
        & " AND EnvEstado = 3 AND EnvTipo In (1, 2) And EnvDocumento > 0" _
        & " ORDER BY EVCID, EnvCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        If Not bReload Then MsgBox "No se encontraron envíos para el código ingresado.", vbExclamation, "Buscar código de impresión"
        Screen.MousePointer = 0
        Exit Sub
    Else
        lCamion.Caption = Trim(RsAux!CamNombre)
        lCamion.Tag = RsAux!EnvCamion
        If RsAux("EnvCodImpresion") > 0 Then tCodigo.Text = RsAux("EnvCodImpresion")
        tCodigo.Tag = Val(tCodigo.Text)
        vsGrid.Redraw = False
        Do While Not RsAux.EOF
            With vsGrid
                If lLastID <> RsAux("EnvCodigo") Then
                If RsAux("EnvCodigo") = 1074923 Then MsgBox "H"
                
                    lLastID = RsAux("EnvCodigo")
                    If lLastVC <> RsAux("VC") Or RsAux("VC") = 0 Then
                        lLastVC = RsAux("VC")
                        .AddItem RsAux!EnvCodigo
                        .Cell(flexcpText, .Rows - 1, 5) = f_GetDireccionRsAux
                        .Cell(flexcpText, .Rows - 1, 6) = Trim(RsAux!Memo)
                        
                        'DATA
                        sFM = RsAux!EnvFModificacion: .Cell(flexcpData, .Rows - 1, 1) = sFM         'F Modificado
                        sFM = RsAux!VC: .Cell(flexcpData, .Rows - 1, 2) = Val(sFM)                      'Va Con
                        
                        If RsAux("EnvTipo") = 2 Then .Cell(flexcpForeColor, .Rows - 1, 0) = &H8000&
                    End If
                End If
                ''..........................ARTICULOS
                loc_AgregoEnColleccionBultos RsAux("ArtID"), Format(RsAux("ArtCodigo"), "(#,000,000)") & " " & RsAux("AD"), RsAux("REvCantidad")
                loc_AddArtNombre RsAux("ArtID"), Format(RsAux("ArtCodigo"), "(#,000,000)") & " " & RsAux("AD")
                loc_SetQTipoArt .Rows - 1, RsAux("ArtID"), RsAux("ArtTipo"), RsAux("REvCantidad")
                totalBultos = totalBultos + RsAux("REvCantidad")
                .Cell(flexcpText, .Rows - 1, 7) = fnc_GetStringArticulos(.Cell(flexcpData, .Rows - 1, 6))
                '..........................Artículos
                
                If Not IsNull(RsAux("DocTipo")) Then
                    If RsAux("DocTipo") = TD_Contado Then
                        .Cell(flexcpText, .Rows - 1, 8) = "Contado"
                    ElseIf RsAux("DocTipo") = TD_Credito Then
                        .Cell(flexcpText, .Rows - 1, 8) = "Crédito"
                    ElseIf RsAux("DocTipo") = 47 Then
                        .Cell(flexcpText, .Rows - 1, 8) = "Cambio"
                    ElseIf RsAux("DocTipo") = 48 Then
                        .Cell(flexcpText, .Rows - 1, 8) = "Retiro"
                    End If
                    .Cell(flexcpData, .Rows - 1, 8) = CStr(RsAux("DocTipo"))
                Else
                    .Cell(flexcpText, .Rows - 1, 8) = "Venta S/F"
                    .Cell(flexcpData, .Rows - 1, 8) = ""
                End If
                
            End With
            RsAux.MoveNext
        Loop
    End If
    RsAux.Close
    vsGrid.Redraw = True
    'Ahora determino si tengo mercadería y saco el nombre del camión
    
    Cons = "Select IsNull(Sum(ReECantidadTotal), 0) as QTotal, IsNull(Sum(ReECantidadEntregada), 0) as QCamion" & _
            " From RenglonEntrega Where ReECodImpresion = " & tCodigo.Text
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        QTotal = RsAux("QTotal")
        QCamion = RsAux("QCamion")
    End If
    RsAux.Close
    
    tooMenu.Buttons("save").Enabled = (QTotal = QCamion)
    If QTotal > QCamion Then
        If Not bReload Then MsgBox "Al camión no se le entregó " & IIf(QCamion > 0, "la totalidad de ", "") & "la mercadería, no se podrá dar todo como entregado.", vbExclamation, "Atención"
    End If
    
    'Busco si el camión tiene mercadería para reclamar.
    Cons = "Select Top 1 MRCCamion From MercaderiaReclamarCamion Where MRCCamion = " & Val(lCamion.Tag) & " And MRCDevuelto Is Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        imgMercaderia.Visible = True
        lbMercaderia.Visible = True
        If Not bReload Then MsgBox "El camionero posee mercadería que debe reclamarle.", vbExclamation, "Mercadería para reclamar"
    End If
    RsAux.Close
    
    sbStatus.Panels("envio").Text = "Envíos: " & vsGrid.Rows - 1
    sbStatus.Panels("bultos").Text = "Bultos: " & totalBultos
    Screen.MousePointer = 0
    Exit Sub
    
errGDR:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar y cargar la información del código de impresión.", Err.Description
End Sub

Private Sub loc_CleanRecepcion()
    lCamion.Caption = ""
    vsGrid.Rows = 1
    tooMenu.Buttons("save").Enabled = False
    imgMercaderia.Visible = False
    lbMercaderia.Visible = False
    sbStatus.Panels("envio").Text = "Envíos: "
End Sub

Private Sub loc_PresentoRecepcion()
Dim iCodImp As Long
    'Si tengo un código en el textbox lo vuelvo a cargar
    'simulo como si fuese un back algo así
    If IsNumeric(tCodigo.Text) Then iCodImp = tCodigo.Text
    If iCodImp > 0 Then db_FillGridDatosCodImpresion True Else loc_CleanRecepcion
End Sub

Private Sub loc_ShowCambiarEnvio(ByVal iCaso As Byte, ByVal iEnvio As Long)
On Error GoTo errSCE
    Dim oFrm As New frmMercaAReclamar
    With oFrm
        .prmInvocacion = iCaso
        .prmEnvio = iEnvio
        .Show vbModal
    End With
    If Val(hlTab(0).Tag) = 2 Then
        s_VerficoEnvioEditado iCamSelect, Impreso
    Else
        db_FillGridDatosCodImpresion True
    End If
    Exit Sub
errSCE:
    objGral.OcurrioError "Error al instanciar el formulario.", Err.Description, "Retornar un envío"
End Sub

Private Sub loc_SetStatusTipoArt()
Dim iQ As Integer, iQ1 As Integer
Dim iTipo As Integer, iArt As Integer
Dim arrArt() As String
Dim arrTipo() As String

    If iIDTipo = 0 And iIDArt = 0 Then Exit Sub
    With vsGrid
        For iQ = 1 To vsGrid.Rows - 1
            arrArt = Split(.Cell(flexcpData, iQ, 6), ";")
            arrTipo = Split(.Cell(flexcpData, iQ, 7), ";")
            
            If iIDTipo > 0 Then
                For iQ1 = 0 To UBound(arrTipo)
                    If InStr(1, arrTipo(iQ1), iIDTipo & ":") = 1 Then
                        arrTipo(iQ1) = Replace(arrTipo(iQ1), iIDTipo & ":", "")
                        iTipo = iTipo + Val(arrTipo(iQ1))
                        Exit For
                    End If
                Next
            End If
            
            If iIDArt > 0 Then
                For iQ1 = 0 To UBound(arrArt)
                    If InStr(1, arrArt(iQ1), iIDArt & ":") = 1 Then
                        arrArt(iQ1) = Replace(arrArt(iQ1), iIDArt & ":", "")
                        iArt = iArt + Val(arrArt(iQ1))
                        Exit For
                    End If
                Next
            End If
            
        Next
    End With
    
    If iIDTipo > 0 Then sbStatus.Panels("tipo").Text = sNTipo & ": " & iTipo
    If iIDArt > 0 Then sbStatus.Panels("articulo").Text = sCodArt & ": " & iArt
    
End Sub

Private Sub loc_AgregoEnColleccionBultos(ByVal idArt As Long, ByVal Nombre As String, ByVal cantidad As Integer)
Dim iQ As Integer
Dim oArtM As clsArticuloMenu

    For Each oArtM In colBultos
        If oArtM.ArticuloCodigo = idArt Then
            oArtM.cantidad = oArtM.cantidad + cantidad
            Exit Sub
        End If
    Next
    Set oArtM = New clsArticuloMenu
    oArtM.ArticuloCodigo = idArt
    oArtM.ArticuloNombre = Nombre
    oArtM.cantidad = cantidad
    colBultos.Add oArtM
End Sub


Private Sub loc_SetQTipoArt(ByVal iRow As Long, ByVal idArt As Long, ByVal idTipo As Long, ByVal iCant As Integer)
Dim arrArt() As String
Dim arrTipo() As String
Dim iQ As Integer, bIns As Boolean
    
    With vsGrid
        arrArt = Split(.Cell(flexcpData, iRow, 6), ";")
        arrTipo = Split(.Cell(flexcpData, iRow, 7), ";")
            
        If UBound(arrArt) = -1 Then
            ReDim arrArt(0)
            arrArt(0) = idArt & ":" & iCant
        Else
            bIns = True
            For iQ = 0 To UBound(arrArt)
                If InStr(1, arrArt(iQ), idArt & ":") = 1 Then
                    arrArt(iQ) = Replace(arrArt(iQ), idArt & ":", "")
                    arrArt(iQ) = idArt & ":" & Val(arrArt(iQ)) + iCant
                    bIns = False
                End If
            Next
            If bIns Then
                ReDim Preserve arrArt(UBound(arrArt) + 1)
                arrArt(UBound(arrArt)) = idArt & ":" & iCant
            End If
        End If
        
        If UBound(arrTipo) = -1 Then
            ReDim arrTipo(0)
            arrTipo(0) = idTipo & ":" & iCant
        Else
            bIns = True
            For iQ = 0 To UBound(arrTipo)
                If InStr(1, arrTipo(iQ), idTipo & ":") = 1 Then
                    arrTipo(iQ) = Replace(arrTipo(iQ), idTipo & ":", "")
                    arrTipo(iQ) = idTipo & ":" & Val(arrTipo(iQ)) + iCant
                End If
            Next
            If bIns Then
                ReDim Preserve arrTipo(UBound(arrTipo) + 1)
                arrTipo(UBound(arrTipo)) = idTipo & ":" & iCant
            End If
        End If
        
        .Cell(flexcpData, iRow, 6) = Join(arrArt, ";")
        .Cell(flexcpData, iRow, 7) = Join(arrTipo, ";")
    End With

End Sub

Private Sub loc_QEnviosClic()
Dim iQ As Integer, iSuma As Integer
    For iQ = vsGrid.FixedRows To vsGrid.Rows - 1
        If vsGrid.Cell(flexcpChecked, iQ, 0) = flexChecked Then iSuma = iSuma + 1
    Next
    sbStatus.Panels("envio").Text = "Seleccionados " & iSuma
End Sub

Private Sub loc_DatosRegistry()

    sNTipo = "": iIDTipo = 0
    sCodArt = "": iIDArt = 0

'Guardo todos para no tener que consultar aca.
    sNTipo = GetSetting(App.Title, "Settings", "AA" & Me.Name & "NTipo", "")
    iIDTipo = Val(GetSetting(App.Title, "Settings", "AA" & Me.Name & "IDTipo", "0"))
    
    sCodArt = Trim(GetSetting(App.Title, "Settings", "AA" & Me.Name & "ArtDesc", ""))
    iIDArt = Val(GetSetting(App.Title, "Settings", "AA" & Me.Name & "IDArt", "0"))
    
    sbStatus.Panels("tipo").Text = IIf(sNTipo <> "", sNTipo, "Sin tipo asignado")
    sbStatus.Panels("articulo").Text = IIf(sCodArt <> "", sCodArt, "Sin artículo asignado")
    
End Sub

Private Function BuscoUsuarioDigito(Digito As Long, Optional Codigo As Boolean = False, Optional Identificacion As Boolean = False, Optional Iniciales As Boolean = False) As Variant
Dim RsUsr As rdoResultset
Dim aRetorno As Variant
On Error GoTo ErrBUD
    Screen.MousePointer = 11
    Cons = "Select * from Usuario Where UsuDigito = " & Digito
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Codigo Then aRetorno = RsUsr!UsuCodigo
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    BuscoUsuarioDigito = aRetorno
    Screen.MousePointer = 0
    Exit Function
ErrBUD:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar el usuario.", Err.Description
End Function


Private Sub UpdateInstalacionesADocumento(ByVal lNroD As Long, ByVal lNroVT As Long)
Dim rsI As rdoResultset
    Cons = "Select * From Instalacion Where InsTipoDocumento = 2 And InsDocumento = " & lNroVT _
        & " And InsAnulada Is Null"
    Set rsI = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsI.EOF
        rsI.Edit
        rsI!InsTipoDocumento = 1
        rsI!InsDocumento = lNroD
        rsI.Update
        rsI.MoveNext
    Loop
    rsI.Close
End Sub

Private Sub loc_ImprimoReparto(ByVal lIDImp As Long)
'Separo los listados en 4:
    '1) Tipo = 1 y Sin Agencia
    '2) Tipo = 1 y Agencia
    '3) Tipo = 2 y Sin Agencia
    '4) Tipo = 2 y Con Agencia
   
   loc_ImprimoRemitosEnvios lIDImp
   
   
   loc_ImprimoRepartoSegunTipo lIDImp, 1, ETDIP_SinAgenciaComun
   loc_ImprimoRepartoSegunTipo lIDImp, 1, ETDIP_SinAgenciaEspecial
   loc_ImprimoRepartoSegunTipo lIDImp, 1, ETDIP_EnviosRemitos
   loc_ImprimoRepartoSegunTipo lIDImp, 1, ETDIP_ConAgencia
   
   loc_ImprimoRepartoSegunTipo lIDImp, 2, ETDIP_SinAgenciaComun
   loc_ImprimoRepartoSegunTipo lIDImp, 2, ETDIP_SinAgenciaEspecial
   loc_ImprimoRepartoSegunTipo lIDImp, 2, ETDIP_ConAgencia
   
   loc_ImprimoHojaArticulos lIDImp
   
End Sub

Private Sub loc_PrintArts(ByVal sArts As String, ByVal lWidthCol As Single)
Dim arrArts() As String
Dim iQ As Integer, iJ As Byte
    bPrintColEsArt = True
    iJ = 0
    arrArts = Split(sArts, "|")
    sArts = ""
    For iQ = 0 To UBound(arrArts)
        iJ = 1 + iJ
        sArts = sArts & IIf(sArts = "", "", "|") & arrArts(iQ)
        If iJ = 3 Then
            vsPrint.AddTable ">" & CInt(lWidthCol / 3) & "|>" & CInt(lWidthCol / 3) & "|>" & CInt(lWidthCol / 3), "", sArts
            iJ = 0
            sArts = ""
        End If
    Next
    If iJ > 0 Then
        If iJ = 1 Then
            sArts = sArts & "||"
        ElseIf iJ = 2 Then
            sArts = sArts & "|"
        End If
        vsPrint.AddTable ">" & CInt(lWidthCol / 3) & "|>" & CInt(lWidthCol / 3) & "|>" & CInt(lWidthCol / 3), "", sArts
    End If
    bPrintColEsArt = False
End Sub

Private Sub loc_TxtPieEnvio(ByVal sTracking As String)
    If sTracking = "" Then
        vsPrint.AddTable "<7500|<2300|<2300|<1650", "", "Observación:|Recibió:|Revisado:|Hora:"
    Else
        vsPrint.AddTable "<7700|<6050", "", "Observación:|" & sTracking
    End If
End Sub

Private Sub loc_ImprimoRepartoSegunTipo(ByVal lIDImp As Long, ByVal iTipo As Integer, ByVal TipoImpresion As TiposDeImpresionPlanillas)
Dim rsE As rdoResultset, rsA As rdoResultset
Dim lCodAnt As Long, lWidth As Long
Dim sZonAge As String
Dim cCobrar As Currency, cFlete As Currency, cPiso As Currency
Dim sArt As String

    On Error GoTo errPR
    bPrintPlanilla = True
    Screen.MousePointer = 11
    
    Dim strVaCon As String
    Cons = "SELECT Envio.*, rTrim(CamNombre) as CamN, rTrim(isNull(AgeNombre, ZonNombre)) as ZAN, IsNull(EVCID, 0) as VC" & _
                ", (rTrim(CPeApellido1) + ', ' + rTrim(CPeNombre1)) as NP, IsNull(CEmNombre, CEmFantasia) as NE, rTrim(DocSerie) + '-' + RTrim(Convert(char(7), DocNumero)) as DocSerieNro, IsNull(DocComentario, '') DocComentario " & _
                "FROM (((((Envio" & _
                        " LEFT OUTER JOIN Documento ON EnvDocumento = DocCodigo)" & _
                        " LEFT OUTER JOIN CPersona ON EnvCliente = CPeCliente)" & _
                        " LEFT OUTER JOIN CEmpresa ON EnvCliente = CEmCliente)" & _
                        " LEFT OUTER JOIN Agencia ON EnvAgencia = AgeCodigo)" & _
                        " LEFT OUTER JOIN EnvioVaCon ON EnvCodigo = EVCEnvio)" _
        & ", Camion, Zona" _
        & " WHERE EnvCodImpresion = " & lIDImp _
        & " And EnvZona = ZonCodigo And EnvCamion = CamCodigo " _
        & " And EnvTipo = " & iTipo
        
    Select Case TipoImpresion
        Case ETDIP_ConAgencia
            Cons = Cons & " AND IsNULL(EnvAgencia, 0) > 0 "
        Case ETDIP_SinAgenciaComun
            Cons = Cons & " AND IsNULL(EnvAgencia, 0) = 0 AND EnvHoraEspecial IS NULL "
        Case ETDIP_SinAgenciaEspecial
            Cons = Cons & " AND IsNULL(EnvAgencia, 0) = 0 AND EnvHoraEspecial IS NOT NULL "
    End Select
    
    If TipoImpresion = ETDIP_EnviosRemitos Then
        Cons = Cons & " AND EnvCodigo IN (SELECT EReEnvio FROM EnviosRemitos) "
    ElseIf TipoImpresion <> ETDIP_ConAgencia Then
        Cons = Cons & " AND EnvCodigo NOT IN (SELECT EReEnvio FROM EnviosRemitos) "
    End If
    
    Cons = Cons & " And EnvDocumento > 0" _
        & " Order by ZAN, EVCID"
        
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    If Not rsE.EOF Then
        loc_StartCtrlPrint True
        
        With vsPrint
            lWidth = .PageWidth - .MarginLeft - .MarginRight
            
            .MarginTop = 250
            .MarginLeft = 0
            .MarginBottom = 300
            .MarginRight = 550
            .MarginBottom = 400
            .MarginTop = 400
            .HdrFontName = "Tahoma"
            .HdrFontBold = True
            .HdrFontSize = "9"
            .Header = "Camión: " & rsE("CamN") & IIf(TipoImpresion = 3, " - FLETES ESPECIALES", IIf(TipoImpresion = ETDIP_EnviosRemitos, "  - CAMBIOS/DEVOLUCIÓN", "")) & "|Fecha: " & Format(rsE("EnvFechaPrometida"), "dd/mm/yy") & "|Código: " & lIDImp
            .Footer = Now & "||Pag.: %d"
        End With
    Dim sTracking As String
        'Agrupo por Código de Envío y por Zona o Agencia.
        Do While Not rsE.EOF
        
            If sZonAge <> rsE("ZAN") Then
                
                If lCodAnt > 0 And sArt <> "" Then
                    With vsPrint
                        .TableBorder = tbNone
                        loc_PrintArts sArt, lWidth
                        
                
                        sTracking = vbNullString
                        'Si es tipo agencia --> busco el tracking.
                        If TipoImpresion = ETDIP_ConAgencia Then
                            sTracking = FindTrackingEnvio(lCodAnt) ' rsE("EnvCodigo"))
                        End If
                        
                        .TableBorder = tbBottom
                        loc_TxtPieEnvio sTracking
                        .TableBorder = tbNone
                        .Paragraph = ""
                    End With
                    sArt = ""
                End If
                
                'Encabezado por Zona o Agencia.
                sZonAge = rsE("ZAN")
                With vsPrint
                    If .CurrentY + (.TextHeight("HQRPQERASDFW") * 3) > .PageHeight - .MarginBottom - .MarginTop Then .NewPage
                    If .CurrentY < .MarginTop Then .CurrentY = .MarginTop + 200
                    .TableBorder = tbTopBottom
                    .AddTable lWidth, "", sZonAge, , &HEFEFEF
                    .TableBorder = tbNone
                End With
                
            End If
                        
            If lCodAnt <> rsE("EnvCodigo") And (InStr(1, strVaCon, "," & rsE("VC") & ",") = 0 Or rsE("VC") = 0) Then
                
                If lCodAnt > 0 And sArt <> "" Then
                    loc_PrintArts sArt, lWidth
                    sTracking = vbNullString
                    'Si es tipo agencia --> busco el tracking.
                    If TipoImpresion = ETDIP_ConAgencia Then
                        sTracking = FindTrackingEnvio(lCodAnt)  'rsE("EnvCodigo"))
                    End If
                    
                    With vsPrint
                        .TableBorder = tbBottom
                        loc_TxtPieEnvio sTracking
                        .TableBorder = tbNone
                        .LineSpacing = 50
                        .Paragraph = ""
                        .LineSpacing = 100
                    End With
                    sArt = ""
                End If
                
                lCodAnt = rsE("EnvCodigo")
                
                cCobrar = 0
                cFlete = 0
                cPiso = 0
                'Sumo el importe de las diferencia de envíos .
                'Ojo acá sumo el importe de todos los envíos que forman el vacon si lo tiene.
                If rsE("VC") > 0 Then
                
                    'Consulto todos los documentos pendientes que tengo para todos los documentos.
                    Cons = "Select Sum(DPeImporte) From DocumentoPendiente " & _
                        " Where DPeIDTipo IN(Select EVCEnvio From EnvioVaCon Where EVCID = " & rsE("VC") & ") And DPeTipo = 1 And DPeIDLiquidacion Is Null"
                    Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If Not rsA.EOF Then
                        If Not IsNull(rsA(0)) Then cCobrar = rsA(0)
                    End If
                    rsA.Close
                    
                    'Busco todos los envíos que tengan como forma de pago paga a camión.
                    Cons = "Select ISNULL(SUM(EnvValorFlete), 0), IsNull(SUM(EnvValorPiso), 0) From Envio Where EnvCodigo IN (Select EVCEnvio From EnvioVaCon Where EVCID = " & rsE("VC") & ") And EnvFormaPago = 3"
                    Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If Not rsA.EOF Then cFlete = cFlete + rsA(0): cPiso = rsA(1)
                    rsA.Close
                    
                Else

                    Cons = "Select Sum(DPeImporte) From DocumentoPendiente Where DPeIDTipo = " & rsE("EnvCodigo") & " And DPeTipo = 1 And DPeIDLiquidacion Is Null"
                    Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If Not rsA.EOF Then
                        If Not IsNull(rsA(0)) Then cCobrar = rsA(0)
                    End If
                    rsA.Close
                    
                    'Paga al camión
                    If rsE("EnvFormaPago") = 3 Then
                        If Not IsNull(rsE("EnvValorFlete")) Then cFlete = rsE("EnvValorFlete")
                    End If
                    If Not IsNull(rsE("EnvValorPiso")) Then cPiso = rsE("EnvValorPiso")
                    
                End If
                
                With vsPrint
                    If .CurrentY + (.TextHeight("HQRPQERASDFW") * 2) > .PageHeight - .MarginBottom - .MarginTop Then .NewPage
                'Código Hora Documento Nombre Dirección Telefono Memo Cobrar
                    .AddTable "<890|<1000|<950|<3100|<4560|<2250|<2500|>2100", "", _
                                            rsE("EnvCodigo") _
                                            & "|" & IIf(IsNull(rsE("EnvRangoHora")), "", Trim(rsE("EnvRangoHora"))) _
                                            & "|" & rsE("DocSerieNro") _
                                            & "|" & IIf(Not IsNull(rsE("NE")), rsE("NE"), rsE("NP")) _
                                            & "|" & objGral.ArmoDireccionEnTexto(cBase, rsE("EnvDireccion"), True, True, False, True, True, False, False) _
                                            & "|" & IIf(IsNull(rsE("EnvTelefono")), "", Trim(rsE("EnvTelefono"))) & IIf(IsNull(rsE("EnvTelefono2")), "", Chr(13) & Trim(rsE("EnvTelefono2"))) _
                                            & "|" & "Cobrar: " & Format(cCobrar, "#,##0.00") & Chr(13) & " Flete:" & Format(cFlete, "#,##0.00") & Chr(13) & "Piso:" & Format(cPiso, "#,##0.00")
                    
                    If Not (IsNull(rsE("EnvComentario")) And Trim(rsE("DocComentario")) = "") Then
                        Dim sMemo As String, sDocMemo As String
                        If Not IsNull(rsE("EnvComentario")) Then sMemo = Replace(Trim(rsE("EnvComentario")), ";", ",") Else sMemo = ""
                        If Not IsNull(rsE("DocComentario")) Then
                            sDocMemo = Trim(rsE("DocComentario"))
                        Else
                            sDocMemo = "."
                        End If
                        
                        .AddTable "890|<6000|<6000", "", _
                                            "Memo|" & sMemo & "|" & sDocMemo
                    End If
                End With
            End If
            If (InStr(1, strVaCon, "," & rsE("VC") & ",") = 0 Or rsE("VC") = 0) Then
                sArt = f_GetArticulosEnvioConNombre(rsE("EnvCodigo"), rsE("VC"))
            End If
            If rsE("VC") > 0 And InStr(1, strVaCon, "," & rsE("VC") & ",") = 0 Then
                If strVaCon = "" Then strVaCon = ","
                strVaCon = strVaCon & rsE("VC") & ","
            End If
            rsE.MoveNext
            
            If rsE.EOF Then
                If lCodAnt > 0 And sArt <> "" Then
                    loc_PrintArts sArt, lWidth
                    
                    sTracking = vbNullString
                    'Si es tipo agencia --> busco el tracking.
                    If TipoImpresion = ETDIP_ConAgencia Then
                        sTracking = FindTrackingEnvio(lCodAnt)
                    End If
                    With vsPrint
                        .TableBorder = tbBottom
                        loc_TxtPieEnvio sTracking
                        .TableBorder = tbNone
                    End With
                    sArt = ""
                End If
            End If
        Loop
        With vsPrint
            .EndDoc
            .PrintDoc
        End With
    End If
    rsE.Close
    bPrintPlanilla = False
Screen.MousePointer = 0
Exit Sub
errPR:
    bPrintPlanilla = False
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al imprimir las hojas " & IIf(iTipo = 2, " de servicios ", " de reparto ") & IIf(TipoImpresion = 1, " para Agencias.", "."), Err.Description, "Imprimir"
End Sub

Private Function FindTrackingEnvio(ByVal idEnvio As String) As String
On Error GoTo errT
FindTrackingEnvio = ""
    Dim rsEB As rdoResultset
    Dim sQry As String
    sQry = "SELECT EBuTrackingAg FROM EtiquetasBultos WHERE EBuEnvio = " & idEnvio & " AND EBuTrackingAg IS NOT NULL GROUP BY EBuTrackingAg"
    Set rsEB = cBase.OpenResultset(sQry, rdOpenDynamic, rdConcurValues)
    Do While Not rsEB.EOF
        FindTrackingEnvio = FindTrackingEnvio & IIf(FindTrackingEnvio <> "", ", ", "") & rsEB(0)
        rsEB.MoveNext
    Loop
    rsEB.Close
    FindTrackingEnvio = "Tracking: " & FindTrackingEnvio
errT:
End Function

Private Function fnc_GetArticulosEnvio(ByVal iEnvio As Long) As String
Dim iQ As Integer
    With vsGrid
        For iQ = .FixedRows To .Rows - 1
            If Val(.Cell(flexcpValue, iQ, 0)) = iEnvio Then
                If Val(hlTab(0).Tag) = 0 Then
                    fnc_GetArticulosEnvio = .Cell(flexcpText, iQ, 8)
                Else
                    fnc_GetArticulosEnvio = .Cell(flexcpText, iQ, 7)
                End If
                Exit Function
            End If
        Next
    End With
End Function

Private Sub loc_ImprimoPlanilla(ByVal Query As String)
'Imprimo lo que tengo en la grilla.
Dim iQ As Integer
Dim rsE As rdoResultset
Dim lCodAnt As Long
Dim sZonAge As String
Dim sArt As String
Dim lWidth As Long
    
    On Error GoTo errPP
    Screen.MousePointer = 11
    bPrintPlanilla = True
    sArt = ""
    lCodAnt = 0

    Set rsE = cBase.OpenResultset(Query, rdOpenDynamic, rdConcurValues)
    
    If Not rsE.EOF Then
    
        loc_StartCtrlPrint True
        
        With vsPrint
            lWidth = .PageWidth - .MarginLeft - .MarginRight
            .HdrFontName = "Tahoma"
            .HdrFontBold = True
            .HdrFontSize = "9"
            .Header = "Reparto Auxiliar||Fecha: " & tFecha.Text
            .Footer = Now
        End With
    
        Do While Not rsE.EOF
        
            If sZonAge <> rsE("ZAN") Then
                If lCodAnt > 0 And sArt <> "" Then
                    With vsPrint
                        .TableBorder = tbBottom
                        .AddTable "1|<" & lWidth - 100, "", "|" & sArt
                        .TableBorder = tbNone
                        .Paragraph = ""
                    End With
                    sArt = ""
                End If
                
                'Encabezado por Zona o Agencia.
                sZonAge = rsE("ZAN")
                With vsPrint
                    If .CurrentY + (.TextHeight("HQRPQERASDFW") * 3) > .PageHeight - .MarginBottom - .MarginTop Then .NewPage
                    .TableBorder = tbTopBottom
                    .AddTable lWidth - 100, "", sZonAge, , &HEFEFEF
                    .TableBorder = tbNone
                End With
            End If
                        
            If lCodAnt <> rsE("EnvCodigo") Then
                If lCodAnt > 0 And sArt <> "" Then
                    With vsPrint
                        .TableBorder = tbBottom
                        .AddTable "1|<" & lWidth - 100, "", "|" & sArt
                        .TableBorder = tbNone
                        .LineSpacing = 50
                        .Paragraph = ""
                        .LineSpacing = 100
                    End With
                    sArt = ""
                End If
                lCodAnt = rsE("EnvCodigo")
                With vsPrint
                    If .CurrentY + (.TextHeight("HQRPQERASDFW") * 2) > .PageHeight - .MarginBottom - .MarginTop Then .NewPage
                'Código Dirección Memo Hora
                    .AddTable ">850|<1000|<6500|<4500", "", Replace(rsE("EnvCodigo") _
                                            & "|" & IIf(IsNull(rsE("RH")), "", Trim(rsE("RH"))) _
                                            & "|" & objGral.ArmoDireccionEnTexto(cBase, rsE("EnvDireccion"), True, True, False, True, True, False, False) _
                                            & "|" & IIf(IsNull(rsE("Memo")), "", Trim(rsE("Memo"))), ";", ",")
                End With
            End If
            sArt = fnc_GetArticulosEnvio(rsE("EnvCodigo"))
            rsE.MoveNext
            If rsE.EOF Then
                If lCodAnt > 0 And sArt <> "" Then
                    With vsPrint
                        .TableBorder = tbBottom
                        .AddTable "1|<" & lWidth - 100, "", "|" & sArt
                        .TableBorder = tbNone
                    End With
                End If
            End If
        Loop
        
        With vsPrint
            .EndDoc
            .PrintDoc
        End With
        
    End If
    rsE.Close
    bPrintPlanilla = False
    Screen.MousePointer = 0
Exit Sub

errPP:
    bPrintPlanilla = False
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al imprimir la planilla auxiliar.", Err.Description, "Error (planilla auxiliar)"
End Sub


Private Sub loc_ImprimoHojaArticulos(ByVal idCodigo As Long)
'Imprimo lo que tengo en la grilla.
Dim iQ As Integer
Dim rsE As rdoResultset
Dim lCodAnt As Long
Dim sZonAge As String
Dim sArt As String
Dim lWidth As Long
    
    On Error GoTo errPP
    Screen.MousePointer = 11
    bPrintPlanilla = True
    sArt = ""
    lCodAnt = 0
    
    Dim Camion As String
    Cons = "SELECT Top 1 ISNULL(CamNombre, '') FROM Envio INNER JOIN Camiones ON EnvCamion = CamID WHERE EnvCodImpresion = " & idCodigo
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Camion = Trim(rsE(0))
    rsE.Close
    
    Cons = "select ReECantidadTotal, ArtCodigo, ArtNombre, ISNULL(LocNombre, '') LocNombre from RenglonEntrega inner join Articulo on ArtId = ReEArticulo " & _
        "left outer join (SELECT ALeLocal, ALEArticulo, MIN(ALEOrden) ORDEN FROM ArticulosLocalesEntrega WHERE ALETipo = 1 GROUP BY ALeLocal, ALEArticulo) as locArt ON locArt.ALEArticulo = ArtId " & _
        "LEFT OUTER JOIN Locales on LocID = locArt.ALELocal " & _
        "WHERE ReECodImpresion = " & idCodigo & " Order by LocNombre"
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Dim locAnt As String
    
    If Not rsE.EOF Then
    
        loc_StartCtrlPrint False
                
        With vsPrint
            lWidth = .PageWidth - .MarginLeft - .MarginRight
            .HdrFontName = "Tahoma"
            .HdrFontBold = True
            .HdrFontSize = "9"
            .MarginLeft = 200
            .MarginTop = 400
            .Header = "Reparto - Artículos|" & Camion & "|Código: " & idCodigo
            .Footer = Now
            
            With vsPrint
                .TableBorder = tbAll
                .AddTable "1000|<1200|4000|2000", "Cant.|Còdigo|Nombre|Local", ""  '"|" & rsE("ReECantidadTotal") & "|" & "|" & rsE("ArtCodigo") & "|" & rsE("ReEArtNombre")
'                .TableBorder = tbNone
            End With
        End With
    
        Do While Not rsE.EOF
            If (locAnt <> Trim(rsE("LocNombre"))) Then
                If locAnt <> "" Then
                    With vsPrint
                        .TableBorder = tbBottom
                        .AddTable "1000|<1200|4000|2000", "", "|" & "|" & "|"
                        .TableBorder = tbNone
                    End With
                End If
                locAnt = Trim(rsE("LocNombre"))
            End If
            
            With vsPrint
                .TableBorder = tbBottom
                .AddTable "1000|<1200|4000|2000", "", rsE("ReECantidadTotal") & "|" & rsE("ArtCodigo") & "|" & Trim(rsE("ArtNombre")) & "|" & Trim(rsE("LocNombre"))
                .TableBorder = tbNone
            End With
            rsE.MoveNext
        Loop
        
        With vsPrint
            .EndDoc
            .PrintDoc
        End With
        
    End If
    rsE.Close
    bPrintPlanilla = False
    Screen.MousePointer = 0
Exit Sub

errPP:
    bPrintPlanilla = False
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al imprimir la planilla auxiliar.", Err.Description, "Error (planilla auxiliar)"
End Sub


Private Sub loc_StartCtrlPrint(ByVal landscape As Boolean)
On Error Resume Next
    SeteoImpresoraPorDefecto paPrintConfD
        
    With vsPrint
        'Le digo que lo haga directo en la impresora.
        .AbortWindow = False

        .PaperBin = paPrintConfB
        .Device = paPrintConfD
        .PaperSize = paPrintConfPaperSize
        
        .Orientation = IIf(landscape, orLandscape, orPortrait)
        .StartDoc
        
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
        .FileName = "Distribuir Envíos"
        .FontSize = 8.25
        .TableBorder = tbNone
    End With
        
End Sub

Private Function fnc_ArmoNombre(Ape1 As String, Ape2 As String, Nom1 As String, Nom2 As String) As String
    fnc_ArmoNombre = Trim(Ape1) & " " & Trim(Ape2) & ", " & Trim(Nom1) & " " & Trim(Nom2)
End Function

Private Sub loc_SaveSettingColGrid()
    SaveSetting App.Title, "Grid" & hlTab(0).Tag, "SizeGrid", vsGrid.Tag
End Sub

Private Sub loc_SetSizeColGrid()
Dim vCol() As String
Dim vSize() As String
Dim iQ As Integer
    With vsGrid
        If .Tag = "" Then Exit Sub
        vCol = Split(.Tag, "|")
        For iQ = 0 To UBound(vCol)
            If InStr(1, vCol(iQ), ":", vbTextCompare) = 0 Then
                .ColWidth(iQ) = vCol(iQ)
            Else
                vSize = Split(vCol(iQ), ":")
                .ColWidth(iQ) = IIf(vSize(1) = 0, vSize(0), vSize(2))
                .Cell(flexcpPicture, 0, iQ) = imgMini.ListImages(vSize(1) + 1).Picture
            End If
        Next
    End With
End Sub

Private Sub act_Save()
    
    'Para grabar que tiene que haber algo en la grilla.
    If vsGrid.Rows = vsGrid.FixedRows Then Exit Sub
    
    Select Case Val(hlTab(0).Tag)
        Case 0
            If MsgBox("Los envíos que no tengan camión sugerido serán asignados automáticamente." & vbCr & vbCr & "¿Confirma asignar los envíos seleccionados?", vbQuestion + vbYesNo, "Asignarle camión al envío") = vbYes Then
                s_SaveAsignoCamionAEnvio
            End If
        
        Case 1
            If Not ValidarVersionEFactura Then
                MsgBox "La versión del componente CGSAEFactura está desactualizado, debe distribuir software." _
                            & vbCrLf & vbCrLf & "Se cancelará la ejecución.", vbCritical, "EFactura"
                End
            End If
            
            If MsgBox("¿Confirma imprimir los envíos de la lista?", vbQuestion + vbYesNo, "Imprimir Reparto") = vbYes Then
                If paCodigoDeSucursal <> 6 Then
                    MsgBox "Su sucursal no admite impresión de envíos, verifique.", vbExclamation, "Posible error"
                    Exit Sub
                End If
                s_SaveImprimoEnvio
            End If
            
        Case 5
            If MsgBox("¿Confirmar pasar el resto de los envíos como entregados?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
                db_RecepcionGraboEntregado True, 0, True
            End If
    End Select
End Sub

Private Function ValidarVersionEFactura() As Boolean
On Error GoTo errEC
    With New clsCGSAEFactura
        ValidarVersionEFactura = .ValidarVersion()
    End With
    Exit Function
errEC:
End Function

Private Sub s_SetMenu()
    With tooMenu
        .Buttons("sugerircamion").Visible = (Val(hlTab(0).Tag) = 0)
        .Buttons("save").Visible = False
        .Buttons("save").Enabled = False
        .Buttons("envio").Enabled = False
        
        Select Case Val(hlTab(0).Tag)
            Case 0
                .Buttons("save").Visible = True
                .Buttons("save").Caption = "Asignar camión"
                .Buttons("save").Enabled = (vsGrid.Rows > vsGrid.FixedRows)
                .Buttons("sugerircamion").Enabled = .Buttons("save").Enabled
                .Buttons("envio").Enabled = .Buttons("save").Enabled
            
            Case 1
                .Buttons("save").Visible = True
                .Buttons("save").Caption = "Imprimir envíos"
                .Buttons("save").Enabled = (vsGrid.Rows > vsGrid.FixedRows)
                .Buttons("sugerircamion").Enabled = .Buttons("save").Enabled
                .Buttons("envio").Enabled = .Buttons("save").Enabled
            
            Case 2, 3
                .Buttons("envio").Enabled = (vsGrid.Rows > vsGrid.FixedRows)
            
            Case 5
                .Buttons("save").Visible = True
                .Buttons("save").Caption = "Entregó todo"
        End Select
    End With
End Sub

Private Sub s_ChangeData()
    iCamSelect = 0
    s_SetTitle
    s_SetMenu
End Sub

Private Sub s_GridSelectAll(ByVal bCheck As Boolean)
On Error Resume Next
    vsGrid.Cell(flexcpChecked, vsGrid.FixedRows, 0, vsGrid.Rows - 1) = IIf(bCheck, flexChecked, flexUnchecked)
End Sub

Private Sub frm_ShowPopUp()
Dim iIDC As Integer
On Error GoTo errSPP

    If Val(hlTab(0).Tag) = 4 Or Val(hlTab(0).Tag) = 6 Then Exit Sub

    MnuRecepEntregado.Visible = (Val(hlTab(0).Tag) = 5)
    MnuCambiarDividir.Visible = (Val(hlTab(0).Tag) = 5 And Val(vsGrid.Cell(flexcpData, vsGrid.RowSel, 8)) <> 48)
    MnuCambiarLinea.Visible = (Val(hlTab(0).Tag) = 2 Or Val(hlTab(0).Tag) = 5)
    MnuCambiarHora.Visible = MnuCambiarLinea.Visible
    MnuCambiarCamion.Visible = MnuCambiarLinea.Visible
    MnuCambiarAnular.Visible = MnuCambiarLinea.Visible
    MnuEliminoVaCon.Visible = MnuCambiarDividir.Visible

    MnuGridAConfirmar.Visible = True
    MnuGridAAsignar.Visible = True
    MnuGridSelectAll.Visible = False
    MnuGridUnSelectAll.Visible = False
    MnuGridL2.Visible = True
    MnuGridCamion(0).Tag = ""
    MnuGridSugerirC.Visible = (Val(hlTab(0).Tag) = 0)
    MnuGridL3.Visible = True
    MnuGridL4.Visible = True
    MnuGridL5.Visible = True
    MnuPrint.Visible = True
    MnuPrintAux.Enabled = False
    MnuGridACamion.Visible = True
    MnuGridACamion.Enabled = True
    MnuGridTodosA.Visible = False
    
    frm_HideShowMenuCamion -1
    
    MnuGridACamion.Caption = IIf(Val(hlTab(0).Tag) = 0, "Asignar envío a ...", "Asignar envío(s) a ...")
    
    If Val(hlTab(0).Tag) <> 2 And Val(hlTab(0).Tag) <> 5 And Val(hlTab(0).Tag) <> 6 Then
        MnuGridCamion(0).Caption = "No hay camión disponible"
        If Val(hlTab(0).Tag) = 3 Then
            If IsDate(vsGrid.Cell(flexcpText, vsGrid.Rows - 1, 1)) Then
                'Busco el id del camión que sería para asignar manual.
                iIDC = f_GetCamionAutomático(vsGrid.Cell(flexcpValue, vsGrid.Row, 0))
            End If
        Else
            iIDC = f_GetCamionAutomático(vsGrid.Cell(flexcpValue, vsGrid.Row, 0))
        End If
        
        If iIDC > 0 Then
            If iCamSelect = iIDC Then
                MnuGridCamion(0).Caption = "Sugerido es el mismo"
                frm_HideShowMenuCamion iCamSelect
            Else
                frm_HideShowMenuCamion iCamSelect
                iIDC = f_GetCamionByID(iIDC)
                If iIDC > 0 Then
                    frm_HideShowMenuCamion arrCamion(iIDC).Codigo
                    MnuGridCamion(0).Caption = arrCamion(iIDC).Nombre
                    MnuGridCamion(0).Tag = arrCamion(iIDC).Codigo
                Else
                    MnuGridCamion(0).Caption = "No hay camión disponible"
                End If
            End If
        ElseIf iCamSelect > 0 Then
            frm_HideShowMenuCamion iCamSelect
        End If
    End If
    
    If vsGrid.Rows > vsGrid.FixedRows Then
        Select Case Val(hlTab(0).Tag)
            Case 0
                MnuGridSelectAll.Visible = True
                MnuGridUnSelectAll.Visible = True
                MnuGridEditEnvio.Enabled = True
                MnuGridAAsignar.Visible = False
                MnuPrintAux.Enabled = True
                MnuGridTodosA.Visible = True
                
            Case 1
                MnuGridL2.Visible = False
                MnuGridEditEnvio.Enabled = True
                MnuGridAConfirmar.Visible = True
                MnuGridAAsignar.Visible = True
                MnuPrintAux.Enabled = True

            Case 2
                MnuGridAConfirmar.Visible = False
                MnuGridAAsignar.Visible = False
                MnuGridL2.Visible = False
                MnuGridEditEnvio.Enabled = True
                MnuGridL4.Visible = False
                MnuGridACamion.Enabled = False
                MnuCambiarAnular.Enabled = (Val(vsGrid.Cell(flexcpData, vsGrid.Row, 2)) = 0)
            
            Case 3
                MnuGridL2.Visible = False
                MnuGridEditEnvio.Enabled = True
                MnuGridAAsignar.Enabled = True  '(vsGrid.Cell(flexcpText, vsGrid.Row, 4) <> "")
                MnuGridAConfirmar.Visible = False
                MnuGridACamion.Enabled = MnuGridAAsignar.Enabled
                MnuGridACamion.Visible = MnuGridACamion.Enabled
                
            Case 5
                MnuGridL2.Visible = False
                MnuGridEditEnvio.Enabled = True
                MnuGridAAsignar.Visible = False
                MnuGridAConfirmar.Visible = False
                MnuGridL4.Visible = False
                MnuGridL3.Visible = False
                MnuGridACamion.Visible = False
                If vsGrid.Row >= 1 Then
                    MnuEliminoVaCon.Enabled = (Val(vsGrid.Cell(flexcpData, vsGrid.Row, 2)) > 0)
                    MnuCambiarAnular.Enabled = (Val(vsGrid.Cell(flexcpData, vsGrid.Row, 2)) = 0)
                End If
                
            Case 7
                MnuGridL2.Visible = False
                MnuGridEditEnvio.Enabled = True
                MnuGridAAsignar.Visible = False
                MnuGridAConfirmar.Visible = False
                MnuGridL4.Visible = False
                MnuGridL3.Visible = False
                MnuGridACamion.Visible = False
'                If vsGrid.Row >= 1 Then
'                    MnuEliminoVaCon.Enabled = (Val(vsGrid.Cell(flexcpData, vsGrid.Row, 2)) > 0)
'                    MnuCambiarAnular.Enabled = (Val(vsGrid.Cell(flexcpData, vsGrid.Row, 2)) = 0)
'                End If
                
        End Select
        PopupMenu MnuGrid
    End If
    
Exit Sub
errSPP:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al desplegar el menú de opciones.", Err.Description, "Error (showpopup)"
End Sub

Private Sub db_FillGridAImprimirImpreso(ByVal lCamion As Long, ByVal eEstado As EstadoEnvio)
On Error GoTo errAI
Dim sFM As String, lLastID As Long, lLastVC As Long

    Screen.MousePointer = 11
    ReDim arrNomArt(0)
    
    Cons = "Select EnvCodigo, EnvFModificacion, IsNull(EnvRangoHora, '') as RH, EnvTipo, LocCodigo, LocNombre, " _
                    & " CalNombre, DirPuerta, DirLetra, DirApartamento, rTrim(TFlNombreCorto) as TF, " _
                    & " IsNull(EVCID, 0) as VC, IsNull(EnvComentario, '') as Memo, IsNull(rTrim(AgeNombre), rTrim(ZonNombre) + ' (' + RTRIM(DepNombre) COLLATE Modern_Spanish_CI_AI  + ')') as ZN, REvCantidad, IsNull(AEsID, ArtCodigo) ArtCodigo, IsNull(AEsNombre, ArtNombre) as AD, ArtTipo, IsNull(AEsID, ArtID) ArtID " _
                    & ", CQCIdQuePaga RedPagos, TFLTipoFlete, EnvHoraEspecial, DocTipo" _
        & " From ((((((((((Envio Left Outer Join Agencia On EnvAgencia = AgeCodigo) LEFT OUTER JOIN EnvioVaCon ON EnvCodigo = EVCEnvio)" _
        & " INNER JOIN Direccion ON EnvDireccion = DirCodigo)" _
        & " INNER JOIN Calle ON DirCalle = CalCodigo) INNER JOIN Localidad ON CalLocalidad = LocCodigo INNER JOIN Departamento ON DepCodigo = LocDepartamento) " _
        & " INNER JOIN TipoFlete ON EnvTipoFlete = TFlCodigo)" _
        & " INNER JOIN Zona ON EnvZona = ZonCodigo ) INNER JOIN RenglonEnvio ON EnvCodigo = REvEnvio) " _
        & " INNER JOIN Articulo ON REvArticulo = ArtID)" _
        & " LEFT OUTER JOIN ArticuloEspecifico ON AEsDocumento = EnvDocumento AND AEsArticulo = REvArticulo AND ((AEsTipoDocumento IN (1, 6) AND ENVTipo = 1) OR (AEsTipoDocumento IN (7, 33) AND ENVTipo = 3)))" _
        & " LEFT OUTER JOIN ConQueCobra ON CQCIdQueCobra = EnvDocumento AND CQCTipoQueCobra IN (1,2) And CQCTipoQuePaga = 15 " _
        & " LEFT OUTER JOIN Documento ON DocCodigo = EnvDocumento AND EnvTipo = 1" _
        & " Where EnvFechaPrometida = '" & Format(tFecha.Value, "mm/dd/yyyy") & "'" _
        & " And EnvEstado = " & eEstado & " And EnvCamion = " & lCamion _
        & " And EnvTipo In (1, 2, 3) And EnvDocumento > 0" _
        & " Order by EVCID, EnvZona, RH, EnvCodigo"


'& " LEFT OUTER JOIN ArticuloEspecifico ON AEsTipoDocumento IN (1, 2, 6) AND AEsDocumento = EnvDocumento AND AEsArticulo = REvArticulo)" _

Dim totalBultos As Integer

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    vsGrid.Redraw = False
    Do While Not RsAux.EOF
        With vsGrid
            If lLastID <> RsAux("EnvCodigo") Then
                
                lLastID = RsAux("EnvCodigo")
                If lLastVC <> RsAux("VC") Or RsAux("VC") = 0 Then
                    lLastVC = RsAux("VC")
                    .AddItem RsAux!EnvCodigo
                    .Cell(flexcpText, .Rows - 1, 2) = RsAux!RH
                    If Not IsNull(RsAux("EnvHoraEspecial")) Then .Cell(flexcpFontBold, .Rows - 1, 2) = True
                    .Cell(flexcpText, .Rows - 1, 3) = RsAux!TF
                    If RsAux("TFLTipoFlete") = eTiposDeTipoFlete.CostoEspecial Then
                        .Cell(flexcpFontBold, .Rows - 1, 3) = True
                        .Cell(flexcpForeColor, .Rows - 1, 2, .Rows - 1, .Cols - 1) = vbRed
                    End If
                    .Cell(flexcpText, .Rows - 1, 4) = RsAux!ZN
                    .Cell(flexcpText, .Rows - 1, 5) = f_GetDireccionRsAux
                    .Cell(flexcpText, .Rows - 1, 6) = Trim(RsAux!Memo)
                    
                    If Not IsNull(RsAux("RedPagos")) Then .Cell(flexcpForeColor, .Rows - 1, 4) = &H80FF& ': .Cell(flexcpFontBold, .Rows - 1, 4) = True
                                    
                    'DATA
                    sFM = RsAux!EnvFModificacion: .Cell(flexcpData, .Rows - 1, 1) = sFM         'F Modificado
                    sFM = RsAux!VC: .Cell(flexcpData, .Rows - 1, 2) = Val(sFM)                      'Va Con
                    
                    If RsAux("EnvTipo") = 3 Then
                        .Cell(flexcpForeColor, .Rows - 1, 0) = &HC0
                    ElseIf RsAux("EnvTipo") = 2 Then
                        .Cell(flexcpForeColor, .Rows - 1, 0) = &H8000&
                    End If
                    
                    If Not IsNull(RsAux("DocTipo")) Then
                        If RsAux("DocTipo") = TD_Contado Then
                            .Cell(flexcpText, .Rows - 1, 8) = "Contado"
                        ElseIf RsAux("DocTipo") = TD_Credito Then
                            .Cell(flexcpText, .Rows - 1, 8) = "Crédito"
                        ElseIf RsAux("DocTipo") = 47 Then
                            .Cell(flexcpText, .Rows - 1, 8) = "Cambio"
                        ElseIf RsAux("DocTipo") = 48 Then
                            .Cell(flexcpText, .Rows - 1, 8) = "Retiro"
                        End If
                        .Cell(flexcpData, .Rows - 1, 8) = CStr(RsAux("DocTipo"))
                    Else
                        .Cell(flexcpText, .Rows - 1, 8) = "Venta S/F"
                        .Cell(flexcpData, .Rows - 1, 8) = ""
                    End If
                    
                End If
            End If
            '..........................ARTICULOS
            loc_AgregoEnColleccionBultos RsAux("ArtID"), Format(RsAux("ArtCodigo"), "(#,000,000)") & " " & RsAux("AD"), RsAux("REvCantidad")
            loc_AddArtNombre RsAux("ArtID"), Format(RsAux("ArtCodigo"), "(#,000,000)") & " " & RsAux("AD")
            loc_SetQTipoArt .Rows - 1, RsAux("ArtID"), RsAux("ArtTipo"), RsAux("REvCantidad")
            totalBultos = totalBultos + RsAux("REvCantidad")
            .Cell(flexcpText, .Rows - 1, 7) = fnc_GetStringArticulos(.Cell(flexcpData, .Rows - 1, 6))
            '..........................Artículos
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    vsGrid.Redraw = True
    sbStatus.Panels("envio").Text = "Envíos: " & vsGrid.Rows - 1
    sbStatus.Panels("bultos").Text = "Bultos: " & totalBultos
    
    loc_SetStatusTipoArt
    Screen.MousePointer = 0
    
    Exit Sub
    Screen.MousePointer = 0
    Exit Sub
errAI:
    vsGrid.Redraw = True
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al cargar la lista de envíos", Err.Description
End Sub

Private Function fnc_CajaCerrada() As Boolean
    
    Cons = "Select * FROM SaldoDisponibilidad " _
        & " Where SDiFecha = '" & Format(Date + 1, "mm/dd/yyyy") & "'" _
        & " And SDiHora = '00:00:00'" _
        & " And SDiDisponibilidad = " & modMaeDisponibilidad.dis_DisponibilidadPara(paCodigoDeSucursal, 1)
    Dim rsC As rdoResultset
    Set rsC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsC.EOF Then
        fnc_CajaCerrada = True
        MsgBox "Atención la caja ya fué cerrada no podrá imprimir envíos hasta mañana.", vbExclamation, "Validar cierre de caja"
    End If
    rsC.Close
    
End Function

Private Sub s_SaveImprimoEnvio()
'Ocurre al apretar grabar en toolbar.
Dim iQ As Integer, sEnvios As String
Dim lIDImp As Long
Dim sErr As String
Dim rsVC As rdoResultset

    Set docToPrint = New Collection
    Set docCopia = New Collection
    
    'Controles:
    '   1) Si hay vtas telefónicas que tengan el cobro en otro envío y con fecha posterior.
    '   2) Facturar diferencias de envíos.
    '   3) Facturar Vtas telefonicas y Servicios
    '   4) Facturar pago
    '   5) Cambio estado al envío
    
    With vsGrid
        For iQ = .FixedRows To .Rows - 1
            'Copio la fecha de edición.
            .Cell(flexcpData, iQ, 4) = .Cell(flexcpData, iQ, 1)
            
            If sEnvios <> "" Then sEnvios = sEnvios & ", "
            sEnvios = sEnvios & .Cell(flexcpValue, iQ, 0)
        Next
    End With
    
    
    '13/6/2011 si el camión empieza con bolsa entonces no dejo grabar.
    Cons = "SELECT Count(*) FROM Envio INNER JOIN Camion ON EnvCamion = CamCodigo AND CamNombre Like 'bolsa%' " & _
        " WHERE EnvCodigo IN (" & sEnvios & ")"
    Dim rsC As rdoResultset
    Set rsC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsC.EOF Then
        If rsC(0) > 0 Then
            MsgBox "No puede imprimir envíos con el camión 'Bolsa'", vbExclamation, "Posible error"
            rsC.Close
            Exit Sub
        End If
    End If
    rsC.Close
    
    'Control de las ventas telefónicas que puedan tener el cobro con fecha posterior.
    
    '10/8 anulo este primer control quedamos en que no se dejan imprimir envíos de vtas telefónicas
    ' que no se haya impreso aún el de cobranza.
    'If Not fnc_ControlVtaTelFechaPosterior(sEnvios) Then Exit Sub
    
    If Not fnc_ControlVtaTelCobranzaEnOtroEnvio(sEnvios) Then Exit Sub
    
    If Not fnc_FrenoPorVentas40KUI(sEnvios) Then Exit Sub
    
    'Verifico que la caja no este cerrada.
    If fnc_CajaCerrada Then Exit Sub
    
    FechaDelServidor
    Screen.MousePointer = 11
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo errRB
    
    'Código de Impresión.
    lIDImp = Mid(NumeroDocumento("CodigoImpresion"), 2)

'    Erase arrHojaBlanca
'    ReDim arrHojaBlanca(0)

    Set docCopia = New Collection

    For iQ = vsGrid.FixedRows To vsGrid.Rows - 1
        If vsGrid.Cell(flexcpData, iQ, 2) > 0 Then
            'Es un vacon
            'Debug.Print vsGrid.Cell(flexcpData, iQ, 2) & " Arriba"
            Cons = "Select EVCEnvio From EnvioVaCon Where EVCID = " & vsGrid.Cell(flexcpData, iQ, 2)
            Set rsVC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            Do While Not rsVC.EOF
                db_EnvioAEntregar rsVC("EVCEnvio"), lIDImp, CDate(vsGrid.Cell(flexcpData, iQ, 4)), sErr, vsGrid.Cell(flexcpData, iQ, 2), Val(vsGrid.Cell(flexcpData, iQ, 8))
                rsVC.MoveNext
            Loop
            rsVC.Close
        Else
'            Debug.Print vsGrid.Cell(flexcpValue, iQ, 0)
            db_EnvioAEntregar vsGrid.Cell(flexcpValue, iQ, 0), lIDImp, CDate(vsGrid.Cell(flexcpData, iQ, 4)), sErr, 0, Val(vsGrid.Cell(flexcpData, iQ, 8))
        End If
    Next
    cBase.CommitTrans

   If MnuOtrSendWhatsapp.Checked Then whatsappEnviarMensajeCodigo lIDImp, ""

    On Error GoTo errGo
    'oLog.InsertoLog TL_Debug, "Código generado = " & lIDImp
    loc_ImprimoContados
    
    loc_ImprimoHojaBlanca lIDImp
    loc_ImprimoReparto lIDImp
    
    vsGrid.Rows = vsGrid.FixedRows
    
    Dim documentosImpresos As String
    documentosImpresos = ""
    documentosImpresos = "El código de impresión que se acaba de generar es el " & lIDImp & "."
    MsgBox documentosImpresos, vbInformation, "Atención"
    
    Set docToPrint = New Collection
    Set docCopia = New Collection
    Screen.MousePointer = 0

Exit Sub
errGo:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al restaurar el formulario.", Err.Description
    On Error Resume Next
    vsGrid.Rows = vsGrid.FixedRows
    Set docToPrint = New Collection
    Set docCopia = New Collection
    Exit Sub

ErrBT:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al bloquear.", Err.Description
    Exit Sub

errRB:
    Resume errVT
    Exit Sub
    
errVT:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al grabar los envíos." & IIf(sErr <> "", vbCr & vbCr & sErr, ""), Err.Description
End Sub

Private Sub loc_ImprimoRemitosEnvios(ByVal codImp As Long)
Dim rsIR As rdoResultset
Dim oPrintEF As New ComPrintEfactura.ImprimoCFE

On Error GoTo errIR

    Cons = "SELECT EnvCodigo FROM Envio" _
        & " WHERE EnvCodImpresion = " & codImp & " And EnvTipo = 2"
    
    Set rsIR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    
    Do While Not rsIR.EOF
        oPrintEF.ImprimirRemitoCadeteria rsIR("EnvCodigo"), paPrintRemRepD, paPrintRemRepB, paPrintConfPaperSize
        rsIR.MoveNext
    Loop
    rsIR.Close
    
    Exit Sub

errIR:
    MsgBox "Error al imprimir los remitos: " & Err.Description, vbError, "ATENCIÓN"
End Sub


Private Sub db_EnvioAEntregar(ByVal lEnvio As Long, ByVal lIDImp As Long, ByVal dFedit As Date, ByRef sErr As String, ByVal lIDVaCon As Long, ByVal docTipo As Integer)
Dim rsE As rdoResultset, rsA As rdoResultset
Dim lIDArtF As Long, lIDDocFact As Long, lIDDocVta As Long
Dim cImp As Currency, cIVA As Currency
Dim bVTaAFact As Boolean, bChangeFP As Boolean
Dim strNroDoc As String
Dim CAE As New clsCAEDocumento
Dim caeG As New clsCAEGenerador
Dim oDocCtdo As clsDocumentoCGSA
Dim oRen As clsDocumentoRenglon

    Dim sMVta As String
    
    bVTaAFact = False
    lIDDocFact = 0
    cImp = 0
    cIVA = 0
    lIDDocVta = 0
        
    Dim bolAHojaBlanca As Boolean
    'Dim bClienteEsRUT As Boolean
        
    Dim oDoc As clsDocToPrint
    
    'Consulto todos los envíos inclusive los vaCon que no están en la grilla.
    Cons = "Select * From Envio Where EnvCodigo = " & lEnvio
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsE.EOF Then
        
        sErr = "Envío: " & rsE("EnvCodigo")
        
        'Válido Fecha de Modificación.
        'OJO Una vta telefónica puede alterar envíos propios que estén en otro va con o en otra fila.
        If dFedit <> rsE("EnvFModificacion") And gFechaServidor <> rsE("EnvFModificacion") Then
            sErr = "Envío: " & rsE("EnvCodigo") & " fue modificado."
            rsE.Close
            rsE.Update
        End If

        Dim oCliFact As clsClienteCFE
        If rsE("EnvTipo") = 3 Then
            Set oCliFact = RetornoCliente(rsE("EnvCliente"), rsE("EnvDocumento"))
        Else
            Set oCliFact = RetornoCliente(rsE("EnvCliente"), 0)
        End If
        
        If (oCliFact.RUT <> "") Then
            Dim oValida As New clsValidaRUT
            If Not oValida.ValidarRUT(oCliFact.RUT) Then
                Set oValida = Nothing
                sErr = "El RUT " & oCliFact.RUT & " no es correcto, por favor edite la ficha del cliente, no podrá continuar."
                rsE.Close
                rsE.Update
            End If
            Set oValida = Nothing
        End If
        
        'Facturo diferencias de envío.
        db_FacturoDiferenciaEnvio rsE("EnvCodigo"), lIDVaCon, oCliFact
    
        If rsE("EnvTipo") = 1 Then
            
            Dim bytF As Byte
            Dim bolEsta As Boolean
            
            
            'Veo si está en un remitoretiro y le modifico el documento
            Cons = "SELECT * FROM EnviosRemitos " _
                & " WHERE EReEnvio = " & rsE("EnvCodigo") & " AND EReNota IS NOT NULL AND EReFactura = " & rsE("EnvDocumento")
            Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsA.EOF Then
            
                Set oDoc = New clsDocToPrint
                docToPrint.Add oDoc
                oDoc.Documento = rsA("EReRemito")
                oDoc.idEnvio = rsE("EnvCodigo")
                oDoc.TipoDoc = TD_RemitoRetiro
                
                If Not IsNull(rsE("EnvAgencia")) Then
                    oDoc.EsConAgencia = (rsE("EnvAgencia") > 0)
                End If
                
            End If
            rsA.Close
            
            
            'Si es una venta telefónica siempre la envío en papel blanco.
            If IsNull(rsE("EnvReclamoCobro")) Then
            
                If docTipo = CGSA_TipoDocumento.TD_RemitoEntrega Or docTipo = CGSA_TipoDocumento.TD_RemitoRetiro Then
                    'Este lo saco en hoja blanca
                    Set oDoc = New clsDocToPrint
                    docToPrint.Add oDoc
                    oDoc.Documento = rsE("EnvDocumento")
                    oDoc.idEnvio = rsE("EnvCodigo")
                    oDoc.TipoDoc = docTipo
                    
                    If docTipo = CGSA_TipoDocumento.TD_RemitoEntrega Then
                        'Saco el remito de regreso para que lo imprima.
                        Cons = "SELECT EReRemito, EReNota FROM EnviosRemitos WHERE EReEnvio = " & rsE("EnvCodigo")
                        Set rsA = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                                                
                        Set oDoc = New clsDocToPrint
                        docToPrint.Add oDoc
                        oDoc.Documento = rsA("EReRemito")
                        oDoc.idEnvio = rsE("EnvCodigo")
                        oDoc.TipoDoc = CGSA_TipoDocumento.TD_RemitoRetiro
                        rsA.Close
                    Else
                        'imprimo 2 copas.
                        Set oDoc = New clsDocToPrint
                        docToPrint.Add oDoc
                        oDoc.Documento = rsE("EnvDocumento")
                        oDoc.idEnvio = rsE("EnvCodigo")
                        oDoc.TipoDoc = docTipo
                    End If
                Else
                
                    bolEsta = ValidarCopiaInsertada(rsE("EnvDocumento"), lEnvio)
                    If Not bolEsta Then
                        Set oDoc = New clsDocToPrint
                        oDoc.Documento = rsE("EnvDocumento")
                        oDoc.idEnvio = rsE("EnvCodigo")
                        If Not IsNull(rsE("EnvAgencia")) Then
                            oDoc.EsConAgencia = (rsE("EnvAgencia") > 0)
                        End If
                        docCopia.Add oDoc
                        'arrHojaBlanca(UBound(arrHojaBlanca)).tipo = tipoDocHB
                    End If
                End If
            Else
                Set oDoc = New clsDocToPrint
                docToPrint.Add oDoc
                oDoc.Documento = rsE("EnvDocumento")
                oDoc.idEnvio = rsE("EnvCodigo")
                If Not IsNull(rsE("EnvAgencia")) Then
                    oDoc.EsConAgencia = (rsE("EnvAgencia") > 0)
                End If
            End If
            
            If rsE("EnvFormaPago") = 2 And IsNull(rsE("EnvDocumentoFactura")) Then
                                
                Cons = "Select TFlArticulo From TipoFlete Where TFlCodigo = " & rsE("EnvTipoFlete")
                Set rsA = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                lIDArtF = rsA(0)
                rsA.Close
                                    
                If Not IsNull(rsE("EnvValorFlete")) Then cImp = rsE("EnvValorFlete"): cIVA = rsE("EnvIvaFlete")
                If Not IsNull(rsE("EnvValorPiso")) Then cImp = cImp + rsE("EnvValorPiso"): cIVA = cIVA + rsE("EnvIvaPiso")
                
                If cImp > 0 Then
                    
                    Set CAE = New clsCAEDocumento
                    If Val(prmEFacturaProductivo) = 0 Then
                        strNroDoc = NumeroDocumento(paDContado)
                        With CAE
                            .Desde = 1
                            .Hasta = 9999999
                            .Serie = Mid(strNroDoc, 1, 1)
                            .Numero = Mid(strNroDoc, 2)
                            .IdDGI = "9014"
                            .TipoCFE = IIf(oCliFact.RUT <> "", CGSA_TiposCFE.CFE_eFactura, CFE_eTicket)
                            .Vencimiento = "31/12/" & CStr(Year(Date))
                        End With
                    Else
                        Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, IIf(oCliFact.RUT <> "", CGSA_TiposCFE.CFE_eFactura, CFE_eTicket), paCodigoDGI)
                    End If
                    Set oDocCtdo = New clsDocumentoCGSA
                    With oDocCtdo
                        Set .cliente = oCliFact
                        .Emision = gFechaServidor
                        .Tipo = TD_Contado
                        .Numero = CAE.Numero
                        .Serie = CAE.Serie
                        .Moneda.Codigo = rsE("EnvMoneda")
                        .Total = cImp
                        .IVA = cIVA
                        .sucursal = paCodigoDeSucursal
                        .Digitador = paCodigoDeUsuario
                        .Comentario = ""
                        .Vendedor = paCodigoDeUsuario
                    End With
                    If Not IsNull(rsE("EnvValorPiso")) Then
                        If rsE("EnvValorPiso") > 0 Then
                            If oArtPisoAgencia Is Nothing Then Set oArtPisoAgencia = CargoArticulosPrms(paArticuloPisoAgencia)
                            Set oRen = New clsDocumentoRenglon
                            Set oRen.Articulo = oArtPisoAgencia
                            oRen.cantidad = 1
                            oRen.IVA = rsE("EnvIvaPiso")
                            oRen.Precio = rsE("EnvValorPiso")
                            oRen.CantidadARetirar = 1
                            oRen.Descripcion = ""
                            oDocCtdo.Renglones.Add oRen
                        End If
                    End If
                    
                    If Not IsNull(rsE("EnvValorFlete")) Then
                        If rsE("EnvValorFlete") > 0 Then
                            If Not oArtFlete Is Nothing Then
                                If oArtFlete.ID <> lIDArtF Then Set oArtFlete = Nothing
                            End If
                            If oArtFlete Is Nothing Then Set oArtFlete = CargoArticulosPrms(lIDArtF)
                            
                            Set oRen = New clsDocumentoRenglon
                            Set oRen.Articulo = oArtFlete
                            oRen.cantidad = 1
                            oRen.IVA = rsE("EnvIvaFlete")
                            oRen.Precio = rsE("EnvValorFlete")
                            oRen.CantidadARetirar = 1
                            oRen.Descripcion = ""
                            oDocCtdo.Renglones.Add oRen
                            
                        End If
                    End If
                    Set oDocCtdo.Conexion = cBase
                    oDocCtdo.Codigo = oDocCtdo.InsertoDocumentoBD(0)
                    
                    Set oDoc = New clsDocToPrint
                    docToPrint.Add oDoc
                    With oDoc
                        .Documento = oDocCtdo.Codigo
                        .idEnvio = rsE("EnvCodigo")
                        If Not IsNull(rsE("EnvAgencia")) Then
                            .EsConAgencia = (rsE("EnvAgencia") > 0)
                        End If
                    End With
                    lIDDocFact = oDocCtdo.Codigo
                    'Inserto dentro del array a imprimir.
                    loc_InsertDocumentoPendiente oDocCtdo.Codigo, 1, rsE("EnvCodigo"), cImp, rsE("EnvMoneda")
                End If
            End If
            
        ElseIf rsE("EnvTipo") = 2 Then
            'SERVICIO NO ASOCIADO A COMPRA
            
            'Pago Domicilio
            If rsE("EnvFormaPago") = 2 Then
                
                If IsNull(rsE("EnvDocumentoFactura")) Then
                    If Not IsNull(rsE("EnvValorFlete")) Then cImp = rsE("EnvValorFlete"): cIVA = rsE("EnvIvaFlete")
                    If Not IsNull(rsE("EnvValorPiso")) Then cImp = cImp + rsE("EnvValorPiso"): cIVA = cIVA + rsE("EnvIvaPiso")
                End If
                
                If cImp > 0 Then
                    Cons = "Select TFlArticulo From TipoFlete Where TFlCodigo = " & rsE("EnvTipoFlete")
                    Set rsA = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    lIDArtF = rsA(0)
                    rsA.Close
                
                    'Levanto la Venta telefónica o Servicio
                    Cons = "Select * From VentaTelefonica Where VTeCodigo = " & rsE("EnvDocumento") & " AND VTeTipo <> 44"
                    Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    'Inserto documento.
                    If Not IsNull(rsA("VTeComentario")) Then sMVta = Trim(rsA("VTeComentario")) Else sMVta = ""
                    Set CAE = New clsCAEDocumento
                    If Val(prmEFacturaProductivo) = 0 Then
                        strNroDoc = NumeroDocumento(paDContado)
                        With CAE
                            .Desde = 1
                            .Hasta = 9999999
                            .Serie = Mid(strNroDoc, 1, 1)
                            .Numero = Mid(strNroDoc, 2)
                            .IdDGI = "9014"
                            .TipoCFE = IIf(oCliFact.RUT <> "", CGSA_TiposCFE.CFE_eFactura, CFE_eTicket)
                            .Vencimiento = "31/12/" & CStr(Year(Date))
                        End With
                    Else
                        Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, IIf(oCliFact.RUT <> "", CGSA_TiposCFE.CFE_eFactura, CFE_eTicket), paCodigoDGI)
                    End If
                    Set oDocCtdo = New clsDocumentoCGSA
                    With oDocCtdo
                        Set .cliente = oCliFact
                        .Emision = gFechaServidor
                        .Tipo = TD_Contado
                        .Numero = CAE.Numero
                        .Serie = CAE.Serie
                        .Moneda.Codigo = rsA("VTeMoneda")
                        .Total = cImp
                        .IVA = cIVA
                        .sucursal = paCodigoDeSucursal
                        .Digitador = paCodigoDeUsuario
                        .Comentario = sMVta
                        .Vendedor = paCodigoDeUsuario
                    End With
                    
                    Set oDocCtdo.Renglones = CopioRenglonesVentaTelefonica(rsE("EnvDocumento"))
                    
                    If Not IsNull(rsE("EnvValorPiso")) Then
                        If rsE("EnvValorPiso") > 0 Then
                            If oArtPisoAgencia Is Nothing Then Set oArtPisoAgencia = CargoArticulosPrms(paArticuloPisoAgencia)
                            Set oRen = New clsDocumentoRenglon
                            Set oRen.Articulo = oArtPisoAgencia
                            oRen.cantidad = 1
                            oRen.IVA = rsE("EnvIvaPiso")
                            oRen.Precio = rsE("EnvValorPiso")
                            oRen.CantidadARetirar = 1
                            oRen.Descripcion = ""
                            oDocCtdo.Renglones.Add oRen
                        End If
                    End If
                    
                    If Not IsNull(rsE("EnvValorFlete")) Then
                        If rsE("EnvValorFlete") > 0 Then
                            If Not oArtFlete Is Nothing Then
                                If oArtFlete.ID <> lIDArtF Then Set oArtFlete = Nothing
                            End If
                            If oArtFlete Is Nothing Then Set oArtFlete = CargoArticulosPrms(lIDArtF)
                        
                            Set oRen = New clsDocumentoRenglon
                            Set oRen.Articulo = oArtFlete
                            oRen.cantidad = 1
                            oRen.IVA = rsE("EnvIvaFlete")
                            oRen.Precio = rsE("EnvValorFlete")
                            oRen.CantidadARetirar = 1
                            oRen.Descripcion = ""
                            oDocCtdo.Renglones.Add oRen
                        End If
                    End If
                    Set oDocCtdo.Conexion = cBase
                    oDocCtdo.Codigo = oDocCtdo.InsertoDocumentoBD(0)
                    
                    loc_InsertDocumentoPendiente oDocCtdo.Codigo, 1, rsE("EnvCodigo"), cImp, rsA("VTeMoneda")
                    
                    Set oDoc = New clsDocToPrint
                    docToPrint.Add oDoc
                    With oDoc
'                        Set .CAE = CAE
                        .Documento = oDocCtdo.Codigo
                        .idEnvio = rsE("EnvCodigo")
                        If Not IsNull(rsE("EnvAgencia")) Then
                            oDoc.EsConAgencia = (rsE("EnvAgencia") > 0)
                        End If
                    End With
                                        
                    rsA.Edit
                    rsA("VTeDocumento") = oDocCtdo.Codigo
                    rsA("VTeFModificacion") = Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss")
                    rsA.Update
                    rsA.Close
                    
                    lIDDocFact = oDocCtdo.Codigo
                End If
            End If
                        
            'No tenía forma de pago domicilio.
            If cImp = 0 Then
                Cons = "Select * From VentaTelefonica Where VTeCodigo = " & rsE("EnvDocumento")
                Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                rsA.Edit
                rsA("VTeDocumento") = 0
                rsA("VTeFModificacion") = Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss")
                rsA.Update
                rsA.Close
            End If
                       
        ElseIf rsE("EnvTipo") = 3 Then
            'VENTA TELEFONICA
            
            If rsE("EnvFormaPago") = 2 Then
                Cons = "Select TFlArticulo From TipoFlete Where TFlCodigo = " & rsE("EnvTipoFlete")
                Set rsA = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                lIDArtF = rsA(0)
                rsA.Close
            End If
            
            Cons = "Select * From VentaTelefonica Where VTeCodigo = " & rsE("EnvDocumento") & " AND VTeTipo <> 44"
            Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                            
            If IsNull(rsA("VTeDocumento")) Then
                
                If IsNull(rsE("EnvReclamoCobro")) Then
                    'El que cobra es otro envío y está en la lista pero me salio posterior a este.
                    'Por lo tanto hago el contado modifico el otro envío y sigo adelante con este.
                    lIDDocVta = fnc_FacturoVtaTelefonicaEnOtro(rsA("VTeCodigo"), lIDVaCon, oCliFact)
                    
                    'UPDATEO LOS ARTICULOS ESPECIFICOS Y LOS PONGO AHORA EN EL CONTADO
                    Cons = "UPDATE ArticuloEspecifico SET AEsTipoDocumento = 1, AEsDocumento = " & lIDDocVta _
                            & " WHERE AEsTipoDocumento = " & rsA("VTeTipo") & " And AEsDocumento = " & rsA("VTeCodigo")
                    cBase.Execute Cons
                    
                    'Arriba me la edita x lo que tengo que recargar.
                    rsA.Close
                    Cons = "Select * From VentaTelefonica Where VTeCodigo = " & rsE("EnvDocumento")
                    Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    bChangeFP = (rsE("EnvFormaPago") = 2)
                    
                    'ATENCIÓN CIERRO Y ABRO YA QUE ME DA ERROR por modificación del registro.
                    rsE.Close
                    Cons = "Select * From Envio " & _
                        " Where EnvCodigo = " & lEnvio
                    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    'VA A REMITO HAY MAS DE UN ENVIO EN LA VENTA
                Else
                
                    bVTaAFact = True
                    'Si o sí es la que tiene el reclamo de cobro
                    If rsE("EnvReclamoCobro") <> rsA("VTeTotal") Then
                        rsA.Close
                        sErr = "Envío : " & rsE("EnvCodigo") & " tiene el importe de cobranza mal."
                        rsE.Close
                        rsA.Edit
                    End If
                    cImp = rsA("VTeTotal")
                    cIVA = rsA("VTeIVA")
                End If
            Else 'tengo que cambiar el documento al envío.
                lIDDocVta = rsA("VTeDocumento")
            End If
                    
            'Pago Domicilio
            If rsE("EnvFormaPago") = 2 Then
                If Not IsNull(rsE("EnvValorFlete")) Then cImp = rsE("EnvValorFlete") + cImp: cIVA = rsE("EnvIvaFlete") + cIVA
                If Not IsNull(rsE("EnvValorPiso")) Then cImp = cImp + rsE("EnvValorPiso"): cIVA = cIVA + rsE("EnvIvaPiso")
            End If
            
            If cImp > 0 Then
                
                If Not IsNull(rsA("VTeComentario")) Then sMVta = Trim(rsA("VTeComentario")) Else sMVta = ""
                
                Set CAE = New clsCAEDocumento
                Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, IIf(oCliFact.RUT <> "", CGSA_TiposCFE.CFE_eFactura, CFE_eTicket), paCodigoDGI)
                Set oDocCtdo = New clsDocumentoCGSA
                With oDocCtdo
                    Set .cliente = oCliFact
                    .Emision = gFechaServidor
                    .Tipo = TD_Contado
                    .Numero = CAE.Numero
                    .Serie = CAE.Serie
                    .Moneda.Codigo = rsA("VTeMoneda")
                    .Total = cImp
                    .IVA = cIVA
                    .sucursal = paCodigoDeSucursal
                    .Digitador = paCodigoDeUsuario
                    .Comentario = sMVta
                    .Vendedor = paCodigoDeUsuario
                End With
                
                If bVTaAFact Then Set oDocCtdo.Renglones = CopioRenglonesVentaTelefonica(rsE("EnvDocumento"))
                
                If rsE("EnvFormaPago") = 2 Then
                    If Not IsNull(rsE("EnvValorPiso")) Then
                        If rsE("EnvValorPiso") > 0 Then
                            If oArtPisoAgencia Is Nothing Then Set oArtPisoAgencia = CargoArticulosPrms(paArticuloPisoAgencia)
                            Set oRen = New clsDocumentoRenglon
                            Set oRen.Articulo = oArtPisoAgencia
                            oRen.cantidad = 1
                            oRen.IVA = rsE("EnvIvaPiso")
                            oRen.Precio = rsE("EnvValorPiso")
                            oRen.CantidadARetirar = 1
                            oRen.Descripcion = ""
                            oDocCtdo.Renglones.Add oRen
                        End If
                    End If
                    
                    If Not IsNull(rsE("EnvValorFlete")) Then
                        If rsE("EnvValorFlete") > 0 Then
                            If Not oArtFlete Is Nothing Then
                                If oArtFlete.ID <> lIDArtF Then Set oArtFlete = Nothing
                            End If
                            If oArtFlete Is Nothing Then Set oArtFlete = CargoArticulosPrms(lIDArtF)
                        
                            Set oRen = New clsDocumentoRenglon
                            Set oRen.Articulo = oArtFlete
                            oRen.cantidad = 1
                            oRen.IVA = rsE("EnvIvaFlete")
                            oRen.Precio = rsE("EnvValorFlete")
                            oRen.CantidadARetirar = 1
                            oRen.Descripcion = ""
                            oDocCtdo.Renglones.Add oRen
                        End If
                     End If
                End If
                Set oDocCtdo.Conexion = cBase
                oDocCtdo.Codigo = oDocCtdo.InsertoDocumentoBD(0)
                
                Set oDoc = New clsDocToPrint
                docToPrint.Add oDoc
                With oDoc
'                    Set .CAE = CAE
                    .Documento = oDocCtdo.Codigo
                    .idEnvio = rsE("EnvCodigo")
                    If Not IsNull(rsE("EnvAgencia")) Then
                        oDoc.EsConAgencia = (rsE("EnvAgencia") > 0)
                    End If
                End With
                    
                'Inserto dentro del array a imprimir.
                loc_InsertDocumentoPendiente oDocCtdo.Codigo, 1, rsE("EnvCodigo"), cImp, rsA("VTeMoneda")
                                    
                'Si no hice la factura por otro envío edito sino no.
                If IsNull(rsA("VTeDocumento")) Then
                    rsA.Edit
                    If bVTaAFact Then rsA("VTeDocumento") = oDocCtdo.Codigo
                    rsA("VTeFModificacion") = Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss")
                    rsA.Update
                End If
                
                If bVTaAFact Then
'                    db_InsertoRenglonesVTaTelefonica lIDDoc, rsE("EnvDocumento")
                    lIDDocVta = oDocCtdo.Codigo
                    
                    'Updateo en la relación del va con el documento
                    Cons = "Update EnvioVaCon SET EVCDocumento = " & oDocCtdo.Codigo & " Where EVCEnvio = " & rsE("EnvCodigo")
                    cBase.Execute (Cons)
                    
                    Cons = "UPDATE ArticuloEspecifico SET AEsTipoDocumento = 1, AEsDocumento = " & lIDDocVta _
                            & " WHERE AEsTipoDocumento = " & rsA("VTeTipo") & " And AEsDocumento = " & rsA("VTeCodigo")
                    cBase.Execute Cons
                    
                End If
                
                If rsE("EnvFormaPago") = 2 Then lIDDocFact = oDocCtdo.Codigo
                        
            End If
            rsA.Close
            
            
            If lIDDocVta > 0 Then UpdateInstalacionesADocumento lIDDocVta, rsE("EnvDocumento")
        End If
    
        If docTipo <> TD_RemitoRetiro Then
            loc_InsertRenglonEntrega lIDImp, rsE("EnvCodigo"), rsE("EnvCamion")
        End If
        
        If bVTaAFact Then
            'Tengo que updatear todos los envíos que pertenecen a esta vta.
            Cons = "Select * From Envio " & _
                        " Where EnvTipo = " & rsE("EnvTipo") & _
                        " And EnvDocumento = " & rsE("EnvDocumento") '& _
                        " And EnvCodigo <> " & rsE("EnvCodigo")
            'NO Descarto a este envío ya que puede ser un vacon
            Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            Do While Not rsA.EOF
                
                loc_ChangeFechaEdit rsA("EnvCodigo"), gFechaServidor
                If rsA("EnvCodigo") <> rsE("EnvCodigo") Then
                    rsA.Edit
                    rsA("EnvTipo") = 1
                    rsA("EnvDocumento") = lIDDocVta
                    rsA("EnvFModificacion") = Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss")
                    rsA.Update
                    
                    'Updateo en la relación del va con el documento
                    Cons = "Update EnvioVaCon SET EVCDocumento = " & oDocCtdo.Codigo & " Where EVCEnvio = " & rsA("EnvCodigo")
                    cBase.Execute (Cons)
                    
                End If
                rsA.MoveNext
            Loop
            rsA.Close
            
            'Veo si está en un remitoretiro y le modifico el documento
            Cons = "SELECT * FROM EnviosRemitos " _
                & " WHERE EReEnvio = " & rsE("EnvCodigo") & " AND EReNota IS NOT NULL AND EReFactura = " & rsE("Envdocumento")
            Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsA.EOF Then
            
                Set oDoc = New clsDocToPrint
                docToPrint.Add oDoc
                oDoc.Documento = rsA("EReRemito")
                oDoc.idEnvio = rsA("EReEnvio")
                oDoc.TipoDoc = TD_RemitoRetiro
            
                rsA.Edit
                rsA("EReFactura") = oDocCtdo.Codigo
                rsA.Update
            End If
            rsA.Close
            
            
            If lIDVaCon > 0 Then
                'Tengo que cambiar la fecha de todos los envíos del Va con que no sean vtas telefonicas
                Cons = "UPDATE Envio SET EnvFModificacion = '" & Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss") & "'" & _
                    "FROM Envio, EnvioVaCon WHERE EVCID = " & lIDVaCon & " And EVCEnvio = EnvCodigo And EnvCodigo <> " & rsE("EnvCodigo")
                cBase.Execute Cons
            End If
            
        End If
        
        'EDITO EL ENVIO
        rsE.Edit
        If (rsE("EnvFormaPago") = 2 And rsE("EnvTipo") = 3) Or bChangeFP Then
            'Pongo como pago del flete Caja.
            rsE("EnvFormaPago") = 1
        End If
        If bVTaAFact Then
        'Cambio el tipo de documento a contado y pongo el id del documento generado.
            rsE("EnvTipo") = 1
            rsE("EnvDocumento") = lIDDocVta
        End If
        rsE("EnvCodImpresion") = lIDImp
        rsE("EnvEstado") = EstadoEnvio.Impreso
        rsE("EnvFModificacion") = Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss")
        If lIDDocFact > 0 Then rsE("EnvDocumentoFactura") = lIDDocFact
        rsE.Update
        rsE.Close
        
        'Esto no va más lo dejo hasta que funcione lo nuevo.
        Cons = "Update RenglonEnvio Set RevCodImpresion = " & lIDImp & " Where RevEnvio = " & lEnvio
        cBase.Execute Cons
        
        'Si emitio documento lo envío a firmar.
        If Not oDocCtdo Is Nothing Then
            If EmitirCFE(oDocCtdo, CAE) <> "" Then rsE.Close: rsE.Edit
            'cBase.Execute "EXEC prg_PosInsertoDocumentosATickets '" & oDocCtdo.Codigo & "', " & oCnfgPrint.ImpresoraTickets
            Set oDocCtdo = Nothing
        End If
        
    End If

End Sub

'Private Function fnc_InsertarRemito(ByVal lEnvio As Long, ByVal iCliente As Long, ByVal iMoneda As Integer, ByVal iDocCtdo As Long, ByVal iDocRem As Long) As Long
'Dim iIDDoc As Long
'    If iDocRem > 0 Then
'        'Ya lo hice en otro envío
'        iIDDoc = iDocRem
'    Else
'        Dim strDoc As String
'        strDoc = NumeroDocumento(paDRemito)
'        iIDDoc = fnc_InsertoDocumentoBD(TipoDocumento.Remito, Mid(strDoc, 1, 1), Mid(strDoc, 2), iCliente, iMoneda, 0, 0, False, "")
'
'        ReDim Preserve arrDoc(UBound(arrDoc) + 1)
'        With arrDoc(UBound(arrDoc))
'            .Documento = iIDDoc
'            .Envio = lEnvio
'            .tipo = TipoDocumento.Remito
'            .Remito = Mid(strDoc, 1, 1) & "-" & Mid(strDoc, 2)
'        End With
'    End If
'    'Relaciono el documento con el remito.
'    If iDocRem > 0 Then
'        Dim rsR As rdoResultset
'        Set rsR = cBase.OpenResultset("Select * FROM RemitoDocumento Where RDoRemito = " & iDocRem & " And RDoDocumento = " & iDocCtdo, rdOpenDynamic, rdConcurValues)
'        If rsR.EOF Then
'            rsR.AddNew
'            rsR("RDoRemito") = iIDDoc
'            rsR("RDoDocumento") = iDocCtdo
'            rsR.Update
'        End If
'        rsR.Close
'    Else
'        Cons = "INSERT INTO RemitoDocumento (RDoRemito, RDoDocumento) Values (" & iIDDoc & ", " & iDocCtdo & ")"
'        cBase.Execute Cons
'    End If
'    loc_InsertoRenglonesRemito lEnvio, iIDDoc, (iDocRem = 0)
'
'    'Ahora busco articulos especificos asociados al documento y los pongo dentro del remito
'    cBase.Execute "UPDATE ArticuloEspecifico SET AEsTipoDocumento = 1, AEsDocumento = " & iIDDoc & " WHERE AEsTipoDocumento = 1 And AEsDocumento = " & iDocCtdo
'    '......................................................................................................
'
'    fnc_InsertarRemito = iIDDoc
'End Function

Private Sub loc_InsertoRenglonesRemito(ByVal iEnvio As Long, ByVal iRemito As Long, ByVal bNew As Boolean)
Dim rsR As rdoResultset, rsRR As rdoResultset, bIns As Boolean
Dim sQy As String

    sQy = "Select REvArticulo, REvCantidad From RenglonEnvio Where RevEnvio = " & iEnvio
    Set rsR = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    Do While Not rsR.EOF
        bIns = True
        If Not bNew Then
            Cons = "Select RenCantidad From Renglon Where RenDocumento = " & iRemito & " And RenArticulo = " & rsR("REvArticulo")
            Set rsRR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsRR.EOF Then
                rsRR.Edit
                rsRR("RenCantidad") = rsRR("RenCantidad") + rsR(1)
                rsRR.Update
                bIns = False
            Else
                bIns = True
            End If
            rsRR.Close
        End If
        If bIns Then
            sQy = "Insert into Renglon (RenDocumento, RenArticulo, RenCantidad, RenPrecio, RenIVA, RenARetirar) " & _
                "VALUES (" & iRemito & ", " & rsR("REvArticulo") & ", " & rsR(1) & ", 0, 0, 0)"
            cBase.Execute sQy
        End If
        rsR.MoveNext
    Loop
    rsR.Close
End Sub

Private Function fnc_GetIDArticuloFlete(ByVal lTipoFlete As Long) As Long
Dim rsA As rdoResultset
    Cons = "Select TFlArticulo From TipoFlete Where TFlCodigo = " & lTipoFlete
    Set rsA = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    fnc_GetIDArticuloFlete = rsA(0)
    rsA.Close
End Function

Private Function fnc_FacturoVtaTelefonicaEnOtro(ByVal lVtaTelefonica As Long, ByVal lIDVaCon As Long, ByVal oCliFact As clsClienteCFE) As Long
'Retorno el ID del nuevo documento.
Dim rsV As rdoResultset, rsA As rdoResultset
Dim lIDArtF As Long
Dim bCofis As Currency
Dim cImp As Currency, cIVA As Currency, cCofis As Currency
Dim strNroDoc As String
    
    Cons = "Select Envio.* From Envio " & _
                "Where EnvDocumento = " & lVtaTelefonica & " And EnvTipo = 3 And EnvReclamoCobro Is Not Null"
    Set rsV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If rsV("EnvFormaPago") = 2 Then lIDArtF = fnc_GetIDArticuloFlete(rsV("EnvTipoFlete"))
    
    Cons = "Select * From VentaTelefonica Where VTeCodigo = " & rsV("EnvDocumento")
    Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    cImp = rsA("VTeTotal")
    cIVA = rsA("VTeIVA")
                    
    'Pago Domicilio
    If rsV("EnvFormaPago") = 2 Then
        If Not IsNull(rsV("EnvValorFlete")) Then cImp = rsV("EnvValorFlete") + cImp: cIVA = rsV("EnvIvaFlete") + cIVA
        If Not IsNull(rsV("EnvValorPiso")) Then cImp = cImp + rsV("EnvValorPiso"): cIVA = cIVA + rsV("EnvIvaPiso")
    End If
    
    Dim CAE As New clsCAEDocumento
    If Val(prmEFacturaProductivo) = 0 Then
        strNroDoc = NumeroDocumento(paDContado)
        With CAE
            .Desde = 1
            .Hasta = 9999999
            .Serie = Mid(strNroDoc, 1, 1)
            .Numero = Mid(strNroDoc, 2)
            .IdDGI = "9014"
            .TipoCFE = IIf(oCliFact.RUT <> "", CGSA_TiposCFE.CFE_eFactura, CFE_eTicket)
            .Vencimiento = "31/12/" & CStr(Year(Date))
        End With
    Else
        Dim caeG As New clsCAEGenerador
        Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, IIf(oCliFact.RUT <> "", CGSA_TiposCFE.CFE_eFactura, CFE_eTicket), paCodigoDGI)
        Set caeG = Nothing
    End If
    Dim doc As New clsDocumentoCGSA
    With doc
        Set .cliente = oCliFact
        .Emision = gFechaServidor
        .Tipo = TD_Contado
        .Numero = CAE.Numero
        .Serie = CAE.Serie
        .Moneda.Codigo = rsA("VTeMoneda")
        .Total = cImp
        .IVA = cIVA
        .sucursal = paCodigoDeSucursal
        .Digitador = paCodigoDeUsuario
        .Comentario = ""
        .zona = rsV("EnvZona")
        .Vendedor = paCodigoDeUsuario
    End With
    
    Dim oRen As clsDocumentoRenglon
    Set doc.Renglones = CopioRenglonesVentaTelefonica(rsV("EnvDocumento"))
    If rsV("EnvFormaPago") = 2 Then
    
        If Not IsNull(rsV("EnvValorPiso")) Then
            If rsV("EnvValorPiso") > 0 Then
                If oArtPisoAgencia Is Nothing Then Set oArtPisoAgencia = CargoArticulosPrms(paArticuloPisoAgencia)
                Set oRen = New clsDocumentoRenglon
                Set oRen.Articulo = oArtPisoAgencia
                oRen.cantidad = 1
                oRen.IVA = rsV("EnvIvaPiso")
                oRen.Precio = rsV("EnvValorPiso")
                oRen.CantidadARetirar = 1
                oRen.Descripcion = ""
                doc.Renglones.Add oRen
            End If
        End If
        
        If Not IsNull(rsV("EnvValorFlete")) Then
            If rsV("EnvValorFlete") > 0 Then
                If Not oArtFlete Is Nothing Then
                    If oArtFlete.ID <> lIDArtF Then Set oArtFlete = Nothing
                End If
                If oArtFlete Is Nothing Then Set oArtFlete = CargoArticulosPrms(lIDArtF)
            
                Set oRen = New clsDocumentoRenglon
                Set oRen.Articulo = oArtFlete
                oRen.cantidad = 1
                oRen.IVA = rsV("EnvIvaFlete")
                oRen.Precio = rsV("EnvValorFlete")
                oRen.CantidadARetirar = 1
                oRen.Descripcion = ""
                doc.Renglones.Add oRen
            End If
        End If
    End If
    Set doc.Conexion = cBase
    doc.Codigo = doc.InsertoDocumentoBD(0)
            
    Dim oDoc As New clsDocToPrint
    docToPrint.Add oDoc
    With oDoc
'        Set .CAE = CAE
        .Documento = doc.Codigo
        .idEnvio = rsV("EnvCodigo")
    End With

                            
    rsA.Edit
    rsA("VTeDocumento") = doc.Codigo
    rsA("VTeFModificacion") = Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss")
    rsA.Update
    rsA.Close
    
    loc_InsertDocumentoPendiente doc.Codigo, 1, rsV("EnvCodigo"), cImp, rsA("VTeMoneda")
    If doc.Codigo > 0 Then UpdateInstalacionesADocumento doc.Codigo, rsV("EnvDocumento")

    'Tengo que updatear todos los envíos que pertenecen a esta vta.
    Cons = "Select * From Envio " & _
                " Where EnvTipo = " & rsV("EnvTipo") & _
                " And EnvDocumento = " & rsV("EnvDocumento") '& _
                " And EnvCodigo <> " & rsV("EnvCodigo")
    Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsA.EOF
        loc_ChangeFechaEdit rsA("EnvCodigo"), gFechaServidor
        
        'Updateo en la relación del va con el documento
        Cons = "Update EnvioVaCon SET EVCDocumento = " & doc.Codigo & " Where EVCEnvio = " & rsA("EnvCodigo")
        cBase.Execute (Cons)
        
        If rsA("EnvCodigo") <> rsV("EnvCodigo") Then
            rsA.Edit
            rsA("EnvTipo") = 1
            rsA("EnvDocumento") = doc.Codigo
            rsA("EnvFModificacion") = Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss")
            rsA.Update
        End If
        rsA.MoveNext
    Loop
    rsA.Close
    
    loc_ChangeFechaEdit rsV("EnvCodigo"), gFechaServidor
    
    'EDITO EL ENVIO
    rsV.Edit
    If rsV("EnvFormaPago") = 2 Then
        rsV("EnvFormaPago") = 1
        rsV("EnvDocumentoFactura") = doc.Codigo
    End If
    rsV("EnvTipo") = 1
    rsV("EnvDocumento") = doc.Codigo
    rsV("EnvFModificacion") = Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss")
    rsV.Update
    rsV.Close
    
    If EmitirCFE(doc, CAE) <> "" Then rsA.Close: rsA.Edit
    
    fnc_FacturoVtaTelefonicaEnOtro = doc.Codigo
    Set doc = Nothing
    
End Function


Private Sub loc_InsertRenglonEntrega(ByVal lIDImp As Long, ByVal lEnvio As Long, ByVal iCamion As Integer)
Dim rsR As rdoResultset
    Cons = "Select REvAEntregar, REvArticulo From RenglonEnvio " & _
                " Where REvEnvio = " & lEnvio & _
                " And REvArticulo Not IN (Select ArtID From Articulo Where ArtTipo IN(SELECT TipID from  dbo.InTipos(" & paTipoArticuloServicio & ")))"
    Set rsR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsR.EOF
        loc_AddRenglonEntrega lIDImp, rsR("REvArticulo"), rsR("REvAEntregar"), iCamion
        rsR.MoveNext
    Loop
    rsR.Close
    
    Cons = "SELECT * FROM EnvioCodigoImpresion WHERE ECIEnvio = " & lEnvio & " AND ECICodImpresion = " & lIDImp
    Set rsR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsR.EOF Then
        'cons = "INSERT INTO EnvioCodigoImpresion (ECIEnvio, ECICodImpresion) VALUES ( " & lenvio
        rsR.AddNew
        rsR("ECIEnvio") = lEnvio
        rsR("ECICodImpresion") = lIDImp
        rsR.Update
    End If
    rsR.Close
    
End Sub

Private Sub loc_AddRenglonEntrega(ByVal lIDImp As Long, ByVal lArt As Long, ByVal iQ As Integer, ByVal iCamion As Integer)
Dim rsRE As rdoResultset

    Cons = "Select * From RenglonEntrega" & _
                " Where ReECodImpresion = " & lIDImp & _
                " And ReEArticulo = " & lArt
    Set rsRE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsRE.EOF Then
        rsRE.AddNew
        rsRE("ReECantidadTotal") = iQ
        rsRE("ReECantidadEntregada") = 0
    Else
        rsRE.Edit
        rsRE("ReECantidadTotal") = iQ + rsRE("ReECantidadTotal")
    End If
    rsRE("ReECodImpresion") = lIDImp
    rsRE("ReEArticulo") = lArt
    rsRE("ReEFModificacion") = Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss")
    rsRE("ReEUsuario") = paCodigoDeUsuario
    rsRE("ReECamion") = iCamion
    rsRE("ReEEstado") = paEstadoArticuloEntrega
    rsRE.Update
    rsRE.Close
End Sub

Private Sub loc_ImprimoHojaBlanca(ByVal codImpresion As Long)
Dim iQ As Integer
On Error GoTo errPrint
    
    'No hay documentos
    If docCopia.Count = 0 Then Exit Sub
        
    vsPrint.Header = ""
    vsPrint.Footer = ""
    vsPrint.Orientation = orPortrait
    Screen.MousePointer = 11
    
    '5/1/2015 a pedido de Rafa los ordeno por idEnvio.
    'Set docCopia = OrdenoGrillaPorIDEnvio(docCopia)
    bPrintPlanilla = False
    Dim oCopia As clsDocToPrint
    For Each oCopia In docCopia
        If (oCopia.EsConAgencia) Then
            '27/02/2018 si es envío de agencia de omnibus no imprimo.
            Dim rsA As rdoResultset
            Cons = "SELECT IsNull(AgeTipo, 0) FROM Envio LEFT OUTER JOIN Agencia ON EnvAgencia = AgeCodigo WHERE EnvCodigo = " & oCopia.idEnvio
            Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If (rsA(0) <> 1) Then
                ImprimoEFactura oCopia.Documento, oCopia.idEnvio, True, 0
            End If
            rsA.Close
        Else
            ImprimoEFactura oCopia.Documento, oCopia.idEnvio, True, 0
        End If
    Next
    Screen.MousePointer = 0
Exit Sub

errPrint:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al imprimir documentos en hoja en blanco.", Err.Description
    Exit Sub
End Sub

Public Sub EsperaConEvento()
    Dim dHora As Date
    dHora = DateAdd("s", 1, Now)
    Do While dHora > Now
        DoEvents
    Loop
End Sub

Private Sub loc_ImprimoContados()
Dim iQ As Integer
On Error GoTo errPrint

    If docToPrint.Count = 0 Then Exit Sub
        
    Screen.MousePointer = 11
    bPrintPlanilla = False
    Dim oDoc As clsDocToPrint
    For Each oDoc In docToPrint
        If (oDoc.EsConAgencia) Then
            '27/02/2018 si es envío de agencia de omnibus no imprimo.
            Dim rsA As rdoResultset
            Cons = "SELECT IsNull(AgeTipo, 0) FROM Envio LEFT OUTER JOIN Agencia ON EnvAgencia = AgeCodigo WHERE EnvCodigo = " & oDoc.idEnvio
            Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If (rsA(0) <> 1) Then
                ImprimoEFactura oDoc.Documento, oDoc.idEnvio, False, oDoc.TipoDoc
            End If
            rsA.Close
        Else
            ImprimoEFactura oDoc.Documento, oDoc.idEnvio, False, oDoc.TipoDoc
        End If
    Next
    Screen.MousePointer = 0
Exit Sub

errPrint:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al imprimir los documentos, debe acceder a reimprimir aquellos que no fueron impresos", Err.Description
    Exit Sub
End Sub

Public Sub ImprimoEFactura(ByVal doc As Long, ByVal Envio As Long, ByVal EsCopia As Boolean, ByVal TipoDoc As TipoDocumento)
On Error GoTo errIEF
    
    Dim rsd As rdoResultset
    'Valido si está firmado.
    Cons = "SELECT EComID, DocSerie, DocNumero From Documento LEFT OUTER JOIN eComprobantes ON DocCodigo = EComID where eComID = " & doc
    Set rsd = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If IsNull(rsd("EComID")) Then
        MsgBox "El documento " & rsd("DocSerie") & "-" & rsd("DocNumero") & " no está firmado, solicite que se corrija la situación", vbExclamation, "ATENCIÓN"
        rsd.Close
        Exit Sub
    End If
    rsd.Close
        
    Dim oPrintEF As New ComPrintEfactura.ImprimoCFE
    If Not EsCopia Then
        '.SetDevice paIContadoN, paIContadoB, paPrintCtdoPaperSize
        'oPrintEF.ImprimirCFE doc, envio, paIContadoN, paIContadoB, paPrintCtdoPaperSize
        oPrintEF.ImprimirCFE doc, Envio, IIf(TipoDoc = RemitoEntrega Or TipoDoc = RemitoRetiro, paPrintRemRepD, paIContadoN), _
                    IIf(TipoDoc = RemitoEntrega Or TipoDoc = RemitoRetiro, paPrintRemRepB, paIContadoB), paPrintCtdoPaperSize
        
    Else
        If TipoDoc = RemitoEntrega Or TipoDoc = RemitoRetiro Then
            oPrintEF.ImprimirCFE doc, Envio, paPrintRemRepD, paPrintRemRepB, paPrintConfPaperSize
        Else
            oPrintEF.ImprimirCFE doc, Envio, paPrintConfD, printBandejaCopiaeTicket, paPrintConfPaperSize
        End If
    End If
    Set oPrintEF = Nothing
    Exit Sub
    
errIEF:
    objGral.OcurrioError "Error al imprimir eFactura", Err.Description
    oLog.InsertoLog TL_Error, "documento ID=" & doc & " Error: " & Err.Description
End Sub

'Private Sub loc_ShowErrorDocumentosCtdo()
'Dim iQ As Integer
'Dim sDocs As String
'    MsgBox "Error al imprimir los documentos." & vbCr & "A continuación se da un detalle de los documentos que se intentaban imprimir para que pueda reimprimirlos.", vbInformation, "ATENCIÓN"
'
'    For iQ = 1 To UBound(arrDoc)
'        If sDocs <> "" Then sDocs = sDocs & ", "
'        sDocs = sDocs & arrDoc(iQ).Documento
'    Next
'
'    Cons = "Select DocCodigo, Serie = DocSerie, Numero = DocNumero  From Documento " _
'        & " Where DocCodigo In (" & sDocs & ")"
'    Dim objLista As New clsListadeAyuda
'    objLista.ActivarAyuda cBase, Cons, 2450, 1, "Documentos a reimprimir"
'    Set objLista = Nothing
'    Me.Refresh
'End Sub

Private Sub loc_ChangeFechaEdit(ByVal lEnvio As Long, ByVal sFEdit As String)
    Dim iQ As Integer
    With vsGrid
        For iQ = .FixedRows To .Rows - 1
            If lEnvio = .Cell(flexcpValue, iQ, 0) Then
                .Cell(flexcpData, iQ, 4) = sFEdit
                Exit For
            End If
        Next
    End With
End Sub

Private Sub db_InsertoRenglonesVTaTelefonica(ByVal lIDDoc As Long, ByVal lIDVta As Long)
Dim rsR As rdoResultset
    Cons = "Select * From RenglonVtaTelefonica Where RVTVentaTelefonica = " & lIDVta
    Set rsR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsR.EOF
        loc_InsertRenglonDocumento lIDDoc, rsR("RVtArticulo"), rsR("RVTCantidad"), rsR("RVTPrecio"), rsR("RVTIVA"), rsR("RVTARetirar")
        rsR.MoveNext
    Loop
    rsR.Close
End Sub

Private Function CopioRenglonesVentaTelefonica(ByVal idVtaTelef As Long) As Collection
Dim rsR As rdoResultset
Dim oRen As clsDocumentoRenglon
    
    Set CopioRenglonesVentaTelefonica = New Collection
    Cons = "SELECT RVTCantidad, RVTPrecio, RVTIva, RVTARetirar, ArtId, ArtTipo, IvaPorcentaje, AESID, IsNull(AEsNombre, ArtNombre) ArtNombre " & _
        "FROM RenglonVtaTelefonica INNER JOIN Articulo ON RVTArticulo = ArtId INNER JOIN ArticuloFacturacion ON AFaArticulo = ArtId " & _
        "INNER JOIN TipoIva ON IvaCodigo = AFaIva " & _
        "LEFT OUTER JOIN ArticuloEspecifico ON AEsTipoDocumento IN (7) And AEsDocumento = RVTVentaTelefonica AND AEsArticulo = ArtID " & _
        "WHERE RVTVentaTelefonica = " & idVtaTelef
    Set rsR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsR.EOF
        'loc_InsertRenglonDocumento lIDDoc, rsR("RVtArticulo"), rsR("RVTCantidad"), rsR("RVTPrecio"), rsR("RVTIVA"), rsR("RVTARetirar")
        Set oRen = New clsDocumentoRenglon
        With oRen
            .Articulo.ID = rsR("ArtID")
            If Not IsNull(rsR("AESID")) Then .Articulo.IDEspecifico = rsR("AESID")
            .Articulo.Nombre = Trim(rsR("ArtNombre"))
            .Articulo.TipoIVA.Porcentaje = rsR("IvaPorcentaje")
            .Articulo.TipoArticulo = rsR("ArtTipo")
            .cantidad = rsR("RVTCantidad")
            .CantidadARetirar = rsR("RVTARetirar")
            .IVA = rsR("RVTIva")
            .Precio = rsR("RVTPrecio")
        End With
        CopioRenglonesVentaTelefonica.Add oRen
        rsR.MoveNext
    Loop
    rsR.Close
    
End Function

Private Sub db_FacturoDiferenciaEnvio(ByVal lEnvio As Long, ByVal lIDVaCon As Long, ByVal clienteEnv As clsClienteCFE)
Dim rsd As rdoResultset
Dim cImp As Currency, cIVA As Currency
Dim strNroDoc As String
Dim tipoCAE As Byte

    Dim oDoc As New clsDocToPrint

    Dim caeG As New clsCAEGenerador
    tipoCAE = IIf(clienteEnv.RUT <> "", 111, 101)
    Dim CAE As clsCAEDocumento
    Dim doc As clsDocumentoCGSA
    Dim oRen As clsDocumentoRenglon
    
'FORMA DE PAGO = 2 --> Domicilio.
    Cons = "SELECT DiferenciaEnvio.*, EnvCliente, ISNULL(EnvAgencia, 0) EnvAgencia FROM DiferenciaEnvio, Envio " & _
            "WHERE EnvCodigo = " & lEnvio _
                & " AND EnvEstado = " & EstadoEnvio.AImprimir _
                & " AND EnvTipo In (1, 2)" _
            & " AND DEvFormaPago = 2" & _
            " AND DEvDocumento Is Null And DevEnvio = EnvCodigo"

    Set rsd = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
    Do While Not rsd.EOF

        If Not IsNull(rsd!DevValorFlete) Then cImp = rsd!DevValorFlete: cIVA = rsd!DevIvaFlete
        If Not IsNull(rsd!DevValorPiso) Then cImp = cImp + rsd!DevValorPiso: cIVA = cIVA + rsd!DevIvaPiso
        
        'Si es la primera vez que lo voy a utilizar.
        If oArtDifEnvio Is Nothing Then Set oArtDifEnvio = CargoArticulosPrms(paArticuloDiferenciaEnvio)
        
        Set CAE = New clsCAEDocumento
        Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, tipoCAE, paCodigoDGI)
        Set doc = New clsDocumentoCGSA
        With doc
            Set .cliente = clienteEnv
            .Emision = gFechaServidor
            .Tipo = TD_Contado
            .Numero = CAE.Numero
            .Serie = CAE.Serie
            .Moneda.Codigo = rsd("DevMoneda")
            .Total = cImp
            .IVA = cIVA
            .sucursal = paCodigoDeSucursal
            .Digitador = paCodigoDeUsuario
            .Comentario = ""
            .Vendedor = paCodigoDeUsuario
        End With
        Set oRen = New clsDocumentoRenglon
        Set oRen.Articulo = oArtDifEnvio
        oRen.cantidad = 1
        oRen.IVA = cIVA
        oRen.Precio = cImp
        oRen.CantidadARetirar = 1
        oRen.Descripcion = ""
        doc.Renglones.Add oRen
        Set doc.Conexion = cBase
        doc.Codigo = doc.InsertoDocumentoBD(0)
        If EmitirCFE(doc, CAE) <> "" Then rsd.Close: rsd.Edit
        
        loc_InsertDocumentoPendiente doc.Codigo, 1, rsd("DevEnvio"), cImp, rsd("DevMoneda")
                    
        'Le asocio el documento a la diferencia de Envio.
        Cons = "Update DiferenciaEnvio Set DEvDocumento = " & doc.Codigo & "Where DEvCodigo = " & rsd("DEvCodigo")
        cBase.Execute Cons
        
        Set oDoc = New clsDocToPrint
        docToPrint.Add oDoc
'        Set oDoc.CAE = CAE
        oDoc.Documento = doc.Codigo
        oDoc.idEnvio = rsd("DevEnvio")
        If rsd("EnvAgencia") > 0 Then
            oDoc.EsConAgencia = True
        End If
        Set oDoc = Nothing
        rsd.MoveNext
    Loop
    rsd.Close
    
    Set caeG = Nothing
    
End Sub

Private Sub loc_InsertRenglonDocumento(ByVal lDoc As Long, ByVal lArt As Long, ByVal iQ As Integer, ByVal cPrecio As Currency, ByVal cIVA As Currency, ByVal iARetirar As Integer)
    Cons = "INSERT INTO Renglon (RenDocumento, RenArticulo, RenCantidad, RenPrecio, RenIVA, RenARetirar)" _
            & " VALUES (" & lDoc & ", " & lArt & ", " & iQ & ", " & cPrecio _
            & ", " & cIVA & ", " & iARetirar & ")"
        cBase.Execute (Cons)
End Sub

Private Function fnc_InsertoDocumentoBD(ByVal iTipoDocumento As Integer, ByVal serieDoc As String, ByVal numeroDoc As Long, _
                                                        ByVal lCliente As Long, ByVal iMoneda As Integer, ByVal cTotal As Currency, _
                                                        ByVal cIVA As Currency, Anulado As Boolean, _
                                                        ByVal Comentario As String, Optional lZona As Long = 0)

Dim RsDoc As rdoResultset
    
    Cons = "INSERT INTO Documento (DocFecha, DocTipo, DocSerie, DocNumero, DocCliente, DocMoneda, DocTotal, DocIVA, DocAnulado, DocSucursal, DocUsuario, DocFModificacion, DocZona, DocVendedor, DocComentario) " _
        & "VALUES ('" & Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss") & "'" _
        & ", " & iTipoDocumento _
        & ", '" & serieDoc & "'" _
        & ", " & numeroDoc _
        & ", " & lCliente & ", " & iMoneda _
        & ", " & cTotal & ", " & cIVA _
        & IIf(Anulado, ", 1", ", 0") _
        & ", " & paCodigoDeSucursal & ", " & paCodigoDeUsuario & ",'" & Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss") & "'" _
        & ", " & IIf(lZona > 0, lZona, " Null") & ", " & paCodigoDeUsuario _
        & ", " & IIf(Comentario <> "", "'" & Comentario & "'", " Null") _
        & ")"
    cBase.Execute Cons
    
    Cons = "SELECT MAX(DocCodigo) From Documento" _
        & " WHERE DocTipo = " & iTipoDocumento _
        & " AND DocSerie = '" & serieDoc & "'" _
        & " AND DocNumero = " & numeroDoc
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    fnc_InsertoDocumentoBD = RsDoc(0)
    RsDoc.Close

End Function

Private Sub loc_InsertDocumentoPendiente(ByVal lDoc As Long, ByVal iTipo As Integer, ByVal lIDTipo As Long, ByVal cImporte As Currency, ByVal iMon As Integer)
Dim m_Disponibilidad As Long
    m_Disponibilidad = modMaeDisponibilidad.dis_DisponibilidadPara(paCodigoDeSucursal, CLng(iMon))
    Cons = "Insert Into DocumentoPendiente (DPeDocumento, DPeTipo, DPeIDTipo, DPeImporte, DPeMoneda, DPeDisponibilidad) Values (" & _
                    lDoc & ", " & iTipo & ", " & lIDTipo & ", " & Format(cImporte, "###0.00") & ", " & iMon & ", " & m_Disponibilidad & ")"
    cBase.Execute Cons
End Sub

Private Function fnc_FrenoPorVentas40KUI(ByVal sEnvios As String) As Boolean
    
    fnc_FrenoPorVentas40KUI = True
    Cons = "Select EnvCodigo From  Envio INNER JOIN VentaTelefonica ON VTeCodigo = EnvDocumento AND VTeTotal > 40000 * " & paValorUIUltMes & _
        " Where EnvCodigo In (" & sEnvios & ") And EnvReclamoCobro Is Not Null " & _
        "And EnvTipo = 3 "
    Dim rsV As rdoResultset
    Set rsV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    sEnvios = ""
    Do While Not rsV.EOF
        If sEnvios <> "" Then sEnvios = sEnvios & ", "
        sEnvios = sEnvios & rsV("EnvCodigo")
        rsV.MoveNext
    Loop
    rsV.Close
    
    Screen.MousePointer = 0
    If sEnvios <> "" Then
        MsgBox "Los siguientes envíos superan las 40000 UI a facturar.", vbExclamation, "ATENCIÓN"
        fnc_FrenoPorVentas40KUI = False
    End If
    
    
End Function

Private Function fnc_ControlVtaTelCobranzaEnOtroEnvio(ByVal sEnvios As String) As Boolean
On Error GoTo errCVT
Dim rsV As rdoResultset

    Screen.MousePointer = 11
    fnc_ControlVtaTelCobranzaEnOtroEnvio = True
    'Busco si tengo en alguno de estos envíos alguno que sea de vta telefónica y no tenga el cobro.
    Cons = "Select EnvCodigo, EnvDocumento From Envio " & _
            "Where EnvCodigo Not In (" & sEnvios & ") And EnvReclamoCobro Is Not Null " & _
            "And EnvTipo = 3 And EnvDocumento In(Select EnvDocumento From Envio " & _
                " Where EnvCodigo IN (" & sEnvios & ") And EnvTipo = 3 And EnvReclamoCobro Is Null)"

    Set rsV = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    sEnvios = ""
    Do While Not rsV.EOF
        If sEnvios <> "" Then sEnvios = sEnvios & ", "
        sEnvios = sEnvios & rsV("EnvCodigo")
        rsV.MoveNext
    Loop
    rsV.Close
    Screen.MousePointer = 0
    If sEnvios <> "" Then
        'fnc_ControlVtaTelCobranzaEnOtroEnvio = (MsgBox("Existen envíos en la lista que tienen la cobranza de la venta telefónica en otro envío que aún no fue impreso." & vbCr & _
                        " Envíos: " & sEnvios & _
                        vbCr & vbCr & "Si continúa se imprimirá un contado para facturar dicha venta telefónica." & vbCr & vbCr & "¿Desea continuar?", vbExclamation + vbYesNo, "Posible Error") = vbYes)
        MsgBox "Existen envíos en la lista que tienen la cobranza de la venta telefónica en otro envío que aún no fue impreso." & vbCr & _
            " Envíos: " & sEnvios & _
            vbCr & vbCr & "Debe imprimir el envío que cobra la venta primero.", vbExclamation, "Ventas telefónicas"
        fnc_ControlVtaTelCobranzaEnOtroEnvio = False
    End If
    Exit Function
errCVT:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al validar la cobranza de ventas telefónicas.", Err.Description

End Function

Private Sub s_SaveAsignoCamionAEnvio()
'Ocurre al apretar grabar en toolbar.
Dim iQ As Integer
    With vsGrid
        For iQ = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, iQ, 0) = flexChecked Then
                If Val(.Cell(flexcpData, iQ, 3)) = 0 Then
                'Si no tiene camión se lo asigno.
                    .Cell(flexcpData, iQ, 3) = f_GetCamionAutomático(vsGrid.Cell(flexcpValue, iQ, 0))
                End If
                If Val(.Cell(flexcpData, iQ, 3)) > 0 Then
                    .RowHidden(iQ) = db_EnvioAsignoCamion(.Cell(flexcpData, iQ, 3), iQ)
                End If
            End If
        Next
        'Cargo de nuevo
        s_FillGrid
    End With
End Sub

Private Function db_EnvioAsignoCamion(ByVal lCamion As Long, ByVal iRowGrid As Integer) As Boolean
Dim rsE As rdoResultset
On Error GoTo errAC
    FechaDelServidor
    '.......................................................................................
    'ATENCIÓN
        'Si un envío es con Va Con cambio todos.
    '.......................................................................................
    db_EnvioAsignoCamion = False
    Screen.MousePointer = 11
    Cons = "Select * From Envio Where EnvCodigo = " & vsGrid.Cell(flexcpValue, iRowGrid, 0)
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsE.EOF Then
        If CDate(vsGrid.Cell(flexcpData, iRowGrid, 1)) <> rsE!EnvFModificacion Then
            Screen.MousePointer = 0
            rsE.Close
            MsgBox "El envío " & vsGrid.Cell(flexcpValue, iRowGrid, 0) & " fue modificado, no se asigno.", vbExclamation, "Atención"
        Else
            If Val(vsGrid.Cell(flexcpData, iRowGrid, 2)) > 0 Then
                rsE.Close
                db_EnvioAsignoCamion = db_CambioEstadoEnvioVaCon(EstadoEnvio.AImprimir, lCamion, Val(vsGrid.Cell(flexcpData, iRowGrid, 2)))
            Else
                If Not IsNull(rsE("EnvZona")) Then
                    rsE.Edit
                    rsE!EnvCamion = lCamion
                    rsE!EnvEstado = EstadoEnvio.AImprimir
                    rsE!EnvFModificacion = Format(gFechaServidor, cte_FormatFH)
                    rsE.Update
                    rsE.Close
                    db_EnvioAsignoCamion = True
                End If
            End If
        End If
    Else
        rsE.Close
        Screen.MousePointer = 0
        MsgBox "El envío " & vsGrid.Cell(flexcpValue, iRowGrid, 0) & " fue eliminado.", vbExclamation, "Atención"
    End If
    Screen.MousePointer = 0
Exit Function
errAC:
End Function

Private Sub db_EnvioAAsignar(Optional sF As String)
Dim rsE As rdoResultset
Dim bUnoSolo As Boolean, bEdit As Boolean, bDel As Boolean

    FechaDelServidor
    Screen.MousePointer = 11
    bUnoSolo = (vsGrid.SelectedRows = 1)
    
    '.......................................................................................
    'ATENCIÓN
        'Si un envío es con Va Con cambio todos.
    '.......................................................................................
    Do While (vsGrid.SelectedRows > 0)
    
        Cons = "Select * From Envio Where EnvCodigo = " & vsGrid.Cell(flexcpValue, vsGrid.SelectedRow(0), 0)
        Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsE.EOF Then
            If CDate(vsGrid.Cell(flexcpData, vsGrid.SelectedRow(0), 1)) <> rsE!EnvFModificacion Then
                Screen.MousePointer = 0
                rsE.Close
                If bUnoSolo Then
                    MsgBox "El envío fue modificado refresque la información.", vbExclamation, "Atención"
                Else
                    bEdit = True
                End If
                vsGrid.RemoveItem vsGrid.SelectedRow(0)
            Else
                If Val(vsGrid.Cell(flexcpData, vsGrid.SelectedRow(0), 2)) > 0 Then
                    rsE.Close
                    If db_CambioEstadoEnvioVaCon(EstadoEnvio.AImprimir, 0, Val(vsGrid.Cell(flexcpData, vsGrid.SelectedRow(0), 2)), sF) Then
                        vsGrid.RemoveItem vsGrid.SelectedRow(0)
                    End If
                Else
                    If Not IsNull(rsE("EnvZona")) Then
                        On Error GoTo ErrBT
                        rsE.Edit
                        rsE!EnvCamion = Null
                        rsE!EnvEstado = EstadoEnvio.AImprimir
                        If IsDate(sF) Then rsE!EnvFechaPrometida = Format(sF, "mm/dd/yyyy")
                        rsE!EnvFModificacion = Format(gFechaServidor, cte_FormatFH)
                        rsE.Update
                        rsE.Close
                        vsGrid.RemoveItem vsGrid.SelectedRow(0)
                    End If
                End If
            End If
        Else
            rsE.Close
            Screen.MousePointer = 0
            vsGrid.RemoveItem vsGrid.SelectedRow(0)
            If bUnoSolo Then MsgBox "El envío fue eliminado.", vbExclamation, "Atención" Else bDel = True
        End If
    Loop
    s_SetMenu
    If (bDel Or bEdit) And Not bUnoSolo Then MsgBox "Algunos envíos fueron eliminados o modificados por otra terminal.", vbExclamation, "Atención"
    Screen.MousePointer = 0
    Exit Sub
ErrBT:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al cambiar de estado al envío.", Err.Description
    Exit Sub
End Sub

Private Sub db_EnvioAConfirmar(Optional sF As String)
Dim bEdit As Boolean, bDel As Boolean
    
    FechaDelServidor
    Screen.MousePointer = 11
    
    If Val(hlTab(0).Tag) <> 0 Then
        Do While (vsGrid.SelectedRows > 0)
            '.......................................................................................
            'ATENCIÓN
                'Si un envío es con Va Con cambio todos.
            '.......................................................................................
            db_SaveEnvioAConfirmar vsGrid.SelectedRow(0), True, bDel, bEdit
        Loop
        s_SetMenu
    Else
        db_SaveEnvioAConfirmar vsGrid.Row, False, bDel, bEdit
        If bDel Or bEdit Then
            s_FillGrid
        Else
            s_SetMenu
        End If
    End If
    
    If bDel Or bEdit Then MsgBox "Algunos envíos fueron eliminados o modificados por otra terminal.", vbExclamation, "Atención"
    Screen.MousePointer = 0
    Exit Sub
ErrBT:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al cambiar de estado al envío.", Err.Description
    Exit Sub
End Sub

Private Sub db_SaveEnvioAConfirmar(ByVal iRow As Integer, ByVal bUnoSolo As Boolean, ByRef bDel As Boolean, ByRef bEdit As Boolean)
Dim rsE As rdoResultset

    Cons = "Select * From Envio Where EnvCodigo = " & vsGrid.Cell(flexcpValue, iRow, 0)
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsE.EOF Then
        If CDate(vsGrid.Cell(flexcpData, iRow, 1)) <> rsE!EnvFModificacion Then
            Screen.MousePointer = 0
            rsE.Close
            If bUnoSolo Then MsgBox "El envío fue modificado, se refrescará la información.", vbExclamation, "Atención" Else bEdit = True
        Else
            If Val(vsGrid.Cell(flexcpData, iRow, 2)) > 0 Then
                rsE.Close
                If db_CambioEstadoEnvioVaCon(AConfirmar, 0, Val(vsGrid.Cell(flexcpData, iRow, 2))) Then
                    vsGrid.RemoveItem iRow
                End If
            Else
                rsE.Edit
                rsE!EnvCamion = Null
                rsE!EnvEstado = EstadoEnvio.AConfirmar
                rsE!EnvFModificacion = Format(gFechaServidor, cte_FormatFH)
                rsE.Update
                rsE.Close
                vsGrid.RemoveItem iRow
            End If
        End If
    Else
        rsE.Close
        Screen.MousePointer = 0
        If bUnoSolo Then MsgBox "El envío fue eliminado.", vbExclamation, "Atención" Else bDel = True
        vsGrid.RemoveItem iRow
    End If

End Sub

Private Sub s_GridFiltrar(ByVal bSel As Boolean)
Dim sFiltro As String
Dim iR As Integer, iC As Integer
    With vsGrid
        If .Rows = .FixedRows Then Exit Sub
        If .Row < .FixedRows Then Exit Sub
        Screen.MousePointer = 11
        iC = .Col
        sFiltro = .Cell(flexcpText, .Row, iC)
        iR = .FixedRows
        Do While iR <= .Rows - 1
            If bSel Then
                If .Cell(flexcpText, iR, iC) <> sFiltro Then .RemoveItem iR Else iR = iR + 1
            Else
                If .Cell(flexcpText, iR, iC) = sFiltro Then .RemoveItem iR Else iR = iR + 1
            End If
        Loop
        s_SetMenu
        Screen.MousePointer = 0
    End With
End Sub
Private Sub db_FillArrayCamiones()
On Error GoTo errFAC
Dim iQ As Integer
    iQ = 0
    ReDim arrCamion(iQ)
    Cons = "Select CamID, rTrim(CamNombre), CamHabilitado From Camiones Order By CamNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        iQ = iQ + 1
        ReDim Preserve arrCamion(iQ)
        With arrCamion(iQ)
            .Codigo = RsAux(0)
            .Nombre = RsAux(1)
            .Habilitado = RsAux(2)
        End With
        If RsAux(2) Then
            Load MnuGridCamion(iQ + 1)
            With MnuGridCamion(iQ + 1)
                .Visible = True
                .Caption = RsAux(1)
                .Tag = RsAux(0)
            End With
                        
            If MnuGridAsiTodosCamion(0).Tag <> "" Then
                Load MnuGridAsiTodosCamion(MnuGridAsiTodosCamion.UBound + 1)
            End If
            With MnuGridAsiTodosCamion(MnuGridAsiTodosCamion.UBound)
                .Visible = True
                .Caption = RsAux(1)
                .Tag = RsAux(0)
            End With
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    Exit Sub
errFAC:
    objGral.OcurrioError "Error al cargar el array con los camiones.", Err.Description
End Sub

Private Sub s_SetFontTab(ByVal Index As Byte)
    With hlTab(hlTab(0).Tag)
        .Font.Bold = False
        .FontOver.Bold = False
        .Refresh
    End With
    hlTab(0).Tag = Index
    With hlTab(Index)
        .Font.Bold = True
        .FontOver.Bold = True
        .Refresh
    End With
    s_SetColor
End Sub

Private Sub s_SetTitle()
    Select Case Val(hlTab(0).Tag)
        Case 0
            lTitle.Caption = "Envíos sin camión asignado a ser impresos el día " & tFecha.Text
        Case 1
            lTitle.Caption = "Envíos asignados al camión a ser impresos el día " & tFecha.Text
        Case 2
            lTitle.Caption = "Envíos que están en el camión para ser entregados el día " & tFecha.Text
        Case 3
            lTitle.Caption = "Envíos con el estado a confirmar"
        Case 4
            lTitle.Caption = "Estadísticas de Envíos para Asignar y Asignados para la fecha."
        Case 5
            lTitle.Caption = "Recepción de envíos"
        Case 7
            lTitle.Caption = "Posibles envíos duplicados"
    End Select
End Sub

Private Sub s_FillGrid()
Dim sValue As String
On Error GoTo errFG
    sbStatus.Panels("envio").Text = "Envíos: "
    sbStatus.Panels("tipo").Text = ""
    sbStatus.Panels("articulo").Text = ""
    sbStatus.Panels("bultos").Text = ""
    Set colBultos = New Collection
    
    picCodImpresion.Visible = (Val(hlTab(0).Tag) = 5)
'Reposiciono
    Form_Resize
    
    iCamSelect = 0
    'Armo Grilla según opción seleccionada en el tag del tab 0
    With vsGrid
        .Editable = False
        .Rows = 1: .Cols = 5: .ColHidden(0) = False: .ColHidden(1) = False: .ColHidden(2) = False: .ColHidden(3) = False: .ColHidden(4) = False
        .Tag = ""
        .BackColorAlternate = &HEFEFEF
        .MergeCells = flexMergeFree
        .AllowSelection = False
        .SelectionMode = flexSelectionByRow
        
        Select Case Val(hlTab(0).Tag)
            Case 0
                .FormatString = "Envío|<Horario|<T|<Flete|<Camión Sugerido|<Zona/Agencia|<Dirección|<Comentario|<Artículos|Tipo"
                sValue = GetSetting(App.Title, "Grid0", "SizeGrid", "1000|1000|300|800|1300|700:0:2000|1000:0:3500|1000:0:3500|1000:0:3500|1500")
                If UBound(Split(sValue, "|")) <> .Cols - 1 Then sValue = "1000|1000|300|800|1300|700:0:2000|1000:0:3500|1000:0:3500|1000:0:3500|1500"
                .Tag = sValue
            
            Case 1, 2, 5, 7
                If Val(hlTab(0).Tag) = 1 Then
                    .AllowSelection = True
                    .SelectionMode = flexSelectionListBox
                End If
                .FormatString = "Envío|Orden|Horario|Flete|Zona/Agencia|Dirección|Comentario|Artículos|Tipo"
                sValue = GetSetting(App.Title, "Grid" & hlTab(0).Tag, "SizeGrid", "900|700|1000|800|700:0:2000|1000:0:3500|1000:0:3500|1000:0:3500|1500")
                If UBound(Split(sValue, "|")) <> .Cols - 1 Then sValue = "900|700|1000|800|700:0:2000|1000:0:3500|1000:0:3500|1000:0:3500|1500"
                .ColHidden(1) = True
                .Tag = sValue
                
                .ColHidden(1) = (Val(hlTab(0).Tag) = 5 Or Val(hlTab(0).Tag) = 7)
                .ColHidden(2) = (Val(hlTab(0).Tag) = 5 Or Val(hlTab(0).Tag) = 7)
                .ColHidden(3) = (Val(hlTab(0).Tag) = 5 Or Val(hlTab(0).Tag) = 7)
                .ColHidden(4) = (Val(hlTab(0).Tag) = 5 Or Val(hlTab(0).Tag) = 7)
                
            Case 3
                .FormatString = "Envío|Fecha|Horario|Flete|Zona/Agencia|Dirección|Comentario|Artículos|Tipo"
                sValue = GetSetting(App.Title, "Grid" & hlTab(0).Tag, "SizeGrid", "900|1000|1000|800|700:0:2000|1000:0:3500|1000:0:3500|1000:0:3500|1500")
                .Tag = sValue
                .AllowSelection = True
                .SelectionMode = flexSelectionListBox
                
            Case 4
                .FormatString = "Flete|Camión|Horario|Zona/Agencia|Cantidad"
                sValue = GetSetting(App.Title, "Grid" & hlTab(0).Tag, "SizeGrid", "1400|1400|700|1400|1000")
                .Tag = sValue
                
            Case 6
                .Cols = 1
                .FormatString = "Camión|Artículo|>Cantidad"
                .ColWidth(0) = 2000: .ColWidth(1) = 4000
                sValue = "" 'GetSetting(App.title, "Grid" & hlTab(0).Tag, "SizeGrid", "1400|1400|700|1400|1000")
                .Tag = sValue
        End Select
        
        .MergeCol(0) = False
        .MergeCol(1) = False
        If .Cols > 3 Then .MergeCol(3) = False
    End With
    loc_SetSizeColGrid
    loc_SaveSettingColGrid
    
    With tFecha
        .Enabled = True
        .BackColor = vbWhite
    End With
    CantidadPosiblesDuplicados
    
    If Val(hlTab(0).Tag) <> 3 And Not tFecha.HasValue Then Exit Sub
    
    
    Select Case Val(hlTab(0).Tag)
        Case 0
            db_FillGridAAsignar
            vsGrid.Editable = True
            
        Case 1
            db_FillCamionesAImprimirImpreso AImprimir
            
        Case 2
            db_FillCamionesAImprimirImpreso Impreso
            
        Case 3
            With tFecha
                .Enabled = False: .BackColor = vbButtonFace
            End With
            db_FillGridAConfirmar
        
        Case 4
            db_FillEstadistica
            
        Case 5
            With tFecha
                .Enabled = False: .BackColor = vbButtonFace
            End With
            loc_PresentoRecepcion
            
        Case 6
            db_FillMercaderiaAReclamar
        
        Case 7
            db_FillPosiblesDuplicados
    End Select
    s_SetMenu
    Exit Sub
errFG:
    objGral.OcurrioError "Error al llenar la grilla.", Err.Description, Me.Caption
    
End Sub

Private Sub s_SetColor()
    
    s_ChangeData
    
    Select Case Val(hlTab(0).Tag)
        Case 0
            picLink.BackColor = &HB88D7B   ' &H900000        '&HB19365
        Case 1
            picLink.BackColor = 224807
        Case 2
            picLink.BackColor = &HA0A0&     '&H8080&
        Case 3
            picLink.BackColor = &H80&
        Case 4
            picLink.BackColor = &H60A4FA
        Case 5
            picLink.BackColor = &H608600     '&H336633 '&H40C0&
        Case 6
            picLink.BackColor = &HFFC0C0
        Case 7
            picLink.BackColor = &H8A9E82
        
    End Select

    hlTab(0).BackColor = picLink.BackColor
    hlTab(1).BackColor = picLink.BackColor
    hlTab(2).BackColor = picLink.BackColor
    hlTab(3).BackColor = picLink.BackColor
    hlTab(4).BackColor = picLink.BackColor
    hlTab(5).BackColor = picLink.BackColor
    hlTab(6).BackColor = picLink.BackColor
    hlTab(7).BackColor = picLink.BackColor
    lTitle.ForeColor = picLink.BackColor
    Me.Refresh
End Sub

Private Function f_GetDireccionRsAux() As String
    
    If Not IsNull(RsAux!CalNombre) Then
        If RsAux("LocCodigo") > 1 Then f_GetDireccionRsAux = f_GetDireccionRsAux & "(" & Trim(RsAux("LocNombre")) & ") "
        f_GetDireccionRsAux = f_GetDireccionRsAux & Trim(RsAux!CalNombre) & " " & Trim(RsAux!DirPuerta)
        If Not IsNull(RsAux!DirLetra) Then f_GetDireccionRsAux = f_GetDireccionRsAux & " " & Trim(RsAux!DirLetra)
        If Not IsNull(RsAux!DirApartamento) Then f_GetDireccionRsAux = f_GetDireccionRsAux & " / " & Trim(RsAux!DirApartamento)
    End If
    
End Function

Private Sub db_FillCamionesAImprimirImpreso(ByVal eEstado As EstadoEnvio)
On Error GoTo errFG
Dim iQ As Integer, iLastC As Long
Dim iQ1 As Integer
Dim sCam As String
    
    Screen.MousePointer = 11
    iQ = 0
    s_DeleteMenuCamion
    Cons = "Select CamCodigo, rTrim(CamNombre) CamNombre, Count(*)" _
            & " From Envio, Camion " _
            & " Where EnvFechaPrometida = '" & Format(tFecha.Value, "mm/dd/yyyy") & "'" _
            & " And EnvEstado = " & eEstado & " And EnvDocumento Is Not Null And EnvCamion = CamCodigo" _
            & " And EnvCodigo Not IN (SELECT EVCEnvio From EnvioVaCon)" _
            & " Group by CamCodigo, CamNombre" _
            & " UNION ALL " _
            & "Select CamCodigo, rTrim(CamNombre) CamNombre, Count(DISTINCT(EVCID))" _
            & " From Envio, EnvioVaCon, Camion " _
            & " Where EnvFechaPrometida = '" & Format(tFecha.Value, "mm/dd/yyyy") & "'" _
            & " And EnvEstado = " & eEstado & " And EnvDocumento Is Not Null And EnvCamion = CamCodigo" _
            & " And EnvCodigo = EVCEnvio Group by CamCodigo, CamNombre ORDER BY CamNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        Do While Not RsAux.EOF
            If iLastC <> RsAux(0) Then
                If iLastC <> 0 Then
                    If iQ > 0 Then
                        Load MnuAIECamion(iQ)
                        MnuAIECamion(iQ).Visible = True
                    End If
                    MnuAIECamion(iQ).Caption = sCam & "( " & iQ1 & ")" 'RsAux(1) & " (" & RsAux(2) & ")"
                    MnuAIECamion(iQ).Tag = iLastC 'RsAux(0)
                    iQ = iQ + 1
                    iQ1 = 0
                    sCam = ""
                End If
                iLastC = RsAux(0)
            End If
            sCam = RsAux(1)
            iQ1 = iQ1 + RsAux(2)
            RsAux.MoveNext
        Loop
        RsAux.Close
        If iQ > 0 Then
            Load MnuAIECamion(iQ)
            MnuAIECamion(iQ).Visible = True
        End If
        MnuAIECamion(iQ).Caption = sCam & "( " & iQ1 & ")" 'RsAux(1) & " (" & RsAux(2) & ")"
        MnuAIECamion(iQ).Tag = iLastC 'RsAux(0)
        If iQ = 0 Then iQ = 1
    End If
    Screen.MousePointer = 0
    If iQ = 0 Then
        'MsgBox "No hay envíos para la fecha seleccionada.", vbExclamation, "Atención"
    ElseIf eEstado = AImprimir Then
        PopupMenu MnuCamion, , hlTab(1).Left, hlTab(1).Top + hlTab(1).Height + picLink.Top
    Else
        PopupMenu MnuCamion, , hlTab(2).Left, hlTab(2).Top + hlTab(2).Height + picLink.Top
    End If
    Exit Sub
errFG:
    vsGrid.Redraw = True
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al cargar la lista a confirmar", Err.Description
End Sub

Private Sub db_FillMercaderiaAReclamar()
On Error GoTo errFM
    Screen.MousePointer = 11
    vsGrid.BackColorAlternate = vsGrid.BackColor
    Cons = "Select ArtID, ArtNombre , CamCodigo, CamNombre, Sum(MRCCantidad) as Cantidad, 1 Tipo" & _
            " From (MercaderiaReclamarCamion INNER JOIN Articulo ON ArtID = MRCArticulo)" & _
            " INNER JOIN Camion ON MRCCamion = CamCodigo" & _
            " WHERE MRCDevuelto Is Null Group by ArtID, ArtNombre, CamCodigo, CamNombre " & _
        " UNION ALL " & _
            "SELECT ArtID, ArtNombre, CamCodigo, CamNombre, Sum(RenARetirar) as Cantidad, 2 Tipo " & _
            "FROM envio INNER JOIN Camion ON EnvCamion = CamCodigo INNER JOIN EnviosRemitos ON EnvCodigo = EReEnvio " & _
            "INNER JOIN Documento ON EReRemito = DocCodigo and DocTipo = 48 INNER JOIN Renglon ON RenDocumento = DocCodigo AND RenARetirar > 0 " & _
            "INNER JOIN Articulo ON ArtId = REnArticulo WHERE EnvEstado in (3,4) Group by ArtID, ArtNombre, CamCodigo, CamNombre ORDER BY CamNombre"
   
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    vsGrid.Redraw = False
    Do While Not RsAux.EOF
        With vsGrid
            .AddItem Trim(RsAux("CamNombre"))
            .Cell(flexcpData, .Rows - 1, 0) = Trim(RsAux("ArtID"))
            .Cell(flexcpData, .Rows - 1, 1) = Trim(RsAux("CamCodigo"))
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux("ArtNombre"))
            .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux("Cantidad"))
            
            If RsAux("Tipo") = 2 Then
                .Cell(flexcpForeColor, .Rows - 1, 1) = &H608600
            End If
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    With vsGrid
        If .Rows > .FixedRows Then
            .MergeCells = flexMergeRestrictColumns
            '.Cell(flexcpBackColor, .FixedRows, 1, .Rows - 1) = &HEFEFEF
            .Cell(flexcpBackColor, .FixedRows, 2, .Rows - 1) = &HEFEFEF
            .MergeCol(0) = True
            .SubtotalPosition = flexSTBelow
            .Subtotal flexSTSum, 0, 2, "#,##0", &HB3DEF5, &HC0
        End If
        .Redraw = True
    End With
    Screen.MousePointer = 0
Exit Sub
errFM:
    vsGrid.Redraw = True
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al cargar la lista de estadísticas.", Err.Description
End Sub

Private Sub db_FillEstadistica()
On Error GoTo errFG
    
    Screen.MousePointer = 11
    vsGrid.BackColorAlternate = vsGrid.BackColor
    'Armo cuerpo que es común a la unión.
    Cons = "From Envio " & _
                    "Left Outer Join Camion On CamCodigo = EnvCamion " & _
                ", TipoFlete, Zona " & _
                " Where EnvFechaPrometida = '" & Format(tFecha.Value, "mm/dd/yyyy") & "' " & _
                " And (EnvCodigo NOT IN (SELECT EVCEnvio From EnvioVaCon) " & _
                    " Or EnvCodigo IN (SELECT Min(EVCEnvio) From EnvioVaCon, Envio E2 Where E2.EnvEstado = 0 And E2.EnvCodigo = EVCEnvio And E2.EnvFechaPrometida = '" & Format(tFecha.Value, "mm/dd/yyyy") & "'))" & _
                "And EnvEstado = 0 And TFlCodigo = EnvTipoFlete And EnvZona = ZonCodigo "
    
    Cons = "Select TFlNombreCorto, IsNull(CamNombre, 'A Asignar') as Cam, ZonNombre, Count(*) as Q, 'M' as Hora " & _
                Cons & " And (EnvRangoHora < '" & paHoraTarde & "' or EnvRangoHora Is Null) " & _
                "Group By TFlNombreCorto, CamNombre, ZonNombre " & _
                "Union All " & _
                "Select TFlNombreCorto, IsNull(CamNombre, 'A Asignar') as Cam, ZonNombre, Count(*) as Q, 'T' as Hora " & _
                Cons & " And EnvRangoHora > '" & paHoraTarde & "'" & _
                "Group By TFlNombreCorto, CamNombre, ZonNombre " & _
                "Order By TFlNombreCorto, Cam, ZonNombre, Hora"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    vsGrid.Redraw = False
    Do While Not RsAux.EOF
        With vsGrid
            .AddItem Trim(RsAux("TFlNombreCorto"))
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux("Cam"))
            If Trim(Trim(RsAux("Cam"))) = "A Asignar" Then .Cell(flexcpForeColor, .Rows - 1, 1) = &H900000
            .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux("Hora"))
            .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux("ZonNombre"))
            .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux("Q"))
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    With vsGrid
        If .Rows > .FixedRows Then
            .MergeCells = flexMergeRestrictColumns
            .Cell(flexcpBackColor, .FixedRows, 1, .Rows - 1) = &HEFEFEF
            .Cell(flexcpBackColor, .FixedRows, 3, .Rows - 1) = &HEFEFEF
            .MergeCol(0) = True
            .MergeCol(1) = True
            .MergeCol(3) = True
            .SubtotalPosition = flexSTBelow
            .Subtotal flexSTSum, 0, 4, "#,##0", &HB3DEF5, &HC0
        End If
        .Redraw = True
    End With
    Screen.MousePointer = 0
Exit Sub
errFG:
    vsGrid.Redraw = True
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al cargar la lista de estadísticas.", Err.Description
End Sub

Private Sub loc_AddArtNombre(ByVal iArt As Long, ByVal sNom As String)
On Error Resume Next
Dim iQ As Integer
    For iQ = 0 To UBound(arrNomArt)
        If arrNomArt(iQ).ID = iArt Then Exit For
    Next
    If arrNomArt(0).ID > 0 Then ReDim Preserve arrNomArt(UBound(arrNomArt) + 1)
    With arrNomArt(UBound(arrNomArt))
        .ID = iArt
        .Nombre = sNom
    End With
End Sub

Private Function fnc_GetNomArticulo(ByVal iArt As Long) As String
On Error Resume Next
Dim iQ As Integer
    For iQ = 0 To UBound(arrNomArt)
        If arrNomArt(iQ).ID = iArt Then fnc_GetNomArticulo = arrNomArt(iQ).Nombre: Exit For
    Next
End Function

Private Function fnc_GetStringArticulos(ByVal sArrArtsQ As String) As String
Dim arrQArt() As String
Dim arrAux() As String
Dim iQ As Integer
Dim sRet As String
    arrQArt = Split(sArrArtsQ, ";")
    For iQ = 0 To UBound(arrQArt)
        If InStr(1, arrQArt(iQ), ":") > 0 Then
            arrAux = Split(arrQArt(iQ), ":")
            arrAux(0) = fnc_GetNomArticulo(arrAux(0))
            If sRet <> "" Then sRet = sRet & ", "
            sRet = sRet & arrAux(1) & " " & arrAux(0)
        End If
    Next
    fnc_GetStringArticulos = sRet
End Function

Private Sub db_FillGridAAsignar()
On Error GoTo errFG
Dim sFM As String, lLastID As Long, lLastVC As Long
Dim totalBultos As Integer

    Screen.MousePointer = 11
    ReDim arrNomArt(0)
    
    '& ", (SELECT CQCIdQuePaga FROM Documento INNER JOIN ConQueCobra ON CQCIdQueCobra = DocCodigo AND CQCTipoQueCobra IN (1,2) And CQCTipoQuePaga = 15 WHERE CQCIdQueCobra = EnvDocumento AND DocTipo IN (1,2)) RedPagos" _
    'Saco todos los envíos y si tiene VaCon le agrego los artículos de los otros envíos.
'    Cons = "Select EnvCodigo, EnvFModificacion, IsNull(EnvRangoHora, '') as RH, EnvTipo, LocCodigo, LocNombre, " _
'                & " CalNombre, DirPuerta, DirLetra, DirApartamento, rTrim(TFlNombreCorto) as TF, " _
'                & " IsNull(EVCID, 0) as VC, IsNull(EnvComentario, '') as Memo, IsNull(rTrim(AgeNombre), rTrim(ZonNombre)) as ZN, REvCantidad, IsNull(AEsID, ArtCodigo) ArtCodigo, IsNull(AEsNombre, ArtDescripcion) as AD, ArtTipo, IsNull(AEsID, ArtID) ArtID " _
'                & ", CQCIdQuePaga RedPagos" _
'        & " From ((((((((((Envio Left Outer Join Agencia On EnvAgencia = AgeCodigo) LEFT OUTER JOIN EnvioVaCon ON EnvCodigo = EVCEnvio)" _
'        & " INNER JOIN Direccion ON EnvDireccion = DirCodigo)" _
'        & " INNER JOIN Calle ON DirCalle = CalCodigo) INNER JOIN Localidad ON CalLocalidad = LocCodigo) " _
'        & " INNER JOIN TipoFlete ON EnvTipoFlete = TFlCodigo)" _
'        & " INNER JOIN Zona ON EnvZona = ZonCodigo ) INNER JOIN RenglonEnvio ON EnvCodigo = REvEnvio) " _
'        & " INNER JOIN Articulo ON REvArticulo = ArtID)" _
'        & " LEFT OUTER JOIN ArticuloEspecifico ON AEsDocumento = EnvDocumento AND AEsArticulo = REvArticulo AND ((AEsTipoDocumento IN (1, 6) AND ENVTipo = 1) OR (AEsTipoDocumento IN (7, 33) AND ENVTipo = 3)))" _
'        & " LEFT OUTER JOIN ConQueCobra ON CQCIdQueCobra = EnvDocumento AND CQCTipoQueCobra IN (1,2) And CQCTipoQuePaga = 15 " _
'        & " Where EnvFechaPrometida = '" & Format(tFecha.Value, "mm/dd/yyyy") & "'" _
'        & " And EnvEstado = " & EstadoEnvio.AImprimir & " And EnvCamion Is Null " _
'        & " And EnvTipo In (1, 2, 3) And EnvDocumento > 0" _
'        & " Order by EVCID, EnvCodigo"
   
   Cons = "Select EnvCodigo, EnvFModificacion, IsNull(EnvRangoHora, '') as RH, EnvTipo, LocCodigo, LocNombre, " _
                & " CalNombre, DirPuerta, DirLetra, DirApartamento, rTrim(TFlNombreCorto) as TF, " _
                & " IsNull(EVCID, 0) as VC, IsNull(EnvComentario, '') as Memo, IsNull(rTrim(AgeNombre), rTrim(ZonNombre)+ ' (' + RTRIM(DepNombre) COLLATE Modern_Spanish_CI_AI  + ')') as ZN, REvCantidad, IsNull(AEsID, ArtCodigo) ArtCodigo, IsNull(AEsNombre, ArtNombre) as AD, ArtTipo, IsNull(AEsID, ArtID) ArtID " _
                & ", CQCIdQuePaga RedPagos, CZoZona, EnvZona, EnvAgencia, TFLTipoFlete, EnvHoraEspecial, DocTipo " _
        & " From ((((((((((Envio Left Outer Join Agencia On EnvAgencia = AgeCodigo) LEFT OUTER JOIN EnvioVaCon ON EnvCodigo = EVCEnvio)" _
        & " INNER JOIN Direccion ON EnvDireccion = DirCodigo)" _
        & " INNER JOIN Calle ON DirCalle = CalCodigo) INNER JOIN Localidad ON CalLocalidad = LocCodigo INNER JOIN Departamento ON DepCodigo = LocDepartamento) " _
        & " INNER JOIN TipoFlete ON EnvTipoFlete = TFlCodigo)" _
        & " INNER JOIN Zona ON EnvZona = ZonCodigo ) INNER JOIN RenglonEnvio ON EnvCodigo = REvEnvio) " _
        & " INNER JOIN Articulo ON REvArticulo = ArtID) " _
        & " LEFT OUTER JOIN ArticuloEspecifico ON AEsDocumento = EnvDocumento AND AEsArticulo = REvArticulo AND ((AEsTipoDocumento IN (1, 6) AND ENVTipo = 1) OR (AEsTipoDocumento IN (7, 33) AND ENVTipo = 3)))" _
        & " LEFT OUTER JOIN ConQueCobra ON CQCIdQueCobra = EnvDocumento AND CQCTipoQueCobra IN (1,2) And CQCTipoQuePaga = 15 LEFT OUTER JOIN CalleZona ON CalCodigo = CZoCalle AND CZoDesde <= DirPuerta AND CZoHasta >= DirPuerta " _
        & " LEFT OUTER JOIN Documento ON DocCodigo = EnvDocumento AND EnvTipo = 1" _
        & " Where EnvFechaPrometida = '" & Format(tFecha.Value, "mm/dd/yyyy") & "'" _
        & " And EnvEstado = " & EstadoEnvio.AImprimir & " And EnvCamion Is Null " _
        & " And EnvTipo In (1, 2, 3) And EnvDocumento > 0" _
        & " Order by EVCID, EnvCodigo"
   
    Dim bZonaMal As Boolean
    bZonaMal = False
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    vsGrid.Redraw = False
    Do While Not RsAux.EOF
        With vsGrid
            If lLastID <> RsAux("EnvCodigo") Then
                lLastID = RsAux("EnvCodigo")
                If lLastVC <> RsAux("VC") Or RsAux("VC") = 0 Then
                    lLastVC = RsAux("VC")
                    .AddItem RsAux!EnvCodigo
                    .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked
                    .Cell(flexcpText, .Rows - 1, 1) = RsAux!RH
                    If Not IsNull(RsAux("EnvHoraEspecial")) Then .Cell(flexcpFontBold, .Rows - 1, 1) = True
                    .Cell(flexcpText, .Rows - 1, 2) = IIf(RsAux!RH > paHoraTarde, "T ", IIf(RsAux!RH <> "", "M ", ""))
                    .Cell(flexcpText, .Rows - 1, 3) = RsAux!TF
                    If (RsAux("TFLTipoFlete") = eTiposDeTipoFlete.CostoEspecial) Then
                        .Cell(flexcpFontBold, .Rows - 1, 3) = True
                        .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = vbRed
                    End If
                    
                    .Cell(flexcpText, .Rows - 1, 4) = ""
                    .Cell(flexcpText, .Rows - 1, 5) = RsAux!ZN
                    .Cell(flexcpText, .Rows - 1, 6) = f_GetDireccionRsAux
                    .Cell(flexcpText, .Rows - 1, 7) = Trim(RsAux!Memo)
                    'DATA
                    sFM = RsAux!EnvFModificacion: .Cell(flexcpData, .Rows - 1, 1) = sFM         'F Modificado
                    sFM = RsAux!VC: .Cell(flexcpData, .Rows - 1, 2) = Val(sFM)                      'Va Con
                    
                    If RsAux("EnvTipo") = 3 Then
                        .Cell(flexcpForeColor, .Rows - 1, 0) = &HC0
                    ElseIf RsAux("EnvTipo") = 2 Then
                        .Cell(flexcpForeColor, .Rows - 1, 0) = &H8000&
                    End If
                    
                    If Not IsNull(RsAux("RedPagos")) Then .Cell(flexcpForeColor, .Rows - 1, 5) = &H80FF& ': .Cell(flexcpFontBold, .Rows - 1, 5) = True
                    
                    If IsNull(RsAux("EnvAgencia")) Then
                        If RsAux("EnvZona") <> RsAux("CZoZona") Then
                            .Cell(flexcpForeColor, .Rows - 1, 5) = vbWhite
                            .Cell(flexcpBackColor, .Rows - 1, 5) = &HC0
                            bZonaMal = True
                        End If
                    End If
                    
                    If Not IsNull(RsAux("DocTipo")) Then
                        If RsAux("DocTipo") = TD_Contado Then
                            .Cell(flexcpText, .Rows - 1, 9) = "Contado"
                        ElseIf RsAux("DocTipo") = TD_Credito Then
                            .Cell(flexcpText, .Rows - 1, 9) = "Crédito"
                        ElseIf RsAux("DocTipo") = 47 Then
                            .Cell(flexcpText, .Rows - 1, 9) = "Cambio"
                        ElseIf RsAux("DocTipo") = 48 Then
                            .Cell(flexcpText, .Rows - 1, 9) = "Retiro"
                        End If
                    ElseIf RsAux("EnvTipo") = 3 Then
                        .Cell(flexcpText, .Rows - 1, 9) = "Vta.Telef."
                    End If
                    
                End If
            End If
            '..........................ARTICULOS
            loc_AgregoEnColleccionBultos RsAux("ArtID"), Format(RsAux("ArtCodigo"), "(#,000,000)") & " " & RsAux("AD"), RsAux("REvCantidad")
            loc_AddArtNombre RsAux("ArtID"), Format(RsAux("ArtCodigo"), "(#,000,000)") & " " & RsAux("AD")
            loc_SetQTipoArt .Rows - 1, RsAux("ArtID"), RsAux("ArtTipo"), RsAux("REvCantidad")
            totalBultos = totalBultos + RsAux("REvCantidad")
            .Cell(flexcpText, .Rows - 1, 8) = fnc_GetStringArticulos(.Cell(flexcpData, .Rows - 1, 6))
            '..........................Artículos
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    vsGrid.Redraw = True
    If bZonaMal Then
        MsgBox "Existen envíos con la zona incorrecta.", vbInformation, "ATENCIÓN"
    End If
    sbStatus.Panels("bultos").Text = "Bultos: " & totalBultos
    loc_QEnviosClic
    loc_SetStatusTipoArt
    Screen.MousePointer = 0

Exit Sub
errFG:
    vsGrid.Redraw = True
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al cargar la lista a confirmar", Err.Description
End Sub

Private Sub CantidadPosiblesDuplicados()
On Error Resume Next
    Cons = "SELECT IsNull(Count(DISTINCT(E1.EnvCodigo)), 0)" & _
        " FROM Envio E1 INNER JOIN RenglonEnvio RE1 ON RE1.REvEnvio = E1.EnvCodigo" & _
        " INNER JOIN Envio E2 ON E2.EnvCliente = E1.EnvCliente AND E1.EnvCodigo <> E2.EnvCodigo AND E1.EnvDocumento <> E2.EnvDocumento AND E2.EnvEstado IN (0, 1, 3) AND E2.EnvCodigo NOT IN(SELECT EVCEnvio FROM EnvioVaCon)" & _
        " INNER JOIN RenglonEnvio RE2 ON RE2.REvEnvio = E2.EnvCodigo AND RE1.REvArticulo = RE2.REvArticulo" & _
        " INNER JOIN Direccion Dir1 ON E1.EnvDireccion = Dir1.DirCodigo" & _
        " INNER JOIN Direccion Dir2 ON E2.EnvDireccion = Dir2.DirCodigo AND Dir2.DirCalle = Dir1.DirCalle AND Dir1.DirPuerta = Dir2.DirPuerta" & _
        " WHERE E1.EnvEstado in (0, 1, 3) AND e1.EnvCodigo NOT IN(SELECT EVCEnvio FROM EnvioVaCon)"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    hlTab(7).Caption = IIf(RsAux(0) > 0, RsAux(0), "") & " Duplicados"
    RsAux.Close
End Sub

Private Sub db_FillPosiblesDuplicados()
Dim sEnvs As String
    
    sEnvs = ""
    On Error GoTo errPD
    Cons = "Select IsNull(EnvCodImpresion, 0) EnvCodImpresion, EnvCodigo, EnvFModificacion, EnvTipo, IsNull(EVCID, 0) as VC, " _
                & " EnvCamion, LocCodigo, LocNombre, CamNombre, CalNombre, DirPuerta, DirLetra, DirApartamento, " _
                & " IsNull(EnvComentario, '') as Memo, REvCantidad, IsNull(AEsID, ArtCodigo) as ArtCodigo, IsNull(AEsNombre, ArtNombre) as AD, ArtTipo, IsNull(AEsID, ArtID) ArtID " _
        & " From ((((((((Envio LEFT OUTER JOIN EnvioVaCon ON EnvCodigo = EVCEnvio) LEFT OUTER JOIN Direccion ON EnvDireccion = DirCodigo) " _
        & " LEFT OUTER JOIN Calle ON DirCalle = CalCodigo) INNER JOIN Localidad ON CalLocalidad = LocCodigo) INNER JOIN RenglonEnvio ON EnvCodigo = REvEnvio) " _
        & " INNER JOIN Articulo ON REvArticulo = ArtID) LEFT OUTER JOIN Camion ON EnvCamion = CamCodigo)" _
        & " LEFT OUTER JOIN ArticuloEspecifico ON AEsTipoDocumento IN (1, 6) AND AEsDocumento = EnvDocumento AND AEsArticulo = REvArticulo)" _
        & " WHERE EnvCodigo IN ("
        
    Cons = Cons & "SELECT DISTINCT(E1.EnvCodigo)" & _
        " FROM Envio E1 INNER JOIN RenglonEnvio RE1 ON RE1.REvEnvio = E1.EnvCodigo" & _
        " INNER JOIN Envio E2 ON E2.EnvCliente = E1.EnvCliente AND E1.EnvCodigo <> E2.EnvCodigo AND E1.EnvDocumento <> E2.EnvDocumento AND E2.EnvEstado IN (0, 1, 3) AND E2.EnvCodigo NOT IN(SELECT EVCEnvio FROM EnvioVaCon)" & _
        " INNER JOIN RenglonEnvio RE2 ON RE2.REvEnvio = E2.EnvCodigo AND RE1.REvArticulo = RE2.REvArticulo" & _
        " INNER JOIN Direccion Dir1 ON E1.EnvDireccion = Dir1.DirCodigo" & _
        " INNER JOIN Direccion Dir2 ON E2.EnvDireccion = Dir2.DirCodigo AND Dir2.DirCalle = Dir1.DirCalle AND Dir1.DirPuerta = Dir2.DirPuerta" & _
        " WHERE E1.EnvEstado in (0, 1, 3) AND e1.EnvCodigo NOT IN(SELECT EVCEnvio FROM EnvioVaCon)"
    
    Cons = Cons & ") ORDER BY EVCID, EnvCodigo"
    
    Dim lLastID As Long, lLastVC As Long
    Dim sFM As String
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        With vsGrid
            If lLastID <> RsAux("EnvCodigo") Then
                lLastID = RsAux("EnvCodigo")
                If lLastVC <> RsAux("VC") Or RsAux("VC") = 0 Then
                    lLastVC = RsAux("VC")
                    .AddItem RsAux!EnvCodigo
                    .Cell(flexcpText, .Rows - 1, 5) = f_GetDireccionRsAux
                    .Cell(flexcpText, .Rows - 1, 6) = Trim(RsAux!Memo)

                    'DATA
                    sFM = RsAux!EnvFModificacion: .Cell(flexcpData, .Rows - 1, 1) = sFM         'F Modificado
                    sFM = RsAux!VC: .Cell(flexcpData, .Rows - 1, 2) = Val(sFM)                      'Va Con

                    If RsAux("EnvTipo") = 2 Then .Cell(flexcpForeColor, .Rows - 1, 0) = &H8000&
                End If
            End If
            ''..........................ARTICULOS
            loc_AddArtNombre RsAux("ArtID"), Format(RsAux("ArtCodigo"), "(#,000,000)") & " " & RsAux("AD")
            loc_SetQTipoArt .Rows - 1, RsAux("ArtID"), RsAux("ArtTipo"), RsAux("REvCantidad")
            .Cell(flexcpText, .Rows - 1, 7) = fnc_GetStringArticulos(.Cell(flexcpData, .Rows - 1, 6))
            '..........................Artículos
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    Exit Sub
    
errPD:
End Sub

Private Function f_GetArticulosEnvioConNombre(ByVal lEnvio As Long, ByVal iVaCon As Long) As String
Dim rsFG As rdoResultset
    f_GetArticulosEnvioConNombre = ""
'    Cons = "Select Sum(RevAEntregar), ArtCodigo, rTrim(ArtNombre)" & _
'            " From Envio, RenglonEnvio, Articulo "
'    If iVaCon > 0 Then
'        Cons = Cons & " Where EnvCodigo IN (SELECT EVCEnvio From EnvioVaCon Where EVCID = " & iVaCon & ")"
'    Else
'        Cons = Cons & " Where EnvCodigo = " & lEnvio
'    End If
'    Cons = Cons & " And EnvCodigo = REvEnvio And REvArticulo = ArtID Group By ArtCodigo, ArtNombre"
    
    Cons = "SELECT Sum(RevAEntregar), AEsID, ArtCodigo, rTrim(ISNULL(AESNombre, ArtNombre)) Nombre " & _
        "FROM Envio INNER JOIN RenglonEnvio ON EnvCodigo = REvEnvio " & _
        "INNER JOIN Articulo ON ArtId = REvArticulo " & _
        "LEFT OUTER JOIN ArticuloEspecifico On AEsTipoDocumento IN(1, 6) " & _
        "AND AEsDocumento = EnvDocumento And AEsArticulo = RevArticulo"
    
    If iVaCon > 0 Then
        Cons = Cons & " Where EnvCodigo IN (SELECT EVCEnvio From EnvioVaCon Where EVCID = " & iVaCon & ")"
    Else
        Cons = Cons & " Where EnvCodigo = " & lEnvio
    End If
    Cons = Cons & " GROUP BY ArtCodigo, AESNombre, AEsID, ArtNombre"
    
    
    Set rsFG = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsFG.EOF
        If f_GetArticulosEnvioConNombre <> "" Then f_GetArticulosEnvioConNombre = f_GetArticulosEnvioConNombre & "|"
        Dim sIDEsp As String
        If Not IsNull(rsFG(1)) Then sIDEsp = " AEsp:" & rsFG(1) Else sIDEsp = ""
        f_GetArticulosEnvioConNombre = f_GetArticulosEnvioConNombre & rsFG(0) & sIDEsp & " " & Format(rsFG(2), "(#,000,000)") & " " & rsFG(3)
        rsFG.MoveNext
    Loop
    rsFG.Close
    Screen.MousePointer = 0
End Function

Private Sub db_FillGridAConfirmar()
Dim sFM As String
Dim lLastID As Long, lLastVC As Long
Dim totalBultos As Integer
On Error GoTo errFG

    Screen.MousePointer = 11
    ReDim arrNomArt(0)
    Cons = "Select EnvCodigo, EnvFModificacion, EnvFechaPrometida as FP, IsNull(EnvRangoHora, '') as RH, EnvTipo, LocCodigo, LocNombre, " _
                    & " CalNombre, DirPuerta, DirLetra, DirApartamento, rTrim(TFlNombreCorto) as TF, " _
                    & " IsNull(EVCID, 0) as VC, IsNull(EnvComentario, '') as Memo, IsNull(rTrim(AgeNombre), rTrim(ZonNombre)) as ZN, REvCantidad, IsNull(AEsID, ArtCodigo) ArtCodigo, IsNull(AEsNombre, ArtNombre) as AD, ArtTipo, IsNull(AEsID, ArtID) ArtID, TFLTipoFlete " _
        & " From ((((((((((Envio Left Outer Join Agencia On EnvAgencia = AgeCodigo) LEFT OUTER JOIN EnvioVaCon ON EnvCodigo = EVCEnvio)" _
        & " INNER JOIN Direccion ON EnvDireccion = DirCodigo)" _
        & " INNER JOIN Calle ON DirCalle = CalCodigo) INNER JOIN Localidad ON CalLocalidad = LocCodigo) " _
        & " INNER JOIN TipoFlete ON EnvTipoFlete = TFlCodigo)" _
        & " INNER JOIN Zona ON EnvZona = ZonCodigo ) INNER JOIN RenglonEnvio ON EnvCodigo = REvEnvio) " _
        & " INNER JOIN Articulo ON REvArticulo = ArtID)" _
        & " LEFT OUTER JOIN ArticuloEspecifico ON AEsDocumento = EnvDocumento AND AEsArticulo = REvArticulo AND ((AEsTipoDocumento IN (1, 6) AND ENVTipo = 1) OR (AEsTipoDocumento IN (7, 33) AND ENVTipo = 3)))" _
        & " Where EnvEstado = 1 And EnvDocumento > 0 And EnvTipo In (1, 2, 3) " _
        & " Union All " _
        & "Select EnvCodigo, EnvFModificacion, EnvFechaPrometida as FP, IsNull(EnvRangoHora, '') as RH, EnvTipo, 0 LocCodigo, Null as LocNombre, " _
                    & " Null as CalNombre, Null as DirPuerta, Null as DirLetra, Null as DirApartamento, rTrim(TFlNombreCorto) as TF, " _
                    & " IsNull(EVCID, 0) as VC, IsNull(EnvComentario, '') as Memo, Null as ZN, REvCantidad, IsNull(AEsID, ArtCodigo) ArtCodigo, IsNull(AEsNombre, ArtNombre) as AD, ArtTipo, IsNull(AEsID, ArtID) ArtID, TFLTipoFlete " _
        & " From (((((Envio LEFT OUTER JOIN EnvioVaCon On EnvCodigo = EVCEnvio) INNER JOIN TipoFlete ON EnvTipoFlete = TFlCodigo)" _
        & " INNER JOIN RenglonEnvio ON EnvCodigo = REvEnvio) " _
        & " INNER JOIN Articulo ON REvArticulo = ArtID)" _
        & " LEFT OUTER JOIN ArticuloEspecifico ON AEsDocumento = EnvDocumento AND AEsArticulo = REvArticulo AND ((AEsTipoDocumento IN (1, 6) AND ENVTipo = 1) OR (AEsTipoDocumento IN (7, 33) AND ENVTipo = 3)))" _
        & " Where EnvDireccion Is Null And EnvEstado = 1 And EnvDocumento > 0 " _
        & " And EnvTipo In (1, 2, 3) " _
        & " Order by VC, EnvCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    vsGrid.Redraw = False
    '"Envío|Fecha|Horario|Tipo Flete|Zona|Dirección|Comentario|Artículos"
    Do While Not RsAux.EOF
        With vsGrid
            If lLastID <> RsAux!EnvCodigo Then
                lLastID = RsAux("EnvCodigo")
                If lLastVC <> RsAux("VC") Or RsAux("VC") = 0 Then
                    lLastVC = RsAux("VC")
                    .AddItem RsAux!EnvCodigo
                    If Not IsNull(RsAux!FP) Then .Cell(flexcpText, .Rows - 1, 1) = RsAux!FP
                    .Cell(flexcpText, .Rows - 1, 2) = RsAux!RH
                    .Cell(flexcpText, .Rows - 1, 3) = RsAux!TF
                    If Not IsNull(RsAux("ZN")) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux("ZN"))
                    .Cell(flexcpText, .Rows - 1, 5) = f_GetDireccionRsAux
                    .Cell(flexcpText, .Rows - 1, 6) = Trim(RsAux("Memo"))
                    'Valores en el DATA
                    sFM = RsAux!EnvFModificacion
                    .Cell(flexcpData, .Rows - 1, 1) = sFM
                    
                    sFM = RsAux!VC
                    .Cell(flexcpData, .Rows - 1, 2) = Val(sFM)
                    
                    If (RsAux("TFLTipoFlete") = eTiposDeTipoFlete.CostoEspecial) Then
                        .Cell(flexcpFontBold, .Rows - 1, 3) = True
                        .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = vbRed
                    End If
                    
                End If
            End If
            
            '..........................ARTICULOS
            loc_AgregoEnColleccionBultos RsAux("ArtID"), Format(RsAux("ArtCodigo"), "(#,000,000)") & " " & RsAux("AD"), RsAux("REvCantidad")
            loc_AddArtNombre RsAux("ArtID"), Format(RsAux("ArtCodigo"), "(#,000,000)") & " " & RsAux("AD")
            loc_SetQTipoArt .Rows - 1, RsAux("ArtID"), RsAux("ArtTipo"), RsAux("REvCantidad")
            totalBultos = totalBultos + RsAux("REvCantidad")
            .Cell(flexcpText, .Rows - 1, 7) = fnc_GetStringArticulos(.Cell(flexcpData, .Rows - 1, 6))
            '..........................Artículos
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    vsGrid.Redraw = True
    sbStatus.Panels("envio").Text = "Envíos: " & vsGrid.Rows - 1
    sbStatus.Panels("bultos").Text = "Bultos: " & totalBultos
    Screen.MousePointer = 0
    
    Exit Sub
errFG:
    vsGrid.Redraw = True
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al cargar la lista a confirmar", Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad

    ObtengoSeteoForm Me, 1000, 500
    oLog.FileLog = App.Path & "\logDistEnv.txt"
    oLog.InsertoLog TL_Info, "Inicio la aplicación"
    loc_DatosRegistry
    With vsGrid
        .Editable = False: .Rows = 1: .Cols = 1 ': .ExtendLastCol = True
        .FormatString = "Envío|Horario|Dirección"
    End With
        
    tFecha.Value = IIf(Weekday(Date) = 7, Date + 2, Date + 1)
    
'    oCnfgPrint.CargarConfiguracion "FacturaContado", "CuotasImpresora"

    MnuOtrSendWhatsapp.Checked = (paSendWApp = 1)

    hlTab(0).Tag = "0"
    s_SetFontTab 0
    s_FillGrid
    
    Dim fHeader As New StdFont, ffooter As New StdFont
    With fHeader
        .Bold = True
        .Name = "Arial"
        .Size = 11
    End With
    With ffooter
        .Bold = True
        .Name = "Tahoma"
        .Size = 10
    End With
        
    db_FillArrayCamiones
    
    With oMenu
        .MenuBackColor = vbWindowBackground
        .MenuBorderColor = &HC0C0C0
        .Init
    End With
    
    With vsPrint
        .MarginBottom = 350
        .MarginLeft = 50
        .MarginRight = 150
        .MarginTop = 350
        .PageBorder = pbNone
    End With
    sbStatus.Panels("tarde").Text = "Tarde a partir: " & paHoraTarde & " "
    
    
    If iIDArt > 0 And sCodArt = "" Then
        Dim rsA As rdoResultset
        Set rsA = cBase.OpenResultset("Select ArtID, ArtDescripcion from articulo where ArtID = " & iIDArt, rdOpenDynamic, rdConcurValues)
        If Not rsA.EOF Then
            iIDArt = rsA(0)
            sCodArt = Trim(rsA(1))
            SaveSetting App.Title, "Settings", "AA" & Me.Name & "ArtDesc", sCodArt
            SaveSetting App.Title, "Settings", "AA" & Me.Name & "IDArt", iIDArt
        Else
            MsgBox "No hay un artículo con ese código.", vbExclamation, "Atención"
        End If
        rsA.Close
        sbStatus.Panels("articulo").Text = IIf(sCodArt = "", "art: " & sCodArt, "Sin artículo asignado")
    End If
    
    printBandejaCopiaeTicket = LeoSeteoBandejaCopiaeTicket()
    If printBandejaCopiaeTicket = 0 Then
        MsgBox "IMPORTANTE!!!" & vbCrLf & vbCrLf & "No tiene configurada la bandeja para las copias de eticket." & vbCrLf & "Acceda al menú impresión para asignarla.", vbExclamation, "IMPORTANTE"
    End If
    
    
Exit Sub
ErrLoad:
    objGral.OcurrioError "Error al inicializar el formulario.", Err.Description
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    picCodImpresion.Move 0, picLink.Top + picLink.Height, ScaleWidth
    If Val(hlTab(0).Tag) = 5 Then
        vsGrid.Top = picCodImpresion.Top + picCodImpresion.Height
    Else
        vsGrid.Top = picCodImpresion.Top
    End If
    vsGrid.Move 0, vsGrid.Top, ScaleWidth, ScaleHeight - (vsGrid.Top + sbStatus.Height)  '(picLink.Top + picLink.Height + sbStatus.Height)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    GuardoSeteoForm Me
    Set objGral = Nothing
    CierroConexion
'    crCierroEngine
End Sub

Private Sub hlTab_Click(Index As Integer)
    s_SetFontTab Index
    s_FillGrid
End Sub

Private Sub imgMercaderia_Click()
    loc_AppReclamar Val(lCamion.Tag)
End Sub

Private Sub Label1_Click()
On Error Resume Next
    tFecha.SetFocus
End Sub

Private Sub lbMercaderia_Click()
    loc_AppReclamar Val(lCamion.Tag)
End Sub

Private Sub MnuAIECamion_Click(Index As Integer)
    If Val(hlTab(0).Tag) = 1 Then
        ' A IMPRIMIR
        db_FillGridAImprimirImpreso MnuAIECamion(Index).Tag, AImprimir
        iCamSelect = MnuAIECamion(Index).Tag
        lTitle.Caption = "Envíos asignados a " & Mid(MnuAIECamion(Index).Caption, 1, InStrRev(MnuAIECamion(Index).Caption, "(") - 1) & " para ser impresos el día " & tFecha.Text
    Else
        ' IMPRESOS
        db_FillGridAImprimirImpreso MnuAIECamion(Index).Tag, Impreso
        iCamSelect = MnuAIECamion(Index).Tag
        lTitle.Caption = "Envíos que tiene " & MnuAIECamion(Index).Caption & " para ser entregados el día " & tFecha.Text
    End If
    On Error Resume Next
    If vsGrid.Rows > vsGrid.FixedRows Then vsGrid.SetFocus
End Sub

Private Sub MnuCambiarAnular_Click()
    If (Val(hlTab(0).Tag) = 2 Or Val(hlTab(0).Tag) = 5) And vsGrid.Rows > 1 Then
        tmModal.Tag = "4"
        tmModal.Enabled = True
    End If
End Sub

Private Sub MnuCambiarCamion_Click()
On Error Resume Next
    If (Val(hlTab(0).Tag) = 2 Or Val(hlTab(0).Tag) = 5) And vsGrid.Rows > 1 Then
        tmModal.Tag = "1"
        tmModal.Enabled = True
    End If
End Sub

Private Sub MnuCambiarDividir_Click()
    If Val(hlTab(0).Tag) = 5 And vsGrid.Rows > 1 Then
        tmModal.Tag = "3"
        tmModal.Enabled = True
    End If
End Sub

Private Sub MnuCambiarHora_Click()
On Error Resume Next
    If (Val(hlTab(0).Tag) = 2 Or Val(hlTab(0).Tag) = 5) And vsGrid.Rows > 1 Then
        tmModal.Tag = "2"
        tmModal.Enabled = True
    End If
End Sub

Private Sub MnuEliminoVaCon_Click()
    loc_EliminoVaCon
End Sub

Private Sub MnuGridAAsignar_Click()
Dim sF As String
    If vsGrid.Rows = vsGrid.FixedRows Then Exit Sub
    If Val(hlTab(0).Tag) = 3 Then
        'El Estado = a Confirmar
        If vsGrid.Cell(flexcpText, vsGrid.Row, 4) <> "" Then
        
            If vsGrid.Cell(flexcpFontBold, vsGrid.Row, 3) = True Then
                MsgBox "El envío es de flete especial, para cambiar la fecha y/o el estado debe acceder al formulario de envío.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        
            sF = InputBox("Ingrese la fecha del envío", "Fecha a Enviar", vsGrid.Cell(flexcpText, vsGrid.Row, 1))
            If IsDate(sF) Then
                db_EnvioAAsignar sF
            ElseIf sF <> "" Then
                MsgBox "Fecha incorrecta.", vbExclamation, "Atención"
            End If
        Else
            MsgBox "No se puede cambiar de estado el envío no tiene zona.", vbExclamation, "Atención"
        End If
        
    ElseIf Val(hlTab(0).Tag) = 1 Then
        db_EnvioAAsignar
    End If
End Sub


Private Sub MnuGridAConfirmar_Click()
    db_EnvioAConfirmar
End Sub

Private Sub MnuGridAsiTodosCamion_Click(Index As Integer)
'Recorro la lista y asigno todos al camión seleccionado.
    If Val(MnuGridAsiTodosCamion(Index).Tag) > 0 Then
        If MsgBox("¿Confirma asignar los envíos seleccionados al camión '" & MnuGridAsiTodosCamion(Index).Caption & "'", vbQuestion + vbYesNo, "Asignar a camión") = vbNo Then Exit Sub
        Dim iQ As Integer
        iQ = 1
        Do While iQ <= vsGrid.Rows - 1
            If vsGrid.Cell(flexcpChecked, iQ, 0) = flexChecked Then
                If db_EnvioAsignoCamion(MnuGridAsiTodosCamion(Index).Tag, iQ) Then
                    vsGrid.RemoveItem iQ
                    iQ = iQ - 1
                End If
            End If
            iQ = iQ + 1
        Loop
        s_FillGrid
    End If
End Sub

Private Sub MnuGridCamion_Click(Index As Integer)
    If Val(MnuGridCamion(Index).Tag) > 0 Then
        
        If vsGrid.SelectedRows > 0 Then
            If MsgBox("¿Confirma asignar los envíos seleccionados al camión '" & MnuGridCamion(Index).Caption & "'", vbQuestion + vbYesNo, "Asignar a camión") = vbNo Then Exit Sub
        Else
            If MsgBox("¿Confirma asignar el envío " & vbCr & "(" & vsGrid.Cell(flexcpText, vsGrid.Row, 0) & ") " & _
                    vsGrid.Cell(flexcpText, vsGrid.Row, 3) & vbCr & " al camión '" & MnuGridCamion(Index).Caption & "'", vbQuestion + vbYesNo, "Asignar a camión") = vbNo Then Exit Sub
        End If
        On Error GoTo errEAC
        
        'En algunas grillas tengo
        If Val(hlTab(0).Tag) = 1 Or Val(hlTab(0).Tag) = 3 Then
            Dim iQ As Integer
            iQ = 1
            Do While (vsGrid.SelectedRows > 0)
                If Not (Val(hlTab(0).Tag) = 3 And vsGrid.Cell(flexcpText, vsGrid.SelectedRow(0), 4) = "") Then
                    If db_EnvioAsignoCamion(Val(MnuGridCamion(Index).Tag), vsGrid.SelectedRow(0)) Then
                        vsGrid.RemoveItem vsGrid.SelectedRow(0)
                    End If
                Else
                    MsgBox "Atención hay envíos que no tienen asignada una zona, refresque la información y luego edite el envío.", vbExclamation, "Atención"
                    vsGrid.RemoveItem vsGrid.SelectedRow(0)
                End If
            Loop
            s_SetMenu
        Else
            If db_EnvioAsignoCamion(Val(MnuGridCamion(Index).Tag), vsGrid.Row) Then
                vsGrid.RemoveItem vsGrid.Row
                s_SetMenu
            Else
                s_FillGrid
            End If
        End If
    End If
    Exit Sub
errEAC:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al asignarle camión al envío.", Err.Description
End Sub

Private Sub MnuGridEditEnvio_Click()
On Error GoTo errEnv
    If vsGrid.Rows = vsGrid.FixedRows Then Exit Sub
    Screen.MousePointer = 11
    Dim objEnv As New clsEnvio
    objEnv.InvocoEnvio vsGrid.Cell(flexcpValue, vsGrid.Row, 0), "C:\Programas O&R\Reportes"
    Set objEnv = Nothing
    Screen.MousePointer = 0
    Me.Refresh
    'Si la opción es la uno tengo que mantener la lista ya que puede estar con datos ya editados.
    If Val(hlTab(0).Tag) < 2 Then
        s_VerficoEnvioEditado iCamSelect, AImprimir
    ElseIf Val(hlTab(0).Tag) = 2 Then
        s_VerficoEnvioEditado iCamSelect, Impreso
    ElseIf Val(hlTab(0).Tag) <> 5 Then
        s_FillGrid
    End If
    Exit Sub
errEnv:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al editar el envío.", Err.Description
End Sub

Private Sub MnuGridFiltroExcSel_Click()
    s_GridFiltrar False
End Sub

Private Sub MnuGridFiltroSel_Click()
    s_GridFiltrar True
End Sub

Private Sub MnuGridSelectAll_Click()
    s_GridSelectAll True
End Sub

Private Sub MnuGridUnSelectAll_Click()
    s_GridSelectAll False
End Sub

Private Sub MnuGSCResto_Click()
    s_SetCamionSugerido False
End Sub

Private Sub MnuGSCSelect_Click()
    s_SetCamionSugerido True
End Sub

Private Sub MnuMerDar_Click()
On Error Resume Next
    Shell App.Path & "\EntDev_Arts_Reparto.exe 1", vbNormalFocus
End Sub

Private Sub MnuMerDevolucion_Click()
On Error Resume Next
    Shell App.Path & "\EntDev_Arts_Reparto.exe 2", vbNormalFocus
End Sub

Private Sub MnuMerReclamar_Click()
On Error Resume Next
    loc_AppReclamar
End Sub

Private Sub MnuMerStockTotal_Click()
    RunApp App.Path & "\stock total.exe"
End Sub

Private Sub MnuOtrAgendaFlete_Click()
    RunApp App.Path & "\Cerrar horario.exe"
End Sub

Private Sub MnuOtrCambioEstadoEnvio_Click()
On Error Resume Next
    Dim oFrm As New CamEstadoEnvio
    oFrm.Show vbModal
    Set oFrm = Nothing
End Sub

Private Sub MnuOtrDeshabilitoCamion_Click()
    HabilitarCamion False
End Sub

Private Sub MnuOtrEnviarWhatsApp_Click()
    Dim sID As String
    sID = InputBox("Ingrese el código de impresión")
    If IsNumeric(sID) Then
        Dim celular As String
        celular = InputBox("Ingrese a que celular desea envíar, vacio es el telf. del cliente.")
        whatsappEnviarMensajeCodigo sID, celular
    End If
End Sub

Private Sub MnuOtrEnvio_Click()
    RunApp App.Path & "\envios.exe"
End Sub

Private Sub MnuOtrFichaAgencia_Click()
    RunApp App.Path & "\Ficha de Agencia.exe"
End Sub

Private Sub MnuOtrHabilitarCamion_Click()
    HabilitarCamion True
End Sub

Private Sub MnuOtrSendWhatsapp_Click()
    MnuOtrSendWhatsapp.Checked = Not MnuOtrSendWhatsapp.Checked
End Sub

Private Sub MnuOtrTiposFletes_Click()
    RunApp App.Path & "\Tipos de flete.exe"
End Sub

Private Sub MnuPrintAux_Click()
On Error GoTo errMP
'Tengo que recorrer la lista y sacar los código de envíos y luego consultar todo para que ordene por zona y agencia.
    Dim sEnvios As String
    Dim iQ As Integer
    With vsGrid
        For iQ = .FixedRows To .Rows - 1
            If sEnvios <> "" Then sEnvios = sEnvios & ", "
            sEnvios = sEnvios & .Cell(flexcpValue, iQ, 0)
        Next
    End With
    
    Cons = "Select EnvCodigo, EnvDireccion, IsNull(EnvRangoHora, '') as RH, rTrim(TFlNombreCorto) as TF, " _
                    & " IsNull(EnvComentario, '') as Memo, rTrim(isNull(AgeNombre, ZonNombre)) as ZAN " _
        & " FROM (Envio Left Outer Join Agencia On EnvAgencia = AgeCodigo)" _
        & ", TipoFlete, Zona" _
        & " Where EnvCodigo In (" & sEnvios & ")" _
        & " And EnvZona = ZonCodigo And EnvTipoFlete = TFlCodigo "
                
        If Val(hlTab(0).Tag) = 0 Then
            Cons = Cons & " And EnvCamion Is Null "
        Else
            Cons = Cons & " And EnvCamion = " & iCamSelect
        End If
    Cons = Cons & " Order by ZAN, EnvCodigo"
    loc_ImprimoPlanilla Cons
    Exit Sub
errMP:
    objGral.OcurrioError "Error al armar la consulta para imprimir la planilla auxiliar.", Err.Description, "Error (planilla auxiliar)"
End Sub

Private Sub MnuPrintBultos_Click()
On Error GoTo errPB
Dim oArt As clsArticuloMenu
    loc_StartCtrlPrint False

    Dim sCamion As String
    Dim iQ As Integer
    For iQ = MnuAIECamion.LBound To MnuAIECamion.UBound
        If MnuAIECamion(iQ).Tag = iCamSelect Then
            sCamion = Mid(MnuAIECamion(iQ).Caption, 1, InStrRev(MnuAIECamion(iQ).Caption, "(") - 1)
        End If
    Next

    With vsPrint
        .HdrFontName = "Tahoma"
        .HdrFontBold = True
        .HdrFontSize = "9"
        .Header = "Plantilla de bultos camión " & sCamion & "||Fecha: " & tFecha.Text
        .Footer = Now
    
        .TableBorder = tbColTopBottom
        .FontBold = True
        .AddTable "1000|5000", "", "Cant|Artículo", False, False
        .FontBold = False
        For Each oArt In colBultos
            .TableBorder = tbBottom
            .AddTable "1000|5000", "", oArt.cantidad & "|" & oArt.ArticuloNombre, False, False
            .TableBorder = tbNone
        Next
        .Paragraph = ""
        .Paragraph = "RECUERDE QUE ESTA LISTA PUEDE SER ALTERADA."
        .EndDoc
        .PrintDoc
    End With
    Exit Sub
errPB:
    objGral.OcurrioError "Error al intentar imprimir los bultos.", Err.Description, "Error"
End Sub

Private Sub MnuPrintConfig_Click()
    prj_GetPrinter True
End Sub

Private Sub MnuPrintDocumento_Click()
    tmModal.Tag = 5
    tmModal.Enabled = True
End Sub

Private Sub MnuPrintPapelRosa_Click()
On Error GoTo errPPR
    
    Dim sNroBandeja As String
    sNroBandeja = InputBox("Ingrese el número de bandeja.", "Bandeja copia eTicket")
    If Not IsNumeric(sNroBandeja) Then
        MsgBox "Valor incorrecto", vbExclamation, "Bandeja copia eTicket"
    Else
        GraboSeteoBandejaCopiaeTicket CInt(sNroBandeja)
    End If
    Exit Sub
errPPR:
    objGral.OcurrioError "Error al setear la bandeja.", Err.Description, "Configuración de bandeja"
End Sub

Private Sub MnuPrintReimprimir_Click()
Dim sID As String
    sID = InputBox("Ingrese el código de impresión a reimprimir.", "Reimpresión")
    If IsNumeric(sID) Then loc_ImprimoReparto CLng(sID)
End Sub


Private Sub MnuRecepEntregado_Click()
    If MsgBox("¿Confirma dar por entregado el envío " & vsGrid.Cell(flexcpText, vsGrid.Row, 0) & "?", vbQuestion + vbYesNo, "Dar por entregado") = vbYes Then
        db_RecepcionGraboEntregado False, Val(vsGrid.Cell(flexcpText, vsGrid.Row, 0)), True
    End If
End Sub


Private Sub sbStatus_PanelClick(ByVal Panel As MSComctlLib.Panel)

    If Panel.Key = "bultos" Then
        ArmoMenuBultos
        PopupMenu MnuBultos
    End If

End Sub

Private Sub sbStatus_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
Dim sDato As String
On Error GoTo errSel
    Select Case Panel.Key
        Case "tipo"
            sDato = InputBox("Ingrese parte o todo el nombre del tipo a seleccionar (vacio limpia datos)", "Seleccionar tipo")
            If sDato = "" Then
                sNTipo = ""
                iIDTipo = 0
            Else
                Dim objL As New clsListadeAyuda
                If objL.ActivarAyuda(cBase, "Select TipCodigo, TipNombre From Tipo Where TipNombre like '" & Replace(sDato, " ", "%") & "%'", 3000, 1, "Tipos") > 0 Then
                    sNTipo = objL.RetornoDatoSeleccionado(1)
                    iIDTipo = objL.RetornoDatoSeleccionado(0)
                    SaveSetting App.Title, "Settings", "AA" & Me.Name & "NTipo", sNTipo
                    SaveSetting App.Title, "Settings", "AA" & Me.Name & "IDTipo", iIDTipo
                End If
                Set objL = Nothing
            End If
        Case "articulo"
            sDato = InputBox("Ingrese el código del artículo a seleccionar (vacio limpia datos)", "Seleccionar tipo")
            If sDato = "" Then
                sCodArt = ""
                iIDArt = 0
            ElseIf Not IsNumeric(sDato) Then
                MsgBox "Ingrese un número.", vbCritical, "Atención"
            Else
                Dim rsA As rdoResultset
                Set rsA = cBase.OpenResultset("Select ArtID, ArtDescripcion from articulo where artCodigo = " & Val(sDato), rdOpenDynamic, rdConcurValues)
                If Not rsA.EOF Then
                    iIDArt = rsA(0)
                    sCodArt = Trim(rsA(1))
                    SaveSetting App.Title, "Settings", "AA" & Me.Name & "ArtDesc", sCodArt
                    SaveSetting App.Title, "Settings", "AA" & Me.Name & "IDArt", iIDArt
                Else
                    MsgBox "No hay un artículo con ese código.", vbExclamation, "Atención"
                End If
                rsA.Close
            End If
        Case "tarde"
            sDato = InputBox("Ingrese desde a que hora comienza la tarde (####).", "Hora Tarde", "1300")
            If sDato = "" Then Exit Sub
            If Len(sDato) <> 4 Or Not IsNumeric(sDato) Then
                MsgBox "El formato tiene que contener 4 dígitos." & vbCr & "NO MODIFICO EL VALOR", vbCritical, "Atención"
                Exit Sub
            Else
                If Val(sDato) < 0 Or Val(sDato) > 2359 Then
                    MsgBox "El formato tiene que contener 4 dígitos (entre 0001 hasta 2359)." & vbCr & "NO MODIFICO EL VALOR", vbCritical, "Atención"
                    Exit Sub
                Else
                    SaveSetting App.Title, "Parametros", "ComienzoHoraTarde", sDato
                    paHoraTarde = sDato
                    sbStatus.Panels("tarde").Text = "Tarde a partir: " & paHoraTarde & " "
                End If
            End If
            Exit Sub
        Case "bultos"
            ArmoMenuBultos
    End Select
    sbStatus.Panels("tipo").Text = IIf(sNTipo <> "", sNTipo, "Sin tipo asignado")
    sbStatus.Panels("articulo").Text = IIf(sCodArt <> "", sCodArt, "Sin artículo asignado")
    
    loc_SetStatusTipoArt
Exit Sub
errSel:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error inesperado.", Err.Description, "Status"
End Sub

Private Sub tCodigo_Change()
    If Val(tCodigo.Tag) > 0 Then tCodigo.Tag = "": loc_CleanRecepcion
End Sub
Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If Val(tCodigo.Tag) > 0 Then
            txtEnvioEntregado.SetFocus
            'vsGrid.SetFocus
        Else
            db_FillGridDatosCodImpresion
        End If
    End If
End Sub

Private Sub tFecha_Change()
    If Val(hlTab(0).Tag) < 6 Then
        vsGrid.Rows = 1
        s_ChangeData
    End If
End Sub

Private Sub tFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            s_FillGrid
    End Select
End Sub

Private Sub tmModal_Timer()
    tmModal.Enabled = False
    Select Case Val(tmModal.Tag)
        Case 1  'Cambiar camión
            loc_ShowCambiarEnvio 2, vsGrid.Cell(flexcpValue, vsGrid.Row, 0)
        Case 2  'Cambiar hora.
            loc_ShowCambiarEnvio 0, vsGrid.Cell(flexcpValue, vsGrid.Row, 0)
        Case 3
            loc_ShowDividirEnvio
        Case 4
            loc_ShowCambiarEnvio 1, vsGrid.Cell(flexcpValue, vsGrid.Row, 0)
        Case 5
            frmRePrint.Show vbModal
    Exit Sub
    End Select
    tmModal.Tag = ""
End Sub

Private Sub tooMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case LCase(Button.Key)
'        Case "cnfgprint": prj_GetPrinter True
        Case "save": act_Save
        Case "sugerircamion"
            tooMenu.Refresh
            PopupMenu MnuGridSugerirC, , Button.Left, Button.Top + Button.Height
        
        Case "envio"
            If vsGrid.Row >= vsGrid.FixedRows Then
                tooMenu.Refresh
                MnuGridEditEnvio_Click
            End If
            
        Case "print"
            tooMenu.Refresh
            MnuPrintAux.Enabled = (Val(hlTab(0).Tag) = 1 And vsGrid.Rows > vsGrid.FixedRows)
            MnuPrintBultos.Enabled = MnuPrintAux.Enabled
            PopupMenu MnuPrint, , Button.Left, Button.Top + Button.Height
            
        Case "otros"
            PopupMenu MnuOtros, , Button.Left, Button.Top + Button.Height
            
        Case "mercaderia"
            PopupMenu MnuMercaderia, , tooMenu.Buttons("mercaderia").Left, tooMenu.Buttons("mercaderia").Height
    End Select

End Sub

Private Sub txtEnvioEntregado_GotFocus()
    With txtEnvioEntregado
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Function CargoEnvíoPorDocumento(ByVal barcode As String) As Long
On Error GoTo errCED
    Screen.MousePointer = 11
    
    Dim iTipo As Byte
    Dim codDoc As Long
    barcode = Trim(barcode)
    If InStr(1, barcode, "d", vbTextCompare) = 0 Then
        MsgBox "Formato incorrecto.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    iTipo = Mid(barcode, 1, InStr(1, barcode, "d", vbTextCompare) - 1)
    codDoc = Mid(barcode, InStr(1, barcode, "d", vbTextCompare) + 1)

    Cons = "SELECT EnvCodigo Envío " & _
        "FROM Envio INNER JOIN Documento ON EnvDocumento = DocCodigo AND DocTipo = " & iTipo & " AND DocCodigo = " & codDoc & _
        " WHERE EnvCodImpresion = " & Val(tCodigo.Text)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        CargoEnvíoPorDocumento = RsAux(0)
        RsAux.MoveNext
        If Not RsAux.EOF Then
            CargoEnvíoPorDocumento = 0
            Dim oHelp As New clsListadeAyuda
            oHelp.CerrarSiEsUnico = True
            If oHelp.ActivarAyuda(cBase, Cons, 3000, , "Buscar envío por documento") > 0 Then
                CargoEnvíoPorDocumento = oHelp.RetornoDatoSeleccionado(0)
            End If
            Set oHelp = Nothing
        End If
    Else
        MsgBox "No hay datos para el código ingresado.", vbInformation, "Búsqueda"
    End If
    Screen.MousePointer = 0
    
    Exit Function
errCED:
    objGral.OcurrioError "Error al buscar el envío.", Err.Description, "Cargar envío por documento"
End Function


Private Function ValidoEnvioCodImpresion(ByVal idEnvio As Long) As Boolean
On Error GoTo errVEI
    Cons = "SELECT EnvCodigo Envío " & _
        "FROM Envio " & _
        " WHERE EnvCodImpresion = " & Val(tCodigo.Text) & _
        " AND EnvCodigo = " & idEnvio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        ValidoEnvioCodImpresion = True
    End If
    RsAux.Close
    Exit Function
errVEI:
    objGral.OcurrioError "Error al buscar el envío.", Err.Description, "Cargar envío por documento"
End Function


Private Sub txtEnvioEntregado_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Len(txtEnvioEntregado.Text) > 0 Then
            If Val(tCodigo.Text) = 0 Then
                MsgBox "Es necesario tener ingresado el código de impresión.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
            
            Dim bDirecto As Boolean
            Dim idEnvio As Long
            If InStr(1, txtEnvioEntregado.Text, "D", vbTextCompare) > 0 Then
                idEnvio = CargoEnvíoPorDocumento(txtEnvioEntregado.Text)
            ElseIf InStr(1, txtEnvioEntregado.Text, "EB", vbTextCompare) = 1 Then
                idEnvio = Mid(txtEnvioEntregado.Text, 3)
            ElseIf IsNumeric(txtEnvioEntregado.Text) Then
                bDirecto = True
                If ValidoEnvioCodImpresion(Val(txtEnvioEntregado.Text)) Then
                    idEnvio = Val(txtEnvioEntregado.Text)
                Else
                    MsgBox "El envío no pertenece al código de impresión ingresado.", vbExclamation, "ATENCIÓN"
                End If
            End If
            
            If idEnvio > 0 Then
                If bDirecto Then
                    If MsgBox("¿Confirma dar por entregado el envío " & idEnvio & "?", vbQuestion + vbYesNo, "Dar por entregado") = vbNo Then
                        Exit Sub
                    End If
                End If
                txtEnvioEntregado.Enabled = False
                db_RecepcionGraboEntregado False, idEnvio, False
                txtEnvioEntregado.Enabled = True
                'Else
                    txtEnvioEntregado.Text = ""
                    txtEnvioEntregado.SetFocus
                'End If
            End If
        End If
    End If
End Sub

Private Sub vsGrid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'Dada la columna y el tag hago lo sgte
' primero me fijo si está expandido sino supongo que si
' segundo ahora considero que este es el mayor width para la columna.
' tercero guardo el nuevo seteo.
Dim vCol() As String, vSize() As String
    On Error Resume Next
    With vsGrid
        vCol = Split(.Tag, "|")
        If .Cell(flexcpData, 0, Col) = 0 And Col > 3 Then
            'Supongo que expande.
            .Cell(flexcpPicture, 0, Col) = imgMini.ListImages(IIf(.Cell(flexcpData, 0, Col) = 0, 2, 1)).Picture
            .Cell(flexcpData, 0, Col) = IIf(.Cell(flexcpData, 0, Col) = 0, 1, 0)
            vSize = Split(vCol(Col), ":")
            If UBound(vSize) >= 1 Then
                If .ColWidth(Col) < Val(vSize(0)) Then vSize(2) = vSize(0) Else vSize(2) = .ColWidth(Col)
                vSize(1) = 1
            End If
            vCol(Col) = Join(vSize, ":")
        Else
            If Col < 4 Then
                vCol(Col) = .ColWidth(Col)
            Else
                vSize = Split(vCol(Col), ":")
                If UBound(vSize) >= 1 Then
                    If .ColWidth(Col) < Val(vSize(0)) Then vSize(2) = vSize(0) Else vSize(2) = .ColWidth(Col)
                    vSize(1) = 1
                End If
                vCol(Col) = Join(vSize, ":")
            End If
        End If
        .Tag = Join(vCol, "|")
        loc_SaveSettingColGrid
    End With
    Erase vCol
    Erase vSize

End Sub

Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col <> 0)
End Sub

Private Sub vsGrid_Click()
On Error GoTo errGC
    If vsGrid.MouseRow = 0 And vsGrid.MouseCol > 3 Then
        Dim vCol() As String
        Dim vSize() As String
        With vsGrid
            vCol = Split(.Tag, "|")
            If InStr(1, vCol(.Col), ":") > 0 Then
                vSize = Split(vCol(.Col), ":")
                .ColWidth(.Col) = IIf(.Cell(flexcpData, 0, .Col) = 0, vSize(2), vSize(0))
                .Cell(flexcpPicture, 0, .Col) = imgMini.ListImages(IIf(.Cell(flexcpData, 0, .MouseCol) = 0, 2, 1)).Picture
                .Cell(flexcpData, 0, .Col) = IIf(.Cell(flexcpData, 0, .Col) = 0, 1, 0)
                vSize(1) = Val(.Cell(flexcpData, 0, .Col))
                vCol(.Col) = Join(vSize, ":")
                .Tag = Join(vCol, "|")
            End If
        End With
        loc_SaveSettingColGrid
    Else
        If vsGrid.Col = 0 And Val(hlTab(0).Tag) = 0 Then loc_QEnviosClic
    End If
    Exit Sub
errGC:
'Este error lo puse por debug ya que en el comercio desaparece el status.
    MsgBox "Error: " & Err.Description, vbCritical, "Grilla clic"
End Sub

Private Sub vsGrid_DblClick()
    If Val(hlTab(0).Tag) = 6 Then
        On Error Resume Next
            
        Cons = "SELECT EnvCodigo Envío, DocSerie + '-' + Cast(Docnumero as varchar(7)) Remito, RenARetirar as Cantidad  " & _
            "FROM envio INNER JOIN Camion ON EnvCamion = CamCodigo INNER JOIN EnviosRemitos ON EnvCodigo = EReEnvio " & _
            "INNER JOIN Documento ON EReRemito = DocCodigo and DocTipo = 48 INNER JOIN Renglon ON RenDocumento = DocCodigo AND RenARetirar > 0 " & _
            "INNER JOIN Articulo ON ArtId = REnArticulo WHERE EnvEstado in (3,4) AND ArtId = " & vsGrid.Cell(flexcpData, vsGrid.RowSel, 0) & _
            " AND CamCodigo = " & vsGrid.Cell(flexcpData, vsGrid.RowSel, 1)
            
        Dim objH As New clsListadeAyuda
        objH.ActivarAyuda cBase, Cons, 6000, 0, "Remitos"
        Set objH = Nothing
    
    End If
End Sub

Private Sub vsGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 93: frm_ShowPopUp
    End Select
End Sub

Private Sub vsGrid_LostFocus()
On Error Resume Next
    tooMenu.Buttons("envio").Enabled = False
End Sub

Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Shift <> 0 Then Exit Sub
    If Button = 2 And vsGrid.Rows > vsGrid.FixedRows Then
        frm_ShowPopUp
    End If
End Sub

Private Sub vsGrid_RowColChange()
On Error Resume Next
    tooMenu.Buttons("envio").Enabled = vsGrid.Row >= 1 'And Val(hlTab(0).Tag) = 5)
End Sub

Private Sub s_DeleteMenuCamion()
Dim iQ As Integer
    For iQ = MnuAIECamion.LBound + 1 To MnuAIECamion.UBound
        Unload MnuAIECamion(iQ)
    Next
End Sub

Private Function f_GetCamionPorVolumen(ByVal bEsAgencia As Boolean, ByVal cEnvVT As Currency, ByVal lTFlete As Long, _
                                                            ByVal lIDZonaoAgencia As Long, ByVal sFecha As String) As Integer
Dim rsC As rdoResultset
    
    f_GetCamionPorVolumen = 0
    Cons = " And CamCodigo IN " _
            & "(Select CTFCamion From CamionFlete Where CTFTipoFlete = " & lTFlete & ")"
    If bEsAgencia Then
        Cons = "Select CAgPrioridad, CamCodigo, IsNull(CamVolumen, 0) as CamVolumen From Camion, CamionAgencia" _
            & " Where CamCodigo = CAgCamion " _
            & " And CAgAgencia = " & lIDZonaoAgencia _
            & Cons _
            & " Order by CAgPrioridad"
    Else
        Cons = "Select CZoPrioridad, CamCodigo, IsNull(CamVolumen, 0) as CamVolumen From Camion, CamionZona" _
            & " Where CamCodigo = CZoCamion " _
            & " And CZoZona = " & lIDZonaoAgencia _
            & Cons _
            & " Order by CZoPrioridad"
    End If
    Set rsC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    'Recorro hasta encontrar el primer camión que tenga lugar.
    Do While Not rsC.EOF
        If cEnvVT > 0 Then
            If cEnvVT <= rsC!CamVolumen - f_GetVolumenAdjudicadoAlCamion(rsC!CamCodigo, sFecha) Then
                f_GetCamionPorVolumen = rsC!CamCodigo
                Exit Do
            End If
        Else
            If rsC!CamVolumen > f_GetVolumenAdjudicadoAlCamion(rsC!CamCodigo, sFecha) Then
                f_GetCamionPorVolumen = rsC!CamCodigo
                Exit Do
            End If
        End If
        rsC.MoveNext
    Loop
    rsC.Close
End Function

Private Function f_GetCamionAutomático(ByVal lEnvio As Long) As Integer
On Error GoTo errGCA
Dim rsE As rdoResultset
    f_GetCamionAutomático = 0
    Cons = "Select EnvFechaPrometida, IsNull(EnvZona, 0) as Z, IsNull(EnvAgencia, 0) as EAge, IsNull(EnvVolumenTotal, 0) as VT, EnvTipoFlete From Envio Where EnvCodigo = " & lEnvio
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsE.EOF Then
        MsgBox "El envío " & lEnvio & " fue eliminado, verifique.", vbExclamation, "Atención"
    Else
        If IsNull(rsE("Z")) Then
            MsgBox "El envío " & lEnvio & " no tiene asignada la zona de reparto.", vbCritical, "Atención"
        Else
            f_GetCamionAutomático = f_GetCamionPorVolumen((rsE!EAge > 0), CCur(rsE!VT), rsE!EnvTipoFlete, IIf(rsE!EAge > 0, rsE!EAge, rsE!Z), rsE!EnvFechaPrometida)
        End If
    End If
    rsE.Close
Exit Function
errGCA:
    Screen.MousePointer = 0
    objGral.OcurrioError "Error al buscar la asignación automática de camión.", Err.Description, "Error (getcamionautomatico)"
End Function

Private Function f_GetVolumenAdjudicadoAlCamion(ByVal lCamion As Integer, ByVal sFecha As String)
Dim rsV As rdoResultset
    
    f_GetVolumenAdjudicadoAlCamion = 0
    Cons = "Select IsNull(SUM(EnvVolumenTotal), 0) " _
        & " From Envio (index = iFePEstHabCamZon) " _
        & " Where EnvFechaPrometida = '" & Format(CDate(sFecha), "mm/dd/yyyy") & "'" _
        & " And EnvEstado = " & EstadoEnvio.AImprimir _
        & " And EnvCamion = " & lCamion & " And EnvZona Is Not Null And EnvDocumento Is Not Null"
    Set rsV = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    f_GetVolumenAdjudicadoAlCamion = rsV(0)
    rsV.Close

End Function

Private Function f_GetCamionByID(ByVal iIDCamion As Integer) As Integer
On Error GoTo errGC
Dim iQ As Integer
    f_GetCamionByID = 0
    For iQ = LBound(arrCamion) To UBound(arrCamion)
        If arrCamion(iQ).Codigo = iIDCamion Then f_GetCamionByID = iQ: Exit For
    Next
errGC:
End Function

Private Sub frm_HideShowMenuCamion(ByVal iID As Integer)
On Error Resume Next
Dim iQ As Integer
    For iQ = MnuGridCamion.LBound To MnuGridCamion.UBound
        MnuGridCamion(iQ).Visible = (iID <> Val(MnuGridCamion(iQ).Tag))
    Next
End Sub

Private Sub s_SetCamionSugerido(ByVal bSelect As Boolean)
Dim iQ As Integer, iIDC As Integer
    Screen.MousePointer = 11
    With vsGrid
        For iQ = .FixedRows To .Rows - 1
            If (.Cell(flexcpChecked, iQ, 0) = flexChecked And bSelect) Or (.Cell(flexcpData, iQ, 3) = 0 And Not bSelect) Then
                iIDC = f_GetCamionAutomático(.Cell(flexcpValue, iQ, 0))
                If iIDC > 0 Then
                    iIDC = f_GetCamionByID(iIDC)
                    If iIDC > 0 Then
                        .Cell(flexcpText, iQ, 4) = arrCamion(iIDC).Nombre
                        .Cell(flexcpData, iQ, 3) = arrCamion(iIDC).Codigo
                    End If
                End If
            End If
        Next iQ
    End With
    Screen.MousePointer = 0
End Sub

Private Sub s_VerficoEnvioEditado(ByVal lCamion As Long, ByVal eEstado As EstadoEnvio)
Dim sFM As String
Dim lLastID As Long, iCol As Byte
On Error GoTo errVEE
    
    ReDim arrNomArt(0)
    Cons = "Select EnvCodigo, EnvFModificacion, IsNull(EnvRangoHora, '') as RH, EnvTipo, LocCodigo, LocNombre, " _
                & " CalNombre, DirPuerta, DirLetra, DirApartamento, rTrim(TFlNombreCorto) as TF, " _
                & " IsNull(EVCID, 0) as VC, IsNull(EnvComentario, '') as Memo, IsNull(rTrim(AgeNombre), rTrim(ZonNombre)) as ZN, REvCantidad, IsNull(AEsID, ArtCodigo) ArtCodigo, IsNull(AEsNombre, ArtNombre) as AD, ArtTipo, IsNull(AEsID, ArtID) ArtID " _
        & " From ((((((((((Envio Left Outer Join Agencia On EnvAgencia = AgeCodigo) LEFT OUTER JOIN EnvioVaCon ON EVCEnvio = EnvCodigo)" _
        & " INNER JOIN Direccion ON EnvDireccion = DirCodigo)" _
        & " INNER JOIN Calle ON DirCalle = CalCodigo) INNER JOIN Localidad ON CalLocalidad = LocCodigo) " _
        & " INNER JOIN TipoFlete ON EnvTipoFlete = TFlCodigo)" _
        & " INNER JOIN Zona ON EnvZona = ZonCodigo ) INNER JOIN RenglonEnvio ON EnvCodigo = REvEnvio) " _
        & " INNER JOIN Articulo ON REvArticulo = ArtID)" _
        & " LEFT OUTER JOIN ArticuloEspecifico ON AEsDocumento = EnvDocumento AND AEsArticulo = REvArticulo AND ((AEsTipoDocumento IN (1, 6) AND ENVTipo = 1) OR (AEsTipoDocumento IN (7, 33) AND ENVTipo = 3)))" _
        & " Where (EnvCodigo = " & vsGrid.Cell(flexcpValue, vsGrid.Row, 0) _
        & " OR EnvCodigo IN (Select EVCEnvio From EnvioVaCon WHERE EVCID = (SELECT EVCID FROM EnvioVaCon WHERE EVCEnvio = " & vsGrid.Cell(flexcpValue, vsGrid.Row, 0) & ")))" _
        & " And EnvFechaPrometida = '" & Format(tFecha.Text, "mm/dd/yyyy") & "'" _
        & " And EnvEstado = " & eEstado & " And EnvZona = ZonCodigo " _
        & " And EnvTipo In (1, 2, 3) And EnvCamion " & IIf(lCamion > 0, " = " & lCamion, " Is Null") _
        & " And EnvDocumento > 0" _
        & " Order by EVCID, EnvCodigo"
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        vsGrid.RemoveItem vsGrid.Row
    Else
        If Val(hlTab(0).Tag) = 0 Then iCol = 8 Else iCol = 7
        'Limpio los artículos y los tipos de artículos
        vsGrid.Cell(flexcpText, vsGrid.Row, iCol) = ""
        vsGrid.Cell(flexcpData, vsGrid.Row, 7) = ""
        vsGrid.Cell(flexcpData, vsGrid.Row, 6) = ""
            
        Do While Not RsAux.EOF
            
            With vsGrid
                
                If lLastID = 0 Then
                    lLastID = RsAux("EnvCodigo")
                    
                    If Val(hlTab(0).Tag) = 0 Then
                        .Cell(flexcpText, .Row, 1) = RsAux!RH
                        .Cell(flexcpText, .Row, 5) = RsAux!ZN
                        .Cell(flexcpText, .Row, 6) = f_GetDireccionRsAux
                        .Cell(flexcpText, .Row, 7) = Trim(RsAux!Memo)
                    Else
                        .Cell(flexcpText, .Row, 2) = RsAux!RH
                        .Cell(flexcpText, .Row, 4) = RsAux!ZN
                        .Cell(flexcpText, .Row, 5) = f_GetDireccionRsAux
                        .Cell(flexcpText, .Row, 6) = Trim(RsAux!Memo)
                    End If
                    .Cell(flexcpText, .Row, 3) = RsAux!TF
                    'DATA
                    sFM = RsAux!EnvFModificacion: .Cell(flexcpData, .Row, 1) = sFM         'F Modificado
                    sFM = RsAux!VC: .Cell(flexcpData, .Row, 2) = Val(sFM)                      'Va Con
                End If
                '..........................ARTICULOS
                'Si no va con otro cargo los artículos del resultado sino busco la suma de todos los artículos para todos los envíos involucrados.
                
                If Val(hlTab(0).Tag) = 0 Then iCol = 8 Else iCol = 7
                
                '..........................ARTICULOS
                loc_AddArtNombre RsAux("ArtID"), Format(RsAux("ArtCodigo"), "(#,000,000)") & " " & RsAux("AD")
                loc_SetQTipoArt .Row, RsAux("ArtID"), RsAux("ArtTipo"), RsAux("REvCantidad")
                .Cell(flexcpText, .Row, iCol) = fnc_GetStringArticulos(.Cell(flexcpData, .Row, 6))
                '..........................Artículos
                
            End With
            RsAux.MoveNext
        Loop
    End If
    RsAux.Close
    s_SetMenu
    Exit Sub
errVEE:
    objGral.OcurrioError "Error al validar el envío editado.", Err.Description
End Sub

Private Sub vsPrint_NewTableCell(Row As Integer, Column As Integer, Cell As String)
    If bPrintPlanilla Then
        If Column = 1 Then
            vsPrint.FontBold = True
        Else
            'Si imprimo los artículos entonces va en bold
            vsPrint.FontBold = bPrintColEsArt
        End If
    End If
End Sub

Private Sub ArmoMenuBultos()
Dim iQ As Integer
Dim oArt As clsArticuloMenu

    For iQ = 1 To MnuBulIdx.Count - 1
        Unload MnuBulIdx(iQ)
    Next
    MnuBulIdx(0).Caption = ""
    For Each oArt In colBultos
        If MnuBulIdx(0).Caption <> "" Then
            Load MnuBulIdx(MnuBulIdx.UBound + 1)
        End If
        MnuBulIdx(MnuBulIdx.UBound).Caption = oArt.cantidad & Space(5) & oArt.ArticuloNombre
    Next
    
    
End Sub

'Private Sub TestingSort()
'    Dim col1 As New Collection
'    Set col1 = OrdenoGrillaPorIDEnvio(col1)
'
'    Dim oNodo As New clsDocToPrint
'    oNodo.idEnvio = 10
'    col1.Add oNodo
'
'    Set col1 = OrdenoGrillaPorIDEnvio(col1)
'
'
'    Set oNodo = New clsDocToPrint
'    oNodo.idEnvio = 6
'    col1.Add oNodo
'
'    Set oNodo = New clsDocToPrint
'    oNodo.idEnvio = 1
'    col1.Add oNodo
'
'
'    Set oNodo = New clsDocToPrint
'    oNodo.idEnvio = 5
'    col1.Add oNodo
'
'
'    Set oNodo = New clsDocToPrint
'    oNodo.idEnvio = 4
'    col1.Add oNodo
'
'    Dim col2 As New Collection
'    Set col2 = OrdenoGrillaPorIDEnvio(col1)
'
'End Sub

'Private Function OrdenoGrillaPorIDEnvio(ByVal Col As Collection) As Collection
'    If Col.Count = 0 Then Set OrdenoGrillaPorIDEnvio = Col: Exit Function
'
'    Dim nodo1 As clsDocToPrint
'    Dim nodo2 As clsDocToPrint
'    Dim nodo3 As clsDocToPrint
'    Dim pos As Integer
'    pos = 1
'    Do While pos < Col.Count
'        Set nodo1 = Col(pos)
'        Set nodo2 = Col(pos + 1)
'        If nodo1.idEnvio < nodo2.idEnvio Then
'            pos = pos + 1
'        Else
'            Set nodo3 = New clsDocToPrint
'            nodo3.Documento = nodo2.Documento
'            nodo3.idEnvio = nodo2.idEnvio
'            nodo3.tipo = nodo2.tipo
'
'            nodo2.Documento = nodo1.Documento
'            nodo2.idEnvio = nodo1.idEnvio
'            nodo2.tipo = nodo1.tipo
'
'            nodo1.Documento = nodo3.Documento
'            nodo1.idEnvio = nodo3.idEnvio
'            nodo1.tipo = nodo3.tipo
'
'            pos = 1
'        End If
'    Loop
'    Set OrdenoGrillaPorIDEnvio = Col
'
'End Function

