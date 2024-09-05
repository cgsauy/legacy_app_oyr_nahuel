VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form MaEmbarque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Embarques"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10590
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MaEmbarque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   10200
      Top             =   4920
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   27
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "carpeta"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "limpiar"
            Object.ToolTipText     =   "Limpiar datos"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "primero"
            Object.ToolTipText     =   "Ir al Primer Registro"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "anterior"
            Object.ToolTipText     =   "Ir al Registro Anterior"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1200
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "siguiente"
            Object.ToolTipText     =   "Ir al Siguiente Registro"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ultimo"
            Object.ToolTipText     =   "Ir al Último Registro"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
            Object.ToolTipText     =   "Refrescar datos en combos"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "transporte"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "flete"
            Object.ToolTipText     =   "Mantenimiento de Fletes"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "linea"
            Object.ToolTipText     =   "Mantenimiento de Líneas de Fletes"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   2000
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.TextBox tRegistro 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4220
         TabIndex        =   67
         Top             =   50
         Width           =   975
      End
   End
   Begin AACombo99.AACombo cProveedor 
      Height          =   315
      Left            =   6960
      TabIndex        =   5
      Top             =   840
      Width           =   3255
      _ExtentX        =   5741
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
   Begin VB.TextBox tFApertura 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4080
      MaxLength       =   12
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox tCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Embarque"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4035
      Left            =   120
      TabIndex        =   41
      Top             =   2160
      Width           =   10335
      Begin VSFlex6DAOCtl.vsFlexGrid vsContenedor 
         Height          =   1155
         Left            =   7140
         TabIndex        =   70
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2037
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsArticulo 
         Height          =   1575
         Left            =   4560
         TabIndex        =   32
         Top             =   1800
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2778
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
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
         ShowComboButton =   0   'False
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   1080
         TabIndex        =   34
         Top             =   2760
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         BackColor       =   12648447
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
         Left            =   1080
         TabIndex        =   23
         Top             =   1320
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
      Begin AACombo99.AACombo cDestino 
         Height          =   315
         Left            =   4560
         TabIndex        =   16
         Top             =   600
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
      Begin AACombo99.AACombo cAgencia 
         Height          =   315
         Left            =   7620
         TabIndex        =   11
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
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
      Begin AACombo99.AACombo cOrigen 
         Height          =   315
         Left            =   4560
         TabIndex        =   9
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
      Begin AACombo99.AACombo cTransporte 
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
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
      Begin AACombo99.AACombo cMTransporte 
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
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
      Begin VB.TextBox tUltFechaEmbarque 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   31
         Text            =   "22-Oct-1999"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox tConocimiento 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox tFlete 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         MaxLength       =   12
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chFletePago 
         Alignment       =   1  'Right Justify
         Caption         =   "&Flete Pago:"
         Height          =   255
         Left            =   5880
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox tEmbPrevisto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   25
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox tEmbarco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         MaxLength       =   12
         TabIndex        =   27
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox tArriboPrevisto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   29
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox tDivisa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   35
         Top             =   2760
         Width           =   1710
      End
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   1080
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   3420
         Width           =   9015
      End
      Begin VB.TextBox tArbitraje 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2880
         TabIndex        =   37
         Top             =   2400
         Width           =   1575
      End
      Begin AACombo99.AACombo cboPrioridad 
         Height          =   315
         Left            =   1080
         TabIndex        =   18
         Top             =   960
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
      Begin VB.Label lblGastoMVD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "28,199.90"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5880
         TabIndex        =   72
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Gasto MVD:"
         Height          =   255
         Left            =   4560
         TabIndex        =   71
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "&Últ. Fecha:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&BL:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "O&rigen:"
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "A&Gencia:"
         Height          =   255
         Left            =   6900
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "&Transporte:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "&Destino:"
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Prioridad:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lblInfoFlete 
         BackStyle       =   0  'Transparent
         Caption         =   "F&lete:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   960
         Width           =   495
      End
      Begin VB.Label labEmbCosteado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Costeado: NO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2760
         TabIndex        =   47
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "E&mb. Prev.:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "E&mbarcó:"
         Height          =   255
         Left            =   2400
         TabIndex        =   26
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Arribo Pre&v.:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Arribó:"
         Height          =   255
         Left            =   2400
         TabIndex        =   46
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label labArribo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "28-Oct-1999"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3120
         TabIndex        =   45
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Di&visa:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Comenta&io:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   3420
         Width           =   975
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Arb.:"
         Height          =   255
         Left            =   2400
         TabIndex        =   36
         Top             =   2400
         Width           =   435
      End
      Begin VB.Label labDivisa 
         BackStyle       =   0  'Transparent
         Caption         =   "Divisa Paga:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "&Local final:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label labDivisaPaga 
         BackStyle       =   0  'Transparent
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   43
         Top             =   3135
         Width           =   495
      End
      Begin VB.Label labFArriboLocal 
         Caption         =   "12-Oct-99"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3240
         TabIndex        =   42
         Top             =   1320
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   68
      Top             =   6255
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5362
            MinWidth        =   5362
            Text            =   "Tasa de Cambio"
            TextSave        =   "Tasa de Cambio"
            Key             =   "tasa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10663
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "ZUREO"
            TextSave        =   "ZUREO"
            Key             =   "zureo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":0BA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":0EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":11D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":13B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":158C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":1766
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":1940
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":1C5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":1E34
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":214E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":2468
            Key             =   "linea"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":2F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":305C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":31B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":3310
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaEmbarque.frx":346A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lAnulada 
      BackStyle       =   0  'Transparent
      Caption         =   "Anulada: 15/01/2001"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   69
      Top             =   540
      Width           =   2475
   End
   Begin VB.Label labPlazo 
      BackStyle       =   0  'Transparent
      Caption         =   "____"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7320
      TabIndex        =   66
      Top             =   1440
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin VB.Label labFormaPago 
      BackStyle       =   0  'Transparent
      Caption         =   "__________"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9240
      TabIndex        =   65
      Top             =   1440
      UseMnemonic     =   0   'False
      Width           =   975
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Plazo:"
      Height          =   255
      Left            =   6480
      TabIndex        =   64
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Forma de Pago:"
      Height          =   255
      Left            =   8040
      TabIndex        =   63
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Datos de la Carpeta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   320
      TabIndex        =   62
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label labComentario 
      BackStyle       =   0  'Transparent
      Caption         =   "___________________________________________________________________________________________________"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1320
      TabIndex        =   61
      Top             =   1680
      UseMnemonic     =   0   'False
      Width           =   8895
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario:"
      Height          =   255
      Left            =   360
      TabIndex        =   60
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label labCosteada 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "____________________"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9240
      TabIndex        =   59
      Top             =   1200
      UseMnemonic     =   0   'False
      Width           =   375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Costeada:"
      Height          =   255
      Left            =   8040
      TabIndex        =   58
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label labIncoterm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "____________________"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7320
      TabIndex        =   57
      Top             =   1200
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Incoterm:"
      Height          =   255
      Left            =   6480
      TabIndex        =   56
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label labLC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "____________________"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5040
      TabIndex        =   55
      Top             =   1440
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "L/C:"
      Height          =   255
      Left            =   4320
      TabIndex        =   54
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label labFactura 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "____________________"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5040
      TabIndex        =   53
      Top             =   1200
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Factura:"
      Height          =   255
      Left            =   4320
      TabIndex        =   52
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label labBcoCorresponsal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "____________________"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1800
      TabIndex        =   51
      Top             =   1440
      UseMnemonic     =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Bco. Corresponsal:"
      Height          =   255
      Left            =   360
      TabIndex        =   50
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label labBcoEmisor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "____________________"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1800
      TabIndex        =   49
      Top             =   1200
      UseMnemonic     =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Bco. Emisor:"
      Height          =   255
      Left            =   360
      TabIndex        =   48
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Proveedor:"
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Apertura:"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Código:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   120
      Top             =   480
      Width           =   10335
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
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
      Begin VB.Menu MnuLinea 
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
      Begin VB.Menu MnuLineaCarpeta 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCarpeta 
         Caption         =   "&Ir a Carpeta"
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuLineaBorrar 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEditMemo 
         Caption         =   "Modificar comentario"
      End
      Begin VB.Menu MnuLimpiar 
         Caption         =   "&Limpiar datos"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu MnuRegistros 
      Caption         =   "&Registros"
      Begin VB.Menu MnuPrimero 
         Caption         =   "&Primer Registro"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuAnterior 
         Caption         =   "Registro &Anterior"
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuSiguiente 
         Caption         =   "Registro &Siguiente"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuUltimo 
         Caption         =   "&Último Registro"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "MaEmbarque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cambios.
    '26-6-2000 : Si el embarque esta costeado dejo modificar el arbitraje y la divisa. Si cambia estos valores mando a otro subrubro <> de divisa.
    '                    Si la divisa esta paga, al recibo lo asigno tambien al nuevo documento.
    
    '6-7-00 : en modificoembarque cuando invoco hagonnotacredito le paso (divisaanterior / arbitraje) antes solo pasba la variable divisaanterior
                    
    '8-9-00: Si me ingresa un artículo que ya existe en la lista hago prorrateo de precio.
    
    '15/12/00 En modifico embarque me falto preguntar si modifico el arbitraje cuando la divisa esta paga y no hace nuevo embarque.
                    
    '28/5/2001 Cambios de manejo pedidos por matilde: Si ingresa costo cero que no pregunte ni modifique el costo del artículo. Y si el art. tiene costo cero que se pare para editar.
    
    '28/5/2001    Agregamos campos CarAnulada y le damos mensaje si no está ingresada la tasa de cambio del último día habil del mes requerido.
    
    '9-2001 Agregue actualizar combo, y acceso a mantenimiento de trasporte.
    '           Dejamos ingresar + de un pedido de boquilla al embarque.
    
    '11-2001  Al hacer nota de crédito o un crédito no registro gastos para compensar dif. de Cambio.
        
Option Explicit

'estos types me los defini para ingresar el gasto en zureo

Private Type tGastoNota
    idEmb As Long
    Valor As Currency
    DivAnterior As Currency
    DivPaga As Boolean
    SRubro As Long
End Type

Private Type tGastoMovimiento
    idEmb As Long
    idEmbViejo As Long          'Si este está cargado va nota.
    Valor As Currency
    HacerNotaPor As Currency
    DivPaga As Boolean
End Type


Private Type tGastoZureo
    idEmb As Long
    CodEmbarque As String
    Valor As Currency
    Fecha As Date
    Proveedor As Integer
    Carpeta As String
    SerieNro As String
    TipoDoc As Integer
    Arbitraje As Double
    BcoNombre As String
    BcoCodigo As Long
    LC As String
    FormaPago As String
    Cuenta As Long
    SaldoCero As Boolean
End Type

Private embarque As clsEmbarque


'RDO.---------------------------------------------------------
'Definición del entorno RDO
Private RsEmbarque As rdoResultset, RsAuxCar As rdoResultset
Private RsAuxE As rdoResultset

'Booleanas.---------------------------------------------------
Private sNuevo As Boolean, sModificar As Boolean

'String
Private sContenedor As String, sLinea As String

Private PrioridadDefault As Integer

'Property.----------------------------------------------------------
Private iSeleccionado As Long       'Código de Embarque seleccionado.
Private frmModal As Boolean

Public Property Get pModal() As Boolean
    pModal = iSeleccionado
End Property
Public Property Let pModal(Tipo As Boolean)
    frmModal = Tipo
End Property
Public Property Get pSeleccionado() As Long
    pSeleccionado = iSeleccionado
End Property
Public Property Let pSeleccionado(Codigo As Long)
    iSeleccionado = Codigo
End Property
Private Sub cAgencia_GotFocus()
    With cAgencia
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Seleccione una agencia."
End Sub
Private Sub cAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Foco cMTransporte
    End If
End Sub
Private Sub cAgencia_LostFocus()
    cAgencia.SelStart = 0
    Status.SimpleText = ""
    If (cAgencia.ListIndex > -1 And embarque.Agencia = 0) Then ValidoAsignarFlete
End Sub

Private Sub cAgencia_Validate(Cancel As Boolean)
    If (cAgencia.ListIndex > -1 And embarque.Agencia = 0) Then ValidoAsignarFlete
    'ValidoAsignarFlete
End Sub

Private Sub cboPrioridad_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If vbKeyReturn = KeyAscii Then Foco tFlete
End Sub

Private Sub cDestino_GotFocus()
    With cDestino
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Seleccione la ciudad destino."
End Sub
Private Sub cDestino_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cboPrioridad
End Sub
Private Sub cDestino_LostFocus()
    cDestino.SelStart = 0
    Status.SimpleText = ""
End Sub
Private Sub cLocal_GotFocus()
    With cLocal
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Seleccione el local donde arribará la mercadería."
End Sub
Private Sub cLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tEmbPrevisto
End Sub
Private Sub cLocal_LostFocus()
    Status.SimpleText = ""
    cLocal.SelStart = 0
End Sub
Private Sub cMoneda_GotFocus()
    With cMoneda
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Seleccione una moneda para la divisa."
End Sub
Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDivisa
End Sub
Private Sub cMoneda_LostFocus()
    Status.SimpleText = ""
    cMoneda.SelStart = 0
    If cMoneda.ListIndex > -1 Then
        Cons = "Select * From Moneda Where MonCodigo = " & cMoneda.ItemData(cMoneda.ListIndex)
        Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAuxE.EOF Then
            If RsAuxE!MonArbitraje Then tArbitraje.Enabled = True: tArbitraje.BackColor = Obligatorio Else tArbitraje.Enabled = False: tArbitraje.BackColor = Inactivo: tArbitraje.Text = ""
        End If
        RsAuxE.Close
    End If
End Sub
Private Sub cMTransporte_GotFocus()
    With cMTransporte
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Seleccione un medio de transporte."
End Sub
Private Sub cMTransporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cTransporte
End Sub
Private Sub cMTransporte_LostFocus()
    Status.SimpleText = ""
    cMTransporte.SelStart = 0
End Sub
Private Sub cOrigen_GotFocus()
    With cOrigen
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Seleccione la ciudad origen."
End Sub
Private Sub cOrigen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cAgencia
End Sub
Private Sub cOrigen_LostFocus()
    Status.SimpleText = ""
    cOrigen.SelStart = 0
End Sub
Private Sub cProveedor_GotFocus()
    With cProveedor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Seleccione un proveedor."
End Sub

Private Sub cProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrCs
    If vbKeyF1 = KeyCode And cProveedor.ListIndex > -1 Then
        Cons = "Select CarID, 'Código' = CarCodigo, Apertura = CarFApertura, Proveedor = PExNombre From Carpeta, ProveedorExterior " _
                & " Where CarProveedor = " & cProveedor.ItemData(cProveedor.ListIndex) _
                & " And CarProveedor = PExCodigo"
                
        If IsDate(tFApertura.Text) Then Cons = Cons & " And CarFApertura >= '" & Format(tFApertura.Text, "mm/dd/yy") & "'"
        Cons = Cons & " Order by CarFApertura"
        Screen.MousePointer = 11
        Dim objAyuda As New clsListadeAyuda
        If objAyuda.ActivarAyuda(cBase, Cons, 5500, 1, "Lista de Carpetas") Then
'        Ayuda.ActivoListaAyuda Cons, False, miconexion.TextoConexion(logImportaciones), 5500
'        Screen.MousePointer = 11
'        DoEvents
'        If Ayuda.ValorSeleccionado > 0 Then BuscoCarpeta Ayuda.ValorSeleccionado
            BuscoCarpeta objAyuda.RetornoDatoSeleccionado(0)
        End If
        Set objAyuda = Nothing
        Screen.MousePointer = 0
    End If
    Exit Sub
ErrCs:
    clsGeneral.OcurrioError "Error al instanciar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cProveedor_LostFocus()
    Status.SimpleText = ""
    cProveedor.SelStart = 0
    'If cProveedor.ListIndex > -1 And sNuevo Then CargoArticulosProveedor cProveedor.ItemData(cProveedor.ListIndex)
End Sub
Private Sub cTransporte_GotFocus()
    With cTransporte
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Seleccione un transporte."
End Sub
Private Sub cTransporte_KeyPress(KeyAscii As Integer)
Dim idLinea As Long, iPos As Integer
    
    If KeyAscii = vbKeyReturn Then
        
        'Veo si el transporte posee línea si no la tiene invoco el amb.
        If cTransporte.ListIndex > -1 Then
            If embarque.Transporte <> cTransporte.ItemData(cTransporte.ListIndex) Then
                Cons = "SELECT Codigo, Texto Linea FROM TransporteLinea INNER JOIN CodigoTexto ON Codigo = TLiLinea " & _
                    "WHERE TLiTransporte = " & cTransporte.ItemData(cTransporte.ListIndex) & " ORDER BY Texto"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If RsAux.EOF Then
                    MsgBox "Atención el transporte no posee lineas, se abrirá el formulario de ingreso.", vbExclamation, "Transporte sin línea"
                    EjecutarApp "Mantenimiento de Transportes.exe", cTransporte.ItemData(cTransporte.ListIndex), True
                End If
                RsAux.Close
            End If
        End If
        
        
        Dim idLineaActual As Long
        For iPos = 1 To vsContenedor.Rows - 1
            If Trim(vsContenedor.Cell(flexcpText, iPos, 1)) <> "" And Trim(vsContenedor.Cell(flexcpText, iPos, 2)) <> "" Then
                idLineaActual = Val(vsContenedor.Cell(flexcpText, iPos, 2))
                Exit For
            End If
        Next iPos
        
        'Verifico si la línea la contiene el vapor si no invoco al ingreso.
        If idLineaActual > 0 Then
            Cons = "SELECT TLiLinea FROM TransporteLinea WHERE TLiLinea = " & idLineaActual & _
                " AND TLiTransporte = " & cTransporte.ItemData(cTransporte.ListIndex)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If RsAux.EOF Then
                MsgBox "Atención el transporte no posee la linea asociada, se abrirá el formulario de ingreso.", vbExclamation, "Transporte sin línea asociada"
                EjecutarApp "Mantenimiento de Transportes.exe", cTransporte.ItemData(cTransporte.ListIndex), True
            End If
            RsAux.Close
        End If
        
        'Veo si hay contenedores ingresados y le cambio la línea.
        If cTransporte.ListIndex > -1 And vsContenedor.Rows > 1 Then
            idLinea = 0
            'Cons = "Select * From Transporte Where TraCodigo = " & cTransporte.ItemData(cTransporte.ListIndex)
            Cons = "SELECT Codigo, Texto Linea FROM TransporteLinea INNER JOIN CodigoTexto ON Codigo = TLiLinea " & _
                "WHERE TLiTransporte = " & cTransporte.ItemData(cTransporte.ListIndex) & " ORDER BY Texto"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            Do While Not RsAux.EOF
                'Me quedo con la primera a no ser que la siguiente sea la actual.
                If idLinea = 0 Then idLinea = RsAux("Codigo")
                If RsAux("Codigo") = idLineaActual Then
                    idLinea = RsAux("Codigo")
                    'Ya tengo la línea pero me pudo agregar contenedores.
                    Exit Do
                End If
                RsAux.MoveNext
            Loop
            RsAux.Close
            
            If idLinea > 0 Then
                For iPos = 1 To vsContenedor.Rows - 1
                    If Trim(vsContenedor.Cell(flexcpText, iPos, 1)) <> "" Then
                        vsContenedor.Cell(flexcpText, iPos, 2) = idLinea
                    End If
                Next iPos
            End If
        End If
        Foco cDestino
    End If
End Sub
Private Sub cTransporte_LostFocus()
    Status.SimpleText = ""
    cTransporte.SelStart = 0
End Sub

Private Sub chFletePago_GotFocus()
    Status.SimpleText = "Indique si el flete esta pago."
End Sub

Private Sub chFletePago_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        vsContenedor.SetFocus
        If vsContenedor.Rows > 1 Then vsContenedor.Select vsContenedor.Rows - 1, 0
    End If
End Sub

Private Sub Form_Activate()

    If iSeleccionado = -1 Then AccionNuevo
    Me.Refresh
    RsEmbarque.Requery
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    'Inicializo......................
    sNuevo = False: sModificar = False
    Set embarque = New clsEmbarque
    
    'CARGO LOS COMBOS.------------------------------------------------
    CargoCombosForm

    'Armo el formulario.--------------------
    sContenedor = ""
    Cons = "Select ConCodigo, ConAbreviacion from Contenedor" _
        & " Order by ConNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        If sContenedor = "" Then
            sContenedor = "#" & Trim(RsAux!ConCodigo) & ";" & Trim(RsAux!ConAbreviacion)
        Else
            sContenedor = sContenedor & "|" & "#" & Trim(RsAux!ConCodigo) & ";" & Trim(RsAux!ConAbreviacion)
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    sLinea = ""
    Cons = "Select Codigo, Texto From CodigoTexto Where Tipo = 67 Order by Texto"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        If sLinea = "" Then
            sLinea = "#" & Trim(RsAux!Codigo) & ";" & Trim(RsAux!Texto)
        Else
            sLinea = sLinea & "|" & "#" & Trim(RsAux!Codigo) & ";" & Trim(RsAux!Texto)
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    With vsContenedor
        .Rows = 2: .Cols = 3
        .ExtendLastCol = True
        .FormatString = "Cantidad|<Contenedor|<Línea"
        .ColDataType(0) = flexDTCurrency
        .ColComboList(1) = sContenedor
        .ColComboList(2) = sLinea
    End With
    
    'Presento la información de la tasa de Cambio del último mes anterior.
    BuscoTCParaPresentar
    
    InicioResultSet iSeleccionado
    
    Timer1.Enabled = True
    Timer1.Interval = 10
    
    Exit Sub
ErrLoad:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar el formulario."
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    RsEmbarque.Close
    CierroConexion
    rdoCZureo.Close
    Set objUsers = Nothing
End Sub

Private Sub labDivisaPaga_Click()
'Consulta pedida 3-8-2004
    If labDivisaPaga.Caption = "SI" Then s_ListaGastosDivisa
End Sub

Private Sub Label11_Click()
    Foco cAgencia
End Sub

Private Sub Label13_Click()
    Foco cMTransporte
End Sub

Private Sub Label18_Click()
    Foco cDestino
End Sub
Private Sub Label19_Click()
    Foco cboPrioridad
End Sub

Private Sub Label2_Click()
    Foco tCodigo
End Sub


Private Sub Label22_Click()
    Foco tEmbPrevisto
End Sub
Private Sub Label23_Click()
    Foco tEmbarco
End Sub

Private Sub Label24_Click()
    Foco tArriboPrevisto
End Sub
Private Sub Label29_Click()
    Foco cMoneda
End Sub
Private Sub Label3_Click()
    Foco tFApertura
End Sub

Private Sub Label30_Click()
    Foco tComentario
End Sub
Private Sub Label31_Click()
    Foco tArbitraje
End Sub

Private Sub Label32_Click()
    Foco cLocal
End Sub

Private Sub Label35_Click()
    Foco tUltFechaEmbarque
End Sub
Private Sub Label4_Click()
    Foco cProveedor
End Sub
Private Sub Label7_Click()
    Foco tConocimiento
End Sub
Private Sub Label9_Click()
    Foco cOrigen
End Sub

Private Sub lblInfoFlete_Click()
    Foco tFlete
End Sub

Private Sub MnuAnterior_Click()
    AccionRegistroAnterior
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuCarpeta_Click()
    AccionFormularioCarpeta
End Sub

Private Sub MnuEditMemo_Click()
    
    'Habilito y Desabilito Botones'-------------------------
    Botones False, False, False, True, True, Toolbar1, Me
    BotonesRegistro False, False, False, False, Toolbar1, Me
    Toolbar1.Buttons("carpeta").Enabled = False: MnuCarpeta.Enabled = False
    MnuEditMemo.Enabled = False
    '--------------------------------------------------------------
    MnuEditMemo.Tag = 1
    tComentario.Enabled = True: tComentario.BackColor = vbWhite
    InhabilitoCamposCarpeta
    tComentario.SetFocus

End Sub

Private Sub MnuEliminar_Click()

    AccionEliminar

End Sub

Private Sub MnuGrabar_Click()

    AccionGrabar

End Sub

Private Sub MnuLimpiar_Click()
    AccionLimpiar
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuPrimero_Click()
    AccionPrimerRegistro
End Sub

Private Sub MnuSiguiente_Click()
    AccionRegistroSiguiente
End Sub

Private Sub MnuUltimo_Click()
    AccionUltimoRegistro
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub AccionNuevo()
    
    'Si ya hay otro embarque dejo los valores que quedan iguales.
    If Not RsEmbarque.EOF Then
        'tCodigo.Text = Label2.Tag & "." & RsEmbarque!EmbCodigo
        'DejoCamposIguales
        MsgBox "Para ingresar un embarque parcial, DEBE CEDERLE Artículos del último embarque que no haya embarcado." _
            & "Seleccione el embarque y MODIFIQUELO.", vbExclamation, "ATENCIÓN"
    Else
        PreparoAccionNuevo
        'Automáticamente invoco al formulario de Carpetas.---------
        If InvocoNuevaCarpeta Then BuscoUltimosDatos Else AccionCancelar: Exit Sub
        iSeleccionado = 0
        tConocimiento.SetFocus
        '-----------------------------------------------------------------
    End If

End Sub

Private Sub AccionModificar()
On Error GoTo ErrAM
   
    RelojA
    sModificar = True
    'Habilito y Desabilito Botones'-------------------------
    Botones False, False, False, True, True, Toolbar1, Me
    BotonesRegistro False, False, False, False, Toolbar1, Me
    Toolbar1.Buttons("carpeta").Enabled = False: MnuCarpeta.Enabled = False
    MnuEditMemo.Enabled = False
    '--------------------------------------------------------------
    
    HabilitoCamposEmbarque
    InhabilitoCamposCarpeta
    
    'Si arribó no dejamos modificar el local de destino.
    If Not IsNull(RsEmbarque!EmbFArribo) Then cLocal.Enabled = False: cLocal.BackColor = Inactivo
    
    'Por las dudas cargo nuevamente los articulos.
    LimpioGrilla
    CargoArticulosEmbarque RsEmbarque!EmbID
    
    'Donde un Nivel este costeado no dejo modificar los artículos.
    If Not RsEmbarque!EmbCosteado Then
        'Verifico si hay subcarpetas costeadas.--------------
        Cons = "Select * From SubCarpeta Where SubEmbarque = " & RsEmbarque!EmbID _
            & " And SubCosteada = 1"
        Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAuxE.EOF Then
            vsArticulo.Enabled = False: vsArticulo.BackColor = Inactivo
            RelojD
            RsAuxE.Close: Exit Sub
        End If
        RsAuxE.Close
        'Verifico si hay otros embarques que esten costeados.--------------
        Cons = "Select * From Embarque Where EmbID <> " & RsEmbarque!EmbID _
            & " And EmbCarpeta = " & RsEmbarque!Embcarpeta _
            & " And EmbCosteado = 1"
        Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAuxE.EOF Then
            vsArticulo.Enabled = False: vsArticulo.BackColor = Inactivo
            RelojD
            RsAuxE.Close: Exit Sub
        End If
        RsAuxE.Close
        'Verifico si la carpeta esta costeada.--------------
        Cons = "Select * From Carpeta Where CarID = " & RsEmbarque!Embcarpeta _
            & " And CarCosteada = 1"
        Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAuxE.EOF Then
            vsArticulo.Enabled = False: vsArticulo.BackColor = Inactivo
            tDivisa.Enabled = False: tDivisa.BackColor = Inactivo
            tArbitraje.Enabled = False: tArbitraje.BackColor = Inactivo
        End If
        RsAuxE.Close
    Else
        vsArticulo.Enabled = False: vsArticulo.BackColor = Inactivo
        'tDivisa.Enabled = False: tDivisa.BackColor = Inactivo
        'tArbitraje.Enabled = False: tArbitraje.BackColor = Inactivo
        cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    End If
    '--------------------------------------------------------------
    
'    If RsEmbarque!EmbDivisaPaga Then
'        vsArticulo.Enabled = False: vsArticulo.BackColor = Inactivo
'        tDivisa.Enabled = False: tDivisa.BackColor = Inactivo
'        tArbitraje.Enabled = False: tArbitraje.BackColor = Inactivo
'        cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
'    End If
    
    If vsArticulo.Enabled Then
        
        If Not IsNull(RsEmbarque!EmbFLocal) Then
            MsgBox "La mercadería ya arribó y tiene ingreso en algún local, para modificar las cantidades debe seguir los siguientes pasos:" & vbCrLf & vbCrLf _
                & " 1) Elimine el ingreso al local (Administracion\Stock\INGRESO DE MERCADERÍA)." & vbCrLf _
                & " 2) Vuelva a cargar la carpeta y modifique el embarque." & vbCrLf _
                & " 3) Realice el arribo al local." _
                , vbInformation, "ATENCIÓN"
            vsArticulo.Enabled = False: vsArticulo.BackColor = Inactivo
            RelojD
            Exit Sub
        End If
        
        'Puede modificar los artículos.--------------------
        If Not IsNull(RsEmbarque!EmbFArribo) And IsNull(RsEmbarque!EmbFLocal) Then
            If IsNull(RsEmbarque!EmbFLocal) Then
                MsgBox "Si se realizan cambios de artículos o en sus cantidades se afectará el stock físico. NO PODRÁ ASIGNAR ARTÍCULOS A UN NUEVO EMBARQUE pues ya arribo al local.", vbInformation, "ATENCIÓN"
            Else
                MsgBox "Si se realizan cambios en la cantidad de artículos, se afectará el stock en la diferencia que exista con la mercadería arribada al local.", vbExclamation, "ATENCIÓN"
            End If
        End If
        'Veo si posee gastos.-------------------------------
        If ExistenGastos Then MsgBox "Existen gastos ingresados para los niveles al cual pertenece el embarque, si modifica los datos de un artículo afectará al costeo.", vbInformation, "ATENCIÓN"
    End If
    RelojD
    tConocimiento.SetFocus
    Exit Sub
ErrAM:
    clsGeneral.OcurrioError "Error al intentar dar acción modificar.", Trim(Err.Description)
    RelojD
    AccionCancelar
End Sub
Private Sub AccionGrabar()

    vsArticulo_LostFocus
    If Val(MnuEditMemo.Tag) = 1 Then
        On Error GoTo errEditMemo
        Cons = "Select * From Embarque Where EmbID = " & RsEmbarque!EmbID
        RsEmbarque.Close
        Set RsEmbarque = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsEmbarque.Edit
        If Trim(tComentario.Text) <> "" Then
            RsEmbarque("EmbComentario") = Trim(tComentario.Text)
        Else
            RsEmbarque("EmbComentario") = Null
        End If
        RsEmbarque.Update
        AccionCancelar
    Else
        If Not ValidoCampos Then Exit Sub
        FechaDelServidor
        If sNuevo Then NuevoEmbarque Else ModificoEmbarque
    End If
Exit Sub
errEditMemo:
    clsGeneral.OcurrioError "Error al intentar grabar el comentario.", Err.Description, "Grabar comentario"
End Sub
Private Sub AccionEliminar()
On Error GoTo ErrAE
Dim FechaModificado  As Date
Dim LetraEmb As String
Dim MuevoStock As Boolean
    
    If MsgBox("¿Confirma eliminar el embarque seleccionado?", vbQuestion + vbYesNo, "ELIMINAR EMBARQUE") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    'Controles de eliminación.--------------------------------------
    'Verifico que el embarque no posea subcarpetas.--------
    Cons = "Select * From SubCarpeta Where SubEmbarque = " & RsEmbarque!EmbID
    Set RsAuxCar = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAuxCar.EOF Then
        MsgBox "El embarque tiene asociadas subcarpetas, no se podrá eliminarlo.", vbInformation, "ATENCIÓN"
        Screen.MousePointer = 0
        Exit Sub
    End If
    RsAuxCar.Close
    '------------------------------------------------------------------------
    'Veo si posee Gastos.--------------------------------------------
    If ExistenGastos Then
        MsgBox "Existen gastos ingresados, no podrá eliminar el embarque.", vbExclamation, "ATENCIÓN"
        RelojD
        Exit Sub
    End If
    '------------------------------------------------------------------------
    MuevoStock = False
    If Not IsNull(RsEmbarque!EmbFArribo) Then
        If Not IsNull(RsEmbarque!EmbFLocal) Then
            'Si tiene FLocal y el local es zf o puerto no se cual es el local final.
            If RsEmbarque!EmbLocal = paLocalZF Or RsEmbarque!EmbLocal = paLocalPuerto Then
                MsgBox "El embarque arribó al local destino, al eliminar el embarque NO SE AFECTARÁ AL STOCK.", vbExclamation, "ATENCIÓN"
            Else
                MuevoStock = True
            End If
        Else
            MsgBox "El embarque arribó al local destino, al eliminar el embarque NO SE AFECTARÁ AL STOCK.", vbExclamation, "ATENCIÓN"
            'Si no es uno de estos locales, hay inconsistencias.------------------
            If RsEmbarque!EmbLocal = paLocalZF And RsEmbarque!EmbLocal = paLocalPuerto Then MuevoStock = True
        End If
    End If
    '-------------------------------------------------------------------------------------------
    FechaModificado = RsEmbarque!EmbFModificacion
    'Me quedo con la letra del embarque.--------------------------------------------
    LetraEmb = RsEmbarque!EmbCodigo
    Cons = "Select * From Embarque Where EmbID = " & RsEmbarque!EmbID
    RsEmbarque.Close
    
    On Error GoTo ErrInicio
    cBase.BeginTrans
    On Error GoTo ErrResumo
    
    Set RsEmbarque = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    'Verifico Multiusuario.-----------------------------------------------------------------
    If RsEmbarque.EOF Then GoTo EliminoEmbarque
    If FechaModificado <> RsEmbarque!EmbFModificacion Then GoTo ModificaronEmbarque
    
    'Muevo el Stock.------------------------------------------------------------------------
    If MuevoStock Then
        Cons = "Select * From ArticuloFolder" _
            & " Where AFoTipo = " & Folder.cFEmbarque & " And AFoCodigo = " & RsEmbarque!EmbID
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            MarcoMovimientoStockFisico UsuLogueado, TipoLocal.Deposito, RsEmbarque!EmbLocal, RsAux!AFoArticulo, RsAux!AFoCantidad, paEstadoArticuloEntrega, -1
            MarcoMovimientoStockTotal RsAux!AFoArticulo, TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, RsAux!AFoCantidad, -1
            MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), RsAux!AFoArticulo, RsAux!AFoCantidad, paEstadoArticuloEntrega, -1
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    '------------------------------------------------------------------------
    
    'Borro los artículos del embarque.-----------------------------
    Cons = "Delete ArticuloFolder Where AFoTipo = " & Folder.cFEmbarque & " And AFoCodigo = " & RsEmbarque!EmbID
    cBase.Execute (Cons)

    'Elimino el registro.
    RsEmbarque.Delete
    cBase.CommitTrans
    
    BuscoCarpeta tCodigo.Tag, ""
    Screen.MousePointer = 0
    Exit Sub
    
EliminoEmbarque:
    cBase.RollbackTrans
    RsEmbarque.Requery
    MsgBox "El embarque seleccionado ha sido eliminado por otra terminal, verifique.", vbInformation, "ATENCIÓN"
    GoTo Fin
    
ModificaronEmbarque:
    cBase.RollbackTrans
    RsEmbarque.Requery
    MsgBox "El embarque seleccionado ha sido modificado por otra terminal, verifique.", vbInformation, "ATENCIÓN"
    GoTo Fin
    
Fin:
    BuscoCarpeta tCodigo.Tag, LetraEmb
    Screen.MousePointer = 0
    Exit Sub

ErrInicio:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al iniciar la transacción."
    Exit Sub
    
ErrResumo:
    Resume ErrEliminacion
    
ErrEliminacion:
    cBase.RollbackTrans
    RsEmbarque.Requery
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar eliminar el embarque."
    Exit Sub

ErrAE:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar eliminar el embarque."
    
End Sub

Private Sub AccionCancelar()
    Screen.MousePointer = 11
    sNuevo = False: sModificar = False
    InhabilitoCamposEmbarque
    HabilitoCamposCarpeta
    
    MnuEditMemo.Tag = ""
    
    LimpioCamposEmbarque
    If Not RsEmbarque.EOF Then
        CargoDatosEmbarque
        CargoArticulosEmbarque RsEmbarque!EmbID
        Botones True, True, True, False, False, Toolbar1, Me
        MnuEditMemo.Enabled = True
        If RsEmbarque.RowCount > 1 Then BotonesRegistro True, True, True, True, Toolbar1, Me Else BotonesRegistro False, False, False, False, Toolbar1, Me
        IndicoRegistro
    Else
        LimpioCamposCarpeta
        Botones True, False, False, False, False, Toolbar1, Me
        MnuEditMemo.Enabled = False
        BotonesRegistro False, False, False, False, Toolbar1, Me
        tRegistro.Text = ""
    End If
    MnuCarpeta.Enabled = True: Toolbar1.Buttons("carpeta").Enabled = True
    Screen.MousePointer = 0
End Sub
Private Sub AccionLimpiar()
    iSeleccionado = 0
    InicioResultSet iSeleccionado
End Sub
Private Sub AccionFormularioCarpeta()
On Error GoTo ErrAFC
    If miconexion.AccesoAlMenu("MaCarpeta") Then
        Dim frmCarpeta As New MaCarpeta
        Screen.MousePointer = 11
        If RsEmbarque.EOF Then
            frmCarpeta.pSeleccionado = 0
        Else
            frmCarpeta.pSeleccionado = RsEmbarque!Embcarpeta
        End If
        'If frmModal Then frmCarpeta.Show vbModal, Me Else frmCarpeta.Show vbModeless
        frmCarpeta.Show vbModal, Me
        Set frmCarpeta = Nothing
        If Trim(tCodigo.Text) <> "" Then
            If Not IsNumeric(tCodigo.Text) Then
                BuscoCarpetaPorCodigo Mid(tCodigo.Text, 1, InStr(tCodigo.Text, ".") - 1), Mid(tCodigo.Text, InStr(tCodigo.Text, ".") + 1, Len(tCodigo.Text))
            Else
                BuscoCarpetaPorCodigo tCodigo.Text
            End If
        End If
    Else
        MsgBox "Ud. no posee permisos para acceder al formulario de Carpetas.", vbInformation, "ATENCIÓN"
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrAFC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al invocar el formulario de carpetas."
End Sub

Private Sub Status_PanelClick(ByVal Panel As MSComCtlLib.Panel)
    If Panel.Key = "zureo" Then
        fnc_DoTestLogin
    End If
End Sub

Private Sub tArbitraje_GotFocus()
    With tArbitraje
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Ingrese el arbitraje."
End Sub
Private Sub tArbitraje_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentario
End Sub
Private Sub tArbitraje_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub tArriboPrevisto_GotFocus()
    With tArriboPrevisto
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Ingrese la fecha prevista de arribo del embarque."
End Sub
Private Sub tArriboPrevisto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
         'Veo si la altero y si hay calendario descarga entonces aviso que voy a cambiar.
            If embarque.ID > 0 And embarque.Arribo = DateMinValue And embarque.ArriboPrevisto <> DateMinValue And embarque.ArriboPrevisto <> CDate(tArriboPrevisto.Text) Then
                If EmbarqueEnCalendario Then
                    If CDate(tArriboPrevisto.Text) > embarque.ArriboPrevisto Then
                        MsgBox "Se va a modificar la fecha en el calendario descarga agregando " & DateDiff("d", embarque.ArriboPrevisto, CDate(tArriboPrevisto.Text)) & " días.", vbInformation, "ATENCIÓN"
                    Else
                        MsgBox "Se va a modificar la fecha en el calendario descarga restando " & DateDiff("d", CDate(tArriboPrevisto.Text), embarque.ArriboPrevisto) & " días.", vbInformation, "ATENCIÓN"
                    End If
                End If
            End If
        Foco tUltFechaEmbarque
    End If
End Sub
Private Sub tArriboPrevisto_LostFocus()
    If IsDate(tArriboPrevisto.Text) Then tArriboPrevisto.Text = Format(CDate(tArriboPrevisto.Text), "dd/MM/yyyy") Else tArriboPrevisto.Text = ""
    Status.SimpleText = ""
End Sub


'Private Sub tCarpetaDespachante_GotFocus()
'    With tCarpetaDespachante
'        .SelStart = 0
'        .SelLength = Len(.Text)
'    End With
'    Status.SimpleText = " Ingrese el número de carpeta del despachante."
'End Sub
'Private Sub tCarpetaDespachante_KeyPress(KeyAscii As Integer)
'    If vbKeyReturn = KeyAscii Then Foco tFlete
'End Sub
'Private Sub tCarpetaDespachante_LostFocus()
'    Status.SimpleText = ""
'End Sub

Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese el código de embarque."
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If InStr(tCodigo.Text, ".") > 0 Then
            If Not IsNumeric(Mid(tCodigo.Text, 1, InStr(tCodigo.Text, ".") - 1)) Then
                MsgBox "El formato del código de carpeta no es numérico.", vbExclamation, "ATENCIO"
                Exit Sub
            End If
            BuscoCarpetaPorCodigo Mid(tCodigo.Text, 1, InStr(tCodigo.Text, ".") - 1), Mid(tCodigo.Text, InStr(tCodigo.Text, ".") + 1, Len(tCodigo.Text))
        Else
            If Not IsNumeric(tCodigo.Text) Then
                MsgBox "El formato del código de carpeta no es numérico.", vbExclamation, "ATENCIO"
                Exit Sub
            End If
            BuscoCarpetaPorCodigo tCodigo.Text
            If tCodigo.Text = "" Then
                If MsgBox("¿Desea ingresar una nueva carpeta?", vbQuestion + vbYesNo, "NUEVA CARPETA") = vbYes Then
                    AccionNuevo
                End If
            End If
        End If
    End If
End Sub
Private Sub tCodigo_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub tComentario_GotFocus()
    With tComentario
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Ingrese un comentario para el embarque."
End Sub
Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub
Private Sub tComentario_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub tConocimiento_GotFocus()
    With tConocimiento
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Ingrese el conocimiento de embarque."
End Sub
Private Sub tConocimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cOrigen
End Sub
Private Sub tConocimiento_LostFocus()
    Status.SimpleText = ""
    tConocimiento.Text = UCase(tConocimiento.Text)
End Sub

Private Sub tDivisa_GotFocus()
    With tDivisa
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese la divisa."
End Sub
Private Sub tDivisa_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tArbitraje.Enabled Then Foco tArbitraje Else Foco tComentario
    End If
End Sub
Private Sub tDivisa_LostFocus()
    Status.SimpleText = ""
    If IsNumeric(tDivisa.Text) Then tDivisa.Text = Format(tDivisa.Text, FormatoMonedaP)
End Sub

Private Sub tEmbarco_GotFocus()
    With tEmbarco
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = ""
End Sub


Private Sub tEmbarco_KeyPress(KeyAscii As Integer)
On Error GoTo ErrFE
    If KeyAscii = vbKeyReturn Then
        
        If IsDate(tEmbarco.Text) Then
            If cOrigen.ListIndex <> -1 Then
                Screen.MousePointer = 11
                Cons = "Select CiuDemora From Ciudad" _
                    & " Where CiuCodigo = " & cOrigen.ItemData(cOrigen.ListIndex)
                Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAuxE.EOF Then
                    If Not IsNull(RsAuxE!CiuDemora) Then tArriboPrevisto.Text = Format(SumoDias(tEmbarco.Text, RsAuxE!CiuDemora), "d-Mmm-yy") Else MsgBox "No hay datos de demora en la ciudad origen.", vbInformation, "ATENCIÓN"
                End If
                RsAuxE.Close
                Screen.MousePointer = 0
            End If
            'Sumo el importe de los fletes.
            'ConsultoPrecioContenedor
            If tFlete.Text = "" Or embarque.Embarco = DateMinValue Then ValidoAsignarFlete
        End If
        
        tArriboPrevisto.SetFocus
    End If
    Exit Sub
ErrFE:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error inesperado."
End Sub

Private Sub tEmbarco_LostFocus()
    Status.SimpleText = ""
    If IsDate(tEmbarco.Text) Then tEmbarco.Text = Format(tEmbarco.Text, "d-Mmm-yyyy")
End Sub

Private Sub tEmbPrevisto_GotFocus()
    With tEmbPrevisto
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese la fecha prevista del embarque."
End Sub

Private Sub tEmbPrevisto_KeyPress(KeyAscii As Integer)
On Error GoTo ErrFE
    If KeyAscii = vbKeyReturn Then
        If IsDate(tEmbPrevisto.Text) And cOrigen.ListIndex <> -1 Then
            Screen.MousePointer = 11
            Cons = "Select CiuDemora From Ciudad" _
                & " Where CiuCodigo = " & cOrigen.ItemData(cOrigen.ListIndex)
            Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAuxE.EOF Then
                If Not IsNull(RsAuxE!CiuDemora) Then tArriboPrevisto.Text = Format(SumoDias(tEmbPrevisto.Text, RsAuxE!CiuDemora), FormatoFP) Else MsgBox "No hay datos de demora en la ciudad origen.", vbInformation, "ATENCIÓN"
            End If
            RsAuxE.Close
            Screen.MousePointer = 0
        End If
        If sNuevo Then tArriboPrevisto.SetFocus Else tEmbarco.SetFocus
    End If
    Exit Sub
ErrFE:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error inesperado."
End Sub
Private Sub tEmbPrevisto_LostFocus()
    Status.SimpleText = ""
    If IsDate(tEmbPrevisto.Text) Then tEmbPrevisto.Text = Format(tEmbPrevisto.Text, "d-Mmm-yyyy")
End Sub

Private Sub tFApertura_GotFocus()
    With tFApertura
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese la fecha de apertura de la carpeta. [F1] - Búsqueda"
End Sub

Private Sub tFApertura_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrFA
    
    If KeyCode = vbKeyF1 Then
        If IsDate(tFApertura.Text) Then
            If Trim(tFApertura.Text) = "" Then
                Cons = "Select CarID, 'Código' = CarCodigo, Apertura = CarFApertura, Proveedor = PExNombre From Carpeta, ProveedorExterior " _
                        & " Where CarProveedor = PExCodigo" _
                        & " Order by CarFApertura"
            Else
                If IsDate(tFApertura.Text) Then
                    Cons = "Select CarID, 'Código' = CarCodigo, Apertura = CarFApertura, Proveedor = PExNombre From Carpeta, ProveedorExterior " _
                        & " Where CarFApertura >= '" & Format(tFApertura.Text, "mm/dd/yy") & "'" _
                        & " And CarProveedor = PExCodigo" _
                        & " Order by CarFApertura"
                End If
            End If
            Screen.MousePointer = 11
            Dim objAyuda As New clsListadeAyuda
            If objAyuda.ActivarAyuda(cBase, Cons, 5500, 1, "Lista de Carpetas") Then
                BuscoCarpeta objAyuda.RetornoDatoSeleccionado(0)
            End If
            Set objAyuda = Nothing
'            Ayuda.ActivoListaAyuda Cons, False, miconexion.TextoConexion(logImportaciones), 5500
'            Screen.MousePointer = 11
'            DoEvents
'            If Ayuda.ValorSeleccionado > 0 Then BuscoCarpeta Ayuda.ValorSeleccionado
'            Set Ayuda = Nothing
            Screen.MousePointer = 0
        Else
            tFApertura.Text = ""
        End If
    ElseIf KeyCode = vbKeyReturn Then
        cProveedor.SetFocus
    End If
    Exit Sub
ErrFA:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error inesperado."
End Sub

Private Sub tFApertura_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub tFlete_GotFocus()
    With tFlete
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese el valor del flete."
End Sub

Private Sub tFlete_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        InvocoAsignacionFlete
    End If
End Sub

Private Sub tFlete_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And chFletePago.Enabled Then chFletePago.SetFocus
End Sub

Private Sub tFlete_LostFocus()
    If IsNumeric(tFlete.Text) Then tFlete.Text = Format(tFlete.Text, FormatoMonedaP) Else tFlete.Text = ""
    Status.SimpleText = ""
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    fnc_DoTestLogin
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
        
    Select Case Button.Key
        
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
        Case "carpeta": AccionFormularioCarpeta
        Case "limpiar": AccionLimpiar
        Case "primero": AccionPrimerRegistro
        Case "anterior": AccionRegistroAnterior
        Case "siguiente": AccionRegistroSiguiente
        Case "ultimo": AccionUltimoRegistro
        Case "transporte": EjecutarApp App.Path & "\Mantenimiento de Transportes.exe"
        Case "flete": EjecutarApp App.Path & "\Flete_Embarque.exe"
        Case "refresh": AccionRefrescoCombos
        Case "linea": EjecutarApp App.Path & "\lineas_flete.exe"
    End Select


End Sub

Private Sub tUltFechaEmbarque_GotFocus()
    With tUltFechaEmbarque
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese la última fecha posible de embarque."
End Sub

Private Sub tUltFechaEmbarque_KeyPress(KeyAscii As Integer)
    If vbKeyReturn = KeyAscii Then
        If vsArticulo.Enabled Then vsArticulo.Select 1, 1, 1, 1: vsArticulo.SetFocus
    End If
End Sub

Private Sub tUltFechaEmbarque_LostFocus()
    Status.SimpleText = ""
    If IsDate(tUltFechaEmbarque.Text) Then tUltFechaEmbarque.Text = Format(tUltFechaEmbarque.Text, "d-Mmm-yyyy") Else tUltFechaEmbarque.Text = ""
End Sub

'--------------------------------------------------------------------------------------------------------
'   Se buscan los ultimos datos de un embarque anterior para desplegar en pantalla
'   Pedido de Carlos el 25/9/98
'--------------------------------------------------------------------------------------------------------
Private Sub BuscoUltimosDatos()

    On Error GoTo errCargar
    If cProveedor.ListIndex <> -1 Then
        Screen.MousePointer = 11
        
        Cons = " Select * from Carpeta, Embarque" _
                & " Where CarID = EmbCarpeta " _
                & " And CarProveedor = " & cProveedor.ItemData(cProveedor.ListIndex) _
                & " And CarFApertura = (Select Max(CarFApertura) From Carpeta Where CarProveedor = " & cProveedor.ItemData(cProveedor.ListIndex) _
                                & "And CarID <> " & tCodigo.Tag & ")"
        
        Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic)
    
        If Not RsAuxE.EOF Then
            'Cargo la moneda.
            BuscoCodigoEnCombo cMoneda, RsAuxE!EmbMoneda
            If cMoneda.ListIndex > -1 Then
                Cons = "Select * From Moneda Where MonCodigo = " & cMoneda.ItemData(cMoneda.ListIndex)
                Set RsAuxCar = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAuxCar.EOF Then
                    If RsAuxCar!MonArbitraje Then tArbitraje.Enabled = True: tArbitraje.BackColor = Obligatorio Else tArbitraje.Enabled = False: tArbitraje.BackColor = Inactivo: tArbitraje.Text = ""
                End If
                RsAuxCar.Close
            End If
            
            'Cargo Ciudad Destino
            If Not IsNull(RsAuxE!EmbCiudadDestino) Then BuscoCodigoEnCombo cDestino, RsAuxE!EmbCiudadDestino
            
            'Cargo Ciudad Origen
            If Not IsNull(RsAuxE!EmbCiudadOrigen) Then BuscoCodigoEnCombo cOrigen, RsAuxE!EmbCiudadOrigen
                
            '12/11/2013 Matilde pide que no sugiera más la agencia.
            'Cargo Agencia
            'If Not IsNull(RsAuxE!EmbAgencia) Then BuscoCodigoEnCombo cAgencia, RsAuxE!EmbAgencia
            
        End If
        RsAuxE.Close
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ha ocurrido un error al cargar los datos del último embarque del proveedor."
End Sub
Private Function BuscoCosto(Articulo As Long) As Currency
On Error GoTo ErrBC

    RelojA
    Cons = "Select PCoImporte From PrecioDeCosto" _
        & " Where PCoArticulo = " & Articulo _
        & " And PCoFecha = " _
            & "(Select MAX(PCoFecha) From PrecioDeCosto" _
            & " Where PCoArticulo = " & Articulo & ")"

    Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic)
    
    If RsAuxE.EOF Then
        BuscoCosto = 0
    Else
        BuscoCosto = RsAuxE!PCoImporte
    End If
    RsAuxE.Close
    RelojD
    Exit Function
    
ErrBC:
    RelojD
    clsGeneral.OcurrioError "Error al buscar el precio del artículo."
End Function
Private Sub InicioResultSet(EmbarqueID As Long)
On Error GoTo ErrIR
    LimpioCamposCarpeta
    LimpioCamposEmbarque
    HabilitoCamposCarpeta
    InhabilitoCamposEmbarque
    Cons = "Select * From Embarque Where EmbID = " & EmbarqueID
    Set RsEmbarque = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsEmbarque.EOF Then
        'La carpeta puede tener más embarques.------------------------------
        Cons = "Select * From Embarque Where EmbCarpeta = " & RsEmbarque!Embcarpeta
        RsEmbarque.Close
        Set RsEmbarque = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        CargoDatosCarpetaPorID RsEmbarque!Embcarpeta
        CargoDatosEmbarque
        CargoArticulosEmbarque RsEmbarque!EmbID
    End If
    DetalloRegistroEmbarque
    Exit Sub
ErrIR:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio error al iniciar el set de embarques."
End Sub
Private Sub LimpioCamposCarpeta()
    tCodigo.Text = "": tCodigo.Tag = "": Label2.Tag = ""
    tFApertura.Text = ""
    cProveedor.Text = ""
    labBcoEmisor.Caption = ""
    labBcoCorresponsal.Caption = ""
    labFactura.Caption = ""
    labLC.Caption = ""
    labIncoterm.Caption = ""
    labCosteada.Caption = ""
    labPlazo.Caption = ""
    labFormaPago.Caption = ""
    labComentario.Caption = ""
    lAnulada.Caption = ""
End Sub
Private Sub LimpioCamposEmbarque()
    tConocimiento.Text = ""
    cOrigen.Text = ""
    cAgencia.Text = ""
    cDestino.Text = ""
    cMTransporte.Text = ""
    cTransporte.Text = ""
'    tCarpetaDespachante.Text = ""
    cboPrioridad.Text = ""
    tFlete.Text = ""
    lblInfoFlete.ForeColor = vbBlack
    chFletePago.Value = 0
    labEmbCosteado.Caption = "Costeado: "
    tEmbPrevisto.Text = ""
    tEmbarco.Text = ""
    tArriboPrevisto.Text = ""
    labArribo.Caption = ""
    tUltFechaEmbarque.Text = ""
    cMoneda.Text = ""
    tDivisa.Text = ""
    tArbitraje.Text = ""
    labDivisaPaga.Caption = ""
    labDivisaPaga.FontUnderline = False
    cLocal.Text = ""
    labFArriboLocal.Caption = ""
    tComentario.Text = ""
    vsContenedor.Rows = 2
    lblGastoMVD.Caption = ""
    LimpioGrilla
End Sub
Private Sub InhabilitoCamposCarpeta()
    tCodigo.Enabled = False: tCodigo.BackColor = Inactivo
    tFApertura.Enabled = False: tFApertura.BackColor = Inactivo
    cProveedor.Enabled = False: cProveedor.BackColor = Inactivo
End Sub
Private Sub HabilitoCamposCarpeta()
    tCodigo.Enabled = True: tCodigo.BackColor = Obligatorio
    tFApertura.Enabled = True: tFApertura.BackColor = vbWhite
    cProveedor.Enabled = True: cProveedor.BackColor = vbWhite
End Sub
Private Sub InhabilitoCamposEmbarque()
    tConocimiento.Enabled = False: tConocimiento.BackColor = Inactivo
    cOrigen.Enabled = False: cOrigen.BackColor = Inactivo
    cAgencia.Enabled = False: cAgencia.BackColor = Inactivo
    cDestino.Enabled = False: cDestino.BackColor = Inactivo
    cMTransporte.Enabled = False: cMTransporte.BackColor = Inactivo
    cTransporte.Enabled = False: cTransporte.BackColor = Inactivo
'    tCarpetaDespachante.Enabled = False: tCarpetaDespachante.BackColor = Inactivo
    cboPrioridad.Enabled = False: cboPrioridad.BackColor = Inactivo
    tFlete.Enabled = False: tFlete.BackColor = Inactivo
    chFletePago.Enabled = False
    vsContenedor.Editable = False: vsContenedor.BackColor = Inactivo
    tEmbPrevisto.Enabled = False: tEmbPrevisto.BackColor = Inactivo
    tEmbarco.Enabled = False: tEmbarco.BackColor = Inactivo
    tArriboPrevisto.Enabled = False: tArriboPrevisto.BackColor = Inactivo
    tUltFechaEmbarque.Enabled = False: tUltFechaEmbarque.BackColor = Inactivo
    cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    tDivisa.Enabled = False: tDivisa.BackColor = Inactivo
    tArbitraje.Enabled = False: tArbitraje.BackColor = Inactivo
    cLocal.Enabled = False: cLocal.BackColor = Inactivo
    tComentario.Enabled = False: tComentario.BackColor = Inactivo
    vsArticulo.BackColor = Inactivo
    vsArticulo.Enabled = True
End Sub
Private Sub HabilitoCamposEmbarque()
    tConocimiento.Enabled = True: tConocimiento.BackColor = vbWhite
    cOrigen.Enabled = True: cOrigen.BackColor = vbWhite
    cAgencia.Enabled = True: cAgencia.BackColor = vbWhite
    cDestino.Enabled = True: cDestino.BackColor = vbWhite
    cMTransporte.Enabled = True: cMTransporte.BackColor = vbWhite
    cTransporte.Enabled = True: cTransporte.BackColor = vbWhite
    cboPrioridad.Enabled = True: cboPrioridad.BackColor = vbWhite
    tFlete.Enabled = True: tFlete.BackColor = vbWhite
    chFletePago.Enabled = True
    tEmbPrevisto.Enabled = True: tEmbPrevisto.BackColor = Obligatorio
    tEmbarco.Enabled = True: tEmbarco.BackColor = vbWhite
    tArriboPrevisto.Enabled = True: tArriboPrevisto.BackColor = vbWhite
    tUltFechaEmbarque.Enabled = True: tUltFechaEmbarque.BackColor = vbWhite
    cMoneda.Enabled = True: cMoneda.BackColor = Obligatorio
    tDivisa.Enabled = True: tDivisa.BackColor = Obligatorio
    tArbitraje.Enabled = True: tArbitraje.BackColor = Obligatorio
    cLocal.Enabled = True: cLocal.BackColor = vbWhite
    tComentario.Enabled = True: tComentario.BackColor = vbWhite
    vsArticulo.BackColor = vbWhite
    vsContenedor.Editable = True: vsContenedor.BackColor = vbWhite
End Sub
Private Sub CargoDatosCarpetaPorCodigo(Codigo As Long)
On Error GoTo ErrCDC
    Screen.MousePointer = 11
    Cons = "Select * From Carpeta Where CarCodigo = " & Codigo
    Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAuxE.EOF Then
        MsgBox "La carpeta seleccionada no existe, verifique.", vbInformation, "ATENCIÓN"
        tCodigo.Tag = ""
        Label2.Tag = ""
    Else
        tCodigo.Tag = RsAuxE!CarID: tCodigo.Text = RsAuxE!CarCodigo
        Label2.Tag = RsAuxE!CarCodigo
        If Not IsNull(RsAuxE!CarFAnulada) Then lAnulada.Caption = "Anulada: " & Format(RsAuxE!CarFAnulada, "dd/mm/yy hh:mm") Else lAnulada.Caption = ""
        tFApertura.Text = Format(RsAuxE!CarFApertura, FormatoFP)
        If Not IsNull(RsAuxE!CarProveedor) Then BuscoCodigoEnCombo cProveedor, RsAuxE!CarProveedor
        If Not IsNull(RsAuxE!CarBcoEmisor) Then
            Cons = "Select * From BancoLocal Where BLoCodigo = " & RsAuxE!CarBcoEmisor
            Set RsAuxCar = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAuxCar.EOF Then labBcoEmisor.Caption = Trim(RsAuxCar!BLoNombre)
            RsAuxCar.Close
        End If
        If Not IsNull(RsAuxE!CarBcoCorresponsal) Then
            Cons = "Select * From BancoExterior Where BExCodigo = " & RsAuxE!CarBcoCorresponsal
            Set RsAuxCar = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAuxCar.EOF Then labBcoCorresponsal.Caption = Trim(RsAuxCar!BExNombre)
            RsAuxCar.Close
        End If
        If Not IsNull(RsAuxE!CarFactura) Then labFactura.Caption = Trim(RsAuxE!CarFactura)
        If Not IsNull(RsAuxE!CarCartaCredito) Then labLC.Caption = Trim(RsAuxE!CarCartaCredito)
        If Not IsNull(RsAuxE!CarFormaPago) Then labFormaPago.Caption = RetornoFormaPago(RsAuxE!CarFormaPago)
        If Not IsNull(RsAuxE!CarPlazo) Then labPlazo.Caption = RsAuxE!CarPlazo
        If Not IsNull(RsAuxE!CarIncoterm) Then
            Cons = "Select IncNombre From IncoTerm Where IncCodigo = " & RsAuxE!CarIncoterm
            Set RsAuxCar = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAuxCar.EOF Then
                If Not IsNull(RsAuxCar(0)) Then labIncoterm.Caption = Trim(RsAuxCar(0))
            End If
            RsAuxCar.Close
        End If
        If RsAuxE!CarCosteada Then labCosteada.Caption = "SI" Else labCosteada.Caption = "NO"
        If Not IsNull(RsAuxE!CarComentario) Then labComentario.Caption = Trim(RsAuxE!CarComentario)
    End If
    RsAuxE.Close
    Screen.MousePointer = 0
    Exit Sub
ErrCDC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar los datos de la carpeta."
End Sub
Private Sub CargoDatosCarpetaPorID(ID As Long)
On Error GoTo ErrCDC
    Screen.MousePointer = 11
    Cons = "Select * From Carpeta Where CarID = " & ID
    Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAuxE.EOF Then
        MsgBox "La carpeta seleccionada no existe, verifique.", vbInformation, "ATENCIÓN"
        tCodigo.Tag = ""
        Label2.Tag = ""
    Else
        tCodigo.Tag = RsAuxE!CarID: tCodigo.Text = RsAuxE!CarCodigo
        Label2.Tag = RsAuxE!CarCodigo
        tFApertura.Text = Format(RsAuxE!CarFApertura, FormatoFP)
        If Not IsNull(RsAuxE!CarBcoEmisor) Then
            Cons = "Select * From BancoLocal Where BLoCodigo = " & RsAuxE!CarBcoEmisor
            Set RsAuxCar = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAuxCar.EOF Then labBcoEmisor.Caption = Trim(RsAuxCar!BLoNombre)
            RsAuxCar.Close
        End If
        If Not IsNull(RsAuxE!CarBcoCorresponsal) Then
            Cons = "Select * From BancoExterior Where BExCodigo = " & RsAuxE!CarBcoCorresponsal
            Set RsAuxCar = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAuxCar.EOF Then labBcoCorresponsal.Caption = Trim(RsAuxCar!BExNombre)
            RsAuxCar.Close
        End If
        If Not IsNull(RsAuxE!CarProveedor) Then BuscoCodigoEnCombo cProveedor, RsAuxE!CarProveedor
        If Not IsNull(RsAuxE!CarFactura) Then labFactura.Caption = Trim(RsAuxE!CarFactura)
        If Not IsNull(RsAuxE!CarCartaCredito) Then labLC.Caption = Trim(RsAuxE!CarCartaCredito)
        If Not IsNull(RsAuxE!CarFormaPago) Then labFormaPago.Caption = RetornoFormaPago(RsAuxE!CarFormaPago)
        If Not IsNull(RsAuxE!CarPlazo) Then labPlazo.Caption = RsAuxE!CarPlazo
        If Not IsNull(RsAuxE!CarIncoterm) Then
            Cons = "Select IncNombre From IncoTerm Where IncCodigo = " & RsAuxE!CarIncoterm
            Set RsAuxCar = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAuxCar.EOF Then
                If Not IsNull(RsAuxCar(0)) Then labIncoterm.Caption = Trim(RsAuxCar(0))
            End If
            RsAuxCar.Close
        End If
        If RsAuxE!CarCosteada Then labCosteada.Caption = "SI" Else labCosteada.Caption = "NO"
        If Not IsNull(RsAuxE!CarComentario) Then labComentario.Caption = Trim(Trim(RsAuxE!CarComentario))
    End If
    RsAuxE.Close
    Screen.MousePointer = 0
    Exit Sub
ErrCDC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar los datos de la carpeta."
End Sub
Private Sub CargoDatosEmbarque()
On Error GoTo ErrCDE
    Screen.MousePointer = 11
    Set embarque = New clsEmbarque
    embarque.ID = RsEmbarque("EmbID")
    'Pongo el código de Embarque.
    tCodigo.Text = Label2.Tag & "." & RsEmbarque!EmbCodigo
    If Not IsNull(RsEmbarque!EmbBL) Then tConocimiento.Text = Trim(RsEmbarque!EmbBL)
    If Not IsNull(RsEmbarque!EmbCiudadOrigen) Then BuscoCodigoEnCombo cOrigen, RsEmbarque!EmbCiudadOrigen
    If Not IsNull(RsEmbarque!EmbAgencia) Then
        BuscoCodigoEnCombo cAgencia, RsEmbarque!EmbAgencia
        embarque.Agencia = RsEmbarque("EmbAgencia")
    End If
    If Not IsNull(RsEmbarque!EmbMedioTransporte) Then BuscoCodigoEnCombo cMTransporte, RsEmbarque!EmbMedioTransporte
    If Not IsNull(RsEmbarque!EmbTransporte) Then BuscoCodigoEnCombo cTransporte, RsEmbarque!EmbTransporte: embarque.Transporte = RsEmbarque("EmbTransporte")
'    If Not IsNull(RsEmbarque!EmbCarpetaDesp) Then tCarpetaDespachante.Text = Trim(RsEmbarque!EmbCarpetaDesp)
    If Not IsNull(RsEmbarque("EmbPrioridad")) Then BuscoCodigoEnCombo cboPrioridad, RsEmbarque("EmbPrioridad")
    If Not IsNull(RsEmbarque!EmbFlete) Then
        tFlete.Text = Format(RsEmbarque!EmbFlete, FormatoMonedaP)
        If IsNull(RsEmbarque!EmbFEmbarque) Then lblInfoFlete.ForeColor = ColorNaranja
    End If
    If RsEmbarque!EmbFletePago Then chFletePago.Value = 1 Else chFletePago.Value = 0
    If Not IsNull(RsEmbarque!EmbCiudadDestino) Then BuscoCodigoEnCombo cDestino, RsEmbarque!EmbCiudadDestino
    If Not IsNull(RsEmbarque!EmbMoneda) Then BuscoCodigoEnCombo cMoneda, RsEmbarque!EmbMoneda
    If Not IsNull(RsEmbarque!EmbDivisa) Then tDivisa.Text = Format(RsEmbarque!EmbDivisa, FormatoMonedaP)
    If Not IsNull(RsEmbarque!EmbArbitraje) Then tArbitraje.Text = RsEmbarque!EmbArbitraje Else tArbitraje.Text = "1.0000000"
    If RsEmbarque!EmbDivisaPaga Then labDivisaPaga.Caption = "SI": labDivisaPaga.FontUnderline = True Else labDivisaPaga.Caption = "NO"
    If Not IsNull(RsEmbarque!EmbLocal) Then BuscoCodigoEnCombo cLocal, RsEmbarque!EmbLocal
    If RsEmbarque!EmbCosteado Then labEmbCosteado.Caption = "Costeado: Si" Else labEmbCosteado.Caption = "Costeado: No"
    If Not IsNull(RsEmbarque!EmbFEPrometido) Then tEmbPrevisto.Text = Format(RsEmbarque!EmbFEPrometido, FormatoFP)
    If Not IsNull(RsEmbarque!EmbFEmbarque) Then tEmbarco.Text = Format(RsEmbarque!EmbFEmbarque, FormatoFP): embarque.Embarco = RsEmbarque("EmbFEmbarque")
    If Not IsNull(RsEmbarque!EmbFAPrometido) Then tArriboPrevisto.Text = Format(RsEmbarque!EmbFAPrometido, FormatoFP): embarque.ArriboPrevisto = RsEmbarque("EmbFAPrometido")
    If Not IsNull(RsEmbarque!EmbFArribo) Then labArribo.Caption = Format(RsEmbarque!EmbFArribo, FormatoFP): embarque.Arribo = RsEmbarque("EmbFArribo")
    If Not IsNull(RsEmbarque!EmbFLocal) Then labFArriboLocal.Caption = Format(RsEmbarque!EmbFLocal, FormatoFP)
    If Not IsNull(RsEmbarque!EmbComentario) Then tComentario.Text = Trim(RsEmbarque!EmbComentario)
    If Not IsNull(RsEmbarque!EmbUltFechaEmbarque) Then tUltFechaEmbarque.Text = Format(RsEmbarque!EmbUltFechaEmbarque, FormatoFP)
    'INDICO EL CODIGO EN EL TEXTO.---------
    If RsEmbarque!EmbCodigo <> "A" Then tCodigo.Text = Label2.Tag & "." & Trim(RsEmbarque!EmbCodigo)
    '---------------------------------------------------------
    CargoContenedores
    
    lblGastoMVD.Caption = Format(fnc_ObtengoGastosMVD(RsEmbarque!EmbID), "#,##0.00")
    
    Screen.MousePointer = 0
    Exit Sub
ErrCDE:
    clsGeneral.OcurrioError "Error al cargar la información del embarque."
End Sub
Private Sub BuscoArticuloXCodigo(Articulo As Long, Fila As Long)
On Error GoTo ErrBAXC

    Screen.MousePointer = 0
    'Si es código de fábrica ahí puedo tener más de uno.
    'Para no complicarla hago primero por código y si no tengo busco por nombre o código de fáb.
    
    Cons = "Select ArtID, ArtNombre, AImProveedor From Articulo, ArticuloImportacion Where ArtCodigo = " & Articulo _
        & " And ArtSeImporta = 1 And ArtID = AImArticulo "
    
    Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAuxE.EOF Then
        RsAuxE.Close
        Screen.MousePointer = 0
        
        BuscoArticuloPorNombre vsArticulo.Cell(flexcpText, Fila, 1), Fila
        
        'MsgBox "No existe un artículo con ese código, o el mismo no es de importación.", vbInformation, "ATENCIÓN"
        'vsArticulo.Cell(flexcpText, Fila, 1) = ""
    Else
        If RsAuxE!AImProveedor <> cProveedor.ItemData(cProveedor.ListIndex) Then
            If MsgBox("El artículo no tiene asignado el proveedor." & Chr(13) & "¿Desea ingresarlo de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                vsArticulo.Cell(flexcpText, Fila, 0, Fila, 3) = "": Exit Sub
            End If
        End If
        For I = 1 To vsArticulo.Rows - 1
            If Val(vsArticulo.Cell(flexcpText, I, 0)) = RsAuxE!ArtID And I <> Fila Then
                'Lo pongo en rojo porque quiere porratear.
                vsArticulo.Cell(flexcpForeColor, Fila, 0, Fila, 3) = vbRed
            End If
        Next I
        vsArticulo.Cell(flexcpText, Fila, 0) = RsAuxE!ArtID
        vsArticulo.Cell(flexcpText, Fila, 1) = Trim(RsAuxE!ArtNombre)
        RsAuxE.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBAXC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar el artículo por código."
End Sub
Private Sub BuscoCarpeta(Carpeta As Long, Optional embarque As String = "")
On Error GoTo ErrBC
    Screen.MousePointer = 11
    LimpioCamposCarpeta
    LimpioCamposEmbarque
    CargoDatosCarpetaPorID Carpeta
    If tCodigo.Text <> "" Then
        RsEmbarque.Close
        Cons = "Select * From Embarque Where EmbCarpeta = " & tCodigo.Tag
        Set RsEmbarque = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        DetalloRegistroEmbarque
        If Not RsEmbarque.EOF Then
            If embarque = "" Then
                CargoDatosEmbarque
                CargoArticulosEmbarque RsEmbarque!EmbID
            Else
                'Recorro hasta encontrar el embarque.
                Do While Not RsEmbarque.EOF
                    If UCase(Trim(RsEmbarque!EmbCodigo)) = UCase(Trim(embarque)) Then Exit Do
                    RsEmbarque.MoveNext
                Loop
                If RsEmbarque.EOF Then MsgBox "No existe el embarque seleccionado, se despliega el primer embarque de esa carpeta.", vbInformation, "ATENCIÓN": RsEmbarque.MoveFirst
                IndicoRegistro
                CargoDatosEmbarque
                CargoArticulosEmbarque RsEmbarque!EmbID
            End If
        End If
    Else
        RsEmbarque.Close
        Cons = "Select * From Embarque Where EmbCarpeta = 0"
        Set RsEmbarque = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        DetalloRegistroEmbarque
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar la carpeta."
End Sub

Private Function fnc_ObtengoGastosMVD(ByVal idEmbarque As Long) As Currency
Dim RsF As rdoResultset
    
    'Hable con Matilde y me confirmo que tome el 1ero ya que sólo tendría q haber uno.
    On Error GoTo errODL
    Cons = "SELECT FEmGastosMvd FROM FleteEmbarque, Embarque, EmbarqueContenedor " & _
        "WHERE EmbID = " & idEmbarque & _
        " AND FEmOrigen = EmbCiudadOrigen " & _
        "AND FEmDestino = EmbCiudadDestino " & _
        "AND FEmAgencia = EmbAgencia " & _
        "AND FEmFAPartir <= EmbFEmbarque " & _
        "AND FEmFHasta >= EmbFEmbarque " & _
        "AND FEmContenedor = ECoContenedor " & _
        "AND FEmLinea = ECoLinea " & _
        "AND EmbID = ECoEmbarque"
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsF.EOF Then fnc_ObtengoGastosMVD = RsF(0)
    RsF.Close
    Exit Function
errODL:
    
End Function
Private Sub BuscoCarpetaPorCodigo(Carpeta As Long, Optional embarque As String = "")
On Error GoTo ErrBC
    
    Screen.MousePointer = 11
    LimpioCamposCarpeta
    LimpioCamposEmbarque
    CargoDatosCarpetaPorCodigo Carpeta
    If tCodigo.Text <> "" Then
        RsEmbarque.Close
        Cons = "Select * From Embarque Where EmbCarpeta = " & tCodigo.Tag
        Set RsEmbarque = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsEmbarque.EOF Then
            DetalloRegistroEmbarque
            If embarque = "" Then
                CargoDatosEmbarque
                CargoArticulosEmbarque RsEmbarque!EmbID
            Else
                'Recorro hasta encontrar el embarque.
                Do While Not RsEmbarque.EOF
                    If UCase(Trim(RsEmbarque!EmbCodigo)) = UCase(Trim(embarque)) Then Exit Do
                    RsEmbarque.MoveNext
                Loop
                If RsEmbarque.EOF Then MsgBox "No existe el embarque " & UCase(embarque) & " en dicha carpeta.", vbInformation, "ATENCIÓN": RsEmbarque.MoveFirst
                IndicoRegistro
                CargoDatosEmbarque
                CargoArticulosEmbarque RsEmbarque!EmbID
            End If
        Else
            'No posee embarques automáticamente opto por nuevo.
            PreparoAccionNuevo
            CargoDatosBoquilla Val(tCodigo.Tag)
            BuscoUltimosDatos
            iSeleccionado = 0
            tConocimiento.SetFocus
            '-----------------------------------------------------------------
        End If
    Else
        RsEmbarque.Close
        Cons = "Select * From Embarque Where EmbCarpeta = 0"
        Set RsEmbarque = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        DetalloRegistroEmbarque
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrBC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar la carpeta.", Trim(Err.Description)
End Sub
Private Sub DetalloRegistroEmbarque()
On Error GoTo ErrDRE
    Screen.MousePointer = 11
    MnuCarpeta.Enabled = True: Toolbar1.Buttons("carpeta").Enabled = True
    If Not RsEmbarque.EOF Then
        If lAnulada.Caption = "" Then
            Botones True, True, True, False, False, Toolbar1, Me
        Else
            Botones True, False, False, False, False, Toolbar1, Me
        End If
        MnuEditMemo.Enabled = (lAnulada.Caption = "")
        RsEmbarque.MoveLast
        tRegistro.Text = "1 de " & Trim(RsEmbarque.RowCount)
        If RsEmbarque.RowCount > 1 Then BotonesRegistro True, True, True, True, Toolbar1, Me Else BotonesRegistro False, False, False, False, Toolbar1, Me
        RsEmbarque.MoveFirst
    Else
        Botones True, False, False, False, False, Toolbar1, Me
        BotonesRegistro False, False, False, False, Toolbar1, Me
        MnuEditMemo.Enabled = False
        tRegistro.Text = ""
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrDRE:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al contabilizar los envíos.", Trim(Err.Description)
End Sub
Private Sub IndicoRegistro()
On Error GoTo ErrIR
    tRegistro.Text = RsEmbarque.AbsolutePosition & " " & Mid(tRegistro.Text, InStr(tRegistro.Text, "d"), Len(tRegistro.Text))
    Exit Sub
ErrIR:
    clsGeneral.OcurrioError "Error al indicar el nro. de registro.", Trim(Err.Description)
End Sub

Private Function ValidoCampos() As Boolean
On Error GoTo ErrVC
    Screen.MousePointer = 11
    'FECHAS.--------------------------------------------------------------------------------
    If tArriboPrevisto.Text <> "" And Not IsDate(tArriboPrevisto.Text) Then
        MsgBox "La fecha de arribo previsto no es correcta.", vbExclamation, "ATENCION"
        tArriboPrevisto.SetFocus
        ValidoCampos = False
        Screen.MousePointer = 0
        Exit Function
    End If
    If tEmbPrevisto.Text = "" Or Not IsDate(tEmbPrevisto.Text) Then
        MsgBox "La fecha prevista  para el embarque no es correcta.", vbExclamation, "ATENCION"
        tEmbPrevisto.SetFocus: ValidoCampos = False: Screen.MousePointer = 0: Exit Function
    End If
    If tEmbarco.Text <> "" And Not IsDate(tEmbarco.Text) Then
        MsgBox "La fecha que embarcó no es correcta.", vbExclamation, "ATENCION"
        tEmbarco.SetFocus: ValidoCampos = False: Screen.MousePointer = 0: Exit Function
    End If
    If tUltFechaEmbarque.Text <> "" And Not IsDate(tUltFechaEmbarque.Text) Then
        MsgBox "La última fecha de embarque no es correcta.", vbExclamation, "ATENCION"
        tUltFechaEmbarque.SetFocus: ValidoCampos = False: Screen.MousePointer = 0: Exit Function
    End If
    '.--------------------------------------------------------------------------------
    ValidoCampos = False
    'Cantidad de Contenedores.--------------------------------------------
    Dim lPos1 As Long, lPos As Long
    For lPos = 1 To vsContenedor.Rows - 1
        For lPos1 = lPos + 1 To vsContenedor.Rows - 1
            If vsContenedor.Cell(flexcpText, lPos, 1) = vsContenedor.Cell(flexcpText, lPos1, 1) Then
                MsgBox "El tipo de contenedor " & vsContenedor.Cell(flexcpTextDisplay, lPos, 1) & " esta duplicado.", vbInformation, "ATENCIÓN"
                vsContenedor.SetFocus
                Screen.MousePointer = 0: Exit Function
            End If
        Next lPos1
    Next
    '.--------------------------------------------------------------------------------
    ValidoCampos = False
'    'Nro. de Carpeta del despachante.-------------------------------------
'    If Trim(tCarpetaDespachante.Text) <> "" And Not IsNumeric(tCarpetaDespachante.Text) Then
'        MsgBox "El formato de la carpeta del despachante no es numérica.", vbExclamation, "ATENCION"
'        tCarpetaDespachante.SetFocus:  Screen.MousePointer = 0: Exit Function
'    End If
    '.--------------------------------------------------------------------------------
    'Prioridad.
    If cboPrioridad.ListIndex = -1 Then
        MsgBox "Seleccione una prioridad.", vbExclamation, "ATENCIÓN"
        cboPrioridad.SetFocus: Screen.MousePointer = 0: Exit Function
    End If
    
    'Valor del flete.---------------------------------------------------------------
    If Trim(tFlete.Text) <> "" And Not IsNumeric(tFlete.Text) Then
        MsgBox "El formato del costo del flete no es numérico.", vbExclamation, "ATENCION"
        tFlete.SetFocus:  Screen.MousePointer = 0: Exit Function
    End If
    'Flete.---------------------------------------------------------------------------
    If chFletePago.Value And Trim(tFlete.Text) = "" Then
        MsgBox "Se debe ingresar el gasto del flete, si el mismo está pago.", vbExclamation, "ATENCION"
        tFlete.SetFocus: Screen.MousePointer = 0: Exit Function
    End If
    '.--------------------------------------------------------------------------------
    'Chequeo datos divisa.------------------------------------------------------
    If cMoneda.ListIndex = -1 Or Trim(tDivisa.Text) = "" Then
        MsgBox "Los datos ingresados para la divisa no son correctos.", vbExclamation, "ATENCION"
        cMoneda.SetFocus: Screen.MousePointer = 0: Exit Function
    End If
    If Not IsNumeric(tDivisa.Text) Then
        MsgBox "Los datos ingresados para la divisa no son correctos.", vbExclamation, "ATENCION"
        tDivisa.SetFocus: ValidoCampos = False: Screen.MousePointer = 0: Exit Function
    End If
    If tArbitraje.Enabled And (Trim(tArbitraje.Text) = "" Or Not IsNumeric(tArbitraje.Text)) Then
        MsgBox "Los datos ingresados para el arbitraje no son correctos.", vbExclamation, "ATENCION"
        tArbitraje.SetFocus: ValidoCampos = False: Screen.MousePointer = 0: Exit Function
    End If
    'Fecha de arribo sea mayor a la de embarque.-----------------------------
    If IsDate(tArriboPrevisto.Text) And IsDate(tEmbPrevisto.Text) Then
        If CDate(tArriboPrevisto.Text) < CDate(tEmbPrevisto.Text) Then
            MsgBox "La fecha de arribo previsto es menor a la fecha de embarque previsto.", vbExclamation, "ATENCION"
            tArriboPrevisto.SetFocus: ValidoCampos = False: Screen.MousePointer = 0: Exit Function
        End If
    End If
    'Fecha últ. de embarque contra emb. previsto.-----------------------------
    If IsDate(tUltFechaEmbarque.Text) And IsDate(tEmbPrevisto.Text) Then
        If CDate(tUltFechaEmbarque.Text) < CDate(tEmbPrevisto.Text) Then
            If MsgBox("La última fecha de embarque es menor a la fecha de embarque previsto." & Chr(vbKeyReturn) & "Desea continuar.", vbQuestion + vbYesNo + vbDefaultButton2, "ATENCION") = vbNo Then
                tUltFechaEmbarque.SetFocus: ValidoCampos = False: Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    'Controlo que por lo menos tenga un artículo.------------------------------
    If vsArticulo.Rows < 2 Then
        MsgBox "No se ingresaron artículos.", vbExclamation, "ATENCIÓN"
        vsArticulo.Select 1, 1, 1, 1: vsArticulo.SetFocus: ValidoCampos = False: Screen.MousePointer = 0: Exit Function
    End If
    ValidoCampos = True
    Screen.MousePointer = 0
    Exit Function
ErrVC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al validar los campos."
End Function

Private Sub NuevoEmbarque()
Dim idEmbarque As Long
Dim LetraEmb As String * 1

    On Error GoTo ErrEmb
    'Veo la forma de pago de la carpeta.
    Cons = "Select * From Carpeta Where CarID = " & tCodigo.Tag
    Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAuxE!CarFormaPago = FormaPago.cFPAnticipado Or RsAuxE!CarFormaPago = FormaPago.cFPCobranza Then
        MsgBox "Se va a registrar un gasto automáticamente con el valor de la divisa asignado al proveedor N/D.", vbInformation, "ATENCIÓN"
    Else
        If Not IsNull(RsAuxE!CarBcoEmisor) And Not IsNull(RsAuxE!CarCartaCredito) Then
            MsgBox "Se va a registrar un gasto automáticamente con el valor de la divisa asignado al Banco Emisor.", vbInformation, "ATENCIÓN"
        End If
    End If
    RsAuxE.Close
    
    Screen.MousePointer = 0
    If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    
    'COMIENZO TRANSACCION.------------------------------------------------------
    cBase.BeginTrans
    On Error GoTo ErrResumo
    
    'Saco el mayor código de Embarque que hay para la carpeta.--------------
    Cons = "Select Max(EmbCodigo) From Embarque Where EmbCarpeta = " & tCodigo.Tag
    Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAuxE.EOF Then
        If Not IsNull(RsAuxE(0)) Then LetraEmb = RsAuxE(0) Else LetraEmb = vbNullString
    Else
        LetraEmb = vbNullString
    End If
    RsAuxE.Close
    If Trim(LetraEmb) = "" Then LetraEmb = "A" Else LetraEmb = Chr(Asc(LetraEmb) + 1)
    
    RsEmbarque.Close
    Cons = "Select * From Embarque Where EmbCarpeta = " & tCodigo.Tag
    Set RsEmbarque = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'INSERTO EL NUEVO EMBARQUE.----------------------------------------------
    RsEmbarque.AddNew
    InsertoCamposBD LetraEmb
    RsEmbarque.Update
    
    'Saco el ID del insertado.-------------------------------------------------------------
    Cons = "Select Max(EmbID) From Embarque"
    Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    idEmbarque = RsAuxE(0)
    RsAuxE.Close
    'Guardo los artículos del embarque.------------------------------------------------
    GuardoArticulosEmbarque idEmbarque
    
    GuardoContenedores idEmbarque
    
    Dim dFechaGasto As Date, bGasto As Boolean
    Dim iBcoEm As Long, iProv As Integer
    Dim sCarpeta As String, sSerie As String
    
    bGasto = False
    'Ingreso gasto automáticamente, si la forma de pago no es plazobl o vista.
    'Veo la forma de pago de la carpeta.
    Cons = "Select * From Carpeta Where CarID = " & tCodigo.Tag
    Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    
    If Not IsNull(RsAuxE!CarBcoEmisor) Then iBcoEm = RsAuxE("CarBcoEmisor") Else iBcoEm = 0
    sCarpeta = RsAuxE!CarCodigo & "." & LetraEmb
    
    If RsAuxE!CarFormaPago = FormaPago.cFPAnticipado Or RsAuxE!CarFormaPago = FormaPago.cFPCobranza Then
        If Trim(tArbitraje.Text) = "" Or tArbitraje.Text = "0" Then tArbitraje.Text = 1
        dFechaGasto = RsAuxE!CarFApertura
        sSerie = "C" & LetraEmb & " " & RsAuxE!CarCodigo
        bGasto = True
        'IngresoGastoAutomatico idEmbarque, Format(CCur(tDivisa.Text) / tArbitraje.Text, FormatoMonedaP), RsAuxE!CarFApertura, 0, RsAuxE!CarCodigo & "." & LetraEmb, LetraEmb, "C" & LetraEmb, RsAuxE!CarCodigo, TipoDocumento.CompraCredito, tArbitraje.Text, labBcoEmisor.Caption, labLC.Caption, labFormaPago.Caption, paSubrubroDivisa
    Else
        If Not IsNull(RsAuxE!CarBcoEmisor) And Not IsNull(RsAuxE!CarCartaCredito) Then
            If Trim(tArbitraje.Text) = "" Or tArbitraje.Text = "0" Then tArbitraje.Text = 1
            iProv = iBcoEm
            dFechaGasto = gFechaServidor
            sSerie = "LC " & RsAuxE!CarCartaCredito
            bGasto = True
            'IngresoGastoAutomatico idEmbarque, Format(CCur(tDivisa.Text) / tArbitraje.Text, FormatoMonedaP), Format(gFechaServidor, FormatoFP), RsAuxE!CarBcoEmisor, RsAuxE!CarCodigo & "." & LetraEmb, LetraEmb, "LC", RsAuxE!CarCartaCredito, TipoDocumento.CompraCredito, tArbitraje.Text, labBcoEmisor.Caption, labLC.Caption, labFormaPago.Caption, paSubrubroDivisa
        End If
    End If
    RsAuxE.Close
    
    cBase.CommitTrans
    'FIN TRANSACCION.-------------------------------------------------------------------
    
    'PASO EL GASTO
    If bGasto Then
        ModCarpeta.InsertoGastoImportacionZureo idEmbarque, Format(CCur(tDivisa.Text) / tArbitraje.Text, FormatoMonedaP), dFechaGasto, iProv, sCarpeta, LetraEmb, sSerie, TipoDocumento.CompraCredito, tArbitraje.Text, labBcoEmisor.Caption, labLC.Caption, labFormaPago.Caption, paSubrubroDivisa, iBcoEm
    End If
    
    'Restauro el formulario.----------------------------------------------------------------
    sNuevo = False
    InhabilitoCamposEmbarque
    HabilitoCamposCarpeta
    BuscoCarpeta tCodigo.Tag, LetraEmb
    tCodigo.SetFocus
    Screen.MousePointer = 0
    Exit Sub
    
ErrEmb:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al iniciar la transacción."
    Exit Sub
    
ErrResumo:
    Resume ErrorNuevo
    
ErrorNuevo:
    cBase.RollbackTrans
    RsEmbarque.Requery
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar almacenar la información."
    Exit Sub
End Sub

Private Function fnc_GetImporteGastos(ByVal idEmb As Long, ByRef iImpGasto As Double) As Boolean
Dim rsG As rdoResultset
    Cons = "Select Sum(GimImporte) From GastoImportacion, Compra Where GImIDSubRubro = " & paSubrubroDivisa _
        & " And GImNivelFolder = " & Folder.cFEmbarque _
        & " And GImFolder = " & idEmb _
        & " And GImIDCompra = ComCodigo And ComMoneda = " & paMonedaDolar
    Set rsG = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not IsNull(rsG(0)) Then
        fnc_GetImporteGastos = True
        iImpGasto = rsG(0)
    End If
    rsG.Close
End Function

Private Sub ModificoEmbarque()
Dim sTodos As Boolean, sNuevoEmbarque As Boolean, sDivisaPaga As Boolean
Dim FechaModificado As Date
Dim LetraEmb As String * 1
Dim ViejoID As Long, NuevoID As Long
Dim aValor As Currency, DivisaAnterior As Currency, ArbitrajeAnterior As Double

Dim regNota As tGastoNota
Dim regMov As tGastoMovimiento
Dim regGasto As tGastoZureo

    On Error GoTo ErrEmb
    Screen.MousePointer = 11
    DivisaAnterior = 0
    sTodos = False: sNuevoEmbarque = False
    'Verificamos si hay embarque similares en destino y buque.
    If Not IsNull(RsEmbarque!EmbFAPrometido) Then
        If CDate(RsEmbarque!EmbFAPrometido) <> CDate(tArriboPrevisto.Text) And cDestino.ListIndex <> -1 And cTransporte.ListIndex <> -1 Then
            Cons = "Select * From Embarque" _
                    & " Where (EmbCarpeta <> " & tCodigo.Tag & " OR EmbID <> '" & RsEmbarque!EmbID & "')" _
                    & " And EmbTransporte = " & cTransporte.ItemData(cTransporte.ListIndex) _
                    & " And EmbCiudadDestino = " & cDestino.ItemData(cDestino.ListIndex) _
                    & " And EmbFAPrometido BETWEEN '" & Format(RestoDias(RsEmbarque!EmbFAPrometido, 5), "mm/dd/yyyy") & "' AND '" & Format(SumoDias(RsEmbarque!EmbFAPrometido, 5), "mm/dd/yyyy") & "'" _
                    & " And EmbFArribo IS Null"
            Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAuxE.EOF Then
                If MsgBox("¿Desea modificar la fecha de embarque prometida para todos los embarques con el mismo buque y destino.", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then sTodos = True
            End If
            RsAuxE.Close
        End If
    End If
    
    '----------------------------------------------------------------------------------------
    If vsArticulo.Enabled Then
        'Veo si me modificó las cantidades de artículos.
        Dim IDArticulos As String
        IDArticulos = ""
        For I = 1 To vsArticulo.Rows - 1
            If IDArticulos = "" Then IDArticulos = Val(vsArticulo.Cell(flexcpText, I, 0)) Else IDArticulos = IDArticulos & ", " & Val(vsArticulo.Cell(flexcpText, I, 0))
        Next I
        Cons = "Select * From ArticuloFolder Where AFoTipo = " & Folder.cFEmbarque _
            & " And AFoCodigo = " & RsEmbarque!EmbID _
            & " And AFoArticulo NOT IN (" & IDArticulos & ")"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Close
            If MsgBox("Ha eliminado artículos del embarque, los mismos serán dados de Baja de la carpeta (NO VAN A NUEVO EMBARQUE)." & Chr(13) & "¿Desea continuar?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Screen.MousePointer = 0: Exit Sub
        Else
            RsAux.Close
        End If
        'Veo si modifico las cantidades
        sNuevoEmbarque = DeseaNuevoEmbarque(RsEmbarque!EmbID)
        If sNuevoEmbarque Then
            ' a pedido de Irma hago 2da pregunta.
            If MsgBox("¿Confirma dejar la diferencia a un nuevo embarque?", vbQuestion + vbYesNo, "SEGUNDA PREGUNTA") = vbNo Then
                sNuevoEmbarque = False
            End If
        End If
    Else
        sNuevoEmbarque = False
    End If
    '----------------------------------------------------------------------------------------
    
    If vsArticulo.Enabled And Not IsNull(RsEmbarque!EmbFArribo) And sNuevoEmbarque Then
        MsgBox "No puede asignar los artículos a un nuevo embarque, elimine el arribó y luego realice esta operación.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If Trim(tArbitraje.Text) = "" Or tArbitraje.Text = "0" Then tArbitraje.Text = "1"
    
    If RsEmbarque!EmbDivisaPaga Then sDivisaPaga = True Else sDivisaPaga = False
    
'    If sDivisaPaga And (CCur(tDivisa.Text) <> RsEmbarque!EmbDivisa Or Val(tArbitraje.Text) <> RsEmbarque!EmbArbitraje) Then
'        If Not sNuevoEmbarque Then
'            MsgBox "Atención se registrará en zureo el ajuste del gasto de divisa.", vbInformation, "IMPORTANTE"
'        End If
    If sDivisaPaga And Not sNuevoEmbarque And (CCur(tDivisa.Text) <> RsEmbarque!EmbDivisa Or Val(tArbitraje.Text) <> RsEmbarque!EmbArbitraje) Then
        Screen.MousePointer = 0
        MsgBox "Alteró el valor de la divisa y la misma ya está paga, no puede realizar esta acción.", vbExclamation, "ATENCIÓN"
        Exit Sub
    Else
        If Not sNuevoEmbarque Then
           If Not sDivisaPaga And (CCur(tDivisa.Text) <> RsEmbarque!EmbDivisa Or Val(tArbitraje.Text) <> RsEmbarque!EmbArbitraje) Then
                If SeHaceNotaoCredito Then
                    If Not DifCambioMesAnterior Then
                        MsgBox "Las diferencias de cambio al mes anterior no están generadas." & vbCrLf _
                            & "No se puede modificar la divisa.", vbExclamation, "ATENCIÓN"
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                End If
            End If
        Else
            If SeHaceNotaoCredito Then
                If Not DifCambioMesAnterior Then
                    MsgBox "Las diferencias de cambio al mes anterior no están generadas." & vbCrLf _
                            & "No se puede modificar la divisa.", vbExclamation, "ATENCIÓN"
                        Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        End If
    End If
    
    Screen.MousePointer = 0
    If MsgBox("¿Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    'COMIENZO TRANSACCION.------------------------------------------------------
    cBase.BeginTrans
'    FechaDelServidor
    On Error GoTo ErrResumo
    FechaModificado = RsEmbarque!EmbFModificacion
    
    'Me quedo con la letra del embarque.
    LetraEmb = RsEmbarque!EmbCodigo
    
    Dim iImpGasto As Double
    
    Cons = "Select * From Embarque Where EmbID = " & RsEmbarque!EmbID
    
    RsEmbarque.Close
    
    Set RsEmbarque = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Verifico Multiusuario.-----------------------------------------------------------------
    If RsEmbarque.EOF Then GoTo EliminoEmbarque
    
    If FechaModificado <> RsEmbarque!EmbFModificacion Then GoTo ModificaronEmbarque
    If sDivisaPaga <> RsEmbarque!EmbDivisaPaga Then GoTo ModificaronEmbarque
    
    If embarque.Arribo = DateMinValue And embarque.ArriboPrevisto <> DateMinValue And embarque.ArriboPrevisto <> CDate(tArriboPrevisto.Text) Then
        CambioFechaEnCalendario DateDiff("d", embarque.ArriboPrevisto, CDate(tArriboPrevisto.Text))
    End If
    
    DivisaAnterior = RsEmbarque!EmbDivisa
    ArbitrajeAnterior = RsEmbarque!EmbArbitraje
    
    If Not sNuevoEmbarque Then
        
        'Guardo los artículos del embarque.------------------------------------------------
        If vsArticulo.Enabled Then
            If IsNull(RsEmbarque!EmbFArribo) Then
                GuardoArticulosEmbarque RsEmbarque!EmbID
            Else
                'Si arribó tengo que verificar el stock y si modificó el precio y posee subcarpetas modificar el importe
                'en la misma.
                GuardoArticulosEmbarqueModificados RsEmbarque!EmbID
            End If
        End If
    
        GuardoContenedores RsEmbarque!EmbID
        
        'Modifico el EMBARQUE.----------------------------------------------
        RsEmbarque.Edit
        InsertoCamposBD LetraEmb
        RsEmbarque.Update
        
        Dim iBcoEm As Long
        
        '20120410 A pedido de Matilde dejo editar la divisa estando paga y sin pasarse a un nuevo embarque.
        'lo volvimos atrás con Irma.
        If Not sDivisaPaga Then
            
            If fnc_GetImporteGastos(RsEmbarque("EmbID"), iImpGasto) Then
                
                'Hay gastos ingresados.
                If Trim(tArbitraje.Text) = "" Or tArbitraje.Text = "0" Then tArbitraje.Text = "1"
                
                If CCur(Format(iImpGasto, "#,##0.0000")) > CCur(Format(RsEmbarque!EmbDivisa / RsEmbarque!EmbArbitraje, "#,##0.000000")) Then
                    
                    'Hago Nota
                    aValor = Format(iImpGasto, FormatoMonedaP) - Format(RsEmbarque!EmbDivisa / RsEmbarque!EmbArbitraje, "#,##0.000000")
                    aValor = Format(aValor, FormatoMonedaP)
                    '14-9-2000 el valor era ej. 0.001111 entonces se veia el gasto como cero.
                    
                    '5/7/2000 aparemente se registran notas con importe cero.
                    If aValor > 0 Then
                        With regNota
                            .idEmb = RsEmbarque("EmbID")
                            .Valor = aValor
                            .DivAnterior = DivisaAnterior / ArbitrajeAnterior
                            .DivPaga = sDivisaPaga
                        End With
                        If Not RsEmbarque!EmbCosteado Then
                            'HagoNotaCredito RsEmbarque!EmbID, aValor, DivisaAnterior / ArbitrajeAnterior, sDivisaPaga, 0, paSubrubroDivisa
                            regNota.SRubro = paSubrubroDivisa
                        Else
                            'HagoNotaCredito RsEmbarque!EmbID, aValor, DivisaAnterior / ArbitrajeAnterior, sDivisaPaga, 0, paSubrubroDifCostoImp
                            regNota.SRubro = paSubrubroDifCostoImp
                        End If
                    End If
                    
                ElseIf CCur(Format(iImpGasto, "#,##0.0000000")) < CCur(Format(RsEmbarque!EmbDivisa / RsEmbarque!EmbArbitraje, "#,##0.0000000")) Then
                    
                    aValor = Format(RsEmbarque!EmbDivisa / RsEmbarque!EmbArbitraje, "#,##0.0000000") - CCur(iImpGasto)
                    aValor = Format(aValor, FormatoMonedaP)
                                        
                    If aValor > 0 Then
                        'Consulto la compra para sacar la misma serie , numero y proveedor del documento anterior.
                        Cons = "Select * from Compra, GastoImportacion, Embarque, Carpeta " _
                                & " Where GImIDSubRubro = " & paSubrubroDivisa _
                                & " And GImNivelFolder = " & Folder.cFEmbarque _
                                & " And GImFolder = " & RsEmbarque!EmbID _
                                & " And ComTipoDocumento = " & TipoDocumento.CompraCredito _
                                & " And GImIDCompra = ComCodigo And GImFolder = EmbID And EmbCarpeta = CarID"
                        
                        Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        
                        If Not RsAuxE.EOF Then
                            
                            If Not IsNull(RsAuxE!CarBcoEmisor) Then iBcoEm = RsAuxE("CarBcoEmisor") Else iBcoEm = 0
                        
                            With regGasto
                                .idEmb = RsEmbarque("EmbID")
                                .Arbitraje = RsEmbarque!EmbArbitraje
                                .BcoCodigo = iBcoEm
                                .BcoNombre = labBcoEmisor.Caption
                                .Carpeta = RsAuxE!CarCodigo & "." & RsAuxE!EmbCodigo
                                .Fecha = gFechaServidor
                                If Not IsNull(RsAuxE("ComProveedor")) Then .Proveedor = RsAuxE("ComProveedor")
                                .Valor = aValor
                                
                                .CodEmbarque = RsEmbarque!EmbCodigo
                                If Not IsNull(RsAuxE!ComSerie) Then
                                    .SerieNro = RsAuxE!ComSerie & " " & RsAuxE!ComNumero
                                Else
                                    .SerieNro = RsAuxE!ComNumero
                                End If
                                .TipoDoc = TipoDocumento.CompraCredito
                                .LC = labLC.Caption
                                .FormaPago = labFormaPago.Caption
                            End With
                        
                            If Not RsEmbarque!EmbCosteado Then
                                regGasto.Cuenta = paSubrubroDivisa
                                'IngresoGastoAutomatico RsEmbarque!EmbID, aValor, Format(gFechaServidor, FormatoFP), RsAuxE!ComProveedor, RsAuxE!CarCodigo & "." & RsAuxE!EmbCodigo, RsEmbarque!EmbCodigo, RsAuxE!ComSerie, RsAuxE!ComNumero, TipoDocumento.CompraCredito, RsEmbarque!EmbArbitraje, labBcoEmisor.Caption, labLC.Caption, labFormaPago.Caption, paSubrubroDivisa
                            Else
                                regGasto.Cuenta = paSubrubroDifCostoImp
                                'IngresoGastoAutomatico RsEmbarque!EmbID, aValor, Format(gFechaServidor, FormatoFP), RsAuxE!ComProveedor, RsAuxE!CarCodigo & "." & RsAuxE!EmbCodigo, RsEmbarque!EmbCodigo, RsAuxE!ComSerie, RsAuxE!ComNumero, TipoDocumento.CompraCredito, RsEmbarque!EmbArbitraje, labBcoEmisor.Caption, labLC.Caption, labFormaPago.Caption, paSubrubroDifCostoImp
                            End If
                        Else
                            'Dejo como esta porque lo que me devuelve aca tiene que ser por un contado.
                        End If
                        RsAuxE.Close
                    End If
                End If
                
            Else
            
                
                'Cambio 16-6-2000
                'Pregunto si la carpeta es de las nuevas. FApertura MAyor al 1/4/2000
                Cons = "Select * From Carpeta Where CarID = " & RsEmbarque!Embcarpeta
                Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                
                If Not IsNull(RsAuxE!CarBcoEmisor) Then iBcoEm = RsAuxE("CarBcoEmisor") Else iBcoEm = 0
                
                If (RsAuxE!CarFApertura > CDate("01/04/2000")) And _
                    ((RsAuxE!CarFormaPago = FormaPago.cFPAnticipado Or RsAuxE!CarFormaPago = FormaPago.cFPCobranza) Or (Not IsNull(RsAuxE!CarBcoEmisor) And Not IsNull(RsAuxE!CarCartaCredito))) Then
                    
                    With regGasto
                        .idEmb = RsEmbarque("EmbID")        '18/11/2008
                        
                        .Arbitraje = Val(tArbitraje.Text)
                        .BcoCodigo = iBcoEm
                        .BcoNombre = labBcoEmisor.Caption
                        .Carpeta = RsAuxE!CarCodigo & "." & RsEmbarque!EmbCodigo
                        .CodEmbarque = RsEmbarque!EmbCodigo
                        .Cuenta = paSubrubroDivisa
                        
                        .FormaPago = labFormaPago.Caption
                        .LC = labLC.Caption
                        .SaldoCero = False
                        .TipoDoc = TipoDocumento.CompraCredito
                        .Valor = Format(tDivisa.Text / tArbitraje.Text, "#,##0.0000000")
                        
                    End With
                
                
                    If RsAuxE!CarFormaPago = FormaPago.cFPAnticipado Or RsAuxE!CarFormaPago = FormaPago.cFPCobranza Then
                        If Trim(tArbitraje.Text) = "" Or tArbitraje.Text = "0" Then tArbitraje.Text = 1
                        'IngresoGastoAutomatico RsEmbarque!EmbID, Format(CCur(tDivisa.Text) / CCur(tArbitraje.Text), "#,##0.000"), RsAuxE!CarFApertura, 0, RsAuxE!CarCodigo & "." & RsEmbarque!EmbCodigo, RsEmbarque!EmbCodigo, "C" & RsEmbarque!EmbCodigo, RsAuxE!CarCodigo, TipoDocumento.CompraCredito, Val(tArbitraje.Text), labBcoEmisor.Caption, labLC.Caption, labFormaPago.Caption, paSubrubroDivisa
                        With regGasto
                            .Fecha = RsAuxE!CarFApertura
                            .Proveedor = 0
                            .SerieNro = "C" & RsEmbarque!EmbCodigo & " " & RsAuxE!CarCodigo
                        End With
                    Else
                        If Not IsNull(RsAuxE!CarBcoEmisor) And Not IsNull(RsAuxE!CarCartaCredito) Then
                            If Trim(tArbitraje.Text) = "" Or tArbitraje.Text = "0" Then tArbitraje.Text = 1
                            'IngresoGastoAutomatico RsEmbarque!EmbID, Format(tDivisa.Text / tArbitraje.Text, "#,##0.0000000"), Format(gFechaServidor, FormatoFP), RsAuxE!CarBcoEmisor, RsAuxE!CarCodigo & "." & RsEmbarque!EmbCodigo, RsEmbarque!EmbCodigo, "LC", RsAuxE!CarCartaCredito, TipoDocumento.CompraCredito, tArbitraje.Text, labBcoEmisor.Caption, labLC.Caption, labFormaPago.Caption, paSubrubroDivisa
                            With regGasto
                                .Fecha = gFechaServidor
                                .Proveedor = RsAuxE!CarBcoEmisor
                                .SerieNro = "LC " & RsAuxE!CarCartaCredito
                            End With
                        End If
                    End If
                End If
                RsAuxE.Close
            End If
            
        End If
        
    Else
        
        'Modifico el EMBARQUE.----------------------------------------------
        ViejoID = RsEmbarque!EmbID
        
        RsEmbarque.Edit
        InsertoCamposBD LetraEmb
        RsEmbarque.Update
        
        
        'Busco si hay un embarque que no arribo para darle la mercadería.
        Cons = "Select * From Embarque Where EmbID = " _
            & "(Select Max(EmbID) From Embarque Where EmbCarpeta = " & tCodigo.Tag _
            & " And EmbFEmbarque Is Null And EmbCodigo > '" & LetraEmb & "')"
        Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAuxE.EOF Then
            RsAuxE.Close
            Cons = "Select Max(EmbCodigo) From Embarque Where EmbCarpeta = " & tCodigo.Tag
            Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAuxE.EOF Then
                If Not IsNull(RsAuxE(0)) Then LetraEmb = RsAuxE(0) Else LetraEmb = vbNullString
            Else
                LetraEmb = vbNullString
            End If
            RsAuxE.Close
            If Trim(LetraEmb) = "" Then LetraEmb = "A" Else LetraEmb = Chr(Asc(LetraEmb) + 1)
            
            RsEmbarque.AddNew
            InsertoCamposBDNuevo LetraEmb
            
            'Si la divisa esta paga tengo que mantener el valor de la misma (no puedo cambiar por más o por menos).
            If sDivisaPaga Then
                RsEmbarque!EmbDivisa = DivisaAnterior - CCur(tDivisa.Text)
                RsEmbarque!EmbDivisaPaga = 1
            End If
            
            RsEmbarque.Update
            
            Cons = "Select Max(EmbID) From Embarque"
            Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            NuevoID = RsAuxE(0)
            RsAuxE.Close
            GuardoArticulosNuevoEmbarqueyModificoActual ViejoID, NuevoID, False
        Else
            LetraEmb = RsAuxE!EmbCodigo
            NuevoID = RsAuxE!EmbID
            
            If sDivisaPaga Then
                RsAuxE.Edit
                RsAuxE!EmbDivisa = RsAuxE!EmbDivisa + (DivisaAnterior - Format(tDivisa.Text, "#,##0.00000"))
                RsAuxE!EmbDivisaPaga = 1
                RsAuxE.Update
            End If
            
            RsAuxE.Close
            GuardoArticulosNuevoEmbarqueyModificoActual ViejoID, NuevoID, True
        End If
        If Not sDivisaPaga Then
            'Updateo la divisa en base al importe que me da la suma de la tabla pues me pudo modificar el costo
            'de algún artículo y no coincide la resta de la divisa anterior con la nueva por esa causa.
            Cons = "Select Sum(AFoCantidad * AFoPUnitario) from ArticuloFolder Where AFoTipo = " & Folder.cFEmbarque _
                & " And AFoCodigo = " & NuevoID
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            Cons = "Update Embarque Set EmbDivisa = " & RsAux(0) _
                & " Where EmbID = " & NuevoID
            cBase.Execute (Cons)
            RsAux.Close
        End If
        
        If fnc_GetImporteGastos(ViejoID, iImpGasto) Then
            
            'Hay gastos en el viejo x lo que hago gasto en el nuevo
'            Dim aValorEmb As Double
'            fnc_GetImporteGastos NuevoID, aValorEmb
            With regMov
                'OJO paso cero como valor ya que el registro de gasto hace (Divisa en BD - Valor) por lo que voy a levantar el valor mismo de la divisa.
                .Valor = 0 'aValorEmb
                .DivPaga = sDivisaPaga
                .idEmb = NuevoID
                .idEmbViejo = ViejoID
                .HacerNotaPor = iImpGasto     'si cargo este valor --> tengo que hacer la nota.
            End With

        Else
        
            Cons = "Select * From Carpeta, Embarque Where EmbID = " & ViejoID & " And EmbCarpeta = CarID"
            Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            If Not IsNull(RsAuxE!CarBcoEmisor) Then iBcoEm = RsAuxE("CarBcoEmisor") Else iBcoEm = 0
            
            If RsAuxE!CarFormaPago = FormaPago.cFPAnticipado Or RsAuxE!CarFormaPago = FormaPago.cFPCobranza Then
                If Trim(tArbitraje.Text) = "" Or tArbitraje.Text = "0" Then tArbitraje.Text = 1
                'IngresoGastoAutomatico ViejoID, Format(tDivisa.Text, "#,##0.0000") / Format(tArbitraje.Text, "#,##0.0000000"), Format(gFechaServidor, FormatoFP), 0, RsAuxE!CarCodigo & "." & LetraEmb, LetraEmb, "C" & LetraEmb, RsAuxE!CarCodigo, TipoDocumento.CompraCredito, tArbitraje.Text, labBcoEmisor.Caption, labLC.Caption, labFormaPago.Caption, paSubrubroDivisa
                With regGasto
                    .Arbitraje = tArbitraje.Text
                    .BcoCodigo = iBcoEm
                    .BcoNombre = labBcoEmisor.Caption
                    .Carpeta = RsAuxE!CarCodigo & "." & LetraEmb
                    .CodEmbarque = LetraEmb
                    .Cuenta = paSubrubroDivisa
                    .Fecha = gFechaServidor
                    .FormaPago = labFormaPago.Caption
                    .idEmb = ViejoID
                    .LC = labLC.Caption
                    .Proveedor = 0
                    .SerieNro = "C" & LetraEmb & " " & RsAuxE!CarCodigo
                    .TipoDoc = TipoDocumento.CompraCredito
                    .Valor = Format(tDivisa.Text, "#,##0.0000") / Format(tArbitraje.Text, "#,##0.0000000")
                End With
                
            Else
                If Not IsNull(RsAuxE!CarBcoEmisor) And Not IsNull(RsAuxE!CarCartaCredito) Then
                    If Trim(tArbitraje.Text) = "" Or tArbitraje.Text = "0" Then tArbitraje.Text = 1
                    'IngresoGastoAutomatico ViejoID, Format(tDivisa.Text, "#,##0.0000") / Format(tArbitraje.Text, "#,##0.0000000"), Format(gFechaServidor, FormatoFP), RsAuxE!CarBcoEmisor, RsAuxE!CarCodigo & "." & LetraEmb, LetraEmb, "LC", RsAuxE!CarCartaCredito, TipoDocumento.CompraCredito, tArbitraje.Text, labBcoEmisor.Caption, labLC.Caption, labFormaPago.Caption, paSubrubroDivisa
                    With regGasto
                        .Arbitraje = tArbitraje.Text
                        .BcoCodigo = iBcoEm
                        .BcoNombre = labBcoEmisor.Caption
                        .Carpeta = RsAuxE!CarCodigo & "." & LetraEmb
                        .CodEmbarque = LetraEmb
                        .Cuenta = paSubrubroDivisa
                        .Fecha = gFechaServidor
                        .FormaPago = labFormaPago.Caption
                        .idEmb = ViejoID
                        .LC = labLC.Caption
                        .Proveedor = iBcoEm
                        .SerieNro = "LC " & RsAuxE!CarCartaCredito
                        .TipoDoc = TipoDocumento.CompraCredito
                        .Valor = Format(tDivisa.Text, "#,##0.0000") / Format(tArbitraje.Text, "#,##0.0000000")
                    End With
                End If
            End If
            RsAuxE.Close
            
            If regGasto.idEmb > 0 Then
                fnc_GetImporteGastos NuevoID, iImpGasto
                With regMov
                    .idEmb = NuevoID
                    .DivPaga = sDivisaPaga
                    .idEmbViejo = ViejoID           'Con este id tomo la información de proveedor y datos del gasto.
                    .Valor = iImpGasto
                End With
            End If
            
        End If
    End If
    
    If sTodos Then
        'Si modifica todos los embarques similares.--------------------------------------
        Cons = "Update Embarque Set EmbFAprometido = '" & Format(tArriboPrevisto.Text, "mm/dd/yyyy") & "'" _
                & " Where (EmbCarpeta <> " & tCodigo.Tag & " OR EmbID <> '" & RsEmbarque!EmbID & "')" _
                & " And EmbTransporte = " & cTransporte.ItemData(cTransporte.ListIndex) _
                & " And EmbCiudadDestino = " & cDestino.ItemData(cDestino.ListIndex) _
                & " And EmbFAPrometido BETWEEN '" & Format(RestoDias(RsEmbarque!EmbFAPrometido, 5), "mm/dd/yyyy") & "' AND '" & Format(SumoDias(RsEmbarque!EmbFAPrometido, 5), "mm/dd/yyyy") & "'" _
                & " And EmbFArribo = Null"
        cBase.Execute (Cons)
    End If
    
    cBase.CommitTrans
    'FIN TRANSACCION.-------------------------------------------------------------------
    
    On Error GoTo errPasoGastos
    Dim sPrms As String
    'PASO GASTOS DE ZUREO
    If regGasto.idEmb > 0 Or regNota.idEmb > 0 Or regMov.idEmb > 0 Then
    
        With regGasto
            If .idEmb > 0 Then
                sPrms = "Emb= " & .idEmb & ", Valor=" & .Valor & ", Fecha=" & .Fecha & ", Prov=" & .Proveedor & ", TipoDoc= " & .TipoDoc & ", Arb=" & .Arbitraje & ", Cuenta=" & .Cuenta & ", Banco=" & .BcoCodigo
                ModCarpeta.InsertoGastoImportacionZureo .idEmb, .Valor, .Fecha, .Proveedor, .Carpeta, .CodEmbarque, .SerieNro, .TipoDoc, .Arbitraje, .BcoNombre, .LC, .FormaPago, .Cuenta, .BcoCodigo
            End If
        End With
        
        If regNota.idEmb > 0 Then
            sPrms = "Emb=" & regNota.idEmb & ", Valor=" & regNota.Valor & ", DivPaga=" & regNota.DivPaga & ", Cta=" & regNota.SRubro
            HagoNotaCredito regNota.idEmb, regNota.Valor, regNota.DivPaga, 0, regNota.SRubro
        End If
        
        If regMov.idEmb > 0 Then
            Dim iID As Long
            sPrms = "Emb=" & regMov.idEmb & ", Valor=" & regMov.Valor & ", DivPaga=" & regMov.DivPaga & ", EmbVijeo=" & regMov.idEmbViejo
            iID = RealizoMovimientosDeGastos(regMov.idEmb, regMov.Valor, regMov.DivPaga, regMov.idEmbViejo, 0)
            If regMov.HacerNotaPor > 0 And iID > 0 Then
                sPrms = "Emb=" & regMov.idEmbViejo & ", Valor=" & regMov.Valor & ", DivPaga=" & regMov.DivPaga & ", Crédito=" & iID
                RealizoMovimientosDeGastos regMov.idEmbViejo, regMov.HacerNotaPor, regMov.DivPaga, 0, iID
            End If
        End If
        
    End If
    
    'Restauro el formulario.----------------------------------------------------------------
    GoTo Fin
    
errPasoGastos:
    clsGeneral.OcurrioError "Error al pasar los gastos a zureo." & vbCrLf & sPrms, Err.Description, "Embarque"
    GoTo Fin
    
EliminoEmbarque:
    cBase.RollbackTrans
    RsEmbarque.Requery
    MsgBox "El embarque seleccionado ha sido eliminado por otra terminal, verifique.", vbInformation, "ATENCIÓN"
    GoTo Fin
    
ModificaronEmbarque:
    cBase.RollbackTrans
    RsEmbarque.Requery
    
    MsgBox "El embarque seleccionado ha sido modificado por otra terminal, verifique.", vbInformation, "ATENCIÓN"
    GoTo Fin
    
Fin:
    sModificar = False
    InhabilitoCamposEmbarque
    HabilitoCamposCarpeta
    BuscoCarpeta tCodigo.Tag, LetraEmb
    tCodigo.SetFocus
    Screen.MousePointer = 0
    Exit Sub

ErrEmb:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al iniciar la transacción.", Trim(Err.Description)
    Exit Sub
    
ErrResumo:
    Resume ErrModifico
    
ErrModifico:
    cBase.RollbackTrans
    RsEmbarque.Requery
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar almacenar la información.", Trim(Err.Description)
    Exit Sub
End Sub
Private Function RealizoMovimientosDeGastos(idEmbarque As Long, Valor As Currency, sDivisaPaga As Boolean, ByVal iIDViejo As Long, ByVal iCredNew As Long) As Long
'En valor viene la sumatoria del gasto.
Dim DivisaAnterior As Double, aAux As Long
    
    'Cargo el valor de la divisa anterior.
    DivisaAnterior = Valor
    
    Cons = "Select * From Embarque Where EmbID = " & idEmbarque
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Valor = 0 Then
        Valor = Format(CCur(RsAux!EmbDivisa) / CCur(RsAux!EmbArbitraje), FormatoMonedaP)
    Else
        Valor = Format(CCur(RsAux!EmbDivisa) / CCur(RsAux!EmbArbitraje), FormatoMonedaP) - Format(Valor, FormatoMonedaP)
    End If
    RsAux.Close
    
    If Trim(tArbitraje.Text) = "" Or tArbitraje.Text = "0" Then tArbitraje.Text = "1"
    
    If Valor < 0 Then
        'No puede estar costeado e ingresar aca.
        RealizoMovimientosDeGastos = HagoNotaCredito(idEmbarque, Valor * -1, sDivisaPaga, iCredNew, paSubrubroDivisa)
    Else
        If Valor > 0 Then
            If iIDViejo > 0 Then
                'Consulto la compra para sacar la misma serie , numero y proveedor del documento anterior.
                Cons = "Select * from Compra, GastoImportacion, Embarque, Carpeta " _
                        & " Where GImIDSubRubro = " & paSubrubroDivisa _
                        & " And GImNivelFolder = " & Folder.cFEmbarque _
                        & " And GImFolder = " & iIDViejo _
                        & " And ComTipoDocumento = " & TipoDocumento.CompraCredito _
                        & " And GImIDCompra = ComCodigo And GImFolder = EmbID And EmbCarpeta = CarID"
            Else
                'Consulto la compra para sacar la misma serie , numero y proveedor del documento anterior.
                Cons = "Select * from Compra, GastoImportacion, Embarque, Carpeta " _
                        & " Where GImIDSubRubro = " & paSubrubroDivisa _
                        & " And GImNivelFolder = " & Folder.cFEmbarque _
                        & " And GImFolder = " & idEmbarque _
                        & " And ComTipoDocumento = " & TipoDocumento.CompraCredito _
                        & " And GImIDCompra = ComCodigo And GImFolder = EmbID And EmbCarpeta = CarID"
            End If
            
            Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAuxE.EOF Then
                Dim iBcoEm As Integer, iProv As Integer
                If Not IsNull(RsAuxE!CarBcoEmisor) Then iBcoEm = RsAuxE("CarBcoEmisor") Else iBcoEm = 0
                If Not IsNull(RsAuxE("ComProveedor")) Then iProv = RsAuxE("ComProveedor")
                Dim sSerie As String
                If Not IsNull(RsAuxE!ComSerie) Then
                    sSerie = RsAuxE!ComSerie & " " & RsAuxE!ComNumero
                Else
                    sSerie = RsAuxE!ComNumero
                End If
                
                If Not sDivisaPaga Then
                    'aAux = IngresoGastoAutomatico(idEmbarque, Valor, Format(gFechaServidor, FormatoFP), RsAuxE!ComProveedor, RsAuxE!CarCodigo & "." & sCodEmb, sCodEmb, RsAuxE!ComSerie, RsAuxE!ComNumero, TipoDocumento.CompraCredito, RsAuxE!EmbArbitraje, labBcoEmisor.Caption, labLC.Caption, labFormaPago.Caption, paSubrubroDivisa)
                    aAux = ModCarpeta.InsertoGastoImportacionZureo(idEmbarque, Valor, Format(gFechaServidor, FormatoFP), iProv, RsAuxE!CarCodigo & "." & RsAuxE("EmbCodigo"), RsAuxE("EmbCodigo"), sSerie, TipoDocumento.CompraCredito, RsAuxE!EmbArbitraje, labBcoEmisor.Caption, labLC.Caption, labFormaPago.Caption, paSubrubroDivisa, iBcoEm)
                Else
                    aAux = ModCarpeta.InsertoGastoImportacionZureo(idEmbarque, Valor, Format(gFechaServidor, FormatoFP), iProv, RsAuxE!CarCodigo & "." & RsAuxE("EmbCodigo"), RsAuxE("EmbCodigo"), sSerie, TipoDocumento.CompraCredito, RsAuxE!EmbArbitraje, labBcoEmisor.Caption, labLC.Caption, labFormaPago.Caption, paSubrubroDivisa, iBcoEm, True)
                End If
                RealizoMovimientosDeGastos = aAux
                
            End If
            RsAuxE.Close
        End If
    End If

End Function
Private Function HagoNotaCredito(idEmbarque As Long, aValor As Currency, sDivisaPaga As Boolean, CreditoNuevo As Long, SubRubro As Long) As Long
On Error GoTo errHN
Dim IDNota As Long
Dim rsDC As rdoResultset, rsCP As rdoResultset
Dim cDifCambio As Currency, PorcDivisa As Currency, aAmortiza As Currency
Dim aValorTC As Currency, aCantTC As Integer
    
Dim sPrms As String
    aValorTC = 0: aCantTC = 0
    
    Cons = "Select * From Compra, GastoImportacion, Embarque, Carpeta " _
            & " Where GImIDSubRubro = " & paSubrubroDivisa _
            & " And GImNivelFolder = " & Folder.cFEmbarque _
            & " And GImFolder = " & idEmbarque _
            & " And ComMoneda = " & paMonedaDolar _
            & " And ComTipoDocumento = " & TipoDocumento.CompraCredito _
            & " And GImIDCompra = ComCodigo And GImFolder = EmbID And EmbCarpeta = CarID"
            
    Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAuxE.EOF Then
    
        'Inserto Nota en tabla compra.
        Dim iBcoEm As Integer, iProv As Integer
        If Not IsNull(RsAuxE("ComProveedor")) Then iProv = RsAuxE("ComProveedor")
        sPrms = "Asigno Banco"
        If Not IsNull(RsAuxE!CarBcoEmisor) Then
            sPrms = "Asigno Banco = " & RsAuxE("CarBcoEmisor")
            iBcoEm = RsAuxE("CarBcoEmisor")
        Else
            iBcoEm = 0
        End If
        sPrms = "Emb=" & idEmbarque & "Valor=" & aValor & ", Prov=" & iProv & ", TipoDoc=" & TipoDocumento.CompraNotaCredito & ", Arbitraje=" & RsAuxE!EmbArbitraje & ", Cta=" & SubRubro & ", Banco=" & iBcoEm
        IDNota = ModCarpeta.InsertoGastoImportacionZureo(idEmbarque, aValor * -1, Format(gFechaServidor, FormatoFP), iProv, RsAuxE!CarCodigo & "." & RsAuxE!EmbCodigo, RsAuxE!EmbCodigo, Trim(RsAuxE!ComNumero), TipoDocumento.CompraNotaCredito, RsAuxE!EmbArbitraje, labBcoEmisor.Caption, labLC.Caption, labFormaPago.Caption, SubRubro, iBcoEm)
        sPrms = "NOTA OK"
        HagoNotaCredito = IDNota
        If IDNota > 0 Then
            sPrms = "INVOCO SP"
            cBase.Execute "EXEC prg_Embarque_NotaGastoImportacion " & idEmbarque & "," & IDNota & ", " & aValor & ", " & CreditoNuevo & ", " & IIf(sDivisaPaga, 1, 0)
        End If
    End If
    RsAuxE.Close
    Exit Function

errHN:
    clsGeneral.OcurrioError "Error la intentar hacer la nota." & sPrms, Err.Description, "Nota de crédito"
End Function
Private Sub InsertoCamposBD(LetraCodigo As String)
'Letra es el código de embarque. A, B, C, .......
    
    'ID de carpeta.-------------------------------------------------
    RsEmbarque!Embcarpeta = tCodigo.Tag
    RsEmbarque!EmbCodigo = UCase(LetraCodigo)
    If Trim(tConocimiento.Text) = "" Then RsEmbarque!EmbBL = Null Else RsEmbarque!EmbBL = Trim(tConocimiento.Text)
    RsEmbarque!EmbMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    RsEmbarque!EmbDivisa = Trim(tDivisa.Text)
    If sNuevo Then RsEmbarque!EmbDivisaPaga = 0
    If Trim(tArbitraje.Text) = "" Then RsEmbarque!EmbArbitraje = 1 Else RsEmbarque!EmbArbitraje = tArbitraje.Text
    If cAgencia.ListIndex > -1 Then RsEmbarque!EmbAgencia = cAgencia.ItemData(cAgencia.ListIndex) Else RsEmbarque!EmbAgencia = Null
    If cMTransporte.ListIndex > -1 Then RsEmbarque!EmbMedioTransporte = cMTransporte.ItemData(cMTransporte.ListIndex) Else RsEmbarque!EmbMedioTransporte = Null
    If cTransporte.ListIndex > -1 Then RsEmbarque!EmbTransporte = cTransporte.ItemData(cTransporte.ListIndex) Else RsEmbarque!EmbTransporte = Null
    If cDestino.ListIndex > -1 Then RsEmbarque!EmbCiudadDestino = cDestino.ItemData(cDestino.ListIndex) Else RsEmbarque!EmbCiudadDestino = Null
    If cOrigen.ListIndex > -1 Then RsEmbarque!EmbCiudadOrigen = cOrigen.ItemData(cOrigen.ListIndex) Else RsEmbarque!EmbCiudadOrigen = Null
    If cLocal.ListIndex > -1 Then RsEmbarque!EmbLocal = cLocal.ItemData(cLocal.ListIndex) Else RsEmbarque!EmbLocal = Null
    If Trim(tFlete.Text) <> "" Then RsEmbarque!EmbFlete = tFlete.Text Else RsEmbarque!EmbFlete = Null
    If chFletePago.Value = 1 Then RsEmbarque!EmbFletePago = 1 Else RsEmbarque!EmbFletePago = 0
    If cboPrioridad.ListIndex > -1 Then RsEmbarque("EmbPrioridad") = cboPrioridad.ItemData(cboPrioridad.ListIndex)
    If sNuevo Then RsEmbarque!EmbCosteado = 0
    If IsDate(tEmbPrevisto.Text) Then RsEmbarque!EmbFEPrometido = Format(tEmbPrevisto.Text, "mm/dd/yyyy") Else RsEmbarque!EmbFEPrometido = Null
    If IsDate(tEmbarco.Text) Then RsEmbarque!EmbFEmbarque = Format(tEmbarco.Text, "mm/dd/yyyy") Else RsEmbarque!EmbFEmbarque = Null
    If IsDate(tArriboPrevisto.Text) Then RsEmbarque!EmbFAPrometido = Format(tArriboPrevisto.Text, "mm/dd/yyyy") Else RsEmbarque!EmbFAPrometido = Null
    If Trim(tComentario.Text) <> "" Then RsEmbarque!EmbComentario = Trim(tComentario.Text) Else RsEmbarque!EmbComentario = Null
    If IsDate(tUltFechaEmbarque.Text) Then RsEmbarque!EmbUltFechaEmbarque = Format(tUltFechaEmbarque.Text, "mm/dd/yyyy") Else RsEmbarque!EmbUltFechaEmbarque = Null
    RsEmbarque!EmbFModificacion = Format(Now, "mm/dd/yyyy hh:mm:ss")
    
End Sub
Private Sub InsertoCamposBDNuevo(LetraCodigo As String)
'Letra es el código de embarque. A, B, C, .......
    
    'ID de carpeta.-------------------------------------------------
    RsEmbarque!Embcarpeta = tCodigo.Tag
    RsEmbarque!EmbCodigo = UCase(LetraCodigo)
    If Trim(tConocimiento.Text) = "" Then RsEmbarque!EmbBL = Null Else RsEmbarque!EmbBL = Trim(tConocimiento.Text)
    RsEmbarque!EmbMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    'Consulto la plata que tengo en la tabla AFoArticulo
    RsEmbarque!EmbDivisa = Trim(tDivisa.Text)
    RsEmbarque!EmbDivisaPaga = 0
    If Trim(tArbitraje.Text) = "" Then RsEmbarque!EmbArbitraje = 1 Else RsEmbarque!EmbArbitraje = Trim(tArbitraje.Text)
    If cAgencia.ListIndex > -1 Then RsEmbarque!EmbAgencia = cAgencia.ItemData(cAgencia.ListIndex) Else RsEmbarque!EmbAgencia = Null
    If cMTransporte.ListIndex > -1 Then RsEmbarque!EmbMedioTransporte = cMTransporte.ItemData(cMTransporte.ListIndex) Else RsEmbarque!EmbMedioTransporte = Null
    If cTransporte.ListIndex > -1 Then RsEmbarque!EmbTransporte = cTransporte.ItemData(cTransporte.ListIndex) Else RsEmbarque!EmbTransporte = Null
    If cDestino.ListIndex > -1 Then RsEmbarque!EmbCiudadDestino = cDestino.ItemData(cDestino.ListIndex) Else RsEmbarque!EmbCiudadDestino = Null
    If cOrigen.ListIndex > -1 Then RsEmbarque!EmbCiudadOrigen = cOrigen.ItemData(cOrigen.ListIndex) Else RsEmbarque!EmbCiudadOrigen = Null
    If cLocal.ListIndex > -1 Then RsEmbarque!EmbLocal = cLocal.ItemData(cLocal.ListIndex) Else RsEmbarque!EmbLocal = Null
    If Trim(tFlete.Text) <> "" Then RsEmbarque!EmbFlete = tFlete.Text Else RsEmbarque!EmbFlete = Null
    If chFletePago.Value = 1 Then RsEmbarque!EmbFletePago = 1 Else RsEmbarque!EmbFletePago = 0
    If cboPrioridad.ListIndex > -1 Then RsEmbarque("EmbPrioridad") = cboPrioridad.ItemData(cboPrioridad.ListIndex)
    RsEmbarque!EmbCosteado = 0
    If IsDate(tEmbPrevisto.Text) Then RsEmbarque!EmbFEPrometido = Format(tEmbPrevisto.Text, "mm/dd/yyyy") Else RsEmbarque!EmbFEPrometido = Null
    RsEmbarque!EmbFEmbarque = Null
    RsEmbarque!EmbFAPrometido = Null
    RsEmbarque!EmbComentario = Null
    RsEmbarque!EmbUltFechaEmbarque = Null
    RsEmbarque!EmbFModificacion = Format(Now, "mm/dd/yyyy hh:mm:ss")
    
End Sub

Private Function DeseaNuevoEmbarque(ByVal idEmbarque As Long) As Boolean

    DeseaNuevoEmbarque = False
    Cons = "Select * From ArticuloFolder Where AFoTipo = " & Folder.cFEmbarque _
        & " And AFoCodigo = " & idEmbarque
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        For I = 1 To vsArticulo.Rows - 1
            If Val(vsArticulo.Cell(flexcpText, I, 0)) = RsAux!AFoArticulo Then
                'Si es mayor supongo que esta agregando a este.
                If RsAux!AFoCantidad > Val(vsArticulo.Cell(flexcpText, I, 2)) Then
                    If MsgBox("Ha modificado la cantidad de los artículos del embarque, si desea puede ceder la diferencia a un nuevo embarque de lo contrario se bajaran de la cantidad total de la carpeta." & Chr(13) _
                        & "¿Desea dejar la diferencia a un nuevo embarque?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
                        DeseaNuevoEmbarque = True: RsAux.Close: Exit Function
                    Else
                        DeseaNuevoEmbarque = False: RsAux.Close: Exit Function
                    End If
                End If
            End If
        Next I
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Function
Private Sub GuardoArticulosNuevoEmbarqueyModificoActual(OldID As Long, NewID As Long, HayDatos As Boolean)
Dim RsArt As rdoResultset, RsA As rdoResultset
    
    'Primero borro aquellos que no están más en la lista.
    Dim IDArticulos As String
    IDArticulos = ""
    For I = 1 To vsArticulo.Rows - 1
        If IDArticulos = "" Then IDArticulos = Val(vsArticulo.Cell(flexcpText, I, 0)) Else IDArticulos = IDArticulos & ", " & Val(vsArticulo.Cell(flexcpText, I, 0))
    Next I
    Cons = "Select * From ArticuloFolder Where AFoTipo = " & Folder.cFEmbarque _
            & " And AFoCodigo = " & OldID _
            & " And AFoArticulo Not IN (" & IDArticulos & ")"
    Set RsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsArt.EOF
        RsArt.Delete
        RsArt.MoveNext
    Loop
    RsArt.Close
    
    With vsArticulo
        For I = 1 To .Rows - 2
            'Veo si el artículo esta en el viejo
            Cons = "Select * From ArticuloFolder Where AFoTipo = " & Folder.cFEmbarque _
                & " And AFoCodigo = " & OldID _
                & " And AFoArticulo = " & Val(.Cell(flexcpText, I, 0))
            Set RsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If RsArt.EOF Then
                'Me da un artículo nuevo
                Cons = "Insert into ArticuloFolder (AFoTipo, AFoCodigo, AFoArticulo, AFoCantidad, AFoPUnitario) Values (" _
                    & Folder.cFEmbarque & ", " & OldID & ", " & Val(.Cell(flexcpText, I, 0)) & ", " & Val(.Cell(flexcpText, I, 2)) & ", " & CCur(.Cell(flexcpText, I, 3)) & ")"
                cBase.Execute (Cons)
            Else
                If RsArt!AFoCantidad > Val(vsArticulo.Cell(flexcpText, I, 2)) Then
                    Cons = "Select * From ArticuloFolder Where AFoTipo = " & Folder.cFEmbarque _
                            & " And AFoCodigo = " & NewID _
                            & " And AFoArticulo = " & RsArt!AFoArticulo
                    Set RsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If RsA.EOF Then
                        'INSERTO DIFERENCIA EN EL NUEVO.
                        Cons = "Insert into ArticuloFolder (AFoTipo, AFoCodigo, AFoArticulo, AFoCantidad, AFoPUnitario) Values (" _
                            & Folder.cFEmbarque & ", " & NewID & ", " & RsArt!AFoArticulo & ", " & RsArt!AFoCantidad - Val(.Cell(flexcpText, I, 2)) & ", " & CCur(vsArticulo.Cell(flexcpText, I, 3)) & ")"
                        cBase.Execute (Cons)
                    Else
                        RsA.Edit
                        RsA!AFoCantidad = RsA!AFoCantidad + (RsArt!AFoCantidad - Val(.Cell(flexcpText, I, 2)))
                        RsA.Update
                    End If
                    RsA.Close
                    'Arreglo el artículo del embarque que estoy modificando osea ViejoID
                    If Val(.Cell(flexcpText, I, 2)) > 0 Then
                        RsArt.Edit
                        RsArt!AFoCantidad = Val(.Cell(flexcpText, I, 2))
                        RsArt.Update
                    Else
                        RsArt.Delete
                    End If
                ElseIf RsArt!AFoCantidad < Val(.Cell(flexcpText, I, 2)) Then
                    RsArt.Edit
                    RsArt!AFoCantidad = Val(.Cell(flexcpText, I, 2))
                    RsArt.Update
                End If
            End If
            RsArt.Close
        Next I
    End With
    
End Sub
Private Sub GuardoArticulosEmbarqueModificados(idEmbarque As Long)
Dim IDArticulos As String
Dim RsArt As rdoResultset, RsRC As rdoResultset

    IDArticulos = ""
    For I = 1 To vsArticulo.Rows - 1
        If Val(vsArticulo.Cell(flexcpText, I, 0)) > 0 Then
            'Guardo string para verificar si eliminó algún artículo.----------------------
            If IDArticulos = "" Then IDArticulos = Val(vsArticulo.Cell(flexcpText, I, 0)) Else IDArticulos = IDArticulos & ", " & Val(vsArticulo.Cell(flexcpText, I, 0))
            
            Cons = "Select * From ArticuloFolder Where AFoTipo = " & Folder.cFEmbarque _
                & " And AFoCodigo = " & idEmbarque & " And AFoArticulo = " & vsArticulo.Cell(flexcpText, I, 0)
            Set RsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            If RsArt.EOF Then
                'Es nuevo.---------------
                Cons = "Insert into ArticuloFolder (AFoTipo, AFoCodigo, AFoArticulo, AFoCantidad, AFoPUnitario) Values (" _
                    & Folder.cFEmbarque & ", " & idEmbarque & ", " & vsArticulo.Cell(flexcpText, I, 0) & ", " & Val(vsArticulo.Cell(flexcpText, I, 2)) & ", " & CCur(vsArticulo.Cell(flexcpText, I, 3)) & ")"
                cBase.Execute (Cons)
                'Veo si esta en la tabla RemitoCompraRenglon
                Cons = "Select * From RemitoCompra, RemitoCompraRenglon" _
                    & " Where RCoTipoFolder = " & Folder.cFEmbarque & " And RCoIDFolder = " & idEmbarque _
                    & " And RCoCodigo = RCRRemito And RCRArticulo = " & vsArticulo.Cell(flexcpText, I, 0)
                Set RsRC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If RsRC.EOF Then
                    'Agregó el artículo al stock.---------
                    MarcoMovimientoStockFisico UsuLogueado, TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), vsArticulo.Cell(flexcpText, I, 0), Val(vsArticulo.Cell(flexcpText, I, 2)), paEstadoArticuloEntrega, 1
                    MarcoMovimientoStockTotal vsArticulo.Cell(flexcpText, I, 0), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, Val(vsArticulo.Cell(flexcpText, I, 2)), 1
                    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), vsArticulo.Cell(flexcpText, I, 0), Val(vsArticulo.Cell(flexcpText, I, 2)), paEstadoArticuloEntrega, 1
                Else
                    If Val(vsArticulo.Cell(flexcpText, I, 2)) > RsRC!RCRCantidad Then
                        'Cantidad ingresada es mayor a la que ingresó por el local.
                        MarcoMovimientoStockFisico UsuLogueado, TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), vsArticulo.Cell(flexcpText, I, 0), Val(vsArticulo.Cell(flexcpText, I, 2)) - RsRC!RCRCantidad, paEstadoArticuloEntrega, 1
                        MarcoMovimientoStockTotal vsArticulo.Cell(flexcpText, I, 0), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, Val(vsArticulo.Cell(flexcpText, I, 2)) - RsRC!RCRCantidad, 1
                        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), vsArticulo.Cell(flexcpText, I, 0), Val(vsArticulo.Cell(flexcpText, I, 2)) - RsRC!RCRCantidad, paEstadoArticuloEntrega, 1
                    End If
                End If
                RsRC.Close
            Else
                Cons = "Select * From RemitoCompra, RemitoCompraRenglon" _
                    & " Where RCoTipoFolder = " & Folder.cFEmbarque & " And RCoIDFolder = " & idEmbarque _
                    & " And RCoCodigo = RCRRemito And RCRArticulo = " & vsArticulo.Cell(flexcpText, I, 0)
                    
                Set RsRC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                
                If RsArt!AFoCantidad <> Val(vsArticulo.Cell(flexcpText, I, 2)) Or RsArt!AFoPUnitario <> CCur(vsArticulo.Cell(flexcpText, I, 3)) Then
                    If RsArt!AFoPUnitario <> CCur(vsArticulo.Cell(flexcpText, I, 3)) Then
                        'Modifico el importe del ArticuloFolder de la SubCarpeta.---------------
                        Cons = "Update ArticuloFolder Set AFoPUnitario = " & CCur(vsArticulo.Cell(flexcpText, I, 3)) _
                            & " Where AFoTipo = " & Folder.cFSubCarpeta & " And AFoArticulo = " & vsArticulo.Cell(flexcpText, I, 0) _
                            & " And AFoCodigo IN (Select Distinct(SubID) From SubCarpeta Where SubEmbarque = " & idEmbarque & ")"
                        cBase.Execute (Cons)
                    End If
                    If RsRC.EOF Then
                        'Si modificó la cantidad modificó el stock.-------------------------------
                        If RsArt!AFoCantidad > Val(vsArticulo.Cell(flexcpText, I, 2)) Then
                            'Quito artículos al stock.
                            MarcoMovimientoStockFisico UsuLogueado, TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), RsArt!AFoArticulo, RsArt!AFoCantidad - Val(vsArticulo.Cell(flexcpText, I, 2)), paEstadoArticuloEntrega, -1
                            MarcoMovimientoStockTotal vsArticulo.Cell(flexcpText, I, 0), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, RsArt!AFoCantidad - Val(vsArticulo.Cell(flexcpText, I, 2)), -1
                            MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), vsArticulo.Cell(flexcpText, I, 0), RsArt!AFoCantidad - Val(vsArticulo.Cell(flexcpText, I, 2)), paEstadoArticuloEntrega, -1
                        ElseIf RsArt!AFoCantidad < Val(vsArticulo.Cell(flexcpText, I, 2)) Then
                            'Agrego artículos al stock.
                            MarcoMovimientoStockFisico UsuLogueado, TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), RsArt!AFoArticulo, Val(vsArticulo.Cell(flexcpText, I, 2)) - RsArt!AFoCantidad, paEstadoArticuloEntrega, 1
                            MarcoMovimientoStockTotal vsArticulo.Cell(flexcpText, I, 0), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, Val(vsArticulo.Cell(flexcpText, I, 2)) - RsArt!AFoCantidad, 1
                            MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), vsArticulo.Cell(flexcpText, I, 0), Val(vsArticulo.Cell(flexcpText, I, 2)) - RsArt!AFoCantidad, paEstadoArticuloEntrega, 1
                        End If
                    Else
                        'Si modificó la cantidad modificó el stock.-------------------------------
'                        If RsArt!AFoCantidad > Val(vsArticulo.Cell(flexcpText, I, 2)) And RsArt!AFoCantidad = RsRC!RCRCantidad Then
'                            'Quito artículos al stock.
'                            MarcoMovimientoStockFisico UsuLogueado, TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), RsArt!AFoArticulo, RsArt!AFoCantidad - Val(vsArticulo.Cell(flexcpText, I, 2)), paEstadoArticuloEntrega, -1
'                            MarcoMovimientoStockTotal vsArticulo.Cell(flexcpText, I, 0), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, RsArt!AFoCantidad - Val(vsArticulo.Cell(flexcpText, I, 2)), -1
'                            MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), vsArticulo.Cell(flexcpText, I, 0), RsArt!AFoCantidad - Val(vsArticulo.Cell(flexcpText, I, 2)), paEstadoArticuloEntrega, -1
                        If RsArt!AFoCantidad < Val(vsArticulo.Cell(flexcpText, I, 2)) And Val(vsArticulo.Cell(flexcpText, I, 2)) > RsRC!RCRCantidad Then
                            'Agrego artículos al stock.
                            MarcoMovimientoStockFisico UsuLogueado, TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), RsArt!AFoArticulo, Val(vsArticulo.Cell(flexcpText, I, 2)) - RsRC!RCRCantidad, paEstadoArticuloEntrega, 1
                            MarcoMovimientoStockTotal vsArticulo.Cell(flexcpText, I, 0), TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, Val(vsArticulo.Cell(flexcpText, I, 2)) - RsRC!RCRCantidad, 1
                            MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), vsArticulo.Cell(flexcpText, I, 0), Val(vsArticulo.Cell(flexcpText, I, 2)) - RsRC!RCRCantidad, paEstadoArticuloEntrega, 1
                        End If
                    End If

                    RsArt.Edit
                    RsArt!AFoCantidad = Val(vsArticulo.Cell(flexcpText, I, 2))
                    RsArt!AFoPUnitario = CCur(vsArticulo.Cell(flexcpText, I, 3))
                    RsArt.Update
                End If
            End If
            RsArt.Close
        End If
    Next
    
    'Verifico si eliminó alguna fila.-------
    Cons = "Select * From ArticuloFolder Where AFoTipo = " & Folder.cFEmbarque _
        & " And AFoCodigo = " & idEmbarque _
        & " And AFoArticulo NOT IN (" & IDArticulos & ")"
    Set RsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsArt.EOF
        Cons = "Delete ArticuloFolder Where AFoTipo = " & Folder.cFEmbarque & " And AFoCodigo = " & idEmbarque _
            & " And AFoArticulo = " & RsArt!AFoArticulo
        cBase.Execute (Cons)
        
        'Doy de baja al stock.--------------------------
        MarcoMovimientoStockFisico UsuLogueado, TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), RsArt!AFoArticulo, RsArt!AFoCantidad, paEstadoArticuloEntrega, -1
        MarcoMovimientoStockTotal RsArt!AFoArticulo, TipoEstadoMercaderia.Fisico, paEstadoArticuloEntrega, RsArt!AFoCantidad, -1
        MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cLocal.ItemData(cLocal.ListIndex), RsArt!AFoArticulo, RsArt!AFoCantidad, paEstadoArticuloEntrega, -1
        RsArt.MoveNext
    Loop
    RsArt.Close
    
End Sub
Private Sub GuardoArticulosEmbarque(idEmbarque As Long)
    
    'Primero borro los que pueda tener y luego inserto lo que hay en la lista.
    Cons = "Delete ArticuloFolder Where AFoTipo = " & Folder.cFEmbarque & " And AFoCodigo = " & idEmbarque
    cBase.Execute (Cons)
    
    For I = 1 To vsArticulo.Rows - 1
        'Tenía solo por el text 0 pero ahí solo guardo el id de articulo.
        If Val(vsArticulo.Cell(flexcpText, I, 2)) > 0 And Val(vsArticulo.Cell(flexcpText, I, 0)) > 0 Then
            Cons = "Insert into ArticuloFolder (AFoTipo, AFoCodigo, AFoArticulo, AFoCantidad, AFoPUnitario) Values (" _
                & Folder.cFEmbarque & ", " & idEmbarque & ", " & vsArticulo.Cell(flexcpText, I, 0) & ", " & Val(vsArticulo.Cell(flexcpText, I, 2)) & ", " & CCur(vsArticulo.Cell(flexcpText, I, 3)) & ")"
            cBase.Execute (Cons)
        End If
    Next
    
End Sub

'Private Sub DejoCamposIguales()
'On Error GoTo ErrDCI
'    RelojA
'    If Not IsNull(RsEmbarque!EmbCiudadDestino) Then BuscoCodigoEnCombo cDestino, RsEmbarque!EmbCiudadDestino
'    If Not IsNull(RsEmbarque!EmbAgencia) Then BuscoCodigoEnCombo cAgencia, RsEmbarque!EmbAgencia
'    If Not IsNull(RsEmbarque!EmbMedioTransporte) Then BuscoCodigoEnCombo cMTransporte, RsEmbarque!EmbMedioTransporte
''    If Not IsNull(RsEmbarque!EmbContenedor) Then BuscoCodigoEnCombo cContenedor, RsEmbarque!EmbContenedor
'    If Not IsNull(RsEmbarque!EmbMoneda) Then BuscoCodigoEnCombo cMoneda, RsEmbarque!EmbMoneda
'    If cMoneda.ListIndex > -1 Then
'        Cons = "Select * From Moneda Where MonCodigo = " & cMoneda.ItemData(cMoneda.ListIndex)
'        Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'        If Not RsAuxE.EOF Then
'            If RsAuxE!MonArbitraje Then tArbitraje.Enabled = True: tArbitraje.BackColor = Obligatorio Else tArbitraje.Enabled = False: tArbitraje.BackColor = Inactivo: tArbitraje.Text = ""
'        End If
'        RsAuxE.Close
'    End If
'    RelojD
'    Exit Sub
'ErrDCI:
'    clsGeneral.OcurrioError "Error al dejar los campos identicos.", Trim(Err.Description)
'    RelojD
'End Sub

Private Sub AccionPrimerRegistro()
On Error GoTo ErrAPR
    Screen.MousePointer = 11
    RsEmbarque.MoveFirst
    BotonesRegistro False, False, True, True, Toolbar1, Me
    LimpioCamposEmbarque
    CargoDatosEmbarque
    CargoArticulosEmbarque RsEmbarque!EmbID
    IndicoRegistro
    Screen.MousePointer = 0
    Exit Sub
ErrAPR:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar acceder al primer registro."
End Sub
Private Sub AccionRegistroAnterior()
On Error GoTo ErrARA
    Screen.MousePointer = 11
    RsEmbarque.MovePrevious
    If Not RsEmbarque.BOF Then
        BotonesRegistro True, True, True, True, Toolbar1, Me
        LimpioCamposEmbarque
        CargoDatosEmbarque
        CargoArticulosEmbarque RsEmbarque!EmbID
        IndicoRegistro
    Else
        'Tengo que mover el resultset al primero.------------------
        RsEmbarque.MoveFirst
        '-----------------------------------------------------------------------
        BotonesRegistro False, False, True, True, Toolbar1, Me
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrARA:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar acceder al registro anterior."
End Sub
Private Sub AccionRegistroSiguiente()
On Error GoTo ErrARS
    Screen.MousePointer = 11
    RsEmbarque.MoveNext
    If Not RsEmbarque.EOF Then
        BotonesRegistro True, True, True, True, Toolbar1, Me
        LimpioCamposEmbarque
        CargoDatosEmbarque
        CargoArticulosEmbarque RsEmbarque!EmbID
        IndicoRegistro
    Else
        'Me pase en uno.--------------------------
        RsEmbarque.MoveLast
        '-----------------------------------------------
        BotonesRegistro True, True, False, False, Toolbar1, Me
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrARS:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar acceder al siguiente registro."
End Sub
Private Sub AccionUltimoRegistro()
On Error GoTo ErrAUR
    Screen.MousePointer = 11
    RsEmbarque.MoveLast
    BotonesRegistro True, True, False, False, Toolbar1, Me
    LimpioCamposEmbarque
    CargoDatosEmbarque
    CargoArticulosEmbarque RsEmbarque!EmbID
    IndicoRegistro
    Screen.MousePointer = 0
    Exit Sub
ErrAUR:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar acceder al último registro."
End Sub

Private Function InvocoNuevaCarpeta() As Boolean
On Error GoTo ErrINC
Dim frmCarpeta As New MaCarpeta
    
    frmCarpeta.pSeleccionado = -1
    frmCarpeta.pCodBoquilla = 0
    frmCarpeta.Show vbModal, Me
    RelojA
    RsEmbarque.Requery
    If frmCarpeta.pSeleccionado > 0 Then
        'Pudo haber levantado otra carpeta que ya tiene embarques.
        Cons = "Select * From Embarque Where EmbCarpeta = " & frmCarpeta.pSeleccionado
        Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAuxE.EOF Then
            RsAuxE.Close
            MsgBox "Se cargo una carpeta que posee embarques, esta operación no esta permitida.", vbExclamation, "ATENCIÓN"
            InvocoNuevaCarpeta = False
        Else
            RsAuxE.Close
            CargoDatosCarpetaPorID frmCarpeta.pSeleccionado
            CargoDatosBoquilla frmCarpeta.pSeleccionado, frmCarpeta.pCodigosBoquilla
            InvocoNuevaCarpeta = True
        End If
    Else
        InvocoNuevaCarpeta = False
    End If
    Set frmCarpeta = Nothing
    RelojD
    Exit Function
ErrINC:
    clsGeneral.OcurrioError "Error al invocar al formaulario de carpeta.", Trim(Err.Description)
    RelojD
End Function
Private Sub CargoCombosForm()
    'Proveedores.----------------------------
    Cons = "Select PExCodigo, PExNombre From ProveedorExterior" _
        & " Order by PExNombre"
    CargoCombo Cons, cProveedor, ""
    '----------------------------------------------
    'Cargo los transportes.-----------------
    Cons = "Select TraCodigo, TraNombre From Transporte Order by TraNombre"
    CargoCombo Cons, cTransporte, ""
    '----------------------------------------------
    'Cargo Ciudad Origen y Ciudad Destino.
    Cons = "Select CiuCodigo, CiuNombre from Ciudad" _
        & " Order by CiuNombre"
    CargoCombo Cons, cOrigen, ""
    CargoCombo Cons, cDestino, ""
    '----------------------------------------------
    'Cargo locales.-----------------------------
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal"
    CargoCombo Cons, cLocal, ""
    '----------------------------------------------
    'Cargo los Medios de Transporte.-----
    cMTransporte.AddItem RetornoMedioTransporte(cVAereo)
    cMTransporte.ItemData(cMTransporte.NewIndex) = cVAereo
    cMTransporte.AddItem RetornoMedioTransporte(cVMaritimo)
    cMTransporte.ItemData(cMTransporte.NewIndex) = cVMaritimo
    cMTransporte.AddItem RetornoMedioTransporte(CVTerrestre)
    cMTransporte.ItemData(cMTransporte.NewIndex) = CVTerrestre
    '----------------------------------------------
    'Cargo las Monedas.--------------------
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda, ""
    '----------------------------------------------
    'Cargo las Agencias de Transportes.--------------------
    Cons = "Select ATrCodigo, ATrNombre From AgenciaTransporte Order by ATrNombre"
    CargoCombo Cons, cAgencia, ""
    '----------------------------------------------
    
    Cons = "SELECT CodID, RTRIM(CodTexto), IsNull(CodValor1, 0) FROM Codigos WHERE CodCual = 165 ORDER BY CodTexto"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        cboPrioridad.AddItem Trim(RsAux(1))
        cboPrioridad.ItemData(cboPrioridad.NewIndex) = RsAux(0)
        If RsAux(2) = 1 Then PrioridadDefault = RsAux(0)
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub

Private Sub CargoArticulosEmbarque(embarque As Long)
On Error GoTo ErrCAE
Dim RsSC As rdoResultset
    Cons = "Select ArticuloFolder.*, ArtNombre From ArticuloFolder, Articulo Where AFoTipo = " & Folder.cFEmbarque _
        & " And AFoCodigo = " & embarque & " And AFoArticulo = ArtID"
    Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAuxE.EOF
        vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 0) = RsAuxE!AFoArticulo
        vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 1) = Trim(RsAuxE!ArtNombre)
        vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 2) = Trim(RsAuxE!AFoCantidad)
        vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 3) = Format(RsAuxE!AFoPUnitario, FormatoMonedaP)
        'Cargo la cantidad de artículos que hay en subcarpetas.---------------------------
        Cons = "Select Sum(AFoCantidad) From SubCarpeta, ArticuloFolder " _
            & " Where SubEmbarque = " & RsEmbarque!EmbID _
            & " And AFoTipo = " & Folder.cFSubCarpeta & " And AFoArticulo = " & RsAuxE!AFoArticulo _
            & " And SubID = AFoCodigo"
        Set RsSC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not IsNull(RsSC(0)) Then
            vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 4) = RsSC(0)
        Else
            vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 4) = 0
        End If
        RsSC.Close
        vsArticulo.AddItem ""
        RsAuxE.MoveNext
    Loop
    RsAuxE.Close
    Exit Sub
ErrCAE:
    clsGeneral.OcurrioError "Error al cargar los artículos del embarque.", Trim(Err.Description)
End Sub
Private Function EstaEnGrilla(ByVal idArt As Long) As Integer
Dim iPos As Integer
    EstaEnGrilla = 0
    For iPos = 1 To vsArticulo.Rows - 1
        If Val(vsArticulo.Cell(flexcpText, iPos, 0)) = idArt Then EstaEnGrilla = iPos: Exit Function
    Next
End Function
Private Sub CargoDatosBoquilla(ByVal idCarpeta As Long, Optional sCodigo As String = "")
On Error GoTo ErrCDB
Dim iFilaAux As Integer
    Screen.MousePointer = 11
    Cons = "Select * From Pedido, ArticuloFolder, Articulo Where PedCarpeta = " & idCarpeta _
        & " And AFoTipo = " & Folder.cFPedido _
        & " And AFoCodigo = PedCodigo And AFoArticulo = ArtID"
    If sCodigo <> "" Then Cons = Cons & " and PedCodigo IN (" & sCodigo & ")"
    Set RsAuxE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAuxE.EOF Then
        If Not IsNull(RsAuxE!PedFEmbarque) Then tEmbPrevisto.Text = Format(RsAuxE!PedFEmbarque, FormatoFP)
        Do While Not RsAuxE.EOF
            'Veo si ya inserte el artículo, si esto ocurre sumo cantidad.
            iFilaAux = EstaEnGrilla(RsAuxE!AFoArticulo)
            If iFilaAux = 0 Then
                vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 0) = RsAuxE!AFoArticulo
                vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 1) = Trim(RsAuxE!ArtNombre)
                vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 2) = Trim(RsAuxE!AFoCantidad)
                vsArticulo.Cell(flexcpText, vsArticulo.Rows - 1, 3) = Format(RsAuxE!AFoPUnitario, FormatoMonedaP)
                vsArticulo.AddItem ""
            Else
                'Sumo la cantidad y válido el precio unitario.
                vsArticulo.Cell(flexcpText, iFilaAux, 2) = Val(vsArticulo.Cell(flexcpText, iFilaAux, 2)) + Val(RsAuxE!AFoCantidad)
            End If
            RsAuxE.MoveNext
        Loop
        
    End If
    RsAuxE.Close
    Screen.MousePointer = 0
    Exit Sub
ErrCDB:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar la información del pedido de boquilla."
End Sub
Private Sub PreparoAccionNuevo()
    
    'Prendo Señal que es uno nuevo.----------------------
    sNuevo = True
    '-----------------------------------------------------------------

    'Habilito y Desabilito Botones.---------------------------
    Botones False, False, False, True, True, Toolbar1, Me
    BotonesRegistro False, False, False, False, Toolbar1, Me
    Toolbar1.Buttons("carpeta").Enabled = False: MnuCarpeta.Enabled = False
    MnuEditMemo.Enabled = False
    '-----------------------------------------------------------------

    'Preparo el formulario.-------------------------------------
    LimpioCamposEmbarque
    HabilitoCamposEmbarque
    BuscoCodigoEnCombo cboPrioridad, CLng(PrioridadDefault)
    InhabilitoCamposCarpeta
    
End Sub

Private Sub LimpioGrilla()

    With vsArticulo
        .Redraw = False
        .ExtendLastCol = True
        .Clear
        .Editable = True
        .Rows = 2
        .Cols = 5
        .ForeColor = vbBlack
        .FormatString = "IDArticulo|Articulo|>Q|>Precio|SubCarpeta"
        .ColWidth(1) = 2950
        .ColWidth(2) = 900
        .ColWidth(3) = 900
        .ColFormat(3) = "#,##0.0000"
        .ColHidden(0) = True: .ColHidden(4) = True
        .AllowUserResizing = flexResizeColumns
        .Redraw = True
    End With

End Sub

Private Sub vsArticulo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrAE
Dim curSuma As Currency
    
    Select Case Col
        Case 1
            If IsNumeric(vsArticulo.Cell(flexcpText, Row, Col)) Then
                BuscoArticuloXCodigo vsArticulo.Cell(flexcpText, Row, Col), Row
                If vsArticulo.Cell(flexcpText, Row, Col) <> "" Then vsArticulo.Col = Col + 1
            ElseIf vsArticulo.Cell(flexcpText, Row, Col) <> "" Then
                BuscoArticuloPorNombre vsArticulo.Cell(flexcpText, Row, Col), Row
                If vsArticulo.Cell(flexcpText, Row, Col) <> "" Then vsArticulo.Col = Col + 1
            Else
                BorroFila Row
            End If
        Case 2
            'Cantidad
            If vsArticulo.Cell(flexcpText, Row, Col + 1) = "" Then
                vsArticulo.Cell(flexcpText, Row, Col + 1) = BuscoCosto(vsArticulo.Cell(flexcpText, Row, 0))
                If vsArticulo.Cell(flexcpForeColor, Row, Col) <> vbRed Then
                    If vsArticulo.Rows - 1 = Row Then vsArticulo.AddItem ""
                    If Val(vsArticulo.Cell(flexcpText, Row, Col + 1)) <> 0 Then
                        vsArticulo.Row = vsArticulo.Rows - 1: vsArticulo.Col = 1
                    Else
                        vsArticulo.Col = Col + 1
                    End If
                Else
                    vsArticulo.Col = Col + 1
                End If
            End If
            
        Case 3 'Precio
        
            'NO ingreso artículo ni cantidad
            If vsArticulo.Cell(flexcpText, Row, 0) = "" And vsArticulo.Cell(flexcpText, Row, 2) = "" Then Exit Sub
            
            If vsArticulo.Cell(flexcpForeColor, Row, Col) <> vbRed Then
                If vsArticulo.Rows - 1 = Row Then
                    vsArticulo.AddItem "": vsArticulo.Row = vsArticulo.Rows - 1: vsArticulo.Col = 1
                End If
                If IsNumeric(vsArticulo.EditText) Then
                    If BuscoCosto(vsArticulo.Cell(flexcpText, Row, 0)) <> CCur(vsArticulo.EditText) Then
                        If MsgBox("El precio no coincide con el último costo ingresado." & Chr(13) & "¿Desea modificar el costo del artículo.?", vbQuestion + vbYesNo, "CAMBIO DE COSTO") = vbYes Then
                            ModificoCosto vsArticulo.Cell(flexcpText, Row, 0), vsArticulo.EditText
                        End If
                    End If
                Else
                    If vsArticulo.Rows - 1 = Row + 1 Then vsArticulo.Row = vsArticulo.Row + 1: vsArticulo.Col = 1
                End If
            Else
                With vsArticulo
                'Tengo que validar que sea un precio y buscar la ubicación del artículo y calcular el precio.
                    If Not IsNumeric(.EditText) Then Exit Sub
                        For I = 1 To .Rows - 1
                            If Val(.Cell(flexcpValue, I, 0)) = Val(.Cell(flexcpValue, Row, 0)) And .Cell(flexcpForeColor, I, 0) = vbBlack And Val(.Cell(flexcpValue, Row, 0)) > 0 Then
                                curSuma = ((CCur(.Cell(flexcpValue, I, 2)) * CCur(.Cell(flexcpValue, I, 3))) + (CCur(.Cell(flexcpValue, Row, 2)) * CCur(.Cell(flexcpValue, Row, 3)))) / (CCur(.Cell(flexcpValue, I, 2)) + CCur(.Cell(flexcpValue, Row, 2)))
                                .Cell(flexcpText, I, 2) = CCur(.Cell(flexcpValue, I, 2)) + CCur(.Cell(flexcpValue, Row, 2))
                                .Cell(flexcpText, I, 3) = Format(curSuma, "#,##0.0000")
                                'Borro la fila que me ingreso.
                                .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
                                .Cell(flexcpText, Row, 0, Row, .Cols - 1) = ""
                                .Col = 1
                            End If
                        Next I
                End With
            End If
    End Select
    Exit Sub
ErrAE:
    clsGeneral.OcurrioError "Error inesperado al editar la celda.", Trim(Err.Description)
End Sub
Private Sub vsArticulo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not sNuevo And Not sModificar Then Cancel = True
    If vsArticulo.Cell(flexcpText, Row, 0) = "" And Col > 1 Then Cancel = True
    If vsArticulo.Cell(flexcpText, Row, 2) = "" And Col = 3 Then Cancel = True
End Sub
Private Sub vsArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If sNuevo Or sModificar Then
        If KeyCode = vbKeyDelete And Val(vsArticulo.Cell(flexcpText, vsArticulo.Row, 4)) = 0 Then
            With vsArticulo
                'Verifico que no haya ninguna en rojo para este artículo.
                For I = 1 To .Rows - 1
                    If Val(.Cell(flexcpText, I, 0)) = Val(.Cell(flexcpText, .Row, 0)) And .Cell(flexcpForeColor, I, 0) = vbRed Then
                        .Cell(flexcpForeColor, I, 0, I, .Cols - 1) = vbBlack
                        .Cell(flexcpText, I, 0, I, .Cols - 1) = ""
                    End If
                Next I
                BorroFila .Row
            End With
        End If
    End If
    
End Sub
Private Sub vsArticulo_LostFocus()
Dim Calc As Currency
    If Not sNuevo And Not sModificar Then Exit Sub
    Calc = 0
    With vsArticulo
        'Verifico que no haya ninguna en rojo para este artículo.
        For I = 1 To vsArticulo.Rows - 1
            If .Cell(flexcpForeColor, I, 0) = vbRed Then
                .Cell(flexcpForeColor, I, 0, I, .Cols - 1) = vbBlack
                .Cell(flexcpText, I, 0, I, .Cols - 1) = ""
            End If
            Calc = Calc + Val(vsArticulo.Cell(flexcpValue, I, 3)) * Val(vsArticulo.Cell(flexcpValue, I, 2))
        Next I
    End With
    If tDivisa.Text <> "" Then
        If CCur(tDivisa.Text) <> Calc Then
            MsgBox "El valor de la divisa no coincide con la suma de precios en la lista de artículos." & Chr(13) & Chr(10) & "Verifique si el importe es el real.", vbInformation, "ATENCIÓN"
            'If MsgBox("La divisa es distinta a la suma de precios en la grilla." & Chr(13) & "¿Desea modificar este valor?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
            '    tDivisa.Text = Format(Calc, FormatoMonedaP)
            'End If
        End If
    Else
        tDivisa.Text = Format(Calc, FormatoMonedaP)
    End If
End Sub

Private Sub vsArticulo_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
        Case 1
            
        Case 2, 3
            'Cantidad y precio.
            If Not IsNumeric(vsArticulo.EditText) Then
                MsgBox "El formato no es numérico.", vbExclamation, "ATENCIÓN"
                Cancel = True
            End If
            If Col = 2 Then
                If Val(vsArticulo.Cell(flexcpText, vsArticulo.Row, 4)) > Val(vsArticulo.EditText) Then
                    MsgBox "Hay artículos ingresados en subcarpetas que superan la cantidad ingresada.", vbExclamation, "ATENCIÓN"
                    Cancel = True
                End If
            End If
    End Select
End Sub
Private Sub BuscoArticuloPorNombre(Articulo As String, Fila As Long)
On Error GoTo ErrPLDA
Dim Resultado As String

    Cons = "Select ArtCodigo, 'Código' = ArtCodigo, Nombre = ArtNombre" _
        & " From Articulo" _
        & " Where ArtNombre LIKE '" & Trim(Articulo) & "%'" _
        & " OR ArtID IN (SELECT ACFArticulo FROM ArticuloCodigoFabrica WHERE ACFCodigo = '" & Trim(Articulo) & "')"
    
    RelojA
    
    Dim objAyuda As New clsListadeAyuda
    If objAyuda.ActivarAyuda(cBase, Cons, , 1, "Lista de Artículos") Then
        Resultado = objAyuda.RetornoDatoSeleccionado(0)
    End If
    Set objAyuda = Nothing
'    sqlAyuda.ActivoListaAyuda Cons, False, cBase.Connect
    'Obtengo si hay seleccionado.---------------
'    Resultado = sqlAyuda.ItemSeleccionado
    'Destruyo la clase.------------------------------
'    Set sqlAyuda = Nothing
    RelojA
    If Resultado <> "" Then
        If IsNumeric(Resultado) Then
           BuscoArticuloXCodigo CLng(Resultado), Fila
        Else
            RelojD
            MsgBox "Se espera que se retorne el código de artículo.", vbInformation, "ATENCIÓN"
        End If
    Else
        BorroFila Fila
    End If
    RelojD
    Exit Sub
    
ErrPLDA:
    RelojD
    clsGeneral.OcurrioError "Error al presentar la lista de ayuda."
End Sub

Private Sub BorroFila(Fila As Long)
    If sModificar And Val(vsArticulo.Cell(flexcpText, Fila, 0)) > 0 Then
        If MsgBox("Si desea ceder este artículo al nuevo embarque debe ponerle como cantidad CERO sino será dado de baja del embarque." & Chr(13) & "¿Confirma eliminar el artículo del embarque?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    End If
    If Fila = vsArticulo.Rows - 1 Then vsArticulo.Cell(flexcpText, Fila, 0, Fila, 3) = "": vsArticulo.Cell(flexcpForeColor, Fila, 0, Fila, 3) = vbBlack Else vsArticulo.RemoveItem (Fila)
End Sub
Private Function ExistenGastos() As Boolean
    
    Cons = "Select GImNivelFolder " _
        & " From GastoImportacion " _
        & " Where GImNivelFolder = " & Folder.cFCarpeta & " And GImFolder = " & RsEmbarque!Embcarpeta
    Cons = Cons & " Union All "
    'Consulta para Embarques.-------------------------------------------------------------
    Cons = Cons & "Select GImNivelFolder From GastoImportacion " _
        & " Where GImNivelFolder = " & Folder.cFEmbarque & " And GImFolder = " & RsEmbarque!EmbID
    Cons = Cons & " Union All "
    'Consulta para SubCarpetas.-------------------------------------------------------------
    Cons = Cons & "Select GImNivelFolder From GastoImportacion, Embarque, SubCarpeta " _
        & " Where GImNivelFolder = " & Folder.cFSubCarpeta & " And GImFolder = SubID And EmbID = SubEmbarque " _
        & " And EmbID = " & RsEmbarque!EmbID
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsAux.EOF Then ExistenGastos = True Else ExistenGastos = False
    RsAux.Close
    
End Function
Private Sub ModificoCosto(Articulo As Long, Costo As Currency)
On Error GoTo ErrMC
    RelojA
    FechaDelServidor
    Cons = "INSERT INTO PrecioDeCosto (PCoArticulo, PCoFecha, PCoImporte) Values (" & Articulo & ", '" _
        & Format(gFechaServidor, sqlFormatoFH) & "', " & Costo & ")"
    cBase.Execute (Cons)
    RelojD
    Exit Sub
ErrMC:
    clsGeneral.OcurrioError "Error al modificar el costo del artículo.", Trim(Err.Description)
    RelojD
End Sub

Private Sub BuscoTCParaPresentar()
Dim cTC As Currency
Dim dFecha As Date, retVal As Integer, dAux As Date
On Error GoTo ErrBD

    'Veo cual es el último día habil del mes pasado.
    dFecha = PrimerDia(Date) - 1
    retVal = Weekday(dFecha)
    Do While retVal < 2 Or retVal > 6
        dFecha = dFecha - 1
        retVal = Weekday(dFecha)
    Loop
    dAux = dFecha
    
     cTC = TasadeCambio(paMonedaDolar, paMonedaPesos, dFecha)
    
    If dAux = dFecha Then
        'Veo si
        Status.Panels("tasa").Text = "Tasa de Cambio del " & Format(dFecha, "d/mm") & " : " & Format(cTC, "#.000")
    Else
        MsgBox "No esta ingresada la tasa de cambio para el último día hábil del mes anterior.", vbInformation, "ATENCIÓN"
        EjecutarApp App.Path & "\Tasa de cambio.exe"
    End If
    Exit Sub
ErrBD:
    clsGeneral.OcurrioError "Ocurrió un error al buscar la tasa de cambio.", Err.Description
End Sub

Private Sub AccionRefrescoCombos()
Dim idSeleccionado As Long
    
    Screen.MousePointer = 11
    If cProveedor.ListIndex > -1 Then
        idSeleccionado = cProveedor.ItemData(cProveedor.ListIndex)
    Else
        idSeleccionado = 0
    End If
    'Proveedores.----------------------------
    Cons = "Select PExCodigo, PExNombre From ProveedorExterior" _
        & " Order by PExNombre"
    CargoCombo Cons, cProveedor, ""
    '----------------------------------------------
    If idSeleccionado > 0 Then BuscoCodigoEnCombo cProveedor, idSeleccionado
    
    If cTransporte.ListIndex > -1 Then
        idSeleccionado = cTransporte.ItemData(cTransporte.ListIndex)
    Else
        idSeleccionado = 0
    End If
    'Cargo los transportes.-----------------
    Cons = "Select TraCodigo, TraNombre From Transporte Order by TraNombre"
    CargoCombo Cons, cTransporte, ""
    '----------------------------------------------
    If idSeleccionado > 0 Then BuscoCodigoEnCombo cTransporte, idSeleccionado
    
    If cOrigen.ListIndex > -1 Then
        idSeleccionado = cOrigen.ItemData(cOrigen.ListIndex)
    Else
        idSeleccionado = 0
    End If
    'Cargo Ciudad Origen y Ciudad Destino.
    Cons = "Select CiuCodigo, CiuNombre from Ciudad" _
        & " Order by CiuNombre"
    CargoCombo Cons, cOrigen, ""
    If idSeleccionado > 0 Then BuscoCodigoEnCombo cOrigen, idSeleccionado
    If cDestino.ListIndex > -1 Then
        idSeleccionado = cDestino.ItemData(cDestino.ListIndex)
    Else
        idSeleccionado = 0
    End If
    CargoCombo Cons, cDestino, ""
    If idSeleccionado > 0 Then BuscoCodigoEnCombo cDestino, idSeleccionado
    '----------------------------------------------
    
    If cLocal.ListIndex > -1 Then
        idSeleccionado = cLocal.ItemData(cLocal.ListIndex)
    Else
        idSeleccionado = 0
    End If
    'Cargo locales.-----------------------------
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal"
    CargoCombo Cons, cLocal, ""
    If idSeleccionado > 0 Then BuscoCodigoEnCombo cLocal, idSeleccionado
    '----------------------------------------------
   
    If cMoneda.ListIndex > -1 Then
        idSeleccionado = cMoneda.ItemData(cMoneda.ListIndex)
    Else
        idSeleccionado = 0
    End If
    'Cargo las Monedas.--------------------
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda, ""
    If idSeleccionado > 0 Then BuscoCodigoEnCombo cMoneda, idSeleccionado
    '----------------------------------------------
    
    If cAgencia.ListIndex > -1 Then
        idSeleccionado = cAgencia.ItemData(cAgencia.ListIndex)
    Else
        idSeleccionado = 0
    End If
    'Cargo las Agencias de Transportes.--------------------
    Cons = "Select ATrCodigo, ATrNombre From AgenciaTransporte Order by ATrNombre"
    CargoCombo Cons, cAgencia, ""
    If idSeleccionado > 0 Then BuscoCodigoEnCombo cAgencia, idSeleccionado
    '----------------------------------------------
    
    sContenedor = ""
    Cons = "Select ConCodigo, ConAbreviacion from Contenedor" _
        & " Order by ConNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        If sContenedor = "" Then
            sContenedor = "#" & Trim(RsAux!ConCodigo) & ";" & Trim(RsAux!ConAbreviacion)
        Else
            sContenedor = sContenedor & "|" & "#" & Trim(RsAux!ConCodigo) & ";" & Trim(RsAux!ConAbreviacion)
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    sLinea = ""
    Cons = "Select Codigo, Texto From CodigoTexto Where Tipo = 67 Order by Texto"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        If sLinea = "" Then
            sLinea = "#" & Trim(RsAux!Codigo) & ";" & Trim(RsAux!Texto)
        Else
            sLinea = sLinea & "|" & "#" & Trim(RsAux!Codigo) & ";" & Trim(RsAux!Texto)
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    With vsContenedor
        .ColComboList(1) = sContenedor
        .ColComboList(2) = sLinea
    End With
    Screen.MousePointer = 0
End Sub

Private Sub CargoContenedores()
Dim rsCont As rdoResultset
On Error GoTo ErrCC
    vsContenedor.Rows = 1
    vsContenedor.ColComboList(1) = sContenedor
    vsContenedor.ColComboList(2) = sLinea
    Cons = "Select * From EmbarqueContenedor Where ECoEmbarque = " & RsEmbarque!EmbID
    Set rsCont = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsCont.EOF
        With vsContenedor
            .AddItem Trim(rsCont!ECoCantidad)
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsCont!ECoContenedor)
            If Not IsNull(rsCont!ECoLinea) Then .Cell(flexcpText, .Rows - 1, 2) = Trim(rsCont!ECoLinea): embarque.LineaAsignada = rsCont("EcoLinea")
        End With
        rsCont.MoveNext
    Loop
    rsCont.Close
    vsContenedor.AddItem ""
    Exit Sub
ErrCC:
    MsgBox "Ocurrió el siguiente error al cargar los contenedores.", vbCritical, "ATENCIÓN"
End Sub

Private Sub vsContenedor_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    'Despues que edito.
    Select Case Col
        Case 0
            If vsContenedor.Cell(flexcpText, Row, 1) <> "" And Not IsNumeric(vsContenedor.Cell(flexcpText, Row, 0)) Then
                vsContenedor.RemoveItem Row
            Else
                vsContenedor.Col = Col + 1
            End If
        Case 1
            vsContenedor.Col = Col + 1
        Case 2
            If vsContenedor.Rows - 1 = Row And vsContenedor.Cell(flexcpText, Row, 0) <> "" And vsContenedor.Cell(flexcpText, Row, 1) <> "" Then
                vsContenedor.AddItem ""
            End If
    End Select
End Sub

Private Sub vsContenedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsContenedor.Editable = False Then Exit Sub
    If KeyCode = vbKeyDelete Then
        If vsContenedor.Cell(flexcpText, vsContenedor.Row, 0) <> "" Or vsContenedor.Cell(flexcpText, vsContenedor.Row, 1) <> "" Then vsContenedor.RemoveItem vsContenedor.Row
        If vsContenedor.Rows = 1 Then vsContenedor.AddItem ""
    End If
End Sub

Private Sub vsContenedor_LostFocus()
    If vsContenedor.Editable = False Then Exit Sub
    Dim lPos As Long
    lPos = 1
    Do While lPos <= vsContenedor.Rows - 1
        If (vsContenedor.Cell(flexcpText, lPos, 0) = "" And Trim(vsContenedor.Cell(flexcpText, lPos, 1)) <> "") Or _
            (vsContenedor.Cell(flexcpText, lPos, 0) <> "" And vsContenedor.Cell(flexcpText, lPos, 1) = "") Then
            vsContenedor.RemoveItem lPos
        Else
            lPos = lPos + 1
        End If
    Loop
    
    Dim lPos1 As Long
    For lPos = 1 To vsContenedor.Rows - 1
        For lPos1 = lPos + 1 To vsContenedor.Rows - 1
            If vsContenedor.Cell(flexcpText, lPos, 1) = vsContenedor.Cell(flexcpText, lPos1, 1) Then
                MsgBox "El tipo de contenedor " & vsContenedor.Cell(flexcpTextDisplay, lPos, 1) & " está duplicado.", vbInformation, "ATENCIÓN"
                vsContenedor.SetFocus
                Exit Sub
            End If
        Next lPos1
    Next
    
    If vsContenedor.Rows = 1 Then vsContenedor.AddItem ""
    
End Sub

Private Sub vsContenedor_Validate(Cancel As Boolean)
    If embarque.LineaAsignada = 0 Then
        ValidoAsignarFlete
    End If
End Sub

Private Sub vsContenedor_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Col = 0 Then
        If vsContenedor.EditText <> "" Then
            If Not IsNumeric(vsContenedor.EditText) Then
                vsContenedor.EditText = ""
                Cancel = True
            End If
        End If
    End If
End Sub

Private Function fnc_GetPrecioContenedor(ByVal iRowGrilla As Integer, ByVal sCons As String) As Currency
Dim rsCont As rdoResultset
    fnc_GetPrecioContenedor = 0
    Set rsCont = cBase.OpenResultset(sCons, rdOpenDynamic, rdConcurValues)
    Dim iQ As Integer
    Do While Not rsCont.EOF
        If IsNull(rsCont!FEMFAPartir) Then
            If MsgBox("Existe un flete sin fecha, el mismo tiene un precio de : " & Format(rsCont!FEmImporte, "#,##0.00") & " para el contenedor " & vsContenedor.Cell(flexcpTextDisplay, iRowGrilla, 1) & vbCr & "¿Desea tomar este precio?", vbQuestion + vbYesNo, "Flete sin fecha a partir") = vbYes Then
                fnc_GetPrecioContenedor = (rsCont!FEmImporte * Val(vsContenedor.Cell(flexcpText, iRowGrilla, 0)))
                Exit Do
            End If
        Else
            fnc_GetPrecioContenedor = (rsCont!FEmImporte * Val(vsContenedor.Cell(flexcpText, iRowGrilla, 0)))
        End If
        rsCont.MoveNext
    Loop
    rsCont.Close
End Function

'Private Sub ConsultoPrecioContenedor()
'Dim lPos As Long
'Dim cSuma As Currency, cAux As Currency
'Dim bSinPrecio As Boolean
'
'    If cAgencia.ListIndex = -1 Or cDestino.ListIndex = -1 Or cOrigen.ListIndex = -1 Then
'        Exit Sub
'    End If
'    bSinPrecio = False
'    For lPos = 1 To vsContenedor.Rows - 1
'
'        If Val(vsContenedor.Cell(flexcpText, lPos, 0)) > 0 Then
'            Cons = "SELECT TOP 1 FEmImporte, FEMFAPartir From FleteEmbarque " _
'                & " Where FEmOrigen = " & cOrigen.ItemData(cOrigen.ListIndex) _
'                & " And FEmDestino = " & cDestino.ItemData(cDestino.ListIndex) _
'                & " And FEmAgencia = " & cAgencia.ItemData(cAgencia.ListIndex) _
'                & " And FEmContenedor = " & vsContenedor.Cell(flexcpText, lPos, 1) _
'                & " And (FEmFAPartir = (Select MAX(FEmFAPartir) From FleteEmbarque " _
'                        & " Where FEmFAPartir < '" & Format(CDate(tEmbarco.Text) + 1, "mm/dd/yyyy") & "'" _
'                        & " And FEmFHasta >= '" & Format(CDate(tEmbarco.Text), "mm/dd/yyyy") & "'" _
'                        & " And FEmOrigen = " & cOrigen.ItemData(cOrigen.ListIndex) _
'                        & " And FEmDestino = " & cDestino.ItemData(cDestino.ListIndex) _
'                        & " And FEmAgencia = " & cAgencia.ItemData(cAgencia.ListIndex) _
'                        & " And FEmContenedor = " & vsContenedor.Cell(flexcpText, lPos, 1)
'
'                If Val(vsContenedor.Cell(flexcpText, lPos, 2)) > 0 Then Cons = Cons & " And FEmLinea = " & vsContenedor.Cell(flexcpText, lPos, 2)
'                Cons = Cons & ") or FEMFAPartir Is Null)"
'
'            If Val(vsContenedor.Cell(flexcpText, lPos, 2)) > 0 Then Cons = Cons & " And FEmLinea = " & vsContenedor.Cell(flexcpText, lPos, 2)
'            Cons = Cons & " ORDER BY FEmImporte DESC"
'
'            cAux = fnc_GetPrecioContenedor(lPos, Cons)
'            If Not bSinPrecio Then bSinPrecio = (cAux = 0)
'            cSuma = cSuma + cAux
'        End If
'    Next
'
'    If bSinPrecio Then
'        MsgBox "Para algún contenedor no se encontró precio.", vbExclamation, "ATENCIÓN"
'    End If
'
'    '12/11/2013 en lugar de dar el precio abro lista de ayuda y le dejo elegir al usuario.
'
'
'
'    If IsNumeric(tFlete.Text) Then
'        If cSuma <> CCur(tFlete.Text) Then
'            If Trim(tFlete.Text) <> "" Then
'                If MsgBox("La suma de los contenedores da " & cSuma & " y este importe es distinto al valor del flete." _
'                    & vbCrLf & "¿Desea modificar este valor?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
'                        tFlete.Text = Format(cSuma, "#,#00.00")
'                End If
'            Else
'                tFlete.Text = Format(cSuma, "#,#00.00")
'            End If
'        End If
'    Else
'        tFlete.Text = Format(cSuma, "#,#00.00")
'    End If
'
'End Sub

Private Sub GuardoContenedores(ByVal idEmb As Long)
Dim rsCont As rdoResultset
Dim lPos As Long

    Cons = "Delete EmbarqueContenedor Where ECoEmbarque = " & idEmb
    cBase.Execute (Cons)
    
    Cons = "Select * From EmbarqueContenedor Where ECoEmbarque = " & idEmb
    Set rsCont = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    For lPos = 1 To vsContenedor.Rows - 1
        If Val(vsContenedor.Cell(flexcpText, lPos, 0)) > 0 And Val(vsContenedor.Cell(flexcpText, lPos, 1)) > 0 Then
            rsCont.AddNew
            rsCont!ECoEmbarque = idEmb
            rsCont!ECoContenedor = Val(vsContenedor.Cell(flexcpText, lPos, 1))
            rsCont!ECoCantidad = Val(vsContenedor.Cell(flexcpText, lPos, 0))
            If Val(vsContenedor.Cell(flexcpText, lPos, 2)) > 0 Then rsCont!ECoLinea = Val(vsContenedor.Cell(flexcpText, lPos, 2))
            rsCont.Update
        End If
    Next lPos
    rsCont.Close
    
End Sub

Private Function DifCambioMesAnterior() As Boolean
On Error Resume Next
Dim rsDC As rdoResultset
    DifCambioMesAnterior = False
    'Cons = "Select top 1 * From Compra Where ComFecha = '" & Format(UltimoDia(DateAdd("m", -1, Date)), "mm/dd/yyyy") & "'" _
        & " And ComDCDe Is Not Null And ComSerie = 'DC'"
    Cons = "Select top 1 * From Compra Where ComFecha = '" & Format(UltimoDia(DateAdd("m", -1, Date)), "mm/dd/yyyy") & "'" _
        & " And ComNumero = 'DC" & Format(UltimoDia(DateAdd("m", -1, Date)), "yyyymm") & "'"
    Set rsDC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsDC.EOF Then
        DifCambioMesAnterior = True
    End If
    rsDC.Close
    
    If Not DifCambioMesAnterior Then
        'Tengo que validar que realmente el crédito sea de meses anteriores.
        Cons = "Select Min(ComFecha) From GastoImportacion, Compra Where GImIDSubRubro = " & paSubrubroDivisa _
            & " And GImNivelFolder = " & Folder.cFEmbarque _
            & " And GImFolder = " & RsEmbarque!EmbID _
            & " And GImIDCompra = ComCodigo And ComMoneda = " & paMonedaDolar
        Set rsDC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsDC.EOF Then
            If Not IsNull(rsDC(0)) Then
                If PrimerDia(Date) <= rsDC(0) Then
                    DifCambioMesAnterior = True
                End If
            End If
        End If
        rsDC.Close
    End If
    
End Function

Private Function SeHaceNotaoCredito() As Boolean
On Error GoTo errHNM
Dim rsHN As rdoResultset
Dim cValor As Currency

    'Si hay gasto ingresado para el embarque y el nuevo valor de la divisa
    'es menor al viejo -----> hace nota.
    SeHaceNotaoCredito = False
    Cons = "Select Sum(GimImporte) From GastoImportacion, Compra Where GImIDSubRubro = " & paSubrubroDivisa _
        & " And GImNivelFolder = " & Folder.cFEmbarque _
        & " And GImFolder = " & RsEmbarque!EmbID _
        & " And GImIDCompra = ComCodigo And ComMoneda = " & paMonedaDolar
        
    Set rsHN = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not IsNull(rsHN(0)) Then
        
        If Trim(tArbitraje.Text) = "" Or tArbitraje.Text = "0" Then tArbitraje.Text = "1"
        
        cValor = CCur(Format(rsHN(0), "#,##0.0000")) - CCur(Format(CCur(tDivisa.Text) / CCur(tArbitraje.Text), "#,##0.000000"))
        
        If cValor <> 0 Then SeHaceNotaoCredito = True
    
    End If
    rsHN.Close

errHNM:

End Function


Private Function HaceNotaModificacion() As Boolean
On Error GoTo errHNM
Dim rsHN As rdoResultset
Dim cValor As Currency

    'Si hay gasto ingresado para el embarque y el nuevo valor de la divisa
    'es menor al viejo -----> hace nota.
    HaceNotaModificacion = False
    Cons = "Select Sum(GimImporte) From GastoImportacion, Compra Where GImIDSubRubro = " & paSubrubroDivisa _
        & " And GImNivelFolder = " & Folder.cFEmbarque _
        & " And GImFolder = " & RsEmbarque!EmbID _
        & " And GImIDCompra = ComCodigo And ComMoneda = " & paMonedaDolar
        
    Set rsHN = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not IsNull(rsHN(0)) Then
        
        If Trim(tArbitraje.Text) = "" Or tArbitraje.Text = "0" Then tArbitraje.Text = "1"
        
        cValor = CCur(Format(rsHN(0), "#,##0.0000")) - CCur(Format(CCur(tDivisa.Text) / CCur(tArbitraje.Text), "#,##0.000000"))
        
        If cValor > 0 Then HaceNotaModificacion = True
    
    End If
    rsHN.Close

errHNM:
End Function

Private Sub s_ListaGastosDivisa()
On Error GoTo errLGD
    
    If RsEmbarque.EOF Then Exit Sub
    Cons = "Select ComCodigo as Código, ComFecha as 'Fecha', ComSerie as 'Serie', ComNumero as Número, GImImporte as Importe, ComTC as 'TC', IsNull(ComComentario, '') as Comentario " _
        & " From GastoImportacion, Compra Where GImIDSubRubro = " & paSubrubroDivisa _
        & " And GImNivelFolder = " & Folder.cFEmbarque _
        & " And GImFolder = " & RsEmbarque!EmbID _
        & " And GImIDCompra = ComCodigo And ComMoneda = " & paMonedaDolar
        
    Dim objH As New clsListadeAyuda
    objH.ActivarAyuda cBase, Cons, 7500, 0, "Gastos de Divisa"
    Set objH = Nothing
    Exit Sub
    
errLGD:
    clsGeneral.OcurrioError "Error al cargar la lista de gastos.", Err.Description
End Sub

Private Function fnc_DoTestLogin(Optional ByVal bSave As Boolean = False) As Boolean
    Screen.MousePointer = 11
    Dim sRet As String
    If Not fnc_ValidoAcceso() Then
        fnc_DoTestLogin = False
        Status.Panels("zureo").Text = "ZUREO OFF"
        If Not bSave Then
            MsgBox "El programa no logró hacer el login en zureo.", vbExclamation, "Acceso a Zureo"
        End If
    Else
        fnc_DoTestLogin = True
        Status.Panels("zureo").Text = "ZUREO ON"
    End If
'    Set objUsers = Nothing
    Screen.MousePointer = 0
End Function

'Private Sub ConsultoOpcionMasBarataEnFlete()
'On Error GoTo errCF
'
'    If embarque.Embarco < Date And embarque.Embarco > DateMinValue Then Exit Sub
'    If cAgencia.ListIndex < 0 Or cOrigen.ListIndex < 0 Or cDestino.ListIndex < 0 Then Exit Sub
'
'Dim vOpcionActual As Currency
'Dim vOtraOpcion As Currency
'Dim lineaActual As Long
'Dim lPos As Integer
'
'    'Busco la línea que tienen los contenedores.
'    For lPos = 1 To vsContenedor.Rows - 1
'        If Val(vsContenedor.Cell(flexcpText, lPos, 0)) > 0 Then
'            If Val(vsContenedor.Cell(flexcpText, lPos, 2)) > 0 Then lineaActual = Val(vsContenedor.Cell(flexcpText, lPos, 2)): Exit For
'        End If
'    Next
'
'    If lineaActual = 0 Then Exit Sub
'    If cAgencia.ItemData(cAgencia.ListIndex) = embarque.Agencia And lineaActual = embarque.LineaAsignada Then
'        Exit Sub
'    End If
'
'
'    Dim sContenedores As String
'    For lPos = 1 To vsContenedor.Rows - 1
'        If Val(vsContenedor.Cell(flexcpText, lPos, 0)) > 0 Then
'            If Val(vsContenedor.Cell(flexcpText, lPos, 1)) > 0 Then sContenedores = sContenedores & IIf(sContenedores <> "", ",", "") & Val(vsContenedor.Cell(flexcpText, lPos, 1))
'        End If
'    Next
'
'Dim rsCont As rdoResultset
'
'    ReDim vLineasPosibles(0) As Long
'    Dim fEmbarque As Date
'    fEmbarque = DateMinValue
'    If IsDate(tEmbarco.Text) Then fEmbarque = CDate(tEmbarco.Text)
'    If IsDate(tEmbPrevisto.Text) Then fEmbarque = CDate(tEmbPrevisto.Text)
'    If fEmbarque = DateMinValue Then Exit Sub
'
'    'Para las condiciones cargo todas las líneas que tienen precio.
'    Cons = "SELECT DISTINCT(FEmLinea) FROM FleteEmbarque " _
'        & " WHERE FEmOrigen = " & cOrigen.ItemData(cOrigen.ListIndex) _
'        & " AND FEmDestino = " & cDestino.ItemData(cDestino.ListIndex) _
'        & " And FEmAgencia = " & cAgencia.ItemData(cAgencia.ListIndex) _
'        & " And FEmContenedor IN( " & sContenedores & ")" _
'        & " AND FEmLinea <> " & lineaActual _
'        & " And FEmFAPartir = (Select MAX(FEmFAPartir) From FleteEmbarque " _
'                & " WHERE FEmFAPartir < '" & Format(fEmbarque + 1, "mm/dd/yyyy") & "'" _
'                & " And FEmFHasta >= '" & Format(fEmbarque, "mm/dd/yyyy") & "'" _
'                & " And FEmOrigen = " & cOrigen.ItemData(cOrigen.ListIndex) _
'                & " And FEmDestino = " & cDestino.ItemData(cDestino.ListIndex) _
'                & " And FEmAgencia = " & cAgencia.ItemData(cAgencia.ListIndex) _
'                & " And FEmContenedor IN( " & sContenedores & "))"
'
'    Set rsCont = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'    Do While Not rsCont.EOF
'        If vLineasPosibles(0) = 0 Then
'            vLineasPosibles(0) = rsCont(0)
'        Else
'            ReDim Preserve vLineasPosibles(UBound(vLineasPosibles) + 1)
'            vLineasPosibles(UBound(vLineasPosibles)) = rsCont(0)
'        End If
'        rsCont.MoveNext
'    Loop
'    rsCont.Close
'
'    'Si no tengo línea me voy.
'    If vLineasPosibles(0) = 0 Then Exit Sub
'
'    Dim idMasBarato As Long
'    Dim importeMasBarato As Currency
'    importeMasBarato = PrecioParaUnaLinea(lineaActual, fEmbarque)
'
'    Dim cSumaAct As Currency
'    Dim iLinea As Integer
'    For iLinea = 0 To UBound(vLineasPosibles)
'        cSumaAct = PrecioParaUnaLinea(vLineasPosibles(iLinea), fEmbarque)
'        If cSumaAct < importeMasBarato And cSumaAct > 0 Then
'            idMasBarato = vLineasPosibles(iLinea)
'            importeMasBarato = cSumaAct
'        End If
'    Next
'
'    If idMasBarato > 0 Then
'
'        Cons = "SELECT Texto FROM CodigoTexto WHERE Codigo = " & idMasBarato
'        Set rsCont = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
'        If Not rsCont.EOF Then
'            Cons = Trim(rsCont("Texto"))
'        End If
'        rsCont.Close
'
'        MsgBox "ATENCIÓN!!!" & vbCrLf & vbCrLf & "Existe una opción más barata con la línea " & Cons & " a un costo de " & importeMasBarato
'    End If
'    Exit Sub
'errCF:
'End Sub

Private Function PrecioParaUnaLinea(ByVal idLinea As Long, ByVal fEmbarque As Date) As Currency
On Error GoTo errPU
Dim lPos As Integer
Dim cSuma As Currency
Dim rsCont As rdoResultset

    For lPos = 1 To vsContenedor.Rows - 1
        If Val(vsContenedor.Cell(flexcpText, lPos, 0)) > 0 Then
            Cons = "SELECT TOP 1 FEmImporte FROM FleteEmbarque " _
                & " WHERE FEmOrigen = " & cOrigen.ItemData(cOrigen.ListIndex) _
                & " AND FEmDestino = " & cDestino.ItemData(cDestino.ListIndex) _
                & " And FEmAgencia = " & cAgencia.ItemData(cAgencia.ListIndex) _
                & " And FEmContenedor = " & vsContenedor.Cell(flexcpText, lPos, 1) _
                & " And (FEmFAPartir = (Select MAX(FEmFAPartir) From FleteEmbarque " _
                        & " WHERE FEmFAPartir < '" & Format(fEmbarque + 1, "mm/dd/yyyy") & "'" _
                        & " And FEmFHasta >= '" & Format(fEmbarque, "mm/dd/yyyy") & "'" _
                        & " And FEmOrigen = " & cOrigen.ItemData(cOrigen.ListIndex) _
                        & " And FEmDestino = " & cDestino.ItemData(cDestino.ListIndex) _
                        & " And FEmAgencia = " & cAgencia.ItemData(cAgencia.ListIndex) _
                        & " And FEmContenedor = " & vsContenedor.Cell(flexcpText, lPos, 1) _
                        & " And FEmLinea = " & idLinea & ") or FEMFAPartir Is Null)" _
                & " And FEmLinea = " & vsContenedor.Cell(flexcpText, lPos, 2) _
                & " ORDER BY FEmImporte DESC"
            Set rsCont = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsCont.EOF Then
                cSuma = cSuma + (rsCont!FEmImporte * Val(vsContenedor.Cell(flexcpText, lPos, 0)))
            End If
            rsCont.Close
        End If
    Next
    PrecioParaUnaLinea = cSuma
    Exit Function
errPU:
End Function

Private Function ObtenerLinea() As clsCodigoNombre
    Set ObtenerLinea = New clsCodigoNombre
    Dim idLineaActual As Long
    Dim iPos As Integer
    For iPos = 1 To vsContenedor.Rows - 1
        If Trim(vsContenedor.Cell(flexcpText, iPos, 1)) <> "" And Trim(vsContenedor.Cell(flexcpText, iPos, 2)) <> "" Then
            idLineaActual = Val(vsContenedor.Cell(flexcpText, iPos, 2))
            Exit For
        End If
    Next iPos
    
    If idLineaActual > 0 Then
        Cons = "Select Codigo, Texto From CodigoTexto Where Codigo = " & idLineaActual
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            ObtenerLinea.Codigo = RsAux("Codigo")
            ObtenerLinea.Nombre = Trim(RsAux("Texto"))
        End If
        RsAux.Close
    End If
    
End Function

Private Function ObtenerContenedores() As Collection
    
    Set ObtenerContenedores = New Collection
    Dim oCont As clsContenedoresEmbarque
    
    Dim iPos As Integer
    For iPos = 1 To vsContenedor.Rows - 1
        If Trim(vsContenedor.Cell(flexcpText, iPos, 1)) <> "" And Trim(vsContenedor.Cell(flexcpText, iPos, 2)) <> "" Then
            Cons = "SELECT ConCodigo, ConNombre From Contenedor WHERE ConCodigo = " & Val(vsContenedor.Cell(flexcpText, iPos, 1))
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                Set oCont = New clsContenedoresEmbarque
                Set oCont.Contenedor = New clsCodigoNombre
                oCont.Contenedor.Codigo = RsAux("ConCodigo")
                oCont.Contenedor.Nombre = RsAux("ConNombre")
                oCont.Cantidad = Val(vsContenedor.Cell(flexcpValue, iPos, 0))
                ObtenerContenedores.Add oCont
            End If
            RsAux.Close
        End If
    Next iPos
        
End Function

Private Function PuedoAbrirAsignacionFlete() As Boolean
    
    Dim oLinea As clsCodigoNombre
    Set oLinea = ObtenerLinea
        
    If (oLinea.Codigo = 0 Or cOrigen.ListIndex = -1 Or cAgencia.ListIndex = -1) Or (Not IsDate(tEmbarco.Text) And Not IsDate(tEmbPrevisto.Text)) Then
        PuedoAbrirAsignacionFlete = False
    Else
        PuedoAbrirAsignacionFlete = True
    End If
    
End Function

Private Sub InvocoAsignacionFlete()
On Error GoTo errIAF
    
    If cAgencia.ListIndex = -1 Then
        MsgBox "Falta agencia", vbInformation, "Datos precio flete"
        Exit Sub
    End If
    
    If cOrigen.ListIndex = -1 Then
        MsgBox "Falta origen.", vbInformation, "Datos precio flete"
        Exit Sub
    End If
       
    Dim oLinea As clsCodigoNombre
    Set oLinea = ObtenerLinea
    
    If oLinea.Codigo = 0 Then
        MsgBox "Falta la línea.", vbInformation, "Datos precio flete"
        Exit Sub
    End If
    
    Dim oDE As New clsDatosPrecioFlete
    Set oDE.Origen = New clsCodigoNombre
    oDE.Origen.Codigo = cOrigen.ItemData(cOrigen.ListIndex)
    oDE.Origen.Nombre = cOrigen.Text
    
    Set oDE.Agencia = New clsCodigoNombre
    oDE.Agencia.Codigo = cAgencia.ItemData(cAgencia.ListIndex)
    oDE.Agencia.Nombre = cAgencia.Text
    
    Set oDE.Destino = New clsCodigoNombre
    oDE.Destino.Codigo = cDestino.ItemData(cDestino.ListIndex)
    oDE.Destino.Nombre = cDestino.Text
    
    Set oDE.Linea = oLinea
    If IsDate(tEmbarco.Text) Then
        oDE.SiEmbarco = True
        oDE.FechaEmbarque = tEmbarco.Text
    ElseIf IsDate(tEmbPrevisto.Text) Then
        oDE.FechaEmbarque = tEmbPrevisto.Text
    Else
        MsgBox "Falta fecha de embarque previsto o la fecha que embarcó", vbExclamation, "Datos precio flete"
        Exit Sub
    End If
    
    Set oDE.Contenedores = ObtenerContenedores
    With New frmPreciosFlete
        Set .DatosEmbarque = oDE
        .Show vbModal, Me
        If .DialogResult Then
            tFlete.Text = Format(.PrecioSeleccionado, "#,##0.00")
            If oDE.SiEmbarco Then
                lblInfoFlete.ForeColor = vbBlack
            Else
                lblInfoFlete.ForeColor = ColorNaranja
            End If
        End If
    End With
    Exit Sub
errIAF:
    clsGeneral.OcurrioError "Error al asignar el precio del flete.", Err.Description
End Sub

Private Function ValidoAsignarFlete() As Boolean
   
    'Nomina Agencia veo si estoy en condiciones de presentarle el precio de flete.
    If PuedoAbrirAsignacionFlete Then
        InvocoAsignacionFlete
    End If
   
End Function

Private Function EmbarqueEnCalendario() As Boolean
On Error GoTo errEC
Dim rsE As rdoResultset
    If embarque.ID = 0 Then Exit Function
    Set rsE = cBase.OpenResultset("SELECT CDeEmbarque FROM CalendarioDescarga WHERE CDeEmbarque = " & embarque.ID, rdOpenForwardOnly, rdConcurValues)
    EmbarqueEnCalendario = Not rsE.EOF
    rsE.Close
    Exit Function
errEC:
End Function

Private Sub CambioFechaEnCalendario(ByVal Dias As Integer)

    Cons = "SELECT * FROM CalendarioDescarga WHERE CDeFecha IS NOT NULL AND CDeEstado = 0 AND CDeEmbarque = " & embarque.ID
    Dim rsE As rdoResultset
    Set rsE = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsE.EOF
        rsE.Edit
        rsE("CDeFecha") = rsE("CDeFecha") + Dias
        If Weekday(rsE("CDeFecha")) = vbSunday Then rsE("CDeFecha") = rsE("CDeFecha") + 1
        rsE.Update
        rsE.MoveNext
    Loop
    rsE.Close

End Sub
