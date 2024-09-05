VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmInGastos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Gastos de Importación"
   ClientHeight    =   6960
   ClientLeft      =   4695
   ClientTop       =   2040
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInGastos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8070
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   13
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   400
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "embarque"
            Object.ToolTipText     =   "Consulta embarques"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "sub"
            Object.ToolTipText     =   "Consulta Subcarpetas"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "gasto"
            Object.ToolTipText     =   "Consulta gastos"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "dolar"
            Object.ToolTipText     =   "Tasas de cambio"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Comprobante"
      ForeColor       =   &H00000080&
      Height          =   1755
      Left            =   60
      TabIndex        =   28
      Top             =   480
      Width           =   7935
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4140
         MaxLength       =   40
         TabIndex        =   5
         Top             =   600
         Width           =   3675
      End
      Begin VB.TextBox tID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox tTCDolar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4140
         MaxLength       =   6
         TabIndex        =   14
         Top             =   1340
         Width           =   735
      End
      Begin VB.TextBox tIOriginal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   2160
         MaxLength       =   13
         TabIndex        =   12
         Text            =   "1,000,000.00"
         Top             =   1320
         Width           =   1035
      End
      Begin VB.TextBox tNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4140
         MaxLength       =   9
         TabIndex        =   9
         Top             =   980
         Width           =   1215
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   1320
         Width           =   795
         _ExtentX        =   1402
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
      Begin AACombo99.AACombo cComprobante 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   960
         Width           =   1875
         _ExtentX        =   3307
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
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "I&d Compra:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lTC 
         BackStyle       =   0  'Transparent
         Caption         =   "21/08/2000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4920
         TabIndex        =   31
         Top             =   1365
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "T/C:"
         Height          =   255
         Left            =   3300
         TabIndex        =   13
         Top             =   1365
         Width           =   795
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe BRUTO:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1365
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Comprobante:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   3300
         TabIndex        =   4
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha &Gasto:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   255
         Left            =   3300
         TabIndex        =   8
         Top             =   1020
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Distribución de Gastos"
      ForeColor       =   &H00000080&
      Height          =   4380
      Left            =   60
      TabIndex        =   27
      Top             =   2295
      Width           =   7935
      Begin VB.TextBox tAutoriza 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   80
         TabIndex        =   24
         Top             =   4020
         Width           =   2055
      End
      Begin VB.CheckBox chVerificado 
         Appearance      =   0  'Flat
         Caption         =   "Autorizado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3600
         TabIndex        =   25
         Top             =   4065
         Width           =   1395
      End
      Begin VB.CommandButton bDistribuir 
         Caption         =   "&Distribuir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3720
         TabIndex        =   19
         Top             =   1500
         Width           =   975
      End
      Begin VB.CommandButton bAgregar 
         Caption         =   "&Agregar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   18
         Top             =   1500
         Width           =   975
      End
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         MaxLength       =   80
         TabIndex        =   22
         Top             =   3680
         Width           =   6675
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsFolder 
         Height          =   1515
         Left            =   120
         TabIndex        =   20
         Top             =   2100
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   2672
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
         FocusRect       =   2
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
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
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin AACombo99.AACombo cFolder 
         Height          =   315
         Left            =   1080
         TabIndex        =   17
         Top             =   1500
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VSFlex6DAOCtl.vsFlexGrid vsGasto 
         Height          =   1125
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   1984
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
         FocusRect       =   2
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
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
      Begin VB.Label lAutoriza 
         BackStyle       =   0  'Transparent
         Caption         =   "Autoriza:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   4020
         Width           =   975
      End
      Begin VB.Label lArticulos 
         Caption         =   "articulos de la carpeta"
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   120
         TabIndex        =   30
         Top             =   1860
         Width           =   7635
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "C&arpetas:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentarios:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3720
         Width           =   975
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   6705
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
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
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6033
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   7680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInGastos.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInGastos.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInGastos.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInGastos.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInGastos.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInGastos.frx":0A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInGastos.frx":0D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInGastos.frx":0E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInGastos.frx":11B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInGastos.frx":14CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInGastos.frx":17E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInGastos.frx":1B00
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInGastos.frx":1E1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInGastos.frx":2134
            Key             =   ""
         EndProperty
      EndProperty
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
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuGrabar 
         Caption         =   "&Grabar"
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuCancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuExit 
         Caption         =   "Del formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmInGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sModificar As Boolean
Dim gIdComprobante As Long
Dim gFModificacion As Date

Dim RsCom As rdoResultset

Dim aTexto As String

Private Sub bAgregar_Click()
    If cFolder.ListIndex <> -1 Then AgregoFolderALista
End Sub

Private Sub AgregoFolderALista()

Dim aTipoFolder As Integer, aFolder As Long
Dim aValor As Long

    On Error GoTo Error
    aTipoFolder = Mid(cFolder.ItemData(cFolder.ListIndex), 1, 1)
    aFolder = Mid(cFolder.ItemData(cFolder.ListIndex), 2, Len(CStr(cFolder.ItemData(cFolder.ListIndex))))
    
    'Verifico si el folder está en la lista----------------------------------------
    For I = 1 To vsFolder.Rows - 1
        If vsFolder.Cell(flexcpData, I, 0) = cFolder.ItemData(cFolder.ListIndex) Then
            MsgBox "La carpeta seleccionada ya está ingresada. Verifique la lista de carpetas.", vbInformation, "ATENCIÓN"
            Exit Sub
        End If
    Next
    
    Select Case aTipoFolder  'Tipo de Folder
        
        Case Folder.cFCarpeta            'Folder Madre -->Gastos solo del Nivel Madre
            For I = 1 To vsGasto.Rows - 1
                If vsGasto.Cell(flexcpData, I, 1) = aTipoFolder Then
                    
                    With vsFolder
                        .AddItem ""
                        
                        .Cell(flexcpText, .Rows - 1, 0) = cFolder.Text     'Folder
                        aValor = cFolder.ItemData(cFolder.ListIndex): .Cell(flexcpData, .Rows - 1, 0) = aValor
                        
                        .Cell(flexcpText, .Rows - 1, 1) = Trim(vsGasto.Cell(flexcpText, I, 0))                                  'Gasto
                        .Cell(flexcpData, .Rows - 1, 1) = vsGasto.Cell(flexcpData, I, 0)
                        
                        .Cell(flexcpBackColor, .Rows - 1, 3) = Obligatorio
                    End With
                    
                End If
            Next
                    
        Case Folder.cFEmbarque              'Folder Embarque Embarque ---> Gastos Nivel 2 y 3
            For I = 1 To vsGasto.Rows - 1
                If vsGasto.Cell(flexcpData, I, 1) = aTipoFolder Or vsGasto.Cell(flexcpData, I, 1) = Folder.cFSubCarpeta Then
                    
                    With vsFolder
                        .AddItem ""
                        
                        .Cell(flexcpText, .Rows - 1, 0) = cFolder.Text     'Folder
                        aValor = cFolder.ItemData(cFolder.ListIndex): .Cell(flexcpData, .Rows - 1, 0) = aValor
                        
                        .Cell(flexcpText, .Rows - 1, 1) = Trim(vsGasto.Cell(flexcpText, I, 0))                                  'Gasto
                        .Cell(flexcpData, .Rows - 1, 1) = vsGasto.Cell(flexcpData, I, 0)
                        
                        .Cell(flexcpBackColor, .Rows - 1, 3) = Obligatorio
                    End With
                    
                End If
            Next I
        
        
        Case Folder.cFSubCarpeta                  'Folder Subcarpeta ---> Gastos Nivel 3
            For I = 1 To vsGasto.Rows - 1
                If vsGasto.Cell(flexcpData, I, 1) = aTipoFolder Then
                    
                    With vsFolder
                        .AddItem ""
                        
                        .Cell(flexcpText, .Rows - 1, 0) = cFolder.Text     'Folder
                        aValor = cFolder.ItemData(cFolder.ListIndex): .Cell(flexcpData, .Rows - 1, 0) = aValor
                        
                        .Cell(flexcpText, .Rows - 1, 1) = Trim(vsGasto.Cell(flexcpText, I, 0))                                  'Gasto
                        .Cell(flexcpData, .Rows - 1, 1) = vsGasto.Cell(flexcpData, I, 0)
                        
                        .Cell(flexcpBackColor, .Rows - 1, 3) = Obligatorio
                    End With
                    
                End If
            Next I
            
        End Select
    cFolder.Text = "": lArticulos.Caption = ""
    Exit Sub

Error:
    MsgBox "El folder seleccionado ya ha sido ingresado.", vbInformation, "ATENCIÓN"
End Sub


Private Sub bDistribuir_Click()

Dim sHay As Boolean  'Hay distribución hecha
    
    If vsFolder.Rows = 1 Then Exit Sub
    DistribuirGasto
    
    vsFolder.SetFocus

End Sub

Private Sub DistribuirGasto()

Dim iGasto As Integer   'indice para recorrer lista gasto

    On Error Resume Next
    Screen.MousePointer = 11
    For iGasto = 1 To vsGasto.Rows - 1
        Select Case vsGasto.Cell(flexcpData, iGasto, 2)     'Tipo de Distribucion
            
            Case Distribucion.Lineal: DistribuirGastoLineal vsGasto.Cell(flexcpData, iGasto, 0), vsGasto.Cell(flexcpValue, iGasto, 3)
            
            Case Distribucion.Divisa: DistribuirGastoDivisa vsGasto.Cell(flexcpData, iGasto, 0), vsGasto.Cell(flexcpValue, iGasto, 3)
                            
            Case Distribucion.Volumen:  DistribuirGastoVolumen vsGasto.Cell(flexcpData, iGasto, 0), vsGasto.Cell(flexcpValue, iGasto, 3)
        
        End Select
    Next
    Screen.MousePointer = 0
    
End Sub

Private Sub DistribuirGastoLineal(CodigoGasto As Long, Importe As Currency)

Dim aTipoFolder As Integer, aFolder As Long
Dim aCantidad As Long
Dim aIndice As Integer, aSuma As Currency   'Los uso para que las cuentas en la dist. de justa.
    
    On Error GoTo errDistribuir
    aCantidad = 0: aSuma = 0
    
    With vsFolder
    
    For I = 1 To .Rows - 1
        If .Cell(flexcpData, I, 1) = CodigoGasto Then  'Si coincide el codigo de gasto
            'Busco la cantidad de articulos en el folder
            
            aTipoFolder = Mid(.Cell(flexcpData, I, 0), 1, 1)
            aFolder = Mid(.Cell(flexcpData, I, 0), 2, Len(CStr(.Cell(flexcpData, I, 0))))
            
            Select Case aTipoFolder
                
                Case Folder.cFCarpeta                           'CARPETA-------------------------------------------------------------
                    cons = "Select Sum(AFoCantidad) from ArticuloFolder" _
                            & " Where AFoTipo = " & Folder.cFEmbarque _
                            & " And AFoCodigo In ( Select EmbID from Embarque Where EmbCarpeta = " & aFolder & ")"
                    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    .Cell(flexcpText, I, 2) = rsAux(0)
                    aCantidad = aCantidad + rsAux(0)
                    rsAux.Close
            
                Case Folder.cFEmbarque                      'EMBARQUE-------------------------------------------------------------
                    cons = "Select Sum(AFoCantidad) from ArticuloFolder" _
                            & " Where AFoTipo = " & Folder.cFEmbarque _
                            & " And AFoCodigo = " & aFolder
                    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    .Cell(flexcpText, I, 2) = rsAux(0)
                    aCantidad = aCantidad + rsAux(0)
                    rsAux.Close
                
                Case Folder.cFSubCarpeta                    'SUBCARPETA-------------------------------------------------------------
                    cons = "Select Sum(AFoCantidad) from ArticuloFolder" _
                            & " Where AFoTipo = " & Folder.cFSubCarpeta _
                            & " And AFoCodigo = " & aFolder
                    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    .Cell(flexcpText, I, 2) = rsAux(0)
                    aCantidad = aCantidad + rsAux(0)
                    rsAux.Close
            End Select
        End If
    Next
    
    'Realizo la distribucion en base a las cantidades obtenidas
    If aCantidad > 0 Then
        For I = 1 To .Rows - 1
            'Si coincide el codigo de gasto
            If .Cell(flexcpData, I, 1) = CodigoGasto Then
                .Cell(flexcpText, I, 3) = Format(.Cell(flexcpValue, I, 2) * Importe / aCantidad, "##,##0.00")
                aSuma = aSuma + .Cell(flexcpValue, I, 3)
                aIndice = I
            End If
        Next
        .Cell(flexcpText, aIndice, 3) = Format(Importe - aSuma + .Cell(flexcpValue, aIndice, 3), "##,##0.00")
    Else
        For I = 1 To .Rows - 1
            If .Cell(flexcpData, I, 1) = CodigoGasto Then .Cell(flexcpText, I, 3) = "0.00"
        Next
    End If
    
    End With
    Exit Sub

errDistribuir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al distribuir el gasto lineal.", Err.Description
End Sub

Private Sub DistribuirGastoDivisa(CodigoGasto As Long, Importe As Currency)

Dim aTipoFolder As Integer, aFolder As Long
Dim aDivisa As Currency
Dim aIndice As Integer, aSuma As Currency   'Los uso para que las cuentas en la dist. de justa.
Dim RsE As rdoResultset
    
    On Error GoTo errDistribuir
    aDivisa = 0
    'NOTA:
    '   (Q art. * P.U.) / Arbitraje
    '   El arbitraje hay que sacarlo del embarque
    
    With vsFolder
    
    For I = 1 To .Rows - 1
        If .Cell(flexcpData, I, 1) = CodigoGasto Then  'Si coincide el codigo de gasto
            'Busco la cantidad de articulos en el folder
            aSuma = 0
            aTipoFolder = Mid(.Cell(flexcpData, I, 0), 1, 1)
            aFolder = Mid(.Cell(flexcpData, I, 0), 2, Len(CStr(.Cell(flexcpData, I, 0))))
            
            Select Case aTipoFolder
                         
                Case Folder.cFCarpeta                           'CARPETA-------------------------------------------------------------
                    cons = "Select AFoCodigo, Total = AFoPUnitario *AFoCantidad from ArticuloFolder" _
                            & " Where AFoTipo = " & Folder.cFEmbarque _
                            & " And AFoCodigo In ( Select EmbID from Embarque Where EmbCarpeta = " & aFolder & ")"
                    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    Do While Not rsAux.EOF
                        cons = "Select EmbArbitraje from Embarque Where EmbID = " & rsAux!AFoCodigo
                        Set RsE = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                        If Not IsNull(RsE!EmbArbitraje) Then aSuma = aSuma + (rsAux!Total / RsE!EmbArbitraje) Else aSuma = aSuma + rsAux!Total
                        RsE.Close
                        rsAux.MoveNext
                    Loop
                    rsAux.Close
                    
                    .Cell(flexcpText, I, 2) = Format(aSuma, FormatoMonedaP)
                    aDivisa = aDivisa + aSuma
            
                Case Folder.cFEmbarque                      'EMBARQUE-------------------------------------------------------------
                    cons = "Select Total = (AFoPUnitario *AFoCantidad) / EmbArbitraje from ArticuloFolder, Embarque" _
                            & " Where AFoTipo = " & Folder.cFEmbarque _
                            & " And AFoCodigo = " & aFolder _
                            & " And AFoCodigo = EmbID"
                    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    Do While Not rsAux.EOF
                        aSuma = aSuma + rsAux!Total
                        rsAux.MoveNext
                    Loop
                    rsAux.Close
                    
                    .Cell(flexcpText, I, 2) = Format(aSuma, FormatoMonedaP)
                    aDivisa = aDivisa + aSuma
                
                Case Folder.cFSubCarpeta                    'SUBCARPETA-------------------------------------------------------------
                    cons = "Select Total = (AFoPUnitario *AFoCantidad) / EmbArbitraje from ArticuloFolder, SubCarpeta, Embarque" _
                            & " Where AFoTipo = " & Folder.cFSubCarpeta _
                            & " And AFoCodigo = " & aFolder _
                            & " And AFoCodigo = SubID And SubEmbarque = EmbID"
                    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    Do While Not rsAux.EOF
                        aSuma = aSuma + rsAux!Total
                        rsAux.MoveNext
                    Loop
                    rsAux.Close
                    
                    .Cell(flexcpText, I, 2) = Format(aSuma, FormatoMonedaP)
                    aDivisa = aDivisa + aSuma
            End Select
        End If
    Next
    
    'Realizo la distribucion en base a las cantidades obtenidas
    aSuma = 0
    If aDivisa > 0 Then
        For I = 1 To .Rows - 1
            'Si coincide el codigo de gasto
            If .Cell(flexcpData, I, 1) = CodigoGasto Then
                .Cell(flexcpText, I, 3) = Format(.Cell(flexcpValue, I, 2) * Importe / aDivisa, "##,##0.00")
                aSuma = aSuma + .Cell(flexcpValue, I, 3)
                aIndice = I
            End If
        Next
        .Cell(flexcpText, aIndice, 3) = Format(Importe - aSuma + .Cell(flexcpValue, aIndice, 3), "##,##0.00")
    Else
        For I = 1 To .Rows - 1
            If .Cell(flexcpData, I, 1) = CodigoGasto Then .Cell(flexcpText, I, 3) = "0.00"
        Next
    End If
    
    End With
    Exit Sub

errDistribuir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al distribuir el gasto por la disvisa.", Err.Description
End Sub


Private Sub DistribuirGastoVolumen(CodigoGasto As Long, Importe As Currency)

Dim aTipoFolder As Integer, aFolder As Long
Dim aVolumen As Currency
Dim aIndice As Integer, aSuma As Currency   'Los uso para que las cuentas en la dist. de justa.

    On Error GoTo errDistribuir
    aVolumen = 0: aSuma = 0
    
    With vsFolder
    
    For I = 1 To .Rows - 1
        If .Cell(flexcpData, I, 1) = CodigoGasto Then  'Si coincide el codigo de gasto
            'Saco el volumen de articulos en el folder
            
            aTipoFolder = Mid(.Cell(flexcpData, I, 0), 1, 1)
            aFolder = Mid(.Cell(flexcpData, I, 0), 2, Len(CStr(.Cell(flexcpData, I, 0))))
            
            Select Case aTipoFolder
                
                Case Folder.cFCarpeta                           'CARPETA-------------------------------------------------------------
                    cons = "Select Sum(AFoCantidad * ArtVolumen) from ArticuloFolder, Articulo" _
                            & " Where AFoTipo = " & Folder.cFEmbarque _
                            & " And AFoCodigo In ( Select EmbID from Embarque Where EmbCarpeta = " & aFolder & ")" _
                            & " And AFoArticulo = ArtID"
                    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    If Not IsNull(rsAux(0)) Then .Cell(flexcpText, I, 2) = Format(rsAux(0), "#,##0.000") Else .Cell(flexcpText, I, 2) = "0.00"
                    If Not IsNull(rsAux(0)) Then aVolumen = aVolumen + rsAux(0)
                    rsAux.Close
            
                Case Folder.cFEmbarque                      'EMBARQUE-------------------------------------------------------------
                    cons = "Select Sum(AFoCantidad * ArtVolumen) from ArticuloFolder, Articulo" _
                            & " Where AFoTipo = " & Folder.cFEmbarque _
                            & " And AFoCodigo = " & aFolder _
                            & " And AFoArticulo = ArtID"
                    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    If Not IsNull(rsAux(0)) Then .Cell(flexcpText, I, 2) = Format(rsAux(0), "#,##0.000") Else .Cell(flexcpText, I, 2) = "0.00"
                    If Not IsNull(rsAux(0)) Then aVolumen = aVolumen + rsAux(0)
                    rsAux.Close
                
                Case Folder.cFSubCarpeta                    'SUBCARPETA-------------------------------------------------------------
                    cons = "Select Sum(AFoCantidad * ArtVolumen) from ArticuloFolder, Articulo" _
                            & " Where AFoTipo = " & Folder.cFSubCarpeta _
                            & " And AFoCodigo = " & aFolder _
                            & " And AFoArticulo = ArtID"
                    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    If Not IsNull(rsAux(0)) Then .Cell(flexcpText, I, 2) = Format(rsAux(0), "#,##0.000") Else .Cell(flexcpText, I, 2) = "0.000"
                    If Not IsNull(rsAux(0)) Then aVolumen = aVolumen + rsAux(0)
                    rsAux.Close
            End Select
        End If
    Next
    
    'Realizo la distribucion en base a las cantidades obtenidas
    If aVolumen > 0 Then
        For I = 1 To .Rows - 1
            'Si coincide el codigo de gasto
            If .Cell(flexcpData, I, 1) = CodigoGasto Then
                .Cell(flexcpText, I, 3) = Format(.Cell(flexcpValue, I, 2) * Importe / aVolumen, "##,##0.00")
                aSuma = aSuma + .Cell(flexcpValue, I, 3)
                aIndice = I
            End If
        Next
        .Cell(flexcpText, aIndice, 3) = Format(Importe - aSuma + .Cell(flexcpValue, aIndice, 3), "##,##0.00")
    Else
        For I = 1 To .Rows - 1
            If .Cell(flexcpData, I, 1) = CodigoGasto Then .Cell(flexcpText, I, 3) = "0.00"
        Next
    End If
    
    End With
    Exit Sub

errDistribuir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al distribuir el gasto por volumen.", Err.Description
End Sub

Private Sub cFolder_Change()
    lArticulos.Caption = ""
End Sub

Private Sub cFolder_Click()

Dim aTipoFolder As Integer, aFolder As Long
Dim aTexto As String: aTexto = ""

    On Error GoTo errArticulos
    'Cargo los articulos para el folder
    If cFolder.ListIndex = -1 Then Exit Sub
    Screen.MousePointer = 11
    aTipoFolder = Mid(cFolder.ItemData(cFolder.ListIndex), 1, 1)
    aFolder = Mid(cFolder.ItemData(cFolder.ListIndex), 2, Len(CStr(cFolder.ItemData(cFolder.ListIndex))))
    
    'Si el folder es carpeta, hay que consultar los articulos del embarque
    If aTipoFolder = Folder.cFCarpeta Then
        cons = "Select * from ArticuloFolder, Articulo " _
                & " Where AFoTipo = " & Folder.cFEmbarque _
                & " And AFoCodigo IN (Select EmbID from Embarque Where EmbCarpeta = " & aFolder & ")" _
                & " And AFoArticulo = ArtID"
    Else
        cons = "Select * from ArticuloFolder, Articulo " _
                & " Where AFoTipo = " & aTipoFolder _
                & " And AFoCodigo = " & aFolder _
                & " And AFoArticulo = ArtID"
    End If
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsAux.EOF
        aTexto = aTexto & rsAux!AFoCantidad & " " & Trim(rsAux!ArtNombre) & ", "
        rsAux.MoveNext
    Loop
    rsAux.Close
    If Len(aTexto) > 2 Then aTexto = Mid(aTexto, 1, Len(aTexto) - 2)
    lArticulos.Caption = aTexto
    Screen.MousePointer = 0
    Exit Sub
    
errArticulos:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurió un error al cargar los artículos de la carpeta.", Err.Description
End Sub

Private Sub cFolder_GotFocus()
    cFolder.SelStart = 0: cFolder.SelLength = Len(cFolder.Text)
End Sub

Private Sub cFolder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(cFolder.Text) = "" Then bDistribuir.SetFocus
        If cFolder.ListIndex <> -1 Then AgregoFolderALista
    End If
    
End Sub

Private Sub tAutoriza_Change()
    tAutoriza.Tag = 0
End Sub

Private Sub tProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And Not sModificar And Val(tProveedor.Tag) <> 0 Then AccionListaDeAyuda
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    sModificar = False
    InicializoGrillas
    
    CargoDatosCombos
    DeshabilitoIngreso
    LimpioFicha
    
    FechaDelServidor
    
    If Trim(Command()) <> "" Then CargoCamposDesdeBD Val(Command())
    If gIdComprobante <> 0 And vsGasto.Rows > 1 Then Botones True, True, True, False, False, Toolbar1, Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
End Sub


Private Sub Label17_Click()
    Foco cFolder
End Sub

Private Sub Label3_Click()
    Foco tProveedor
End Sub

Private Sub Label4_Click()
    Foco tFecha
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuExit_Click()
    Unload Me
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub AccionModificar()

    If Not VerificoFolderCosteado Then Exit Sub
    
    On Error Resume Next
    Screen.MousePointer = 11
    sModificar = True
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoIngreso
    vsGasto.Editable = True: vsFolder.Editable = True
    CargoComboFolder

    tFecha.Enabled = False: tFecha.BackColor = Inactivo
    tProveedor.Enabled = False: tProveedor.BackColor = Inactivo
    cComprobante.Enabled = False: cComprobante.BackColor = Inactivo
    tNumero.Enabled = False: tNumero.BackColor = Inactivo
    
    cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    tIOriginal.Enabled = False: tIOriginal.BackColor = Inactivo
    tTCDolar.Enabled = False: tTCDolar.BackColor = Inactivo
    
    cFolder.SetFocus
    Screen.MousePointer = 0
    
End Sub

'-----------------------------------------------------------------------------
'   Retorna True si no hay ningun folder costeado
'   Retorna False si hay alguno costeado
Private Function VerificoFolderCosteado() As Boolean

Dim Ok As Boolean
Dim aTipoFolder As Integer, aFolder As Long
Dim aTexto As String

    On Error GoTo errValidar
    Screen.MousePointer = 11
    Ok = True
    
    'Valido los Campos de los gasto GImImporte y GImCostear (x si el gasto fue pasado dps de costear la sub carpeta) ----
    'En teoria tiene q coincidir (si es q el gasto no entro en el costeo), si salta x subs controlo la dif.
    Dim mDiffCosteo As Currency: mDiffCosteo = 0
    cons = "Select * from GastoImportacion Where GImIDCompra = " & gIdComprobante
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        mDiffCosteo = mDiffCosteo + (rsAux!GImImporte - rsAux!GImCostear)
        rsAux.MoveNext
    Loop
    rsAux.Close
    '----------------------------------------------------------------------------------------------------------------------------------------
    
    'Verifico si alguna de las carpetas esta costeada
    With vsFolder
    For I = 1 To .Rows - 1
        
        aTipoFolder = Mid(.Cell(flexcpData, I, 0), 1, 1)
        aFolder = Mid(.Cell(flexcpData, I, 0), 2, Len(CStr(.Cell(flexcpData, I, 0))))
        aTexto = .Cell(flexcpText, I, 0)
        
        Select Case aTipoFolder
            Case Folder.cFSubCarpeta
        
                cons = "Select * from Subcarpeta Where SubID = " & aFolder
                Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                If rsAux!SubCosteada Then Ok = False: rsAux.Close: Exit For
                rsAux.Close
                
            Case Folder.cFEmbarque
                    cons = "Select * from Embarque Where EmbID = " & aFolder
                    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                    If rsAux!EmbCosteado Then Ok = False: rsAux.Close: Exit For
                    rsAux.Close
                    
                    '1) Veo si algunas de las sub del embarque esta costeda
                    cons = "Select * from Subcarpeta Where SubEmbarque = " & aFolder & " And SubCosteada = 1"
                    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                    If Not rsAux.EOF Then Ok = False
                    rsAux.Close
                    '2) Valido si hay diff de costeos (x si el gasto fue pasado dps de costear la sub carpeta)
                    If Not Ok Then
                        If mDiffCosteo > 0.25 Then Exit For
                        Ok = True
                    End If
                    
            Case Folder.cFCarpeta
                    cons = "Select * from Carpeta Where CarID = " & aFolder
                    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                    If rsAux!CarCosteada Then Ok = False: rsAux.Close: Exit For
                    rsAux.Close
                    
                    'Valido si algunos de los embarque o sub de c/embarque están costeadas
                    cons = "Select * from Embarque" _
                                            & " Left outer join SubCarpeta ON SubEmbarque = EmbID And SubCosteada = 1" _
                            & " Where EmbCarpeta = " & aFolder _
                            & " And EmbCosteado = 1"
                    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                    If Not rsAux.EOF Then Ok = False
                    rsAux.Close
                    
                    If Not Ok Then
                        If mDiffCosteo > 0.25 Then Exit For
                        Ok = True
                    End If
        End Select
        
    Next
    End With
    VerificoFolderCosteado = Ok
    Screen.MousePointer = 0
    
    If Not Ok Then
        MsgBox "No podrá modificar el registro del gasto debido a que alguna de las carpetas (o subniveles) ha sido costeado." & Chr(vbKeyReturn) _
                 & "Folder " & aTexto, vbExclamation, "Subnivel costeado"
    End If
    
    Exit Function

errValidar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al validar el costeo de las carpetas.", Err.Description
End Function

Private Sub AccionGrabar()
    On Error GoTo errorBT

    Dim aError As String: aError = ""

    Screen.MousePointer = 11
    If Not ValidoCampos Then Screen.MousePointer = 0: Exit Sub
    If Not ValidoCantidadGastos Then Screen.MousePointer = 0: Exit Sub
    If Not ValidoDocumento Then Screen.MousePointer = 0: Exit Sub
    
    Screen.MousePointer = 0
    If MsgBox("Confirma almacenar la información ingresada", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    
    FechaDelServidor
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
                
    cons = "Delete GastoImportacion Where GImIDCompra = " & gIdComprobante
    cBase.Execute cons
    
    'Cargo Tabla: GastosSubRubro
    cons = "Select * from GastoImportacion Where GImIDCompra = " & gIdComprobante
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    With vsFolder
        For I = 1 To .Rows - 1
            rsAux.AddNew
            rsAux!GImIDCompra = gIdComprobante
            rsAux!GImIDSubrubro = .Cell(flexcpData, I, 1)
            rsAux!GImImporte = .ValueMatrix(I, 3)
            rsAux!GImCostear = .ValueMatrix(I, 3)
            rsAux!GImNivelFolder = Mid(.Cell(flexcpData, I, 0), 1, 1)
            rsAux!GImFolder = Mid(.Cell(flexcpData, I, 0), 2, Len(CStr(.Cell(flexcpData, I, 0))))
            rsAux.Update
        Next I
    End With
    rsAux.Close

    cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
        
    gFModificacion = gFechaServidor
    
    sModificar = False
    DeshabilitoIngreso
    Botones True, True, True, False, False, Toolbar1, Me
    Foco tFecha

    RegistroGastoDivisa
    
    Screen.MousePointer = 0
    Exit Sub
    
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    If Trim(aError) = "" Then aError = "No se ha podido inicializar la transacción. Reintente la operación."
    Screen.MousePointer = 0
    clsGeneral.OcurrioError aError, Err.Description
End Sub

'   Ajusta el ingreso automático del gasto divisa DivisaPaga = true
Private Sub RegistroGastoDivisa()
    
    'On Error GoTo errDivisa
    Screen.MousePointer = 11
    With vsFolder
    For I = 1 To .Rows - 1
    
        'Si el gasto es divisa y está asignado al nivel embarque ...
        If .Cell(flexcpData, I, 1) = paSubrubroDivisa And Folder.cFEmbarque = Mid(.Cell(flexcpData, I, 0), 1, 1) Then
        
            'Si el subrubro es divisa --> consulto para ver si está paga
            cons = "Select * from Embarque Where EmbID = " & Mid(.Cell(flexcpData, I, 0), 2, Len(CStr(.Cell(flexcpData, I, 0))))
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then
                If Not rsAux!EmbDivisaPaga Then
                    If MsgBox("Se ha ingresado el gasto " & .Cell(flexcpText, I, 1) & " para la carpeta " & .Cell(flexcpText, I, 0) & Chr(vbKeyReturn) _
                                & "Desea actualizar la divisa como paga.", vbQuestion + vbYesNo + vbDefaultButton2, "Actualizar Divisa Paga") = vbYes Then
                        rsAux.Edit
                        rsAux!EmbDivisaPaga = True
                        rsAux.Update
                    End If
                End If
            End If
            rsAux.Close
        End If
    Next I
    End With
    Screen.MousePointer = 0
    Exit Sub

errDivisa:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al registrar automáticamente el gasto divisa."
End Sub

Private Sub AccionEliminar()
Dim aError As String
    
    If Not VerificoFolderCosteado Then Exit Sub

    If MsgBox("Confirma eliminar el gasto seleccionado", vbQuestion + vbYesNo, "ELIMINAR") = vbYes Then
        Screen.MousePointer = 11
        
        On Error GoTo errorBT
        cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
        On Error GoTo errorET
        
        cons = "Delete GastoImportacion Where GImIDCompra = " & gIdComprobante
        cBase.Execute cons
        
        cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
        LimpioFicha
        DeshabilitoIngreso
        Botones True, False, False, False, False, Toolbar1, Me
        gIdComprobante = 0
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación.", Err.Description
End Sub

Private Sub AccionCancelar()

    LimpioFicha
    If sModificar Then
        Botones True, True, True, False, False, Toolbar1, Me
        CargoCamposDesdeBD gIdComprobante
    Else
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    
    DeshabilitoIngreso
    sModificar = False
    Foco tFecha

End Sub

Private Sub CargoCamposDesdeBD(IdCompra As Long)

Dim aValor As Long

    Screen.MousePointer = 11
    On Error GoTo errCargar
    'Cargo los datos desde la tabla Compra-----------------------------------------------------------------------------------------
    cons = "Select * from Compra Left Outer Join Usuario On ComUsrAutoriza = UsuCodigo" & _
               " Where ComCodigo = " & IdCompra
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then
        rsAux.Close
        MsgBox "No existe registro de compra para el id ingresado.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: gIdComprobante = 0: Botones True, False, False, False, False, Toolbar1, Me: Exit Sub
    End If
    
    gIdComprobante = rsAux!ComCodigo
    gFModificacion = rsAux!ComFModificacion
    
    tID.Text = Format(rsAux!ComCodigo, "#,###,##0")
    tFecha.Text = Format(rsAux!ComFecha, FormatoFP)
    
    'Datos del Proveedor------------------------------------------------------------------------------------
    Dim rs1 As rdoResultset
    cons = "Select * from ProveedorCliente Where PClCodigo = " & rsAux!ComProveedor
    Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rs1.EOF Then
        tProveedor.Text = Trim(rs1!PClFantasia)
        tProveedor.Tag = rsAux!ComProveedor
    End If
    
    If Not IsNull(rsAux!UsuCodigo) Then
        tAutoriza.Text = Trim(rsAux!UsuIdentificacion)
        tAutoriza.Tag = rsAux!UsuCodigo
    End If
    lAutoriza.Tag = tAutoriza.Tag
    
    If Not IsNull(rsAux!ComVerificado) Then
        If rsAux!ComVerificado = 1 Then chVerificado.Value = vbChecked Else chVerificado.Value = vbUnchecked
    Else
        chVerificado.Value = vbGrayed
    End If
    
    rs1.Close
    '------------------------------------------------------------------------------------------------------------
    
    BuscoCodigoEnCombo cComprobante, rsAux!ComTipoDocumento
    If Not IsNull(rsAux!ComNumero) Then tNumero.Text = rsAux!ComNumero
    
    BuscoCodigoEnCombo cMoneda, rsAux!ComMoneda
    tIOriginal.Text = Format(rsAux!ComImporte, FormatoMonedaP)
    If Not IsNull(rsAux!ComIVA) Then tIOriginal.Text = Format(rsAux!ComImporte + rsAux!ComIVA, FormatoMonedaP)
    
    If Not IsNull(rsAux!ComTC) Then If rsAux!ComTC <> 1 Then tTCDolar.Text = Format(rsAux!ComTC, "0.000")
    
    If Not IsNull(rsAux!ComComentario) Then tComentario.Text = Trim(rsAux!ComComentario)
    rsAux.Close
    
    'Cargo los datos desde la BD GastosSubRubro-----------------------------------------------------------------------------------------
    cons = "Select * from GastoSubrubro, Subrubro " _
           & " Where GSrIDCompra = " & IdCompra _
           & " And GSrIDSubrubro = SRuID" & " And SRuRubro IN (" & paRubroGastosImportaciones & ")"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        With vsGasto
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Trim(rsAux!SRuNombre)                                    'Gasto
            aValor = rsAux!SRuID: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            .Cell(flexcpText, .Rows - 1, 1) = RetornoNombreFolder(rsAux!SRuNivel)     'Nivel
            aValor = rsAux!SRuNivel: .Cell(flexcpData, .Rows - 1, 1) = aValor
            
            .Cell(flexcpText, .Rows - 1, 2) = RetornoNombreDistribucion(rsAux!SRuDistribucion)     'Distribucion
            aValor = rsAux!SRuDistribucion: .Cell(flexcpData, .Rows - 1, 2) = aValor
            
            .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!GSrImporte, FormatoMonedaP)
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    'Cargo los campos desde la tabla GastosImportacion-----------------------------------------------------------------------------------------
    Dim RsFol As rdoResultset
    Dim aTexto As String
    
    cons = "Select * from GastoImportacion, Subrubro " _
           & " Where GImIDCompra = " & IdCompra _
           & " And GImIDSubrubro = SRuID"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        With vsFolder
            .AddItem ""
            
            Select Case rsAux!GImNivelFolder
                Case Folder.cFCarpeta: cons = "Select Carpeta = CarCodigo, Embarque = Null, Sub = Null from Carpeta Where CarId = " & rsAux!GImFolder
                
                Case Folder.cFEmbarque
                    cons = "Select Carpeta = CarCodigo, Embarque = EmbCodigo, Sub = Null from Embarque, Carpeta " _
                           & " Where EmbId = " & rsAux!GImFolder & " And EmbCarpeta = CarID"
                
                Case Folder.cFSubCarpeta
                    cons = "Select Carpeta = CarCodigo, Embarque = EmbCodigo, Sub = SubCodigo from Subcarpeta, Embarque, Carpeta " _
                           & " Where SubId = " & rsAux!GImFolder & " And SubEmbarque = EmbID And EmbCarpeta = CarID"
            End Select
            Set RsFol = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            aTexto = RsFol!Carpeta
            If Not IsNull(RsFol!Embarque) Then aTexto = aTexto & "." & Trim(RsFol!Embarque)
            If Not IsNull(RsFol!Sub) Then aTexto = aTexto & "/" & Trim(RsFol!Sub)
            RsFol.Close
            
            .Cell(flexcpText, .Rows - 1, 0) = aTexto     'Folder
            aValor = rsAux!GImNivelFolder & rsAux!GImFolder: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!SRuNombre)                                  'Gasto
            aValor = rsAux!SRuID: .Cell(flexcpData, .Rows - 1, 1) = aValor
            
            .Cell(flexcpBackColor, .Rows - 1, 3) = Obligatorio
            .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!GImImporte, FormatoMonedaP)
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    Screen.MousePointer = 0
    Exit Sub
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del comprobante.", Err.Description
End Sub

Private Sub tFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And Not sModificar Then AccionListaDeAyuda
    If KeyCode = vbKeyDown Then tFecha.Text = Format(Now, FormatoFP)
End Sub

Private Sub AccionListaDeAyuda()

    On Error GoTo errAyuda
    
    If Not IsDate(tFecha.Text) And Val(tProveedor.Tag) = 0 Then Exit Sub
    Screen.MousePointer = 11
    
    Dim aLista As New clsListadeAyuda
    Dim aSeleccionado As Long: aSeleccionado = 0
    
    cons = " Select ID_Compra = ComCodigo, Fecha = ComFecha, Proveedor = PClFantasia, ComNumero as Comprobante, Moneda = MonSigno , Importe = ComImporte, Comentarios = ComComentario" _
            & " from Compra, ProveedorCliente, Moneda" _
            & " Where ComProveedor = PClCodigo" _
            & " And ComMoneda = MonCodigo"
            
    If IsDate(tFecha.Text) Then cons = cons & " And ComFecha >= '" & Format(tFecha.Text, sqlFormatoF) & "'"
    If Val(tProveedor.Tag) <> 0 Then cons = cons & " And ComProveedor = " & Val(tProveedor.Tag)
    
    cons = cons & " Order by ComFecha DESC"
    
    aLista.ActivoListaAyudaSQL cBase, cons
    Me.Refresh
    
    If IsNumeric(aLista.ItemSeleccionadoSQL) Then aSeleccionado = CLng(aLista.ItemSeleccionadoSQL)
    Set aLista = Nothing
    
    If aSeleccionado <> 0 Then LimpioFicha: CargoCamposDesdeBD aSeleccionado
    If gIdComprobante <> 0 And vsGasto.Rows > 1 Then Botones True, True, True, False, False, Toolbar1, Me
    
    Screen.MousePointer = 0
    Exit Sub
        
errAyuda:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al activar la lista de ayuda.", Err.Description
End Sub

Private Sub tID_Change()
    If tID.Enabled Then Botones True, False, False, False, False, Toolbar1, Me
End Sub

Private Sub tID_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tID.Text) = "" Then Foco tFecha: Exit Sub
        If Not IsNumeric(tID.Text) Then MsgBox "El id ingresado no es correcto. Verifique.", vbExclamation, "ATENCIÓN": Exit Sub
        gIdComprobante = CLng(tID.Text)
        LimpioFicha
        CargoCamposDesdeBD gIdComprobante
        If gIdComprobante <> 0 And vsGasto.Rows > 1 Then Botones True, True, True, False, False, Toolbar1, Me
    End If
    
End Sub

Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0
    tFecha.SelLength = Len(tFecha.Text)
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tProveedor
End Sub

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, FormatoFP) Else tFecha.Text = ""
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    
    Select Case Button.Key
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        
        Case "embarque": EjecutarApp App.Path & "\Consulta Embarques"
        Case "sub": EjecutarApp App.Path & "\Consulta Subcarpetas"
        Case "gasto": EjecutarApp App.Path & "\Consulta de Gastos"
        Case "dolar": EjecutarApp App.Path & "\Tasa de Cambio"
    End Select

End Sub

Private Sub DeshabilitoIngreso()

    tID.Enabled = True: tID.BackColor = Blanco
    tFecha.Enabled = True: tFecha.BackColor = Blanco
    tProveedor.Enabled = True: tProveedor.BackColor = Blanco
    cComprobante.Enabled = False: cComprobante.BackColor = Inactivo
    tNumero.Enabled = False: tNumero.BackColor = Inactivo
    
    cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    tIOriginal.Enabled = False: tIOriginal.BackColor = Inactivo
    
    tTCDolar.Enabled = False: tTCDolar.BackColor = Inactivo
    
    
    cFolder.Enabled = False: cFolder.BackColor = Inactivo
    
    bAgregar.Enabled = False
    bDistribuir.Enabled = False
    
    tComentario.Enabled = False: tComentario.BackColor = Inactivo
    
    vsGasto.BackColor = Inactivo: vsGasto.Editable = True
    vsFolder.BackColor = Inactivo:  vsFolder.Editable = True
    
    tAutoriza.Enabled = False: tAutoriza.BackColor = Inactivo
    chVerificado.Enabled = False
    
End Sub

Private Sub HabilitoIngreso()

    tID.Enabled = False: tID.BackColor = Inactivo
    tFecha.Enabled = False: tFecha.BackColor = Inactivo
    tProveedor.Enabled = False: tProveedor.BackColor = Inactivo
    cComprobante.Enabled = False: cComprobante.BackColor = Inactivo
    tNumero.Enabled = False: tNumero.BackColor = Inactivo
    
    cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    tIOriginal.Enabled = False: tIOriginal.BackColor = Inactivo
    tTCDolar.Enabled = False: tTCDolar.BackColor = Inactivo
    
    
    cFolder.Enabled = True: cFolder.BackColor = Blanco
    
    bAgregar.Enabled = True
    bDistribuir.Enabled = True
    
    tComentario.Enabled = False: tComentario.BackColor = Inactivo
    
    tAutoriza.Enabled = False: tAutoriza.BackColor = Inactivo
    chVerificado.Enabled = False
    
    vsGasto.BackColor = Inactivo: vsGasto.Editable = False
    vsFolder.BackColor = Blanco: vsFolder.Editable = True
    
End Sub

Private Sub LimpioFicha()
    
    tID.Text = ""
    tFecha.Text = ""
    tProveedor.Text = ""
    cComprobante.Text = "": tNumero.Text = ""
    cMoneda.Text = "": tIOriginal.Text = ""
    
    tTCDolar.Text = ""
    lTC.Caption = ""
    
    cFolder.Clear: cFolder.Tag = ""
    lArticulos.Caption = ""
    
    vsGasto.Rows = 1
    vsFolder.Rows = 1
    
    tComentario.Text = ""
    tAutoriza.Text = "": chVerificado.Value = vbGrayed
    
End Sub

Private Sub tProveedor_Change()
    On Error Resume Next
    tProveedor.Tag = 0
    tTCDolar.Tag = "": tTCDolar.Text = "": lTC.Caption = ""

    cMoneda.Text = ""
End Sub

Private Sub tProveedor_GotFocus()
    tProveedor.SelStart = 0: tProveedor.SelLength = Len(tProveedor.Text)
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
    On Error GoTo errBuscar
    
    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) <> 0 Or Trim(tProveedor.Text) = "" Then
            If tID.Enabled Then Foco tID Else Foco cComprobante
            Exit Sub
        End If
        
        Dim aQ As Long, aIdProveedor As Long, aTexto As String
        aQ = 0
        Screen.MousePointer = 11
        cons = "Select PClCodigo, PClFantasia as 'Nombre', PClNombre as 'Razón Social' from ProveedorCliente " _
                & " Where PClNombre like '" & Trim(tProveedor.Text) & "%' Or PClFantasia like '" & Trim(tProveedor.Text) & "%'"
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            aQ = 1: aIdProveedor = rsAux!PClCodigo: aTexto = Trim(rsAux!Nombre)
            rsAux.MoveNext: If Not rsAux.EOF Then aQ = 2
        End If
        rsAux.Close
        
        Select Case aQ
            Case 0:
                    MsgBox "No existe una empresa para el con el nombre ingresado.", vbExclamation, "No existe Empresa"
            Case 1:
                    tProveedor.Text = aTexto: tProveedor.Tag = aIdProveedor
            Case 2:
                    
                    Dim aLista As New clsListadeAyuda
                    
                    If aLista.ActivarAyuda(cBase, cons, 5500, 1) <> 0 Then
                        tProveedor.Text = Trim(aLista.RetornoDatoSeleccionado(1))
                        tProveedor.Tag = aLista.RetornoDatoSeleccionado(0)
                    Else
                        tProveedor.Text = ""
                    End If
                    
                    Set aLista = Nothing
        End Select
        Screen.MousePointer = 0
    
    End If
    Exit Sub

errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al procesar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub InicializoGrillas()

Dim aTexto As String: aTexto = ""
    
    On Error Resume Next
    With vsGasto
        .Rows = 1: .Cols = 4
        .Editable = False
        .FormatString = "Gasto|Nivel|Distribución|>Importe"
            
        .WordWrap = True
        .ColWidth(0) = 3500
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColDataType(3) = flexDTCurrency
    End With
    
    With vsFolder
        .Rows = 1: .Cols = 4
        .Editable = False
        .FormatString = "<Carpeta|Gasto|>Q / Volumen / $|>Importe Asignado"
            
        .WordWrap = True
        .ColWidth(0) = 1000
        .ColWidth(1) = 3400
        .ColWidth(2) = 1500
        .ColDataType(3) = flexDTCurrency
    End With
    
End Sub

Private Sub CargoDatosCombos()

    On Error Resume Next
    
    cons = "Select MonCodigo, MonSigno from Moneda Where MonCodigo In (" & paMonedaDolar & ", " & paMonedaPesos & ")"
    CargoCombo cons, cMoneda
    
    'Cargo los valores para los comprobantes de pago
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraContado)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraContado
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraCredito)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraCredito
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaCredito)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraNotaCredito
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaDevolucion)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraNotaDevolucion
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraRecibo)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraRecibo
    
End Sub

Private Sub vsFolder_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not sModificar Then Cancel = True: Exit Sub
    If Col <> 3 Then Cancel = True
End Sub

Private Sub vsFolder_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        If vsFolder.RowSel < 1 Then Exit Sub
        If Not sModificar Then Exit Sub
        vsFolder.RemoveItem vsFolder.RowSel
    End If
    
End Sub

Private Sub vsFolder_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If sModificar And vsFolder.Rows > vsFolder.FixedRows Then AccionGrabar
    End If
End Sub

Private Sub vsFolder_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not IsNumeric(vsFolder.EditText) Then Cancel = True: Exit Sub
    If CCur(vsFolder.EditText) = 0 Then Cancel = True: Exit Sub
    
    vsFolder.EditText = Format(vsFolder.EditText, FormatoMonedaP)
End Sub

Private Sub vsGasto_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not sModificar Then Cancel = True: Exit Sub
    If Col <> 3 Then Cancel = True
End Sub

Private Sub vsGasto_GotFocus()
    On Error Resume Next
    vsGasto.Select vsGasto.Row, 3
End Sub

Private Sub vsGasto_KeyDown(KeyCode As Integer, Shift As Integer)

Dim sSalir As Boolean: sSalir = False

    If KeyCode = vbKeyDelete Then
        If vsGasto.RowSel < 1 Then Exit Sub
        If Not sModificar Then Exit Sub
        
        On Error Resume Next
        Do While Not sSalir
            sSalir = True
            For I = 1 To vsFolder.Rows - 1
                If vsFolder.Cell(flexcpData, I, 1) = vsGasto.Cell(flexcpData, vsGasto.RowSel, 0) Then
                    vsFolder.RemoveItem I
                    sSalir = False: Exit For
                End If
            Next
        Loop
        
        vsGasto.RemoveItem vsGasto.RowSel
        
        cFolder.Tag = "": cFolder.Clear
        CargoComboFolder
    End If
    
End Sub

Private Sub CargoComboFolder()
'   Segun Carlos, el 28/8, si el nivel del gasto es Madre -> cargo solo C.Madres
'                                                                    Embarque -> cargo solo Embarques
'                                                                    SubCarpeta -> cargo embarques y Subcarpetas
'   Se Mofico el 8/10   Si es Sub -> Cargo embarques q No van a ZF y Subs
'
'   En el tag del combo guardo los niveles que tengo cargados ---> para no repetir Ej.:     1:2:  , despues busco el texto nivel:
'   para ver si está cargado.

Dim Carpeta As Boolean, Parcial As Boolean, SubCarpeta As Boolean

    On Error GoTo errCargar
    Screen.MousePointer = 11
    Carpeta = False: Parcial = False: SubCarpeta = False
    
    For I = 1 To vsGasto.Rows - 1
        If vsGasto.Cell(flexcpData, I, 1) = Folder.cFCarpeta Then Carpeta = True
        If vsGasto.Cell(flexcpData, I, 1) = Folder.cFEmbarque Then Parcial = True
        If vsGasto.Cell(flexcpData, I, 1) = Folder.cFSubCarpeta Then Parcial = True: SubCarpeta = True
    Next
    
    cons = ""
    cFolder.Clear
    
    If Carpeta Then
        cons = "Select Nivel = " & Folder.cFCarpeta & ", ID = CarID, Car = CarCodigo, Emb = '', Sub = 0 " _
               & " From Carpeta Where CarCosteada = 0" _
               & " UNION ALL"
    End If
    
    If Parcial Then
        If Parcial And Not SubCarpeta Then      'El Nivel es Parcial
            cons = cons & " Select Nivel = " & Folder.cFEmbarque & ", ID = EmbID, Car = CarCodigo, Emb = EmbCodigo, Sub = 0 " _
                                & " From Embarque, Carpeta " _
                                & " Where EmbCarpeta = CarID And EmbCosteado = 0" _
                                & " UNION ALL"
            
        Else                                                   'El nivel es Sub Cargo las que No Van a ZF
            cons = cons & " Select Nivel = " & Folder.cFEmbarque & ", ID = EmbID, Car = CarCodigo, Emb = EmbCodigo, Sub = 0 " _
                                & " From Embarque , Carpeta " _
                                & " Where EmbCarpeta = CarID And EmbCosteado = 0 And EmbLocal <> " & paLocalZF _
                                & " UNION ALL"
        End If
    End If

    If SubCarpeta Then
        cons = cons & " Select Nivel = " & Folder.cFSubCarpeta & ", ID = SubID, Car = CarCodigo, Emb = EmbCodigo, Sub = SubCodigo " _
                        & " From SubCarpeta, Embarque, Carpeta " _
                        & " Where SubEmbarque = EmbID And EmbCarpeta = CarID And SubCosteada = 0" _
                        & " UNION ALL"
    End If
    
    If cons <> "" Then
        cons = Mid(cons, 1, Len(cons) - Len("UNION ALL")) & " Order by Car, Emb, Sub"
    
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        Do While Not rsAux.EOF
            aTexto = rsAux!Car
            If Trim(rsAux!Emb) <> "" Then
                aTexto = aTexto & "." & Trim(rsAux!Emb)
                If rsAux!Sub <> 0 Then aTexto = aTexto & "/" & rsAux!Sub
            End If
            
            cFolder.AddItem (aTexto)
            cFolder.ItemData(cFolder.NewIndex) = rsAux!Nivel & rsAux!ID
            rsAux.MoveNext
        Loop
        rsAux.Close
    End If
    
    Screen.MousePointer = 0
    Exit Sub

errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar las carpetas. ", Err.Description
End Sub

Private Sub vsGasto_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error Resume Next
    If Not IsNumeric(vsGasto.EditText) Then Cancel = True: Exit Sub
    If CCur(vsGasto.EditText) = 0 Then Cancel = True: Exit Sub
    
    vsGasto.EditText = Format(vsGasto.EditText, FormatoMonedaP)
    
End Sub

Private Function ValidoCampos() As Boolean

Dim aTotal As Currency
    On Error GoTo errValido
    ValidoCampos = False
    
    
    If vsFolder.Rows = 1 Then
        MsgBox "Debe ingresar las carpetas para realizar la distribución del gasto.", vbExclamation, "ATENCIÓN"
        Foco cFolder: Exit Function
    End If
       
    'Valido importe de los gastos contra importe de distribucion
    Dim J As Integer
    For I = 1 To vsGasto.Rows - 1
        aTotal = 0
        For J = 1 To vsFolder.Rows - 1
            If vsFolder.Cell(flexcpData, J, 1) = vsGasto.Cell(flexcpData, I, 0) Then
                If vsFolder.Cell(flexcpValue, J, 3) = 0 Then
                    MsgBox "Hay carpetas que tienen valor cero de distribución. Vuelva a realiar la distribución automática.", vbExclamation, "ATENCIÓN"
                    Exit Function
                End If
                aTotal = aTotal + vsFolder.Cell(flexcpValue, J, 3)
            End If
        Next J
        If vsGasto.Cell(flexcpValue, I, 3) <> aTotal Then
            MsgBox "No coincide el importe del gasto (" & vsGasto.Cell(flexcpText, I, 0) & ") con la suma de los importes de distribución (" & Format(aTotal, FormatoMonedaP) & ").", vbExclamation, "ATENCIÓN"
            Exit Function
        End If
    Next I
        
    ValidoCampos = True
    Exit Function

errValido:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al validar los datos.", Err.Description
End Function

Private Function ValidoCantidadGastos() As Boolean

Dim aTipoFolder As Integer, aFolder As Long
Dim aGasto As Long, aCantidad As Integer

    On Error GoTo errValido
    ValidoCantidadGastos = False
    With vsFolder
    For I = 1 To .Rows - 1
        aCantidad = 0
        aGasto = .Cell(flexcpData, I, 1)
        aTipoFolder = Mid(.Cell(flexcpData, I, 0), 1, 1)
        aFolder = Mid(.Cell(flexcpData, I, 0), 2, Len(CStr(.Cell(flexcpData, I, 0))))
    
        cons = "Select Count(*) from GastoImportacion " _
                & " Where GImIDSubrubro = " & aGasto _
                & " And GImFolder = " & aFolder _
                & " And GImNivelFolder = " & aTipoFolder _
                & " And GImIDCompra <> " & gIdComprobante
        Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        If Not rsAux.EOF Then If Not IsNull(rsAux(0)) Then aCantidad = rsAux(0)
        rsAux.Close
        
        If aCantidad > 0 Then
            If aGasto = paSubrubroDivisa Then
                If MsgBox("El gasto divisa ya está ingresado, no podrá registrar uno nuevo." & Chr(vbKeyReturn) _
                        & "Para modificar el gasto o registrar un ajuste, se recomienda utilizar el mantenimiento de embarques." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                        & "Está seguro de continuar ", vbExclamation + vbYesNo + vbDefaultButton2, "Registro de Divisa") = vbNo Then _
                Exit Function
            End If
                
            cons = "Select * from Subrubro Where SRuID = " & aGasto
            Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not IsNull(rsAux!SRuCantidad) Then
                If aCantidad >= rsAux!SRuCantidad Then
                    If MsgBox("La cantidad de gastos por carpeta (" & rsAux!SRuCantidad & ") para el gasto " & .TextMatrix(I, 1) & ", en la carpeta " & .TextMatrix(I, 0) & " está cubierta." & Chr(vbKeyReturn) _
                        & "Desea continuar con el ingreso", vbYesNo + vbInformation + vbDefaultButton2, "Gastos por Carpeta") = vbNo Then
                        rsAux.Close: Exit Function
                    End If
                End If
            End If
            rsAux.Close
        End If
    Next
    End With
    
    ValidoCantidadGastos = True
    Exit Function
errValido:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al validar la cantidad de gastos por carpeta.", Err.Description
End Function

Private Function ValidoDocumento() As Boolean
Dim RsA2 As rdoResultset
Dim aDato As String

    'On Error Resume Next
    ValidoDocumento = False
    
    cons = "Select * from Compra" _
           & " Where ComCodigo <> " & gIdComprobante
           
    If Trim(tNumero.Text) <> "" Then cons = cons & " And ComNumero = '" & Trim(tNumero.Text) & "'"
    
    cons = cons & " And ComProveedor = " & Val(tProveedor.Tag) _
                & " And ComMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                & " And ComImporte = " & CCur(tIOriginal.Text)
           
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        'El folder esta en el Data(0)       Valido las carpetas asignadas
        cons = "Select * from GastoImportacion Where GImIDCompra = " & rsAux!ComCodigo
        Set RsA2 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsA2.EOF
            For I = 1 To vsFolder.Rows - 1
                aDato = vsFolder.Cell(flexcpData, I, 0)
                If CLng(Mid(aDato, 1, 1)) = RsA2!GImNivelFolder And CLng(Mid(aDato, 2, Len(aDato))) = RsA2!GImFolder Then
                    Screen.MousePointer = 0
                    If MsgBox("Ya existen gastos registrados con el mismo documento y proveedor. Posiblemente estén asignados a las mismas carpetas." & Chr(vbKeyReturn) _
                        & "Fecha: " & Format(rsAux!ComFecha, "d-mmm yyyy") & Chr(vbKeyReturn) _
                        & "Importe: " & Format(rsAux!ComImporte, "##,##0.00") & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                        & "Desea proseguir con el ingreso del gasto.", vbInformation + vbYesNo + vbDefaultButton2, "ATENCIÓN") = vbNo Then
                            rsAux.Close: RsA2.Close: Exit Function
                    End If
                    Screen.MousePointer = 11
                End If
            Next
            RsA2.MoveNext
        Loop
        RsA2.Close
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    ValidoDocumento = True
                   
End Function

