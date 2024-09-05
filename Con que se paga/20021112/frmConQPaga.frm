VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmConQPaga 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Con Qué Paga"
   ClientHeight    =   4740
   ClientLeft      =   1785
   ClientTop       =   2835
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConQPaga.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   9030
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
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
            Object.Visible         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
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
            Style           =   4
            Object.Width           =   6750
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame fCon 
      Caption         =   "Con Qué se paga"
      ForeColor       =   &H00000080&
      Height          =   2505
      Left            =   60
      TabIndex        =   34
      Top             =   1920
      Width           =   8895
      Begin VB.TextBox tPesos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         MaxLength       =   13
         TabIndex        =   15
         Text            =   "1,000,000.00"
         Top             =   285
         Width           =   1095
      End
      Begin VB.TextBox tImporteD 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7560
         MaxLength       =   13
         TabIndex        =   19
         Text            =   "10/10/2000"
         Top             =   645
         Width           =   1215
      End
      Begin VB.TextBox tChImporte 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7560
         MaxLength       =   13
         TabIndex        =   29
         Text            =   "1,000,000.00"
         Top             =   1030
         Width           =   1215
      End
      Begin VB.TextBox tChVence 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5640
         MaxLength       =   11
         TabIndex        =   27
         Text            =   "10/10/2000"
         Top             =   1030
         Width           =   975
      End
      Begin VB.TextBox tChLibrado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3960
         MaxLength       =   11
         TabIndex        =   25
         Text            =   "10/10/2000"
         Top             =   1030
         Width           =   975
      End
      Begin VB.TextBox tChNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2040
         MaxLength       =   12
         TabIndex        =   23
         Text            =   "123456789012"
         Top             =   1030
         Width           =   1095
      End
      Begin VB.TextBox tChSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1640
         MaxLength       =   2
         TabIndex        =   22
         Text            =   "AA"
         Top             =   1030
         Width           =   340
      End
      Begin VB.TextBox tImporteP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         MaxLength       =   13
         TabIndex        =   13
         Text            =   "1,000,000.00"
         Top             =   285
         Width           =   1095
      End
      Begin VSFlex6DAOCtl.vsFlexGrid vsPago 
         Height          =   975
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   8655
         _ExtentX        =   15266
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
      Begin AACombo99.AACombo cDisponibilidad 
         Height          =   315
         Left            =   1200
         TabIndex        =   17
         Top             =   645
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         ForeColor       =   0
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
      Begin AACombo99.AACombo cChTipo 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   1005
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         ForeColor       =   0
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
      Begin VB.Label lTC 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe en Pe&sos:"
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
         Left            =   5280
         TabIndex        =   40
         Top             =   285
         Width           =   3495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe en Pe&sos:"
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "$ Dis&ponibilidad:"
         Height          =   255
         Left            =   6000
         TabIndex        =   18
         Top             =   645
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Impor&te:"
         Height          =   255
         Left            =   6840
         TabIndex        =   28
         Top             =   1035
         Width           =   615
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "&Vence:"
         Height          =   255
         Left            =   5040
         TabIndex        =   26
         Top             =   1030
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Librado:"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   1030
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nº:"
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   1030
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Disponibilidad:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "I&mporte Orig.:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.Frame fQue 
      Caption         =   "Que se paga"
      ForeColor       =   &H00000080&
      Height          =   1365
      Left            =   60
      TabIndex        =   33
      Top             =   480
      Width           =   8895
      Begin VB.TextBox tCofis 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5520
         MaxLength       =   13
         TabIndex        =   43
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cMoneda 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1200
         Style           =   1  'Simple Combo
         TabIndex        =   39
         Text            =   "Combo1"
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton bBuscar2 
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
         Left            =   5520
         TabIndex        =   11
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton bBuscar1 
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
         Left            =   8460
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox tSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3480
         MaxLength       =   9
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox tNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4200
         MaxLength       =   9
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox tImporte 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   1920
         MaxLength       =   13
         TabIndex        =   36
         Text            =   "1,000,000.00"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox tIva 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4080
         MaxLength       =   13
         TabIndex        =   35
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox tIDCompra 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   1
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2940
         MaxLength       =   11
         TabIndex        =   3
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   5100
         MaxLength       =   40
         TabIndex        =   5
         Top             =   240
         Width           =   3315
      End
      Begin AACombo99.AACombo cComprobante 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         ForeColor       =   12582912
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
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cofis:"
         Height          =   255
         Left            =   4980
         TabIndex        =   44
         Top             =   1005
         Width           =   495
      End
      Begin VB.Label lTotalGasto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1,000,000.00"
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
         Left            =   7500
         TabIndex        =   42
         Top             =   980
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Gasto:"
         Height          =   255
         Left            =   6360
         TabIndex        =   41
         Top             =   1005
         Width           =   1155
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A.:"
         Height          =   255
         Left            =   3480
         TabIndex        =   38
         Top             =   1005
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1005
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Comprobante:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "&ID Compra:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   31
      Top             =   4485
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "terminal"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "bd"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7805
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConQPaga.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConQPaga.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConQPaga.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConQPaga.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConQPaga.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConQPaga.frx":0BA4
            Key             =   ""
         EndProperty
      EndProperty
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
         Visible         =   0   'False
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
   End
   Begin VB.Menu MnuBases 
      Caption         =   "&Bases"
      Begin VB.Menu MnuBx 
         Caption         =   "MnuBx"
         Index           =   0
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
Attribute VB_Name = "frmConQPaga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmIDGasto As Long

Dim sNuevo As Boolean, sModificar As Boolean
Dim RsCom As rdoResultset
Dim bEsNuevo As Boolean
Dim gFModificacion As Date
Dim aValor As Long

Dim aTiposDocs As String

Private Sub bBuscar1_Click()

    If Not IsDate(tFecha.Text) Then
        MsgBox "Ingrese una fecha para buscar los documentos.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Sub
    End If
    
    If Val(tProveedor.Tag) = 0 Then
        MsgBox "Seleccione le proveedor para buscar los documentos.", vbExclamation, "ATENCIÓN"
        Foco tProveedor: Exit Sub
    End If
        
    cons = " Select ID_Compra = ComCodigo, Fecha = ComFecha, Proveedor = PClFantasia, Comprobante = ComSerie + Convert(char(10), ComNumero), Moneda = MonSigno , Importe = ComImporte, Comentarios = ComComentario" _
            & " from Compra, ProveedorCliente, Moneda" _
            & " Where ComProveedor = PClCodigo" _
            & " And ComMoneda = MonCodigo" _
            & " And ComTipoDocumento In (" & aTiposDocs & ")"
            
    If IsDate(tFecha.Text) Then cons = cons & " And ComFecha >= '" & Format(tFecha.Text, sqlFormatoF) & "'"
    If Val(tProveedor.Tag) <> 0 Then cons = cons & " And ComProveedor = " & Val(tProveedor.Tag)
    
    cons = cons & " Order by ComFecha DESC"
    
    AccionListaDeAyuda cons
    
End Sub

Private Sub bBuscar2_Click()
    
    If cComprobante.ListIndex = -1 Then
        MsgBox "Seleccione el tipo de comprobante a buscar.", vbExclamation, "ATENCIÓN"
        Foco cComprobante: Exit Sub
    End If
    
    If Trim(tNumero.Text) = "" Then
        MsgBox "Ingrese el numero de comprobante a buscar.", vbExclamation, "ATENCIÓN"
        Foco tNumero: Exit Sub
    End If
        
    cons = " Select ID_Compra = ComCodigo, Fecha = ComFecha, Proveedor = PClFantasia, Comprobante = ComSerie + Convert(char(10), ComNumero), Moneda = MonSigno , Importe = ComImporte, Comentarios = ComComentario" _
            & " from Compra, ProveedorCliente, Moneda" _
            & " Where ComProveedor = PClCodigo" _
            & " And ComMoneda = MonCodigo" _
            & " And ComTipoDocumento = " & cComprobante.ItemData(cComprobante.ListIndex)
            
    If Trim(tSerie.Text) <> "" Then cons = cons & " And ComSerie = '" & Trim(tSerie.Text) & "'"
    If Trim(tNumero.Text) <> "" Then cons = cons & " And ComNumero = " & Trim(tNumero.Text)
    
    AccionListaDeAyuda cons

End Sub

Private Sub cChTipo_Change()
    DeshabilitoCamposCheque DesdeTipo:=True
End Sub

Private Sub cChTipo_Click()
    DeshabilitoCamposCheque DesdeTipo:=True
End Sub

Private Sub cChTipo_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If cChTipo.ListIndex = -1 Then Exit Sub
        If cChTipo.ItemData(cChTipo.ListIndex) = 1 Then
            tChSerie.Enabled = True: tChSerie.BackColor = Obligatorio
            tChNumero.Enabled = True: tChNumero.BackColor = Obligatorio
            Foco tChSerie
        Else
            AgregoPago
        End If
    End If

End Sub

Private Sub cComprobante_GotFocus()
    With cComprobante: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub cComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco tSerie
End Sub

Private Sub cComprobante_LostFocus()
cComprobante.SelLength = 0
End Sub

Private Sub cDisponibilidad_Change()
    DeshabilitoCamposCheque True
End Sub

Private Sub cDisponibilidad_Click()
    DeshabilitoCamposCheque True
End Sub

Private Sub cDisponibilidad_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        
        If Not IsNumeric(tImporte.Text) Then
            MsgBox "El importe ingresado no es correcto. Verifique", vbExclamation, "ATENCIÓN"
            Foco tImporteP: Exit Sub
        End If
        If cDisponibilidad.ListIndex = -1 Then
            MsgBox "Debe seleccionar la disponibilidad con la que se realizará el pago.", vbExclamation, "ATENCIÓN"
            Foco cDisponibilidad: Exit Sub
        End If
        
        'Consulta para ver que tipo de disponibilidad es-----------------------------------------------------------
        cons = "Select * from Disponibilidad Where DisID = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If rsAux!DisMoneda <> cMoneda.ItemData(cMoneda.ListIndex) Then
            tImporteD.Enabled = True: tImporteD.BackColor = Obligatorio
            If rsAux!DisMoneda = paMonedaPesos And IsNumeric(tPesos.Text) Then tImporteD.Text = tPesos.Text
        End If
        
        If Not IsNull(rsAux!DisSucursal) Then
            cChTipo.Enabled = True: cChTipo.BackColor = Obligatorio
            BuscoCodigoEnCombo cChTipo, 1
            tChNumero.Enabled = True: tChNumero.BackColor = Obligatorio
            tChSerie.Enabled = True: tChSerie.BackColor = Obligatorio
        End If
        rsAux.Close
        '------------------------------------------------------------------------------------------------------------------
        If tImporteD.Enabled Then Foco tImporteD: Exit Sub
        If cChTipo.Enabled Then Foco cChTipo: Exit Sub
        AgregoPago 'Agrego el pago a la lista
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    If paMDPagoDeCompra = 0 Then MsgBox "El parámetro Pago de compras (movimiento de disponibilidades) no está cargado.", vbExclamation, "Parámetros"
End Sub

Private Sub Form_Load()
    On Error Resume Next
    'Center Form
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2, Me.Width, Me.Height
    
    Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    
    aTiposDocs = TipoDocumento.Compracontado & ", " & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraNotaDevolucion & ", " _
                       & TipoDocumento.CompraRecibo & ", " & TipoDocumento.CompraReciboDePago & ", " _
                       & TipoDocumento.CompraSalidaCaja & ", " & TipoDocumento.CompraEntradaCaja
                       
    InicializoGrillas
    
    LoadME
    
    If prmIDGasto <> 0 Then
        CargoCamposDesdeBD prmIDGasto
        If vsPago.Rows = 1 Then
            If Val(tIDCompra.Tag) <> 0 And MnuNuevo.Enabled = True Then AccionNuevo
            bEsNuevo = True
        End If
    End If
    
End Sub

Private Sub LoadME()
On Error Resume Next
    Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
        
    bEsNuevo = False
    sNuevo = False: sModificar = False
    CargoDatosCombo
    
    LimpioFicha
    LimpioCamposPago
    
    DeshabilitoIngreso
    Botones False, False, False, False, False, Toolbar1, Me

End Sub

Private Sub CargoDatosCombo()

    On Error Resume Next
    
    'Cargo los valores para los comprobantes de pago
    cComprobante.Clear
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.Compracontado)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.Compracontado
    
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraEntradaCaja)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraEntradaCaja
    
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaCredito)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraNotaCredito
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaDevolucion)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraNotaDevolucion
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraReciboDePago)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraReciboDePago
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraRecibo)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraRecibo
    
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraSalidaCaja)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraSalidaCaja
            
    'Cargo las monedas en el combo
    cons = "Select MonCodigo, MonSigno from Moneda"
    CargoCombo cons, cMoneda
    
    cons = "Select DisID, DisNombre from Disponibilidad Order by DisNombre"
    CargoCombo cons, cDisponibilidad
    
    cChTipo.Clear
    cChTipo.AddItem "Cheque": cChTipo.ItemData(cChTipo.NewIndex) = 1
    cChTipo.AddItem "Orden": cChTipo.ItemData(cChTipo.NewIndex) = 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub Label1_Click()
    Foco tProveedor
End Sub

Private Sub Label10_Click()
    Foco tIDCompra
End Sub

Private Sub Label2_Click()
    Foco tImporteP
End Sub

Private Sub Label4_Click()
    Foco cComprobante
End Sub

Private Sub Label6_Click()
    Foco tPesos
End Sub

Private Sub Label9_Click()
    Foco tFecha
End Sub

Private Sub lTotalGasto_Change()
    On Error Resume Next
    If Val(lTotalGasto.Caption) < 0 Then
        fCon.Caption = "Con Qué se cobra"
        fQue.Caption = "Que se cobra"
        Me.Caption = "Con Qué Cobra"
    Else
        fCon.Caption = "Con Qué se paga"
        fQue.Caption = "Que se paga"
        Me.Caption = "Con Qué Paga"
    End If
End Sub

Private Sub MnuBx_Click(Index As Integer)

On Error Resume Next

    If Not AccionCambiarBase(MnuBx(Index).Tag, MnuBx(Index).Caption) Then Exit Sub
    Screen.MousePointer = 11
    
    CargoParametrosImportaciones
    CargoParametrosSucursal
    LoadME
   
    'Cambio el Color del fondo de controles ----------------------------------------------------------------------------------------
    prmColorBase Trim(MnuBx(Index).Tag)
    '-------------------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 0
    
End Sub

Public Function prmColorBase(keyConn As String)
    On Error GoTo errColor
    For I = MnuBx.LBound To MnuBx.UBound
        If LCase(Trim(MnuBx(I).Tag)) = LCase(keyConn) Then
            Dim arrC() As String
            arrC = Split(MnuBases.Tag, "|")
            If arrC(I) <> "" Then Me.BackColor = arrC(I) Else Me.BackColor = vbButtonFace
            Exit For
        End If
    Next
    
    fQue.BackColor = Me.BackColor
    fCon.BackColor = Me.BackColor

errColor:
End Function


Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub

Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub AccionNuevo()
   
Dim aImporte As Currency
    
    On Error Resume Next
    sNuevo = True
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoIngreso
    
    'If cMoneda.ItemData(cMoneda.ListIndex) <> paMonedaPesos Then tPesos.Enabled = True: tPesos.BackColor = Obligatorio
    
    aImporte = Abs(CCur(tImporte.Text))
    If Trim(tIva.Text) <> "" Then aImporte = aImporte + Abs(CCur(tIva.Text))
    If Trim(tCofis.Text) <> "" Then aImporte = aImporte + Abs(CCur(tCofis.Text))
    tImporteP.Text = Format(aImporte, FormatoMonedaP)
    If Not Me.Visible Then Me.Show
    Foco tImporteP
    
End Sub

Private Sub AccionModificar()
    
    Exit Sub
    sModificar = True
    
    Botones False, False, False, True, True, Toolbar1, Me
    HabilitoIngreso
        
End Sub

Private Sub AgregoPago()

Dim aSuma, aImporte As Currency

    On Error GoTo errValidar
    If Not ValidoCamposPago Then Exit Sub
    
    'Si es cheque hay que validar que no este ingresado en la lista !!!
    aSuma = 0: aImporte = 0
    With vsPago
        
        For I = 1 To .Rows - 1      '---------------------------------------------------------------------------
            aSuma = aSuma + .Cell(flexcpValue, I, 0)
            If .Cell(flexcpData, I, 1) = cDisponibilidad.ItemData(cDisponibilidad.ListIndex) Then
                If Trim(.Cell(flexcpText, I, 3)) = "" Then
                    MsgBox "La disponibilidad ya está ingersada. No podrá ingresar doble registro.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                End If
                'Veo si es el mismo cheque
                If cChTipo.ItemData(cChTipo.ListIndex) = .Cell(flexcpData, .Rows - 1, 3) And .Cell(flexcpText, .Rows - 1, 4) = Trim(Trim(tChSerie.Text) & " " & Trim(tChNumero.Text)) Then
                    MsgBox "El comprobante de pago ya está ingresado. No podrá ingresar doble registro.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                End If
            End If
        Next                            '---------------------------------------------------------------------------
        
        aImporte = Abs(CCur(tImporte.Text))
        If Trim(tIva.Text) <> "" Then aImporte = aImporte + Abs(CCur(tIva.Text))
        If Trim(tCofis.Text) <> "" Then aImporte = aImporte + Abs(CCur(tCofis.Text))
        If aSuma > aImporte Then
            MsgBox "La suma de los pagos supera al importe del comprobante en " & aSuma - aImporte & Trim(cMoneda.Text) & ". Verifique", vbExclamation, "ATENCIÓN"
            Foco tImporteP: Exit Sub
        End If
        
        'Si es cheque valido que no se exeda en el importe del cheque --> si existe!!
        If Val(tChNumero.Tag) <> 0 Then
            Dim RsCh As rdoResultset
            Dim aValorCheque As Currency
            Dim aValorQMuevo As Currency
            If Trim(tImporteD.Text) <> "" Then aValorQMuevo = CCur(tImporteD.Text) Else aValorQMuevo = CCur(tImporteP.Text)
            
            cons = "Select Sum(CPaImporte) from ChequePago Where CPaIdCheque = " & Val(tChNumero.Tag)
            Set RsCh = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not RsCh.EOF Then
                If Not IsNull(RsCh(0)) Then
                    If RsCh(0) + aValorQMuevo > CCur(tChImporte.Text) Then
                        MsgBox "El importe ingresado supera el saldo disponible en el cheque." & Chr(vbKeyReturn) & Chr(vbKeyReturn) _
                            & "Importe Original del Cheque: " & Trim(tChImporte.Text) & Chr(vbKeyReturn) _
                            & "Importe Disponible para Asignar: " & Format(RsCh(0) - CCur(tChImporte.Text), FormatoMonedaP) & Chr(vbKeyReturn), vbExclamation, "Importe Incorrecto."
                        RsCh.Close: Exit Sub
                    End If
                End If
            End If
            RsCh.Close
        End If
        
        .AddItem Format(tImporteP.Text, FormatoMonedaP)
        .Cell(flexcpForeColor, .Rows - 1, 0) = vbWhite: .Cell(flexcpBackColor, .Rows - 1, 0) = &H80&
        
        .Cell(flexcpText, .Rows - 1, 1) = cDisponibilidad.Text
        aValor = cDisponibilidad.ItemData(cDisponibilidad.ListIndex): .Cell(flexcpData, .Rows - 1, 1) = aValor
        
        'Importe de la disponibilidad
        If Trim(tImporteD.Text) <> "" Then .Cell(flexcpText, .Rows - 1, 2) = tImporteD.Text Else .Cell(flexcpText, .Rows - 1, 2) = tImporteP.Text
        
        .Cell(flexcpText, .Rows - 1, 3) = Trim(Trim(tChSerie.Text) & " " & Trim(tChNumero.Text))
        aValor = Val(tChNumero.Tag): .Cell(flexcpData, .Rows - 1, 3) = aValor
        
        .Cell(flexcpText, .Rows - 1, 4) = Format(tChImporte.Text, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 5) = Format(tChLibrado.Text, "dd/mm/yyyy")
        .Cell(flexcpText, .Rows - 1, 6) = Format(tChVence.Text, "dd/mm/yyyy")
        
        'Importe de la disponibilidad
        If IsNumeric(tPesos.Text) Then .Cell(flexcpText, .Rows - 1, 7) = Format(tPesos.Text, FormatoMonedaP) Else .Cell(flexcpText, .Rows - 1, 7) = tImporteP.Text
        
    End With
    
    cDisponibilidad.Text = "": tPesos.Text = ""
    tImporteP.Text = Format(aImporte - aSuma - CCur(tImporteP.Text), FormatoMonedaP)
    If CCur(tImporteP.Text) = 0 Then
        tImporteP.Text = ""
        AccionGrabar
    Else
        Foco tImporteP
    End If
    Exit Sub
    
errValidar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al validar los datos.", Err.Description
End Sub

Private Sub AccionGrabar()
Dim aMsgErr As String
    
    aMsgErr = ""
    If Not ValidoCampos Then Exit Sub
    
    If MsgBox("Confirma almacenar la información ingresada", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    On Error GoTo errorBT
    Screen.MousePointer = 11
    FechaDelServidor
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    'Veo fecha de modificacion
    cons = "Select * from Compra Where ComCodigo = " & CLng(tIDCompra.Text)
    Set RsCom = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If gFModificacion <> RsCom!ComFModificacion Then
        aMsgErr = "La ficha de compra ha sido modificada desde otra terminal. Para grabar vuelva a cargar los datos."
        RsCom.Close: GoTo errorET: Exit Sub
    Else
        RsCom.Edit
        RsCom!ComFModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsCom.Update: RsCom.Close
    End If
    
    GraboDatosBD CLng(tIDCompra.Text)
    
    cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
        
    sNuevo = False: sModificar = False
    DeshabilitoIngreso
    Botones False, True, True, False, False, Toolbar1, Me
    Foco tIDCompra
    gFModificacion = gFechaServidor
    Screen.MousePointer = 0
    
    If bEsNuevo Then
        If MsgBox("Desea volver a la pantalla que invocó al formulario.", vbQuestion + vbYesNo, "Salir del formulario") = vbYes Then Unload Me Else bEsNuevo = False
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
    If Trim(aMsgErr) = "" Then aMsgErr = "No se ha podido inicializar la transacción. Reintente la operación."
    Screen.MousePointer = 0
    clsGeneral.OcurrioError aMsgErr, Err.Description
End Sub

Private Sub GraboDatosBD(Compra As Long)
Dim aIDMovimiento As Long, aIDCheque As Long
Dim RsMov As rdoResultset

    'Cargo los Movimientos de disponibilidades (Renglones, Cheques y  Relaciones de Pagos de Cheques)
    cons = "Select * from MovimientoDisponibilidad Where MDiID = 0"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    rsAux.AddNew
    rsAux!MDiFecha = Format(tFecha.Text, sqlFormatoF)
    rsAux!MDiHora = Format(gFechaServidor, "hh:mm:ss")
    rsAux!MDiTipo = paMDPagoDeCompra
    rsAux!MDiIdCompra = Compra
    rsAux.Update: rsAux.Close
    
    cons = "Select Max(MDiID) from MovimientoDisponibilidad"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    aIDMovimiento = rsAux(0)
    rsAux.Close
    
    'Movimientos de Disponibilidades
    cons = "Select * from MovimientoDisponibilidadRenglon Where MDRIDMovimiento = " & aIDMovimiento
    Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    With vsPago
    For I = 1 To .Rows - 1
        RsMov.AddNew
        RsMov!MDRIdMovimiento = aIDMovimiento
        RsMov!MDRIdDisponibilidad = .Cell(flexcpData, I, 1)
        
        If Trim(.Cell(flexcpText, I, 3)) = "" Then
            RsMov!MDRIdCheque = 0
        Else
            'Se paga con un cheque----------------------------------------------
            If .Cell(flexcpData, I, 3) <> 0 Then    'El cheque ya existe
                RsMov!MDRIdCheque = .Cell(flexcpData, I, 3)
                aIDCheque = .Cell(flexcpData, I, 3)
            Else
                'Inserto el Cheque
                cons = "Select * from Cheque Where CheID = 0"
                Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                rsAux.AddNew
                rsAux!CheIDDisponibilidad = .Cell(flexcpData, I, 1)
                If InStr(.Cell(flexcpText, I, 3), " ") <> 0 Then rsAux!CheSerie = Trim(Mid(.Cell(flexcpText, I, 3), 1, InStr(.Cell(flexcpText, I, 3), " ") - 1))
                rsAux!CheNumero = Mid(.Cell(flexcpText, I, 3), InStr(.Cell(flexcpText, I, 3), " ") + 1, Len(.Cell(flexcpText, I, 3)))
                rsAux!CheImporte = CCur(.Cell(flexcpText, I, 4))
                If IsDate(.Cell(flexcpText, I, 5)) Then rsAux!CheLibrado = Format(.Cell(flexcpText, I, 5), sqlFormatoF)
                If IsDate(.Cell(flexcpText, I, 6)) Then rsAux!CheVencimiento = Format(.Cell(flexcpText, I, 6), sqlFormatoF)
                rsAux.Update: rsAux.Close
                
                cons = "Select Max(CheID) from Cheque"
                Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                aIDCheque = rsAux(0)
                rsAux.Close
                
                RsMov!MDRIdCheque = aIDCheque
            End If
            
            'Inserto Relacion de Pago
            cons = " Insert Into ChequePago (CPaIDCheque, CPaIDCompra, CPaImporte) " _
                    & " Values(" & aIDCheque & ", " & Compra & ", " & CCur(.Cell(flexcpText, I, 2)) & ")"
            cBase.Execute cons
        End If
        
        Select Case cComprobante.ItemData(cComprobante.ListIndex)       'Lo que se mueve la disponibilidad
            Case TipoDocumento.CompraNotaCredito, TipoDocumento.CompraNotaDevolucion, TipoDocumento.CompraEntradaCaja
                    RsMov!MDRDebe = CCur(.Cell(flexcpText, I, 2))
            Case Else:
                    RsMov!MDRHaber = CCur(.Cell(flexcpText, I, 2))
        End Select
        
        RsMov!MDRImporteCompra = CCur(.Cell(flexcpText, I, 0))    'Relacion en importe de la compra
        RsMov!MDRImportePesos = CCur(.Cell(flexcpText, I, 7))       'Relacion en pesos
        RsMov.Update
    Next
    End With
    RsMov.Close
    
End Sub

Private Sub AccionEliminar()
    
    If Not PidoSucesoEliminar Then Exit Sub
    
    If MsgBox("Confirma eliminar el pago ingresado.", vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Exit Sub
    
    On Error GoTo errorBT
    FechaDelServidor
    Dim RsRen As rdoResultset, RsChe As rdoResultset
    Screen.MousePointer = 11
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    cons = "Select * from MovimientoDisponibilidad Where MDiIDCompra = " & tIDCompra.Tag
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        
        cons = "Select * from MovimientoDisponibilidadRenglon " _
               & " Where MDRIDMovimiento = " & rsAux!MDiID _
               & " And MDRIDCheque > 0"
        Set RsRen = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsRen.EOF
            
            cons = "Delete ChequePago Where CPaIDCheque = " & RsRen!MDRIdCheque & " And CPaIDCompra = " & tIDCompra.Tag
            cBase.Execute cons
            
            'Si no hay mas relaciones de pago borro el cheque
            cons = "Select * from ChequePago Where CPaIDCheque = " & RsRen!MDRIdCheque
            Set RsChe = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If RsChe.EOF Then
                cons = "Delete Cheque Where CheId = " & RsRen!MDRIdCheque
                cBase.Execute cons
            End If
            RsChe.Close
            
            RsRen.MoveNext
        Loop
        RsRen.Close
        
        'Borro los renglones de movimientos
        cons = "Delete MovimientoDisponibilidadRenglon Where MDRIDMovimiento = " & rsAux!MDiID
        cBase.Execute cons
        
    End If
    'Borro el movimiento
    cons = "Delete MovimientoDisponibilidad Where MDiID = " & rsAux!MDiID
    cBase.Execute cons
    
    rsAux.Close
    
    
    clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, dSuceso.Tipo, paCodigoDeTerminal, _
                dSuceso.Usuario, 0, _
                Descripcion:=dSuceso.Titulo, Defensa:=dSuceso.Defensa, _
                Valor:=dSuceso.Valor, idAutoriza:=dSuceso.Autoriza

    cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------
    
    DeshabilitoIngreso
    CargoCamposDesdeBD CLng(tIDCompra.Tag)
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
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
End Sub

Sub AccionCancelar()

Dim aCompra As Long
    
    On Error Resume Next
    aCompra = CLng(tIDCompra.Tag)
    DeshabilitoIngreso
    
    LimpioFicha
    LimpioCamposPago
    
    CargoCamposDesdeBD aCompra
    
    sNuevo = False: sModificar = False
    
End Sub

Private Sub AccionListaDeAyuda(Consulta As String)

    On Error GoTo errAyuda
    Screen.MousePointer = 11
    
    Dim aLista As New clsListadeAyuda
    Dim aSeleccionado As Long: aSeleccionado = 0
            
    aLista.ActivoListaAyudaSQL cBase, Consulta
    Me.Refresh: DoEvents
    
    If IsNumeric(aLista.ItemSeleccionadoSQL) Then aSeleccionado = CLng(aLista.ItemSeleccionadoSQL)
    Set aLista = Nothing
    
    If aSeleccionado <> 0 Then LimpioFicha: CargoCamposDesdeBD aSeleccionado
    
    Screen.MousePointer = 0
    Exit Sub
        
errAyuda:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al activar la lista de ayuda.", Err.Description
End Sub

Private Sub tChImporte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then AgregoPago
End Sub

Private Sub tChLibrado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tChVence
End Sub

Private Sub tChNumero_Change()
    DeshabilitoCamposCheque
End Sub

Private Sub tChNumero_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If cChTipo.ListIndex = -1 Then
            MsgBox "Debe seleccionar el tipo de documento para realizar la salida de la cuenta.", vbExclamation, "ATENCIÓN"
            Foco cChTipo: Exit Sub
        End If
        If tChSerie.Enabled And Trim(tChSerie.Text) = "" Then
            MsgBox "Debe ingresar el número de serie del cheque para buscarlo en la base de datos.", vbExclamation, "ATENCIÓN"
            Foco tChSerie: Exit Sub
        End If
        If Not IsNumeric(tChNumero.Text) Then
            MsgBox "Debe ingresar el número del cheque para buscarlo en la base de datos.", vbExclamation, "ATENCIÓN"
            Foco tChNumero: Exit Sub
        End If
        
        'Hay que buscar en las tablas de cheques para ver si está ingresado
        cons = "Select * from Cheque " _
                & " Where CheIDDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
                & " And CheSerie = '" & Trim(tChSerie.Text) & "'" _
                & " And CheNumero = " & Trim(tChNumero.Text)
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            tChNumero.Tag = rsAux!CheID
            tChImporte.Text = Format(rsAux!CheImporte, FormatoMonedaP)
            If Not IsNull(rsAux!CheLibrado) Then tChLibrado.Text = Format(rsAux!CheLibrado, "dd/mm/yyyy")
            If Not IsNull(rsAux!CheVencimiento) Then tChLibrado.Text = Format(rsAux!CheVencimiento, "dd/mm/yyyy")
        Else
            tChNumero.Tag = 0
            tChImporte.Enabled = True: tChImporte.BackColor = Obligatorio
            tChVence.Enabled = True: tChVence.BackColor = Blanco
            tChLibrado.Enabled = True: tChLibrado.BackColor = Obligatorio: tChLibrado.Text = Format(Now, "dd/mm/yyyy")
        End If
        rsAux.Close
        
        If tChLibrado.Enabled Then Foco tChLibrado: Exit Sub
        AgregoPago 'El cheque existe, agrego a la lista
        
    End If
    
End Sub

Private Sub tChSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then Foco tChNumero
    
End Sub

Private Sub tChVence_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tImporteD.Enabled Then tChImporte.Text = tImporteD.Text Else tChImporte.Text = tImporteP.Text
        Foco tChImporte
    End If
End Sub

Private Sub tFecha_GotFocus()
    With tFecha: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tIDCompra_Change()
    If Val(tIDCompra.Tag) <> 0 Then Botones False, False, False, False, False, Toolbar1, Me
    tIDCompra.Tag = 0
End Sub

Private Sub tIDCompra_GotFocus()
    With tIDCompra: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tIDCompra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Val(tIDCompra.Tag) = 0 And IsNumeric(tIDCompra.Text) Then CargoCamposDesdeBD CLng(tIDCompra.Text) Else Foco tFecha
    End If
End Sub

Private Sub tFecha_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 And Not sNuevo And Not sModificar Then Call bBuscar1_Click
    
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tProveedor: Exit Sub
End Sub


Private Sub tImporteP_KeyPress(KeyAscii As Integer)
    
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tImporteP.Text) Then tImporteP.Text = Format(tImporteP.Text, FormatoMonedaP)
        
        If cMoneda.ItemData(cMoneda.ListIndex) <> paMonedaPesos Then
            tPesos.Text = Format(CCur(tImporteP.Text) * CCur(tPesos.Tag), FormatoMonedaP)
            lTC.Caption = "TC de la Compra: " & Format(tPesos.Tag, "#,##0.000")
        End If
        Foco cDisponibilidad
    End If
    
End Sub

Private Sub tImporteP_LostFocus()

    On Error Resume Next
    If Not sNuevo Then Exit Sub
    If IsNumeric(tImporteP.Text) Then tImporteP.Text = Format(tImporteP.Text, FormatoMonedaP)
        
    If cMoneda.ItemData(cMoneda.ListIndex) <> paMonedaPesos Then
        tPesos.Text = Format(CCur(tImporteP.Text) * CCur(tPesos.Tag), FormatoMonedaP)
        lTC.Caption = "TC de la Compra: " & Format(tPesos.Tag, "#,##0.000")
    End If


End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bBuscar2.SetFocus
End Sub

Private Sub tPesos_Change()
    lTC.Caption = ""
End Sub

Private Sub tPesos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tPesos.Text) Then tPesos.Text = Format(tPesos.Text, FormatoMonedaP)
        Foco cDisponibilidad
    End If
End Sub

Private Sub tProveedor_Change()
    tProveedor.Tag = 0
End Sub

Private Sub tProveedor_GotFocus()
    With tProveedor: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 And Val(tProveedor.Tag) <> 0 Then Call bBuscar1_Click
End Sub


Private Sub tProveedor_KeyPress(KeyAscii As Integer)
    
    On Error GoTo errBuscar
    
    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) <> 0 Or Trim(tProveedor.Text) = "" Then Foco tIDCompra: Exit Sub
        Screen.MousePointer = 11
        cons = "Select PClCodigo, PClFantasia as 'Nombre Fantasía', PClNombre as 'Razón Social' from ProveedorCliente " _
                & " Where PClNombre like '" & Trim(tProveedor.Text) & "%' Or PClFantasia like '" & Trim(tProveedor.Text) & "%'"
        
        Dim aLista As New clsListadeAyuda, mSel As Long
        mSel = aLista.ActivarAyuda(cBase, cons, 5500, 1, "Lista de Proveedores")
        Me.Refresh
        If mSel <> 0 Then
            tProveedor.Text = Trim(aLista.RetornoDatoSeleccionado(1))
            tProveedor.Tag = aLista.RetornoDatoSeleccionado(0)
        Else
            tProveedor.Text = ""
        End If
        Set aLista = Nothing
        Screen.MousePointer = 0
    End If
    Exit Sub
    Screen.MousePointer = 0

errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al procesar la lista de ayuda.", Err.Description
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
    End Select

End Sub

Private Function ValidoCampos() As Boolean

    ValidoCampos = False
    
    With vsPago
    
    Dim aSuma, aImporte As Currency: aSuma = 0: aImporte = 0
    
    For I = 1 To .Rows - 1
        aSuma = aSuma + CCur(.Cell(flexcpText, I, 0))
    Next
    aImporte = Abs(CCur(tImporte.Text))
    If Trim(tIva.Text) <> "" Then aImporte = aImporte + Abs(CCur(tIva.Text))
    If Trim(tCofis.Text) <> "" Then aImporte = aImporte + Abs(CCur(tCofis.Text))
    
    If aImporte <> aSuma Then
        MsgBox "El importe del comprobante (" & Format(aImporte, FormatoMonedaP) & ") no coincide con la suma de los pagos (" & Format(aSuma, FormatoMonedaP) & ")", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    End With
    
    'Valido que no exista un con que se paga ingresado
    Dim bHay As Boolean: bHay = False
    cons = "Select Count(*) from MovimientoDisponibilidad" _
            & " Where MDiIDCompra = " & tIDCompra.Tag
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux(0)) Then If rsAux(0) > 0 Then bHay = True
    End If
    rsAux.Close
    
    If bHay Then
        MsgBox "Ya existe un con que se paga para el gasto seleccionado." & vbCrLf & _
                    "Vuelva a cargar los datos.", vbExclamation, "Hay Pago Ingresado !!!"
        Screen.MousePointer = 0: Exit Function
    End If
    
    ValidoCampos = True
    
End Function

Private Sub DeshabilitoIngreso()
        
    tProveedor.Enabled = True: tProveedor.BackColor = Blanco
    tFecha.Enabled = True: tFecha.BackColor = Blanco
    tIDCompra.Enabled = True: tIDCompra.BackColor = Blanco
    
    cComprobante.Enabled = True: cComprobante.BackColor = Blanco: cComprobante.ForeColor = &HC00000
    tSerie.Enabled = True: tSerie.BackColor = Blanco
    tNumero.Enabled = True: tNumero.BackColor = Blanco
    
    cMoneda.Enabled = False: cMoneda.BackColor = Inactivo
    tImporte.Enabled = False: tImporte.BackColor = Inactivo
    tIva.Enabled = False: tIva.BackColor = Inactivo
    tCofis.Enabled = False: tCofis.BackColor = Inactivo
    
    bBuscar1.Enabled = True: bBuscar2.Enabled = True
            
    vsPago.BackColor = Inactivo
    vsPago.Editable = False
    
    'Campos de Pago
    tImporteP.Enabled = False: tImporteP.BackColor = Inactivo
    tPesos.Enabled = False: tPesos.BackColor = Inactivo
    cDisponibilidad.Enabled = False: cDisponibilidad.BackColor = Inactivo
    
    tImporteD.Enabled = False: tImporteD.BackColor = Inactivo
    cChTipo.Enabled = False: cChTipo.BackColor = Inactivo
    tChSerie.Enabled = False: tChSerie.BackColor = Inactivo
    tChNumero.Enabled = False: tChNumero.BackColor = Inactivo
    tChVence.Enabled = False: tChVence.BackColor = Inactivo
    tChLibrado.Enabled = False: tChLibrado.BackColor = Inactivo
    tChImporte.Enabled = False: tChImporte.BackColor = Inactivo
    
End Sub

Private Sub HabilitoIngreso()

    tProveedor.Enabled = False: tProveedor.BackColor = Inactivo
    tFecha.Enabled = False: tFecha.BackColor = Inactivo
    tIDCompra.Enabled = False: tIDCompra.BackColor = Inactivo
    
    cComprobante.Enabled = False: cComprobante.BackColor = Inactivo
    tSerie.Enabled = False: tSerie.BackColor = Inactivo
    tNumero.Enabled = False: tNumero.BackColor = Inactivo
    
    bBuscar1.Enabled = False: bBuscar2.Enabled = False
        
    vsPago.BackColor = Blanco
    
    tImporteP.Enabled = True: tImporteP.BackColor = Obligatorio
    cDisponibilidad.Enabled = True: cDisponibilidad.BackColor = Obligatorio
    
End Sub

Private Sub HabilitoIngresoCheque()

    cChTipo.Enabled = True: cChTipo.BackColor = Obligatorio
    tChSerie.Enabled = True: tChSerie.BackColor = Obligatorio
    tChNumero.Enabled = True: tChNumero.BackColor = Obligatorio
    tChLibrado.Enabled = True: tChLibrado.BackColor = Obligatorio
    tChVence.Enabled = True: tChVence.BackColor = Blanco
    tChImporte.Enabled = True: tChImporte.BackColor = Obligatorio
        
End Sub

Private Sub CargoCamposDesdeBD(aCompra As Long)
Dim rs1 As rdoResultset

    On Error GoTo errCargo
    If aCompra = 0 Then Exit Sub
    
    Screen.MousePointer = 11
    LimpioFicha
    
    cons = "Select * from Compra, ProveedorCliente " _
               & " Where ComCodigo = " & aCompra _
               & " And ComProveedor = PClCodigo" _
               & " And ComTipoDocumento In (" & aTiposDocs & ")"
    Set RsCom = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsCom.EOF Then
        MsgBox "No existe una compra para el código ingresado." & Chr(vbKeyReturn) & "(*) Contado, nota de devolución, nota de crédito, recibo de pago, recibo provisorio, ingreso o salida de caja.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        RsCom.Close: Exit Sub
    End If
        
    tIDCompra.Text = Format(RsCom!ComCodigo, "#,###,##0")
    tIDCompra.Tag = RsCom!ComCodigo
    tFecha.Text = Format(RsCom!ComFecha, FormatoFP)
    tProveedor.Text = Trim(RsCom!PClFantasia)
    tProveedor.Tag = RsCom!ComProveedor
    
    BuscoCodigoEnCombo cComprobante, RsCom!ComTipoDocumento
    If Not IsNull(RsCom!ComSerie) Then tSerie.Text = Trim(RsCom!ComSerie)
    If Not IsNull(RsCom!ComNumero) Then tNumero.Text = Trim(RsCom!ComNumero)
    
    BuscoCodigoEnCombo cMoneda, RsCom!ComMoneda
    If Not IsNull(RsCom!ComIva) Then tIva.Text = Format(RsCom!ComIva, FormatoMonedaP)
    If Not IsNull(RsCom!ComCofis) Then tCofis.Text = Format(RsCom!ComCofis, FormatoMonedaP)
    tImporte.Text = Format(RsCom!ComImporte, FormatoMonedaP)
    
    Dim aTotal As Currency
    aTotal = CCur(tImporte.Text)
    If Trim(tIva.Text) <> "" Then aTotal = aTotal + CCur(tIva.Text)
    If Trim(tCofis.Text) <> "" Then aTotal = aTotal + CCur(tCofis.Text)
    lTotalGasto.Caption = Format(aTotal, FormatoMonedaP)
    
    If Not IsNull(RsCom!ComTC) Then tPesos.Tag = RsCom!ComTC Else tPesos.Tag = 1
    
    gFModificacion = RsCom!ComFModificacion

    RsCom.Close
       
    'Cargo los pagos ingresados para el comprobante
    cons = "Select * from MovimientoDisponibilidad, MovimientoDisponibilidadRenglon, Disponibilidad" _
            & " Where MDiIDCompra = " & tIDCompra.Tag _
            & " And MDiID = MDRIDMovimiento " _
            & " And MDRIDDisponibilidad = DisID"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    With vsPago
    
    If Not rsAux.EOF Then
        .Rows = 1
        Do While Not rsAux.EOF
            .AddItem Format(rsAux!MDRImporteCompra, FormatoMonedaP)
            .Cell(flexcpForeColor, .Rows - 1, 0) = vbWhite: .Cell(flexcpBackColor, .Rows - 1, 0) = &H80&
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!DisNombre)
            aValor = rsAux!DisID: .Cell(flexcpData, .Rows - 1, 1) = aValor
            
            'Importe de la disponibilidad
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!MDRHaber, FormatoMonedaP)
            
            If Not IsNull(rsAux!DisSucursal) Then       'Si la Disponibilidad es bancaria
                'Busco el Cheque
                cons = "Select * from Cheque Where CheID = " & rsAux!MDRIdCheque
                Set rs1 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                If Not rs1.EOF Then
                    .Cell(flexcpText, .Rows - 1, 3) = Trim(rs1!CheSerie) & " " & rs1!CheNumero
                     aValor = rs1!CheID: .Cell(flexcpData, .Rows - 1, 3) = aValor
                     
                     .Cell(flexcpText, .Rows - 1, 4) = Format(rs1!CheImporte, FormatoMonedaP)
                     .Cell(flexcpText, .Rows - 1, 5) = Format(rs1!CheLibrado, "dd/mm/yyyy")
                     If Not IsNull(rs1!CheVencimiento) Then .Cell(flexcpText, .Rows - 1, 6) = Format(rs1!CheVencimiento, "dd/mm/yyyy")
                End If
                rs1.Close
            End If
                
            .Cell(flexcpText, .Rows - 1, 7) = rsAux!MDRImportePesos
            rsAux.MoveNext
        Loop
        Botones False, True, True, False, False, Toolbar1, Me
    Else
        Select Case cComprobante.ItemData(cComprobante.ListIndex)
            Case TipoDocumento.Compracontado, TipoDocumento.CompraRecibo, TipoDocumento.CompraReciboDePago, TipoDocumento.CompraNotaDevolucion, TipoDocumento.CompraNotaCredito, TipoDocumento.CompraEntradaCaja, TipoDocumento.CompraSalidaCaja
                    Botones True, False, False, False, False, Toolbar1, Me
            Case Else: Botones False, False, False, False, False, Toolbar1, Me
        End Select
    End If
    rsAux.Close
    
    End With
    Screen.MousePointer = 0
    Exit Sub

errCargo:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos de la compra.", Err.Description
End Sub

Private Sub LimpioFicha()

    tIDCompra.Text = ""
    tFecha.Text = ""
    tProveedor.Text = ""
    
    cMoneda.Text = "": tImporte.Text = "": tIva.Text = "": tCofis.Text = ""
    cComprobante.Text = "": tSerie.Text = "": tNumero.Text = ""
    
    lTotalGasto.Caption = ""
    vsPago.Rows = 1
    
End Sub

Private Sub LimpioCamposPago()
    tImporteP.Text = ""
    tPesos.Text = "": lTC.Caption = ""
    cDisponibilidad.Text = ""
    tImporteD.Text = ""
    cChTipo.Text = "": tChSerie.Text = "": tChNumero.Text = ""
    tChLibrado.Text = "": tChVence.Text = "": tChImporte.Text = ""
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsPago
        .Rows = 1: .Cols = 1
        .Editable = False
        .FormatString = ">Importe|Disponibilidad|>$ Disponibilidad|<Comprobante|>Importe|Librado|Vence|Pesos"
            
        .WordWrap = False
        .ColHidden(7) = True
        .ColWidth(0) = 1200: .ColWidth(1) = 2100: .ColWidth(2) = 1200: .ColWidth(4) = 1200: .ColWidth(5) = 950
    End With
    
End Sub

Private Sub tSerie_GotFocus()
    With tSerie: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tNumero
End Sub

Private Sub tImporteD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsNumeric(tImporteD.Text) Then tImporteD.Text = Format(tImporteD.Text, FormatoMonedaP)
        If cChTipo.Enabled Then Foco cChTipo: Exit Sub
        AgregoPago
    End If
End Sub

Private Sub vsPago_KeyDown(KeyCode As Integer, Shift As Integer)

    If Not sNuevo And Not sModificar Then Exit Sub
    If KeyCode = vbKeyDelete Then
        If vsPago.Row > 0 Then vsPago.RemoveItem vsPago.Row
    End If
    
End Sub

Private Sub DeshabilitoCamposCheque(Optional DesdeDisponibilidad As Boolean = False, Optional DesdeTipo As Boolean = False)

    If DesdeDisponibilidad Then
        tImporteD.Text = "": tImporteD.Enabled = False: tImporteD.BackColor = Inactivo
        cChTipo.Text = "": cChTipo.Enabled = False: cChTipo.BackColor = Inactivo
        tChNumero.Text = "": tChNumero.Enabled = False: tChNumero.BackColor = Inactivo
        tChSerie.Text = "": tChSerie.Enabled = False: tChSerie.BackColor = Inactivo
    End If
    tChNumero.Tag = 0   'ID del cheque
    
    If DesdeTipo Then
        tChSerie.Text = "": tChSerie.Enabled = False: tChSerie.BackColor = Inactivo
        tChNumero.Text = "": tChNumero.Enabled = False: tChNumero.BackColor = Inactivo
    End If
    
    tChVence.Text = "": tChVence.Enabled = False: tChVence.BackColor = Inactivo
    tChLibrado.Text = "": tChLibrado.Enabled = False: tChLibrado.BackColor = Inactivo
    tChImporte.Text = "": tChImporte.Enabled = False: tChImporte.BackColor = Inactivo

End Sub

Private Function ValidoCamposPago() As Boolean

    ValidoCamposPago = False
    
    If Not IsNumeric(tImporteP.Text) Then
        MsgBox "El importe de pago ingresado no es correcto. Verifique", vbExclamation, "ATENCIÓN"
        Foco tImporteP: Exit Function
    End If
    If tPesos.Enabled And Not IsNumeric(tPesos.Text) Then
        MsgBox "Ingrese el equivalente en pesos para establecer una relación. Verifique", vbExclamation, "ATENCIÓN"
        Foco tPesos: Exit Function
    End If

    If CCur(tImporteP.Text) < 0 Then
        MsgBox "El importe de pago ingresado no es correcto. Verifique", vbExclamation, "ATENCIÓN"
        Foco tImporteP: Exit Function
    End If
    If cDisponibilidad.ListIndex = -1 Then
        MsgBox "Debe seleccionar la disponibilidad para realizar el pago.", vbExclamation, "ATENCIÓN"
        Foco cDisponibilidad: Exit Function
    End If
    If tImporteD.Enabled And Not IsNumeric(tImporteD.Text) Then
        MsgBox "Debe ingresar el importe en la moneda de la disponibilidad para establecer una relación entre las cuentas.", vbExclamation, "ATENCIÓN"
        Foco tImporteD: Exit Function
    End If
    If cChTipo.Enabled And cChTipo.ListIndex = -1 Then
        MsgBox "Debe seleccionar el tipo de comprobante para realizar el pago.", vbExclamation, "ATENCIÓN"
        Foco cChTipo: Exit Function
    End If
    If tChSerie.Enabled And Trim(tChSerie.Text) = "" Then
        MsgBox "Debe ingresar la serie del comprobante de pago.", vbExclamation, "ATENCIÓN"
        Foco tChSerie: Exit Function
    End If
    If tChNumero.Enabled And Not IsNumeric(tChNumero.Text) Then
        MsgBox "Debe ingresar el número del comprobante de pago.", vbExclamation, "ATENCIÓN"
        Foco tChNumero: Exit Function
    End If
    If tChLibrado.Enabled And Not IsDate(tChLibrado.Text) Then
        MsgBox "Debe ingresar la fecha de librado del comprobante de pago.", vbExclamation, "ATENCIÓN"
        Foco tChLibrado: Exit Function
    End If
    If tChVence.Enabled And Not IsDate(tChVence.Text) And (tChVence.Text) <> "" Then
        MsgBox "La fecha de vencimiento del comprobante de pago no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tChVence: Exit Function
    End If
    If tChImporte.Enabled And Not IsNumeric(tChImporte.Text) Then
        MsgBox "Debe ingresar el importe total del comprobante de pago.", vbExclamation, "ATENCIÓN"
        Foco tChImporte: Exit Function
    End If

    ValidoCamposPago = True
    
End Function

Private Function PidoSucesoEliminar() As Boolean
    On Error GoTo errSuceso
    PidoSucesoEliminar = False
    
    Dim objSuceso As New clsSuceso
    objSuceso.TipoSuceso = prmSucesoGastos
    objSuceso.ActivoFormulario paCodigoDeUsuario, "Eliminar Con Que se Paga", cBase
    
    Me.Refresh
    With dSuceso
        .Usuario = objSuceso.RetornoValor(Usuario:=True)
        .Defensa = objSuceso.RetornoValor(Defensa:=True)
        .Autoriza = objSuceso.Autoriza
    End With
    
    Set objSuceso = Nothing
    If dSuceso.Usuario = 0 Or Trim(dSuceso.Defensa) = "" Then Exit Function  'Abortó el ingreso del suceso
    
    'Cargo otros datos en la estructura del suceso
    With dSuceso
        .Tipo = prmSucesoGastos
        .Titulo = "Elimina Con Que se Paga  (ID:" & Trim(tIDCompra.Text) & ")"
        .Defensa = Trim(.Defensa)
        .Valor = 0
    End With
    PidoSucesoEliminar = True
    
    With vsPago
        For I = 1 To .Rows - 1
            dSuceso.Defensa = .Cell(flexcpText, I, 1) & " $" & .Cell(flexcpText, I, 2) & vbCrLf & Trim(dSuceso.Defensa)
            dSuceso.Valor = dSuceso.Valor + .Cell(flexcpText, I, 0)
        Next
    
    End With
    
    Screen.MousePointer = 0
    Exit Function
errSuceso:
    clsGeneral.OcurrioError "Error al pedir datos del suceso.", Err.Description
    Screen.MousePointer = 0
End Function
