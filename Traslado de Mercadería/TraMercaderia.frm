VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form TraMercaderia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia de Mercadería"
   ClientHeight    =   7395
   ClientLeft      =   30
   ClientTop       =   615
   ClientWidth     =   8220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TraMercaderia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   8220
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   27
      Top             =   480
      Width           =   1695
   End
   Begin VSPrinter8LibCtl.VSPrinter vsPrint 
      Height          =   2175
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Visible         =   0   'False
      Width           =   4095
      _cx             =   7223
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
      Zoom            =   8.61742424242424
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
   Begin VB.Timer tmServerPrint 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7440
      Top             =   1800
   End
   Begin VB.TextBox tDoc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6120
      TabIndex        =   25
      Top             =   1740
      Width           =   1695
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4215
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7435
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprimir Planilla Auxiliar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "entregar"
            Object.ToolTipText     =   "Emisión y entrega de mercadería "
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "rePrint"
            Object.ToolTipText     =   "Reimprimir documento"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   4400
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir del formulario"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin AACombo99.AACombo cDestino 
      Height          =   315
      Left            =   6060
      TabIndex        =   5
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
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
   Begin AACombo99.AACombo cOrigen 
      Height          =   315
      Left            =   900
      TabIndex        =   1
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
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
   Begin AACombo99.AACombo cIntermediario 
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Top             =   960
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
   End
   Begin AACombo99.AACombo cEstado 
      Height          =   315
      Left            =   6060
      TabIndex        =   11
      Top             =   1320
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
   End
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   7320
      MaxLength       =   3
      TabIndex        =   17
      Top             =   6780
      Width           =   735
   End
   Begin VB.TextBox tCantidad 
      Height          =   285
      Left            =   4680
      MaxLength       =   12
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.ComboBox cArticulo 
      Height          =   315
      Left            =   900
      Style           =   1  'Simple Combo
      TabIndex        =   7
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   6780
      Width           =   7215
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   7140
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11880
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Picture         =   "TraMercaderia.frx":030A
            Key             =   "printer"
         EndProperty
      EndProperty
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraMercaderia.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraMercaderia.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraMercaderia.frx":0848
            Key             =   "Total"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraMercaderia.frx":0B62
            Key             =   "Parcial"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraMercaderia.frx":0E7C
            Key             =   "No"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraMercaderia.frx":1196
            Key             =   "NoDoy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraMercaderia.frx":14B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraMercaderia.frx":15C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraMercaderia.frx":16D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraMercaderia.frx":17E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraMercaderia.frx":18F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TraMercaderia.frx":1C12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblRecepciono 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Recepción: 88/88/8888"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   3480
      TabIndex        =   29
      Top             =   1740
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Remito:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5400
      TabIndex        =   24
      Top             =   1740
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "&Lista"
      Height          =   315
      Left            =   780
      TabIndex        =   12
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Por:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   23
      Top             =   1740
      Width           =   495
   End
   Begin VB.Label labUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2520
      TabIndex        =   22
      Top             =   1740
      Width           =   735
   End
   Begin VB.Label labCreado 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   960
      TabIndex        =   21
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Creado:"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   1740
      Width           =   615
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Usuario:"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   7320
      TabIndex        =   16
      Top             =   6540
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Destino:"
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Estado:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ca&ntidad:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3900
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artículo:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Co&mentario:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6540
      Width           =   7215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "O&rigen:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Camión:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2820
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   1395
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   7935
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu MnuLinea 
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
      Begin VB.Menu MnuLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu MnuEntregar 
         Caption         =   "En&tregar"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Volver al formulario"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "TraMercaderia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cambios:
'   21-7-00 Listview por grilla
'               Dejo entregar mercadería que no hay en stock y registro suceso silencioso por c/artículo.
'               Agregue campo en la tabla local que me indica si tiene terminal, si no tiene ya recepcionó el traslado.

'   22 -7-00 Si es un traslado sin intermediario no dejo modificar el mismo pues me caga el stock, total si no llego la merca que lo borre y uno nuevo.
'   12-2-01  En entrega de mercadería me borraban el intermediario y cagaba.
'   10-9-2003 Cambie el flexcptext x value en la columna 0, además encontre que si salia por fmodificado en modifico traslado no tenía el exit sub.
'                    Le agregue fecha de modificado a todo los que grabo menos a modfiicar que ya estaba.

'   24-3-2006
'               Cargo prm sucursal tomo el nombre del documento traslado (ojo utilizo prm de Contado), además incluyo qRenglonCtdo.

Option Explicit
Private Rs As rdoResultset
Private sNuevo As Boolean, sModificar As Boolean
Private CodTraslado As Long
Private jobnum As Integer       'Nro. de Trabajo para la contado
Private CantForm As Integer    'Cantidad de formulas del reporte

Private EmpresaEmisora As clsClienteCFE
Private TasaBasica As Currency, TasaMinima As Currency

Private Function EmitirCFE(ByVal Documento As clsDocumentoCGSA, ByVal CAE As clsCAEDocumento) As String
On Error GoTo errEC
    With New clsCGSAEFactura
        .URLAFirmar = prmURLFirmaEFactura
        .TasaBasica = TasaBasica
        .TasaMinima = TasaMinima
        .ImporteConInfoDeCliente = prmImporteConInfoCliente
        Set .Connect = cBase
        Dim sResult As String
        sResult = .FirmarUnDocumento(Documento.Codigo)
        If UCase(sResult) <> "TRUE" Then EmitirCFE = sResult
        'If Not .GenerarEComprobante(CAE, Documento, EmpresaEmisora, paCodigoDGI) Then
        '    EmitirCFE = .XMLRespuesta
        'End If
    End With
    Exit Function
errEC:
    EmitirCFE = "Error en firma: " & Err.Description
End Function

Private Sub ImprimoEFactura(ByVal doc As Long)
On Error GoTo errIEF

    Dim eComCod As Long
    Dim rsE As rdoResultset
    Set RsAux = cBase.OpenResultset("SELECT EComID From eComprobantes WHERE EComID = " & doc, rdOpenForwardOnly, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "NO FUE FIRMADO EL REMITO, pida al administrador que se lo firme.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        Exit Sub
    End If
    RsAux.Close

    Dim oPrintEF As New ComPrintEfactura.ImprimoCFE
    'paPrintConfD , paPrintConfB, paPrintConfXDef, paPrintConfPaperSize
    oPrintEF.ImprimirCFEPorXML doc, paPrintConfD, paPrintConfB, paPrintConfPaperSize
    Set oPrintEF = Nothing
    Exit Sub
errIEF:
    clsGeneral.OcurrioError "Error al imprimir eFactura", Err.Description
End Sub

Private Sub AccionRePrint()

    If Val(tDoc.Tag) = 0 Then Exit Sub

    If MsgBox("¿Confirma reimprimir el documento?" & vbCrLf & vbCrLf & "Impresora: " & paPrintConfD & " en la bandeja " & paPrintConfB, vbQuestion + vbYesNo, "Atención") = vbYes Then

       On Error Resume Next
        Screen.MousePointer = 11
'        Dim gSucesoUsr As Integer, gSucesoDef As String
'        Dim objSuceso As New clsSuceso
'        objSuceso.ActivoFormulario paCodigoDeUsuario, "Reimpresión de Documentos", cBase
'        gSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
'        gSucesoDef = objSuceso.RetornoValor(Defensa:=True)
'        Set objSuceso = Nothing
'        Me.Refresh
'        If gSucesoUsr = 0 Then Screen.MousePointer = 0: Exit Sub 'Abortó el ingreso del suceso
'        '---------------------------------------------------------------------------------------------
'        clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.Reimpresiones, paCodigoDeTerminal, gSucesoUsr, CodTraslado, _
'                                   Descripcion:="Traslado: " & CodTraslado & " Documento: " & tDoc.Text, Defensa:=Trim(gSucesoDef)
        
        ImprimoEFactura Val(tDoc.Tag)
        Screen.MousePointer = 0
    End If
    
End Sub

Private Sub loc_AnuloDocumento()
Dim bDest As Boolean
Dim sUser As String, lUser As Long

    If cIntermediario.ListIndex > -1 Then
        If MsgBox("¿Confirma anular el traslado seleccionado?" & vbCr & vbCr & "La mercadería retornará al local origen.", vbQuestion + vbYesNo, "Eliminar") = vbNo Then Exit Sub
    Else
        If MsgBox("¿Confirma anular el traslado seleccionado?" & vbCr & vbCr & "La mercadería retornará al local origen si el destino la recepcionó.", vbQuestion + vbYesNo, "Eliminar") = vbNo Then Exit Sub
    End If
    
    lUser = 0
    sUser = ""
    sUser = InputBox("Ingrese su dígito de Usuario.", "Anular Traslado")
        
    If Not IsNumeric(sUser) Then
        MsgBox "No se ingresó un dígito correcto", vbExclamation, "ATENCIÓN"
        Exit Sub
    Else
        lUser = BuscoUsuarioDigito(CLng(sUser), True)
        If lUser = 0 Then Exit Sub
    End If
    
    FechaDelServidor
    Dim idRemito As Long
    
    Screen.MousePointer = 11
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrT
    
    Cons = "Select * From Traspaso Where TraCodigo = " & CodTraslado
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Tengo que validar quien tiene la mercadería.
    bDest = Not IsNull(RsAux("TraFechaEntregado"))
    
    Dim idDest As Long
    Dim idCamion As Long
    idCamion = 0
    idDest = RsAux("TraLocalDestino")
    If Not IsNull(RsAux!TraLocalIntermedio) Then idCamion = RsAux!TraLocalIntermedio
    If Not IsNull(RsAux("TraRemito")) Then idRemito = RsAux("TraRemito")
    
    RsAux.Edit
    RsAux("TraAnulado") = Format(gFechaServidor, "mm/dd/yyyy hh:nn")
    RsAux!TraFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:nn")
    RsAux.Update
    RsAux.Close


    'Si le di la mercadería lo anulo.
    If IsDate(lblRecepciono.Tag) Then
        Cons = "Select * From RenglonTraspaso Where RTrTraspaso = " & CodTraslado
        Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not Rs.EOF
            'Está en destino o está en camión
            If bDest Then
                MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, idDest, Rs("RTrArticulo"), Rs("RTrCantidad"), Rs("RTrEstado"), -1
                MarcoMovimientoStockFisico lUser, TipoLocal.Deposito, idDest, CLng(Rs!RTrArticulo), Rs!RTrCantidad, Rs!RTrEstado, -1, TipoDocumento.Traslados, CodTraslado
            Else
                If idCamion > 0 Then
                    'La tiene el camión
                    MarcoMovimientoStockFisicoEnLocal TipoLocal.Camion, idCamion, Rs("RTrArticulo"), Rs!RTrCantidad, Rs("RtrEstado"), -1
                    MarcoMovimientoStockFisico lUser, TipoLocal.Camion, idCamion, Rs!RTrArticulo, Rs!RTrCantidad, Rs!RTrEstado, -1, TipoDocumento.Traslados, CodTraslado
                End If
            End If
            'El origen no la entregó cuando es sin camión y no fue recepcionado.
            If Not (cIntermediario.ListIndex = -1 And Not bDest) Then
                'ORIGEN ... SUBO EL STOCK
                MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, cOrigen.ItemData(cOrigen.ListIndex), Rs("RTrArticulo"), Rs("RTrCantidad"), Rs("RTrEstado"), 1
                MarcoMovimientoStockFisico lUser, TipoLocal.Deposito, cOrigen.ItemData(cOrigen.ListIndex), CLng(Rs!RTrArticulo), Rs!RTrCantidad, Rs!RTrEstado, 1, TipoDocumento.Traslados, CodTraslado
            End If
            Rs.MoveNext
        Loop
        Rs.Close
    End If
    
    cBase.CommitTrans
    
    Screen.MousePointer = 0
    PresentoTraspaso CodTraslado
    Exit Sub
    
ErrBT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al intentar iniciar la transaccion.", Err.Description
    Exit Sub
   
Resumo:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al intentar almacenar la información.", Err.Description
    Exit Sub
ErrT:
    Resume Resumo
End Sub

Private Sub loc_DBAddRenglones(ByVal iCodT As Long)
    With vsConsulta
        For I = 1 To .Rows - 1
            If Val(.Cell(flexcpText, I, 0)) > 0 Then
                'Inserto el renglón del traspaso.-----------------------------
                Cons = "Insert into RenglonTraspaso (RTrTraspaso, RTrArticulo, RTrEstado, RTrCantidad, RTrPendiente)" _
                    & " Values (" & iCodT & ", " & Val(.Cell(flexcpData, I, 0)) _
                    & ", " & Val(.Cell(flexcpData, I, 1)) & ", " & Val(.Cell(flexcpText, I, 0)) _
                    & ", " & Val(.Cell(flexcpText, I, 0)) & ")"
                cBase.Execute (Cons)
            End If
        Next
    End With
End Sub

Private Sub loc_FindTrasladoByDoc()
On Error GoTo errFTD
Dim sSerie As String, sNro As String
    
    If InStr(tDoc.Text, "-") <> 0 Then
        sSerie = Mid(tDoc.Text, 1, InStr(tDoc.Text, "-") - 1)
        sNro = Val(Mid(tDoc.Text, InStr(tDoc.Text, "-") + 1))
    Else
        sSerie = Mid(tDoc.Text, 1, 1)
        sNro = Val(Mid(tDoc.Text, 2))
    End If
    tDoc.Text = UCase(sSerie) & "-" & sNro
    
    Cons = "Select TraCodigo, TraSerie as Serie, TraNumero as Numero, SucAbreviacion as Sucursal " & _
                " From Traspaso, Sucursal Where TraSerie = '" & sSerie & "' And TraNumero = " & CLng(sNro) & _
                " And TraSucursal = SucCodigo"
    Screen.MousePointer = 11
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        sNro = RsAux("TraCodigo")
        RsAux.MoveNext
        If Not RsAux.EOF Then
            'Presento lista de ayuda hay más de uno.
            sNro = ""
            Dim objLista As New clsListadeAyuda
            If objLista.ActivarAyuda(cBase, Cons, 6000, 1, "Traslados de mercadería") > 0 Then
                sNro = objLista.RetornoDatoSeleccionado(0)
            End If
            Me.Refresh
            Set objLista = Nothing
        End If
    Else
        sNro = ""
        Screen.MousePointer = 0
        MsgBox "No existe un traslado que corresponda a los datos ingresados.", vbExclamation, "Atención"
    End If
    RsAux.Close
    
    If sNro <> "" Then PresentoTraspaso CLng(sNro)
Exit Sub
errFTD:
    Screen.MousePointer = 0
    MsgBox "Error al buscar el traslado por documento.", vbExclamation, "Atención"
End Sub

Private Sub cArticulo_GotFocus()
    Status.SimpleText = " Ingrese el código o nombre de un artículo."
    cArticulo.SelStart = 0: cArticulo.SelLength = Len(cArticulo.Text)
End Sub

Private Sub cArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(cArticulo.Text) <> vbNullString Then
        If prmQRenglon > vsConsulta.Rows - 1 Then
            If Not IsNumeric(cArticulo.Text) Then
                BuscoArticuloXNombre
            Else
                BuscoArticuloCodigo cArticulo.Text
            End If
            If cArticulo.ListIndex > -1 Then Foco tCantidad
        Else
            MsgBox "No puede agregar más renglones.", vbExclamation, "Atención"
        End If
    Else
        If KeyAscii = vbKeyReturn And vsConsulta.Rows > 1 Then vsConsulta.SetFocus
    End If
End Sub
Private Sub cArticulo_LostFocus()
    Status.SimpleText = vbNullString
End Sub

Private Sub BuscoArticuloCodigo(aCodigo As Long)
On Error GoTo ErrBAC
    Screen.MousePointer = 11
    Cons = "Select * From Articulo Where ArtCodigo = " & aCodigo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    cArticulo.Clear
    If Not RsAux.EOF Then
        cArticulo.AddItem Trim(RsAux!ArtNombre)
        cArticulo.ItemData(cArticulo.NewIndex) = RsAux!ArtID
        cArticulo.ListIndex = 0
    End If
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
ErrBAC:
    clsGeneral.OcurrioError "Error al buscar el artículo por código.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cDestino_Click()
    'Si modifica no cambio los valores guardados en los tag.
    If Not sModificar Then
        CodTraslado = 0: cIntermediario.Tag = vbNullString: txtCodigo.Text = ""
    End If
    If Not sNuevo And Not sModificar Then
        OcultoTodo
    End If
End Sub

Private Sub cDestino_Change()
    If Not sNuevo And Not sModificar Then OcultoTodo
End Sub

Private Sub cEstado_GotFocus()
    cEstado.SelStart = 0
    cEstado.SelLength = Len(cEstado.Text)
    Status.SimpleText = " Seleccione el estado físico del artículo."
End Sub

Private Sub cEstado_KeyPress(KeyAscii As Integer)
Dim iStock As Integer
    If KeyAscii = vbKeyReturn And cArticulo.ListIndex > -1 And IsNumeric(tCantidad.Text) And cEstado.ListIndex > -1 Then
        If CLng(tCantidad.Text) > 0 Then
            iStock = StockLocalArticuloyEstado(cArticulo.ItemData(0), cEstado.ItemData(cEstado.ListIndex))
            If iStock < CLng(tCantidad.Text) Then
                If MsgBox("No hay tantos artículos en stock. ¿Desea continuar de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
            End If
            InsertoArticuloEnlaLista (iStock)
            LimpioIngreso
            cArticulo.SetFocus
        Else
            MsgBox "Los datos son incorrectos.", vbExclamation, "ATENCIÓN": Foco tCantidad
        End If
    ElseIf KeyAscii = vbKeyReturn Then
        MsgBox "Los datos son incompletos o incorrectos.", vbExclamation, "ATENCIÓN"
    End If
End Sub

Private Sub cEstado_LostFocus()
    Status.SimpleText = vbNullString
    cEstado.SelStart = 0
End Sub

Private Sub cIntermediario_Click()
    
    If Not sModificar Then CodTraslado = 0: cIntermediario.Tag = vbNullString:: txtCodigo.Text = ""
    If Not sNuevo And Not sModificar Then
        Botones True, False, False, False, False, Toolbar1, Me
        Toolbar1.Buttons("imprimir").Enabled = False: MnuImprimir.Enabled = False
        Toolbar1.Buttons("entregar").Enabled = False: MnuEntregar.Enabled = False
        Toolbar1.Buttons("rePrint").Enabled = False
        vsConsulta.Rows = 1
        tComentario.Text = vbNullString: labCreado.Caption = vbNullString: labUsuario.Caption = vbNullString
    End If
    
End Sub

Private Sub cIntermediario_Change()
    If Not sNuevo And Not sModificar Then OcultoTodo
End Sub

Private Sub cIntermediario_GotFocus()
    cIntermediario.SelStart = 0
    cIntermediario.SelLength = Len(cIntermediario.Text)
    Status.SimpleText = " Seleccione un camión intermediario. - [ F1] Pendientes, [ F2] Realizados -"
End Sub

Private Sub cIntermediario_KeyDown(KeyCode As Integer, Shift As Integer)
    If cIntermediario.ListIndex = -1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyF1: BuscoTraspasoXIntermediario
        Case vbKeyF2: BuscoTraspasoRealizadoXIntermediario
    End Select
End Sub
Private Sub cIntermediario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cDestino.SetFocus
End Sub
Private Sub cIntermediario_LostFocus()
    cIntermediario.SelStart = 0: Status.SimpleText = ""
End Sub

Private Sub cOrigen_Click()

    If Not sModificar Then
        CodTraslado = 0: cIntermediario.Tag = vbNullString: txtCodigo.Text = ""
    End If
    If Not sNuevo And Not sModificar Then
        Botones True, False, False, False, False, Toolbar1, Me
        Toolbar1.Buttons("imprimir").Enabled = False: MnuImprimir.Enabled = False
        Toolbar1.Buttons("entregar").Enabled = False: MnuEntregar.Enabled = False
        Toolbar1.Buttons("rePrint").Enabled = False
        vsConsulta.Rows = 1
        tComentario.Text = vbNullString: labCreado.Caption = vbNullString: labUsuario.Caption = vbNullString
    End If

End Sub

Private Sub cOrigen_Change()
    If Not sNuevo And Not sModificar Then OcultoTodo
End Sub

Private Sub cOrigen_GotFocus()
    cOrigen.SelStart = 0: cOrigen.SelLength = Len(cOrigen.Text)
    Status.SimpleText = " Seleccione un Local. - [ F1] Pendientes, [ F2] Realizados -"
End Sub

Private Sub cOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
    If cOrigen.ListIndex = -1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyF1: BuscoTraspasoXOrigen
        Case vbKeyF2: BuscoTraspasoRealizadoXOrigen
    End Select
End Sub

Private Sub cOrigen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cIntermediario.SetFocus
End Sub
Private Sub cOrigen_LostFocus()
    cOrigen.SelLength = 0: Status.SimpleText = vbNullString
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Screen.MousePointer = vbDefault: Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad

    HagoPausa 1
    Status.Panels("printer").Text = paPrintConfD
    lblRecepciono.Caption = vbNullString
    tmServerPrint.Enabled = False
    With vsConsulta
        .Redraw = False
        .Editable = False: .Rows = 1: .Cols = 1: .ExtendLastCol = True
        .FormatString = "Entregar|<Estado|<Artículo|Stock"
        .ColWidth(1) = 1000: .ColWidth(2) = 3500
        .ColHidden(3) = True
        .Redraw = True
    End With
    
    sNuevo = False: sModificar = False
    
    CargoLocales
    CargoEstados
    DeshabilitoIngreso
    Botones True, False, False, False, False, Toolbar1, Me
    Toolbar1.Buttons("imprimir").Enabled = False: MnuImprimir.Enabled = False
    Toolbar1.Buttons("entregar").Enabled = False: MnuEntregar.Enabled = False
    Toolbar1.Buttons("rePrint").Enabled = False
    
    Cons = "Select * From Sucursal Where SucCodigo = " & paCodigoDeSucursal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        Me.Caption = "Traslado de Mercadería (Sucursal: " & Trim(RsAux!SucAbreviacion) & ") "
    End If
    RsAux.Close
    'tmServerPrint.Enabled = True
    Exit Sub
    
ErrLoad:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error inesperado. " & Err.Description
    
End Sub
Private Sub cDestino_GotFocus()
    cDestino.SelStart = 0
    cDestino.SelLength = Len(cDestino.Text)
    Status.SimpleText = " Seleccione un Local. - [ F1] Pendientes, [ F2] Realizados -"
End Sub

Private Sub cDestino_KeyDown(KeyCode As Integer, Shift As Integer)
    If cDestino.ListIndex = -1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyF1: BuscoTraspasoXDestino
        Case vbKeyF2: BuscoTraspasoRealizadosXDestino
    End Select
End Sub

Private Sub cDestino_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If cOrigen.ListIndex = -1 Then
            MsgBox "No selecciono el local Origen.", vbExclamation, "ATENCIÓN"
            cOrigen.SetFocus: Exit Sub
        End If
        If cDestino.ListIndex > -1 Then
            If cDestino.ItemData(cDestino.ListIndex) = cOrigen.ItemData(cOrigen.ListIndex) Then
                MsgBox "Selecciono el mismo local.", vbExclamation, "ATENCIÓN"
            Else
                If cArticulo.Enabled Then
                    Foco cArticulo
                Else
                    If vsConsulta.Enabled Then vsConsulta.SetFocus
                End If
            End If
        ElseIf Not sModificar And Not sNuevo Then
            Toolbar1.Buttons("imprimir").Enabled = False: MnuImprimir.Enabled = False
            DeshabilitoIngreso
            LimpioIngreso
        End If
    End If

End Sub

Private Sub cDestino_LostFocus()
    cDestino.SelLength = 0
    Status.SimpleText = vbNullString
End Sub

Private Sub CargoLocales()
On Error GoTo ErrCO

    cOrigen.Clear: cDestino.Clear: cIntermediario.Clear

    Cons = "Select CamCodigo, CamNombre From Camion Order by CamNombre"
    CargoCombo Cons, cIntermediario, ""

    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cOrigen, ""
    CargoCombo Cons, cDestino, ""
    Exit Sub

ErrCO:
    clsGeneral.OcurrioError "Error al cargar los Locales."
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.SimpleText = vbNullString
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Sub

Private Sub Label1_Click()
    Foco cOrigen
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.SimpleText = " Seleccione un Local. - [ F1] Búsqueda - "
End Sub

Private Sub Label2_Click()
    Foco tComentario
End Sub
Private Sub Label3_Click()
    Foco cIntermediario
End Sub

Private Sub Label4_Click()
    Foco cArticulo
End Sub

Private Sub Label5_Click()
    Foco tCantidad
End Sub

Private Sub Label6_Click()
    Foco cEstado
End Sub

Private Sub Label7_Click()
    Foco cDestino
End Sub

Private Sub Label8_Click()
    Foco tUsuario
End Sub

Private Sub MnuCancelar_Click()
    AccionCancelar
End Sub
Private Sub MnuEliminar_Click()
    AccionEliminar
End Sub

Private Sub MnuEntregar_Click()
    AccionEntregar
End Sub

Private Sub MnuGrabar_Click()
    AccionGrabar
End Sub

Private Function StockLocalArticuloyEstado(lnArticulo As Long, iEstado As Integer) As Integer
On Error GoTo errSTL

    Screen.MousePointer = vbHourglass
    StockLocalArticuloyEstado = 0

    Cons = "Select Sum(StLCantidad) From StockLocal " _
        & " Where StLArticulo = " & lnArticulo & " And StlTipoLocal = " & TipoLocal.Deposito _
        & " And StLLocal = " & cOrigen.ItemData(cOrigen.ListIndex) & " And StLEstado = " & iEstado
        
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then StockLocalArticuloyEstado = Rs(0)
    Rs.Close
    Screen.MousePointer = vbDefault
    Exit Function
        
errSTL:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error inesperado al buscar el stock del local."

End Function

Private Sub MnuModificar_Click()
    AccionModificar
End Sub

Private Sub MnuNuevo_Click()
    AccionNuevo
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub Status_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If "printer" = Panel.Key Then
        prj_GetPrinter True
        Panel.Text = paPrintConfD
    End If
End Sub

Private Sub tCantidad_GotFocus()
    tCantidad.SelStart = 0
    tCantidad.SelLength = Len(tCantidad.Text)
    Status.SimpleText = " Ingrese la cantidad a traspasar."
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCantidad.Text) Then
            If CLng(tCantidad.Text) > 0 Then
                cEstado.SetFocus
                If cEstado.ListIndex = -1 Then BuscoCodigoEnCombo cEstado, CLng(paEstadoArticuloEntrega)
            Else
                MsgBox "Debe ingresar un número mayor que cero.", vbExclamation, "ATENCIÓN"
            End If
        Else
            MsgBox "El formato no es numérico.", vbExclamation, "ATENCIÓN"
        End If
    End If

End Sub

Private Sub tCantidad_LostFocus()
    Status.SimpleText = vbNullString
End Sub

Private Sub tComentario_GotFocus()
    tComentario.SelStart = 0
    tComentario.SelLength = Len(tComentario.Text)
    Status.SimpleText = " Ingrese un comentario."
End Sub
Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tUsuario.SetFocus
End Sub

'Private Sub tDoc_Change()
'    If CodTraslado > 0 Then OcultoTodo False: CodTraslado = 0
'End Sub

'Private Sub tDoc_KeyPress(KeyAscii As Integer)
'On Error Resume Next
'    If KeyAscii = vbKeyReturn Then
'        If CodTraslado = 0 Then loc_FindTrasladoByDoc
'    End If
'End Sub

Private Sub tmServerPrint_Timer()
On Error GoTo errSP
    tmServerPrint.Enabled = False
    MsgBox "TODO", vbExclamation, "PENDIENTE"
'    Dim rsP As rdoResultset
'    Set rsP = cBase.OpenResultset("SELECT IsNull(CodValor1, 0) FROM Codigos WHERE CodCual = 151 AND CodValor2 = " _
'                                & paCodigoDeSucursal & " ORDER BY CodValor1", rdOpenDynamic, rdConcurValues)
'    Do While Not rsP.EOF
'        If rsP(0) > 0 Then
'            fnc_PrintDocumento rsP(0), miConexion.UsuarioLogueado(True)
'        End If
'        rsP.MoveNext
'    Loop
'    rsP.Close
'    tmServerPrint.Interval = 10000
'    tmServerPrint.Enabled = True
    Exit Sub
errSP:
    clsGeneral.OcurrioError "Error al consultar los posibles documentos a imprimir.", Err.Description, "Servidor de impresión"
    Screen.MousePointer = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "salir": Unload Me
        Case "entregar": AccionEntregar
        Case "cancelar": AccionCancelar
        Case "rePrint": AccionRePrint
        Case "imprimir": AccionImprimirAux
    End Select
End Sub

Private Sub CargoValoresIVA()
Dim RsIva As rdoResultset
Dim sQy As String
    sQy = "SELECT IvaCodigo, IvaPorcentaje FROM TipoIva WHERE IvaCodigo IN (1,2)"
    Set RsIva = cBase.OpenResultset(sQy, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsIva.EOF
        Select Case RsIva("IvaCodigo")
            Case 1: TasaBasica = RsIva("IvaPorcentaje")
            Case 2: TasaMinima = RsIva("IvaPorcentaje")
        End Select
        RsIva.MoveNext
    Loop
    RsIva.Close
End Sub

Private Sub AccionEntregar()
Dim Msg As String ', sSinTerminal As Boolean
Dim idOrigen As Long, idCamion As Long, idDestino As Long
   
    idOrigen = 0: idCamion = 0: idDestino = 0
    On Error GoTo errAE
    
    If TasaBasica = 0 Then CargoValoresIVA
    
    'If (EmpresaEmisora.Codigo = 0) Then
    If (EmpresaEmisora Is Nothing) Then
        Set EmpresaEmisora = New clsClienteCFE
        EmpresaEmisora.CargoClienteCarlosGutierrez paCodigoDeSucursal
    End If
    
    Cons = "Select * From Traspaso Where TraCodigo = " & CodTraslado
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not IsNull(RsAux("TraFModificacion")) Then
        If RsAux!TraFModificacion <> CDate(cIntermediario.Tag) Then
            RsAux.Close
            MsgBox "El traslado fue modificado por otra terminal, se cargará nuevamente el mismo para desplegar los cambios.", vbInformation, "ATENCIÓN"
            PresentoTraspaso CodTraslado
            Exit Sub
        End If
    End If
    idOrigen = RsAux!TraLocalOrigen
    If Not IsNull(RsAux!TraLocalIntermedio) Then idCamion = RsAux!TraLocalIntermedio
    idDestino = RsAux!TraLocalDestino
    RsAux.Close

    '10/11/2008 Si el local de destino no es mi local y el origen es sin terminal --> NO recepciono.
    
    Dim bOrigenST As Boolean, bDestST As Boolean
    'sSinTerminal = False
    Cons = "Select LocCodigo, LocSinTerminal  from Local Where LocCodigo in (" & idDestino & ", " & idOrigen & ")"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Select Case RsAux("LocCodigo")
            Case idDestino
                bDestST = RsAux!LocSinTerminal
            Case idOrigen
                bOrigenST = RsAux!LocSinTerminal
        End Select
        'If RsAux!LocSinTerminal Then sSinTerminal = True
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If paCodigoDeSucursal <> idOrigen And Not bOrigenST Then
        MsgBox "Su terminal no corresponde al local que debe entregar la mercadería, no podrá almacenar.", vbCritical, "ATENCIÓN"
        Exit Sub
    End If
    
    Dim oInfoCAE As New clsCAEGenerador
    If Not oInfoCAE.SucursalTieneCae(cBase, CGSA_TiposCFE.CFE_eRemito, paCodigoDGI) Then
        MsgBox "No hay un CAE disponible para emitir el eRemito, por favor comuniquese con administración.", vbCritical, "eFactura"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If bDestST Then
        MsgBox "El Local Destino es sin terminal." & Chr(13) & "Se recepcionará automáticamente el traslado.", vbInformation, "ATENCIÓN"
    End If
    
    Dim sPreg As String
    If MsgBox("¿Confirma emitir el documento del traslado?" & IIf(bDestST Or idCamion > 0, vbCr & vbCr & "Se haran movimientos de stock.", ""), vbQuestion + vbYesNo, "Entregar mercadería") = vbNo Then Exit Sub
    
    Msg = vbNullString
    Msg = InputBox("Ingrese su código de usuario.", "Entregar mercadería de Traslado.")
    
    If Not IsNumeric(Msg) Then
        Exit Sub
    Else
        tUsuario.Tag = BuscoUsuarioDigito(CInt(Msg), True)
        If CInt(tUsuario.Tag) <= 0 Then Exit Sub
    End If
    
    
    If bDestST Or idCamion > 0 Then
        With vsConsulta
            For I = 1 To .Rows - 1
                If Val(.Cell(flexcpValue, I, 0)) > Val(.Cell(flexcpValue, I, 3)) Then
                    If MsgBox("Hay artículos que dejarán el stock negativo. ¿Confirma continuar?", vbQuestion + vbYesNo, "Stock negativo") = vbNo Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next
        End With
    End If
    
    
    Screen.MousePointer = vbHourglass
    Msg = "Error inesperado. " & Err.Description
    FechaDelServidor
    
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrRelajo
    
        
    Cons = "Select * From Traspaso Where TraCodigo = " & CodTraslado
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux!TraFModificacion <> CDate(cIntermediario.Tag) Then
        Msg = "El traslado fue modificado, deberá editar la información nuevamente."
        RsAux.Close
        RsAux.Edit 'Ocasiono el error
    End If
    
        'Genero el CAE
    Dim CAE As New clsCAEDocumento
    Dim caeG As New clsCAEGenerador
    Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, CGSA_TiposCFE.CFE_eRemito, paCodigoDGI)
    Set caeG = Nothing
    
    Dim doc As New clsDocumentoCGSA
    With doc
        Set .Cliente = EmpresaEmisora
        .Emision = gFechaServidor
        .Tipo = TD_TrasladosInternos
        .Numero = CAE.Numero
        .Serie = CAE.Serie
        .Moneda.Codigo = 1
        .Total = 0
        .IVA = 0
        .Sucursal = paCodigoDeSucursal
        .Digitador = CInt(tUsuario.Tag)
        .Comentario = "Traslado " & CodTraslado
        .Vendedor = CInt(tUsuario.Tag)
'        .Adenda = "Traslado de mercadería: " & CodTraslado & "<BR/>" & "Origen: " & cOrigen.Text & IIf(cIntermediario.Text <> "", ", Camión: " & cIntermediario.Text, "") & ", Destino: " & cDestino.Text & "<BR/>" & "Memo: " & tComentario.Text
    End With
    Set doc.Conexion = cBase
    doc.Codigo = doc.InsertoCabezalDelDocumento
    
    RsAux.Edit
    RsAux!TraFImpreso = Format(gFechaServidor, sqlFormatoFH)
    RsAux!TraFModificacion = Format(gFechaServidor, sqlFormatoFH)
    
    RsAux("TraSucursal") = paCodigoDeSucursal
    RsAux("TraRemito") = doc.Codigo
    
'    If paDContado <> "" Then
'        RsAux("TraSerie") = sTxt
'        RsAux("TraNumero") = iNro
'    End If

    'Si es sin terminal ya recepciono la mercadería y digo quien la recibe.
    If bDestST Then
        RsAux!TraFechaEntregado = Format(gFechaServidor, sqlFormatoFH)
        RsAux!TraUsuarioReceptor = CInt(tUsuario.Tag)
        RsAux!TraUsuarioFinal = CInt(tUsuario.Tag)
    End If
    
    RsAux!TraTerminal = paCodigoDeTerminal
    RsAux.Update
    RsAux.Close
    
    'Marco como entregado.
    cDestino.Tag = "1"
    
    
    Dim oRenTras As clsDocumentoRenglon
    With vsConsulta
        For I = 1 To .Rows - 1
            
            If CInt(.Cell(flexcpValue, I, 0)) > 0 Then
            
                Set oRenTras = New clsDocumentoRenglon
                doc.Renglones.Add oRenTras
                
                oRenTras.Articulo.ID = Val(.Cell(flexcpData, I, 0))
                oRenTras.EstadoMercaderia = Val(.Cell(flexcpData, I, 1))
                oRenTras.Cantidad = Val(.Cell(flexcpText, I, 0))
                
                
                If bDestST Or idCamion > 0 Then
                
                    'Primero intento dar la baja al origen de forma de ver si queda ese stock.
                    Cons = "Select * From StockLocal " _
                        & " Where StLArticulo = " & Val(.Cell(flexcpData, I, 0)) _
                        & " And StlTipoLocal = " & TipoLocal.Deposito & " And StLLocal = " & idOrigen _
                        & " And StLEstado = " & Val(.Cell(flexcpData, I, 1))
                    
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
                    If RsAux.EOF Then
                        RsAux.AddNew
                        RsAux!StLArticulo = Val(.Cell(flexcpData, I, 0))
                        RsAux!StLTipoLocal = TipoLocal.Deposito
                        RsAux!StlLocal = idOrigen
                        RsAux!StLEstado = Val(.Cell(flexcpData, I, 1))
                        RsAux!StLCantidad = Val(.Cell(flexcpText, I, 0)) * -1
                        RsAux.Update
                        clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.DiferenciaDeArticulos, paCodigoDeTerminal, CLng(tUsuario.Tag), 0, CLng(.Cell(flexcpData, I, 0)), _
                              Descripcion:="Traslado de Mercadería, código: " & CodTraslado, Defensa:="Se ingresaron " & CInt(.Cell(flexcpText, I, 0)) & " artículos " & Trim(.Cell(flexcpText, I, 2)) & "  , estado " & Trim(.Cell(flexcpText, I, 1)) & "sin haber en el local."
                    Else
                        If RsAux!StLCantidad < CInt(.Cell(flexcpValue, I, 0)) Then
                            clsGeneral.RegistroSuceso cBase, gFechaServidor, TipoSuceso.DiferenciaDeArticulos, paCodigoDeTerminal, CLng(tUsuario.Tag), 0, CLng(.Cell(flexcpData, I, 0)), _
                                Descripcion:="Traslado de Mercadería, código: " & CodTraslado, Defensa:="Se entregaron " & CInt(.Cell(flexcpValue, I, 0)) & " artículos  " & Trim(.Cell(flexcpText, I, 2)) & "  , estado " & Trim(.Cell(flexcpText, I, 1)) & " y en el local habían " & RsAux!StLCantidad
                        End If
                        If RsAux!StLCantidad - CInt(.Cell(flexcpValue, I, 0)) = 0 Then
                            RsAux.Delete
                        Else
                            RsAux.Edit
                            RsAux!StLCantidad = RsAux!StLCantidad - CInt(.Cell(flexcpValue, I, 0))
                            RsAux.Update
                        End If
                    End If
                    RsAux.Close
                    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, idOrigen, Val(.Cell(flexcpData, I, 0)), CInt(.Cell(flexcpValue, I, 0)), Val(.Cell(flexcpData, I, 1)), -1, TipoDocumento.Traslados, CodTraslado
                
                End If
                
                'Inserto en el destino.
                'Aca el destino es el camión.
                If idCamion > 0 Then
                    If Not bDestST Then
                        'Como no tiene terminal el destino abajo hago la recepción. entonces no toco el stock del camión.
                        'simplemente simulo con los movimientos fisicos.
                        MarcoMovimientoStockFisicoEnLocal TipoLocal.Camion, idCamion, Val(.Cell(flexcpData, I, 0)), Val(.Cell(flexcpValue, I, 0)), Val(.Cell(flexcpData, I, 1)), 1
                    End If
                    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Camion, idCamion, Val(.Cell(flexcpData, I, 0)), Val(.Cell(flexcpValue, I, 0)), Val(.Cell(flexcpData, I, 1)), 1, TipoDocumento.Traslados, CodTraslado
                End If
                
                'Si es sin terminal hago recepción.
                If bDestST Then
                    If idCamion > 0 Then
                        'Le saco al camión
                        MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Camion, idCamion, Val(.Cell(flexcpData, I, 0)), Val(.Cell(flexcpValue, I, 0)), Val(.Cell(flexcpData, I, 1)), -1, TipoDocumento.Traslados, CodTraslado
                    End If
                    
                    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, idDestino, Val(.Cell(flexcpData, I, 0)), Val(.Cell(flexcpValue, I, 0)), Val(.Cell(flexcpData, I, 1)), 1
                    MarcoMovimientoStockFisico CLng(tUsuario.Tag), TipoLocal.Deposito, idDestino, Val(.Cell(flexcpData, I, 0)), Val(.Cell(flexcpValue, I, 0)), Val(.Cell(flexcpData, I, 1)), 1, TipoDocumento.Traslados, CLng(CodTraslado)
                    
                    Cons = "Update RenglonTraspaso set RTrPendiente = 0 Where RTrTraspaso = " & CodTraslado _
                        & " And RTrArticulo = " & Val(.Cell(flexcpData, I, 0)) & " And RTrEstado = " & Val(.Cell(flexcpData, I, 1))
                    cBase.Execute (Cons)
                End If
                
            End If
        Next
        doc.InsertoRenglonDocumentoEstadoBD
    End With
    cBase.CommitTrans
    On Error Resume Next
    
    Dim sPaso As String
    sPaso = EmitirCFE(doc, CAE)
    Dim resM As VbMsgBoxResult
    resM = vbYes
    Do While sPaso <> ""
        resM = MsgBox("ATENCIÓN no se firmó el documento" & vbCrLf & vbCrLf & "Presione SI para reintentar" & vbCrLf & " Presione NO para abandonar ", vbExclamation + vbYesNo, "ATENCIÓN")
        If resM = vbNo Then Exit Do
        sPaso = EmitirCFE(doc, CAE)
    Loop
    If sPaso = "" Then
        'Hago pausa ya que está dando un error.
        HagoPausa 1
        ImprimoEFactura doc.Codigo
    End If
    Set doc = Nothing
    
    PresentoTraspaso CodTraslado
    Screen.MousePointer = vbDefault
    Exit Sub

errAE:
    clsGeneral.OcurrioError "Error al intentar entregar la mercadería", Err.Description
    Exit Sub
    
ErrBT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al intentar iniciar la transaccion.", Err.Description
    Exit Sub
   
Resumo:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError Msg, Err.Description
    Exit Sub
    
ErrRelajo:
    Resume Resumo
End Sub

Private Sub HagoPausa(ByVal segundos As Integer)
    Dim hasta As Date
    hasta = DateAdd("s", segundos, Now)
    Do While hasta > Now
        'Debug.Print hasta
        DoEvents
    Loop
End Sub


Private Sub CargoEstados()
On Error GoTo ErrCE

    'Levanto los estados que no afecten el stock total.
    Cons = "Select EsMCodigo, EsMAbreviacion From EstadoMercaderia " _
        & " Where EsMBajaStockTotal = 0 Order by EsMAbreviacion"
    CargoCombo Cons, cEstado, ""
    Exit Sub

ErrCE:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al cargar los Estados."

End Sub

Private Sub DeshabilitoIngreso()
    txtCodigo.Enabled = True: txtCodigo.BackColor = vbWindowBackground
    cArticulo.Enabled = False: cArticulo.BackColor = Inactivo
    tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
    cEstado.Enabled = False: cEstado.BackColor = Inactivo
    tComentario.Enabled = False: tComentario.BackColor = Inactivo
    tUsuario.BackColor = Inactivo: tUsuario.Enabled = False
'    vsConsulta.Enabled = False
'    tDoc.Enabled = True
End Sub

Private Sub HabilitoIngreso()
    txtCodigo.Enabled = False: txtCodigo.BackColor = Inactivo
    cArticulo.Enabled = True: cArticulo.BackColor = Blanco
    tCantidad.Enabled = True: tCantidad.BackColor = Blanco
    cEstado.Enabled = True: cEstado.BackColor = Blanco
'    vsConsulta.Enabled = True
    tComentario.Enabled = True: tComentario.BackColor = Blanco
    tUsuario.BackColor = Obligatorio: tUsuario.Enabled = True
    tDoc.Enabled = False
End Sub

Private Sub LimpioIngreso()
    tDoc.Text = "": tDoc.Tag = ""
    cArticulo.Clear
    tCantidad.Text = vbNullString
    cEstado.Text = ""
End Sub
Private Sub BuscoArticuloXNombre()
On Error GoTo ErrBAN
Dim aCodigo As Long: aCodigo = 0

    Screen.MousePointer = vbHourglass
    Cons = "Select ArtCodigo, Código = ArtCodigo, Artículo = ArtNombre from Articulo" _
        & " Where ArtNombre LIKE '" & Replace(cArticulo.Text, " ", "%") & "%'" _
        & " Order by ArtNombre"
    cArticulo.Clear
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un nombre de artículo con esas características.", vbInformation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            aCodigo = RsAux(0)
            RsAux.Close
        Else
            RsAux.Close
            Dim objLista As New clsListadeAyuda
            If objLista.ActivarAyuda(cBase, Cons, 5000, 1, "Lista Artículos") Then
                aCodigo = objLista.RetornoDatoSeleccionado(0)
            End If
            Set objLista = Nothing       'Destruyo la clase.
        End If
        If aCodigo > 0 Then BuscoArticuloCodigo aCodigo
    End If
    Screen.MousePointer = 0
    Exit Sub
    
ErrBAN:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub InsertoArticuloEnlaLista(Stock As Integer)
    
    On Error GoTo ErrClave
    
    With vsConsulta
        'Veo si ya inserte el artículo para el estado seleccionado.
        For I = 1 To .Rows - 1
            If CLng(.Cell(flexcpData, I, 0)) = cArticulo.ItemData(cArticulo.ListIndex) _
                And CInt(.Cell(flexcpData, I, 1)) = cEstado.ItemData(cEstado.ListIndex) Then
                MsgBox "Ya se inserto ese artículo con el estado seleccionado, verifique.", vbExclamation, "ATENCIÓN": Exit Sub
            End If
        Next I
    
        .AddItem tCantidad.Text
        .Cell(flexcpText, .Rows - 1, 1) = cEstado.Text
        .Cell(flexcpText, .Rows - 1, 2) = Trim(cArticulo.Text)
        .Cell(flexcpText, .Rows - 1, 3) = Stock
        'Data
        .Cell(flexcpData, .Rows - 1, 0) = cArticulo.ItemData(cArticulo.ListIndex)
        .Cell(flexcpData, .Rows - 1, 1) = cEstado.ItemData(cEstado.ListIndex)
    End With
    Exit Sub

ErrIAEL:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error inesperado al ingresar el artículo en la lista."
    Exit Sub
    
ErrClave:
    Screen.MousePointer = vbDefault
    MsgBox "El artículo ya fue ingresado con ese estado, verífique.", vbCritical, "ATENCIÓN"
    
End Sub
Private Sub tUsuario_GotFocus()
    tUsuario.SelStart = 0
    tUsuario.SelLength = Len(tUsuario.Text)
    Status.SimpleText = " Ingrese su código de Usuario."
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    tUsuario.Tag = vbNullString
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tUsuario.Text) Then
            tUsuario.Tag = BuscoUsuarioDigito(CInt(tUsuario.Text), True)
            If Val(tUsuario.Tag) > 0 Then AccionGrabar Else tUsuario.Tag = vbNullString
        Else
            MsgBox "El formato del código no es numérico.", vbExclamation, "ATENCIÓN"
            tUsuario.SetFocus
        End If
    End If
End Sub
Private Sub AccionCancelar()

    LimpioIngreso
    DeshabilitoIngreso
    vsConsulta.Rows = 1
    If Not sModificar Then
        CodTraslado = 0
        cOrigen.ListIndex = -1
        cIntermediario.ListIndex = -1
        cDestino.ListIndex = -1
        Botones True, False, False, False, False, Toolbar1, Me
        MnuImprimir.Enabled = False: Toolbar1.Buttons("imprimir").Enabled = False
        Toolbar1.Buttons("rePrint").Enabled = False
    Else
        PresentoTraspaso CLng(CodTraslado)
        cOrigen.Enabled = True
        cDestino.Enabled = True
        cIntermediario.Enabled = True
    End If
    sNuevo = False: sModificar = False
    Foco cOrigen
    
End Sub
Private Sub BuscoTraspasoRealizadoXOrigen()
On Error GoTo ErrBTXO
    
    If sNuevo Or sModificar Then Exit Sub
    
    Botones True, False, False, False, False, Toolbar1, Me
    MnuImprimir.Enabled = False: Toolbar1.Buttons("entregar").Enabled = False
    Toolbar1.Buttons("entregar").Enabled = False: MnuEntregar.Enabled = False
    Toolbar1.Buttons("rePrint").Enabled = False
    CodTraslado = 0
    LimpioIngreso
    vsConsulta.Rows = 1
    cIntermediario.ListIndex = -1
    cDestino.ListIndex = -1
    tComentario.Text = vbNullString
    
    If cOrigen.ListIndex <> -1 Then
        
        Cons = "Select Distinct(TraCodigo), Código = TraCodigo, Fecha = TraFecha,  Intermediario = CamNombre, Destino = LocNombre From Traspaso " _
                & " Left Outer Join Camion ON TraLocalIntermedio = CamCodigo " _
                & ", Local" _
                & " Where TraLocalOrigen = " & cOrigen.ItemData(cOrigen.ListIndex) _
                & " And TraFechaEntregado >= '" & Format(Date - 3, "mm/dd/yyyy") & "'" _
                & " And TraLocalDestino = LocCodigo And TraAnulado IS Null "
                
        Screen.MousePointer = 11
        Dim objLista As New clsListadeAyuda
        Dim aCodigo As Long: aCodigo = 0
        If objLista.ActivarAyuda(cBase, Cons, 6000, 1, "Lista Ayuda") > 0 Then
            aCodigo = objLista.RetornoDatoSeleccionado(0)
        End If
        Me.Refresh
        Set objLista = Nothing
        If aCodigo > 0 Then PresentoTraspaso aCodigo
        Screen.MousePointer = 0
    Else
        Toolbar1.Buttons("imprimir").Enabled = False
        MsgBox "Debe seleccionar un local para poder consultar.", vbExclamation, "ATENCIÓN"
    End If
    Exit Sub
    
ErrBTXO:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al buscar la información."

End Sub
Private Sub BuscoTraspasoXOrigen()
On Error GoTo ErrBTXO
    
    If sNuevo Or sModificar Then Exit Sub
    
    Botones True, False, False, False, False, Toolbar1, Me
    MnuImprimir.Enabled = False: Toolbar1.Buttons("entregar").Enabled = False
    Toolbar1.Buttons("entregar").Enabled = False: MnuEntregar.Enabled = False
    Toolbar1.Buttons("rePrint").Enabled = False
    CodTraslado = 0
    LimpioIngreso
    vsConsulta.Rows = 1
    cIntermediario.ListIndex = -1
    cDestino.ListIndex = -1
    tComentario.Text = vbNullString
    
    If cOrigen.ListIndex <> -1 Then
        
        Cons = "Select Distinct(TraCodigo), Código = TraCodigo, Fecha = TraFecha,  Intermediario = CamNombre, Destino = LocNombre From Traspaso " _
                & " Left Outer Join Camion ON TraLocalIntermedio = CamCodigo " _
                & ", Local, RenglonTraspaso " _
                & " Where TraLocalOrigen = " & cOrigen.ItemData(cOrigen.ListIndex) _
                & " And TraAnulado IS Null And TraCodigo = RTrTraspaso And RTrPendiente > 0 " _
                & " And TraLocalDestino = LocCodigo ORDER BY TraFecha DESC"
                
        Screen.MousePointer = 11
        Dim objLista As New clsListadeAyuda
        Dim aCodigo As Long: aCodigo = 0
        If objLista.ActivarAyuda(cBase, Cons, 6000, 0, "Lista Ayuda") > 0 Then
            aCodigo = objLista.RetornoDatoSeleccionado(0)
        End If
        Me.Refresh
        Set objLista = Nothing
        If aCodigo > 0 Then PresentoTraspaso aCodigo
        Screen.MousePointer = 0
    Else
        Toolbar1.Buttons("imprimir").Enabled = False
        MsgBox "Debe seleccionar un local para poder consultar.", vbExclamation, "ATENCIÓN"
    End If
    Exit Sub
    
ErrBTXO:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al buscar la información."
End Sub

Private Sub PresentoTraspaso(ByVal lnCodTraspaso As Long)
On Error GoTo ErrPT
Dim aValor As Long

    vsConsulta.Rows = 1
    lblRecepciono.Caption = vbNullString: lblRecepciono.Tag = ""
    cOrigen.ListIndex = -1: cDestino.ListIndex = -1: cIntermediario.ListIndex = -1
    tComentario.Text = vbNullString
    Toolbar1.Buttons("rePrint").Tag = ""
    tDoc.Text = "": tDoc.Tag = ""
    Shape1.BackColor = &HC0C000
    
    Cons = "SELECT Traspaso.*, DocSerie, DocNumero From Traspaso LEFT OUTER JOIN Documento ON DocCodigo = TraRemito Where TraCodigo = " & lnCodTraspaso
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsAux.EOF Then
        MnuImprimir.Enabled = False: Toolbar1.Buttons("entregar").Enabled = False
        Toolbar1.Buttons("entregar").Enabled = False: MnuEntregar.Enabled = False
        cDestino.Tag = vbNullString: CodTraslado = 0
        RsAux.Close
        labCreado.Caption = ""
        MsgBox "No se encontro el traspaso seleccionado, verifique si otra terminal no lo elimino."
    
    Else
        
        LimpioIngreso
        BuscoCodigoEnCombo cOrigen, RsAux!TraLocalOrigen
        BuscoCodigoEnCombo cDestino, RsAux!TraLocalDestino
        If Not IsNull(RsAux!TraLocalIntermedio) Then BuscoCodigoEnCombo cIntermediario, RsAux!TraLocalIntermedio
        
        txtCodigo.Text = RsAux("TraCodigo")
        Botones True, IIf(IsNull(RsAux("TraSerie")) And IsNull(RsAux("TraAnulado")), True, False), IsNull(RsAux("TraAnulado")), False, False, Toolbar1, Me
        MnuImprimir.Enabled = IIf(IsNull(RsAux("TraSerie")) And IsNull(RsAux("TraAnulado")), True, False): Toolbar1.Buttons("imprimir").Enabled = MnuImprimir.Enabled
        Toolbar1.Buttons("rePrint").Enabled = (Not IsNull(RsAux("TraSerie")) And IsNull(RsAux("TraAnulado"))) Or Not IsNull(RsAux("TraRemito"))
        
        'NO CONTROLO MAS LA SUCURSAL ya que se reimprime en papel blanco.
'        If Not IsNull(RsAux("TraSucursal")) Then
'            Toolbar1.Buttons("rePrint").Enabled = (Toolbar1.Buttons("rePrint").Enabled And (RsAux("TraSucursal") = paCodigoDeSucursal))
'        End If
        
        If Not IsNull(RsAux!TraFModificacion) Then
            cIntermediario.Tag = RsAux!TraFModificacion
        Else
            cIntermediario.Tag = ""
        End If
    
    '.....................................................................................................................................
    'Si no hice el documento
        MnuEntregar.Enabled = (IsNull(RsAux!TraRemito) And IsNull(RsAux("TraSerie")) And IsNull(RsAux("TraAnulado")))
        Toolbar1.Buttons("entregar").Enabled = MnuEntregar.Enabled
        cDestino.Tag = IIf(MnuEntregar.Enabled, "0", "1")
    '.....................................................................................................................................
    
        If Not IsNull(RsAux!TraFecha) Then labCreado.Caption = Format(RsAux!TraFecha, "d-Mmm-yyyy") Else labCreado.Caption = ""
        If Not IsNull(RsAux!TraUsuarioInicial) Then labUsuario.Caption = Trim(BuscoUsuario(RsAux!TraUsuarioInicial, False, False, True)) Else labUsuario.Caption = ""
        If Not IsNull(RsAux!TraComentario) Then tComentario.Text = Trim(RsAux!TraComentario) Else tComentario.Text = ""
        lnCodTraspaso = RsAux!TraCodigo
        
        
        If Not IsNull(RsAux("TraSerie")) Then
            tDoc.Text = Trim(RsAux("TraSerie")) & "-" & Trim(RsAux("TraNumero"))
        ElseIf Not IsNull(RsAux("TraRemito")) Then
            tDoc.Text = Trim(RsAux("DocSerie")) & "-" & Trim(RsAux("DocNumero")): tDoc.Tag = RsAux("TraRemito")
        End If
        
        If Not IsNull(RsAux("TraAnulado")) Then
            Shape1.BackColor = &H80FF&
            MsgBox "El traslado seleccionado está anulado.", vbExclamation, "Atención"
        End If
        
        If Not IsNull(RsAux("TraFechaEntregado")) Then
            lblRecepciono.Caption = "Recepcionado: " & Format(RsAux("TraFechaEntregado"), "dd/MM/yy hh:mm")
            lblRecepciono.Tag = RsAux("TraFechaEntregado")
        End If
        
        RsAux.Close
        
        Cons = "Select RenglonTraspaso.*, ArtNombre, EsMAbreviacion From RenglonTraspaso, Articulo, EstadoMercaderia" _
            & " Where RTrTraspaso = " & lnCodTraspaso & " And RTrArticulo = ArtID And RTrEstado = ESMCodigo"
            
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        Do While Not RsAux.EOF
            With vsConsulta
                .AddItem RsAux!RTrCantidad
                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!EsMAbreviacion)
                .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 3) = StockLocalArticuloyEstado(RsAux!RTrArticulo, RsAux!RTrEstado)
                'Data..........................................................
                aValor = RsAux!RTrArticulo: .Cell(flexcpData, .Rows - 1, 0) = aValor
                aValor = RsAux!RTrEstado: .Cell(flexcpData, .Rows - 1, 1) = aValor
                aValor = RsAux!RTrCantidad: .Cell(flexcpData, .Rows - 1, 2) = aValor    'Me guardo la cantidad
                If Not IsNull(RsAux!RTrPendiente) Then aValor = RsAux!RTrPendiente Else aValor = 0
                .Cell(flexcpData, .Rows - 1, 3) = aValor   'Me guardo lo pendiente
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
        CodTraslado = lnCodTraspaso
    End If
    Exit Sub
    
ErrPT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al levantar la información."
End Sub

Private Sub AccionNuevo()
    
'    If paDContado = "" Then
'        MsgBox "No tiene asignado un documento de traslado, no podrá crear traslados de su sucursal u otras en la que intervengan camioneros.", vbExclamation, "ATENCIÓN"
'        'Exit Sub
'    End If
    
    cDestino.Tag = vbNullString
    txtCodigo.Text = vbNullString
    LimpioIngreso
    HabilitoIngreso
    cOrigen.ListIndex = -1: cIntermediario.ListIndex = -1: cDestino.ListIndex = -1
    vsConsulta.Rows = 1
    labCreado.Caption = "": labUsuario.Caption = ""
    Foco cOrigen
    sNuevo = True
    Botones False, False, False, True, True, Toolbar1, Me
    MnuImprimir.Enabled = False: Toolbar1.Buttons("imprimir").Enabled = False
    MnuEntregar.Enabled = False: Toolbar1.Buttons("entregar").Enabled = False
    Toolbar1.Buttons("rePrint").Enabled = False
    Shape1.BackColor = &HC0C000
End Sub

Private Sub BuscoTraspasoRealizadoXIntermediario()
On Error GoTo ErrBTXI
    
    If sNuevo Or sModificar Then Exit Sub
    CodTraslado = 0
    LimpioIngreso
    vsConsulta.Rows = 1
    cOrigen.ListIndex = -1
    cDestino.ListIndex = -1
    tComentario.Text = vbNullString
    lblRecepciono.Caption = vbNullString
    MnuImprimir.Enabled = False: Toolbar1.Buttons("entregar").Enabled = False
    Toolbar1.Buttons("entregar").Enabled = False: MnuEntregar.Enabled = False
    
    
    If cIntermediario.ListIndex <> -1 Then
        
        Cons = "Select Distinct(TraCodigo), Código = TraCodigo, Fecha = TraFecha,  Origen = O.LocNombre, Destino = D.LocNombre From Traspaso " _
                & " ,Camion, Local D, Local O " _
                & " Where TraLocalIntermedio = " & cIntermediario.ItemData(cIntermediario.ListIndex) _
                & " And TraFechaEntregado >= '" & Format(Date - 3, "mm/dd/yyyy") & "'" _
                & " And TraAnulado IS Null And TraLocalDestino = D.LocCodigo" _
                & " And TraLocalOrigen = O.LocCodigo ORDER BY TraFecha DESC"
                
        Screen.MousePointer = 11
        Dim objLista As New clsListadeAyuda
        Dim aCodigo As Long: aCodigo = 0
        If objLista.ActivarAyuda(cBase, Cons, 6000, 0, "Lista Ayuda") > 0 Then
            aCodigo = objLista.RetornoDatoSeleccionado(0)
        End If
        Me.Refresh
        Set objLista = Nothing
        If aCodigo > 0 Then PresentoTraspaso aCodigo
        Screen.MousePointer = 0
    Else
        MsgBox "Debe seleccionar un local intermediario para poder consultar.", vbExclamation, "ATENCIÓN"
    End If
    Exit Sub
    
ErrBTXI:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al buscar la información."
End Sub
Private Sub BuscoTraspasoXIntermediario()
On Error GoTo ErrBTXI
    
    If sNuevo Or sModificar Then Exit Sub
    CodTraslado = 0
    LimpioIngreso
    vsConsulta.Rows = 1
    cOrigen.ListIndex = -1
    cDestino.ListIndex = -1
    tComentario.Text = vbNullString
    lblRecepciono.Caption = vbNullString
    MnuImprimir.Enabled = False: Toolbar1.Buttons("entregar").Enabled = False
    Toolbar1.Buttons("entregar").Enabled = False: MnuEntregar.Enabled = False
    
    If cIntermediario.ListIndex <> -1 Then
        
        Cons = "Select Distinct(TraCodigo), Código = TraCodigo, Fecha = TraFecha,  Origen = O.LocNombre, Destino = D.LocNombre From Traspaso " _
                & " ,Camion, Local D, Local O, RenglonTraspaso " _
                & " Where TraLocalIntermedio = " & cIntermediario.ItemData(cIntermediario.ListIndex) _
                & " And TraAnulado IS Null And TraCodigo = RTrTraspaso And RTrPendiente > 0 " _
                & " And TraLocalDestino = D.LocCodigo" _
                & " And TraLocalOrigen = O.LocCodigo"
                
        Screen.MousePointer = 11
        Dim objLista As New clsListadeAyuda
        Dim aCodigo As Long: aCodigo = 0
        
        If objLista.ActivarAyuda(cBase, Cons, 6000, 0, "Lista Ayuda") > 0 Then
            aCodigo = objLista.RetornoDatoSeleccionado(0)
        End If
        Me.Refresh
        Set objLista = Nothing
        If aCodigo > 0 Then PresentoTraspaso aCodigo
        Screen.MousePointer = 0
    Else
        MsgBox "Debe seleccionar un local intermediario para poder consultar.", vbExclamation, "ATENCIÓN"
    End If
    Exit Sub
    
ErrBTXI:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al buscar la información."

End Sub
Private Sub BuscoTraspasoRealizadosXDestino()
On Error GoTo ErrBTXD
    
    If sNuevo Or sModificar Then Exit Sub
    
    CodTraslado = 0
    LimpioIngreso
    vsConsulta.Rows = 1
    cIntermediario.ListIndex = -1
    cOrigen.ListIndex = -1
    MnuImprimir.Enabled = False: Toolbar1.Buttons("entregar").Enabled = False
    Toolbar1.Buttons("entregar").Enabled = False: MnuEntregar.Enabled = False
    tComentario.Text = vbNullString
    lblRecepciono.Caption = vbNullString
    
    If cDestino.ListIndex <> -1 Then
        
        Cons = "Select Distinct(TraCodigo), Código = TraCodigo, Fecha = TraFecha, Origen = LocNombre, Intermediario = CamNombre From Traspaso " _
                & " Left Outer Join Camion ON TraLocalIntermedio = CamCodigo " _
                & ", Local" _
                & " Where TraLocalDestino = " & cDestino.ItemData(cDestino.ListIndex) _
                & " And TraAnulado IS Null And TraFechaEntregado >= '" & Format(Date - 3, "mm/dd/yyyy") & "'" _
                & " And TraLocalOrigen = LocCodigo ORDER BY TraFecha DESC"
                
        Screen.MousePointer = 11
        Dim objLista As New clsListadeAyuda
        Dim aCodigo As Long: aCodigo = 0
        If objLista.ActivarAyuda(cBase, Cons, 6000, 0, "Lista Ayuda") > 0 Then
            aCodigo = objLista.RetornoDatoSeleccionado(0)
        End If
        Me.Refresh
        Screen.MousePointer = 0
        Set objLista = Nothing
        If aCodigo > 0 Then PresentoTraspaso aCodigo
    Else
        MsgBox "Debe seleccionar un local de destino para poder consultar.", vbExclamation, "ATENCIÓN"
    End If
    Exit Sub
    
ErrBTXD:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al buscar la información."
    
End Sub
Private Sub BuscoTraspasoXDestino()
On Error GoTo ErrBTXD
    
    If sNuevo Or sModificar Then Exit Sub
    
    CodTraslado = 0
    LimpioIngreso
    vsConsulta.Rows = 1
    cIntermediario.ListIndex = -1
    cOrigen.ListIndex = -1
    MnuImprimir.Enabled = False: Toolbar1.Buttons("entregar").Enabled = False
    Toolbar1.Buttons("entregar").Enabled = False: MnuEntregar.Enabled = False
    tComentario.Text = vbNullString
    lblRecepciono.Caption = vbNullString
    
    If cDestino.ListIndex <> -1 Then
        
        Cons = "Select Distinct(TraCodigo), Código = TraCodigo, Fecha = TraFecha, Origen = LocNombre, Intermediario = CamNombre From Traspaso " _
                & " Left Outer Join Camion ON TraLocalIntermedio = CamCodigo " _
                & ", Local, RenglonTraspaso " _
                & " Where TraLocalDestino = " & cDestino.ItemData(cDestino.ListIndex) _
                & " And TraAnulado IS Null And TraCodigo = RTrTraspaso And RTrPendiente > 0 " _
                & " And TraLocalOrigen = LocCodigo ORDER BY TraFecha DESC"
                
        Screen.MousePointer = 11
        Dim objLista As New clsListadeAyuda
        Dim aCodigo As Long: aCodigo = 0
        
        If objLista.ActivarAyuda(cBase, Cons, 6000, 0, "Lista Ayuda") > 0 Then
            aCodigo = objLista.RetornoDatoSeleccionado(0)
        End If
        Me.Refresh
        
        Set objLista = Nothing
        If aCodigo > 0 Then PresentoTraspaso aCodigo
        Screen.MousePointer = 0
    Else
        MsgBox "Debe seleccionar un local de destino para poder consultar.", vbExclamation, "ATENCIÓN"
    End If
    Exit Sub
    
ErrBTXD:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al buscar la información."

End Sub

Private Sub AccionGrabar()
Dim Msg As String

'    If paDContado = "" Then
'        MsgBox "No tiene asignado un documento, no podrá editar el traslado.", vbExclamation, "ATENCIÓN"
'        Exit Sub
'    End If

    If tUsuario.Tag = vbNullString Then
        MsgBox " Debe ingresar su código de usuario.", vbExclamation, "ATENCIÓN"
        tUsuario.SetFocus: Exit Sub
    End If

    If cOrigen.ListIndex = -1 Or cDestino.ListIndex = -1 Then
        MsgBox "Los locales origen y destino son obligatorios.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If cIntermediario.Text <> vbNullString And cIntermediario.ListIndex = -1 Then
        MsgBox "El local intermediario no es correcto.", vbExclamation, "ATENCIÓN"
        cIntermediario.SetFocus: Exit Sub
    End If
    
    If cDestino.ItemData(cDestino.ListIndex) = cOrigen.ItemData(cOrigen.ListIndex) Then
        MsgBox "Selecciono el mismo local de origen y destino.", vbExclamation, "ATENCIÓN"
        cDestino.SetFocus: Exit Sub
    End If
    
    If Not clsGeneral.TextoValido(tComentario.Text) Then
        MsgBox "Se ingreso un carácter no válido en el comentario.", vbExclamation, "ATENCIÓN"
        tComentario.SetFocus: Exit Sub
    End If
    
    Msg = vbNullString
    If vsConsulta.Rows = 1 Then
        MsgBox "No hay artículos a trasladar.", vbExclamation, "ATENCIÓN": Exit Sub
    Else
        For I = 1 To vsConsulta.Rows - 1
            If Val(vsConsulta.Cell(flexcpText, I, 0)) > 0 Then Msg = "hay": Exit For
        Next
        If Msg = vbNullString Then
            MsgBox "No hay datos a trasladar.", vbExclamation, "ATENCIÓN": Exit Sub
        End If
        Msg = vbNullString
    End If
    
    If MsgBox("¿Confirma grabar el traslado de mercadería.", vbQuestion + vbYesNo, "IMPRIMIR") = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    If sNuevo Then
        NuevoTraslado Msg
    Else
        ModificoTraslado Msg
    End If
    
    tUsuario.Text = "": tUsuario.Tag = ""
    If Not sNuevo And Not sModificar Then
        LimpioIngreso
        DeshabilitoIngreso
        Botones True, True, True, False, False, Toolbar1, Me
        cOrigen.Enabled = True
        cDestino.Enabled = True
        cIntermediario.Enabled = True
        PresentoTraspaso CodTraslado
    End If
    Screen.MousePointer = vbDefault
    Exit Sub

End Sub
Private Sub NuevoTraslado(Msg As String)
Dim iAux As Long
    
    FechaDelServidor
    
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrRelajo
    
    'Primero inserto en la tabla traspaso, luego obtengo su código.
    'Seguido inserto los renglones y cambio los movimientos de stock.
    
    Cons = "Select MAX(TraCodigo) From Traspaso"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If IsNull(RsAux(0)) Then iAux = 0 Else iAux = RsAux(0)
    RsAux.Close
    
    Cons = "Insert into Traspaso (TraFecha, TraFModificacion, TraLocalOrigen, TraLocalIntermedio,  TraLocalDestino, TraComentario, TraFechaEntregado, TraUsuarioInicial, TraUsuarioFinal) " _
        & " Values ('" & Format(gFechaServidor, sqlFormatoFH) & "','" & Format(gFechaServidor, sqlFormatoFH) & "'" _
        & ", " & cOrigen.ItemData(cOrigen.ListIndex)
    
    If cIntermediario.ListIndex > -1 Then
        Cons = Cons & ", " & cIntermediario.ItemData(cIntermediario.ListIndex)
    Else
        Cons = Cons & " , Null"
    End If
    Cons = Cons & ", " & cDestino.ItemData(cDestino.ListIndex)
        
    If Trim(tComentario.Text) <> vbNullString Then
        Cons = Cons & ", '" & Trim(tComentario.Text) & "'"
    Else
        Cons = Cons & ", Null"
    End If
    Cons = Cons & ", Null" _
        & "," & tUsuario.Tag & ",Null)"
    cBase.Execute (Cons)
    
    'Saco el código del insertado.
    Cons = "Select MAX(TraCodigo) From Traspaso Where TraCodigo > " & iAux _
        & " And TraLocalOrigen = " & cOrigen.ItemData(cOrigen.ListIndex) _
        & " And TraLocalDestino = " & cDestino.ItemData(cDestino.ListIndex)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not IsNull(RsAux(0)) Then iAux = RsAux(0) Else iAux = 0
    RsAux.Close
    loc_DBAddRenglones iAux
    cBase.CommitTrans
    
    CodTraslado = iAux
    sNuevo = False
    Toolbar1.Buttons("imprimir").Enabled = True:  MnuImprimir.Enabled = True
    'Si no lleva intermediario no es necesario asignarle la mercadería, lo hace la impresión.
    If cIntermediario.ListIndex > -1 Then
        Toolbar1.Buttons("entregar").Enabled = True:  MnuEntregar.Enabled = True
    Else
        Toolbar1.Buttons("entregar").Enabled = False:  MnuEntregar.Enabled = False
    End If
    Exit Sub
ErrBT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al intentar iniciar la transaccion."
    Exit Sub
    
ErrRelajo:
    Resume Resumo
    
Resumo:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError Msg, Err.Description
    
End Sub
Private Sub AccionModificar()
On Error GoTo ErrAM
    
'    If paDContado = "" Then
'        MsgBox "No tiene asignado un documento, no podrá editar el traslado.", vbExclamation, "ATENCIÓN"
'        Exit Sub
'    End If
    
    PresentoTraspaso CLng(CodTraslado)
    sModificar = True
    If CodTraslado = 0 Then Exit Sub
    HabilitoIngreso
    
    Botones False, False, False, True, True, Toolbar1, Me
    Toolbar1.Buttons("imprimir").Enabled = False: MnuImprimir.Enabled = False
    Toolbar1.Buttons("entregar").Enabled = False: MnuEntregar.Enabled = False
    Toolbar1.Buttons("rePrint").Enabled = False
    cDestino.Enabled = False
    If cDestino.Tag = "" Or cDestino.Tag = "0" Then
        Foco cArticulo
        cOrigen.Enabled = True
        cIntermediario.Enabled = True
        'Si no se recepcionó ningún artículo, permito modificar el local de destino.
        Cons = "Select * from RenglonTraspaso Where RTrTraspaso = " & CLng(CodTraslado) _
            & " And RTrCantidad <> RTrPendiente "
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly)
        If RsAux.EOF Then
          cDestino.Enabled = True: cDestino.BackColor = Obligatorio
        End If
        RsAux.Close
    ElseIf cDestino.Tag = "1" Then
        cOrigen.Enabled = False
        cIntermediario.Enabled = False
        cDestino.Enabled = False
        cArticulo.Enabled = False
        tCantidad.Enabled = False
        cEstado.Enabled = False
        vsConsulta.SetFocus
    End If
    Exit Sub
    
ErrAM:
    clsGeneral.OcurrioError "Error al acceder a modo de edición. " & Trim(Err.Description)
    AccionCancelar
    
End Sub
Private Sub ModificoTraslado(Msg As String)

    FechaDelServidor
    On Error GoTo ErrBT
    cBase.BeginTrans
    On Error GoTo ErrRelajo
    
    Cons = "Select * From Traspaso Where TraCodigo = " & CLng(CodTraslado)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        Msg = "No se encontró el traslado seleccionado, verifique si no fue eliminado."
        RsAux.Close
        RsAux.Edit
    Else
        'Marco señal de si fue impreso.
        If Not IsNull(RsAux!TraFImpreso) Or Not IsNull(RsAux("TraSerie")) Then
            RsAux.Close
            cBase.RollbackTrans
            MsgBox "Traslado con documento asignado, refresque la información.", vbExclamation, "Atención"
            Exit Sub
        End If
        If Not IsNull(RsAux!TraFModificacion) Then
            If RsAux!TraFModificacion <> CDate(cIntermediario.Tag) Then
                RsAux.Close
                cBase.RollbackTrans
                MsgBox "El traslado fue modificado por otra terminal, verifique.", vbInformation, "ATENCIÓN"
                Exit Sub
            End If
        End If
        
        RsAux.Edit
        RsAux!TraLocalOrigen = cOrigen.ItemData(cOrigen.ListIndex)
        
        If cIntermediario.ListIndex > -1 Then
            RsAux!TraLocalIntermedio = cIntermediario.ItemData(cIntermediario.ListIndex)
        Else
            RsAux!TraLocalIntermedio = Null
        End If
        RsAux!TraFModificacion = Format(gFechaServidor, sqlFormatoFH)
        If Trim(tComentario.Text) = vbNullString Then
            RsAux!TraComentario = Null
        Else
            RsAux!TraComentario = Trim(tComentario.Text)
        End If
        RsAux!TraLocalDestino = cDestino.ItemData(cDestino.ListIndex)
        RsAux!TraUsuarioInicial = tUsuario.Tag
        RsAux.Update
        RsAux.Close
    End If
    'Elimino los renglones
    cBase.Execute "Delete From RenglonTraspaso Where RTrTraspaso = " & CodTraslado
    loc_DBAddRenglones CodTraslado
    cBase.CommitTrans
    
    sModificar = False
    Exit Sub
    
ErrBT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al intentar iniciar la transaccion."
    Exit Sub
    
ErrRelajo:
    Resume Resumo
    
Resumo:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError Msg, Err.Description
    
End Sub

Private Sub AccionEliminar()
Dim sImpreso As Boolean
Dim Rs As rdoResultset
Dim Msg As String
Dim lnCodUsuario As Long
Dim strUsuario As String


    If tDoc.Text <> "" Then
        MsgBox "Ya se emitió un eRemito, no se puede anular, si desea volver atrás los movimientos haga el movimiento contrario.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If


    If tDoc.Text <> "" Or cDestino.Tag = "1" Then
        
        Cons = "SELECT RTrArticulo FROM RenglonTraspaso WHERE RTrTraspaso = " & CodTraslado _
            & " And RTrCantidad <> ISNULL(RTrPendiente, 0)"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Close
            Screen.MousePointer = vbDefault
            MsgBox "No se puede eliminar el traslado debido a que ya se recepcionó la mercadería.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        RsAux.Close
    
        'Si hay intermediario y no se recepcionó entonces dejo anular el traslado.
        If Not IsDate(lblRecepciono.Tag) And cIntermediario.ListIndex > -1 Then
            'Debe existir intermediario por lo tanto dejo tirar para atrás todo el traslado.
            loc_AnuloDocumento
        Else
            MsgBox "Ya se emitió un eRemito, no se puede anular, si desea volver atrás los movimientos haga el movimiento contrario.", vbExclamation, "ATENCIÓN"
        End If
        Exit Sub
    End If


    If paCodigoDeSucursal <> cOrigen.ItemData(cOrigen.ListIndex) Then
        If MsgBox("ATENCIÓN su terminal no pertenece a la sucursal origen." & vbCrLf & vbCrLf & "¿Confirma continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Posible Error") = vbNo Then Exit Sub
    End If
        
    If MsgBox("¿Confirma eliminar el Traslado seleccionado?", vbQuestion + vbYesNo, "Eliminar") = vbNo Then Exit Sub
        
    lnCodUsuario = 0
    strUsuario = vbNullString
    strUsuario = InputBox("Ingrese su dígito de Usuario.", "Eliminar traslado")
        
    If Not IsNumeric(strUsuario) Then
        MsgBox "No se ingresó un dígito correcto", vbExclamation, "ATENCIÓN"
        Exit Sub
    Else
        lnCodUsuario = BuscoUsuarioDigito(CLng(strUsuario), True)
        If lnCodUsuario = 0 Then Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    Cons = "Select * From RenglonTraspaso Where RTrTraspaso = " & CodTraslado _
        & " And RTrCantidad <> RTrPendiente"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        sImpreso = False
        
        Cons = "Select * From Traspaso Where TRaCodigo = " & CodTraslado
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not IsNull(RsAux!TraFImpreso) Then
            'Por el sistema viejo.
            If IsNull(RsAux("TraSerie")) Then
                sImpreso = True
            Else
                'Sistema nuevo NO DEJO ELIMINAR
                Screen.MousePointer = 0
                RsAux.Close
                MsgBox "El traslado tiene documento, no se puede eliminar.", vbExclamation, "Atención"
                Exit Sub
            End If
        End If
        RsAux.Close
         
         
        On Error GoTo ErrBT
        cBase.BeginTrans
        On Error GoTo ErrRelajo
        
        If sImpreso And cIntermediario.ListIndex > -1 Then
        
            Cons = "Select * From RenglonTraspaso Where RTrTraspaso = " & CodTraslado
            Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            
            Do While Not Rs.EOF
            
                Cons = "Select * From StockLocal " _
                    & " Where StLArticulo = " & Rs!RTrArticulo _
                    & " And StlTipoLocal = " & TipoLocal.Deposito _
                    & " And StLLocal = " & cOrigen.ItemData(cOrigen.ListIndex) _
                    & " And StLEstado = " & Rs!RTrEstado
                
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
                If RsAux.EOF Then
                    'No tiene más ese local.
                    RsAux.AddNew
                    RsAux!StLArticulo = Rs!RTrArticulo
                    RsAux!StLTipoLocal = TipoLocal.Deposito
                    RsAux!StlLocal = cOrigen.ItemData(cOrigen.ListIndex)
                    RsAux!StLEstado = Rs!RTrEstado
                    RsAux!StLCantidad = Rs!RTrCantidad
                    RsAux.Update
                Else
                    RsAux.Edit
                    RsAux!StLCantidad = RsAux!StLCantidad + Rs!RTrCantidad
                    RsAux.Update
                End If
                RsAux.Close
                
                MarcoMovimientoStockFisico lnCodUsuario, TipoLocal.Deposito, cOrigen.ItemData(cOrigen.ListIndex), CLng(Rs!RTrArticulo), Rs!RTrCantidad, Rs!RTrEstado, 1, TipoDocumento.Traslados, CodTraslado
                
                'Quito del intermediario.
                Cons = "Select * From StockLocal " _
                    & " Where StLArticulo = " & Rs!RTrArticulo _
                    & " And StlTipoLocal = " & TipoLocal.Camion _
                    & " And StLLocal = " & cIntermediario.ItemData(cIntermediario.ListIndex) _
                    & " And StLEstado = " & Rs!RTrEstado
                
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
                If RsAux.EOF Then
                    Msg = "No se encontraron todos los artículos en el local intermediario."
                    RsAux.Close
                    RsAux.Edit
                Else
                    RsAux.Edit
                    RsAux!StLCantidad = RsAux!StLCantidad - Rs!RTrCantidad
                    RsAux.Update
                End If
                RsAux.Close
                MarcoMovimientoStockFisico lnCodUsuario, TipoLocal.Camion, cIntermediario.ItemData(cIntermediario.ListIndex), Rs!RTrArticulo, Rs!RTrCantidad, Rs!RTrEstado, -1, TipoDocumento.Traslados, CodTraslado
                
                Rs.MoveNext
            Loop
            Rs.Close
            
        End If
        
        'Elimino los renglones y luego el traslado.
        Cons = " Delete RenglonTraspaso Where RTrTraspaso = " & CodTraslado
        cBase.Execute (Cons)
        
        Cons = " Delete Traspaso Where TraCodigo = " & CodTraslado
        cBase.Execute (Cons)
        
        cBase.CommitTrans
        'Limpio los datos.
        Botones True, False, False, False, False, Toolbar1, Me
        MnuImprimir.Enabled = False
        Toolbar1.Buttons("imprimir").Enabled = False
        Toolbar1.Buttons("entregar").Enabled = False
        MnuEntregar.Enabled = False
        Toolbar1.Buttons("rePrint").Enabled = False

        CodTraslado = 0
        cOrigen.ListIndex = -1
        cDestino.ListIndex = -1
        cIntermediario.ListIndex = -1
        tComentario.Text = vbNullString
        lblRecepciono.Caption = vbNullString: lblRecepciono.Tag = ""
        tUsuario.Text = vbNullString
    Else
        RsAux.Close
        Screen.MousePointer = vbDefault
        MsgBox "No se puede eliminar el traslado debido a que ya se entrego parcialmente la mercadería.", vbExclamation, "ATENCIÓN"
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
ErrBT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al intentar iniciar la transaccion."
    Exit Sub
    
ErrRelajo:
    Resume Resumo
    
Resumo:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError Msg, Err.Description

End Sub

Private Sub txtCodigo_Change()
On Error Resume Next

    If (CodTraslado > 0) Then CodTraslado = 0
    If (Me.ActiveControl.Name = "txtCodigo") Then
        CodTraslado = 0: cIntermediario.Tag = vbNullString
        Botones True, False, False, False, False, Toolbar1, Me
        Toolbar1.Buttons("imprimir").Enabled = False: MnuImprimir.Enabled = False
        Toolbar1.Buttons("entregar").Enabled = False: MnuEntregar.Enabled = False
        Toolbar1.Buttons("rePrint").Enabled = False
        vsConsulta.Rows = 1
        tComentario.Text = vbNullString: labCreado.Caption = vbNullString: labUsuario.Caption = vbNullString
        lblRecepciono.Caption = vbNullString: lblRecepciono.Tag = ""
    End If
    
    
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
On Error GoTo errC
    If KeyAscii = vbKeyReturn Then
        If CodTraslado = 0 And IsNumeric(txtCodigo.Text) Then
            PresentoTraspaso Val(txtCodigo.Text)
        End If
    End If
    Exit Sub
errC:
    clsGeneral.OcurrioError "Error al cargar el traslado.", Err.Description, "Buscar traslado"
End Sub

Private Sub vsConsulta_GotFocus()
    Status.SimpleText = " Seleccione un artículo y modifique su cantidad ('+', '-'), elimina con ('Supr')."
End Sub

Private Sub vsConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (sModificar Or sNuevo) Then Exit Sub
    
    With vsConsulta
        If vsConsulta.Rows > 1 Then
            Select Case KeyCode
                Case vbKeyAdd
                    If cDestino.Tag <> "1" Then
                        'NO Fue impreso
                        .Cell(flexcpText, .Row, 0) = CLng(.Cell(flexcpText, .Row, 0)) + 1
                    Else
                        'Fue impreso
                        If CLng(.Cell(flexcpText, .Row, 0)) < CLng(.Cell(flexcpData, .Row, 2)) Then
                            .Cell(flexcpText, .Row, 0) = CLng(.Cell(flexcpText, .Row, 0)) + 1
                        End If
                    End If
    
                Case vbKeySubtract
                    
                    If cDestino.Tag = "1" Then Exit Sub
                    
                    If CLng(.Cell(flexcpText, .Row, 0)) >= 1 Then
                        If .Cell(flexcpData, .Row, 3) <> vbNullString Then
                            'Veo si ya le di mercadería.
                            If CInt(.Cell(flexcpData, .Row, 3)) = 0 Then
                                MsgBox "No se pueden quitar más artículos debido a que ya se entrego la cantidad total de este artículo.", vbExclamation, "ATENCIÓN"
                                Exit Sub
                            End If
                            If CLng(.Cell(flexcpText, .Row, 0)) > CLng(.Cell(flexcpData, .Row, 2)) - CLng(.Cell(flexcpData, .Row, 3)) Then
                                .Cell(flexcpText, .Row, 0) = CLng(.Cell(flexcpText, .Row, 0)) - 1
                            Else
                                MsgBox "No se pueden quitar más artículos debido a que ya entrego la cantidad restante.", vbExclamation, "ATENCIÓN"
                            End If
                        Else
                            .Cell(flexcpText, .Row, 0) = CLng(.Cell(flexcpText, .Row, 0)) - 1
                        End If
                    End If
    
                Case vbKeyDelete
                    If cDestino.Tag = "1" Then Exit Sub
                    'Tengo que ver si entrego algo de este artículo.
'                    If .Cell(flexcpData, .Row, 3) <> "" And CLng(.Cell(flexcpData, .Row, 2)) > 0 Then
'                        MsgBox "No se pueden eliminar artículos que tienen entregas, si no quiere trasladar este artículo intente poner cantidad igual a cero.", vbExclamation, "ATENCIÓN"
'                        Exit Sub
'                    End If

                    If sNuevo Then
                        .RemoveItem .Row
                    Else
                        .Cell(flexcpText, .Row, 0) = Val(.Cell(flexcpData, .Row, 2)) - Val(.Cell(flexcpData, .Row, 3))
                    End If
                
                Case vbKeyReturn: If tComentario.Enabled Then tComentario.SetFocus
                   
            End Select
        End If
    End With

End Sub

Private Sub vsConsulta_LostFocus()
    Status.SimpleText = vbNullString
End Sub

Private Sub OcultoTodo(Optional ByVal bCleanDoc As Boolean = True)
    Botones True, False, False, False, False, Toolbar1, Me
    Toolbar1.Buttons("imprimir").Enabled = False: MnuImprimir.Enabled = False
    Toolbar1.Buttons("entregar").Enabled = False: MnuEntregar.Enabled = False
    Toolbar1.Buttons("rePrint").Enabled = False
    vsConsulta.Rows = 1
    tComentario.Text = vbNullString: labCreado.Caption = vbNullString: labUsuario.Caption = vbNullString
    lblRecepciono.Caption = vbNullString: lblRecepciono.Tag = ""
    tDoc.Tag = ""
    If bCleanDoc Then tDoc.Text = ""
End Sub

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

Private Function BuscoUsuario(Codigo As Long, Optional Identificacion As Boolean = False, Optional Digito As Boolean = False, Optional Iniciales As Boolean = False)
Dim RsUsr As rdoResultset
Dim aRetorno As String: aRetorno = ""
    
    On Error Resume Next
    Cons = "Select * from Usuario Where UsuCodigo = " & Codigo
    Set RsUsr = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsUsr.EOF Then
        If Identificacion Then aRetorno = Trim(RsUsr!UsuIdentificacion)
        If Digito Then aRetorno = Trim(RsUsr!UsuDigito)
        If Iniciales Then aRetorno = Trim(RsUsr!UsuInicial)
    End If
    RsUsr.Close
    BuscoUsuario = aRetorno
End Function

'Private Function fnc_PrintDocumento(ByVal lCodTra As Long, ByVal iUser As Integer) As Boolean
'    'On Error Resume Next
'    On Error GoTo errMPD
'    SeteoImpresoraPorDefecto paPrintCtdoD
'    With vsPrint
'        .Device = paPrintCtdoD
'        .PaperSize = paPrintCtdoPaperSize
'        .PaperBin = paPrintCtdoB
'        .Header = ""
'        .Footer = ""
'        .Orientation = orPortrait
'    End With
'    'On Error GoTo errMPD
'    Dim oPrint As New clsPrintManager
'    With oPrint
'     '   SeteoImpresoraPorDefecto paPrintCtdoD
'        .SetDevice paPrintCtdoD, paPrintCtdoB, paPrintCtdoPaperSize
'        If .LoadFileData(prmReportes & "\rpttraslado.txt") Then
'            fnc_PrintDocumento = .PrintDocumento("prg_Traslados_ImprimirDocumento " & lCodTra & ", " & iUser & ", '" & paDContado & "'", vsPrint)
'        End If
'    End With
'    Set oPrint = Nothing
'    Exit Function
'errMPD:
'    MsgBox "Error al imprimir el documento, error: " + Err.Description, vbCritical, "Impresión de traslados"
'End Function

'Public Sub ImprimirDocumento(ByVal lCodTra As Long, ByVal iUser As Integer)
'
'    Screen.MousePointer = 11
'
'    iUser = BuscoUsuario(CLng(iUser), False, True)
'
'    Cons = "Select Traspaso.*, RTrCantidad, ArtCodigo, ArtNombre From Traspaso, RenglonTraspaso, Articulo " & _
'            " Where TRaCodigo = " & CodTraslado & " And TraCodigo = RTrTraspaso And RTrArticulo = ArtID"
'    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
'    If RsAux.EOF Then
'        MsgBox "No se encontró el traslado, no se imprime.", vbExclamation, "Atención"
'        RsAux.Close
'        Screen.MousePointer = 0
'        Exit Sub
'    End If
'
'    modPrintDocumento.StarDocument
'    modPrintDocumento.SetDataDocument Format(RsAux("TraFImpreso"), "dd/mm/yy  hh:nn:ss"), paDContado, RsAux("TraSerie") & " " & RsAux("TraNumero"), iUser, CodTraslado, tComentario.Text
'    modPrintDocumento.SetClienteDocument "De: " & Trim(cOrigen.Text) & IIf(cIntermediario.ListIndex > -1, Space(10) & "Por: " & Trim(cIntermediario.Text), ""), "Para: " & Trim(cDestino.Text)
'
'    'Agrego los artículos.
'    Do While Not RsAux.EOF
'        modPrintDocumento.SetNewArticuloDocument Format(RsAux("ArtCodigo"), "0000000"), RsAux("RTrCantidad"), Trim(RsAux("ArtNombre"))
'        RsAux.MoveNext
'    Loop
'    RsAux.Close
'
'    Screen.MousePointer = 0
'    Dim sDef As String
'    sDef = Printer.DeviceName
'    SeteoImpresoraPorDefecto paPrintCtdoD
'    vsPrint.Device = paPrintCtdoD
'    vsPrint.PaperBin = paPrintCtdoB
'    vsPrint.PaperSize = paPrintCtdoPaperSize
'    modPrintDocumento.PrintDocument vsPrint
'    SeteoImpresoraPorDefecto sDef
'
'End Sub

Private Sub AccionImprimirAux()
Dim iLeftM As Long, iTopM As Long

    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub

    Dim sPDef As String
    sPDef = Printer.DeviceName
    SeteoImpresoraPorDefecto paPrintConfD
    
    'Determino si la imprsora donde voy a imprimir está instalada.
    Dim X As Printer
    Dim bolPrint As Boolean
    For Each X In Printers
        If Trim(X.DeviceName) = Trim(paPrintConfD) Then
            bolPrint = True
            Exit For
        End If
    Next
    
    If Not bolPrint Then
        MsgBox "La impresora " & Trim(paPrintConfD) & " no se encuentrá en la lista de impresoras, la impresión será destinada a la impresora por defecto.", vbInformation, "Atención"
    End If

    With vsPrint
        iLeftM = .MarginLeft
        iTopM = .MarginTop
        .Device = Trim(paPrintConfD)
        .Orientation = orPortrait
        .PaperSize = 1                     'Hoja carta
        .PaperBin = paPrintConfB  'Bandeja por defecto.
        .MarginLeft = 700
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
    
        .FileName = "Traslado de Mercaderia"
    
        EncabezadoListado vsPrint, "Traslado de Mercadería", False
        .FontBold = True
        .Paragraph = "Código de Traslado = " & CodTraslado
        .Paragraph = "Comentario:  " & Trim(tComentario.Text)
        .Paragraph = ""
        .FontBold = False
        vsConsulta.ExtendLastCol = False: .RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        .Paragraph = ""
        .EndDoc
        .PrintDoc False
        .MarginLeft = iLeftM: .MarginTop = iTopM
    End With
    SeteoImpresoraPorDefecto sPDef
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    SeteoImpresoraPorDefecto sPDef
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
    
End Sub

