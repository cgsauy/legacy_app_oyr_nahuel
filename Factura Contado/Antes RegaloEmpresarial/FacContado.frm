VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{5EA2D00A-68AC-4888-98E6-53F6035BBEE3}#1.3#0"; "CGSABuscarCliente.ocx"
Begin VB.Form FacContado 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación Contado"
   ClientHeight    =   5025
   ClientLeft      =   1095
   ClientTop       =   2400
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FacContado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8610
   Begin VB.Timer tmArticuloLimitado 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   3000
      Top             =   120
   End
   Begin prjBuscarCliente.ucBuscarCliente txtCliente 
      Height          =   285
      Left            =   780
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      Text            =   "_.___.___-_"
      DocumentoCliente=   1
      QueryFind       =   "EXEC [dbo].[prg_BuscarCliente] 0, '', '', '', '', '', '[KeyQuery]', 0, 0, '', '', 7"
      KeyQuery        =   "[KeyQuery]"
      NeedCheckDigit  =   0   'False
      Comportamiento  =   1
   End
   Begin VB.PictureBox picTransaccion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3960
      ScaleHeight     =   315
      ScaleWidth      =   2775
      TabIndex        =   49
      Top             =   90
      Width           =   2775
      Begin VB.TextBox txtTransaccion 
         Appearance      =   0  'Flat
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Text            =   "25"
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkRetiraAqui 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Retira aquí?"
      Height          =   255
      Left            =   2160
      TabIndex        =   48
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tArticulo 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Text            =   "8888"
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CheckBox chRetira 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Retira"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   855
   End
   Begin VB.CheckBox chPagaCheque 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pa&ga c/cheque"
      Height          =   255
      Left            =   4680
      TabIndex        =   26
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox tNombreC 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4380
      TabIndex        =   7
      Top             =   600
      Width           =   4095
   End
   Begin VB.CheckBox chNomDireccion 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   4140
      TabIndex        =   5
      Top             =   960
      Width           =   195
   End
   Begin VB.ComboBox cDireccion 
      Height          =   315
      Left            =   2580
      TabIndex        =   4
      Text            =   "cDireccion"
      Top             =   945
      Width           =   1515
   End
   Begin AACombo99.AACombo cPendiente 
      Height          =   315
      Left            =   1200
      TabIndex        =   21
      Top             =   3720
      Width           =   4995
      _ExtentX        =   8811
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
   Begin AACombo99.AACombo cEnvio 
      Height          =   315
      Left            =   7800
      TabIndex        =   17
      Top             =   1680
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
   Begin AACombo99.AACombo cMoneda 
      Height          =   315
      Left            =   7680
      TabIndex        =   32
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
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
   Begin VB.TextBox tVendedor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   4080
      MaxLength       =   3
      TabIndex        =   25
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox tComentarioDocumento 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   70
      TabIndex        =   28
      Top             =   4440
      Width           =   5895
   End
   Begin VB.TextBox tFRetiro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      TabIndex        =   23
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   7920
      MaxLength       =   3
      TabIndex        =   30
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox tCantidad 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   3840
      MaxLength       =   5
      TabIndex        =   11
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox tUnitario 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   6600
      MaxLength       =   12
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox tComentario 
      Height          =   315
      Left            =   4440
      MaxLength       =   15
      TabIndex        =   13
      Top             =   1680
      Width           =   2175
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   4770
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9024
            MinWidth        =   2
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   4604
            MinWidth        =   2
            Text            =   "F2-Modificar, F3-Nuevo, F4-Buscar"
            TextSave        =   "F2-Modificar, F3-Nuevo, F4-Buscar"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   1482
            MinWidth        =   1482
            Text            =   "F9-Envíos "
            TextSave        =   "F9-Envíos "
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrilla 
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2778
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
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
      Editable        =   -1  'True
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
   End
   Begin VB.TextBox tANombre 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      MaxLength       =   70
      TabIndex        =   42
      Top             =   4800
      Width           =   5295
   End
   Begin MSMask.MaskEdBox tBanco 
      Height          =   285
      Left            =   1200
      TabIndex        =   45
      Top             =   4800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   -2147483630
      PromptInclude   =   0   'False
      MaxLength       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##-###"
      PromptChar      =   "_"
   End
   Begin VB.Label lblInfoCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "&C.I.:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblRucPersona 
      BackStyle       =   0  'Transparent
      Caption         =   "213 783 510 017"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2280
      TabIndex        =   47
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lPrint 
      BackStyle       =   0  'Transparent
      Caption         =   "Impresora: Fact Colonia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   120
      TabIndex        =   46
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lTelCant 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel.(0):"
      Height          =   255
      Left            =   240
      TabIndex        =   41
      Top             =   960
      Width           =   555
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "&Vend.:"
      Height          =   255
      Left            =   3480
      TabIndex        =   24
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label53 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mentarios:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Pen&diente por:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "&LISTA"
      Height          =   255
      Left            =   300
      TabIndex        =   18
      Top             =   2460
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "En&viar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7800
      TabIndex        =   16
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuar&io:"
      Height          =   255
      Left            =   7200
      TabIndex        =   29
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6240
      TabIndex        =   40
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I.V.A."
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6240
      TabIndex        =   39
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sub total"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6240
      TabIndex        =   38
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label labSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   37
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label labIVA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   36
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label labTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   35
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label labArticulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Artículo. [F12 Servicio]"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ca&nt."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Precio Unitario"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6600
      TabIndex        =   14
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Moneda:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   31
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comen&tario"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lTelefono 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Trabajo1 099-645014"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   780
      TabIndex        =   3
      Top             =   960
      Width           =   1755
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nom&bre:"
      Height          =   255
      Left            =   3750
      TabIndex        =   6
      Top             =   600
      Width           =   675
   End
   Begin VB.Label labDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Niagara 2345"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4380
      TabIndex        =   34
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   855
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   8415
   End
   Begin VB.Label lANombre 
      BackStyle       =   0  'Transparent
      Caption         =   "A nombre de:"
      Height          =   255
      Left            =   2160
      TabIndex        =   44
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco:"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   4800
      Width           =   855
   End
   Begin VB.Menu MnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu MnuEmitir 
         Caption         =   "&Emitir Factura"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuEnvio 
         Caption         =   "&Realizar Envíos"
         Enabled         =   0   'False
         Shortcut        =   {F9}
      End
      Begin VB.Menu MnuOpUltimoEnvio 
         Caption         =   "&Visualizar último envío"
         Shortcut        =   ^U
      End
      Begin VB.Menu MnuOpRemito 
         Caption         =   "&Hacer remito a último contado"
         Shortcut        =   {F8}
      End
      Begin VB.Menu MnuLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpFactSinIVA 
         Caption         =   "Facturar sin IVA (Zona Franca)"
      End
      Begin VB.Menu MnuLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLimpiar 
         Caption         =   "&Limpiar Factura"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuLineRP 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOptTransacciones 
         Caption         =   "Ingresar transacciones"
      End
      Begin VB.Menu MnuLSalir 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVolver 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuPrinter 
      Caption         =   "Impresora"
      Begin VB.Menu MnuPrintDonde 
         Caption         =   "Dónde imprimo?"
      End
      Begin VB.Menu MnuPrintConfig 
         Caption         =   "Configurar"
      End
      Begin VB.Menu MnuPrintLine1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrintOpt 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu MnuEntrega 
      Caption         =   "Mercadería"
      Begin VB.Menu MnuEntRetirar 
         Caption         =   "Retira ahora en este local"
      End
   End
End
Attribute VB_Name = "FacContado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim aTexto As String
 
'Modificaciones
'   15-6-2001
        'Válido conexión.
        'Muestro el último envío almacenado.
        'Si el artículo está inhabilitado pero ingreso por código le pregunto si igual lo desea facturar.
'   18-6-2001
        'Verifico si hay ruc o si es empresa lleva cofis.
        'Ahora todos los artículos llevan cofis.
'   8/8/2001    Agregamos artículo es combo y sus concecuencias.
'   19-10-01    Carlos no quiere suceso cuando no tiene ingresado precio un artículo.
'   29-12       Elimine campo RUC y marco si no tiene dir y agregue Telef.
'   21-10-03    Agregue instalaciones
'   19-12-03    Si el cliente tiene ventas telef. doy aviso, si coinciden los art. se lo abro modal.
'   30-12-03    Agregue opción de ir a remito con el último contado hecho.
'   10-5-04     Cambie manejo de impresora agregando opciones.
'   18-1-05     Cambio en combo dirección carga por top.
'   28-9-05     Ajuste consulta con precios haciendo left outer.
'   28-10-05    Si supera prmImpCed y no tiene ced o la ced = 99999999 --> no dejo y doy aviso.
'....................................................................................................................
Private DeptoDir As String, LocalDir As String
Private EmpresaEmisora As clsClienteCFE

Private TasaBasica As Currency, TasaMinima As Currency

Private dFRetirar As Date
Private Const prmKeyApp = "FacturaContado"
Private Const keyTicketContado As String = "TicketContado"
Dim oCnfgPrint As New clsImpresoraTicketsCnfg
Dim oCnfgPrintSalidaCaja As New clsImpresoraTicketsCnfg
    
Private Const cte_KeyFindDir = "Buscar ......?"
Private oCliente As clsContacto

'String.----------------------------------------
Public strCodigoEnvio As String         'Cdo. vuelvo de envio si no graba tengo que borrar los que esten aca.

Private gDirFactura As Long
Private jobnum As Integer       'Nro. de Trabajo para la contado
Private CantForm As Integer    'Cantidad de formulas del reporte

Private lUltimoEnvio As Long            'Retiene el último envío almacenado.
Private lUltimoDoc As Long              'Retengo el último ctdo realizado (para ir a remitos)

Private Type tFechaRetirar
    IDArticulo As Long
    FRetira As Date
End Type
Private arrArtFechaRetira() As tFechaRetirar

Private miRenglon As tRenglonFact
Private m_Patron As String
Private Const vbRojoFuerte = &HC0&

Private Sub MenuAccionRetiraEnLocal()
    Dim lRet As String
    lRet = GetSetting("Ventas", "Config", "RetiraAqui", "0")
    If Val(lRet) = 1 Then
        MnuEntRetirar.Checked = True
    Else
        MnuEntRetirar.Checked = False
    End If
    SeteoCtrlsRetiraEnLocal
End Sub

Private Sub SeteoCtrlsRetiraEnLocal()
    If MnuEntRetirar.Checked Then
        tFRetiro.Left = 1080
    Else
        tFRetiro.Left = 1200
    End If
    chkRetiraAqui.Visible = MnuEntRetirar.Checked
End Sub

Private Sub ValidarFechaRetiro()
    If IsDate(tFRetiro.Text) Then
        Dim dFecha As Date: dFecha = ObtenerFechaRetiro
        If CDate(tFRetiro.Text) < dFecha Then
            MsgBox "Fecha incorrecta!!! se sustituirá.", vbExclamation, "Atención"
            tFRetiro.Text = dFecha
        End If
    End If
End Sub

Private Sub EliminarArrayFecha(ByVal Articulo As Long)
On Error GoTo errEAF
    Dim intA As Integer
    For intA = 1 To UBound(arrArtFechaRetira)
        If arrArtFechaRetira(intA).IDArticulo = Articulo Then
            arrArtFechaRetira(intA).FRetira = Date - 1
            arrArtFechaRetira(intA).IDArticulo = 0
            Exit Sub
        End If
    Next
Exit Sub
errEAF:
End Sub

Private Function ObtenerFechaRetiro() As Date
On Error GoTo errOFR
    ObtenerFechaRetiro = DateSerial(2000, 1, 1)
    Dim intR As Integer
    Dim intA As Integer
    With vsGrilla
        For intR = 1 To .Rows - 1
            If CInt(.Cell(flexcpText, intR, 0)) - CInt(.Cell(flexcpData, intR, 6)) > 0 Then
                For intA = 1 To UBound(arrArtFechaRetira)
                    If arrArtFechaRetira(intA).IDArticulo = CLng(.Cell(flexcpData, intR, 0)) And arrArtFechaRetira(intA).FRetira > ObtenerFechaRetiro Then
                        ObtenerFechaRetiro = arrArtFechaRetira(intA).FRetira
                    End If
                Next
            End If
        Next
    End With
    Exit Function
errOFR:
    clsGeneral.OcurrioError "Error al buscar la fecha disponible de los artículos, se pondrá por defecto la última encontrada.", Err.Description, "Buscar fecha disponible"
End Function

Private Sub loc_InsertArticuloEspecifico()
On Error GoTo errTC
    
'    If IsNumeric(tArticulo.Text) Then
'        Cons = " WHERE ArtCodigo = " & tArticulo.Text
'    Else
'        Cons = " WHERE ArtNombre LIKE '" & clsGeneral.Replace(Trim(tArticulo.Text), " ", "%") & "%'"
'    End If
'
'    'presento lista con los artículos específicos que esten para vender.
'    Cons = "Select AEsID 'Código', AEsNombre 'Articulo', AEsNroSerie 'Nro.Serie', SucAbreviacion 'Local' " & _
'            " From (Articulo INNER JOIN ArticuloEspecifico ON AEsArticulo = ArtID )Left Outer Join Sucursal On AEsLocal = SucCodigo " & _
'            Cons & _
'            " And AEsEstado = 1 And AEsDocumento Is Null"
    Screen.MousePointer = 11
    
    Dim iRet As Long, iVarP As Currency, bEnv As Boolean
    Dim objLista As New clsListadeAyuda
    iRet = objLista.ActivarAyuda(cBase, "EXEC prg_BuscarArticuloEspecifico '" + tArticulo.Text + "'", 5200, 3, "Ayuda")
    Me.Refresh
    If iRet = 0 Then
        
        Set objLista = Nothing
        Exit Sub
    
    Else

        If Ingresado(objLista.RetornoDatoSeleccionado(0)) Then
            Set objLista = Nothing
            InicializoVarRenglon
            MsgBox "El artículo seleccionado ya está ingresado.", vbExclamation, "ATENCIÓN"
            Screen.MousePointer = 0
            Exit Sub
        End If
        iRet = objLista.RetornoDatoSeleccionado(3)
        miRenglon.IDArticulo = objLista.RetornoDatoSeleccionado(0)
        miRenglon.Tipo = objLista.RetornoDatoSeleccionado(2)
        miRenglon.CodArticulo = objLista.RetornoDatoSeleccionado(1)
        miRenglon.EsInhabilitado = False
                
        If objLista.RetornoDatoSeleccionado(6) Then
            'ES COMBO.
            Cons = "Select PreID, PreArticulo, PreImporte From Presupuesto Where PreArtCombo = " & miRenglon.IDArticulo _
                & " And PreMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
            Dim rsC As rdoResultset
            Set rsC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not rsC.EOF Then
                miRenglon.IDCombo = rsC!PreID
                miRenglon.ArtCombo = rsC!PreArticulo
            End If
            rsC.Close
        End If
        
        If Not PrecioArticulo(miRenglon.IDArticulo, cMoneda.ItemData(cMoneda.ListIndex), miRenglon.Precio) Then
            MsgBox "El artículo seleccionado no posee precios ingresados para la moneda seleccionada.", vbInformation, "ATENCIÓN"
        Else
            miRenglon.PrecioOriginal = miRenglon.Precio
        End If
        Cons = Trim(objLista.RetornoDatoSeleccionado(4))
        If Val(objLista.RetornoDatoSeleccionado(5)) <> 0 Then iVarP = objLista.RetornoDatoSeleccionado(5)
        
        Dim idEspecifico As Long
        idEspecifico = objLista.RetornoDatoSeleccionado(3)
        
        'Inserto en la grilla.
        If MsgBox("¿El artículo va para envío?", vbQuestion + vbYesNo, "Enviar Artículo específico") = vbYes Then
            bEnv = True
        End If
        
        If miRenglon.ArtCombo > 0 Then
            tCantidad.Text = "1"
            If bEnv Then cEnvio.Text = "Si" Else cEnvio.Text = "No"
            InsertoArticulosCombo iVarP, idEspecifico
        Else
            CargoArticuloEnGrilla miRenglon.IDArticulo, miRenglon.Tipo, 1, Especifico, miRenglon.Precio + iVarP, Cons, "", miRenglon.Precio + iVarP, IIf(bEnv, "Si", "No"), miRenglon.EsInhabilitado, miRenglon.CodArticulo, iRet
        End If
        
        labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), FormatoMonedaP)
        MnuEmitir.Enabled = True
        If bEnv Then MnuEnvio.Enabled = True
        LimpioRenglon
        'Pongo por defecto artículo normal.
        labArticulo.Tag = "0": labArticulo.Caption = "&Artículo"
        tArticulo.SetFocus
        cMoneda.Enabled = False

    End If
    Set objLista = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
errTC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar artículos específicos.", Err.Description, "Artículo específico"
End Sub
Private Sub cDireccion_Change()
    If labDireccion.Caption <> "" And cDireccion.ListIndex = -1 Then labDireccion.Caption = ""
End Sub

Private Sub cDireccion_Click()
On Error GoTo errCargar

    If cDireccion.ListIndex <> -1 Then
        If Val(cDireccion.ItemData(cDireccion.ListIndex)) > -1 Then
            Screen.MousePointer = 11
            labDireccion.Caption = ""
            labDireccion.Caption = " " & clsGeneral.ArmoDireccionEnTexto(cBase, cDireccion.ItemData(cDireccion.ListIndex))
            Screen.MousePointer = 0
        Else
            labDireccion.Caption = ""
            cDireccion.SelStart = 0: cDireccion.SelLength = Len(cDireccion.Text)
        End If
    Else
        labDireccion.Caption = ""
    End If

errCargar:
    Screen.MousePointer = 0
End Sub

Private Sub cDireccion_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If cDireccion.ListIndex = -1 Then
            If Val(cDireccion.ItemData(cDireccion.ListIndex)) = -1 And cte_KeyFindDir <> cDireccion.Text Then
                loc_FindDireccionAuxiliarTexto
            End If
        Else
            chNomDireccion.SetFocus
        End If
    End If
End Sub

Private Sub cEnvio_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And cEnvio.ListIndex > -1 _
        And IsNumeric(tUnitario.Text) And (Val(miRenglon.IDArticulo) > 0 Or Val(tArticulo.Tag) > 0) _
        And Val(tCantidad.Text) >= 1 Then
        
        If CCur(tUnitario.Text) < 0 And CInt(labArticulo.Tag) <> 2 Then
            MsgBox "No se puede facturar artículos con costo negativo.", vbExclamation, "ATENCIÓN"
            tUnitario.SetFocus: Exit Sub
        End If
        
        If Trim(tComentario.Text) <> vbNullString Then
            If Not clsGeneral.TextoValido(tComentario.Text) Then
                MsgBox "Se ingreso un carácter no válido en el campo comentario.", vbExclamation, "ATENCIÓN"
                tComentario.SetFocus: Exit Sub
            End If
        End If
        
        If cDireccion.ListCount = 0 And oCliente.ID > 0 Then LabelMensaje labDireccion, True
        If Val(lTelefono.Tag) = 0 And oCliente.ID > 0 Then LabelMensaje lTelefono, True
        
        If labArticulo.Tag = "0" Then InsertoFila Else InsertoFilaCombo
        
        f_ValidateRetira False
        
    End If
End Sub

Private Sub chkRetiraAqui_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tVendedor.SetFocus
End Sub

Private Sub chPagaCheque_Click()
Dim bVisible As Boolean
    If chPagaCheque.Value = 0 Then
        CamposBanco False
    Else
        CamposBanco True
        If oCliente.Cheque Then CargoDatosCheque
    End If
End Sub

Private Sub chPagaCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComentarioDocumento
End Sub

Private Sub chRetira_Click()
    
    With tFRetiro
        .Locked = IIf(chRetira.Value = 1, False, True)
        .BackColor = IIf(chRetira.Value = 1, Obligatorio, Me.BackColor)
        .ForeColor = IIf(chRetira.Value = 1, vbBlack, vbRojoFuerte)
        .FontBold = IIf(chRetira.Value = 1, False, True)
        .BorderStyle = IIf(chRetira.Value = 1, 1, 0)
        .Text = IIf(chRetira.Value = 1, "", "NO")
    End With
    
End Sub

Private Sub chRetira_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And vsGrilla.Rows > 1 Then
        If Not tFRetiro.Locked Then
            Foco tFRetiro
        Else
            tVendedor.SetFocus
        End If
    End If
    
End Sub

Private Sub cMoneda_Click()
    LimpioRenglon
End Sub
Private Sub cMoneda_GotFocus()
    cMoneda.SelStart = 0
    cMoneda.SelLength = Len(cMoneda.Text)
    Status.Panels(1).Text = " Seleccione una moneda."
    Status.Panels(2).Enabled = False
End Sub
Private Sub cMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco txtCliente
End Sub
Private Sub cMoneda_LostFocus()
    cMoneda.SelLength = 0
End Sub

Private Sub cPendiente_GotFocus()
    cPendiente.SelStart = 0: cPendiente.SelLength = Len(cPendiente.Text)
    Status.Panels(1).Text = " Seleccione un motivo de pendiente."
End Sub
Private Sub cPendiente_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If chRetira.Enabled Then
            chRetira.SetFocus
        Else
            If Not tFRetiro.Locked Then
                Foco tFRetiro
            Else
                tVendedor.SetFocus
            End If
        End If
    End If
End Sub
Private Sub cPendiente_LostFocus()
    cPendiente.SelLength = 0
    Status.Panels(1).Text = vbNullString
End Sub

Private Sub chNomDireccion_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then Foco tArticulo
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault: Me.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errKD
    Select Case KeyCode
        Case vbKeyF2
            If Me.ActiveControl.Name <> "txtCliente" Then txtCliente.EditarCliente
        Case vbKeyF12
            Screen.MousePointer = 11
            EjecutarApp App.Path & "\visualizacion de operaciones.exe", CStr(oCliente.ID)
        Case vbKeyC
            If Shift = vbAltMask Then txtCliente.SetFocus
    End Select
    Screen.MousePointer = 0
    Exit Sub
errKD:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error inesperado.", Trim(Err.Description)
End Sub

Private Sub Form_Load()
On Error GoTo errInicializar
    
    If Not ValidarVersionEFactura Then
        MsgBox "La versión del componente CGSAEFactura está desactualizado, debe distribuir software." _
                    & vbCrLf & vbCrLf & "Se cancelará la ejecución.", vbCritical, "EFactura"
        End
    End If
    
    Set EmpresaEmisora = New clsClienteCFE
    'EmpresaEmisora.CargoInformacionCliente cBase, 1, False
    EmpresaEmisora.CargoClienteCarlosGutierrez paCodigoDeSucursal
    
    picTransaccion.BackColor = Me.BackColor
    picTransaccion.Visible = False
    
    Set txtCliente.Connect = cBase
    txtCliente.NeedCheckDigit = True
    
    MnuOptTransacciones.Visible = False
'    MnuOptTransacciones.Visible = miConexion.AccesoAlMenu("Cajero redpagos")
    MnuLineRP.Visible = MnuOptTransacciones.Visible
    
    lPrint.Caption = "Impresora: " & paIContadoN
    s_LoadMenuOpcionPrint
    
    oCnfgPrint.CargarConfiguracion prmKeyApp, keyTicketContado ' "CuotasImpresora"
    If Val(oCnfgPrint.ImpresoraTickets) = 0 Then
        MsgBox "INDIQUE EN DONDE IMPRIME LOS CONTADOS", vbExclamation, "ATENCIÓN"
        frmDondeImprimo.prmKeyTicket = keyTicketContado
        frmDondeImprimo.prmKeyApp = prmKeyApp
        frmDondeImprimo.Show vbModal
        oCnfgPrint.CargarConfiguracion prmKeyApp, keyTicketContado '"CuotasImpresora"
    End If
    
    oCnfgPrintSalidaCaja.CargarConfiguracion "MovimientosDeCaja", "TickeadoraMovimientosDeCaja"
        
    If Val(oCnfgPrint.ImpresoraTickets) = 0 Then
        MsgBox "Debe indicar la tickeadora a utilizar para imprimir los documentos.", vbExclamation, "ATENCIÓN"
        End
    End If
    
    Erase arrArtFechaRetira
    ReDim arrArtFechaRetira(0)
    
    Me.Height = 5685
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    dis_CargoArrayMonedas
    
    MenuAccionRetiraEnLocal

    With vsGrilla
        .Redraw = False: .Rows = 1: .Cols = 1
        .FormatString = ">Q|<Artículo|>Comentario|>Unitario|>I.V.A.|>SubTotal|<Envío|Habilitado|Instalador|codigo"
        .ColWidth(0) = 500: .ColWidth(1) = 3300: .ColWidth(2) = 1100: .ColWidth(3) = 1000: .ColWidth(4) = 700: .ColWidth(5) = 1000
        .Editable = False
        .ColHidden(7) = True
        .ColHidden(8) = True
        .ColHidden(9) = True
        .AllowUserResizing = flexResizeColumns: .Redraw = True
    End With
        
    labArticulo.Tag = "0": labArticulo.Caption = "&Artículo"
    
    tFRetiro.Text = ""
    chRetira.Value = 1
    chRetira.Enabled = False
    
    LimpioDatosCliente
    LimpioRenglon
    LimpioDatosBanco
    CamposBanco False
    strCodigoEnvio = ""
    LabTotalesEnCero
    
    FechaDelServidor
    '-----------------------------------------------------------------------------------------------------------
    cEnvio.AddItem "Si": cEnvio.AddItem "No"
    
    Cons = "Select MonCodigo, MonSigno From Moneda Where MonFactura = 1 Order by MonSigno"
    CargoCombo Cons, cMoneda, ""
    
    Cons = "Select PEnCodigo, PEnNombre From PendienteEntrega Order by PEnNombre"
    CargoCombo Cons, cPendiente, ""
    '-----------------------------------------------------------------------------------------------------------

    If paMonedaFacturacion > 0 Then BuscoCodigoEnCombo cMoneda, paMonedaFacturacion
    
'    InicializoCrystalEngine
    Exit Sub
    
errInicializar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario."
End Sub

'Private Sub InicializoCrystalEngine()
'
'    'Inicializa el Engine del Crystal y setea la impresora para el JOB
'    On Error GoTo ErrCrystal
'
'    'Abro el Engine del Crystal
'    If crAbroEngine = 0 Then GoTo ErrCrystal
'
'    'Inicializo el Reporte y SubReportes
'    jobnum = crAbroReporte(gPathListados & "Contado.RPT")
'    If jobnum = 0 Then GoTo ErrCrystal
'
'    'Configuro la Impresora
'    If Trim(Printer.DeviceName) <> Trim(paIContadoN) Then SeteoImpresoraPorDefecto paIContadoN
'    If Not crSeteoImpresora(jobnum, Printer, paIContadoB) Then GoTo ErrCrystal
'
'    'Obtengo la cantidad de formulas que tiene el reporte.
'    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
'    If CantForm = -1 Then GoTo ErrCrystal
'
'    Exit Sub
'
'ErrCrystal:
'    Screen.MousePointer = 0
'    crMsgErr = crMsgErr & IIf(crMsgErr <> "", vbCr, "")
'    clsGeneral.OcurrioError Trim(crMsgErr) & "No se podrán imprimir facturas." & vbCr & "Error: " & Err.Description, "Inicializo Crystal"
'End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Status.Panels(1).Text = vbNullString
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If strCodigoEnvio <> vbNullString And vsGrilla.Rows > 1 Then BorroEnvios
End Sub

Private Sub BorroEnvios()

On Error GoTo ErrBE
Dim CodEnvios As String
Dim lngCodEnvio As Long
    If strCodigoEnvio = "0" Then strCodigoEnvio = ""
    Do While strCodigoEnvio <> ""
    
        If InStr(1, strCodigoEnvio, ",") > 0 Then
            CodEnvios = Left(strCodigoEnvio, InStr(1, strCodigoEnvio, ","))
            lngCodEnvio = CLng(Left(CodEnvios, InStr(1, CodEnvios, ",") - 1))
            strCodigoEnvio = Right(strCodigoEnvio, Len(strCodigoEnvio) - InStr(1, strCodigoEnvio, ","))
        Else
            lngCodEnvio = CLng(strCodigoEnvio)
            strCodigoEnvio = ""
        End If
        
        cBase.BeginTrans
        On Error GoTo ErrResumo
        
        Dim idEVC As Long
        Cons = "SELECT EVCID, IsNull(count(*), 0) From EnvioVaCon WHERE EVCEnvio = " & lngCodEnvio & " GROUP BY EVCID"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If RsAux(1) > 2 Then
                cBase.Execute ("DELETE EnvioVaCon WHERE EVCID = " & RsAux(0) & " AND EVCEnvio = " & lngCodEnvio)
            Else
                cBase.Execute ("DELETE EnvioVaCon WHERE EVCID = " & RsAux(0))
            End If
        End If
        RsAux.Close
        
        'Borro los renglones del envío.
        Cons = "DELETE RenglonEnvio Where REvEnvio = " & lngCodEnvio
        cBase.Execute (Cons)
        
        'Borro el envío
        Cons = "Select EnvDireccion From Envio Where EnvCodigo = " & lngCodEnvio
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not IsNull(RsAux("EnvDireccion")) Then lngCodEnvio = RsAux!EnvDireccion Else lngCodEnvio = 0
        RsAux.Delete
        RsAux.Close
        
        'Borro la dirección.
        If lngCodEnvio > 0 Then
            Cons = "DELETE Direccion Where DirCodigo = " & lngCodEnvio
            cBase.Execute (Cons)
        End If
        
        cBase.CommitTrans
    Loop
    Exit Sub
    
ErrBE:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error inesperado al intentar la transacción."
    
ErrResumo:
    Resume Relajo
    
Relajo:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "No se pudo eliminar algún envío."

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    cBase.Close
    End
End Sub

Private Sub labArticulo_Click()
    Foco tArticulo
End Sub

Private Sub Label12_Click()
    Foco cMoneda
End Sub

Private Sub Label13_Click()
    Foco tUsuario
    Status.Panels(1).Text = " Ingrese el dígito de usuario."
End Sub

Private Sub Label16_Click()
    Foco tVendedor
End Sub
Private Sub Label17_Click()
    Foco tComentario
End Sub
Private Sub Label26_Click()
    Foco cPendiente
End Sub

Private Sub Label3_Click()
On Error Resume Next
    With tNombreC
        .SelStart = Len(.Text)
    End With
End Sub

Private Sub Label53_Click()
    Foco tComentarioDocumento
End Sub
Private Sub Label6_Click()
    Foco tCantidad
End Sub
Private Sub Label7_Click()
    Foco tUnitario
End Sub
Private Sub Label8_Click()
    cEnvio.SetFocus
End Sub

Private Sub lblInfoCliente_Click()
On Error Resume Next
    txtCliente.SetFocus
End Sub

Private Sub lPrint_Change()
    If paPrintEsXDef Then
        lPrint.ForeColor = vbBlack
    Else
        lPrint.ForeColor = &HFF&
    End If
End Sub

Private Sub lPrint_DblClick()
    frmEnQueVino.Show
End Sub

Private Sub MnuEmitir_Click()
    AccionEmitir
End Sub

Private Sub MnuEntRetirar_Click()
    Dim lVal As String
    lVal = IIf(MnuEntRetirar.Checked, "0", "1")
    SaveSetting "Ventas", "Config", "RetiraAqui", lVal
    MenuAccionRetiraEnLocal
End Sub

Private Sub MnuEnvio_Click()
    Dim bolVoy As Boolean
    bolVoy = False
    For I = 1 To vsGrilla.Rows - 1
        If Trim(vsGrilla.Cell(flexcpText, I, 6)) = "Si" Then bolVoy = True: Exit For
    Next
    If bolVoy Then
        If oCliente.ID = 0 Then
            MsgBox "No se pueden ingresar envíos sin seleccionar un cliente.", vbInformation, "ATENCIÓN": Exit Sub
        Else
            Dim idTabla As Integer
            idTabla = NumeroAuxiliarEnvio
            If idTabla = 0 Then MsgBox "Reintente la operación.", vbExclamation, "ATENCIÓN": Exit Sub
            Dim objEnvio As New clsEnvio
            objEnvio.NuevoEnvio cBase, strCodigoEnvio, idTabla, oCliente.ID, cMoneda.ItemData(cMoneda.ListIndex), TipoEnvio.Entrega
            Me.Refresh
            strCodigoEnvio = objEnvio.RetornoEnvios
            Set objEnvio = Nothing
            If strCodigoEnvio = vbNullString Then strCodigoEnvio = "0"
            CalculoArticulosEnEnvio
            If strCodigoEnvio <> "0" Then MnuEnvio.Enabled = True: CargoRenglonesEnvio (strCodigoEnvio)
            f_ValidateRetira False
        End If
    Else
        MsgBox "No hay artículos marcados para envío.", vbInformation, "ATENCIÓN"
    End If

End Sub

Private Sub MnuLimpiar_Click()

    If Trim(strCodigoEnvio) <> vbNullString Then BorroEnvios
    strCodigoEnvio = ""
    
    MnuOpFactSinIVA.Checked = False
    BackColor = &HC0E0FF
    Shape2.BackColor = &H80C0FF
    chPagaCheque.BackColor = BackColor
    chNomDireccion.BackColor = Shape2.BackColor
    
    Erase arrArtFechaRetira
    ReDim arrArtFechaRetira(0)
    
    LimpioRenglon
    vsGrilla.Rows = 1
    tVendedor.Text = vbNullString
    LimpioDatosCliente
    LabTotalesEnCero
    cMoneda.Enabled = True
    BuscoCodigoEnCombo cMoneda, paMonedaFacturacion
    txtCliente.Text = ""
    txtCliente.DocumentoCliente = DC_CI
    tUsuario.Text = vbNullString
    tUsuario.Tag = vbNullString
    cPendiente.ListIndex = -1
    tComentarioDocumento.Text = vbNullString
    MuestroPorServicio
    CamposBanco False
    LimpioDatosBanco
    
    dFRetirar = Date
    
    tFRetiro.Text = ""
    chRetira.Value = 1
    chRetira.Enabled = False
    chkRetiraAqui.Value = 0
    
    CamposRenglon True
    labArticulo.Tag = "0": labArticulo.Caption = "&Artículo"
    txtCliente.SetFocus
   
End Sub

Private Sub MnuOpFactSinIVA_Click()
    
    If vsGrilla.Rows > 1 Then
        MsgBox "Para cambiar la condición se deben eliminar los artículos ya ingresados.", vbInformation, "ATENCIÓN"
    Else
        MnuOpFactSinIVA.Checked = Not MnuOpFactSinIVA.Checked
        If MnuOpFactSinIVA.Checked Then
            BackColor = &H80C0FF
            Shape2.BackColor = &HC0E0FF
        Else
            BackColor = &HC0E0FF
            Shape2.BackColor = &H80C0FF
        End If
        chPagaCheque.BackColor = BackColor
        chNomDireccion.BackColor = Shape2.BackColor
    End If
    
End Sub

Private Sub MnuOpRemito_Click()
    If lUltimoDoc > 0 Then
        EjecutarApp App.Path & "\remitos.exe", "doc:" & CStr(lUltimoDoc)
    End If
End Sub

'Private Sub MnuOptTransacciones_Click()
'    MnuOptTransacciones.Checked = Not MnuOptTransacciones.Checked
'    picTransaccion.Visible = MnuOptTransacciones.Checked
'    Set oTransVta = Nothing
'End Sub

Private Sub MnuOpUltimoEnvio_Click()
    If lUltimoEnvio > 0 Then
        Dim objEnvio As New clsEnvio
        objEnvio.InvocoEnvio lUltimoEnvio, gPathListados
        Set objEnvio = Nothing
    End If
End Sub

Private Sub MnuPrintConfig_Click()
    prj_LoadConfigPrint True
    s_SetPrinter
End Sub

Private Sub MnuPrintDonde_Click()
    frmDondeImprimo.prmKeyTicket = keyTicketContado
    frmDondeImprimo.prmKeyApp = prmKeyApp
    frmDondeImprimo.Show vbModal
    oCnfgPrint.CargarConfiguracion prmKeyApp, keyTicketContado '"CuotasImpresora"
End Sub

Private Sub MnuPrintOpt_Click(Index As Integer)
On Error GoTo errLCP
Dim objPrint As New clsCnfgPrintDocument
Dim sPrint As String
Dim iQ As Integer
    
    With objPrint
        Set .Connect = cBase
        .Terminal = paCodigoDeTerminal
        If .ChangeConfigPorOpcion(MnuPrintOpt(Index).Caption) Then
            sPrint = .getDocumentoImpresora(Contado)
            paOptPrintSel = .GetOpcionActual
        End If
        sPrint = .getDocumentoImpresora(Contado)
    End With
    Set objPrint = Nothing
    
    prj_LoadConfigPrint False
    s_SetPrinter
    Exit Sub
    
errLCP:
    MsgBox "Error al setear los datos de configuración: " & Err.Description, vbExclamation, "ATENCIÓN"
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub tANombre_GotFocus()
On Error Resume Next
    With tANombre
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tANombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tUsuario_KeyPress vbKeyReturn
End Sub

Private Sub tArticulo_Change()
On Error Resume Next
    miRenglon.IDArticulo = 0
    tmArticuloLimitado.Enabled = False
    tArticulo.ForeColor = vbBlack
    tCantidad.ForeColor = vbBlack
End Sub

Private Sub tArticulo_GotFocus()
    Status.Panels(1).Text = " Ingrese un código. con [F1] cambia el tipo."
End Sub

Private Sub tArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errCA
    
    Select Case KeyCode
        Case vbKeyF1
            LimpioRenglon
            If labArticulo.Tag = "0" Then
                labArticulo.Tag = "1": labArticulo.Caption = "&Servicio"
                CamposRenglon False
            ElseIf labArticulo.Tag = "1" Then
                labArticulo.Tag = "2": labArticulo.Caption = "&Combo"
                CamposRenglon True
            ElseIf labArticulo.Tag = "2" Then
                labArticulo.Tag = "3": labArticulo.Caption = "&Artículo Específico"
                CamposRenglon True
            Else
                labArticulo.Tag = "0": labArticulo.Caption = "&Artículo"
                CamposRenglon True
            End If
        
        Case vbKeyF4
            'Busco servicios que tenga el cliente.
            If Shift = 0 Then BuscoServiciosCliente
    End Select
    Exit Sub
errCA:
    clsGeneral.OcurrioError "Ocurrio un error inesperado.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
Dim aIdSeleccionado As Long
    
    If KeyAscii = vbKeyReturn Then
    
        If cDireccion.ListCount = 0 And oCliente.ID > 0 Then LabelMensaje labDireccion, True
        If Val(lTelefono.Tag) = 0 And oCliente.ID > 0 Then LabelMensaje lTelefono, True
        
        If cMoneda.ListIndex = -1 Then
            MsgBox "Seleccione una moneda.", vbExclamation, "ATENCIÓN": Foco cMoneda: Exit Sub
        End If
        
        tComentario.Tag = ""
        If Trim(tArticulo.Text) <> vbNullString Then
            tUnitario.Tag = ""
            If labArticulo.Tag = "0" Then
                
                If miRenglon.IDArticulo = 0 Then CargoArticulosNormales
                If miRenglon.IDArticulo > 0 Then Foco tCantidad
                
            ElseIf labArticulo.Tag = "1" Then
                
                'Servicio
                If Not IsNumeric(tArticulo.Text) Or vsGrilla.Rows > 1 Then Exit Sub
                CargoDatosServicio tArticulo.Text
                
            ElseIf labArticulo.Tag = "2" Then
                'Combo
                If Not IsNumeric(tArticulo.Text) Then
                    Cons = "Select PreCodigo, Nombre = PreNombre From Presupuesto " _
                        & " Where PreNombre Like '" & clsGeneral.Replace(Trim(tArticulo.Text), " ", "%") & "%'" _
                        & " And PreHabilitado = 1"
                    aIdSeleccionado = InvocoListaAyuda(Cons, False)
                    If aIdSeleccionado = 0 Then Exit Sub
                Else
                    aIdSeleccionado = tArticulo.Text
                End If
                Cons = "Select * From Presupuesto " _
                    & " Where PreCodigo = " & aIdSeleccionado _
                    & " And PReMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
                'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If ObtenerResultSet(cBase, RsAux, Cons, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
                If RsAux.EOF Then
                    RsAux.Close
                    MsgBox "No existe un combo con el código ingresado.", vbInformation, "ATENCIÓN"
                    Screen.MousePointer = 0: Exit Sub
                End If
                If RsAux!PreHabilitado Then
                    If Not IsNull(RsAux!PreImporte) Then tUnitario.Tag = Trim(RsAux!PreImporte)
                    If Not Ingresado(RsAux!PreArticulo) Then
                        tArticulo.Text = Trim(RsAux!PreNombre): tArticulo.Tag = RsAux!PreID: Foco tCantidad
                    Else
                        Screen.MousePointer = 0
                        MsgBox "El artículo con el cual se factura el combo ya fue ingresado.", vbExclamation, "ATENCIÓN"
                    End If
                Else
                    Screen.MousePointer = 0
                    MsgBox "El presupuesto seleccionado no está habilitado para la venta.", vbExclamation, "ATENCIÓN"
                End If
                RsAux.Close
            ElseIf labArticulo.Tag = "3" Then
                'ESPECÍFICO
                loc_InsertArticuloEspecifico
            End If
        Else
            aIdSeleccionado = 0
            If vsGrilla.Rows > 1 Then
                For I = 1 To vsGrilla.Rows - 1
                    If Trim(vsGrilla.Cell(flexcpText, I, 6)) = "Si" Then
                        aIdSeleccionado = 1
                        Exit For
                    End If
                Next
                If aIdSeleccionado = 1 Then
                    MnuEnvio_Click
                Else
                    f_ValidateRetira True
                End If
                If Me.ActiveControl.Name = "tArticulo" Then
                    cPendiente.SetFocus
                End If
                If oCliente.ID = 0 And vsGrilla.Rows > vsGrilla.FixedRows Then txtCliente.SetFocus
                Screen.MousePointer = vbDefault
            End If
            
        End If
    End If

End Sub

Private Sub tBanco_Change()
    tBanco.Tag = ""
End Sub

Private Sub tBanco_GotFocus()
    tBanco.SelStart = 0
    tBanco.SelLength = 6
End Sub

Private Sub tBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Len(tBanco.Text) = 5 Then
        If BuscoBancoEmisor(tBanco.Text) Then Foco tANombre Else MsgBox "No existe un banco para el código ingresado.", vbExclamation, "ATENCIÓN"
    End If
End Sub

Private Sub tCantidad_GotFocus()
    With tCantidad
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.Panels(1).Text = " Ingrese la cantidad de artículos."
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsNumeric(tCantidad.Text) Then
        If cDireccion.ListCount = 0 And oCliente.ID > 0 Then LabelMensaje labDireccion, True
        If Val(lTelefono.Tag) = 0 And oCliente.ID > 0 Then LabelMensaje lTelefono, True
        If CLng(tCantidad.Text) > 0 And labArticulo.Tag = "0" Then
            Foco tComentario
        ElseIf CLng(tCantidad.Text) > 0 And labArticulo.Tag = "2" Then
            Foco tComentario
        End If
        
    End If
End Sub

Private Sub tCantidad_LostFocus()
    AplicoCantidadLimitadaPorCantidad
End Sub

Private Sub AplicoCantidadLimitadaPorCantidad()
    tmArticuloLimitado.Enabled = False
    If miRenglon.IDArticulo = 0 Then Exit Sub
    If IsNumeric(tCantidad.Text) Then
        If CInt(tCantidad.Text) > 0 Then
            tCantidad.Text = CInt(tCantidad.Text)
            tmArticuloLimitado.Enabled = (miRenglon.CantidadAlXMayor = 0 And InStr(1, paCategoriaDistribuidor, "," & oCliente.Categoria & ",") > 0) Or miRenglon.CantidadAlXMayor > 1 And miRenglon.CantidadAlXMayor < Val(tCantidad.Text)
        Else
            tCantidad.Text = ""
        End If
    End If
    AplicoTextoDeVentaLimitada
End Sub

Private Sub LimpioDatosCliente()

    tNombreC.Text = ""
    
    Set oCliente = New clsContacto
    
    lTelCant.Caption = "Tel.:"
    lTelefono.Caption = ""
    lTelefono.Tag = ""
    labDireccion.Caption = ""
    chNomDireccion.Value = 0
    LabelMensaje labDireccion, False
    LabelMensaje lTelefono, False
    cDireccion.Clear: cDireccion.BackColor = Colores.Gris
    gDirFactura = 0
    lblRucPersona.Caption = ""
    
End Sub

Private Sub tComentario_GotFocus()
    With tComentario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.Panels(1).Text = " Ingrese un comentario para el artículo."
End Sub

Private Sub tComentario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 And Shift = 0 And miRenglon.IDArticulo > 0 And labArticulo.Tag = "0" Then
        loc_InsertArticuloEspecifico
    End If
End Sub

Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        tUnitario.SetFocus
        If cDireccion.ListCount = 0 And oCliente.ID > 0 Then LabelMensaje labDireccion, True
        If Val(lTelefono.Tag) = 0 And oCliente.ID > 0 Then LabelMensaje lTelefono, True
    End If
End Sub
Private Sub tComentarioDocumento_GotFocus()
    tComentarioDocumento.SelStart = 0
    tComentarioDocumento.SelLength = Len(tComentarioDocumento.Text)
    Status.Panels(1).Text = " Ingrese un comentario para el documento."
End Sub
Private Sub tComentarioDocumento_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then tUsuario.SetFocus
End Sub
Private Sub tComentarioDocumento_LostFocus()
    Status.Panels(1).Text = vbNullString
End Sub

Private Sub tFRetiro_Change()
    tFRetiro.Tag = ""
End Sub

Private Sub tFRetiro_GotFocus()
    If chRetira.Value = 0 Then tVendedor.SetFocus
    
    With tFRetiro
        If Trim(.Text) = "" Then
            Dim dRet As Date: dRet = ObtenerFechaRetiro
            If dRet < Date Then dRet = Date
            .Text = Format(dRet, "d-Mmm yyyy"): .Tag = "D"
        End If
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.Panels(1).Text = " Ingrese la fecha posible de retiro de la mercadería."
    
End Sub
Private Sub tFRetiro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If chRetira.Value = 1 And IsDate(tFRetiro.Text) Then
            On Error Resume Next
            If chRetira.Value = 1 And chkRetiraAqui.Enabled And chkRetiraAqui.Visible Then
                chkRetiraAqui.SetFocus
            Else
                tVendedor.SetFocus
            End If
        End If
    End If
End Sub
Private Sub tFRetiro_LostFocus()
     If IsDate(tFRetiro.Text) Then tFRetiro.Text = Format(tFRetiro.Text, "d-Mmm yyyy")
     Status.Panels(1).Text = vbNullString
End Sub

Private Sub tFRetiro_Validate(Cancel As Boolean)
    ValidarFechaRetiro
End Sub

Private Sub tmArticuloLimitado_Timer()
    tmArticuloLimitado.Enabled = False
    If Val(tmArticuloLimitado.Tag) = 0 Then
        tArticulo.ForeColor = &HFF&
        tmArticuloLimitado.Tag = 1
    Else
        tArticulo.ForeColor = vbBlack
        tmArticuloLimitado.Tag = 0
    End If
    tCantidad.ForeColor = tArticulo.ForeColor
    tmArticuloLimitado.Enabled = True
End Sub

Private Sub tNombreC_GotFocus()
On Error Resume Next
    tNombreC.SelStart = Len(tNombreC.Text)
End Sub

Private Sub tNombreC_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then Foco tArticulo
End Sub

Private Function CargoDatosCliente() As Boolean
    'ATENCION-------------------------------------------------------------------
    'En el tag del nombre me guardo el tipo de Cliente.
    'En el tag de dirección guardo la categoria de descuento del cliente.
    '---------------------------------------------------------------------------
On Error GoTo errCDC
Dim rsC As rdoResultset
    
    Cons = "Select CliDireccion, CPeApellido1, CPeApellido2, CPeNombre1, CPeNombre2, CliCategoria, CliCheque, CEmNombre, CEmFantasia From Cliente " _
            & " Left Outer Join CPersona On CPeCliente = CliCodigo" _
            & " Left Outer Join CEmpresa On CEmCliente = CliCodigo" _
        & " Where CliCodigo = " & txtCliente.Cliente.Codigo
    'Set rsC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, rsC, Cons, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
    
    oCliente.ID = txtCliente.Cliente.Codigo
    oCliente.Tipo = txtCliente.Cliente.Tipo
    If Not IsNull(rsC!CliCategoria) Then oCliente.Categoria = rsC("CliCategoria")
    
    If txtCliente.Cliente.Tipo = TC_Persona Then
        oCliente.Apellido1 = Trim(rsC("CPeApellido1"))
        If Not IsNull(rsC("CPeApellido2")) Then oCliente.Apellido2 = Trim(rsC("CPeApellido2"))
        oCliente.Nombre1 = Trim(rsC("CPeNombre1"))
        If Not IsNull(rsC("CPeNombre2")) Then oCliente.Nombre2 = Trim(rsC("CPeNombre2"))
        tNombreC.Text = oCliente.MostrarComo
        lblRucPersona.Caption = txtCliente.Cliente.RutPersona
    Else
        If Not IsNull(rsC!CEmNombre) Then
            oCliente.Nombre1 = Trim(rsC!CEmNombre)
        Else
            oCliente.Apellido1 = Trim(rsC!CEmFantasia)
        End If
    End If
    tNombreC.Text = " " & oCliente.MostrarComo
    
    If Not IsNull(rsC!CliDireccion) Then
        cDireccion.AddItem "Dirección Principal": cDireccion.ItemData(cDireccion.NewIndex) = rsC!CliDireccion
        cDireccion.Tag = rsC!CliDireccion
        gDirFactura = rsC!CliDireccion
    End If
    'bCheque = False
    If Not IsNull(rsC!CliCheque) Then
        oCliente.Cheque = (UCase(rsC!CliCheque) = "S")   'bCheque = True
    End If
    rsC.Close
    CargoDireccionesAuxiliares oCliente.ID
    CargoTelefonos
    CargoDatosCliente = (oCliente.ID > 0)
    Exit Function
errCDC:
    clsGeneral.OcurrioError "Error al cargar la ficha del cliente.", Err.Description
End Function

Private Sub tUnitario_GotFocus()
On Error Resume Next
    
    If miRenglon.IDArticulo > 0 Then PresentoPrecio
    
    Status.Panels(1).Text = " Costo unitario del artículo."
    tUnitario.SelStart = 0
    tUnitario.SelLength = Len(tUnitario.Text)
    
End Sub

Private Sub tUnitario_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And IsNumeric(tUnitario.Text) And cMoneda.ListIndex > -1 Then
        
        If cDireccion.ListCount = 0 And oCliente.ID > 0 Then LabelMensaje labDireccion, True
        If Val(lTelefono.Tag) = 0 And oCliente.ID > 0 Then LabelMensaje lTelefono, True
        
        If Val(tArticulo.Tag) = 0 And miRenglon.IDArticulo = 0 Then
            MsgBox "No hay un artículo seleccionado.", vbExclamation, "ATENCIÓN": LimpioRenglon: Exit Sub
        End If
        
        If IsNumeric(tUnitario.Text) Then
            m_Patron = dis_arrMonedaProp(cMoneda.ItemData(cMoneda.ListIndex), pRedondeo)
            tUnitario.Text = Redondeo(CCur(tUnitario.Text), m_Patron)
            cEnvio.ListIndex = 1: cEnvio.SetFocus
        End If
        
    Else
        If cMoneda.ListIndex = -1 Then
            MsgBox "No seleccionó una moneda.", vbCritical, "ATENCIÓN"
            LimpioRenglon
            cMoneda.SetFocus: vsGrilla.Rows = 1
        End If
    End If
        
End Sub

Private Sub LimpioRenglon()
    tArticulo.ForeColor = vbBlack
    tCantidad.ForeColor = vbBlack
    tArticulo.Text = "": tArticulo.Tag = "0"
    tCantidad.Text = ""
    tComentario.Text = ""
    tUnitario.Text = "": tUnitario.Tag = ""
    cEnvio.Tag = "": cEnvio.Text = ""
    tComentario.Tag = ""
    tmArticuloLimitado.Enabled = False
    InicializoVarRenglon
End Sub

Private Function BuscoDescuentoCliente(IDArticulo As Long, CatCliente As Long, curUnitario As Currency, intCantidad As Integer) As String
Dim RsDto As rdoResultset
    BuscoDescuentoCliente = curUnitario
    miRenglon.Precio = curUnitario
    
    If paTipoCuotaContado > 0 And CatCliente > 0 Then
        
        m_Patron = dis_arrMonedaProp(cMoneda.ItemData(cMoneda.ListIndex), pRedondeo)
    
        Cons = "Select CDTPorcentaje, AFaCantidadD From ArticuloFacturacion, CategoriaDescuento" _
            & " Where AfaArticulo = " & IDArticulo _
            & " And AfaCategoriaD = CDtCatArticulo And CDtCatCliente = " & CatCliente _
            & " And CDtCatPlazo = " & paTipoCuotaContado
            
        'Set RsDto = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If ObtenerResultSet(cBase, RsDto, Cons, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
        
        If Not RsDto.EOF Then
            If Not IsNull(RsDto!AFaCantidadD) Then
                If RsDto!AFaCantidadD <= intCantidad Then
                    BuscoDescuentoCliente = Redondeo(curUnitario - (curUnitario * RsDto(0)) / 100, m_Patron)
                    miRenglon.Precio = CCur(BuscoDescuentoCliente)
                Else
                    miRenglon.Precio = curUnitario
                    If MsgBox("El cliente tiene descuento para el artículo pero, no cumple con la cantidad mínima." & Chr(13) _
                        & "¿Le aplica el descuento de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
                        BuscoDescuentoCliente = Redondeo(curUnitario - (curUnitario * RsDto(0)) / 100, m_Patron)
                        miRenglon.Precio = CCur(BuscoDescuentoCliente)
                    End If
                End If
            End If
        End If
        RsDto.Close
    End If
    BuscoDescuentoCliente = Format(BuscoDescuentoCliente, FormatoMonedaP)

End Function

Private Sub LabTotalesEnCero()
    labTotal.Caption = "0.00": labIVA.Caption = "0.00": labSubTotal.Caption = "0.00"
End Sub

Private Function Ingresado(IDArticulo As Long)
    
    Ingresado = False
    With vsGrilla
        For I = 1 To .Rows - 1
            If IDArticulo = CLng(.Cell(flexcpData, I, 0)) Then Ingresado = True: Exit For
        Next I
    End With

End Function

Private Sub RestoLabTotales(curTotal As Currency, curIva As Currency)
Dim cIVA As Currency
    cIVA = Format(curTotal - (curTotal / CCur(1 + (curIva / 100))), "#,##0.00")
    labIVA.Caption = Format(CCur(labIVA.Caption) - cIVA, "#,##0.00")
    labTotal.Caption = Format(CCur(labTotal.Caption) - curTotal, "#,##0.00")
    labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
End Sub

Private Sub CambioPreciosEnLista(CodCategoria As Long)
On Error GoTo ErrCPEL
Dim aValor As Currency
    
    'ATENCION.---------------------------------------------------------------
    'Si CodCategoria = 0 then No se aplica descuento.----------
    'Pongo los labels de totales en cero.---------------------------------
    LabTotalesEnCero
    For I = 1 To vsGrilla.Rows - 1
        
        If Trim(vsGrilla.Cell(flexcpData, I, 2)) = "0" Then  'Verifico que el artículo no sea de un combo o de servicio.
            
            Cons = "Select PViPrecio From PrecioVigente" _
                & " Where PViArticulo = " & CLng(vsGrilla.Cell(flexcpData, I, 0)) _
                & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                & " And PViHabilitado = 1 And PViTipoCuota = " & paTipoCuotaContado
                
            'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If ObtenerResultSet(cBase, RsAux, Cons, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
            If RsAux.EOF Then
                vsGrilla.Cell(flexcpData, I, 4) = ""   'No hay precio para ese artículo.
            Else
                aValor = RsAux!PViPrecio
                If MnuOpFactSinIVA.Checked Then
                    m_Patron = dis_arrMonedaProp(cMoneda.ItemData(cMoneda.ListIndex), pRedondeo)
                    aValor = Redondeo(aValor / (1 + (IVAArticulo(CLng(vsGrilla.Cell(flexcpData, I, 0))) / 100)), m_Patron)
                End If
                If CodCategoria = 0 Then
                    vsGrilla.Cell(flexcpData, I, 4) = Redondeo(aValor, m_Patron)
                    vsGrilla.Cell(flexcpText, I, 3) = Format(Redondeo(aValor, m_Patron), FormatoMonedaP)
                    vsGrilla.Cell(flexcpText, I, 5) = Format(vsGrilla.Cell(flexcpText, I, 0) * aValor, FormatoMonedaP)
                Else
                    vsGrilla.Cell(flexcpText, I, 3) = Redondeo(BuscoDescuentoCliente(vsGrilla.Cell(flexcpData, I, 0), CodCategoria, aValor, vsGrilla.Cell(flexcpText, I, 0)), m_Patron)
                    aValor = vsGrilla.Cell(flexcpText, I, 3): vsGrilla.Cell(flexcpData, I, 4) = aValor
                    vsGrilla.Cell(flexcpText, I, 5) = Format(vsGrilla.Cell(flexcpText, I, 0) * vsGrilla.Cell(flexcpText, I, 3), FormatoMonedaP)
                End If
            End If
            labIVA.Caption = Format(CCur(labIVA.Caption) + CCur(vsGrilla.Cell(flexcpText, I, 5)) - (CCur(vsGrilla.Cell(flexcpText, I, 5)) / CCur(1 + (CCur(vsGrilla.Cell(flexcpText, I, 4)) / 100))), FormatoMonedaP)
            labTotal.Caption = Format(CCur(labTotal.Caption) + CCur(vsGrilla.Cell(flexcpText, I, 5)), FormatoMonedaP)
            labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), FormatoMonedaP)
            RsAux.Close
        Else
            labIVA.Caption = Format(CCur(labIVA.Caption) + CCur(vsGrilla.Cell(flexcpText, I, 5)) - (CCur(vsGrilla.Cell(flexcpText, I, 5)) / CCur(1 + (CCur(vsGrilla.Cell(flexcpText, I, 4)) / 100))), FormatoMonedaP)
            labTotal.Caption = Format(CCur(labTotal.Caption) + CCur(vsGrilla.Cell(flexcpText, I, 5)), FormatoMonedaP)
            labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), FormatoMonedaP)
        End If
    Next I
    Exit Sub

ErrCPEL:
    clsGeneral.OcurrioError "Ocurrio un error inesperado al modificar los precios, VERIFIQUE."
    Screen.MousePointer = 0
End Sub

Private Sub tUnitario_LostFocus()
    If IsNumeric(tUnitario.Text) Then tUnitario.Text = Format(tUnitario.Text, FormatoMonedaP)
    AplicoCantidadLimitadaPorCantidad
End Sub

Private Sub tUsuario_GotFocus()
    With tUsuario
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.Panels(1).Text = " Ingrese el dígito de usuario."
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsNumeric(tUsuario.Text) Then
        tUsuario.Tag = 0: tUsuario.Tag = BuscoUsuarioDigito(Val(tUsuario.Text), True)
        If Val(tUsuario.Tag) > 0 Then
            If lBanco.Visible Then
                If Trim(tBanco.Tag) <> "" And tBanco.Text <> "" And tANombre.Text <> "" Then
                    tUsuario.Text = vbNullString
                    AccionEmitir
                Else
                    tBanco.SetFocus
                End If
            Else
                tUsuario.Text = vbNullString
                AccionEmitir
            End If
        End If
    End If
End Sub

Private Sub AccionEmitir()
    
    If cDireccion.ListIndex > -1 Then
        Cons = "SELECT DepNombre, LocNombre " & _
            "FROM Direccion INNER JOIN Calle ON DirCalle = CalCodigo " & _
            "INNER JOIN Localidad ON CalLocalidad = LocCodigo " & _
            "INNER JOIN Departamento ON LocDepartamento = DepCodigo " & _
            "WHERE DirCodigo = " & cDireccion.ItemData(cDireccion.ListIndex)
        'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If ObtenerResultSet(cBase, RsAux, Cons, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
        If Not RsAux.EOF Then
            DeptoDir = Trim(RsAux("DepNombre"))
            LocalDir = Trim(RsAux("LocNombre"))
        End If
        RsAux.Close
    End If
    
    Dim bNoVaRUT As Boolean
    If txtCliente.Cliente.Tipo = TC_Persona And txtCliente.Cliente.RutPersona <> "" Then
        Dim rsp As VbMsgBoxResult
        rsp = vbCancel
        Do While rsp = vbCancel
            rsp = MsgBox("CLIENTE UNIPERSONAL" & vbCrLf & vbCrLf & "¿El cliente desea facturar con RUT?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "FACTURAR CON RUT")
        Loop
        If rsp = vbNo Then
            bNoVaRUT = True
            rsp = vbCancel
            Do While rsp = vbCancel
                rsp = MsgBox("¿El cliente aún posee ese RUT?" & vbCrLf & vbCrLf & "Si responde NO el RUT se eliminará de la ficha del cliente.", vbQuestion + vbYesNoCancel + vbDefaultButton3, "RUT EN USO")
            Loop
            If rsp = vbNo Then
                
                'Updateo la tabla CPERSONA y registro el suceso de cambio de RUT.
                Cons = "UPDATE CPersona SET CPERuc = NULL WHERE CPeCliente = " & txtCliente.Cliente.Codigo
                cBase.Execute Cons
                
                lblRucPersona.Caption = ""
            End If
        End If
    End If
    
    tUsuario.Enabled = False
    If ControloDatos(bNoVaRUT) Then
        
        ControloVentaTelefonica
        'Controlo importe cédula
        If txtCliente.Cliente.Tipo = TC_Persona And oCliente.ID > 0 And CCur(labSubTotal.Caption) > Format(paImpCedula, "#,##0.00") And (txtCliente.Cliente.Documento = "" Or txtCliente.Cliente.Documento = "99999999") And lblRucPersona.Caption = "" Then
            MsgBox "Se debe ingresar el nombre y la cédula correcta del cliente." & vbCrLf & vbCrLf & "Resolución 873/005 de la DGI", vbExclamation, "Atención"
            tUsuario.Enabled = True
            Exit Sub
        End If
        
'        If txtCliente.Cliente.Tipo = TC_Persona And txtCliente.Cliente.RutPersona <> "" Then
'            If MsgBox("¿Consultó con el cliente si desea facturar con RUT?", vbQuestion + vbYesNo + vbDefaultButton2, "Persona factura con RUT") = vbNo Then
'                txtCliente.SetFocus
'                Exit Sub
'            End If
'        End If
        
        If ChangeCnfgPrint Then
            prj_LoadConfigPrint False
            s_SetPrinter
        End If
        
        snd_ActivarSonido Replace(gPathListados, "\reportes\", "\sonidos\", , , vbTextCompare) & "emitirfactura.wav"
        
        Dim sImpresora As String
'        If (oCnfgPrint.Opcion = 0 Or ((txtCliente.Cliente.Documento <> "" And txtCliente.Cliente.Tipo = TC_Empresa) Or (txtCliente.Cliente.RutPersona <> "" And txtCliente.Cliente.Tipo = TC_Persona))) Then
'            sImpresora = paIContadoN
'        Else
        sImpresora = "Tickeadora " & oCnfgPrint.ImpresoraTickets
'        End If
        
        'Valido y el usuario decide cancelar por la condición.
        If Not ControloFacturaDuplicada Then tUsuario.Enabled = True: Exit Sub
        
        Dim bVTaLimitada As Boolean
        bVTaLimitada = AvisoVentaLimitada(True)
        
        If Not ValidoRUT() Then
            txtCliente.SetFocus
            tUsuario.Enabled = True
            Exit Sub
        End If
        
        If Val(oCnfgPrint.ImpresoraTickets) = 0 Then
            MsgBox "Debe indicar la tickeadora a utilizar para imprimir los documentos.", vbExclamation, "ATENCIÓN"
        End If
        
        Dim resp As Byte
        Dim subResp As Byte
        resp = 255
        If RetiraEnDeposito And paCodigoDeSucursal = 5 Then
            Dim fQVino As New frmEnQueVino
            fQVino.Show vbModal
            resp = fQVino.Respuesta
            subResp = fQVino.SubRespuesta
        End If
        
        If MsgBox("¿ Confirma emitir la factura ?" & vbCrLf & vbCrLf & "Impresora: " & sImpresora, vbQuestion + vbYesNo + IIf(bVTaLimitada, vbDefaultButton2, vbDefaultButton1), "ATENCIÓN") = vbYes Then
            If ControlStock Then GraboFacturaRenglon bVTaLimitada, bNoVaRUT, resp, subResp
        Else
            tUsuario.Enabled = True: Foco tUsuario
        End If
        
    End If
    tUsuario.Enabled = True
End Sub

Private Function EsTipoDeServicio(ByVal idTipo As Long) As Boolean
     EsTipoDeServicio = (InStr(1, "," & tTiposArtsServicio & ",", "," & idTipo & ",") > 0)
End Function

Function ControloFacturaDuplicada() As Boolean
Dim sIDsArts As String
Dim iSumaIDCant As Long, iSumaID As Long
Dim iQTot As Long
    
    With vsGrilla
        For I = 1 To .Rows - 1
            'If CLng(.Cell(flexcpData, I, 3)) <> paTipoArticuloServicio Then
            If Not EsTipoDeServicio(CLng(.Cell(flexcpData, I, 3))) Then
                iSumaIDCant = iSumaIDCant + (CLng(.Cell(flexcpData, I, 0)) * CLng(.Cell(flexcpText, I, 0)))
                sIDsArts = sIDsArts & IIf(sIDsArts <> "", ", ", "") & .Cell(flexcpData, I, 0)
                iQTot = iQTot + CInt(.Cell(flexcpText, I, 0))
                iSumaID = iSumaID + CInt(.Cell(flexcpData, I, 0))
            End If
        Next
    End With
    
'AND RenArticulo IN (" & sIDsArts & ")"
    Cons = "SELECT DocCodigo" & _
        " FROM Documento INNER JOIN Renglon ON DocCodigo = RenDocumento" & _
        " INNER JOIN Articulo ON ArtID = RenArticulo AND ArtTipo <> 151" & _
        " WHERE DocCliente = " & oCliente.ID & _
        " AND DocTipo = 1 AND DocAnulado = 0 AND DocFecha >= DATEADD(Minute, -10, GetDate())" & _
        " Group by DocCodigo" & _
        " Having Sum(RenArticulo * RenCantidad) = " & iSumaIDCant & _
        " AND Sum(RenCantidad) = " & iQTot & " AND Sum(RenArticulo) = " & iSumaID
    
    Dim rsDup As rdoResultset
    'Set rsDup = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, rsDup, Cons, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
    If Not rsDup.EOF Then
        If MsgBox("ATENCIÓN!!! no hace 10 mínutos el cliente hizo una factura con los mismos artículos." & vbCrLf & vbCrLf & "¿Está seguro de grabar un documento identico al anterior?", vbQuestion + vbYesNo + vbDefaultButton2, "Documento identico") = vbNo Then
            rsDup.Close
            ControloFacturaDuplicada = False
            Exit Function
        End If
    End If
    rsDup.Close
    ControloFacturaDuplicada = True
    
End Function

Private Function ControloDatos(ByVal NoVaRUT As Boolean) As Boolean
Dim Suma As Currency
    
    ControloDatos = False
    If oCliente.ID = 0 Then
        MsgBox "No se puede facturar sin seleccionar el cliente.", vbExclamation, "ATENCIÓN"
        Foco txtCliente: Exit Function
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Se debe seleccionar una moneda.", vbExclamation, "ATENCIÓN"
        cMoneda.Enabled = True: Foco cMoneda: Exit Function
    End If
    
    If vsGrilla.Rows = 1 Then
        MsgBox "Debe ingresar los artículos a facturar.", vbExclamation, "ATENCIÓN"
        Foco tArticulo: Exit Function
    End If
    
    Suma = 0
    With vsGrilla
        For I = 1 To .Rows - 1
            Suma = Suma + CCur(.Cell(flexcpText, I, 5))
        Next I
    End With
    
    If Suma <> CCur(labTotal.Caption) Then
        MsgBox "La suma total no coincide con la suma de la lista, verifique.", vbCritical, "ATENCIÓN"
        Foco tArticulo: Exit Function
    End If
    
    If Not IsDate(tFRetiro.Text) And UCase(tFRetiro.Text) <> "NO" Then
        MsgBox "La fecha de retiro no es válida.", vbExclamation, "ATENCIÓN"
        Foco tFRetiro: Exit Function
    ElseIf UCase(tFRetiro.Text) <> "NO" Then
        If CDate(tFRetiro.Text) Then
            ValidarFechaRetiro
        End If
    End If
    
    If tVendedor.Tag = "" Or Val(tVendedor.Tag) = 0 Then
        MsgBox "No se ingresó el dígito del vendedor.", vbExclamation, "ATENCIÓN"
        Foco tVendedor: Exit Function
    End If
    
    If Not clsGeneral.TextoValido(tComentarioDocumento.Text) Then
        MsgBox "Se ingreso un carácter no válido en el comentario del documento.", vbExclamation, "ATENCIÓN"
        Foco tComentario: Exit Function
    End If
    
    If tUsuario.Tag = vbNullString Or Val(tUsuario.Tag) = 0 Then
        MsgBox "Debe ingresar el dígito de usuario que factura.", vbExclamation, "ATENCIÓN"
        tUsuario.SetFocus: Exit Function
    End If
    
    If chPagaCheque.Value = 1 Then
        If tBanco.Tag = "" Then
            MsgBox "Debe ingresar el banco del cheque.", vbExclamation, "ATENCIÓN"
            tBanco.SetFocus: Exit Function
        End If
        
        If Trim(tANombre.Text) = "" Then
            MsgBox "Debe ingresar el nombre de quien emite el cheque.", vbExclamation, "ATENCIÓN"
            tANombre.SetFocus: Exit Function
        Else
            If Not clsGeneral.TextoValido(tANombre.Text) Then
                MsgBox "Ingreso alguna comilla simple, verifique.", vbExclamation, "ATENCIÓN"
                tANombre.SetFocus: Exit Function
            End If
        End If
        
        If Not BuscoBancoEmisor(tBanco.Text) Then
            MsgBox "No existe un banco para el código ingresado.", vbExclamation, "ATENCIÓN"
            Exit Function
        End If
    End If
    
    'Controles eFactura.
    If (Suma > prmImporteConInfoCliente) Then
        
        If (txtCliente.Cliente.Tipo = TC_Empresa And txtCliente.Cliente.Documento = "") Or (txtCliente.Cliente.Tipo = TC_Persona And (txtCliente.Cliente.Documento = "" And txtCliente.Cliente.RutPersona = "")) Then
            MsgBox "Es necesario facturar con RUT o Cédula.", vbCritical, "EFactura"
            Exit Function
        End If
        
        If (cDireccion.ListIndex = -1) Or (labDireccion.Caption = "" Or Trim(labDireccion.Caption) = "Sin Dirección") Then
            MsgBox "Debe seleccionar una dirección de facturación.", vbExclamation, "Validación EFactura"
            Exit Function
        End If
        
    Else
        
        If (txtCliente.Cliente.Tipo = TC_Empresa And txtCliente.Cliente.Documento = "") Then  'Or (txtCliente.Cliente.Tipo = TC_Persona And (Suma > prmImporteConInfoCliente Or txtCliente.Cliente.RutPersona <> "")) Then
            If MsgBox("Para facturar a una empresa es necesario facturar con RUT." & vbCrLf & vbCrLf & "¿Desea facturar de todas formas?", vbQuestion + vbYesNo, "EFactura") = vbNo Then
                Exit Function
            End If
        ElseIf (txtCliente.Cliente.Tipo = TC_Empresa Or (txtCliente.Cliente.Tipo = TC_Persona And txtCliente.Cliente.RutPersona <> "" And Not NoVaRUT)) Then
            If (cDireccion.ListIndex = -1) Or (labDireccion.Caption = "" Or Trim(labDireccion.Caption) = "Sin Dirección") Then
                MsgBox "Debe seleccionar una dirección de facturación.", vbExclamation, "Validación EFactura"
                Exit Function
            End If
        End If
        
    End If
    
'    If (txtCliente.Cliente.Tipo = TC_Empresa) Or (txtCliente.Cliente.Tipo = TC_Persona And (Suma > prmImporteConInfoCliente Or txtCliente.Cliente.RutPersona <> "")) Then
'
'        If (cDireccion.ListIndex = -1) Or (labDireccion.Caption = "" Or Trim(labDireccion.Caption) = "Sin Dirección") Then
'            MsgBox "Debe seleccionar una dirección de facturación.", vbExclamation, "Validación EFactura"
'            Exit Function
'        End If
'
'        If (txtCliente.Cliente.Documento = "" And txtCliente.Cliente.RutPersona = "") Then
'            MsgBox "Para facturar es necesario ingresar cédula o RUT." & vbCrLf & "¿Desea continuar de todas formas?", vbExclamation, "ATENCIÓN"
'            Exit Function
'        End If
'    End If
    
    ControloDatos = True
    
    
    'Si paso la Validación controlo la direccion que factura-----------------------------------------------------------------
    If cDireccion.ListIndex <> -1 Then
        On Error Resume Next
        If gDirFactura <> cDireccion.ItemData(cDireccion.ListIndex) Then        'Cambio Dir Facutua
            If MsgBox("Ud. a cambiado la dirección con la que el cliente factura habitualmente." & vbCrLf & "Quiere que esta dirección quede por defecto para facturar.", vbQuestion + vbYesNo, "Dirección por Defecto al Facturar") = vbNo Then Exit Function
            
            If cDireccion.ItemData(cDireccion.ListIndex) <> Val(cDireccion.Tag) Then        'Dir. selecc. <> a la Ppal.
                
                'Agregué esto para limpiar ya que hay casos en que los clientes tienen más de uno.
                Cons = "Update DireccionAuxiliar SET DAuFactura = 0 WHERE DauCliente = " & oCliente.ID & " AND DauFactura = 1"
                cBase.Execute (Cons)
                
                Dim rsAD As rdoResultset
                Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & oCliente.ID & " And DAuDireccion = " & cDireccion.ItemData(cDireccion.ListIndex)
                Set rsAD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsAD.EOF Then
                    rsAD.Edit: rsAD!DAuFactura = True: rsAD.Update
                End If
                rsAD.Close
            End If
            
            If gDirFactura <> Val(cDireccion.Tag) Then      'La gDirFactura Anterior no era la ppal, la desmarco
                Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & oCliente.ID & " And DAuDireccion = " & gDirFactura
                Set rsAD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsAD.EOF Then
                    rsAD.Edit: rsAD!DAuFactura = False: rsAD.Update
                End If
                rsAD.Close
            End If
            gDirFactura = cDireccion.ItemData(cDireccion.ListIndex)
        End If
        
    End If
    '----------------------------------------------------------------------------------------------------------------------------------
    
End Function

Private Sub tVendedor_GotFocus()
    With tVendedor
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.Panels(1).Text = " Ingrese el dígito del vendedor."
End Sub

Private Sub tVendedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn And IsNumeric(tVendedor.Text) Then
        tVendedor.Tag = 0: tVendedor.Tag = BuscoUsuarioDigito(Val(tVendedor.Text), True)
        If Val(tVendedor.Tag) > 0 Then chPagaCheque.SetFocus
    End If
    
End Sub
Private Sub tVendedor_LostFocus()
    If Not IsNumeric(tVendedor.Text) Then tVendedor.Text = ""
    Status.Panels(1).Text = ""
End Sub

Private Sub ObtenerCamposFxPorSP(ByVal Documento As Long, ByRef sTxtoERI As String, ByRef sTxtClienteInferior As String, ByRef sTxtEnvios As String)
Dim RsF As rdoResultset
Dim sQy As String
Dim qItems As Integer, qEnvia As Integer, qInstala As Integer
    
    f_GetQEnvioInstala qItems, qEnvia, qInstala
    
    sQy = "EXEC prg_ImpresionContadoDatosRetira " & Documento & ", " & tVendedor.Tag & ", '" & tNombreC.Text & "', '" & strCodigoEnvio & "', " & qItems & ", " & qEnvia & ", " & qInstala
    
    Set RsF = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    If Not RsF.EOF Then
        If Not IsNull(RsF(0)) Then sTxtoERI = Trim(RsF(0))
        If Not IsNull(RsF(1)) Then sTxtClienteInferior = Trim(RsF(1))
        If Not IsNull(RsF(2)) Then sTxtEnvios = Trim(RsF(2))
    End If
    RsF.Close

End Sub

Private Sub AccionImprimir_VSReport(Documento As Long, ByVal TextoDoc As String, ByVal msgACaja As String, ByVal NoVaRUT As Boolean)
On Error GoTo ErrCrystal
Dim sTextoERI As String
Dim sTxtClienteInferior As String
Dim sTxtEnvios As String
Dim aTexto As String
Dim sPaso As String
    
    Err.Clear
    Screen.MousePointer = 11
    sPaso = "1"
    ObtenerCamposFxPorSP Documento, sTextoERI, sTxtClienteInferior, sTxtEnvios
    Dim oImprimo As New clsImpresionContado

    Dim oDonde As New clsParametrosImpresora
    
'    oDonde.Bandeja = paIContadoB
'    oDonde.Impresora = paIContadoN
'    oDonde.Papel = 1
    
    sPaso = "2"
    'Set oImprimo.DondeImprimo = New clsConfigImpresora
    
    On Error GoTo ErrCrystal
    oImprimo.DondeImprimo.Bandeja = paIContadoB
    sPaso = "2.1"
    oImprimo.DondeImprimo.Impresora = paIContadoN
    sPaso = "2.2"
    oImprimo.DondeImprimo.Papel = 1
    sPaso = "2.3"
    oImprimo.PathReportes = gPathListados
    sPaso = "2.4"
    oImprimo.StringConnect = miConexion.TextoConexion("Comercio")
    
    sPaso = "3"
    'Paso campos de consulta
    With oImprimo
        .field_Envios = sTxtEnvios
        .field_ClienteInferior = sTxtClienteInferior
        .field_NombreDocumento = paDContado

        If txtCliente.Cliente.Documento <> "" And txtCliente.Cliente.Tipo = TC_Persona Then aTexto = "(" & txtCliente.Cliente.Documento & ")" Else aTexto = ""
        If chNomDireccion.Value = 1 And Trim(labDireccion.Caption) <> "" Then aTexto = aTexto & " (" & Trim(cDireccion.Text) & ")"
        .field_ClienteNombre = Trim(tNombreC.Text) & " " & aTexto
        .field_ClienteDireccion = Trim(labDireccion.Caption)
        If txtCliente.Cliente.Tipo = TC_Empresa And txtCliente.Cliente.Documento <> "" Then
            .field_RUT = Trim(clsGeneral.RetornoFormatoRuc(txtCliente.Cliente.Documento))
            .field_CFinal = ""      'x defexto está en X
        ElseIf txtCliente.Cliente.Tipo = TC_Persona And txtCliente.Cliente.RutPersona <> "" And Not NoVaRUT Then
            .field_RUT = Trim(clsGeneral.RetornoFormatoRuc(txtCliente.Cliente.RutPersona))
            .field_CFinal = ""      'x defexto está en X
        End If

        .field_CodigoDeBarras = CodigoDeBarras(TipoDocumento.Contado, Documento)
        .field_TextoRetira = sTextoERI

        If AporteACuenta > 0 Then .field_MsgACaja = msgACaja
        sPaso = 4
        .ImprimoFacturaContado_VSReport Documento
    End With
    Screen.MousePointer = 0
    Exit Sub

ErrCrystal:
    On Error Resume Next
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al imprimir el documento, debe reimprimirlo." & vbCrLf & "Nro. Documento = " & Mid(TextoDoc, 1, 1) & " " & CLng(Trim(Mid(TextoDoc, 2, Len(TextoDoc)))) & vbCrLf & "Paso " & sPaso
    Exit Sub
End Sub

Private Sub CalculoArticulosEnEnvio()
Dim aValor As Long
    '---------------------------------------------------------------------------------
    'Pongo los artículos que estaban para envío en cero para consultar luego y poner realmente
    'los que van para envío. También elimino aquellos artículos que pagan flete para insertarlos luego.
    With vsGrilla
        I = 1
        Do While I <= .Rows - 1
            aValor = 0: .Cell(flexcpData, I, 6) = aValor
            If Trim(.Cell(flexcpData, I, 5)) <> "" Then RestoLabTotales CCur(.Cell(flexcpText, I, 5)), CCur(.Cell(flexcpText, I, 4)): .RemoveItem I: I = I - 1
            I = I + 1
        Loop
    End With
    '---------------------------------------------------------------------------------

    Cons = "Select Sum(REvAEntregar), REvArticulo From RenglonEnvio Where REvEnvio IN (" & strCodigoEnvio & ")" _
        & "Group by REvArticulo"
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, RsAux, Cons, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
    Do While Not RsAux.EOF
        With vsGrilla
            For I = 1 To .Rows - 1
                If RsAux!REvArticulo = CLng(.Cell(flexcpData, I, 0)) Then aValor = RsAux(0): .Cell(flexcpData, I, 6) = aValor
            Next I
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close

End Sub

Public Function CambioClienteEnvios(Cliente As Long, Envios As String) As Boolean
On Error GoTo ErrCCE
    
    CambioClienteEnvios = True
    If InStr(Envios, ",") = 0 Then If Envios = "0" Then Exit Function
    
    Cons = "Update Envio Set EnvCliente = " & Cliente & " Where EnvCodigo IN (" & Envios & ")"
    cBase.Execute (Cons)
    Exit Function

ErrCCE:
    clsGeneral.OcurrioError "Ocurrió un error al intentar modificar el cliente en los envíos."
    CambioClienteEnvios = False
End Function
Private Sub InsertoFilaCombo()
On Error GoTo ErrIF
Dim aValor As Currency
    
    Cons = "Select * From Presupuesto, Articulo Where PreID = " & tArticulo.Tag _
        & " And PreArticulo = ArtID"
    
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If ObtenerResultSet(cBase, RsAux, Cons, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
    With vsGrilla
        'Este artículo es el de bonificación tengo que ver si está insertado.
        If Not Ingresado(RsAux!ArtID) Then
            .AddItem CInt(tCantidad.Text)
            aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
            .Cell(flexcpData, .Rows - 1, 1) = TipoArticulo.Presupuesto
            .Cell(flexcpData, .Rows - 1, 2) = Trim(tArticulo.Tag)   'Este campo me dice si es pto. el código del mismo.
            aValor = RsAux!ArtTipo: .Cell(flexcpData, .Rows - 1, 3) = aValor
            .Cell(flexcpData, .Rows - 1, 4) = tUnitario.Tag
            .Cell(flexcpData, .Rows - 1, 6) = 0
            m_Patron = dis_arrMonedaProp(cMoneda.ItemData(cMoneda.ListIndex), pRedondeo)
            tUnitario.Text = Format(Redondeo(tUnitario.Text, m_Patron), FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 2) = Trim(tComentario.Text)
            .Cell(flexcpText, .Rows - 1, 3) = Trim(tUnitario.Text)
            
            If MnuOpFactSinIVA.Checked = False Then
                .Cell(flexcpText, .Rows - 1, 4) = IVAArticulo(tArticulo.Tag)
            Else
                .Cell(flexcpText, .Rows - 1, 4) = 0
            End If
            
            .Cell(flexcpText, .Rows - 1, 5) = Format(CCur(tUnitario.Text) * tCantidad.Text, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = "No"
            
            labIVA.Caption = Format(CCur(labIVA.Caption) + CCur(.Cell(flexcpText, .Rows - 1, 5)) - (CCur(.Cell(flexcpText, .Rows - 1, 5)) / CCur(1 + (CCur(.Cell(flexcpText, .Rows - 1, 4)) / 100))), FormatoMonedaP)
            labTotal.Caption = Format(CCur(labTotal.Caption) + CCur(CCur(tUnitario.Text) * tCantidad.Text), FormatoMonedaP)
            
            RsAux.Close
            
        Else
            RsAux.Close
            MsgBox "El combo seleccionado tiene el artículo " & Trim(RsAux!ArtNombre) & " que ya fue ingresado, no podrá insertar este combo.", vbInformation, "ATENCIÓN"
            Exit Sub
        End If
    End With
    labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), FormatoMonedaP)
    
    'Inserto los artículos del presupuesto.
    Cons = "Select * From PresupuestoArticulo, Articulo " _
            & " Left Outer Join PrecioVigente ON ArtID = PViArticulo " _
            & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
            & " And PViTipoCuota = " & paTipoCuotaContado _
        & " Where PArPresupuesto = " & tArticulo.Tag _
        & " And PArArticulo = ArtID"
    
    'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If ObtenerResultSet(cBase, RsAux, Cons, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
    m_Patron = dis_arrMonedaProp(cMoneda.ItemData(cMoneda.ListIndex), pRedondeo)
    
    Do While Not RsAux.EOF
        With vsGrilla
            If Not Ingresado(RsAux!ArtID) Then
                .AddItem CInt(tCantidad.Text) * RsAux!ParCantidad
                'DATA.
                aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                .Cell(flexcpData, .Rows - 1, 1) = TipoArticulo.Articulo
                .Cell(flexcpData, .Rows - 1, 2) = Trim(tArticulo.Tag)   'Este campo me dice si es pto. el código del mismo.
                aValor = RsAux!ArtTipo: .Cell(flexcpData, .Rows - 1, 3) = aValor
                .Cell(flexcpData, .Rows - 1, 6) = 0
                
                If Not IsNull(RsAux!PViPrecio) And RsAux!PViHabilitado Then
                    aValor = Redondeo(RsAux!PViPrecio, m_Patron)
                Else
                    aValor = 0
                    MsgBox "El artículo " & Trim(RsAux!ArtNombre) & " no posee precio contado o el mismo no está habilitado.", vbExclamation, "ATENCIÓN"
                    If aValor = 0 Then
                        .RemoveItem .Rows - 1   'La última es la que inserte recien
                        RsAux.Close
                        I = 1
                        Do While I <= .Rows - 1
                            If Trim(.Cell(flexcpData, I, 2)) = Trim(tArticulo.Tag) Then RestoLabTotales CCur(.Cell(flexcpText, I, 5)), CCur(.Cell(flexcpText, I, 4)): .RemoveItem I: I = I - 1
                            I = I + 1
                        Loop
                        tArticulo.SetFocus: LimpioRenglon
                        Exit Sub
                    End If
                End If
                
                .Cell(flexcpData, .Rows - 1, 4) = aValor
                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 2) = Trim(tComentario.Text)
                .Cell(flexcpText, .Rows - 1, 3) = Format(aValor, FormatoMonedaP)
                If MnuOpFactSinIVA.Checked = False Then
                    .Cell(flexcpText, .Rows - 1, 4) = IVAArticulo(tArticulo.Tag)
                Else
                    .Cell(flexcpText, .Rows - 1, 4) = 0
                End If
                .Cell(flexcpText, .Rows - 1, 5) = Format(aValor * CInt(tCantidad.Text) * RsAux!ParCantidad, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 6) = cEnvio.Text
            
                labIVA.Caption = Format(CCur(labIVA.Caption) + CCur(.Cell(flexcpText, .Rows - 1, 5)) - (CCur(.Cell(flexcpText, .Rows - 1, 5)) / CCur(1 + (CCur(.Cell(flexcpText, .Rows - 1, 4)) / 100))), FormatoMonedaP)
                labTotal.Caption = Format(CCur(labTotal.Caption) + CCur(.Cell(flexcpText, .Rows - 1, 5)), FormatoMonedaP)
                labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), FormatoMonedaP)
            Else
                MsgBox "El combo seleccionado tiene el artículo " & Trim(RsAux!ArtNombre) & " que ya fue ingresado, no podrá insertar este combo.", vbInformation, "ATENCIÓN"
                RsAux.Close
                'Borro todas las filas que tienen este combo.
                I = 1
                Do While I <= .Rows - 1
                    If Trim(.Cell(flexcpData, I, 2)) = Trim(tArticulo.Tag) Then RestoLabTotales CCur(.Cell(flexcpText, I, 5)), CCur(.Cell(flexcpText, I, 4)): .RemoveItem I: I = I - 1
                    I = I + 1
                Loop
                tArticulo.SetFocus: LimpioRenglon
                Exit Sub
            End If
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    LimpioRenglon
    tArticulo.SetFocus
    cMoneda.Enabled = False
    MnuEmitir.Enabled = True
    If cEnvio.ListIndex = 0 Then MnuEnvio.Enabled = True
    Exit Sub
ErrIF:
    clsGeneral.OcurrioError "Ocurrio un error inesperado al insertar el renglon."
End Sub

Private Sub InsertoFila()
On Error GoTo ErrIF
    
    If miRenglon.ArtCombo > 0 Then
        'Es combo.
        InsertoArticulosCombo 0, 0
    Else
    
        'Guardo desde que fecha está disponible el artículo.
        ReDim Preserve arrArtFechaRetira(UBound(arrArtFechaRetira) + 1)
        With arrArtFechaRetira(UBound(arrArtFechaRetira))
            .IDArticulo = miRenglon.IDArticulo
            .FRetira = miRenglon.DisponibleDesde
        End With
        
        CargoArticuloEnGrilla miRenglon.IDArticulo, miRenglon.Tipo, Val(tCantidad.Text), Articulo, miRenglon.Precio, miRenglon.NombreArticulo, tComentario.Text, tUnitario.Text, cEnvio.Text, miRenglon.EsInhabilitado, miRenglon.CodArticulo
        
        'Si ya tengo cargada fecha y es menor --> pongo esta fecha.
        'Como aún no puedo saber la Q de arts que están en envíos es algo para hacer dps.
        If IsDate(tFRetiro.Text) Then
            If CDate(tFRetiro.Text) < miRenglon.DisponibleDesde Then
                tFRetiro.Text = miRenglon.DisponibleDesde
            End If
        End If
        
    End If
    labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), FormatoMonedaP)
    MnuEmitir.Enabled = True
    If cEnvio.ListIndex = 0 Then MnuEnvio.Enabled = True
    LimpioRenglon
    tArticulo.SetFocus
    cMoneda.Enabled = False
    Exit Sub
ErrIF:
    clsGeneral.OcurrioError "Ocurrio un error inesperado al insertar el renglon."
End Sub

Private Function AvisoVentaLimitada(ByVal doyMsg As Boolean) As Boolean
On Error GoTo errCS
    Screen.MousePointer = 11
    AvisoVentaLimitada = False
    With vsGrilla
        For I = 1 To .Rows - 1
'            .Cell(flexcpFontStrikethru, I, 0) = False
'            .Cell(flexcpFontStrikethru, I, 1) = False
            .Cell(flexcpForeColor, I, 0) = vbBlack
            If .Cell(flexcpData, I, 1) = TipoArticulo.Servicio Then Exit For
            
            If Not EsTipoDeServicio(CLng(.Cell(flexcpData, I, 3))) Then
            'If Val(.Cell(flexcpData, I, 3)) <> paTipoArticuloServicio Then
                If Val(.Cell(flexcpData, I, 8)) = 0 And InStr(1, paCategoriaDistribuidor, "," & oCliente.Categoria & ",") > 0 Then
'                    .Cell(flexcpFontStrikethru, I, 0) = True
'                    .Cell(flexcpFontStrikethru, I, 1) = True
                    .Cell(flexcpForeColor, I, 0) = &HFF&
                    AvisoVentaLimitada = True
                    If doyMsg Then
                        MsgBox "Atención!!! " & vbCrLf & vbCrLf & "No está autorizada la venta del artículo " & .Cell(flexcpText, I, 1) & " a distribuidores." & vbCrLf & vbCrLf & "Debe consultar para vender.", vbExclamation, "POSIBLE ERROR"
                    End If
                    
                Else
                    If Val(.Cell(flexcpData, I, 8)) > 1 And Val(.Cell(flexcpData, I, 8)) < CInt(.Cell(flexcpText, I, 0)) Then
'                        .Cell(flexcpFontStrikethru, I, 0) = True
'                        .Cell(flexcpFontStrikethru, I, 1) = True
                        .Cell(flexcpForeColor, I, 0) = &HFF&
                        AvisoVentaLimitada = True
                        If doyMsg Then
                            MsgBox "Atención!!! " & vbCrLf & vbCrLf & "La cantidad máxima autorizada de venta para el artículo " & .Cell(flexcpText, I, 1) & " es de  " & Val(.Cell(flexcpData, I, 8)) & vbCrLf & vbCrLf & "Debe consultar para exceder dicha cantidad.", vbExclamation, "POSIBLE ERROR"
                        End If
                    End If
                End If
            End If
        Next I
    End With
    Screen.MousePointer = 0: Exit Function
errCS:
    clsGeneral.OcurrioError "Ocurrio un error al intentar controlar el stock.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function ControlStock() As Boolean
On Error GoTo errCS
    Screen.MousePointer = 11
    ControlStock = False
    With vsGrilla
        For I = 1 To .Rows - 1
            If .Cell(flexcpData, I, 1) = TipoArticulo.Servicio Then Exit For
            
            'If Val(.Cell(flexcpData, I, 3)) <> paTipoArticuloServicio Then
            If Not EsTipoDeServicio(CLng(.Cell(flexcpData, I, 3))) Then
                Cons = "Select StTCantidad From StockTotal Where StTArticulo = " & .Cell(flexcpData, I, 0)
                'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If ObtenerResultSet(cBase, RsAux, Cons, logComercio) <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
                If RsAux.EOF Then
                    If MsgBox("No hay stock para el artículo " & Trim(.Cell(flexcpText, I, 1)) & "." _
                        & Chr(13) & "¿Desea facturar de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                        RsAux.Close: Exit Function
                    End If
                Else
                    If RsAux!StTCantidad < CInt(.Cell(flexcpText, I, 0)) Then
                        If MsgBox("No existe tanto stock para el artículo " & Trim(.Cell(flexcpText, I, 1)) & "." & Chr(13) _
                            & "¿Desea facturar de todas formas?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then
                                RsAux.Close: Exit Function
                        End If
                    End If
                End If
                RsAux.Close
                
                'HAGO CONTROL DE LIMITACIÓN DE VENTA.
                
            End If
        Next I
    End With
    ControlStock = True
    Screen.MousePointer = 0: Exit Function
errCS:
    clsGeneral.OcurrioError "Ocurrio un error al intentar controlar el stock.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function VerificoCostoArticulo() As Currency
On Error GoTo errVCA
    ' Retorno  0 = No hay cambios de precio
    VerificoCostoArticulo = 0
    With vsGrilla
        For I = 1 To .Rows - 1
            If .Cell(flexcpData, I, 4) <> "" Then
                If CCur(.Cell(flexcpData, I, 4)) <> CCur(.Cell(flexcpText, I, 3)) And CCur(.Cell(flexcpData, I, 4)) <> 0 Then
                    VerificoCostoArticulo = VerificoCostoArticulo + (CCur(.Cell(flexcpText, I, 3)) - CCur(.Cell(flexcpData, I, 4)))
                End If
            End If
        Next
    End With
    Exit Function
errVCA:
End Function

Private Sub Calc_IvaCofis(ByRef cI As Currency, cC As Currency)
Dim iCont As Integer
    cI = 0: cC = 0
    With vsGrilla
        For iCont = 1 To .Rows - 1
            
        Next
    End With
End Sub

Private Function RetiraEnDeposito() As Boolean
    RetiraEnDeposito = False
    With vsGrilla
        For I = 1 To .Rows - 1
            If CLng(.Cell(flexcpData, I, 1)) <> TipoArticulo.Servicio And CInt(.Cell(flexcpText, I, 0)) - CInt(.Cell(flexcpData, I, 6)) > 0 Then
                Cons = "SELECT * FROM Articulo WHERE ArtID = " & CLng(.Cell(flexcpData, I, 0)) & " AND ArtLocalRetira = 6"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    RetiraEnDeposito = True
                    RsAux.Close
                    Exit Function
                End If
                RsAux.Close
            End If
        Next
    End With
    
End Function

Private Sub GraboFacturaRenglon(ByVal bSucVtaXMayor As Boolean, ByVal NoVaRUT As Boolean, ByVal resQV As Byte, ByVal subRespQV As Byte)
Dim Control As Currency
Dim strDefensa As String
Dim NroDocumento As Long, aUsuario As Long
Dim cAuxCalc As Currency
Dim sDefInh As String, lUsuInh As Long
Dim objSuceso As clsSuceso
Dim sDefCambioNombre As String, bSucesoNombre As Boolean
Dim idAutPrecio As Long, idAutCamName As Long, idAutInh As Long
Dim cIVA As Currency
Dim sDefNoVender As String, iUsuNoV As Long, idAutNoV As Long
2
Dim saldoCtaPers As Currency
Dim idMovCaja As Long

Dim sPaso As String
Dim sFRetira As String ', sTextoImp As String
sPaso = "1"

    If TasaBasica = 0 Then CargoValoresIVA

    saldoCtaPers = SaldoCuentaPersonal(oCliente.ID, False)
    If saldoCtaPers > 0 Then
        Dim dlgRet As VbMsgBoxResult
        Do
            dlgRet = (MsgBox("¿El cliente va a utilizar los aportes pendientes para saldar el documento?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Utilizar aportes"))
        Loop Until dlgRet <> vbCancel
        If dlgRet = vbNo Then saldoCtaPers = 0
    End If

    Screen.MousePointer = vbHourglass
    sPaso = "2"
    strDefensa = "": aUsuario = 0
    idAutPrecio = 0: idAutCamName = 0: idAutInh = 0
    Control = VerificoCostoArticulo
    If Control <> 0 Then
        'Llamo al registro del Suceso-------------------------------------------------------------
        Set objSuceso = New clsSuceso
        aUsuario = 0
        objSuceso.TipoSuceso = TipoSuceso.ModificacionDePrecios
        objSuceso.ActivoFormulario CLng(tUsuario.Tag), "Cambio de Precio", cBase
        Me.Refresh
        aUsuario = objSuceso.Usuario
        strDefensa = objSuceso.Defensa
        idAutPrecio = objSuceso.Autoriza
        Set objSuceso = Nothing
        If aUsuario = 0 Then Screen.MousePointer = 0: Exit Sub
    End If
    
    lUsuInh = 0: sDefInh = ""
    'Ahora visualizo si hay artículos inhabilitados y le pido defensa para todos.
    For I = 1 To vsGrilla.Rows - 1
        If Trim(vsGrilla.Cell(flexcpText, I, 7)) <> "" And Not (Val(vsGrilla.Cell(flexcpData, I, 2)) > 0 And vsGrilla.Cell(flexcpData, I, 1) = TipoArticulo.Especifico) Then
            Set objSuceso = New clsSuceso
            objSuceso.TipoSuceso = TipoSuceso.FacturaArticuloInhabilitado
            objSuceso.ActivoFormulario CLng(tUsuario.Tag), "Artículos inhabilitados", cBase
            Me.Refresh
            lUsuInh = objSuceso.Usuario
            sDefInh = objSuceso.Defensa
            idAutInh = objSuceso.Autoriza
            Set objSuceso = Nothing
            If lUsuInh = 0 Then Screen.MousePointer = 0: Exit Sub Else Exit For
        End If
    Next I
    
    'Veo si le cambio el nombre al cliente.
    If Trim(oCliente.MostrarComo) <> Trim(tNombreC.Text) Then
        bSucesoNombre = True
        'Llamo al registro del Suceso-------------------------------------------------------------
        Set objSuceso = New clsSuceso
        aUsuario = 0
        objSuceso.TipoSuceso = TipoSuceso.FacturaCambioNombre
        objSuceso.ActivoFormulario CLng(tUsuario.Tag), "Cambio de Nombre", cBase
        Me.Refresh
        aUsuario = objSuceso.Usuario
        sDefCambioNombre = objSuceso.Defensa
        idAutCamName = objSuceso.Autoriza
        Set objSuceso = Nothing
        If aUsuario = 0 Then Screen.MousePointer = 0: Exit Sub
    Else
        bSucesoNombre = False
    End If
    
    If txtCliente.Cliente.NoVender Then
        Set objSuceso = New clsSuceso
        objSuceso.TipoSuceso = TipoSuceso.ClienteNoVender
        objSuceso.ActivoFormulario CLng(tUsuario.Tag), "No vender a cliente", cBase
        Me.Refresh
        iUsuNoV = objSuceso.Usuario
        sDefNoVender = objSuceso.Defensa
        idAutNoV = objSuceso.Autoriza
        Set objSuceso = Nothing
        If iUsuNoV = 0 Then Screen.MousePointer = 0: Exit Sub
    End If
    
    Dim sDefVtaxMayor As String, iUsuVtaXMayor As Long, idAutVtaXMayor As Long
    If bSucVtaXMayor Then
        Set objSuceso = New clsSuceso
        objSuceso.TipoSuceso = TipoSuceso.FacturaArticuloInhabilitado
        objSuceso.ActivoFormulario CLng(tUsuario.Tag), "Venta por mayor inhabilitada", cBase
        Me.Refresh
        iUsuVtaXMayor = objSuceso.Usuario
        sDefVtaxMayor = objSuceso.Defensa
        idAutVtaXMayor = objSuceso.Autoriza
        Set objSuceso = Nothing
        If iUsuVtaXMayor = 0 Then Screen.MousePointer = 0: Exit Sub
    End If
    
    'Si tiene artículo especifico el envío no puede estar en un vacon.
    If strCodigoEnvio <> "" Then
        Dim bHayEsp As Boolean
        With vsGrilla
            For I = 1 To .Rows - 1
                If Val(.Cell(flexcpData, I, 2)) > 0 And .Cell(flexcpData, I, 1) = TipoArticulo.Especifico Then
                    bHayEsp = True
                    Exit For
                End If
            Next
        End With
        If bHayEsp Then
            Cons = "SELECT EVCEnvio FROM EnvioVaCon WHERE EVCEnvio IN (" & strCodigoEnvio & ")"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                MsgBox "No puede agregar un artículo específico en un VACON, debe editar el envío y desasignar el mismo.", vbCritical, "ATENCIÓN"
                RsAux.Close
                Screen.MousePointer = 0: Exit Sub
            End If
            RsAux.Close
        End If
    End If
    
    
    On Error GoTo ErrGFR
    '--------------------------------------------------------------------------------------------------
    'GetStringRetira sTextoImp, sFRetira
    GetStringRetira sFRetira
    
    If IsDate(tFRetiro.Text) Then
        sFRetira = CDate(tFRetiro.Text)
        Dim dFRet As Date: dFRet = ObtenerFechaRetiro
        If dFRet > DateSerial(2000, 1, 1) Then
            If CDate(sFRetira) > dFRet Then
                If Abs(DateDiff("d", CDate(sFRetira), dFRet)) > 59 Then
                    sFRetira = sFRetira & " " & "01:59:00"
                Else
                    sFRetira = sFRetira & " " & "01:" & Format(Abs(DateDiff("d", dFRet, CDate(sFRetira))), "00") & ":00"
                End If
            Else
                sFRetira = sFRetira & " " & "01:00:00"
            End If
        End If
    End If
    sPaso = "3"
    
    If cDireccion.ListIndex > -1 Then
        Dim lngZona As Long
        lngZona = BuscoZonaDireccion(cDireccion.ItemData(cDireccion.ListIndex))
    End If
    
    Dim oTransacciones As New clsCobrarConQuePaga
    Dim colDocsAImprimir As Collection
    
    Dim oCli As New clsClienteCFE
    With oCli
        .Codigo = txtCliente.Cliente.Codigo
        .TipoCliente = txtCliente.Cliente.Tipo
        If txtCliente.Cliente.Tipo = TC_Empresa Then
            .RUT = txtCliente.Cliente.Documento
            .CodigoDGICI = TD_RUT
        Else
            .cI = clsGeneral.QuitoFormatoCedula(txtCliente.Cliente.Documento)
            If txtCliente.Cliente.RutPersona <> "" And Not NoVaRUT Then
                .CodigoDGICI = TD_RUT
                .RUT = txtCliente.Cliente.RutPersona
            Else
                .CodigoDGICI = txtCliente.Cliente.TipoDocumento.TipoDocIdDGI
            End If
        End If
        .NombreCliente = tNombreC.Text
        .Direccion.Departamento = DeptoDir
        .Direccion.Localidad = LocalDir
        .Direccion.Domicilio = labDireccion.Caption
        .CodigoDGIPais = txtCliente.Cliente.TipoDocumento.Pais.CodigoDGI
    End With
    
    Dim tipoCAE As Byte
    tipoCAE = IIf((txtCliente.Cliente.Tipo = TC_Empresa And txtCliente.Cliente.Documento <> "") Or _
            (txtCliente.Cliente.Tipo = TC_Persona And txtCliente.Cliente.RutPersona <> "" And NoVaRUT = False), CFE_eFactura, CFE_eTicket)
    
    FechaDelServidor
    
    cBase.BeginTrans 'Comienzo la TRANSACCION-------------------------------------------------------------------------
    On Error GoTo ErrResumo
    
    If saldoCtaPers > 0 Then
        'Valido que no me hayan modificado el saldo y además bloqueo.
        saldoCtaPers = SaldoCuentaPersonal(oCliente.ID, False)
        If saldoCtaPers = 0 Then
            cBase.RollbackTrans
            Screen.MousePointer = 0
            MsgBox "Atención los aportes ya fueron asignados, verifique.", vbExclamation, "ATENCIÓN"
            Exit Sub
        ElseIf saldoCtaPers > CCur(labTotal.Caption) Then
            saldoCtaPers = CCur(labTotal.Caption)
        End If
    End If
    
        
    Dim CAE As New clsCAEDocumento
    Dim caeG As New clsCAEGenerador
    Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, tipoCAE, paCodigoDGI)
    Set caeG = Nothing
'FIN Obtención CAE.
    Dim doc As New clsDocumentoCGSA
    With doc
        Set .Cliente = oCli
        .Emision = gFechaServidor
        .Tipo = TD_Contado
        .Numero = CAE.Numero
        .Serie = CAE.Serie
        .Moneda.Codigo = cMoneda.ItemData(cMoneda.ListIndex)
        .Total = CCur(labTotal.Caption)
        .IVA = CCur(labIVA.Caption)
        .Sucursal = paCodigoDeSucursal
        .Digitador = tUsuario.Tag
        .Comentario = tComentarioDocumento.Text & IIf(tBanco.Text <> "", " c/Chq", "")
        .Zona = lngZona
        .FechaRetira = CDate(sFRetira)
        If cPendiente.ListIndex > -1 Then .Pendiente = cPendiente.ItemData(cPendiente.ListIndex)
        .Vendedor = Val(tVendedor.Tag)
    End With

    cAuxCalc = 0
    Dim oRen As clsDocumentoRenglon
    With vsGrilla
        For I = 1 To .Rows - 1
            
            cAuxCalc = Format((CCur(.Cell(flexcpText, I, 5)) - (CCur(.Cell(flexcpText, I, 5)) / CCur(1 + (CCur(.Cell(flexcpText, I, 4)) / 100)))) / .Cell(flexcpValue, I, 0), "####0.00")
            cIVA = cIVA + (cAuxCalc * .Cell(flexcpValue, I, 0))
            
            Set oRen = New clsDocumentoRenglon
            
            oRen.Articulo.ID = CLng(.Cell(flexcpData, I, 0))
            If Val(.Cell(flexcpData, I, 2)) > 0 And .Cell(flexcpData, I, 1) = TipoArticulo.Especifico Then
                oRen.Articulo.idEspecifico = Val(.Cell(flexcpData, I, 2))
            End If
            If Val(.Cell(flexcpText, I, 9)) > 0 Then oRen.Articulo.Codigo = .Cell(flexcpText, I, 9)
            oRen.Articulo.Nombre = .Cell(flexcpText, I, 1)
                        
            oRen.Articulo.TipoIVA.Porcentaje = .Cell(flexcpText, I, 4)
            oRen.Articulo.TipoArticulo = CLng(.Cell(flexcpData, I, 3))
            oRen.Cantidad = CInt(.Cell(flexcpText, I, 0))
            oRen.IVA = cAuxCalc
            oRen.Precio = CCur(.Cell(flexcpText, I, 3))
            If CLng(.Cell(flexcpData, I, 1)) = TipoArticulo.Servicio Then 'Or CLng(.Cell(flexcpData, I, 3)) = paTipoArticuloServicio Then
                oRen.CantidadARetirar = 0
            Else
                oRen.CantidadARetirar = CInt(.Cell(flexcpText, I, 0)) - CInt(.Cell(flexcpData, I, 6))
            End If
            oRen.Descripcion = Trim(.Cell(flexcpText, I, 2))
            doc.Renglones.Add oRen
        
        Next
    End With
    If CCur(Format(cIVA, "###0.00")) <> CCur(labIVA.Caption) Then doc.IVA = Format(cIVA, "###0.00")
    
    Set doc.Conexion = cBase
    NroDocumento = doc.InsertoDocumentoBD(0)
    doc.Codigo = NroDocumento
    
    With vsGrilla
        For I = 1 To .Rows - 1
            'Stock Virtual
'            If CLng(.Cell(flexcpData, I, 1)) <> TipoArticulo.Servicio And _
                CLng(.Cell(flexcpData, I, 3)) <> paTipoArticuloServicio Then
                    
            If CLng(.Cell(flexcpData, I, 1)) <> TipoArticulo.Servicio And Not EsTipoDeServicio(CLng(.Cell(flexcpData, I, 3))) Then
                    'MarcoStockVenta CLng(tUsuario.Tag), CLng(.Cell(flexcpData, i, 0)), CInt(.Cell(flexcpText, i, 0)) - CInt(.Cell(flexcpData, i, 6)), CInt(.Cell(flexcpData, i, 6)), 0, TipoDocumento.Contado, NroDocumento, paCodigoDeSucursal
                    cBase.Execute "Exec StockMovimientoVenta " & CLng(tUsuario.Tag) & "," & CLng(.Cell(flexcpData, I, 0)) & ", " & CInt(.Cell(flexcpText, I, 0)) - CInt(.Cell(flexcpData, I, 6)) & "," & CInt(.Cell(flexcpData, I, 6)) & "," & TipoDocumento.Contado & ", " & NroDocumento & ", " & paCodigoDeSucursal
            End If
            
            If Val(.Cell(flexcpData, I, 2)) > 0 And .Cell(flexcpData, I, 1) = TipoArticulo.Especifico Then
                'Updateo la tabla.
                Cons = "Update ArticuloEspecifico " & _
                    "Set AEsTipoDocumento = 1, AEsDocumento = " & NroDocumento & _
                    ", AEsModificado = '" & Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss") & "'" & _
                    " Where AEsID = " & Val(.Cell(flexcpData, I, 2))
                cBase.Execute Cons
            End If
            
        Next I
    End With
    
    'Para ver si es un servicio busco en el data 1 si es tipo servicio updateo el mismo.
    If vsGrilla.Cell(flexcpData, 1, 1) = TipoArticulo.Servicio Then
        'Updateo el servicio y le pongo el ID del documento.
        Cons = "Select * From Servicio Where SerCodigo = " & vsGrilla.Cell(flexcpData, 1, 5)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        RsAux!SerDocumento = NroDocumento
        RsAux.Update
        RsAux.Close
    End If
    
    sPaso = "9"
    If Trim(strCodigoEnvio) <> "" Then
        
        If InStr(1, strCodigoEnvio, ",") Then
            lUltimoEnvio = Mid(strCodigoEnvio, 1, InStr(1, strCodigoEnvio, ",") - 1)
        Else
            lUltimoEnvio = strCodigoEnvio
        End If
        
        'si hay envios pongo en envDocumentoFactura el documento que lo facturo.
        'osea este.
        Cons = "UPDATE Envio Set EnvDocumento = " & NroDocumento _
            & " , EnvCliente = " & oCliente.ID & " , EnvDocumentoFactura = " & NroDocumento & " , EnvUsuario = " & tUsuario.Tag _
            & " WHERE EnvCodigo IN (" & strCodigoEnvio & ")" _
            & " And EnvFormaPago = " & TipoPagoEnvio.PagaAhora
        
        cBase.Execute (Cons)
            
        'Si hay envios que no los paga ahora pero los hizo.
        Cons = "UPDATE Envio Set EnvDocumento = " & NroDocumento _
            & " , EnvCliente = " & oCliente.ID & " , EnvUsuario = " & tUsuario.Tag & " WHERE EnvCodigo IN (" & strCodigoEnvio & ")" _
            & " And EnvFormaPago <> " & TipoPagoEnvio.PagaAhora
        cBase.Execute (Cons)
        
        Cons = "UPDATE EnvioVaCon Set EVCDocumento = " & NroDocumento & "WHERE EVCEnvio IN (" & strCodigoEnvio & ")"
        cBase.Execute (Cons)
    End If
        
    sPaso = "10"
    If Control <> 0 Then
        aTexto = "Contado " & CAE.Serie & " " & CAE.Numero & " (" & Trim(gFechaServidor) & ")"
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, TipoSuceso.ModificacionDePrecios, paCodigoDeTerminal, aUsuario, NroDocumento, Descripcion:=aTexto, Defensa:=Trim(strDefensa), Valor:=Control, idCliente:=oCliente.ID, idautoriza:=idAutPrecio
    End If
    
    If lUsuInh > 0 Then
        aTexto = "Artículos: "
        For I = 1 To vsGrilla.Rows - 1
            If Trim(vsGrilla.Cell(flexcpText, I, 7)) <> "" Then aTexto = aTexto & Trim(vsGrilla.Cell(flexcpData, I, 7)) & ", "
        Next I
        aTexto = Mid(aTexto, 1, Len(aTexto) - 2)
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, TipoSuceso.FacturaArticuloInhabilitado, paCodigoDeTerminal, lUsuInh, NroDocumento, Descripcion:=aTexto, Defensa:=Trim(sDefInh), Valor:=1, idCliente:=txtCliente.Cliente.Codigo, idautoriza:=idAutInh
    End If
    If bSucesoNombre Then
        aTexto = "Cambio de Nombre en Contado"
        sDefCambioNombre = "Nuevo nombre: " & Trim(tNombreC.Text) & vbCrLf & sDefCambioNombre
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, TipoSuceso.FacturaCambioNombre, paCodigoDeTerminal, aUsuario, NroDocumento, , aTexto, Trim(sDefCambioNombre), 1, txtCliente.Cliente.Codigo, idAutCamName
    End If
    If iUsuNoV > 0 Then
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, TipoSuceso.ClienteNoVender, paCodigoDeTerminal, iUsuNoV, NroDocumento, , "Cliente No vender", Trim(sDefNoVender), 1, txtCliente.Cliente.Codigo, idAutNoV
    End If
    If iUsuVtaXMayor > 0 Then
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, TipoSuceso.FacturaArticuloInhabilitado, paCodigoDeTerminal, iUsuVtaXMayor, NroDocumento, , "Venta por mayor de artículo", Trim(sDefVtaxMayor), 1, txtCliente.Cliente.Codigo, idAutVtaXMayor
    End If
    
    sPaso = "11"
    If lBanco.Visible Then
    
        Cons = "Select * From Comentario Where ComCliente = " & oCliente.ID _
            & " And ComTipo = " & paTipoComCheque
        If Val(lBanco.Tag) > 0 Then
            Cons = Cons & " And ComCodigo = " & Val(lBanco.Tag)
        End If
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Edit
        Else
            RsAux.AddNew
            RsAux!ComCliente = oCliente.ID
            RsAux!ComTipo = paTipoComCheque
        End If
        RsAux!ComFecha = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
        RsAux!ComComentario = "Banco: " & Trim(tBanco.Tag) _
                & Format(tBanco.Text, " (00-000)") & " A nombre de: " & Trim(tANombre.Text)
        RsAux!ComUsuario = tUsuario.Tag
        RsAux.Update
        RsAux.Close
        
        Cons = "Select * From Cliente Where CliCodigo = " & oCliente.ID
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        RsAux.Edit
        RsAux!CliCheque = "S"
        RsAux.Update
        RsAux.Close
    End If
    
    sPaso = "12"
    'Si tiene prendida la señal de entrega en local.
    If chkRetiraAqui.Value And chkRetiraAqui.Visible Then
        cBase.Execute "prg_EntregaMercaderia_EntregoArticulo 9999, 9898, 0, 0, " & NroDocumento & ", " & paCodigoDeSucursal & ", 0, 0"
    End If
    
    
    'ENCUESTA EN QUE VINO.
    If resQV <> 255 Then
        Cons = "INSERT INTO EnQueVino VALUES (" & NroDocumento & ", " & resQV & ", " & IIf(subRespQV > 0, subRespQV, "Null") & ")"
        cBase.Execute Cons
    End If
    
    cBase.CommitTrans
    'Fin Transaccion-------------------------------------------------------------------------!!!!!!!!!!!!!
    
    Dim sErrEfactura As String
    lUltimoDoc = NroDocumento
    On Error GoTo ErrImpresion
    
    sPaso = EmitirCFE(doc, CAE)
    If sPaso <> "" Then
        MsgBox "ATENCIÓN no se firmo el documento: " & sPaso, vbCritical, "ATENCIÓN"
        EnvioALog "No se firmo el documento: " & sPaso
    Else
        cBase.Execute "EXEC prg_PosInsertoDocumentosATickets '" & doc.Codigo & "', " & oCnfgPrint.ImpresoraTickets
    End If
    Set doc = Nothing
    
    sPaso = "18"
'    If Not bATicket Then
'        sPaso = "19"
'        Dim msgACaja As String
''        If saldoCtaPers > 0 Then
''            If saldoCtaPers >= CCur(labTotal.Caption) Then
''                msgACaja = "Paga factura con aportes por $ " & Format(CCur(labTotal.Caption), FormatoMonedaP)
''            Else
''                msgACaja = "Paga con aportes por $ " & Format(saldoCtaPers, FormatoMonedaP) & ", cobrar $ " & Format(CCur(labTotal.Caption) - saldoCtaPers, FormatoMonedaP)
''            End If
''        End If
'    End If
    strCodigoEnvio = vbNullString
    
'    If saldoCtaPers > 0 And colDocsAImprimir.Count > 0 Then
'        ImprimoDocumentosDeAportes colDocsAImprimir(1), NroSerie & " " & Numero
'    End If
    
    On Error Resume Next
    Cons = "Select RenDocumento From Renglon, Articulo Where RenDocumento = " & NroDocumento _
            & " And ArtInstalador > 0 And RenArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then EjecutarApp App.Path & "\Instalaciones.exe", "doc:" & CStr(NroDocumento)
    RsAux.Close
    
    labArticulo.Tag = "0": labArticulo.Caption = "&Artículo"
    CamposRenglon True
    MnuLimpiar_Click
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrGFR:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al iniciar la transacción.", sPaso & Err.Description
    Exit Sub

ErrResumo:
    Resume ErrRelajo
    
ErrRelajo:
    Screen.MousePointer = vbDefault
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al emitir la factura de contado.", sPaso & Err.Description
    Exit Sub
    
ErrImpresion:
    strCodigoEnvio = vbNullString
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrió un error al imprimir o restaurar el formulario.", sPaso & Err.Description
    strCodigoEnvio = vbNullString
    MnuLimpiar_Click
    Exit Sub
End Sub

Private Sub ImprimoDocumentosDeAportes(ByVal oDocs As clsDocAImprimir, ByVal Documento As String)
On Error GoTo errIDA
Dim oPrint As New clsImpresionDeDocumentos
    
    Err.Clear
    Set oPrint.Conexion = cBase
    oPrint.PathReportes = gPathListados
    oPrint.NombreBaseDatos = miConexion.RetornoPropiedad(False, False, False, True)
'    Set oPrint.DondeImprimo = New clsConfigImpresora
    If oCnfgPrintSalidaCaja.Opcion = 0 Then
        oPrint.DondeImprimo.Impresora = paIRemitoN
        oPrint.DondeImprimo.Bandeja = paIRemitoB
    Else
        oPrint.DondeImprimo.Impresora = oCnfgPrintSalidaCaja.ImpresoraTickets
    End If
    If oCnfgPrintSalidaCaja.Opcion = 0 Then
        oPrint.ImprimoSalidaCaja_Crystal oDocs.IDDocumento, "Señas Recibas", "$", Val(tUsuario.Tag), paNombreSucursal
    Else
        oPrint.ImprimoSalidaCajaTicket oDocs.IDDocumento, paNombreSucursal, tUsuario.Text, "Señas recibidas, Ctdo:" & Documento
    End If
    Exit Sub
errIDA:
    clsGeneral.OcurrioError "Error al imprimir la salida de caja.", Err.Description, "Impresión de salida de caja"
End Sub

Private Sub CargoRenglonesEnvio(ByVal CodEnvios As String)
Dim lngCodEnvio As Long, aValor As Currency, cIVAAux As Currency
Dim strAux As String
Dim RsArt As rdoResultset
Dim bolInserte As Boolean
Dim cIVA As Currency

    strAux = CodEnvios
    Do While strAux <> ""
        If InStr(1, strAux, ",") > 0 Then
            CodEnvios = Left(strAux, InStr(1, strAux, ","))
            lngCodEnvio = CLng(Left(CodEnvios, InStr(1, CodEnvios, ",") - 1))
            strAux = Right(strAux, Len(strAux) - InStr(1, strAux, ","))
        Else
            lngCodEnvio = CLng(strAux)
            strAux = ""
        End If
        
        Cons = "Select * From Envio Where EnvCodigo = " & lngCodEnvio
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        Do While Not RsAux.EOF
            'Si lo factura ahora lo cargo.-----------------------------
            If RsAux!EnvFormaPago = TipoPagoEnvio.PagaAhora Then
            
                Cons = "Select * From TipoFlete, Articulo, ArticuloFacturacion, TipoIva " _
                    & " Where TFlCodigo = " & RsAux!EnvTipoFlete _
                    & " And ArtID = TFlArticulo And ArtId = AFaArticulo  And AFaIva = IvaCodigo"
                
                Set RsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                
                If RsArt.EOF Then
                    MsgBox "Error crítico, se almacenó un valor de tipo de flete incorrecto.", vbCritical, "ATENCIÓN"
                    RsArt.Close: Exit Sub
                End If
                
                bolInserte = False
                
                'Busco si ya ingreso un artículo con ese código. si eso ocurrio cierro el
                'resultset sino  cuando sale lo inserto..
                With vsGrilla
                    
                    For I = 1 To .Rows - 1
                        
                        If RsArt!ArtID = CLng((.Cell(flexcpData, I, 0))) Then   'Es este.------
                            
                            aValor = RsAux!EnvCodigo: .Cell(flexcpData, I, 5) = aValor
                            .Cell(flexcpText, I, 0) = CInt(.Cell(flexcpText, I, 0)) + 1
                            
                            aValor = RsAux!EnvValorFlete
                            .Cell(flexcpText, I, 5) = Format(CCur(.Cell(flexcpText, I, 5)) + aValor, FormatoMonedaP)
                            .Cell(flexcpText, I, 3) = Format(CCur(.Cell(flexcpText, I, 5)) / CCur(.Cell(flexcpText, I, 0)), FormatoMonedaP)
                            
                            'Le sumo el valor al total.
                            labTotal.Caption = Format(CCur(labTotal.Caption) + aValor, "#,##0.00")
                            'Se sumo el iva del valor al total.
                            cIVAAux = RsAux!EnvValorFlete - (RsAux!EnvValorFlete / CCur(1 + (CCur(.Cell(flexcpText, I, 4)) / 100)))
                            labIVA.Caption = Format(CCur(labIVA.Caption) + cIVAAux, "#,##0.00")
                            
                            aValor = .Cell(flexcpValue, I, 3): .Cell(flexcpData, I, 4) = aValor
                            
                            cIVAAux = Format(CCur(.Cell(flexcpText, I, 5)) / CCur(1 + (CCur(.Cell(flexcpText, I, 4)) / 100)), "0.00")
                            cIVAAux = CCur(.Cell(flexcpText, I, 5)) - cIVAAux
                            
                            
                            labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
                            RsArt.Close
                            bolInserte = True
                            Exit For
                        End If
                    Next I
                End With
                If Not bolInserte Then
                    With vsGrilla
                        .AddItem "1"
                        'DATA.
                        aValor = RsArt!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                        .Cell(flexcpData, .Rows - 1, 1) = TipoArticulo.PagoFlete
                        .Cell(flexcpData, .Rows - 1, 2) = 0                                                     'Este campo me dice si es pto. el código del mismo.
                        aValor = RsArt!ArtTipo: .Cell(flexcpData, .Rows - 1, 3) = aValor           'Guardo el tipo de Articulo
                        aValor = RsAux!EnvCodigo: .Cell(flexcpData, .Rows - 1, 5) = aValor
                        .Cell(flexcpData, .Rows - 1, 6) = 0                                                                 'Artículos que están para envio.

                        .Cell(flexcpText, .Rows - 1, 1) = Trim(RsArt!ArtNombre)
                        .Cell(flexcpText, .Rows - 1, 2) = ""
                        
                        cIVA = Format(IVAArticulo(RsArt!ArtID), "#,##0.00")
                        aValor = RsAux!EnvValorFlete
                        

                        .Cell(flexcpText, .Rows - 1, 4) = Format(cIVA, "#,##0.00")
                        .Cell(flexcpText, .Rows - 1, 3) = Format(aValor, "#,##0.00")
                        .Cell(flexcpText, .Rows - 1, 5) = Format(aValor, "#,##0.00")
                        
                        .Cell(flexcpData, .Rows - 1, 4) = aValor     'Guardo el costo unitario.
                        .Cell(flexcpText, .Rows - 1, 6) = "No"
                        
                        .Cell(flexcpText, .Rows - 1, 9) = RsArt!ArtCodigo
                        
                        cIVAAux = Format(CCur(.Cell(flexcpText, .Rows - 1, 5)) / CCur(1 + (CCur(.Cell(flexcpText, .Rows - 1, 4)) / 100)), "0.00")
                        cIVAAux = CCur(.Cell(flexcpText, .Rows - 1, 5)) - cIVAAux
                        
                        labIVA.Caption = Format(CCur(labIVA.Caption) + cIVAAux, "#,##0.00")
                        labTotal.Caption = Format(CCur(labTotal.Caption) + aValor, "#,##0.00")
                    End With
                    labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
                    RsArt.Close
                End If
                
                'Valores de piso.------------------------------
                If Not IsNull(RsAux!EnvValorPiso) Then
                
                    Cons = "Select * From Articulo, ArticuloFacturacion, TipoIva" _
                        & " Where ArtId = " & paArticuloPisoAgencia _
                        & " And ArtID = AFaArticulo And AFaArticulo = " & paArticuloPisoAgencia _
                        & " And AFaIVA = IVACodigo"
                        
                    Set RsArt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                
                    If RsArt.EOF Then
                        MsgBox "Error crítico: existe un costo de piso y no existe un artículo asociado a su facturación.", vbCritical, "ATENCIÓN"
                        RsArt.Close
                        Exit Sub
                    End If
                    bolInserte = False
                    
                    With vsGrilla
                        For I = 1 To .Rows - 1
                            If RsArt!ArtID = CLng((.Cell(flexcpData, I, 0))) Then   'Es este.------
                                aValor = RsAux!EnvCodigo: .Cell(flexcpData, I, 5) = aValor
                                .Cell(flexcpText, I, 0) = CInt(.Cell(flexcpText, I, 0)) + 1
                                
                                aValor = RsAux!EnvValorPiso
                                
                                .Cell(flexcpText, I, 5) = Format(CCur(.Cell(flexcpText, I, 5)) + aValor, "#,##0.00")
                                .Cell(flexcpText, I, 3) = Format(CCur(.Cell(flexcpText, I, 5)) / CCur(.Cell(flexcpText, I, 0)), "#,##0.00")
                                
                                
                                labTotal.Caption = Format(CCur(labTotal.Caption) + aValor, "#,##0.00")
                                
                                aValor = .Cell(flexcpText, I, 3): .Cell(flexcpData, I, 4) = aValor
                                
                                cIVAAux = Format(CCur(.Cell(flexcpText, I, 5)) / CCur(1 + (CCur(.Cell(flexcpText, I, 4)) / 100)), "0.00")
                                cIVAAux = CCur(.Cell(flexcpText, I, 5)) - cIVAAux
                                
                                labIVA.Caption = Format(CCur(labIVA.Caption) + cIVAAux, "#,##0.00")
                                labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
                                RsArt.Close
                                bolInserte = True
                                Exit For
                            End If
                        Next
                        If Not bolInserte Then
                            .AddItem "1"
                            'DATA.
                            aValor = RsArt!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                            .Cell(flexcpData, .Rows - 1, 1) = TipoArticulo.PagoFlete
                            .Cell(flexcpData, .Rows - 1, 2) = 0     'Este campo me dice si es pto. el código del mismo.
                            aValor = RsArt!ArtTipo: .Cell(flexcpData, .Rows - 1, 3) = aValor    'Guardo el tipo de Articulo
                            aValor = RsAux!EnvCodigo: .Cell(flexcpData, .Rows - 1, 5) = aValor
                            .Cell(flexcpData, .Rows - 1, 6) = 0     'Artículos que están para envio.
                            
                            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsArt!ArtNombre)
                            .Cell(flexcpText, .Rows - 1, 2) = ""
                            
                            aValor = RsAux!EnvValorPiso
                            cIVA = IVAArticulo(RsArt!ArtID)
'                            If MnuOpFactSinIVA.Checked = False Then
                                .Cell(flexcpText, .Rows - 1, 4) = IVAArticulo(RsArt!ArtID)
 '                           Else
  '                              aValor = Format(aValor / (1 + (cIVA / 100)), "#,##0")
   '                             .Cell(flexcpText, .Rows - 1, 4) = 0
    '                        End If
                            .Cell(flexcpText, .Rows - 1, 3) = Format(aValor, "#,##0.00")
                            .Cell(flexcpText, .Rows - 1, 5) = Format(aValor, "#,##0.00")
                            .Cell(flexcpData, .Rows - 1, 4) = aValor       'Guardo el costo unitario.
                            
                            .Cell(flexcpText, .Rows - 1, 6) = "No"
                            
                            .Cell(flexcpText, .Rows - 1, 9) = RsArt!ArtCodigo
                            
                            cIVAAux = Format(CCur(.Cell(flexcpText, .Rows - 1, 5)) / CCur(1 + (CCur(.Cell(flexcpText, .Rows - 1, 4)) / 100)), "0.00")
                            cIVAAux = CCur(.Cell(flexcpText, .Rows - 1, 5)) - cIVAAux
                            
                            labIVA.Caption = Format(CCur(labIVA.Caption) + cIVAAux, "#,##0.00")
                            labTotal.Caption = Format(CCur(labTotal.Caption) + aValor, "#,##0.00")
                            labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
                            RsArt.Close
                        End If
                    End With
                End If
            End If
            RsAux.MoveNext
        Loop
        RsAux.Close
    Loop
    
End Sub

Private Function BuscoNombreMoneda(Codigo As Long) As String

    On Error GoTo ErrBU
    Dim rs As rdoResultset
    BuscoNombreMoneda = ""

    Cons = "SELECT * FROM Moneda WHERE MonCodigo = " & Codigo
    Set rs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rs.EOF Then BuscoNombreMoneda = Trim(rs!MonNombre)
    rs.Close
    Exit Function
    
ErrBU:
End Function

Private Sub CamposRenglon(Estado As Boolean)
Dim Color1 As Variant, Color2 As Variant
    tCantidad.Enabled = Estado
    tComentario.Enabled = Estado
    tUnitario.Enabled = Estado
    cEnvio.Enabled = Estado
    If Estado Then
        Color1 = Obligatorio: Color2 = Blanco
    Else
        Color1 = Inactivo: Color2 = Inactivo
    End If
    tCantidad.BackColor = Color1
    tComentario.BackColor = Color2
    tUnitario.BackColor = Color1
    cEnvio.BackColor = Color1
End Sub

Private Sub txtCliente_BorroCliente()
    LimpioDatosCliente
End Sub

Private Sub txtCliente_CambioTipoDocumento()
    SeteoInfoDocumentoCliente
End Sub

Private Sub txtCliente_Focus()
    Status.Panels(1).Text = "Ingrese el documento del cliente."
End Sub

Private Sub txtCliente_PresionoEnter()
'    If txtCliente.Cliente.Codigo > 0 Then
    If tArticulo.Enabled Then tArticulo.SetFocus
'    End If
End Sub

Private Sub txtCliente_SeleccionoCliente()
    LimpioDatosCliente
    If Not CargoDatosCliente Then
        LimpioDatosCliente
        Exit Sub
    End If
    If vsGrilla.Rows > 1 And vsGrilla.Enabled Then CambioPreciosEnLista oCliente.Categoria
    If strCodigoEnvio <> vbNullString Then
        If Not CambioClienteEnvios(oCliente.ID, strCodigoEnvio) Then MnuLimpiar_Click: Exit Sub
    End If
    txtCliente.BuscoComentariosAlerta txtCliente.Cliente.Codigo, True
    If txtCliente.DarMsgClienteNoVender(txtCliente.Cliente.Codigo) Then
        MsgBox "Atención: NO se puede vender sin autorización. Consultar con gerencia!", vbCritical, "ATENCIÓN"
    End If
    db_FindVtaPendiente oCliente.ID
    AvisoVentaLimitada False
    'Si tiene RUT le corro la validación por las dudas.
    ValidoRUT
End Sub

Private Sub txtTransaccion_GotFocus()
    With txtTransaccion
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub vsGrilla_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
'    If Not (Col = 0 Or Col = 2) Then Cancel = True: Exit Sub
    Cancel = True
End Sub

Private Sub vsGrilla_GotFocus()
    Status.Panels(1).Text = "Artículos a facturar. [+,-] Agrega y Quita, [Space] Edita, [E] Envía."
End Sub

Private Sub vsGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyDelete
            If vsGrilla.row < 1 Then Exit Sub
            'Si no es artículo o es de combo.
            If vsGrilla.Cell(flexcpData, vsGrilla.row, 1) = TipoArticulo.Presupuesto Or (vsGrilla.Cell(flexcpData, vsGrilla.row, 2) > 0 And vsGrilla.Cell(flexcpData, vsGrilla.row, 1) <> TipoArticulo.Especifico) Then Exit Sub
            If CInt(vsGrilla.Cell(flexcpData, vsGrilla.row, 6)) > 0 Then MsgBox "El artículo está asignado a envíos, acceda al mismo y eliminelos.", vbInformation, "ATENCIÓN": Exit Sub
            If CInt(vsGrilla.Cell(flexcpData, vsGrilla.row, 1)) = TipoArticulo.PagoFlete Then MsgBox "El artículo paga flete, si desea eleminarlo acceda al envío y cambie la forma de pago.", vbInformation, "ATENCIÓN": Exit Sub
                        
            RestoLabTotales CCur(vsGrilla.Cell(flexcpText, vsGrilla.row, 5)), CCur(vsGrilla.Cell(flexcpText, vsGrilla.row, 4))
            EliminarArrayFecha CLng(vsGrilla.Cell(flexcpData, vsGrilla.row, 0))
            vsGrilla.RemoveItem vsGrilla.row
            f_ValidateRetira False
            
        Case vbKeySubtract
            If vsGrilla.row < 1 Then Exit Sub
            If vsGrilla.Cell(flexcpData, vsGrilla.row, 1) <> TipoArticulo.Articulo Or vsGrilla.Cell(flexcpData, vsGrilla.row, 2) > 0 Then Exit Sub
            If CInt(vsGrilla.Cell(flexcpText, vsGrilla.row, 0)) - 1 = 0 Then Exit Sub
            If CInt(vsGrilla.Cell(flexcpText, vsGrilla.row, 0)) - 1 < CInt(vsGrilla.Cell(flexcpData, vsGrilla.row, 6)) Then Exit Sub
            RestoLabTotales CCur(vsGrilla.Cell(flexcpText, vsGrilla.row, 3)), CCur(vsGrilla.Cell(flexcpText, vsGrilla.row, 4))
            vsGrilla.Cell(flexcpText, vsGrilla.row, 0) = CInt(vsGrilla.Cell(flexcpText, vsGrilla.row, 0)) - 1
            vsGrilla.Cell(flexcpText, vsGrilla.row, 5) = Format(CCur(vsGrilla.Cell(flexcpText, vsGrilla.row, 0)) * CCur(vsGrilla.Cell(flexcpText, vsGrilla.row, 3)), "#,##0.00")
            f_ValidateRetira False
            ValidoVentaLimitadaPorFila vsGrilla.row
            
        Case vbKeyAdd
            If vsGrilla.row < 1 Then Exit Sub
            If vsGrilla.Cell(flexcpData, vsGrilla.row, 1) <> TipoArticulo.Articulo Or vsGrilla.Cell(flexcpData, vsGrilla.row, 2) > 0 Then Exit Sub
            labIVA.Caption = Format(CCur(labIVA.Caption) + CCur(vsGrilla.Cell(flexcpText, vsGrilla.row, 3)) - (CCur(vsGrilla.Cell(flexcpText, vsGrilla.row, 3)) / CCur(1 + (CCur(vsGrilla.Cell(flexcpText, vsGrilla.row, 4)) / 100))), "#,##0.00")
            labTotal.Caption = Format(CCur(labTotal.Caption) + CCur(vsGrilla.Cell(flexcpText, vsGrilla.row, 3)), "#,##0.00")
            labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
            vsGrilla.Cell(flexcpText, vsGrilla.row, 0) = CInt(vsGrilla.Cell(flexcpText, vsGrilla.row, 0)) + 1
            vsGrilla.Cell(flexcpText, vsGrilla.row, 5) = Format(CCur(vsGrilla.Cell(flexcpText, vsGrilla.row, 0)) * CCur(vsGrilla.Cell(flexcpText, vsGrilla.row, 3)), "#,##0.00")
            f_ValidateRetira False
            ValidoVentaLimitadaPorFila vsGrilla.row
            
        Case vbKeyReturn: Foco cPendiente
        
        Case vbKeySpace
            If CInt(vsGrilla.Cell(flexcpData, vsGrilla.row, 1)) = TipoArticulo.PagoFlete Or CInt(vsGrilla.Cell(flexcpData, vsGrilla.row, 1)) = TipoArticulo.Presupuesto _
                Or Val(vsGrilla.Cell(flexcpData, vsGrilla.row, 2)) > 0 Or CInt(vsGrilla.Cell(flexcpData, vsGrilla.row, 1)) = TipoArticulo.Especifico Then
                MsgBox "Este artículo no se puede editar.", vbInformation, "ATENCIÓN": Exit Sub
            End If
            
            'Si no es artículo o es de combo.
            With vsGrilla
                'Si tiene envíos
                If Val(.Cell(flexcpData, .row, 6)) > 0 Then MsgBox "Este artículo ya esta asignado en envíos no podrá editarlo.", vbExclamation, "ATENCION": Exit Sub
                InicializoVarRenglon
                    
                tArticulo.Text = .Cell(flexcpText, .row, 1)
                tArticulo.Tag = .Cell(flexcpData, .row, 0): miRenglon.IDArticulo = Trim(tArticulo.Tag)
                miRenglon.Tipo = .Cell(flexcpData, .row, 3)
                
                tCantidad.Text = .Cell(flexcpText, .row, 0)
'                tUnitario.Tag = .Cell(flexcpData, .Row, 4): miRenglon.Precio = Val(tUnitario.Tag)
                PrecioArticulo miRenglon.IDArticulo, cMoneda.ItemData(cMoneda.ListIndex), miRenglon.PrecioOriginal
                
                tUnitario.Text = miRenglon.PrecioOriginal
                tComentario.Text = .Cell(flexcpText, .row, 2)
                tUnitario.Text = .Cell(flexcpText, .row, 3)
                cEnvio.Text = .Cell(flexcpText, .row, 6)
                tComentario.Tag = .Cell(flexcpText, .row, 7)
                miRenglon.CantidadAlXMayor = Val(.Cell(flexcpData, .row, 8))
                miRenglon.NombreArticulo = .Cell(flexcpText, .row, 1)
                If Trim(.Cell(flexcpText, .row, 7)) <> "" Then miRenglon.EsInhabilitado = True
                If Val(.Cell(flexcpData, .row, 7)) > 0 Then miRenglon.CodArticulo = Val(.Cell(flexcpData, .row, 7))
                RestoLabTotales CCur(vsGrilla.Cell(flexcpText, vsGrilla.row, 5)), CCur(vsGrilla.Cell(flexcpText, vsGrilla.row, 4))
                .RemoveItem .row
                AplicoCantidadLimitadaPorCantidad
                Foco tCantidad
                
            End With
            
        Case vbKeyE
            If vsGrilla.Cell(flexcpData, vsGrilla.row, 1) = TipoArticulo.Articulo Or vsGrilla.Cell(flexcpData, vsGrilla.row, 1) = TipoArticulo.Especifico Then
                If vsGrilla.Cell(flexcpText, vsGrilla.row, 6) = "Si" Then
                    If CLng(vsGrilla.Cell(flexcpData, vsGrilla.row, 6)) = 0 Then vsGrilla.Cell(flexcpText, vsGrilla.row, 6) = "No"
                Else
                    vsGrilla.Cell(flexcpText, vsGrilla.row, 6) = "Si"
                End If
            End If
            
    End Select
    
End Sub

Private Function InvocoListaAyuda(Consulta As String, ByVal CerrarUnicaCoincidencia As Boolean) As Long
    InvocoListaAyuda = 0
    Dim objLista As New clsListadeAyuda
    objLista.CerrarSiEsUnico = CerrarUnicaCoincidencia
    If objLista.ActivarAyuda(cBase, Cons, 5200, 1, "Ayuda") Then InvocoListaAyuda = objLista.RetornoDatoSeleccionado(0)
    Me.Refresh
    Set objLista = Nothing
    Screen.MousePointer = 0
End Function

Private Sub BuscoServiciosCliente()
On Error GoTo ErrBSC
    
    If oCliente.ID > 0 Then
        Dim aQ As Long: aQ = 0
        Dim aIdServicioS As Long
        Cons = "Select * from Servicio, Producto, Taller  Where SerProducto = ProCodigo " & _
                    " And SerCliente = " & oCliente.ID & " And TalServicio = SerCodigo And SerEstadoServicio <>  " & EstadoS.Anulado _
                    & " And TalFReparado <> Null"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aQ = 1: aIdServicioS = RsAux!SerCodigo: RsAux.MoveNext
            If Not RsAux.EOF Then aQ = 2
        End If
        RsAux.Close
        
        Select Case aQ
            Case 0: MsgBox "No hay servicios pendientes para el cliente seleccionado.", vbExclamation, "No hay Servicios Pendientes"
            Case 1: CargoDatosServicio aIdServicioS
            Case 2:
                Cons = "Select SerCodigo as 'Servicio', SerFecha as 'Solicitud', ProCodigo as 'id_Prod.', ArtNombre as 'Producto', ProNroSerie as 'Nº Serie', ProCompra 'F/Compra', SerComentario 'Comentarios' " & _
                        " From Servicio, Producto, Articulo, Taller " & _
                        " Where SerProducto = ProCodigo and ProArticulo = ArtID" & _
                        " And SerCliente = " & oCliente.ID & _
                        " And SerEstadoServicio <> " & EstadoS.Anulado & _
                        " And TalFReparado <> Null And TalServicio = SerCodigo"
                
                Dim miLista As New clsListadeAyuda
                If miLista.ActivarAyuda(cBase, Cons, 8500, 0, "Ayuda de Servicios") Then
                    aIdServicioS = miLista.RetornoDatoSeleccionado(0)
                End If
                Me.Refresh
                Set miLista = Nothing
                If aIdServicioS <> 0 Then CargoDatosServicio aIdServicioS
        End Select
    End If
    Exit Sub

ErrBSC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrio un error al buscar los servicios del cliente."
End Sub

Private Sub CargoDatosServicio(ByVal idServicio As Long)
On Error GoTo ErrCDS
Dim cCostoRep As Currency
    Screen.MousePointer = 11
    
    vsGrilla.Rows = 1: LabTotalesEnCero
    Cons = "Select * From Servicio, Moneda, Producto " _
        & " Where SerCodigo = " & idServicio & " And SerDocumento IS Null " _
        & " And SerProducto = ProCodigo And SerMoneda = MonCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    If RsAux.EOF Then
        MsgBox "No se encontraron datos para ese servicio o el mismo ya fue facturado.", vbExclamation, "ATENCIÓN"
    Else
        Dim sFCosteo As String
        sFCosteo = UltimaFCosteo
        If Not IsNull(RsAux!SerFCumplido) Then
            If CDate(UltimoDia(CDate(sFCosteo))) + 1 > RsAux!SerFCumplido Then
                MsgBox "No se podrá facturar este servicio ya que los artículos fueron incluidos en el costeo correspondiente a la fecha de cumplido del servicio.", vbExclamation, "ATENCIÓN"
                RsAux.Close
                Screen.MousePointer = 0: Exit Sub
            End If
        End If
        
        If MnuOpFactSinIVA.Checked = True Then
            MnuOpFactSinIVA.Checked = False
            BackColor = &HC0E0FF
            Shape2.BackColor = &H80C0FF
            chPagaCheque.BackColor = BackColor
            chNomDireccion.BackColor = Shape2.BackColor
        End If
        
        BuscoCodigoEnCombo cMoneda, RsAux!MonCodigo
        Dim bChange As Boolean
        bChange = True
        If oCliente.ID > 0 And RsAux("SerCliente") <> oCliente.ID Then
            If MsgBox("Este servicio no es de " & tNombreC.Text & vbCrLf & vbCrLf & "¿Lo desea facturar a ese nombre?", vbQuestion + vbYesNo, "OTRO CLIENTE") = vbYes Then
                bChange = False
            End If
        End If
        If bChange Then
            txtCliente.Text = ""
            LimpioDatosCliente
            txtCliente.CargarControl RsAux("SerCliente")
        End If
        
        cCostoRep = CargoRenglonesServicio(RsAux!SerCodigo)
        'Ahora en base al valor del costo final verifico el total que me dio los artículos.
        If CCur(labTotal.Caption) <> RsAux!SerCostoFinal Then
            'Como tengo diferencia en la suma de los artículos con el costo final inserto artículo que me da el costo final.
            CargoArticuloCobroServicio RsAux!SerCostoFinal - cCostoRep, RsAux!SerCodigo
        End If
        If vsGrilla.Rows > 1 Then
            CamposRenglon False
            OcultoPorServicio
            Foco cPendiente
        End If
    End If
    RsAux.Close
    Screen.MousePointer = 0: Exit Sub
ErrCDS:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar el servicio del cliente.", Err.Description
End Sub

Private Sub CargoArticuloCobroServicio(ByVal CostoServicio As Currency, idServicio As Long)
Dim RsSR As rdoResultset
Dim aValor As Currency

    If CostoServicio = 0 Then Exit Sub
    
    Cons = "Select * From Articulo Where ArtID = " & paArticuloCobroServicio
    Set RsSR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsSR.EOF Then
        
        With vsGrilla
            .AddItem "1"
            'DATA.
            aValor = RsSR!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
            .Cell(flexcpData, .Rows - 1, 1) = TipoArticulo.Servicio
            .Cell(flexcpData, .Rows - 1, 2) = 0     'Este campo me dice si es pto. el código del mismo.
            aValor = RsSR!ArtTipo: .Cell(flexcpData, .Rows - 1, 3) = aValor    'Guardo el tipo de Articulo
            
            .Cell(flexcpData, .Rows - 1, 5) = idServicio
            .Cell(flexcpData, .Rows - 1, 6) = 0     'Artículos que están para envio.
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsSR!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 2) = ""
            aValor = IVAArticulo(RsSR!ArtID)

            If MnuOpFactSinIVA.Checked = False Then
                .Cell(flexcpText, .Rows - 1, 4) = aValor
            Else
                .Cell(flexcpText, .Rows - 1, 4) = 0
                CostoServicio = Format(CostoServicio / (1 + (aValor / 100)), "#,##0")
            End If
            
            .Cell(flexcpData, .Rows - 1, 4) = CostoServicio      'Guardo el costo unitario.
            .Cell(flexcpText, .Rows - 1, 3) = Format(CostoServicio, "#,##0.00")
            
            .Cell(flexcpText, .Rows - 1, 5) = Format(CostoServicio, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 6) = "No"
            
            .Cell(flexcpText, .Rows - 1, 9) = RsSR("ArtCodigo")
                
            labIVA.Caption = Format(CCur(labIVA.Caption) + CCur(.Cell(flexcpText, .Rows - 1, 5)) - (CCur(.Cell(flexcpText, .Rows - 1, 5)) / CCur(1 + (CCur(.Cell(flexcpText, .Rows - 1, 4)) / 100))), "#,##0.00")
            labTotal.Caption = Format(CCur(labTotal.Caption) + CostoServicio, "#,##0.00")
            labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
        End With
    End If
    RsSR.Close
End Sub
Private Function CargoRenglonesServicio(idServicio As Long) As Currency
Dim RsSR As rdoResultset
Dim aValor As Currency
Dim cIVA As Currency, cRepImp As Currency
    
    cRepImp = 0
    Cons = "Select * From ServicioRenglon, Articulo " _
        & " Where SReServicio = " & idServicio _
        & " And SReTipoRenglon = " & TipoRenglonS.Cumplido & " And SReMotivo = ArtID"
    Set RsSR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    Do While Not RsSR.EOF
        With vsGrilla
            .AddItem RsSR!SReCantidad
            'DATA.
            aValor = RsSR!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
            .Cell(flexcpData, .Rows - 1, 1) = TipoArticulo.Servicio
            .Cell(flexcpData, .Rows - 1, 2) = 0     'Este campo me dice si es pto. el código del mismo.
            aValor = RsSR!ArtTipo: .Cell(flexcpData, .Rows - 1, 3) = aValor    'Guardo el tipo de Articulo
            
            .Cell(flexcpData, .Rows - 1, 5) = idServicio
            .Cell(flexcpData, .Rows - 1, 6) = 0     'Artículos que están para envio.
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsSR!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 2) = ""
            
            cIVA = IVAArticulo(RsSR!ArtID)
            aValor = Format(RsSR!SReTotal, FormatoMonedaP)
            cRepImp = (aValor * RsSR!SReCantidad) + cRepImp
            
            If MnuOpFactSinIVA.Checked = False Then
                .Cell(flexcpText, .Rows - 1, 4) = Format(cIVA, FormatoMonedaP)
            Else
                cIVA = 1 + (cIVA / 100)
                .Cell(flexcpText, .Rows - 1, 4) = "0"
                aValor = Format(aValor / cIVA, "#,##0")
            End If
            
            .Cell(flexcpData, .Rows - 1, 4) = aValor       'Guardo el costo unitario.
            .Cell(flexcpText, .Rows - 1, 3) = Format(aValor, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 5) = Format(aValor * RsSR!SReCantidad, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 6) = "No"
            
            .Cell(flexcpText, .Rows - 1, 9) = RsSR("ArtCodigo")
            
            labIVA.Caption = Format(CCur(labIVA.Caption) + CCur(.Cell(flexcpText, .Rows - 1, 5)) - (CCur(.Cell(flexcpText, .Rows - 1, 5)) / CCur(1 + (CCur(.Cell(flexcpText, .Rows - 1, 4)) / 100))), "#,##0.00")
            labTotal.Caption = Format(CCur(labTotal.Caption) + CCur(aValor * RsSR!SReCantidad), "#,##0.00")
        End With
        labSubTotal.Caption = Format(CCur(labTotal.Caption) - CCur(labIVA.Caption), "#,##0.00")
        RsSR.MoveNext
    Loop
    RsSR.Close
    CargoRenglonesServicio = cRepImp
    
End Function

Private Sub OcultoPorServicio()
    vsGrilla.Enabled = False: tArticulo.Enabled = False: tArticulo.BackColor = Colores.Inactivo: tArticulo.Text = ""
    cMoneda.Enabled = False
End Sub
Private Sub MuestroPorServicio()
    vsGrilla.Enabled = True: tArticulo.Enabled = True: tArticulo.BackColor = Colores.Obligatorio
    cMoneda.Enabled = True
End Sub

Private Function NumeroAuxiliarEnvio() As Integer
Dim idAux As Integer
    
    NumeroAuxiliarEnvio = 0
    On Error GoTo ErrBT
    
    idAux = Autonumerico(TAutonumerico.AuxiliarEnvio)
    cBase.BeginTrans
    
    If idAux > 10000 And idAux < 10035 Then
        Cons = "Delete EnvioAuxiliar Where EAuID >= 15000"
        cBase.Execute (Cons)
    ElseIf idAux > 20000 And idAux < 20035 Then
        Cons = "Delete EnvioAuxiliar Where EAuID < 15000"
        cBase.Execute (Cons)
    End If
    
    'Para que no exista duplicación.
    Cons = "Delete EnvioAuxiliar Where EAuID = " & idAux
    cBase.Execute (Cons)
    '---------------------------------------------------------------------------
    On Error GoTo ErrRB
    For I = 1 To vsGrilla.Rows - 1
        If Trim(vsGrilla.Cell(flexcpText, I, 6)) = "Si" And (Val(vsGrilla.Cell(flexcpData, I, 1)) = TipoArticulo.Articulo Or Val(vsGrilla.Cell(flexcpData, I, 1)) = TipoArticulo.Especifico) _
            And Not EsTipoDeServicio(Val(vsGrilla.Cell(flexcpData, I, 3))) Then
            Cons = "Insert into EnvioAuxiliar (EAuID, EAuArticulo, EAuCantidad) Values (" & idAux & ", " & vsGrilla.Cell(flexcpData, I, 0) & ", " & vsGrilla.Cell(flexcpText, I, 0) & ")"
            cBase.Execute (Cons)
        End If
    Next I
    cBase.CommitTrans
    NumeroAuxiliarEnvio = idAux
    Exit Function

ErrBT:
    clsGeneral.OcurrioError "Error al intentar abrir la transacción.", Err.Description
    Screen.MousePointer = 0
ErrRB:
    Resume ErrResumo
ErrResumo:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al insertar los artículos para enviar.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub CargoDireccionesAuxiliares(aIdCliente As Long)
On Error GoTo errCDA
    Dim lDPcpal As Long, sNFactura As String
    Dim rsDA As rdoResultset
    
    If gDirFactura > 0 Then lDPcpal = gDirFactura: gDirFactura = 0
    
    'Direcciones Auxiliares-----------------------------------------------------------------------
    Cons = "Select Top 16 * from DireccionAuxiliar Where DAuCliente = " & aIdCliente & _
            " Order by DAuFactura Desc, DAuNombre "
    
    Set rsDA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsDA.EOF Then
        Do While Not rsDA.EOF
                        
            With cDireccion
                .AddItem Trim(rsDA!DAuNombre)
                .ItemData(.NewIndex) = rsDA!DAuDireccion
            End With
            
            'Si es la seleccionada para facturar.
            If rsDA!DAuFactura Then
                gDirFactura = rsDA!DAuDireccion
                sNFactura = Trim(rsDA!DAuNombre)
            End If
            
            rsDA.MoveNext
        Loop
    End If
    rsDA.Close
    
    With cDireccion
        If .ListCount > 1 Then cDireccion.BackColor = Colores.Blanco
        
        If .ListCount > IIf(lDPcpal > 0, 16, 15) Then
            'Cumple Top --> borro pongo nuevamente la dirección del cliente y la opción buscar.
            .Clear
            
            If lDPcpal > 0 Then
                .AddItem "Dirección Principal": .ItemData(.NewIndex) = lDPcpal
                .Tag = lDPcpal
            End If
                        
            If gDirFactura > 0 Then
                .AddItem sNFactura
                .ItemData(.NewIndex) = gDirFactura
            End If
            
            'Para buscar.
            .AddItem cte_KeyFindDir
            .ItemData(.NewIndex) = -1
            
        End If
        
        If gDirFactura = 0 And lDPcpal > 0 Then gDirFactura = lDPcpal

        If gDirFactura <> 0 Then
            BuscoCodigoEnCombo cDireccion, gDirFactura
        Else
            If cDireccion.ListCount > 0 Then
                cDireccion.ListIndex = 0
            Else
                labDireccion.Caption = "  Sin Dirección"
                LabelMensaje labDireccion, True
            End If
       End If
    End With
  
errCDA:
End Sub

Private Function BuscoZonaDireccion(lngIDDir As Long) As Long
Dim RsZona As rdoResultset, strCons As String
    
    BuscoZonaDireccion = 0
    
    strCons = "Select CZoZona from Direccion, CalleZona" _
            & " Where DirCodigo = " & lngIDDir _
            & " And CZoCalle = DirCalle " _
            & " And CZoDesde <= DirPuerta " & " And CZoHasta >= DirPuerta"
    Set RsZona = cBase.OpenResultset(strCons, rdOpenDynamic, rdConcurValues)
    
    If Not RsZona.EOF Then
        If Not IsNull(RsZona(0)) Then BuscoZonaDireccion = RsZona(0)
    End If
    RsZona.Close
    
End Function

Private Function UltimaFCosteo() As String
Dim rsUF As rdoResultset
On Error GoTo errUF
    UltimaFCosteo = "01/01/1990"
    Cons = "Select Max(CabMesCosteo) From CMCabezal"
    Set rsUF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsUF.EOF Then
        If Not IsNull(rsUF(0)) Then UltimaFCosteo = rsUF(0)
    End If
    rsUF.Close
    Exit Function
errUF:
    clsGeneral.OcurrioError "Ocurrió un error al buscar la mayor fecha de costeo.", Trim(Err.Description)
End Function


Private Function RetornoCodigoArticulo(ByVal IDArt As Long) As Long
On Error GoTo errRCA
    Cons = "Select * From Articulo Where ArtID = " & IDArt
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RetornoCodigoArticulo = RsAux!ArtCodigo
    RsAux.Close
    Exit Function
errRCA:
End Function

Private Function Query_CabezalBusquedaArticuloNormal() As String
    Query_CabezalBusquedaArticuloNormal = "SELECT ArtID, ArtCodigo, ArtEnUso " _
        & ", IsNull(ArtEnVentaXMayor, 1) ArtEnVentaXMayor, CASE WHEN ArtHabilitado='S' THEN 1 ELSE 0 END Habilitado " _
        & ", ArtNombre, ArtTipo, ArtEsCombo, ArtDemoraEntrega, ArtDisponibleDesde " _
        & "FROM Articulo "
End Function

Private Sub CargoArticulosNormales()
On Error GoTo errCAN
Dim bPH As Boolean, idTipo As Long, cAuxPrecio As Currency

    Screen.MousePointer = 11
    InicializoVarRenglon
    
    If Not IsNumeric(tArticulo.Text) Then
        
        Cons = "SELECT dbo.ocxarticulo('VtaCdo', '" & tArticulo.Text & "')"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            Cons = RsAux(0)
        Else
            Cons = ""
        End If
        RsAux.Close
        Dim IDArticulo As Long: IDArticulo = 0
        If Cons <> "" Then
            'miRenglon.CodArticulo = InvocoListaAyuda(Cons)
            IDArticulo = InvocoListaAyuda(Cons, True)
        End If
        If IDArticulo = 0 Then Exit Sub
        'Cons = "Select * From Articulo Where ArtID = " & IDArticulo
        Cons = Query_CabezalBusquedaArticuloNormal & " WHERE ArtID = " & IDArticulo
    Else
        miRenglon.CodArticulo = tArticulo.Text
        'Cons = "Select * From Articulo Where ArtCodigo = " & miRenglon.CodArticulo
        Cons = Query_CabezalBusquedaArticuloNormal & " WHERE ArtCodigo = " & miRenglon.CodArticulo
    End If
    
    'Busco el Artículo por código.
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        Screen.MousePointer = 0
        MsgBox "No existe un artículo con ese código.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    miRenglon.CodArticulo = RsAux("ArtCodigo")
    miRenglon.CantidadAlXMayor = 0
    
    If Not IsNull(RsAux("ArtEnVentaXMayor")) Then
        miRenglon.CantidadAlXMayor = RsAux("ArtEnVentaXMayor")
    Else
        miRenglon.VentaXMayor = 1
    End If
    
    'Veo si esta ingresado.
    If Ingresado(RsAux!ArtID) Then
        RsAux.Close: InicializoVarRenglon
        MsgBox "El artículo seleccionado ya esta ingresado.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If RsAux!ArtEnUso = 0 Then
        snd_ActivarSonido Replace(gPathListados, "\reportes\", "\sonidos\", , , vbTextCompare) & "artfuerauso.wav"
        If MsgBox("El artículo ingresado no está en uso." & vbCrLf & "¿Desea facturarlo de todas formas?", vbQuestion + vbYesNo + vbDefaultButton2, "Artículo fuera de uso") = vbYes Then
            miRenglon.EsInhabilitado = True
        Else
            RsAux.Close
            GoTo evInhabilitado
        End If
    Else
        miRenglon.EsInhabilitado = (RsAux("Habilitado") = 0)
        'Si ingreso por código y no está habilitado le consulto.
        'If (UCase(RsAux!ArtHabilitado) <> "S" Or IsNull(RsAux!ArtHabilitado)) Then
        If miRenglon.EsInhabilitado Then
            snd_ActivarSonido Replace(gPathListados, "\reportes\", "\sonidos\", , , vbTextCompare) & "artnohabilitado.wav"
            If MsgBox("El artículo ingresado no está habilitado para la venta." & vbCrLf & "¿Desea facturarlo de todas formas?", vbQuestion + vbYesNo + vbDefaultButton2, "Artículo no habilitado") = vbNo Then
                RsAux.Close
                GoTo evInhabilitado
            End If
        End If
    End If

    miRenglon.NombreArticulo = Trim(RsAux("ArtNombre"))
    miRenglon.Tipo = RsAux!ArtTipo
    miRenglon.IDArticulo = RsAux!ArtID
    AplicoTextoDeVentaLimitada

'DEMORA y Disponibilidad
    If Not IsNull(RsAux("ArtDisponibleDesde")) Or Not IsNull(RsAux("ArtDemoraEntrega")) Then
        miRenglon.DisponibleDesde = Date
    Else
        miRenglon.DisponibleDesde = DateSerial(2000, 1, 1)
    End If
    If Not IsNull(RsAux("ArtDisponibleDesde")) Then
        If RsAux("ArtDisponibleDesde") > Date Then
            miRenglon.DisponibleDesde = RsAux("ArtDisponibleDesde")
        End If
    End If
    If Not IsNull(RsAux("ArtDemoraEntrega")) Then
        miRenglon.DisponibleDesde = DateAdd("d", RsAux("ArtDemoraEntrega"), miRenglon.DisponibleDesde)
    End If
    
    If RsAux!ArtEsCombo Then
        RsAux.Close
        'Busco en tabla Presupuesto el artículo que tiene este artículo.
        Cons = "Select PreID, PreArticulo, PreImporte From Presupuesto Where PreArtCombo = " & miRenglon.IDArticulo _
            & " And PreMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
            
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            cAuxPrecio = 0
            miRenglon.IDCombo = RsAux!PreID
            miRenglon.ArtCombo = RsAux!PreArticulo
            If MnuOpFactSinIVA.Checked Then
                If RsAux!PreImporte <> 0 Then
                    cAuxPrecio = 1 + (IVAArticulo(RsAux!PreArticulo) / 100)
                    cAuxPrecio = Format(RsAux!PreImporte / cAuxPrecio, "#,##0")
                End If
            Else
                If RsAux!PreImporte <> 0 Then
                    m_Patron = dis_arrMonedaProp(cMoneda.ItemData(cMoneda.ListIndex), pRedondeo)
                    cAuxPrecio = Format(Redondeo(RsAux!PreImporte, m_Patron), FormatoMonedaP)
                End If
            End If
            miRenglon.PrecioBonificacion = cAuxPrecio
                        
        Else
            
            RsAux.Close
            MsgBox "No existe un combo ingresado con la moneda seleccionada.", vbExclamation, "ATENCIÓN"
            InicializoVarRenglon
            Screen.MousePointer = 0
            Exit Sub
        End If
        RsAux.Close
        
    Else
        RsAux.Close
        
        If Not PrecioArticulo(miRenglon.IDArticulo, cMoneda.ItemData(cMoneda.ListIndex), miRenglon.Precio) Then
            MsgBox "El artículo seleccionado no posee precios ingresados para la moneda seleccionada.", vbInformation, "ATENCIÓN"
        Else
            miRenglon.PrecioOriginal = miRenglon.Precio
        End If
        
    End If
    Screen.MousePointer = 0
    Exit Sub
    
evInhabilitado:
    InicializoVarRenglon
    Screen.MousePointer = 0
    Exit Sub
    
errCAN:
    clsGeneral.OcurrioError "Ocurrió el siguiente error al buscar el artículo.", Err.Description, "Contado"
    Screen.MousePointer = 0
End Sub

Private Sub AplicoTextoDeVentaLimitada()
    
    Dim codAux As Long
    codAux = miRenglon.IDArticulo
    If IsNumeric(tCantidad.Text) Then
    
        If miRenglon.CantidadAlXMayor = 0 And (InStr(1, paCategoriaDistribuidor, "," & oCliente.Categoria & ",") > 0 Or oCliente.Categoria = 0) Then
            tArticulo.Text = miRenglon.NombreArticulo & " (no vta. a Distr.)"
        ElseIf miRenglon.CantidadAlXMayor > 1 And miRenglon.CantidadAlXMayor < Val(tCantidad.Text) Then
            tArticulo.Text = miRenglon.NombreArticulo & " (limitado a " & miRenglon.CantidadAlXMayor & ")"
        Else
            tArticulo.Text = miRenglon.NombreArticulo
        End If
    Else
        'Defino el nombre en base a la disponibilidad de venta.
        If miRenglon.CantidadAlXMayor = 0 And (InStr(1, paCategoriaDistribuidor, "," & oCliente.Categoria & ",") > 0 Or oCliente.Categoria = 0) Then
            tArticulo.Text = miRenglon.NombreArticulo & " (no vta. a Distr.)"
        ElseIf miRenglon.CantidadAlXMayor > 1 Then
            tArticulo.Text = miRenglon.NombreArticulo & " (limitado a " & miRenglon.CantidadAlXMayor & ")"
        Else
            tArticulo.Text = miRenglon.NombreArticulo
        End If
    End If
    If Not tmArticuloLimitado.Enabled Then
        tmArticuloLimitado.Enabled = (tArticulo.Text <> miRenglon.NombreArticulo)
        If Not tmArticuloLimitado.Enabled Then
            tArticulo.ForeColor = vbBlack
            tCantidad.ForeColor = vbBlack
        End If
    End If
    miRenglon.IDArticulo = codAux
    
End Sub

Private Function PrecioArticulo(ByVal lArticulo As Long, ByVal idMoneda As Long, cPrecio As Currency) As Boolean
On Error GoTo errPA
Dim rsPrecio As rdoResultset
Dim cTC As Currency

    PrecioArticulo = False
    cPrecio = 0
    Cons = "Select * From PrecioVigente Where PViArticulo = " & lArticulo _
        & " And PViMoneda = " & idMoneda & " And PViTipoCuota = " & paTipoCuotaContado
    Set rsPrecio = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsPrecio.EOF Then
        If rsPrecio!PViHabilitado Then cPrecio = rsPrecio!PViPrecio: PrecioArticulo = True
    End If
    rsPrecio.Close
    
    If PrecioArticulo Then
        If MnuOpFactSinIVA.Checked Then
            'Le resto el iva.
            cTC = 1 + (IVAArticulo(lArticulo) / 100)
            cPrecio = cPrecio / cTC
        End If
    End If
    m_Patron = dis_arrMonedaProp(idMoneda, pRedondeo)
    cPrecio = Redondeo(cPrecio, m_Patron)
    
    If PrecioArticulo Or cMoneda.ItemData(cMoneda.ListIndex) <> paMonedaPesos Then Exit Function
    
    'Si la moneda es pesos y no tengo precio, busco el precio en dolares.
    Cons = "Select * From PrecioVigente Where PViArticulo = " & lArticulo _
        & " And PViMoneda = " & paMonedaDolar & " And PViTipoCuota = " & paTipoCuotaContado
    Set rsPrecio = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsPrecio.EOF Then
        If rsPrecio!PViHabilitado Then
            cPrecio = rsPrecio!PViPrecio
            cTC = TasadeCambio(paMonedaDolar, cMoneda.ItemData(cMoneda.ListIndex), Date)
            cPrecio = cPrecio * cTC
        End If
        PrecioArticulo = True
    End If
    rsPrecio.Close
    
    If PrecioArticulo Then
        If MnuOpFactSinIVA.Checked Then
            'Le resto el iva.
            cTC = 1 + (IVAArticulo(lArticulo) / 100)
            cPrecio = cPrecio / cTC
        End If
    End If
    
    m_Patron = dis_arrMonedaProp(idMoneda, pRedondeo)
    cPrecio = Redondeo(cPrecio, m_Patron)
    Exit Function
    
errPA:
    clsGeneral.OcurrioError "Ocurrió el siguiente error al buscar el precio vigente del artículo con ID: " & lArticulo & ".", Err.Description
End Function

Private Sub InicializoVarRenglon()
    
    miRenglon.PrecioBonificacion = 0
    miRenglon.ArtCombo = 0
    miRenglon.CodArticulo = 0
    miRenglon.IDCombo = 0
    miRenglon.EsInhabilitado = False
    miRenglon.IDArticulo = 0
    miRenglon.Precio = 0
    miRenglon.Tipo = 0
    miRenglon.PrecioOriginal = 0
    miRenglon.NombreArticulo = ""
    
End Sub

Private Sub PresentoPrecio()
On Error GoTo errPP
Dim cAuxPrecio As Currency, cSumaPrecio As Currency
    
    cSumaPrecio = 0
    
    If miRenglon.ArtCombo = 0 Then
        cSumaPrecio = BuscoDescuentoCliente(miRenglon.IDArticulo, oCliente.Categoria, miRenglon.PrecioOriginal, Val(tCantidad.Text))
        
    Else
        cSumaPrecio = BuscoDescuentoCliente(miRenglon.ArtCombo, oCliente.Categoria, miRenglon.PrecioBonificacion, Val(tCantidad.Text))

        Cons = "Select * From PresupuestoArticulo Where PArPresupuesto = " & miRenglon.IDCombo
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            cAuxPrecio = 0
            If PrecioArticulo(RsAux!PArArticulo, cMoneda.ItemData(cMoneda.ListIndex), cAuxPrecio) Then
                cAuxPrecio = BuscoDescuentoCliente(RsAux!PArArticulo, oCliente.Categoria, cAuxPrecio, RsAux!ParCantidad * Val(tCantidad.Text))
                'Si la cantidad es mayor a 1 incremento el precio
                cAuxPrecio = cAuxPrecio * RsAux!ParCantidad
            Else
                MsgBox "No existe precio vigente para el artículo del combo con id " & RsAux!PArArticulo & ".", vbInformation, "ATENCIÓN"
            End If
            cSumaPrecio = cSumaPrecio + CCur(cAuxPrecio)
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    tUnitario.Text = Format(CCur(cSumaPrecio), "#,##0.00")
    Exit Sub
    
errPP:
    clsGeneral.OcurrioError "Ocurrió al calcular el precio unitario.", Err.Description, "Error en Presentar Precio"
End Sub

Private Sub InsertoArticulosCombo(ByVal varPEspecifico As Currency, ByVal idEspecifico As Long)
On Error GoTo errIAC
Dim cAuxPrecio As Currency, idAux As Long
Dim cComboParcial As Currency
    
    If Ingresado(miRenglon.ArtCombo) Then
        MsgBox "El artículo bonificación del combo ya esta ingresado, no podrá ingresar el combo.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    'Valido que no este ingresado algún artículo del combo.
    Cons = "Select * From PresupuestoArticulo, Articulo Where PArPresupuesto = " & miRenglon.IDCombo & " And PArArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        If Ingresado(RsAux!PArArticulo) Then
            MsgBox "El artículo " & Trim(RsAux!ArtNombre) & " ya está ingresado, no podrá ingresar el combo.", vbExclamation, "ATENCIÓN"
            RsAux.Close
            Exit Sub
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    cComboParcial = 0
    
    Cons = "Select * From Articulo, PresupuestoArticulo Where PArPresupuesto = " & miRenglon.IDCombo _
        & " And PArArticulo = ArtID "
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        
        cAuxPrecio = 0
        PrecioArticulo RsAux!ArtID, cMoneda.ItemData(cMoneda.ListIndex), cAuxPrecio
        
        cAuxPrecio = BuscoDescuentoCliente(RsAux!ArtID, oCliente.Categoria, cAuxPrecio, Val(tCantidad.Text) * RsAux!ParCantidad)
        
        cComboParcial = cComboParcial + (CCur(cAuxPrecio) * RsAux!ParCantidad)
        
        CargoArticuloEnGrilla RsAux!ArtID, RsAux!ArtTipo, CInt(tCantidad.Text) * RsAux!ParCantidad, IIf(idEspecifico > 0, Especifico, Articulo), miRenglon.Precio, RsAux!ArtNombre, "", cAuxPrecio, cEnvio.Text, False, RsAux!ArtCodigo, idEspecifico
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
        
    'Si el articulo bonificación tiene costo <> 0 ---> lo inserto.
    Cons = "Select * From Articulo Where ArtID = " & miRenglon.ArtCombo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    miRenglon.PrecioBonificacion = BuscoDescuentoCliente(miRenglon.ArtCombo, oCliente.Categoria, miRenglon.PrecioBonificacion, Val(tCantidad.Text))
    
    If tUnitario.Text <> "" Then
        If CCur(tUnitario.Text) - (cComboParcial + miRenglon.PrecioBonificacion) <> 0 Then
            cAuxPrecio = miRenglon.PrecioBonificacion
            miRenglon.PrecioBonificacion = CCur(tUnitario.Text) - cComboParcial
        Else
            cAuxPrecio = miRenglon.PrecioBonificacion
        End If
    Else
        cAuxPrecio = 0
    End If
    
    If miRenglon.PrecioBonificacion + varPEspecifico <> 0 Then
        CargoArticuloEnGrilla miRenglon.ArtCombo, RsAux!ArtTipo, tCantidad.Text, Bonificacion, cAuxPrecio, RsAux!ArtNombre, tComentario.Text, miRenglon.PrecioBonificacion + varPEspecifico, "No", miRenglon.EsInhabilitado, miRenglon.CodArticulo
    End If
    RsAux.Close
    
    Exit Sub
    
errIAC:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los artículos del combo.", Err.Description
End Sub

Private Sub CargoArticuloEnGrilla(ByVal IDArticulo As Long, ByVal idTipoArticulo As Long, ByVal Cantidad As Integer _
                                            , ByVal TipoArticuloRenglon As TipoArticulo, ByVal PrecioUnitarioReal As Currency _
                                            , ByVal sNombreArticulo As String, ByVal sComentario As String, ByVal cUnitarioVenta As Currency _
                                            , ByVal sEnvio As String, ByVal bInhabilitado As Boolean, ByVal lCodArticulo As Long, Optional iIDEspecifico As Long = 0)

    With vsGrilla
        .AddItem Cantidad
        'DATA.
        .Cell(flexcpData, .Rows - 1, 0) = IDArticulo
        .Cell(flexcpData, .Rows - 1, 1) = TipoArticuloRenglon
        
        'Agregué esto por las ventas redpagos
        If (TipoArticuloRenglon = PagoFlete) Then
            .Cell(flexcpData, .Rows - 1, 5) = strCodigoEnvio
        End If
        
        .Cell(flexcpData, .Rows - 1, 2) = iIDEspecifico          'Este campo me dice si es pto. el código del mismo ó para el caso especifico el id del mismo.
        .Cell(flexcpData, .Rows - 1, 3) = idTipoArticulo      'Guardo el tipo de Articulo
        .Cell(flexcpData, .Rows - 1, 4) = PrecioUnitarioReal
        .Cell(flexcpData, .Rows - 1, 8) = miRenglon.CantidadAlXMayor
        '----------------------------------------------------------------------------
        'ATENCIÒN
        'En el tag 5 guardo si es un artículo que paga flete.
        '----------------------------------------------------------------------------
        .Cell(flexcpData, .Rows - 1, 6) = 0     'Artículos que están para envio.
        
        .Cell(flexcpText, .Rows - 1, 1) = Trim(sNombreArticulo)
        .Cell(flexcpText, .Rows - 1, 2) = Trim(sComentario)
        .Cell(flexcpText, .Rows - 1, 3) = Format(cUnitarioVenta, FormatoMonedaP)
        If MnuOpFactSinIVA.Checked = False Then
            .Cell(flexcpText, .Rows - 1, 4) = Format(IVAArticulo(IDArticulo), FormatoMonedaP)
        Else
            .Cell(flexcpText, .Rows - 1, 4) = "0.00"
        End If
        .Cell(flexcpText, .Rows - 1, 5) = Format(cUnitarioVenta * Cantidad, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 6) = sEnvio
        
        If bInhabilitado Then
            .Cell(flexcpText, .Rows - 1, 7) = "1"
            .Cell(flexcpData, .Rows - 1, 7) = lCodArticulo
        End If
        
        .Cell(flexcpText, .Rows - 1, 8) = GetInstaladorArt(IDArticulo)
        .Cell(flexcpText, .Rows - 1, 9) = lCodArticulo
        
        labIVA.Caption = Format(CCur(labIVA.Caption) + Format((CCur(.Cell(flexcpText, .Rows - 1, 5)) - (CCur(.Cell(flexcpText, .Rows - 1, 5)) / CCur(1 + (CCur(.Cell(flexcpText, .Rows - 1, 4)) / 100)))), "0.00"), FormatoMonedaP)
        labTotal.Caption = Format(CCur(labTotal.Caption) + CCur(cUnitarioVenta * Cantidad), FormatoMonedaP)
        
        ValidoVentaLimitadaPorFila .Rows - 1
    End With

End Sub

Private Sub ValidoVentaLimitadaPorFila(ByVal I As Long)

    With vsGrilla
'        .Cell(flexcpFontStrikethru, I, 0) = False
'        .Cell(flexcpFontStrikethru, I, 1) = False
        .Cell(flexcpForeColor, I, 0) = vbBlack
        'If Val(.Cell(flexcpData, I, 3)) <> paTipoArticuloServicio Then
        If Not EsTipoDeServicio(CLng(.Cell(flexcpData, I, 3))) Then
            If Val(.Cell(flexcpData, I, 8)) = 0 And InStr(1, paCategoriaDistribuidor, "," & oCliente.Categoria & ",") > 0 Then
'                .Cell(flexcpFontStrikethru, I, 0) = True
'                .Cell(flexcpFontStrikethru, I, 1) = True
                .Cell(flexcpForeColor, I, 0) = &HFF&
            Else
                If Val(.Cell(flexcpData, I, 8)) > 1 And Val(.Cell(flexcpData, I, 8)) < CInt(.Cell(flexcpText, I, 0)) Then
'                    .Cell(flexcpFontStrikethru, I, 0) = True
'                    .Cell(flexcpFontStrikethru, I, 1) = True
                    .Cell(flexcpForeColor, I, 0) = &HFF&
                End If
            End If
        End If
    End With
    
End Sub

Private Sub LabelMensaje(cLabel As Control, ByVal bResalto As Boolean)

    If bResalto Then
        With cLabel
            If .ForeColor <> vbRojoFuerte Then
                .ForeColor = vbRojoFuerte
                .BackColor = vbWindowBackground
            Else
                .ForeColor = vbWhite
                .BackColor = vbRojoFuerte
            End If
            .FontBold = True
        End With
    Else
        With cLabel
            .ForeColor = vbButtonText
            .BackColor = vbWindowBackground
            .FontBold = False
        End With
    End If

End Sub

Private Sub CargoTelefonos()
Dim rsT As rdoResultset
Dim iCant As Integer
Dim sTelef As String

    iCant = 0
    sTelef = ""
    Cons = "Select TTeCodigo, TTeNombre, TelNumero From Telefono, TipoTelefono Where TelCliente = " & oCliente.ID _
        & " And TelTipo = TTeCodigo"
    Set rsT = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsT.EOF
        If iCant = 0 Then
            'Siempre cargo el primero.
            sTelef = Trim(rsT!TTeNombre) & " " & Trim(rsT!TelNumero)
            lTelefono.ForeColor = vbRojoFuerte
        End If
        'Veo según el tipo de cliente que parámetro cargo.
        If txtCliente.Cliente.Tipo = TC_Persona Then
            If paTipoTelefonoP = rsT!TTeCodigo Then
                sTelef = Trim(rsT!TTeNombre) & " " & Trim(rsT!TelNumero)
                lTelefono.ForeColor = vbButtonText
            End If
        Else
            If paTipoTelefonoE = rsT!TTeCodigo Then
                sTelef = Trim(rsT!TTeNombre) & " " & Trim(rsT!TelNumero)
                lTelefono.ForeColor = vbButtonText
            End If
        End If
        iCant = iCant + 1
        rsT.MoveNext
    Loop
    rsT.Close
    lTelefono.Tag = iCant
    lTelefono.Caption = Trim(sTelef)
    If iCant > 1 Then
        lTelCant.Caption = "Tel.(" & iCant & "):"
    Else
        If Val(lTelefono.Tag) = 0 Then
            lTelefono.Caption = "Sin Teléfono"
            LabelMensaje lTelefono, True
        End If
        lTelCant.Caption = "Tel.:"
    End If
    
End Sub

Private Sub CamposBanco(ByVal bVisible As Boolean)
    lBanco.Visible = bVisible
    tBanco.Visible = bVisible
    lANombre.Visible = bVisible
    tANombre.Visible = bVisible
    If bVisible Then
        Me.Height = 6120
        chPagaCheque.Value = 1
    Else
        Me.Height = 5685
        chPagaCheque.Value = 0
    End If
    Refresh
End Sub

Private Sub LimpioDatosBanco()
    lBanco.Tag = ""
    tBanco.Text = ""
    tANombre.Text = ""
End Sub

Private Function BuscoBancoEmisor(Codigo As String) As Boolean
Dim Banco As String, Sucursal As String
Dim rsBco As rdoResultset

    On Error GoTo errCargar
    Banco = Mid(Codigo, 1, 2)
    Sucursal = Mid(Codigo, 3, 3)
    BuscoBancoEmisor = True
    
    Cons = "Select * from BancoSSFF, SucursalDeBanco" _
          & "  Where BanCodigoB = " & Banco _
          & "  And SBaCodigoS = " & Sucursal _
          & "  And SBaBanco = BanCodigo"
    Set rsBco = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsBco.EOF Then
        tBanco.Tag = Trim(rsBco!BanNombre)
    Else
        tBanco.Tag = ""
        BuscoBancoEmisor = False
    End If
    rsBco.Close
    Exit Function

errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar el banco emisor."
End Function

Private Sub CargoDatosCheque()
Dim sDato As String, sAux As String
    lBanco.Tag = ""
    Cons = "Select * From Comentario Where ComCliente = " & oCliente.ID _
        & " And ComTipo = " & paTipoComCheque
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        sDato = RsAux!ComComentario
        sDato = Trim(sDato)
        If InStr(1, sDato, "banco: ", vbTextCompare) > 0 And InStr(1, sDato, "A nombre de: ", vbTextCompare) > 0 Then
            If InStr(1, sDato, "(") Then
                lBanco.Tag = RsAux!ComCodigo
                sAux = Mid(sDato, InStr(1, sDato, "(") + 1, 6)
                tBanco.Text = Mid(sAux, 1, 2) & Mid(sAux, 4)
                tANombre.Text = Mid(sDato, InStr(1, sDato, "A nombre de: ") + Len("A nombre de: "))
                BuscoBancoEmisor tBanco.Text
                Exit Do
            End If
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
End Sub

Private Sub db_FindVtaPendiente(ByVal lCliente As Long)
Dim iCont As Integer
    iCont = GetQVtaTelefonicaPendiente(oCliente.ID)
    If iCont > 0 Then
        If MsgBox("El cliente tiene " & iCont & " ventas pendientes." & vbCr & "¿Desea visualizarlas?", vbQuestion + vbYesNo, "Ventas Telefónicas") = vbYes Then
            EjecutarApp App.Path & "\Contados a domicilio.exe", CStr(lCliente)
        End If
    End If
End Sub

Private Function GetQVtaTelefonicaPendiente(ByVal lCliente As Long) As Integer
Dim rsVT As rdoResultset
Dim iCant As Integer

    iCant = 0
    Cons = "Select count(*) " _
        & " From VentaTelefonica " _
        & " Where VTeTipo IN(" & TipoDocumento.ContadoDomicilio & ", 32, 33)" _
        & " And VTeCliente = " & lCliente _
        & " And VTeAnulado = Null And VTeDocumento = Null"
        
    Set rsVT = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsVT.EOF Then iCant = rsVT(0)
    rsVT.Close
    GetQVtaTelefonicaPendiente = iCant
    
End Function

Private Function IsArtInVtaPendiente(ByVal sArt As String) As Long
Dim rsVT As rdoResultset
    
    IsArtInVtaPendiente = 0
    Cons = "Select * " _
        & " From VentaTelefonica, RenglonVtaTelefonica " _
        & " Where VTeTipo IN(32,33, " & TipoDocumento.ContadoDomicilio & ")" _
        & " And VTeCliente = " & oCliente.ID _
        & " And RVTArticulo In (" & sArt & ")" _
        & " And VTeCodigo = RVtVentaTelefonica" _
        & " And VTeAnulado = Null And VTeDocumento = Null"
        
    Set rsVT = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    'Puede haber + de una
    If Not rsVT.EOF Then IsArtInVtaPendiente = rsVT!VTeCodigo
    rsVT.Close
    

End Function

Private Sub ControloVentaTelefonica()
Dim lVT As Long
Dim sArt As String
    
    With vsGrilla
        For lVT = 1 To .Rows - 1
            If sArt <> "" Then sArt = sArt & ", "
            sArt = sArt & .Cell(flexcpData, lVT, 0)
        Next
    End With
    lVT = IsArtInVtaPendiente(sArt)
    If lVT > 0 Then
        MsgBox "El cliente tiene una venta telefónica pendiente y se detectó que tiene por lo menos un artículo de los que quiere facturar." & vbCr & "Es conveniente si la venta no se realizará que la misma se elimine.", vbExclamation, "Validación"
        EjecutarApp App.Path & "\Contados a domicilio.exe", "i" & lVT, True
    End If
    
End Sub

Private Sub f_GetQEnvioInstala(iItems As Integer, iEnvia As Integer, iInstala As Integer)
On Error Resume Next
'Recorro la lista y retorno el total de artículos y el total que se envía
Dim iCont As Integer

    iItems = 0
    iEnvia = 0
    iInstala = 0
    
    With vsGrilla
        For iCont = 1 To .Rows - 1
             ' .Cell(flexcpData, .Rows - 1, 3)                tipo del artículo.
                If CLng(.Cell(flexcpData, iCont, 1)) <> TipoArticulo.Servicio And _
                    Not EsTipoDeServicio(CLng(.Cell(flexcpData, iCont, 3))) Then
                    iItems = iItems + .Cell(flexcpValue, iCont, 0)
                    iEnvia = iEnvia + Val(.Cell(flexcpData, iCont, 6))
                    If .Cell(flexcpValue, iCont, 8) > 0 Then iInstala = iInstala + (.Cell(flexcpValue, iCont, 0) - Val(.Cell(flexcpData, iCont, 6)))
                End If
        Next iCont
    End With
    
End Sub

Private Sub f_ValidateRetira(ByVal bSetFoco As Boolean)
Dim iQitem As Integer
Dim iQEnv As Integer
Dim iQInst As Integer
On Error Resume Next

    'Busco cantidades.
    f_GetQEnvioInstala iQitem, iQEnv, iQInst
    
    'Casos para habilitar el check.
'    El Check se habilita    Si se instala todo _
                                    Si se instala algo y hay envios parciales

    chRetira.Enabled = IIf(iQInst = iQitem Or (iQEnv <> iQitem And iQEnv > 0 And iQInst > 0), True, False)
    
    '.........................................................
    If Not chRetira.Enabled Then
        If iQitem = iQEnv And iQInst = 0 Then chRetira.Value = 0 Else chRetira.Value = 1
    Else
        chRetira.Value = IIf(iQitem <> iQEnv And iQitem <> iQInst And (iQitem - iQEnv - iQInst) > 0, 1, 0)
    End If
    
    If iQitem <> iQEnv Then
        'Consulto si hay artículos de instalación.
        If iQInst = 0 And Not chRetira.Enabled Then
            chRetira.Value = 1
            tFRetiro.Text = ""
        End If
    End If
    If bSetFoco Then cPendiente.SetFocus

End Sub

Private Function GetInstaladorArt(ByVal lIDArt As Long) As Long
Dim rsIns As rdoResultset
    
    GetInstaladorArt = 0
    'Ya le pongo que sea positivo .
    Cons = "Select * From Articulo Where ArtID = " & lIDArt _
            & " And ArtInstalador > 0 "
    Set rsIns = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsIns.EOF Then GetInstaladorArt = rsIns!ArtInstalador
    rsIns.Close
    
End Function

Private Sub GetStringRetira(sFechaP As String) '(ByRef sTextoImp As String, sFechaP As String)
Dim iQitem As Integer
Dim iQEnv As Integer
Dim iQInst As Integer
On Error Resume Next
    
    'Busco cantidades.
 '   f_GetQEnvioInstala iQitem, iQEnv, iQInst
    
    If iQEnv > 0 Then
        'Tengo envíos
        If iQEnv = iQitem Then
        '    sTextoImp = IIf(iQEnv = 1, "Se Envía", "Se Envía Todo")
            sFechaP = DateAdd("d", 1, Date) & " 09:00:00"
        Else
            If iQInst = 0 Then
                'No tengo instalación.
             '   sTextoImp = " Se Envían " & iQEnv & " de " & iQitem
                sFechaP = tFRetiro.Text
                If tFRetiro.Tag <> "" Then sFechaP = sFechaP & " 23:00:00"
            Else
                'Tengo instalaciones
              '  sTextoImp = " Se Envían " & iQEnv & " de " & iQitem & ". V.Inst."
                If chRetira.Value = 1 Then
                    sFechaP = tFRetiro.Text
                    If tFRetiro.Tag <> "" Then sFechaP = sFechaP & " 23:00:00"
                Else
                    sFechaP = DateAdd("d", 1, Date) & " 10:00:00"
                End If
            End If
        End If
    Else
        If iQInst > 0 Then
            If iQInst <> iQitem Then
       '         sTextoImp = "Retira " & tFRetiro.Text & ". V.Inst."
                sFechaP = tFRetiro.Text
                If tFRetiro.Tag <> "" Then sFechaP = sFechaP & " 23:00:00"
            Else
                If chRetira.Value = 1 Then
                    sFechaP = tFRetiro.Text
                    If tFRetiro.Tag <> "" Then sFechaP = sFechaP & " 23:00:00"
              '      sTextoImp = "Retira " & tFRetiro.Text & ". V.Inst."
                Else
                    sFechaP = DateAdd("d", 1, Date) & " 10:00:00"
             '       sTextoImp = "Se Instala c/Sum."
                End If
            End If
        Else
            If chRetira.Value = 0 Then
                'Ocurre cuando iqitem = 0
        '        sTextoImp = ""
                sFechaP = DateAdd("d", 1, Date) & " 08:00:00"
            Else
                'Caso todo a retirar.
          '      sTextoImp = "Retira " & tFRetiro.Text
                sFechaP = tFRetiro.Text
                If tFRetiro.Tag <> "" Then sFechaP = sFechaP & " 23:00:00"
            End If
        End If
    End If

End Sub

Private Sub s_LoadMenuOpcionPrint()
Dim vOpt() As String
Dim iQ As Integer
    
    MnuPrintLine1.Visible = (paOptPrintList <> "")
    MnuPrintOpt(0).Visible = (paOptPrintList <> "")
    
    If paOptPrintList = "" Then
        Exit Sub
    ElseIf InStr(1, paOptPrintList, "|", vbTextCompare) = 0 Then
        MnuPrintOpt(0).Caption = paOptPrintList
    Else
        vOpt = Split(paOptPrintList, "|")
        For iQ = 0 To UBound(vOpt)
            If iQ > 0 Then Load MnuPrintOpt(iQ)
            With MnuPrintOpt(iQ)
                .Caption = Trim(vOpt(iQ))
                .Checked = (LCase(.Caption) = LCase(paOptPrintSel))
                .Visible = True
            End With
        Next
    End If
    
End Sub

Private Sub s_SetPrinter()
On Error Resume Next
    
    Dim iQ As Integer
    For iQ = MnuPrintOpt.LBound To MnuPrintOpt.UBound
        MnuPrintOpt(iQ).Checked = False
        MnuPrintOpt(iQ).Checked = (MnuPrintOpt(iQ).Caption = paOptPrintSel)
    Next
    lPrint.Caption = "Impresora: " & paIContadoN
    If Trim(Printer.DeviceName) <> Trim(paIContadoN) Then SeteoImpresoraPorDefecto paIContadoN
    
End Sub

Private Function fnc_DireccionInsertada(ByVal lID As Long) As Boolean
Dim iQ As Integer
    fnc_DireccionInsertada = False
    For iQ = 0 To cDireccion.ListCount - 1
        If cDireccion.ItemData(iQ) = lID Then fnc_DireccionInsertada = True: Exit Function
    Next
End Function

Private Sub loc_FindDireccionAuxiliarTexto()
On Error GoTo errFDA
Dim rsD As rdoResultset
    Cons = "Select DAuDireccion , DAuNombre " & _
                "From DireccionAuxiliar Where DAuCliente = " & oCliente.ID & _
                " And DAuNombre Like '" & Replace(cDireccion.Text, " ", "%") & "%'"
    Set rsD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsD.EOF Then
        rsD.MoveNext
        If rsD.EOF Then
            rsD.MoveFirst
            If fnc_DireccionInsertada(rsD("DAuDireccion")) Then
                BuscoCodigoEnCombo cDireccion, rsD("DAuDireccion")
            Else
                'Inserto
                With cDireccion
                    .AddItem Trim(rsD("DAuNombre"))
                    .ItemData(.NewIndex) = rsD("DAuDireccion")
                    .ListIndex = .NewIndex
                End With
            End If
            chNomDireccion.Value = 1
        Else
            Dim objLista As New clsListadeAyuda
            With objLista
                If .ActivarAyuda(cBase, Cons, 3000, 1, "Dirección Auxiliar") > 0 Then
                    If fnc_DireccionInsertada(.RetornoDatoSeleccionado(0)) Then
                        BuscoCodigoEnCombo cDireccion, .RetornoDatoSeleccionado(0)
                    Else
                        cDireccion.AddItem Trim(.RetornoDatoSeleccionado(1))
                        cDireccion.ItemData(cDireccion.NewIndex) = .RetornoDatoSeleccionado(0)
                        cDireccion.ListIndex = cDireccion.NewIndex
                    End If
                    chNomDireccion.Value = 1
                End If
            End With
            Set objLista = Nothing
        End If
        rsD.Close
    Else
        rsD.Close
        MsgBox "No hay coincidencias.", vbInformation, "Buscar dirección auxiliar"
    End If
Exit Sub
errFDA:
    clsGeneral.OcurrioError "Error al buscar la dirección auxiliar.", Err.Description
End Sub

Private Function ValidarVersionEFactura() As Boolean
On Error GoTo errEC
    With New clsCGSAEFactura
        ValidarVersionEFactura = .ValidarVersion()
    End With
    Exit Function
errEC:
End Function

Private Function EmitirCFE(ByVal Documento As clsDocumentoCGSA, ByVal CAE As clsCAEDocumento) As String
On Error GoTo errEC
    With New clsCGSAEFactura
        .URLAFirmar = prmURLFirmaEFactura
        .TasaBasica = TasaBasica
        .TasaMinima = TasaMinima
        .ImporteConInfoDeCliente = prmImporteConInfoCliente
        Set .Connect = cBase
        If Not .GenerarEComprobante(CAE, Documento, EmpresaEmisora, paCodigoDGI) Then
            EmitirCFE = .XMLRespuesta
        End If
    End With
    Exit Function
errEC:
    EmitirCFE = "Error en firma: " & Err.Description
End Function

Private Sub EnvioALog(ByVal Texto As String)
On Error GoTo errEAL
    Open "\\ibm3200\oyr\efactura\logCtdoEFactura.txt" For Append As #1
    Print #1, Now & Space(5) & "Terminal: " & miConexion.NombreTerminal & Space(5) & "CONTADO" & Space(5) & Texto
    Close #1
    Exit Sub
errEAL:
End Sub

Private Sub SeteoInfoDocumentoCliente()
    lblRucPersona.Visible = (txtCliente.DocumentoCliente <> DC_RUT)
    lblInfoCliente.ForeColor = vbBlack
    lblInfoCliente.FontBold = False
    Select Case txtCliente.DocumentoCliente
        Case DC_CI
            lblInfoCliente.Caption = "C.I.:"
        Case DC_RUT
            lblInfoCliente.Caption = "R.U.T.:"
        Case Else
            If txtCliente.Cliente.TipoDocumento.Nombre = "" Then
                lblInfoCliente.Caption = "Otro:"
            Else
                lblInfoCliente.Caption = txtCliente.Cliente.TipoDocumento.Abreviacion
            End If
            lblInfoCliente.ForeColor = &HFF&
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

Private Function ValidoRUT() As Boolean
On Error GoTo errVR
    ValidoRUT = True
    Dim oValida As New clsValidaRUT
    If txtCliente.Cliente.Tipo = TC_Empresa And txtCliente.Cliente.Documento <> "" Then
        If Not oValida.ValidarRUT(txtCliente.Cliente.Documento) Then
            MsgBox "RUT INCORRECTO!!!, por favor valide con el cliente el número de RUT ya que no cumple con la validación.", vbExclamation, "RUT INCORRECTO"
            ValidoRUT = False
        End If
    ElseIf txtCliente.Cliente.Tipo = TC_Persona And txtCliente.Cliente.RutPersona <> "" Then
        If Not oValida.ValidarRUT(txtCliente.Cliente.RutPersona) Then
            MsgBox "RUT INCORRECTO!!!, por favor valide con el cliente el número de RUT ya que no cumple con la validación.", vbExclamation, "RUT INCORRECTO"
            ValidoRUT = False
        End If
    End If
    Set oValida = Nothing
    Exit Function
errVR:
    clsGeneral.OcurrioError "Error al validar el RUT, no podrá facturar.", Err.Description
    ValidoRUT = False
End Function
