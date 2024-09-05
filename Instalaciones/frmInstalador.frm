VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{923DD7D8-A030-4239-BCD4-51FDB459E0FE}#4.0#0"; "orgComboCalculator.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmInstall 
   BackColor       =   &H00EEFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instalaciones"
   ClientHeight    =   4830
   ClientLeft      =   2895
   ClientTop       =   2235
   ClientWidth     =   8250
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInstalador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8250
   Begin VB.CheckBox chHaceRemito 
      BackColor       =   &H00EEFFFF&
      Caption         =   "&Hacer remito"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   6000
      TabIndex        =   31
      Top             =   3600
      Width           =   2175
   End
   Begin AACombo99.AACombo cInstalador 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
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
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsArticulo 
      Height          =   855
      Left            =   1080
      TabIndex        =   22
      Top             =   3600
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1508
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
      BackColor       =   15663103
      ForeColor       =   -2147483641
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   15663103
      BackColorAlternate=   15663103
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
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   0
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
   Begin VB.TextBox tMemo 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   1080
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   20
      Text            =   "frmInstalador.frx":0442
      Top             =   3000
      Width           =   7095
   End
   Begin orgCalculatorFlat.orgCalculator caCobrar 
      Height          =   285
      Left            =   4680
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorCalculator=   -2147483633
      BackColorOperator=   -2147483636
      ForeColorOperator=   -2147483634
      Text            =   "0.00"
   End
   Begin orgCalculatorFlat.orgCalculator caViatico 
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorCalculator=   -2147483633
      BackColorOperator=   -2147483636
      ForeColorOperator=   -2147483634
      Text            =   "0.00"
   End
   Begin VB.TextBox tTelefono 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6000
      MaxLength       =   15
      TabIndex        =   15
      Top             =   1920
      Width           =   1320
   End
   Begin VB.TextBox tInterno 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   16
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox tContacto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      MaxLength       =   30
      TabIndex        =   18
      Text            =   "QQQQQQQQQQQQQQQQQQQQQQQQQQQQQQ"
      Top             =   2280
      Width           =   3495
   End
   Begin VB.ComboBox cDireccion 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton bDireccion 
      Caption         =   "Dirección&..."
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox tDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      ForeColor       =   &H8000000E&
      Height          =   675
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   25
      Text            =   "frmInstalador.frx":04A7
      Top             =   2280
      Width           =   3585
   End
   Begin VB.ComboBox cRangoHora 
      Height          =   315
      Left            =   6120
      TabIndex        =   6
      Text            =   "8888-8888"
      Top             =   1200
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dpFecha 
      Height          =   315
      Left            =   4680
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   41418755
      CurrentDate     =   37909
   End
   Begin VB.TextBox tID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   500
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "next"
            Object.ToolTipText     =   "Siguiente Instalación del documento"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "viaticos"
            Object.ToolTipText     =   "Abrir plantilla de viáticos"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   4575
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
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
      Left            =   5760
      Top             =   -120
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
            Picture         =   "frmInstalador.frx":04CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInstalador.frx":05DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInstalador.frx":06EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInstalador.frx":0800
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInstalador.frx":0912
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInstalador.frx":0A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInstalador.frx":0D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInstalador.frx":0E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInstalador.frx":116A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInstalador.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInstalador.frx":179E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInstalador.frx":18F8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin AACombo99.AACombo cTipoTelefono 
      Height          =   315
      Left            =   4680
      TabIndex        =   14
      Top             =   1920
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
   Begin VB.Label lbLiquidacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      Caption         =   " Liquidación: 123456888"
      ForeColor       =   &H00EEFFFF&
      Height          =   285
      Left            =   6000
      TabIndex        =   35
      Top             =   4170
      Width           =   2175
   End
   Begin VB.Label lPrecio 
      Appearance      =   0  'Flat
      BackColor       =   &H00608000&
      Caption         =   " de "
      ForeColor       =   &H00EEFFFF&
      Height          =   285
      Left            =   6120
      TabIndex        =   34
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lInstalada 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Instalada: 88-88-8888 por Juan Fernandez"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EEFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   33
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Label lAnulada 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "Anulada: 88-88-8888"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EEFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   32
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artículos:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lRemito 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Remito asociado: 8888888"
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
      Left            =   6000
      TabIndex        =   30
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label lModificada 
      BackStyle       =   0  'Transparent
      Caption         =   "Modificada: 10-06-03 15:50 por Elizabeth"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3840
      TabIndex        =   29
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label lAlta 
      BackStyle       =   0  'Transparent
      Caption         =   "Alta: QQQQQQQQQQ"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2040
      TabIndex        =   28
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00EEFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&Código:"
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lDoc 
      BackColor       =   &H8000000C&
      Caption         =   "Déposito Colonia Crédito B-588888"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2040
      TabIndex        =   27
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "Documento asociado:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   480
      Width           =   8055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mentario:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Viático:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&brar:"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&Teléfono:"
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&ntacto:"
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha:"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Instalador:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   855
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
      Begin VB.Menu MnuLineIndep 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOptIndependizarRemito 
         Caption         =   "Independizar remito"
      End
      Begin VB.Menu MnuSalirLine 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVolver 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuVer 
      Caption         =   "Ver"
      Begin VB.Menu MnuVerFactura 
         Caption         =   "Detalle de Factura"
      End
   End
End
Attribute VB_Name = "frmInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'MODIFICACIONES-----------------------------------------------------------------------------------------------------------------
'13-2-2004
'  ***  Cobro de la instalación  ***
'  Cree campo en tabla rengloninstalacion donde guardo el cobro que no sea del documento que facturo el aire.
'  Este campo tiene el sgte formato
'   QCobrEnInstalacion|QEnDoc:IDDoc;QEnDoc2:IDDoc2
' Donde QCobrEnInstalacion significa la cantidad que se incluyo en el cobro de la instalación
' luego QEnDoc = Q de art. de cobro que se tomaron del documento seguido de los 2 ptos.
'--------------------------------------------------------------------------------------------------------------------------------------

'13-12-04 Agregue consulta con Q de instalaciones
'21-12-04 Controlo que no me haga remito cuando no es un ctdo o crédito.

'06/10/2007 dejo hacer más de un remito y poner Q cero en la grilla para independizar artículos.

'17/9/2008 independización de remito.


Public prmDocumento As Long
Public prmTipoDoc As Byte
Public prmIDInst As Long

Private bLoad As Boolean
Private lIDCli As Long
Private arrInstDoc() As Long

Private Sub bDireccion_Click()
On Error Resume Next
Dim lNewDir As Long
    
    If cDireccion.ListIndex = -1 Or cDireccion.ListIndex = -2 Then Exit Sub
    
    If cDireccion.ListIndex = 0 Then
        If cDireccion.ItemData(0) = 0 Then
            lNewDir = 0
            If Val(bDireccion.Tag) > 0 Then
                lNewDir = CopyDireccion(bDireccion.Tag)
                If lNewDir > 0 Then
                    cDireccion.ItemData(0) = lNewDir
                    cDireccion.Tag = lNewDir
                Else
                    MsgBox "No se logró copiar la dirección principal del cliente, reintente dar click en el botón.", vbCritical, "Error"
                    cDireccion.Tag = ""
                    Exit Sub
                End If
            End If
        Else
            lNewDir = cDireccion.ItemData(0)
        End If
                    
        Dim objDir As New clsDireccion
        If Val(tID.Text) = 0 Then
            objDir.ActivoFormularioDireccion cBase, lNewDir, lIDCli
        Else
            objDir.ActivoFormularioDireccion cBase, cDireccion.ItemData(0), lIDCli, "Instalacion", "InsDireccion", "InsID", Val(tID.Text)
        End If
        
        'Tomo el id de la clase
        lNewDir = objDir.CodigoDeDireccion
                       
        If cDireccion.ItemData(0) = 0 Then
            If lNewDir > 0 Then
                'Me guardo el id diciendo que es nueva.
                cDireccion.ItemData(0) = lNewDir
                cDireccion.Tag = lNewDir
            Else
                'X si estaba cargada para no borrar algo que ya no existe.
                cDireccion.Tag = ""
            End If
        Else
            If Val(cDireccion.Tag) = cDireccion.ItemData(0) And lNewDir = 0 Then cDireccion.Tag = ""
            cDireccion.ItemData(0) = lNewDir
        End If
        If lNewDir = 0 Then tDireccion.Text = "" Else Call cDireccion_Click
        If tTelefono.Text = "" Then
            If cTipoTelefono.Enabled Then
                cTipoTelefono.SetFocus
            ElseIf tContacto.Enabled Then
                tContacto.SetFocus
            End If
        End If
    End If
    
End Sub

Private Sub caCobrar_GotFocus()
    Status.SimpleText = "Ingrese el valor que se cobrará por la instalación."
End Sub

Private Sub caCobrar_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then Foco cDireccion
End Sub

Private Sub caCobrar_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub caViatico_GotFocus()
    Status.SimpleText = "Ingrese el valor del viático del instalador."
End Sub

Private Sub caViatico_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then caCobrar.SetFocus
End Sub

Private Sub caViatico_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub cDireccion_Click()
Dim lDirFind As Long

    tDireccion.Text = ""
    lDirFind = 0
    If cDireccion.ListIndex > -1 Then
        If cDireccion.ListIndex = 0 Then
            'Es la dirección a instalar.
            If cDireccion.ItemData(0) = 0 Then
                'No hay dirección ingresada caso nuevo.
                lDirFind = Val(bDireccion.Tag)
            Else
                'Ya hay dirección grabada.
                lDirFind = cDireccion.ItemData(0)
            End If
        Else
            lDirFind = Val(cDireccion.ItemData(cDireccion.ListIndex))
        End If
        
        If lDirFind > 0 Then
            tDireccion.Text = clsGeneral.ArmoDireccionEnTexto(cBase, lDirFind, False, True, True, True, False, False, True)
        End If
    Else
        tDireccion.Text = ""
    End If
    bDireccion.Enabled = (cDireccion.ListIndex = 0) And cDireccion.Enabled
        
End Sub

Private Sub cDireccion_GotFocus()
    Status.SimpleText = "Seleccione la dirección donde se hará la instalación."
End Sub

Private Sub cDireccion_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If cDireccion.ItemData(cDireccion.ListIndex) = -2 Then
            db_FindDireccionUltimosEnvios
        End If
        If tTelefono.Text = "" Then cTipoTelefono.SetFocus Else tContacto.SetFocus
    End If
End Sub

Private Sub chHaceRemito_Click()
Dim iCont As Integer

    Exit Sub
    With vsArticulo
        If Not .Enabled Then Exit Sub
        
        If chHaceRemito.Value = 0 Then
            
            For iCont = 0 To .Rows - 1
                'Pongo toda la cantidad que tiene el documento disponible.
                .Cell(flexcpText, iCont, 0) = .Cell(flexcpData, iCont, 1) - (.Cell(flexcpData, iCont, 2) - Val(.Cell(flexcpValue, iCont, 3)))
                f_SetTotalCobro iCont
            Next
            
        Else
            'Pongo la cantidad que tiene para retirar
            For iCont = 0 To .Rows - 1
                If .Cell(flexcpValue, iCont, 2) > 0 Then
                    'Tiene para retirar.
                    'Primero pongo todo lo disponible para que se pueda instalar.
                    .Cell(flexcpText, iCont, 0) = .Cell(flexcpData, iCont, 1) - (.Cell(flexcpData, iCont, 2) - Val(.Cell(flexcpValue, iCont, 3)))
                    If .Cell(flexcpValue, iCont, 0) > .Cell(flexcpValue, iCont, 2) Then
                        .Cell(flexcpText, iCont, 0) = .Cell(flexcpValue, iCont, 2)
                    End If
                Else
                    .Cell(flexcpText, iCont, 0) = 0
                End If
                f_SetTotalCobro iCont
            Next
        End If
        f_SetTotalInstalacion
    End With
End Sub

Private Sub cInstalador_Change()
    cInstalador.Tag = ""
End Sub

Private Sub cInstalador_Click()
    cInstalador.Tag = ""
End Sub

Private Sub cInstalador_GotFocus()
    With cInstalador
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Seleccione quien hará la instalación."
End Sub

Private Sub cInstalador_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn And cInstalador.ListIndex > -1 Then
        If cInstalador.Tag = "" Then
            'Cargo automáticamente el primer día de la agenda.
            If loc_SetFechaTipoFlete Then
                cInstalador.Tag = "1"
                caCobrar.SetFocus
            Else
                dpFecha.SetFocus
            End If
        Else
            caCobrar.SetFocus
        End If
    End If
End Sub

Private Sub cInstalador_LostFocus()
    Status.SimpleText = ""
    cInstalador.SelStart = 0
End Sub

Private Sub cRangoHora_GotFocus()
    With cRangoHora
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Seleccione o ingrese un horario."
End Sub

Private Sub cRangoHora_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If cRangoHora.ListIndex = -1 And cRangoHora.Text <> "" Then If Not ValidoRangoHorario Then Exit Sub
        caViatico.SetFocus
    End If
End Sub

Private Sub cRangoHora_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub cTipoTelefono_Change()
    tTelefono.Text = "": tInterno.Text = ""
End Sub

Private Sub cTipoTelefono_Click()
    tTelefono.Text = "": tInterno.Text = ""
End Sub

Private Sub cTipoTelefono_GotFocus()
    With cTipoTelefono
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Seleccione el tipo de teléfono a registrar."
End Sub

Private Sub cTipoTelefono_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If tTelefono.Text = "" Then
            If cTipoTelefono.ListIndex = -1 Then tContacto.SetFocus Else loc_LoadTelefono
            Exit Sub
        Else
            If cTipoTelefono.ListIndex > -1 Then tContacto.SetFocus
        End If
    End If
End Sub

Private Sub cTipoTelefono_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub dpFecha_GotFocus()
    Status.SimpleText = "Seleccione la fecha que se prometió la instalación."
End Sub

Private Sub dpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn Then cRangoHora.SetFocus
End Sub

Private Sub dpFecha_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then cRangoHora.SetFocus
End Sub

Private Sub dpFecha_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then loc_ShowTotalInstalaciones
End Sub

Private Sub dpFecha_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub Form_Activate()
    
    If Not bLoad Then
        
        If prmDocumento > 0 Then
            'Lo más seguro es que venga para ingresar uno nuevo.
            loc_AccesoPorDocumento
            '---------------------------------------------------------------
        Else
            If prmIDInst > 0 Then LoadInstalacion
        End If
        
    End If
    bLoad = True
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    ObtengoSeteoForm Me, 500, 500
    Me.Height = 5430
    Me.Width = 8340
    bLoad = False
    loc_CleanCtrl
    loc_SetCtrl False
    With vsArticulo
        .Cols = 8
        .ColHidden(2) = True
        .ColHidden(3) = True
        .ColHidden(4) = True
        .ColHidden(5) = True
        .ColHidden(7) = True
        .ColWidth(0) = 300
        .ColWidth(1) = 3200
    End With
    
    ReDim arrInstDoc(0)
    CargoCombo "Select InsCodigo,InsNombre From Instaladores Order by InsNombre", cInstalador
    CargoCombo "Select TTeCodigo, TTeNombre From TipoTelefono Order by TTeNombre", cTipoTelefono
    Screen.MousePointer = 0
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al ingresar al formulario."
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    
    If Val(cDireccion.Tag) > 0 And Val(cDireccion.Tag) = cDireccion.ItemData(0) Then
        Cons = "Delete Direccion Where DirCodigo = " & Val(cDireccion.Tag)
        cBase.Execute (Cons)
    End If
    Erase arrInstDoc
    
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    End
    Exit Sub
    
End Sub

Private Sub Label10_Click()
    Foco tID
End Sub

Private Sub Label2_Click()
    Foco cInstalador
End Sub

Private Sub Label3_Click()
    If dpFecha.Enabled Then dpFecha.SetFocus
End Sub

Private Sub Label4_Click()
    Foco tContacto
End Sub

Private Sub Label5_Click()
    Foco cTipoTelefono
End Sub

Private Sub Label6_Click()
On Error Resume Next
    If caCobrar.Enabled Then caCobrar.SetFocus
End Sub

Private Sub Label7_Click()
On Error Resume Next
    If caViatico.Enabled Then caViatico.SetFocus
End Sub

Private Sub Label8_Click()
    Foco tMemo
End Sub

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

Private Sub MnuOptIndependizarRemito_Click()
    
    If MsgBox("Confirma desvincular el remito de la instalación?", vbQuestion + vbYesNo + vbDefaultButton2, "Independizar remito") = vbYes Then
        On Error GoTo errSave
        Dim lUID As Long, sDefSuc As String
        Dim objSuceso As New clsSuceso
        With objSuceso
            .TipoSuceso = TipoSuceso.VariosStock
            .ActivoFormulario paCodigoDeUsuario, "Independizar remito", cBase
            lUID = .Usuario
            sDefSuc = .Defensa
        End With
        Set objSuceso = Nothing
        Me.Refresh
        If lUID = 0 Then Screen.MousePointer = 0: Exit Sub
        
        Set RsAux = cBase.OpenResultset("EXEC prg_Instalaciones_IndependizarRemito " & prmIDInst, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If RsAux(0) = 1 Then
                'Grabo el suceso.
                clsGeneral.RegistroSuceso cBase, Now, TipoSuceso.VariosStock, paCodigoDeTerminal, lUID, IIf(prmTipoDoc = 1, prmDocumento, 0), 0, "Independizó el remito de id: " & Val(lRemito.Tag) & " de la instalación: " & prmIDInst, sDefSuc, , lIDCli
                RsAux.Close
                LoadInstalacion
                Exit Sub
            Else
                MsgBox "ERROR al almacenar la información: " & Trim(RsAux(1)), vbCritical, "Independizar remito"
            End If
        End If
        RsAux.Close
    End If
    Exit Sub
errSave:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al intentar independizar.", Err.Description
End Sub

Private Sub MnuVerFactura_Click()
On Error Resume Next
    If prmDocumento > 0 Then
        If prmTipoDoc = 1 Then
            EjecutarApp App.Path & "\Detalle de factura.exe", CStr(prmDocumento)
        Else
            EjecutarApp App.Path & "\Contados a Domicilio.exe", "i" & CStr(prmDocumento)
        End If
    End If
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Sub AccionNuevo()
On Error GoTo ErrAN

    If prmDocumento = 0 Then Exit Sub

    Screen.MousePointer = 11
    
    'Cargo los artículos que tengo disponible para instalar.
    If prmTipoDoc = 1 Then
        BuscoCodigoEnCombo cInstalador, LoadArticuloDocumento
    Else
        BuscoCodigoEnCombo cInstalador, LoadArticuloVenta
    End If
    
    If vsArticulo.Rows = 0 Then
        Screen.MousePointer = 0
        MsgBox "No quedan artículos disponibles a instalar para el documento seleccionado.", vbExclamation, "ATENCIÓN"
                
        If tID.Text <> "" Then
            prmIDInst = Val(tID.Text)
            LoadInstalacion
        End If
        Exit Sub
    End If
    
    'Prendo Señal que es uno nuevo.
    Toolbar1.Tag = 1
    'Habilito y Desabilito Botones.
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    MnuOptIndependizarRemito.Enabled = False
    Toolbar1.Buttons("next").Enabled = False
    
    'Habilito controles
    With lInstalada
        If .Visible Then .Visible = False: .Caption = ""
    End With
    
    'No borro para dejar los otros campos =.
    lbLiquidacion.Visible = False
    lbLiquidacion.Caption = ""
    
    loc_SetCtrl True
    tID.Text = ""
    
    'Busco si para el documento tengo artículos disponibles a instalar.
    loc_LoadComboDireccion
    
    If cDireccion.ListCount > 0 Then cDireccion.ListIndex = 0
    LoadTelefonoDefault
    'Posiciono
    loc_SetDefaultNew
    f_SetTotalInstalacion
    If prmTipoDoc = 1 Then HaceRemito
    If chHaceRemito.Visible Then lRemito.Caption = "": lRemito.Tag = ""
    
    dpFecha.Tag = ""
    Screen.MousePointer = 0
    Exit Sub
    
ErrAN:
    clsGeneral.OcurrioError "Ocurrio un error inesperado.", Err.Description
    Screen.MousePointer = 0
End Sub

Sub AccionModificar()
    
    Toolbar1.Tag = 2
    'Habilito y Desabilito Botones.
    
    'Habilito todo tal cual pudiese hacer todo.
    loc_SetCtrl True
            
    If lRemito.Caption <> "" Then
        With vsArticulo
            .Enabled = True
            .BackColor = vbButtonFace
            .BackColorBkg = .BackColor
        End With
    Else
        If prmTipoDoc = 1 Then HaceRemito
    End If
        
    Call Botones(False, False, False, True, True, Toolbar1, Me)
    Toolbar1.Buttons("next").Enabled = False
    MnuOptIndependizarRemito.Enabled = False
    If lPrecio.Tag = "" Then f_SetTotalInstalacion
    caCobrar.SetFocus
    Screen.MousePointer = 0

End Sub

Sub AccionGrabar()
Dim dNow As String
Dim iCont As Integer
Dim lRemito As Long

    If Not loc_ValidateSave Then Exit Sub
    
    If MsgBox("Confirma almacenar la información ingresada?", vbQuestion + vbYesNo, "Grabar") = vbYes Then
        
        'Veo Máximo.
        'Además de consultar Cierro la agenda si está abierta para ese día.
        Dim bSuceso As Boolean, lUID As Long, sDefSuc As String
        If Val(Toolbar1.Tag) = 1 Or dpFecha.Tag <> dpFecha.Value Then
            If Not fnc_ControlQMaxima(bSuceso) Then Exit Sub
        End If
        
        If bSuceso Then
            Dim objSuceso As New clsSuceso
            With objSuceso
                .TipoSuceso = 21
                .ActivoFormulario paCodigoDeUsuario, "Instalación", cBase
                lUID = .Usuario
                sDefSuc = .Defensa
            End With
            Set objSuceso = Nothing
            Me.Refresh
            If lUID = 0 Then Screen.MousePointer = 0: Exit Sub
        End If
        
        Screen.MousePointer = 11
        dNow = GetFechaServidor
        
        If prmTipoDoc = 1 Then
            Cons = "Select DocFModificacion From Documento Where DocCodigo = " & prmDocumento
        Else
            Cons = "Select VTeFModificacion From VentaTelefonica Where VTeCodigo = " & prmDocumento
        End If
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If CDate(lDoc.Tag) <> CDate(RsAux(0)) Then
            MsgBox "Documento modificado, cancele y vuelva a cargar la información para el mismo.", vbCritical, "Atención"
            RsAux.Close
            Screen.MousePointer = 0
            Exit Sub
        End If
        RsAux.Close
        
        On Error GoTo errBT
        cBase.BeginTrans
        On Error GoTo ErrRoll
        
        '..........................................................................................................Remito
        If chHaceRemito.Value = 1 And chHaceRemito.Visible = True Then
            'Creo el remito para la mercadería a retirar.
            lRemito = SaveRemito(dNow)
        Else
            lRemito = 0
        End If
        '..........................................................................................................
        
        If Val(Toolbar1.Tag) = 1 Then
            Cons = "Select * From Instalacion Where InsID = 0"
        Else
            Cons = "Select * From Instalacion Where InsID = " & prmIDInst
        End If
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsAux.EOF Then
            RsAux.AddNew
            RsAux!InsUsuarioAlta = paCodigoDeUsuario
        Else
            RsAux.Edit
        End If
        If lRemito <> 0 Then RsAux!InsRemito = lRemito
        RsAux!InsInstalador = cInstalador.ItemData(cInstalador.ListIndex)
        RsAux!InsDocumento = prmDocumento
        RsAux!InsTipoDocumento = prmTipoDoc
        RsAux!InsFechaProm = Format(dpFecha.Value, "yyyy/mm/dd")
        If cRangoHora.Text <> "" Then
            If cRangoHora.ListIndex = -1 Then
                RsAux!InsRangoHora = cRangoHora.Text
            Else
                RsAux!InsRangoHora = GetRangoHora
            End If
        Else
            RsAux!InsRangoHora = Null
        End If
        
        If cDireccion.ItemData(0) = 0 Then
            'Tengo que hacer la copia de la dirección principal o de la que tenga seleccionada en el combo.
            If cDireccion.ListIndex = 0 Then
                RsAux!InsDireccion = CopyDireccion(bDireccion.Tag)
            Else
                If cDireccion.ItemData(cDireccion.ListIndex) > 0 Then
                    RsAux!InsDireccion = CopyDireccion(cDireccion.ItemData(cDireccion.ListIndex))
                End If
            End If
        Else
            If cDireccion.ListIndex = 0 Then
                RsAux!InsDireccion = cDireccion.ItemData(0)
            Else
                Cons = "Delete Direccion Where DirCodigo = " & cDireccion.ItemData(0)
                cBase.Execute (Cons)
                If cDireccion.ItemData(cDireccion.ListIndex) > 0 Then RsAux!InsDireccion = CopyDireccion(cDireccion.ItemData(cDireccion.ListIndex))
            End If
        End If
        
        If Trim(tContacto.Text) = "" Then
            RsAux!InsContacto = Null
        Else
            RsAux!InsContacto = Trim(tContacto.Text)
        End If
        
        If Trim(tTelefono.Text) <> "" Then
            RsAux!InsTelefono = Trim(tTelefono.Text)
        Else
            RsAux!InsTelefono = Null
        End If
        If caCobrar.Text <> 0 Then
            RsAux!InsDebeAbonarInst = caCobrar.Text
        Else
            RsAux!InsDebeAbonarInst = Null
        End If
        If caViatico.Text <> 0 Then
            RsAux!InsViatico = caViatico.Text
        Else
            RsAux!InsViatico = Null
        End If
        If Trim(SacoEnter(tMemo.Text)) <> "" Then
            RsAux!InsComentario = Trim(SacoEnter(tMemo.Text))
        Else
            RsAux!InsComentario = Null
        End If
        RsAux!InsFechaModificacion = Format(dNow, "yyyy/mm/dd hh:mm:ss")
        RsAux!InsUsuario = paCodigoDeUsuario
        
        If Val(lPrecio.Tag) > 0 Then
            RsAux!InsPrecioInstalacion = CCur(lPrecio.Tag)
        Else
            RsAux!InsPrecioInstalacion = Null
        End If
        RsAux.Update
        RsAux.Close
        
        'Si es nuevo tomo el mayor id.
        If Val(Toolbar1.Tag) = 1 Then
            
            Cons = "Select Max(InsID) From Instalacion " _
                & " Where InsDocumento = " & prmDocumento _
                & " And InsTipoDocumento = " & prmTipoDoc _
                & " And InsUsuario = " & paCodigoDeUsuario _
                & " And InsFechaModificacion = '" & Format(dNow, "yyyy/mm/dd hh:mm:ss") & "'"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            prmIDInst = RsAux(0)
            RsAux.Close
        End If
        
        If vsArticulo.Enabled Then
            '.......................................................................................................... Artículos
            Cons = "Delete RenglonInstalacion Where RInInstalacion = " & prmIDInst
            cBase.Execute (Cons)
            
            Cons = "Select * From RenglonInstalacion Where RInInstalacion = " & prmIDInst
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            For iCont = 0 To vsArticulo.Rows - 1
                If vsArticulo.Cell(flexcpValue, iCont, 0) > 0 Then
                    With vsArticulo
                        RsAux.AddNew
                        RsAux!RInInstalacion = prmIDInst
                        RsAux!RInArticulo = .Cell(flexcpData, iCont, 0)
                        RsAux!RInCantidad = .Cell(flexcpValue, iCont, 0)
                        RsAux!RInCobro = .Cell(flexcpData, iCont, 6) & IIf(.Cell(flexcpText, iCont, 7) <> "", "|" & .Cell(flexcpText, iCont, 7), "")
                        RsAux.Update
                    End With
                End If
            Next iCont
            RsAux.Close
            '..........................................................................................................
        End If
        
        'Grabo el telefono del cliente.
        If Trim(tTelefono.Text) <> "" And cTipoTelefono.ListIndex > -1 Then
            
            Cons = " Select * From Telefono " _
                    & " Where TelCliente = " & lIDCli _
                    & " And TelNumero = '" & Trim(tTelefono.Text) & "'"
                    
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            If RsAux.EOF Then
                
                RsAux.Close
                
                Cons = " Select * From Telefono " _
                    & " Where TelCliente = " & lIDCli _
                    & " And TelTipo = " & cTipoTelefono.ItemData(cTipoTelefono.ListIndex)
                
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If RsAux.EOF Then
                    RsAux.AddNew
                    RsAux!TelCliente = lIDCli
                    RsAux!TelTipo = cTipoTelefono.ItemData(cTipoTelefono.ListIndex)
                    RsAux!TelNumero = Trim(tTelefono.Text)
                    If Trim(tInterno.Text) <> "" Then RsAux!TelInterno = Trim(tInterno.Text) Else: RsAux!TelInterno = Null
                    RsAux.Update
                Else
                    RsAux.Edit
                    RsAux!TelNumero = Trim(tTelefono.Text)
                    If Trim(tInterno.Text) <> "" Then RsAux!TelInterno = Trim(tInterno.Text) Else: RsAux!TelInterno = Null
                    RsAux.Update
                End If
                RsAux.Close
            End If
        End If
        '.......................................................................................................... Grabo Telefono.
        
        If bSuceso Then
            clsGeneral.RegistroSuceso cBase, Now, 21, paCodigoDeTerminal, lUID, IIf(prmTipoDoc = 1, prmDocumento, 0), 0, "Se dio mensaje de no poder cumplir por pasar tope.", sDefSuc, , lIDCli
        End If
        
        cBase.CommitTrans
        
        cDireccion.Tag = ""
        On Error Resume Next
        AccionCancelar
        
    End If
    Exit Sub
    

errBT:
    clsGeneral.OcurrioError "No se logro iniciar la transacción.", Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub

ErrET:
    Resume ErrRoll
    Exit Sub

ErrRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al almacenar la información.", Err.Description
    Screen.MousePointer = 0
    
End Sub

Sub AccionEliminar()
Dim lUID As Long
Dim sDefSuc As String

    On Error GoTo errValido
    'Válido que si tiene remito los artículos no hayan sido retirados.
    If Val(lRemito.Tag) > 0 Then
        Cons = "Select * From RenglonRemito Where RReRemito = " & Val(lRemito.Tag) _
            & " And RReCantidad <> RReAEntregar"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            RsAux.Close
            MsgBox "No podrá eliminar la instalación ya que se entregó mercadería del remito.", vbExclamation, "Validación"
            Exit Sub
        End If
        RsAux.Close
    End If
    
    If MsgBox("Confirma anular la instalación?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
        
        Dim objSuceso As New clsSuceso
        With objSuceso
            .TipoSuceso = 21
            .ActivoFormulario paCodigoDeUsuario, "Eliminar Instalación", cBase
            lUID = .Usuario
            sDefSuc = .Defensa
        End With
        Set objSuceso = Nothing
        Me.Refresh
        If lUID = 0 Then Screen.MousePointer = 0: Exit Sub
        
        Screen.MousePointer = 11
        GetFechaServidor
        On Error GoTo errBT
        cBase.BeginTrans
                
        'Si hay remito lo borro y le pongo la cantidad a retirar.
        Cons = "Select * From Instalacion Where InsID = " & prmIDInst
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If CDate(lModificada.Tag) <> RsAux!InsFechaModificacion Then
                RsAux.Close
                cBase.RollbackTrans
                MsgBox "Otra terminal modificó la instalación, verifique.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        End If
        
        RsAux.Edit
        RsAux!InsAnulada = Format(Now, "yyyy-mm-dd hh:nn")
        RsAux.Update
        RsAux.Close
        
        'Elimino el remito si existe
        If Val(lRemito.Tag) > 0 Then del_BorroRemito
        '..................................................................................................
        
        'Registro el suceso
        If prmTipoDoc = 1 Then
            clsGeneral.RegistroSuceso cBase, Now, 21, paCodigoDeTerminal, lUID, prmDocumento, 0, "Se elimina instalación " & prmIDInst, sDefSuc, , lIDCli
        Else
            clsGeneral.RegistroSuceso cBase, Now, 21, paCodigoDeTerminal, lUID, 0, 0, "Se elimina instalación " & prmIDInst, sDefSuc, , lIDCli
        End If
        '..................................................................................................
        
        cBase.CommitTrans
        '..................................................................................................
        
        loc_CleanCtrl
        LoadInstalacion
        Screen.MousePointer = 0
        
    End If
    
    Exit Sub

errValido:
    clsGeneral.OcurrioError "Error al validar la eliminación.", Err.Description, "Eliminar Instalación"
    Screen.MousePointer = 0
    Exit Sub
    
errBT:
    clsGeneral.OcurrioError "No se pudo iniciar la transacción.", Err.Description
    Screen.MousePointer = vbDefault
    Exit Sub
ErrET:
    Resume ErrRoll
ErrRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se pudo eliminar la pregunta.", Err.Description
    Screen.MousePointer = 0
End Sub

Sub AccionCancelar()
On Error Resume Next
    Toolbar1.Tag = 0
    Screen.MousePointer = 11
    If Val(cDireccion.Tag) > 0 And Val(cDireccion.Tag) = cDireccion.ItemData(0) Then
        'Es una copia de algo que al final no se almacena.
        Cons = "Delete Direccion Where DirCodigo = " & Val(cDireccion.Tag)
        cBase.Execute (Cons)
    End If
    Botones True, Val(tID.Text) > 0, Val(tID.Text) > 0, False, False, Toolbar1, Me
    MnuOptIndependizarRemito.Enabled = (Val(tID.Text) > 0)
    loc_SetCtrl False
    
    If prmIDInst > 0 Then
        chHaceRemito.Value = 0
        LoadInstalacion
    Else
        loc_CleanCtrl
        FindDocumento prmTipoDoc, prmDocumento
    End If
    Toolbar1.Buttons("next").Enabled = UBound(arrInstDoc) > 1
    tID.SetFocus
    Screen.MousePointer = 0
End Sub

Private Sub tContacto_GotFocus()
    With tContacto
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese algún dato que haga referencia a quien contactar."
End Sub

Private Sub tContacto_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then tMemo.SetFocus
End Sub

Private Sub tContacto_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub tID_Change()
    If Val(Toolbar1.Tag) = 0 Then
        loc_CleanCtrl
        FindDocumento prmTipoDoc, prmDocumento
        Botones prmDocumento > 0, False, False, False, False, Toolbar1, Me
        Toolbar1.Buttons("next").Enabled = UBound(arrInstDoc) > 1
        MnuOptIndependizarRemito.Enabled = False
        prmIDInst = 0
    End If
End Sub

Private Sub tID_GotFocus()
    With tID
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.SimpleText = " Ingrese un código de instalación y presione <Enter> para buscarla."
End Sub

Private Sub tID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tID.Text) Then
            If prmIDInst = 0 Or prmIDInst <> Val(tID.Text) Then
                prmIDInst = Val(tID.Text)
                LoadInstalacion
                If prmIDInst = 0 Then MsgBox "No existe una instalación con ese código.", vbInformation, "ATENCIÓN"
                If prmIDInst = 0 And prmDocumento > 0 Then
                    Botones True, False, False, False, False, Toolbar1, Me
                    MnuOptIndependizarRemito.Enabled = False
                End If
                Toolbar1.Buttons("next").Enabled = UBound(arrInstDoc) > 1
            End If
        End If
    End If
End Sub

Private Sub tInterno_GotFocus()
    With tInterno
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese un interno o aclaración del teléfono."
End Sub

Private Sub tInterno_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then tContacto.SetFocus
End Sub

Private Sub tMemo_GotFocus()
    With tMemo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese algún comentario. [Enter - Graba]"
End Sub

Private Sub tMemo_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If vsArticulo.Enabled Then vsArticulo.SetFocus Else AccionGrabar
    End If
End Sub

Private Sub tMemo_LostFocus()
    Status.SimpleText = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar": AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
        Case "next": NextInst
        Case "viaticos": EjecutarApp "\\ibm3200\OyR\Programas\HtmlMatriz\ViaticosInstAC.htm"
    End Select

End Sub

Private Function loc_ValidateSave()
Dim iCont As Integer

    loc_ValidateSave = False
    
    If cInstalador.ListIndex = -1 Then
        MsgBox "Se debe seleccionar un instalador.", vbExclamation, "Validación"
        cInstalador.SetFocus: Exit Function
    End If
    
    
    If dpFecha.Value < Date And Val(Toolbar1.Tag) = 1 Then
        MsgBox "No es posible ingresar una fecha menor al día de hoy.", vbCritical, "ATENCIÓN"
        dpFecha.SetFocus
        Exit Function
    End If
    
'    If caCobrar.Tag = "1" And caCobrar.Text = 0 Then
'        If MsgBox("No se ingresó un valor a cobrar por la instalación y en la factura no se encontro un artículo que actue como el pago de la misma." & _
            vbCr & "¿Esta seguro de continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Posible Error") = vbNo Then
            
'            caCobrar.SetFocus
'            Exit Function
            
'        End If
'    End If
    
    If caCobrar.Text < 0 Or caViatico.Text < 0 Then
        If MsgBox("El valor a cobrar y/o el viático son negativos." & _
            vbCr & "¿Esta seguro de continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Posible Error") = vbNo Then
            
            caCobrar.SetFocus
            Exit Function
        End If
    End If
    
    
    If tDireccion.Text = "" Then
        MsgBox "La dirección es necesaria, seleccione o ingrese una.", vbCritical, "Validación"
        cDireccion.SetFocus: Exit Function
    End If
    
    Dim cCobro As Currency
    cCobro = 0
    For iCont = 0 To vsArticulo.Rows - 1
        If vsArticulo.Cell(flexcpBackColor, iCont, 0) = vbYellow Then
            MsgBox "Alguna fila tiene el cobro mal definido.", vbExclamation, "Validación"
            Exit Function
        End If
        cCobro = cCobro + Format((vsArticulo.Cell(flexcpData, iCont, 6) * vsArticulo.Cell(flexcpValue, iCont, 6)), "#,##0.00")
    Next
    
    If Format(cCobro, "#,##0.00") <> caCobrar.Text Then
        MsgBox "El campo cobrar debería ser " & Format(cCobro, "#,##0.00"), vbInformation, "Validación"
    End If
    
    'OJO EL TRUE ESTA DENTRO DEL FOR
    For iCont = 0 To vsArticulo.Rows - 1
        If vsArticulo.Cell(flexcpValue, iCont, 0) > 0 Then
            loc_ValidateSave = True
            Exit For
        End If
    Next
    
    If Not loc_ValidateSave Then
        MsgBox "Seleccione por lo menos un artículo a instalar.", vbExclamation, "Validación"
        On Error Resume Next
        vsArticulo.SetFocus
    End If
    
    
End Function

Private Sub loc_SetCtrl(ByVal bEdit As Boolean)
    
    tID.Enabled = Not bEdit
    cInstalador.Enabled = bEdit
    With dpFecha
        .Enabled = bEdit
        If .Enabled And .CustomFormat <> "dd/MM/yyyy" Then .CustomFormat = "dd/MM/yyyy": .Value = Date
    End With
    caCobrar.Enabled = bEdit
    caViatico.Enabled = bEdit
    cRangoHora.Enabled = bEdit
    bDireccion.Enabled = bEdit
    cDireccion.Enabled = bEdit
    cTipoTelefono.Enabled = bEdit
    tTelefono.Enabled = bEdit
    tInterno.Enabled = bEdit
    tContacto.Enabled = bEdit
    tMemo.Enabled = bEdit
    vsArticulo.Enabled = bEdit
    If bEdit Then
        tID.BackColor = vbButtonFace
        cInstalador.BackColor = vbWhite
    Else
        tID.BackColor = vbWhite
        cInstalador.BackColor = vbButtonFace
    End If
    caCobrar.BackColorDisplay = cInstalador.BackColor
    caViatico.BackColorDisplay = cInstalador.BackColor
    cRangoHora.BackColor = cInstalador.BackColor
    bDireccion.BackColor = cInstalador.BackColor
    cDireccion.BackColor = cInstalador.BackColor
    cTipoTelefono.BackColor = cInstalador.BackColor
    tTelefono.BackColor = cInstalador.BackColor
    tInterno.BackColor = cInstalador.BackColor
    tContacto.BackColor = cInstalador.BackColor
    tMemo.BackColor = cInstalador.BackColor
    With vsArticulo
        If .Enabled Then
            .BackColor = vbWhite
        Else
            .BackColor = vbButtonFace
        End If
        .BackColorBkg = .BackColor
    End With
    lAnulada.Visible = Not bEdit
    
End Sub

Private Sub loc_CleanCtrl()
        
    cInstalador.Text = "": cInstalador.Tag = ""
    With dpFecha
        .CustomFormat = "HH:mm"
        .Value = "00:00:00"
        .Tag = ""
    End With
    cRangoHora.Text = ""
    caViatico.Clean
    caCobrar.Clean: caCobrar.Tag = ""
    cDireccion.Clear: cDireccion.Tag = ""
    tDireccion.Text = ""
    cTipoTelefono.Text = ""
    tTelefono.Text = ""
    tInterno.Text = ""
    tContacto.Text = ""
    tMemo.Text = ""
    lDoc.Caption = ""
    vsArticulo.Rows = 0
    bDireccion.Tag = ""       'Guardo el id de la direccion particular.
    bDireccion.Enabled = ((cDireccion.ListIndex = 0) And cDireccion.Enabled)
    lIDCli = 0                      'Variable con el id del cliente
    
    lAlta.Caption = "Alta:"
    lModificada.Caption = "Modificada:": lModificada.Tag = ""
    lRemito.Caption = "": lRemito.Tag = ""
    chHaceRemito.Visible = False
    lAnulada.Caption = "": lAnulada.Visible = False
    lInstalada.Caption = "": lInstalada.Visible = False
    
    lPrecio.Caption = "": lPrecio.Tag = ""
    
    lbLiquidacion.Caption = ""
    lbLiquidacion.Visible = False

    
End Sub

Private Sub tTelefono_GotFocus()
    With tTelefono
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Status.SimpleText = "Ingrese un número de teléfono."
End Sub

Private Sub tTelefono_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then tInterno.SetFocus
End Sub

Private Sub tTelefono_LostFocus()
Dim aTexto As String
On Error Resume Next
    
    Status.SimpleText = ""
    If Trim(tTelefono.Text) <> "" Then
        aTexto = clsGeneral.RetornoFormatoTelefono(cBase, tTelefono.Text, 0)
        If aTexto <> "" Then
            tTelefono.Text = aTexto
        Else
            MsgBox "El teléfono ingresado no coincide con los formatos establecidos.", vbExclamation, "ATENCIÓN"
            Foco tTelefono
        End If
    End If
    
End Sub

Private Sub LoadInstalacion()
On Error GoTo errLI
Dim lAux As Long, lLiquidacion As Long
Dim bNew As Boolean
Dim iRetira As Integer
    
    Screen.MousePointer = 11
    tID.Text = prmIDInst
    prmIDInst = tID.Text
    If cInstalador.ListIndex > -1 Then loc_CleanCtrl
    lLiquidacion = 0
    bNew = False
    Cons = "Select * From Instalacion " _
            & " Left Outer Join TecnicoInstalador On TInCodigo = InsTecnico " _
            & " Left Outer Join Documento On InsRemito = DocCodigo And DocTipo = 6" _
        & " , RenglonInstalacion, Articulo" _
            & " Left Outer Join PrecioVigente On PViArticulo = ArtInstalacion And PViTipoCuota = " & prmTipoCuotaContado _
        & " Where InsID = " & prmIDInst & " And InsID = RInInstalacion And RInArticulo = ArtID"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        
        'Busco el documento, la función me retorna si el documento fue o no anulado.
        If FindDocumento(RsAux!InsTipoDocumento, RsAux!InsDocumento) <> 0 Then
            RsAux.Close
            MsgBox "El documento fue anulado.", vbExclamation, "ATENCIÓN"
            Toolbar1.Buttons("next").Enabled = UBound(arrInstDoc) > 1
            Exit Sub
        End If
        
        lAlta.Caption = "Alta: " & GetUserIdentity(RsAux!InsUsuarioAlta)
        lModificada.Caption = "Modificada: " & Format(RsAux!InsFechaModificacion, "d/mm/yy hh:nn") & " por " & GetUserIdentity(RsAux!InsUsuario)
        lModificada.Tag = RsAux!InsFechaModificacion
        If Not IsNull(RsAux!InsAnulada) Then lAnulada.Visible = True: lAnulada.Caption = "Anulada: " & Format(RsAux!InsAnulada, "dd/mm/yyyy")
        If Not IsNull(RsAux!InsFechaRealizada) Then
            lInstalada.Visible = True: lInstalada.Caption = "Realizada el " & Format(RsAux!InsFechaRealizada, "dd/mm/yyyy")
            If Not IsNull(RsAux!InsTecnico) Then lInstalada.Caption = lInstalada.Caption & " por " & Trim(RsAux!TInNombre)
        End If
        
        BuscoCodigoEnCombo cInstalador, RsAux!InsInstalador
        With dpFecha
            .CustomFormat = "dd/MM/yyyy": .Value = RsAux!InsFechaProm
            .Tag = .Value
        End With
        If Not IsNull(RsAux!InsRangoHora) Then cRangoHora.Text = Trim(RsAux!InsRangoHora)
        If Not IsNull(RsAux!InsDebeAbonarInst) Then caCobrar.Text = RsAux!InsDebeAbonarInst
        If Not IsNull(RsAux!InsViatico) Then caViatico.Text = RsAux!InsViatico
        loc_LoadComboDireccion
        cDireccion.ItemData(0) = RsAux!InsDireccion
        cDireccion.ListIndex = 0
        If tDireccion.Text = "" Then cDireccion_Click
        If Not IsNull(RsAux!InsTelefono) Then
            BuscoCodigoEnCombo cTipoTelefono, FindTipoTelefono(Trim(RsAux!InsTelefono))
            tTelefono.Text = Trim(RsAux!InsTelefono)
        End If
        If Not IsNull(RsAux!InsContacto) Then tContacto.Text = Trim(RsAux!InsContacto)
        If Not IsNull(RsAux!InsComentario) Then tMemo.Text = Trim(RsAux!InsComentario)
        
        If Not IsNull(RsAux!InsRemito) Then
            If Not IsNull(RsAux("DocCodigo")) Then
                lRemito.Caption = "Remito asociado: " & Trim(RsAux!DocSerie) & "-" & RsAux("DocNumero")
                lRemito.Tag = RsAux!InsRemito
            Else
                lRemito.Caption = "Remito asociado: " & Trim(RsAux!InsRemito)
                lRemito.Tag = RsAux!InsRemito
            End If
        End If
        
        If Not IsNull(RsAux!InsLiquidacion) Then
            lLiquidacion = RsAux!InsLiquidacion
            lbLiquidacion.Visible = True
            lbLiquidacion.Caption = " ID de liquidación: " & lLiquidacion
        End If
        
        If Not IsNull(RsAux!InsPrecioInstalacion) Then lPrecio.Caption = " de " & Format(RsAux!InsPrecioInstalacion, "#,##0.00"): lPrecio.Tag = RsAux!InsPrecioInstalacion
                
        'Cargo los artículos de la instalación.
        Do While Not RsAux.EOF
            With vsArticulo
                .AddItem Trim(RsAux!RInCantidad)
                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ArtNombre)
                lAux = RsAux!ArtID
                .Cell(flexcpData, .Rows - 1, 0) = lAux
                lAux = GetCantArtVendido(RsAux!RInArticulo, iRetira)
                .Cell(flexcpData, .Rows - 1, 1) = lAux
                .Cell(flexcpData, .Rows - 1, 2) = GetCantArtEnInstalacion(RsAux!RInArticulo)
                .Cell(flexcpText, .Rows - 1, 2) = iRetira
                .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!RInCantidad)   'Me guardo la cantidad que tengo en esta instalación.
                If .Cell(flexcpData, .Rows - 1, 1) > .Cell(flexcpData, .Rows - 1, 2) Then bNew = True
                
                'Guardo el artículo que cobra.
                If Not IsNull(RsAux!ArtInstalacion) Then
                    lAux = RsAux!ArtInstalacion
                Else
                    lAux = 0
                End If
                .Cell(flexcpData, .Rows - 1, 3) = lAux
                '............................................................Art. que cobra.
                                
                If Not lAnulada.Visible Then
                
                    'Lo que no se si el preciovigente vario.
                    If Not IsNull(RsAux!PViPrecio) Then
                        .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!PViPrecio, "###0.00")
                    Else
                        .Cell(flexcpText, .Rows - 1, 6) = 0
                    End If
                    
                    'Ahora busco en el documento la Q que se cobro en el mismo.
                    lAux = GetArtCobroDocumento(lAux)
                    .Cell(flexcpData, .Rows - 1, 4) = lAux
                                        
                    'Esto es el total cobrado del documento - lo usado en otras instalaciones
                    .Cell(flexcpData, .Rows - 1, 4) = lAux - db_GetQArtCobroUsadoEnInstalacion(.Cell(flexcpData, .Rows - 1, 3))
                    
                    'De este saldo tengo que determinar cuantos ya di en la lista.
                    '..............................................................................................
                    If .Cell(flexcpData, .Rows - 1, 4) > 0 Then
                        .Cell(flexcpData, .Rows - 1, 4) = .Cell(flexcpData, .Rows - 1, 4) - f_QOtroConCobro(.Rows - 1)
                    End If
                    'A este saldo abajo le tengo que sumar lo que tengo asignado en esta instalación.
                    '..............................................................................................
                                        
                                        
                    If IsNull(RsAux!RInCobro) Then
                        
                        If .Cell(flexcpData, .Rows - 1, 4) > 0 Then
                            'Pongo lo que tomo del documento.
                            If .Cell(flexcpData, .Rows - 1, 4) >= .Cell(flexcpValue, .Rows - 1, 0) Then
                                .Cell(flexcpData, .Rows - 1, 5) = .Cell(flexcpValue, .Rows - 1, 0)
                            Else
                                .Cell(flexcpData, .Rows - 1, 5) = .Cell(flexcpData, .Rows - 1, 4)
                            End If
                        End If
                        'Esto es la Q que estoy cobrando en la instalación.
                        .Cell(flexcpData, .Rows - 1, 6) = .Cell(flexcpValue, .Rows - 1, 0) - .Cell(flexcpData, .Rows - 1, 5)
                        
                    Else
                        'Aca tengo que presentar según sea lo almacenado.
                        'Formato del array es QCobrEnInstalacion|QEnDoc:IDDoc;QEnDoc2:IDDoc2
                        
                        'Dentro de esta rutina cargo el text7 y el Data7 y Data6
                        f_SetArtCobroParaInstalacion .Rows - 1, RsAux!RInCobro
                        
                        'Lo que esta asignado x el documento es el total - lo en otros doc - lo que esta en la instalación.
                        .Cell(flexcpData, .Rows - 1, 5) = .Cell(flexcpValue, .Rows - 1, 0) - .Cell(flexcpData, .Rows - 1, 6) - .Cell(flexcpData, .Rows - 1, 7)
                        .Cell(flexcpData, .Rows - 1, 4) = .Cell(flexcpData, .Rows - 1, 4) + .Cell(flexcpData, .Rows - 1, 5)
                    End If
                    If IsNull(RsAux!RInCobro) Then f_SetTotalCobro .Rows - 1
                End If
            End With
            RsAux.MoveNext
        Loop
    Else
        prmIDInst = 0
    End If
    RsAux.Close
    
    Dim bEdit As Boolean
    bEdit = Not lInstalada.Visible And lLiquidacion = 0 And Not lAnulada.Visible
    Botones prmIDInst > 0, prmIDInst > 0 And bEdit, prmIDInst > 0 And bEdit, False, False, Toolbar1, Me
    MnuOptIndependizarRemito.Enabled = (Val(lRemito.Tag) > 0)
    Toolbar1.Buttons("next").Enabled = UBound(arrInstDoc) > 1
    Screen.MousePointer = 0
    Exit Sub
    
errLI:
    clsGeneral.OcurrioError "Error al levantar la información de la instalación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub loc_AccesoPorDocumento()
On Error GoTo errAD
    
    '1ero veo si existe alguna instalación para el documento.
    
    Cons = "Select * From Instalacion " _
        & " Where InsTipoDocumento = " & prmTipoDoc _
        & " And InsDocumento = " & prmDocumento

    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        'Nuevo
        RsAux.Close
        SetNewInstalacion
    Else
        'Cargo los datos del documento y veo si hay posibilidad de hacer uno nuevo.
        prmIDInst = RsAux!InsID
        RsAux.Close
        LoadInstalacion
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errAD:
    clsGeneral.OcurrioError "Error al intentar armar la información para el documento.", Err.Description
    Screen.MousePointer = 0

End Sub

Private Function LoadArticuloDocumento() As Long
Dim lAux As Long
Dim iRetira As Integer
        
    vsArticulo.Rows = 0
    LoadArticuloDocumento = 0
    
    Cons = "Select * From Documento, Renglon, Articulo " _
            & " Left Outer Join PrecioVigente On PViArticulo = ArtInstalacion And PViTipoCuota = " & prmTipoCuotaContado _
        & " Where DocCodigo = " & prmDocumento _
        & " And ArtInstalador Is Not Null And DocCodigo = RenDocumento And RenArticulo = ArtID And RenCantidad > 0"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    Do While Not RsAux.EOF
        
        With vsArticulo
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ArtNombre)
            
            lAux = RsAux!ArtID
            .Cell(flexcpData, .Rows - 1, 0) = lAux
            
            'Invoco a este metodo ya que puedo tener notas en el medio.
            lAux = GetCantArtVendido(RsAux!ArtID, iRetira)
            .Cell(flexcpData, .Rows - 1, 1) = lAux
            .Cell(flexcpText, .Rows - 1, 0) = lAux
            
            .Cell(flexcpText, .Rows - 1, 2) = iRetira       'Para saber cuantos artículos van para un posible remito.
            .Cell(flexcpData, .Rows - 1, 2) = GetCantArtEnInstalacion(RsAux!ArtID)
            
            .Cell(flexcpText, .Rows - 1, 0) = .Cell(flexcpValue, .Rows - 1, 0) - .Cell(flexcpData, .Rows - 1, 2)
            
            If LoadArticuloDocumento = 0 Then LoadArticuloDocumento = RsAux!ArtInstalador
            
            'No quedan artículos x instalar.
            If .Cell(flexcpValue, .Rows - 1, 0) = 0 Then
                .RemoveItem .Rows - 1
            Else
                If .Rows = 1 Then caCobrar.Text = 0
                
                'Tengo que determinar la cantidad de artículos que quedan x cobrar.
                If Not IsNull(RsAux!PViPrecio) Then
                    .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!PViPrecio, "###0.00")
                Else
                    .Cell(flexcpText, .Rows - 1, 6) = 0
                End If
                
                'Guardo el artículo que cobra.
                If Not IsNull(RsAux!ArtInstalacion) Then
                    lAux = RsAux!ArtInstalacion
                Else
                    lAux = 0
                End If
                .Cell(flexcpData, .Rows - 1, 3) = lAux
                '............................................................Art. que cobra.
                .Cell(flexcpData, .Rows - 1, 6) = 0
                If lAux > 0 Then
                    'El total de artículos de cobro es la Q que estoy poniendo el text0
                    'Osea Data1 + Data2
                            
                    'Ahora busco en el documento la Q que se cobro en el mismo.
                    lAux = GetArtCobroDocumento(lAux)
                    
                    'Esto es el total cobrado del documento - lo usado en otras instalaciones
                    .Cell(flexcpData, .Rows - 1, 4) = lAux - db_GetQArtCobroUsadoEnInstalacion(.Cell(flexcpData, .Rows - 1, 3))
                    
                    'De este saldo tengo que determinar cuantos ya di en la lista.
                    '..............................................................................................
                    .Cell(flexcpData, .Rows - 1, 4) = .Cell(flexcpData, .Rows - 1, 4) - f_QOtroConCobro(.Rows - 1)
                    'Este es el saldo de cobrados que puedo utilizar.
                    '..............................................................................................
                    
                    If .Cell(flexcpData, .Rows - 1, 4) > 0 Then
                        'Pongo lo que tomo del documento.
                        If .Cell(flexcpData, .Rows - 1, 4) >= .Cell(flexcpValue, .Rows - 1, 0) Then
                            .Cell(flexcpData, .Rows - 1, 5) = .Cell(flexcpValue, .Rows - 1, 0)
                        Else
                            .Cell(flexcpData, .Rows - 1, 5) = .Cell(flexcpData, .Rows - 1, 4)
                        End If
                    End If
                    
                    'Esto es la Q que estoy cobrando en la instalación.
                    .Cell(flexcpData, .Rows - 1, 6) = .Cell(flexcpValue, .Rows - 1, 0) - .Cell(flexcpData, .Rows - 1, 5)
                End If
                caCobrar.Text = caCobrar.Text + (Val(.Cell(flexcpData, .Rows - 1, 6)) * .Cell(flexcpValue, .Rows - 1, 6))
            End If
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Function
Private Function f_QOtroConCobro(ByVal iRow As Integer) As Integer
    'Retorno la fila
Dim iCont As Integer
    
    f_QOtroConCobro = 0
    For iCont = 0 To vsArticulo.Rows - 1
        If iCont <> iRow Then
            If vsArticulo.Cell(flexcpData, iCont, 3) = vsArticulo.Cell(flexcpData, iRow, 3) Then
                f_QOtroConCobro = f_QOtroConCobro + vsArticulo.Cell(flexcpData, iCont, 5)
            End If
        End If
    Next iCont
    
End Function

Private Sub SetNewInstalacion()
'En esta opción no hay ninguna instalación para el documento.
    
    loc_CleanCtrl
    FindDocumento prmTipoDoc, prmDocumento
    'Invoco a botón Nuevo
    If prmDocumento > 0 Then
        AccionNuevo
    Else
        MsgBox "No es posible hacer una instalación para el documento.", vbExclamation, "Posible error"
    End If
    
End Sub

Private Sub loc_SetDefaultNew()
On Error Resume Next
    
    caViatico.Text = 0
    If vsArticulo.Rows > 0 Then
        'Armo la información x defecto para el instalador.
        If cInstalador.ListIndex > -1 Then
            If loc_SetFechaTipoFlete Then
                cInstalador.Tag = "1"   'Marco como fecha ingresada.
                If prmArticuloPagaInstalacion <> "" Then
                    SetPagaInstalacion
                    If caCobrar.Tag = "" Then
                        caViatico.SetFocus
                    Else
                        caCobrar.SetFocus
                    End If
                Else
                    caCobrar.SetFocus
                End If
            Else
                dpFecha.SetFocus
            End If
        Else
            cInstalador.SetFocus
        End If
    End If
    
End Sub

Private Sub SetPagaInstalacion()
    Cons = "Select * From Renglon Where RenArticulo IN (" & prmArticuloPagaInstalacion & ") And RenDocumento = " & prmDocumento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        'Tengo que sugerir un precio.
        caCobrar.Tag = "1"
    Else
        caCobrar.Tag = ""
    End If
    RsAux.Close
End Sub

Private Function loc_SetFechaTipoFlete() As Boolean
On Error GoTo errSF
Dim RsF As rdoResultset
    
Dim sMat As String
Dim iSumaDia As Integer
Dim dCierre As Date
Dim douAgenda As Double, douHabilitado As Double
    
    'retorno true si cargue
    loc_SetFechaTipoFlete = False
    
    cRangoHora.Clear
    If cInstalador.ListIndex = -1 Then Exit Function
    
    Cons = "Select TipoFlete.* From Instaladores, TipoFlete" _
        & " Where InsCodigo = " & cInstalador.ItemData(cInstalador.ListIndex) _
        & " And TFlAgenda Is Not Null And InsTipoFlete = TFlCodigo"
    
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsF.EOF Then
    
        If IsNull(RsF!TFlFechaAgeHab) Then dCierre = Date Else dCierre = RsF!TFlFechaAgeHab
        douAgenda = RsF!TFlAgenda
        If IsNull(RsF!TFlAgendaHabilitada) Then douHabilitado = -1 Else douHabilitado = RsF!TFlAgendaHabilitada
        
        If DateDiff("d", dCierre, Date) >= 7 Then
            'Como cerro hace + de una semana tomo la agenda normal.
            dCierre = Date
            sMat = superp_MatrizSuperposicion(douAgenda)
            iSumaDia = loc_BuscoProximoDia(dCierre, sMat)
        Else
            If douHabilitado > 0 Then
                If dCierre < Date Then dCierre = Date
                sMat = superp_MatrizSuperposicion(douHabilitado)
                iSumaDia = loc_BuscoProximoDia(dCierre, sMat)
                If iSumaDia = -1 Then
                    sMat = superp_MatrizSuperposicion(douAgenda)
                    iSumaDia = loc_BuscoProximoDia(dCierre, sMat)
                End If
            Else
                sMat = superp_MatrizSuperposicion(douAgenda)
                iSumaDia = loc_BuscoProximoDia(dCierre, sMat)
            End If
        End If
        
        If iSumaDia <> -1 Then
            dpFecha.Value = dCierre + iSumaDia
            loc_LoadRangoHoraTipoFlete sMat
            loc_SetFechaTipoFlete = True
        Else
            MsgBox "No existe una agenda ingresada para el instalador seleccionado.", vbInformation, "ATENCIÓN"
        End If
        
    End If
    RsF.Close
    
    Exit Function
errSF:
    clsGeneral.OcurrioError "Error al intentar poner la fecha de instalación.", Err.Description
End Function

Private Function loc_BuscoProximoDia(ByVal dFecha As Date, ByVal sMat As String)
Dim rsHora As rdoResultset
Dim iDia As Integer, iSuma As Integer
    
    'Por las dudas que no cumpla en la semana paso la agenda normal.
    
    On Error GoTo errBDER
    loc_BuscoProximoDia = -1
    
    'Consulto en base a la matriz devuelta.
    Cons = "Select Distinct(HFlDiaSemana) From HorarioFlete Where HFlIndice IN (" & sMat & ")"
    Set rsHora = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsHora.EOF Then
        
        'Busco el valor que coincida con el dia de hoy y ahí busco para arriba.
        iSuma = 0
        Do While iSuma < 7
            rsHora.MoveFirst
            iDia = Weekday(dFecha + iSuma)
            Do While Not rsHora.EOF
                If rsHora!HFlDiaSemana = iDia Then
                    loc_BuscoProximoDia = iSuma
                    GoTo Encontre
                End If
                rsHora.MoveNext
            Loop
            iSuma = iSuma + 1
        Loop
        rsHora.Close
    End If
    
Encontre:
    rsHora.Close
    Exit Function
    
errBDER:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el primer día disponible para el tipo de flete.", Trim(Err.Description)
End Function

Private Sub loc_LoadRangoHoraTipoFlete(Optional sMat As String = "")
On Error GoTo errCHEPD
Dim rsRH As rdoResultset
Dim dCierre As Date
Dim douAgenda As Double, douHabilitado As Double

    cRangoHora.Clear
    
    If sMat = "" Then
        Cons = "Select TipoFlete.* From Instaladores, TipoFlete" _
            & " Where InsCodigo = " & cInstalador.ItemData(cInstalador.ListIndex) _
            & " And TFlAgenda Is Not Null And InsTipoFlete = TFlCodigo"
            
        Set rsRH = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        
        If Not rsRH.EOF Then
            
            If IsNull(rsRH!TFlFechaAgeHab) Then dCierre = Date Else dCierre = rsRH!TFlFechaAgeHab
            douAgenda = rsRH!TFlAgenda
            If IsNull(rsRH!TFlAgendaHabilitada) Then douHabilitado = -1 Else douHabilitado = rsRH!TFlAgendaHabilitada
        
            If Abs(DateDiff("d", dCierre, dpFecha.Value)) >= 7 Or douHabilitado = 0 Then
                sMat = superp_MatrizSuperposicion(douAgenda)
            Else
                sMat = superp_MatrizSuperposicion(douHabilitado)
            End If
        End If
        rsRH.Close
    End If
    
    If sMat <> "" Then
        Cons = "Select HFlCodigo, HFlNombre From HorarioFlete Where HFlIndice IN (" & sMat & ")" _
            & " And HFlDiaSemana = " & Weekday(dpFecha.Value) & " Order by HFlInicio"
        CargoCombo Cons, cRangoHora
        If cRangoHora.ListCount > 0 Then cRangoHora.ListIndex = 0
    End If
    Exit Sub
    
errCHEPD:
    clsGeneral.OcurrioError "Error al buscar los horarios para el día de semana.", Trim(Err.Description)
End Sub

Private Sub loc_LoadComboDireccion()
On Error GoTo errLD
Dim rsDA As rdoResultset
Dim lDirP As Long
    
    cDireccion.Clear
    
    'Este item es el que me muestra la dirección que tengo cargada en la copia.
    cDireccion.AddItem "(Dirección a Instalar)": cDireccion.ItemData(cDireccion.NewIndex) = 0
    
    lDirP = 0
    Cons = "Select CliDireccion From CLiente Where CliCodigo = " & lIDCli
    Set rsDA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsDA.EOF Then
        If Not IsNull(rsDA(0)) Then lDirP = rsDA!CliDireccion
    End If
    rsDA.Close
    If lDirP > 0 Then
        cDireccion.AddItem "(Particular)": cDireccion.ItemData(cDireccion.NewIndex) = lDirP
        bDireccion.Tag = lDirP
    Else
        bDireccion.Tag = ""
    End If
        
    'Direcciones Auxiliares-----------------------------------------------------------------------
    Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & lIDCli
    Set rsDA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsDA.EOF Then
        Do While Not rsDA.EOF
            cDireccion.AddItem Trim(rsDA!DAuNombre)
            cDireccion.ItemData(cDireccion.NewIndex) = rsDA!DAuDireccion
            rsDA.MoveNext
        Loop
    End If
    rsDA.Close
    
    cDireccion.AddItem "(Últimos envíos)"
    cDireccion.ItemData(cDireccion.NewIndex) = -2
    
    If cDireccion.ListCount > 1 Then cDireccion.BackColor = vbWhite
    Exit Sub
    
errLD:
    clsGeneral.OcurrioError "Error al cargar las direcciones del cliente.", Err.Description
End Sub

Private Sub loc_LoadTelefono()
Dim RsTel As rdoResultset
    
    On Error GoTo ErrNT
    If cTipoTelefono.ListIndex = -1 Then Exit Sub
        
    Screen.MousePointer = 11
    Cons = " Select TelTipo, TelNumero, TelInterno " _
            & " From Telefono Where TelCliente = " & lIDCli & " And TelTipo = " & cTipoTelefono.ItemData(cTipoTelefono.ListIndex)
    Set RsTel = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsTel.EOF Then
        tTelefono.Text = Trim(RsTel!TelNumero)
        If Not IsNull(RsTel!TelInterno) Then tInterno.Text = Trim(RsTel!TelInterno) Else: tInterno.Text = ""
    End If
    RsTel.Close
    Screen.MousePointer = 0
    Exit Sub
        
ErrNT:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Error al cargar los datos del número de teléfono.", Err.Description
End Sub

Private Function CopyDireccion(ByVal lnCodDireccion As Long) As Long
Dim lIdCalle As Long, lNroPuerta As Long
Dim RsDO As rdoResultset, RsDC As rdoResultset
    
    'Copio la Direccion
    Screen.MousePointer = 11
    On Error GoTo errBT
    CopyDireccion = 0
    
    'Direccion ORIGINAL
    Cons = "Select * from Direccion Where DirCodigo = " & lnCodDireccion
    Set RsDO = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    'Direccion COPIA
    Cons = "Select * from Direccion Where DirCodigo = " & lnCodDireccion
    Set RsDC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    RsDC.AddNew
    If Not IsNull(RsDO!DirComplejo) Then RsDC!DirComplejo = RsDO!DirComplejo
    
    RsDC!DirCalle = RsDO!DirCalle
    lIdCalle = RsDO!DirCalle
    
    RsDC!DirPuerta = RsDO!DirPuerta
    lNroPuerta = RsDO!DirPuerta
    
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
                
    Cons = "Select Max(DirCodigo) from Direccion Where DirCalle = " & lIdCalle _
        & " And DirPuerta = " & lNroPuerta
    Set RsDC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    CopyDireccion = RsDC(0)
    RsDC.Close
    
    Screen.MousePointer = vbDefault
    Exit Function
    
errBT:
    Screen.MousePointer = vbDefault
    Exit Function

End Function

Private Sub LoadTelefonoDefault()
Dim rsTD As rdoResultset
Dim sTel As String, sInterno As String, lTipo As Long
    
    lTipo = 0
    Cons = " Select TelTipo, TelNumero, IsNull(TelInterno, '') as TelInterno" _
            & " From Telefono Where TelCliente = " & lIDCli & " And TelTipo IN(" & prmTipoTelefP & ", " & prmTipoTelefE & ")"

    Set rsTD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsTD.EOF
        sTel = Trim(rsTD!TelNumero)
        sInterno = Trim(rsTD!TelInterno)
        lTipo = rsTD!TelTipo
        If rsTD!TelTipo = prmTipoTelefP Then Exit Do
        rsTD.MoveNext
    Loop
    rsTD.Close
    
    If lTipo > 0 Then
        BuscoCodigoEnCombo cTipoTelefono, lTipo
        tTelefono.Text = sTel
        tInterno.Text = sInterno
    Else
        cTipoTelefono.Text = ""
        tTelefono.Text = ""
        tInterno.Text = ""
    End If
    
End Sub

Private Sub vsArticulo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Col <> 1 Then Cancel = True: Exit Sub
    
    'No hay nada para cobrar.
    If vsArticulo.Cell(flexcpData, Row, 5) = vsArticulo.Cell(flexcpValue, Row, 0) And vsArticulo.Cell(flexcpData, Row, 6) = 0 And vsArticulo.Cell(flexcpData, Row, 7) = 0 Then Cancel = True: Exit Sub
    
    vsArticulo.ComboList = "..."
    
End Sub

Private Sub vsArticulo_CellButtonClick(ByVal Row As Long, ByVal Col As Long)

    If Col = 1 Then
        
        With frmDetCobro
            .prmCliente = lIDCli
            .prmInstalacion = Val(tID.Text)
            If prmTipoDoc = 1 Then .prmDocInstalacion = prmDocumento
            .prmIDArticuloInstalacion = vsArticulo.Cell(flexcpData, Row, 3)
            .prmQNecesitoCobrar = vsArticulo.Cell(flexcpValue, Row, 0) - vsArticulo.Cell(flexcpData, Row, 5)
            .prmQInstalacion = vsArticulo.Cell(flexcpData, Row, 6)
            .prmEnOtrosDocumentos = vsArticulo.Cell(flexcpText, Row, 7)
                        
            .Show vbModal, Me
            
            If .bSeteo Then
                
                vsArticulo.Cell(flexcpText, Row, 7) = .prmEnOtrosDocumentos
                vsArticulo.Cell(flexcpData, Row, 6) = .prmQInstalacion
                vsArticulo.Cell(flexcpData, Row, 7) = .QEnLista
                
                f_SetTotalInstalacion
            End If
            Set frmDetCobro = Nothing
            vsArticulo.Select Row, 0
        End With
    End If
    
End Sub

Private Sub vsArticulo_GotFocus()
    Status.SimpleText = "Seleccione los artículos a instalar. [+ ó -] Agrega o elimina, [Supr] Quitar de la lista"
End Sub

Private Sub vsArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    Select Case KeyCode
        Case vbKeyAdd
            '6/10/2007 agregue la condición siempre y cuando el texto este visible.
            If Shift <> 0 Or (lRemito.Caption <> "" And Not chHaceRemito.Visible) Then Exit Sub    'Or lRemito.Caption <> "" Then Exit Sub
            If Toolbar1.Tag = "2" Then
                'Tengo que ver que la cantidad que tengo es la disponible.
                'tengo en la celda 3 lo que esta en esta instalación, en el data 2 lo que hay en instalaciones, (x eso resto lo que esta en instalaciones - lo de esta instalación)
                If vsArticulo.Cell(flexcpValue, vsArticulo.Row, 0) < vsArticulo.Cell(flexcpData, vsArticulo.Row, 1) - (vsArticulo.Cell(flexcpData, vsArticulo.Row, 2) - vsArticulo.Cell(flexcpValue, vsArticulo.Row, 3)) Then
                    vsArticulo.Cell(flexcpText, vsArticulo.Row, 0) = vsArticulo.Cell(flexcpValue, vsArticulo.Row, 0) + 1
                    f_SetTotalCobro vsArticulo.Row
                    f_SetTotalInstalacion
                End If
            Else
                If vsArticulo.Cell(flexcpValue, vsArticulo.Row, 0) < vsArticulo.Cell(flexcpData, vsArticulo.Row, 1) - vsArticulo.Cell(flexcpData, vsArticulo.Row, 2) Then
                    vsArticulo.Cell(flexcpText, vsArticulo.Row, 0) = vsArticulo.Cell(flexcpValue, vsArticulo.Row, 0) + 1
                    f_SetTotalCobro vsArticulo.Row
                    f_SetTotalInstalacion
                End If
            End If
            
        Case vbKeySubtract
            '6/10/2007 agregue la condición siempre y cuando el texto este visible.
            If lRemito.Caption <> "" Then
                Exit Sub
            End If
            
            If vsArticulo.Cell(flexcpValue, vsArticulo.Row, 0) > 0 And vsArticulo.Rows > 0 Then
                vsArticulo.Cell(flexcpText, vsArticulo.Row, 0) = vsArticulo.Cell(flexcpValue, vsArticulo.Row, 0) - 1
                f_SetTotalCobro vsArticulo.Row
                f_SetTotalInstalacion
            End If
            
        Case vbKeyDelete
            '6/10/2007 agregue la condición siempre y cuando el texto este visible.
            If Shift <> 0 Or (lRemito.Caption <> "" And Not chHaceRemito.Visible) Then Exit Sub    'Or lRemito.Caption <> "" Then Exit Sub
            If vsArticulo.Rows > 0 Then
                
                'Hago esto para ver si tengo artículos con cobro que puedan usar lo que tengo asignado x artículos cobrados en el documento
                vsArticulo.Cell(flexcpText, vsArticulo.Row, 0) = 0
                f_SetTotalCobro vsArticulo.Row
                '.........................................................................................................
                
                vsArticulo.RemoveItem vsArticulo.Row
                f_SetTotalInstalacion
            Else
                If MsgBox("Desea cancelar la instalación?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then AccionCancelar
            End If
    End Select
End Sub

Private Sub vsArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then AccionGrabar
End Sub

Private Function GetFechaServidor() As String

    Dim RsF As rdoResultset
    On Error GoTo errFecha
    Cons = "Select GetDate()"
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    GetFechaServidor = RsF(0)
    RsF.Close
    
    Time = GetFechaServidor
    Date = GetFechaServidor
    Exit Function

errFecha:
    GetFechaServidor = Now
End Function

Private Function SaveRemito(ByVal dFecha As Date) As Long
Dim rsRem As rdoResultset
Dim iCont As Integer

    Cons = "Select * From Remito Where RemDocumento = " & prmDocumento
    Set rsRem = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rsRem.AddNew
    rsRem!RemDocumento = prmDocumento
    rsRem!RemFecha = Format(dFecha, "yyyy/mm/dd hh:mm:ss")
    rsRem!RemModificado = Format(dFecha, "yyyy/mm/dd hh:mm:ss")
    rsRem!RemUsuario = paCodigoDeUsuario
    rsRem.Update
    
    Cons = "Select Max(RemCodigo) From Remito Where RemDocumento = " & prmDocumento
    Set rsRem = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    SaveRemito = rsRem(0)
    rsRem.Close

    '------------------------------------------------------------------------------------------------RENGLON-REMITO
    Cons = "Select * from RenglonRemito Where RReRemito = " & SaveRemito
    Set rsRem = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    With vsArticulo
        For iCont = 0 To .Rows - 1
            If .Cell(flexcpValue, iCont, 2) > 0 And .Cell(flexcpValue, iCont, 0) > 0 Then
                rsRem.AddNew
                rsRem!RReRemito = SaveRemito
                rsRem!RReArticulo = .Cell(flexcpData, iCont, 0)
                
                'Si lo que hay a retirar es mayor a lo que instala
                If .Cell(flexcpValue, iCont, 2) > .Cell(flexcpValue, iCont, 0) Then
                    rsRem!RReCantidad = .Cell(flexcpValue, iCont, 0)
                Else
                    rsRem!RReCantidad = .Cell(flexcpValue, iCont, 2)
                End If
                rsRem!RReAEntregar = rsRem!RReCantidad
                rsRem.Update
                
                '--------------------------------------------------------------------------------------Updateo tabla RenglonDocumento
                If .Cell(flexcpValue, iCont, 2) > .Cell(flexcpValue, iCont, 0) Then
                    Cons = "Update Renglon Set RenARetirar = RenARetirar - " & .Cell(flexcpValue, iCont, 0)
                Else
                    Cons = "Update Renglon Set RenARetirar = RenARetirar - " & .Cell(flexcpValue, iCont, 2)
                End If
                Cons = Cons & " Where RenDocumento = " & prmDocumento _
                                    & " And RenArticulo = " & .Cell(flexcpData, iCont, 0)
                                    
                cBase.Execute (Cons)
                '--------------------------------------------------------------------------------------Updateo tabla RenglonDocumento
            End If
        Next
    End With
    rsRem.Close
    '-----------------------------------------------------------------------------------------------RENGLON-REMITO

    '-----------------------------------------------------------------------------------------------Fecha Modificado al Documento
    Cons = "Select * from Documento where DocCodigo = " & prmDocumento
    Set rsRem = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    rsRem.Edit
    rsRem!DocFModificacion = Format(dFecha, "yyyy/mm/dd hh:mm:ss")
    rsRem.Update
    rsRem.Close
    '----------------------------------------------------------------------------------------------Fecha Modificado al Documento
    
End Function

Private Function GetRangoHora() As String
Dim rsRH As rdoResultset

    Cons = "Select * from CodigoTexto Where Codigo = " & cRangoHora.ItemData(cRangoHora.ListIndex)
    Set rsRH = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Len(Trim(rsRH!Clase)) < 4 Then
        GetRangoHora = "0" & Trim(rsRH!Clase) & "-" & Trim(rsRH!Puntaje)
    Else
        GetRangoHora = Trim(rsRH!Clase) & "-" & Trim(rsRH!Puntaje)
    End If
    rsRH.Close
    
End Function

Private Function FindTipoTelefono(ByVal sTelef As String) As Long
Dim rsTT As rdoResultset
On Error GoTo errFTT

    FindTipoTelefono = -1
    Cons = "Select * From Telefono Where TelCliente = " & lIDCli & " and TelNumero = '" & sTelef & "'"
    Set rsTT = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsTT.EOF Then FindTipoTelefono = rsTT!TelTipo
    rsTT.Close
errFTT:
End Function

Private Function GetUserIdentity(ByVal lCod As Long)
Dim rsUID As rdoResultset
On Error Resume Next
    GetUserIdentity = ""
    Cons = "Select * From Usuario Where UsuCodigo = " & lCod
    Set rsUID = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsUID.EOF Then GetUserIdentity = Trim(rsUID!UsuIdentificacion)
    rsUID.Close
End Function

Private Function GetCantArtEnInstalacion(ByVal lArt As Long) As Integer
On Error GoTo errAEI
Dim rsC As rdoResultset
    
    GetCantArtEnInstalacion = 0
    Cons = "Select Sum(RInCantidad) From Instalacion, RenglonInstalacion " _
        & " Where InsDocumento = " & prmDocumento _
        & " And InsTipoDocumento = " & prmTipoDoc _
        & " And RInArticulo = " & lArt _
        & " And InsAnulada Is Null And InsID = RInInstalacion"
    Set rsC = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsC.EOF Then
        If Not IsNull(rsC(0)) Then GetCantArtEnInstalacion = rsC(0)
    End If
    rsC.Close
    Exit Function
    
errAEI:
    clsGeneral.OcurrioError "Error al buscar.", Err.Description, "Artículos en Instalación"
End Function

Private Function GetCantArtVendido(ByVal lArt As Long, ByRef iRetira As Integer) As Integer
Dim rsA As rdoResultset
    
    GetCantArtVendido = 0
    iRetira = 0
    If prmTipoDoc = 1 Then
        Cons = "Select RenCantidad, RenARetirar  From Renglon Where RenDocumento = " & prmDocumento _
            & " And RenArticulo = " & lArt
    Else
        Cons = "Select RVTCantidad as RenCantidad , RVTARetirar as 'RenARetirar'  From RenglonVTaTelefonica Where RVTVentaTelefonica = " & prmDocumento _
            & " And RVTArticulo = " & lArt
    End If
    
    Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsA.EOF Then
        GetCantArtVendido = rsA("RenCantidad")
        iRetira = rsA("RenARetirar")
    End If
    rsA.Close
    
    If prmTipoDoc <> 1 Then Exit Function
    
    'Ahora saco la cantidad que pueden estar en notas.
    If GetCantArtVendido > 0 Then
        Cons = "Select IsNull(Sum(RenCantidad), 0) From Renglon, Nota Where NotFactura = " & prmDocumento _
            & " And RenArticulo = " & lArt & " And RenDocumento = NotNota"
        Set rsA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not rsA.EOF Then
            GetCantArtVendido = GetCantArtVendido - rsA(0)
        End If
        rsA.Close
    End If
    
End Function

Private Sub HaceRemito()
Dim iCont As Integer
    For iCont = 0 To vsArticulo.Rows - 1
        If vsArticulo.Cell(flexcpValue, iCont, 2) > 0 Then chHaceRemito.Visible = True
    Next iCont
End Sub

Private Sub del_BorroRemito()
Dim rsR As rdoResultset

    Cons = "Select * From RenglonRemito Where RReRemito = " & Val(lRemito.Tag)
    Set rsR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsR.EOF
        Cons = "Update Renglon Set RenARetirar = RenARetirar + " & rsR("RReAEntregar") _
            & " Where RenDocumento = " & prmDocumento & " And RenArticulo = " & rsR("RReArticulo")
        cBase.Execute (Cons)
        rsR.Delete
        rsR.MoveFirst
    Loop
    rsR.Close
    
    If prmTipoDoc = 1 Then
        Cons = "Select * From Documento Where DocCodigo = " & prmDocumento
        Set rsR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        rsR.Edit
        rsR!DocFModificacion = Format(Now, "yyyy-mm-dd hh:nn:ss")
        rsR.Update
        rsR.Close
    End If
    
    Cons = "Select * From Remito Where RemCodigo = " & Val(lRemito.Tag)
    Set rsR = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsR.EOF Then rsR.Delete
    rsR.Close
    
End Sub

Private Function FindDocumento(ByVal byTipo As Byte, ByVal lDocCod As Long) As Integer
Dim rsD As rdoResultset
On Error GoTo errFD
    
    lDoc.Caption = "": lDoc.Tag = ""
    Erase arrInstDoc
    ReDim arrInstDoc(0)
    
    prmDocumento = lDocCod
    prmTipoDoc = byTipo
    lIDCli = 0
    FindDocumento = 0
    If byTipo = 1 Then
        Cons = "Select SucAbreviacion, DocCliente as 'Cli', DocCodigo as 'Cod', DocSerie as 'Serie', DocNumero as 'Nro', DocTipo as 'Tipo', DocAnulado, DocFModificacion as 'FM' From Documento, Sucursal " _
            & " Where DocCodigo = " & prmDocumento _
            & " And DocSucursal = SucCodigo"
    Else
        Cons = "Select SucAbreviacion, VTeCliente as 'Cli', VTeCodigo as 'Cod', '' as 'Serie', VTeCodigo as 'Nro', 7 as 'Tipo', VTeAnulado, VTeFModificacion as 'FM' From VentaTelefonica, Sucursal " _
            & " Where VTeCodigo = " & prmDocumento _
            & " And VTeSucursal = SucCodigo"
    End If
    Set rsD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsD.EOF Then
        If byTipo = 1 Then
            If rsD!DocAnulado <> 0 Then
                FindDocumento = 1
                rsD.Close
                Exit Function
            End If
        Else
            If Not IsNull(rsD!VTeAnulado) Then
                FindDocumento = 1
                rsD.Close
                Exit Function
            End If
        End If
        lIDCli = rsD!Cli
        lDoc.Caption = Trim(rsD!SucAbreviacion)
        If rsD!Tipo = 1 Then
            lDoc.Caption = lDoc.Caption & " Contado "
        ElseIf rsD!Tipo = 2 Then
            lDoc.Caption = lDoc.Caption & " Crédito "
        Else
            lDoc.Caption = lDoc.Caption & " Venta Telefónica "
        End If
        lDoc.Caption = " " & lDoc.Caption & Trim(rsD!Serie) & " " & Trim(rsD!Nro)
        'Guardo fecha de modificacion
        lDoc.Tag = rsD("FM")
    Else
        prmDocumento = 0
    End If
    rsD.Close
    
    'Cargo en menú todas las instalaciones
    Erase arrInstDoc
    ReDim arrInstDoc(0)
    If prmDocumento > 0 Then
        Cons = "Select * From Instalacion Where InsDocumento = " & prmDocumento
        Set rsD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not rsD.EOF
            ReDim Preserve arrInstDoc(UBound(arrInstDoc) + 1)
            arrInstDoc(UBound(arrInstDoc)) = rsD!InsID
            rsD.MoveNext
        Loop
        rsD.Close
    End If
    Exit Function
    
errFD:
    clsGeneral.OcurrioError "Error al buscar la relación del documento.", Err.Description
End Function
Private Function LoadArticuloVenta() As Long
Dim lAux As Long
Dim iRetira As Integer
    
    vsArticulo.Rows = 0
        
    LoadArticuloVenta = 0
    Cons = "Select * From VentaTelefonica, RenglonVtaTelefonica, Articulo " _
            & " Left Outer Join PrecioVigente On PViArticulo = ArtInstalacion And PViTipoCuota = " & prmTipoCuotaContado _
        & " Where VTeCodigo = " & prmDocumento _
        & " And ArtInstalador Is Not Null And VteCodigo = RVTVentaTelefonica And RVTArticulo = ArtID"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    Do While Not RsAux.EOF
        With vsArticulo
            .AddItem ""
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!ArtNombre)
            
            lAux = RsAux!ArtID
            .Cell(flexcpData, .Rows - 1, 0) = lAux
            
            'Invoco a este metodo ya que puedo tener notas en el medio.
            lAux = RsAux!RVTCantidad
            .Cell(flexcpData, .Rows - 1, 1) = lAux
            .Cell(flexcpText, .Rows - 1, 0) = lAux
                        
            .Cell(flexcpText, .Rows - 1, 2) = RsAux!RVTARetirar
            .Cell(flexcpData, .Rows - 1, 2) = GetCantArtEnInstalacion(RsAux!ArtID)
            
            .Cell(flexcpText, .Rows - 1, 0) = .Cell(flexcpValue, .Rows - 1, 0) - .Cell(flexcpData, .Rows - 1, 2)
            
            If LoadArticuloVenta = 0 Then LoadArticuloVenta = RsAux!ArtInstalador
            
            'No quedan artículos x instalar.
            If .Cell(flexcpValue, .Rows - 1, 0) = 0 Then
                .RemoveItem .Rows - 1
            Else
                If .Rows = 1 Then caCobrar.Text = 0
                
                'Tengo que determinar la cantidad de artículos que quedan x cobrar.
                If Not IsNull(RsAux!PViPrecio) Then
                    .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!PViPrecio, "###0.00")
                Else
                    .Cell(flexcpText, .Rows - 1, 6) = 0
                End If
                
                'Guardo el artículo que cobra.
                If Not IsNull(RsAux!ArtInstalacion) Then
                    lAux = RsAux!ArtInstalacion
                Else
                    lAux = 0
                End If
                .Cell(flexcpData, .Rows - 1, 3) = lAux
                '............................................................Art. que cobra.
                            
                .Cell(flexcpData, .Rows - 1, 6) = 0
                If lAux > 0 Then
                    'El total de artículos de cobro es la Q que estoy poniendo el text0
                    'Osea Data1 + Data2
                            
                    'Ahora busco en el documento la Q que se cobro en el mismo.
                    lAux = GetArtCobroDocumento(lAux)
                    
                    'Esto es el total cobrado del documento - lo usado en otras instalaciones
                    .Cell(flexcpData, .Rows - 1, 4) = lAux - db_GetQArtCobroUsadoEnInstalacion(.Cell(flexcpData, .Rows - 1, 3))
                    
                    'De este saldo tengo que determinar cuantos ya di en la lista.
                    '..............................................................................................
                    .Cell(flexcpData, .Rows - 1, 4) = .Cell(flexcpData, .Rows - 1, 4) - f_QOtroConCobro(.Rows - 1)
                    'Este es el saldo de cobrados que puedo utilizar.
                    '..............................................................................................
                    
                    If .Cell(flexcpData, .Rows - 1, 4) > 0 Then
                        'Pongo lo que tomo del documento.
                        If .Cell(flexcpData, .Rows - 1, 4) >= .Cell(flexcpValue, .Rows - 1, 0) Then
                            .Cell(flexcpData, .Rows - 1, 5) = .Cell(flexcpValue, .Rows - 1, 0)
                        Else
                            .Cell(flexcpData, .Rows - 1, 5) = .Cell(flexcpData, .Rows - 1, 4)
                        End If
                    End If
                    
                    'Esto es la Q que estoy cobrando en la instalación.
                    .Cell(flexcpData, .Rows - 1, 6) = .Cell(flexcpValue, .Rows - 1, 0) - .Cell(flexcpData, .Rows - 1, 5)
                End If
                caCobrar.Text = caCobrar.Text + (Val(.Cell(flexcpData, .Rows - 1, 6)) * .Cell(flexcpValue, .Rows - 1, 6))
            End If
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close

End Function

Private Function ValidoRangoHorario() As Boolean

    ValidoRangoHorario = True
    If cRangoHora.ListIndex > -1 Then Exit Function
    
    If InStr(1, cRangoHora.Text, "-") > 0 Then
        Select Case Len(cRangoHora.Text)
            Case 9
                If CLng(Mid(cRangoHora.Text, 1, InStr(1, cRangoHora.Text, "-") - 1)) > CLng(Mid(cRangoHora.Text, InStr(1, cRangoHora.Text, "-") + 1, Len(cRangoHora.Text))) Then
                    MsgBox "El rango de horario ingresado no es válido.", vbExclamation, "ATENCIÓN"
                    cRangoHora.SetFocus
                    ValidoRangoHorario = False
                    Exit Function
                End If
                
            Case 5
                If InStr(1, cRangoHora.Text, "-") = 1 Then
                    If CLng(Mid(cRangoHora.Text, InStr(1, cRangoHora.Text, "-") + 1, Len(cRangoHora.Text))) < prmPrimeraHoraEnvio Then
                        MsgBox "El horario ingresado es menor a la primera hora de entrega.", vbExclamation, "ATENCIÓN"
                        ValidoRangoHorario = False
                        Exit Function
                    Else
                        If prmPrimeraHoraEnvio < 1000 Then
                            cRangoHora.Text = "0" & prmPrimeraHoraEnvio & cRangoHora.Text
                        Else
                            cRangoHora.Text = prmPrimeraHoraEnvio & cRangoHora.Text
                        End If
                        Exit Function
                    End If
                Else
                    If InStr(1, cRangoHora.Text, "-") = 5 Then
                        If CLng(Mid(cRangoHora.Text, 1, InStr(1, cRangoHora.Text, "-") - 1)) > prmUltimaHoraEnvio Then
                            MsgBox "El horario ingresado es mayor que la última hora de envio.", vbExclamation, "ATENCIÓN"
                            ValidoRangoHorario = False
                            Exit Function
                        Else
                            cRangoHora.Text = cRangoHora.Text & prmUltimaHoraEnvio
                        End If
                    Else
                        MsgBox "No se ingreso un horario válido. [####-####]", vbExclamation, "ATENCIÓN"
                        cRangoHora.SetFocus
                        ValidoRangoHorario = False
                        Exit Function
                    End If
                End If
            
            Case 8
                If CLng(Mid(cRangoHora.Text, 1, InStr(1, cRangoHora.Text, "-") - 1)) > CLng(Mid(cRangoHora.Text, InStr(1, cRangoHora.Text, "-") + 1, Len(cRangoHora.Text))) Then
                    MsgBox "El rango de horario ingresado no es válido.", vbExclamation, "ATENCIÓN"
                    cRangoHora.SetFocus
                    ValidoRangoHorario = False
                    Exit Function
                End If
                
                If InStr(1, cRangoHora.Text, "-") = 4 Then
                    cRangoHora.Text = "0" & cRangoHora.Text
                End If
            
            Case Else
                    MsgBox "No se ingreso un horario válido. [####-####]", vbExclamation, "ATENCIÓN"
                    cRangoHora.SetFocus
                    ValidoRangoHorario = False
                    Exit Function
                    
        End Select
    Else
        MsgBox "No se ingreso un horario válido. [####-####]", vbExclamation, "ATENCIÓN"
        cRangoHora.SetFocus
        ValidoRangoHorario = False
        Exit Function
    End If
    
End Function

Private Sub NextInst()
On Error Resume Next
Dim iPos As Integer
    iPos = 1
    For I = 1 To UBound(arrInstDoc)
        If Val(arrInstDoc(I)) = Val(tID.Text) Then
            iPos = I
        End If
    Next
    If iPos = UBound(arrInstDoc) Then
        iPos = 1
    Else
        iPos = iPos + 1
    End If
    tID.Text = arrInstDoc(iPos)
    tID_KeyPress vbKeyReturn
    Foco tID
End Sub

Private Function GetArtCobroDocumento(ByVal lArtCobro As Long) As Integer
On Error GoTo errGACD
Dim rsQ As rdoResultset

    GetArtCobroDocumento = 0
    If lArtCobro = 0 Then Exit Function
    If prmTipoDoc = 1 Then
        Cons = "Select RenCantidad From Renglon Where RenDocumento = " & prmDocumento _
            & " And RenArticulo = " & lArtCobro
    Else
        Cons = "Select RVTCantidad From RenglonVtaTelefonica Where RVtVentaTelefonica = " & prmDocumento _
            & " And RVTArticulo = " & lArtCobro
    End If
    Set rsQ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsQ.EOF Then
        GetArtCobroDocumento = rsQ(0)
    End If
    rsQ.Close
    Exit Function
errGACD:
    clsGeneral.OcurrioError "Error al buscar la cantidad de artículos cobrados en el documento.", Err.Description
End Function

Private Function GetQArtCobroNecesito(ByVal lArtCobro As Long) As Integer
On Error GoTo errGACN
Dim rsQ As rdoResultset

    GetQArtCobroNecesito = 0
    If lArtCobro = 0 Then Exit Function
    If prmTipoDoc = 1 Then
        Cons = "Select Sum(RenCantidad) From Renglon, Articulo Where RenDocumento = " & prmDocumento _
            & " And ArtInstalacion = " & lArtCobro & " And ArtID = RenArticulo"
    Else
        Cons = "Select Sum(RVTCantidad) From RenglonVtaTelefonica, Articulo Where RVtVentaTelefonica = " & prmDocumento _
            & " And ArtInstalacion = " & lArtCobro & " And RVTArticulo = ArtID"
    End If
    Set rsQ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsQ.EOF Then GetQArtCobroNecesito = rsQ(0)
    rsQ.Close
    
    'Ahora dado el artículo que estoy instalando me fijo la Q
    Exit Function
errGACN:
    clsGeneral.OcurrioError "Error al buscar la cantidad de artículos de cobro que se necesitan.", Err.Description
End Function

Private Function GetQArtCobroEnInstalacion(ByVal lArtCobro As Long) As Integer
On Error GoTo errGACI
Dim rsQ As rdoResultset
Dim sValor() As String
Dim sDoc() As String
Dim iCont As Integer, iParcial As Integer
    
    GetQArtCobroEnInstalacion = 0
    If lArtCobro = 0 Then Exit Function
    
    Cons = "Select RInCobro, RInCantidad From Instalacion, RenglonInstalacion, Articulo " _
        & " Where InsTipoDocumento = " & prmTipoDoc & " And InsDocumento = " & prmDocumento _
        & " And ArtInstalacion = " & lArtCobro & " And InsAnulada Is Null And RInCobro Is Not Null" _
        & " And RInInstalacion = InsID And ArtID = RInArticulo"
    
    Set rsQ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsQ.EOF
        sValor = Split(rsQ(0), "|")
        
        'Sumo la cantidad que se pudo cobrar en la instalación.
        iParcial = sValor(0)
        
        If UBound(sValor) = 1 Then
            sDoc = Split(sValor(1), ";")
            For iCont = 0 To UBound(sDoc)
                'Tengo cantidad en otros documentos.
                iParcial = iParcial + Val(Mid(sDoc(iCont), 1, InStr(1, sDoc(iCont), ":", vbTextCompare) - 1))
            Next
        End If
        
        If iParcial <> rsQ("RInCantidad") And iParcial > 0 Then
            'El resto se cobro en el documento.
            iParcial = rsQ!RInCantidad
        End If
        GetQArtCobroEnInstalacion = GetQArtCobroEnInstalacion + iParcial
        rsQ.MoveNext
    Loop
    rsQ.Close
    
    Exit Function
errGACI:
    clsGeneral.OcurrioError "Error al buscar la cantidad de artículos cobrados en instalaciones del documento.", Err.Description
End Function

Private Sub f_SetArtCobroParaInstalacion(ByVal iRow As Integer, ByVal sDato As String)
Dim sValor() As String
Dim sDoc() As String
Dim iCont As Integer

    If sDato = "" Then Exit Sub
    
    sValor = Split(sDato, "|")
    vsArticulo.Cell(flexcpData, iRow, 6) = Val(sValor(0))
    
    If UBound(sValor) = 1 Then
        sDoc = Split(sValor(1), ";")
        vsArticulo.Cell(flexcpText, iRow, 7) = sValor(1)
        For iCont = 0 To UBound(sDoc)
            'Tengo cantidad en otros documentos.
            vsArticulo.Cell(flexcpData, iRow, 7) = vsArticulo.Cell(flexcpData, iRow, 7) + Val(Mid(sDoc(iCont), 1, InStr(1, sDoc(iCont), ":", vbTextCompare) - 1))
        Next
    End If
    
End Sub
Private Sub f_SetTotalInstalacion()
'Recorro la lista y pongo en el label el costo de instalación
Dim iCont As Integer
Dim bCobrar As Boolean
Dim cCobro As Currency
    
    lPrecio.Caption = "0.00"
    cCobro = 0
    bCobrar = True
    With vsArticulo
        For iCont = .FixedRows To .Rows - 1
            .Cell(flexcpBackColor, iCont, 0, iCont, .Cols - 1) = vbWhite
            If .Cell(flexcpValue, iCont, 0) > 0 Then
                lPrecio.Caption = CCur(lPrecio.Caption) + (.Cell(flexcpValue, iCont, 0) * .Cell(flexcpValue, iCont, 6))
                If .Cell(flexcpData, iCont, 6) > 0 Then cCobro = cCobro + (.Cell(flexcpData, iCont, 6) * .Cell(flexcpValue, iCont, 6))
                If .Cell(flexcpValue, iCont, 0) <> .Cell(flexcpData, iCont, 5) + .Cell(flexcpData, iCont, 6) + .Cell(flexcpData, iCont, 7) Then
                    .Cell(flexcpBackColor, iCont, 0, iCont, .Cols - 1) = vbYellow
                End If
            End If
            
        Next iCont
    End With
    lPrecio.Tag = CCur(lPrecio.Caption)
    lPrecio.Caption = " de " & Format(lPrecio.Tag, "#,##0.00")
    
    If bCobrar Then caCobrar.Text = cCobro
    
End Sub

Private Sub f_SetTotalCobro(ByVal iRow As Integer)
'Pongo la cantidad total que necesito cobrar de la instalación del artículo seleccionado.
Dim iOtro As Integer
    
    'Tengo la Sgte información
    'Data 4 = la Q de art. cobrados en el documento que pudo utilizar.
    'Data 5 = la Q que tengo asignada.
    'Data 7 = la Q que me asigno x otros documentos.
            
    With vsArticulo
        If .Cell(flexcpValue, iRow, 0) <= .Cell(flexcpData, iRow, 7) Then
            'Todo lo esta cobrando en otro documento.
            .Cell(flexcpData, iRow, 5) = 0
        Else
            If .Cell(flexcpData, iRow, 4) >= .Cell(flexcpValue, iRow, 0) - .Cell(flexcpData, iRow, 7) Then
                .Cell(flexcpData, iRow, 5) = .Cell(flexcpValue, iRow, 0) - .Cell(flexcpData, iRow, 7)
            Else
                'Me está quedando un saldo a cobrar en la instalación
                .Cell(flexcpData, iRow, 5) = .Cell(flexcpData, iRow, 4)
            End If
        End If
                
        If .Cell(flexcpData, iRow, 4) - .Cell(flexcpData, iRow, 5) > 0 Then
            'No esta usando todos los posibles que se cobraron el documento, busco si tengo otra fila que lo pueda usar.
            iOtro = f_OtroNecesitaCobro(iRow)
            If iOtro > -1 Then
                .Cell(flexcpData, iOtro, 6) = .Cell(flexcpData, iOtro, 6) - (.Cell(flexcpData, iRow, 4) - .Cell(flexcpData, iRow, 5))
                .Cell(flexcpData, iOtro, 5) = .Cell(flexcpData, iOtro, 5) + (.Cell(flexcpData, iRow, 4) - .Cell(flexcpData, iRow, 5))
                .Cell(flexcpData, iOtro, 4) = .Cell(flexcpData, iOtro, 4) + (.Cell(flexcpData, iRow, 4) - .Cell(flexcpData, iRow, 5))
                .Cell(flexcpData, iRow, 4) = .Cell(flexcpData, iRow, 5)
            End If
            'SI NO HAY LO UTILIZA EL MISMO QUE QUITO.
        End If
        
        .Cell(flexcpData, iRow, 6) = .Cell(flexcpValue, iRow, 0) - .Cell(flexcpData, iRow, 7) - .Cell(flexcpData, iRow, 5)
        'Ocurre si la cantidad dada para algún documento supera el total de la instalación
        If .Cell(flexcpData, iRow, 6) < 0 Then .Cell(flexcpData, iRow, 6) = 0
    End With

End Sub

Private Function db_GetQArtCobroUsadoEnInstalacion(ByVal lArtCobro As Long) As Integer
On Error GoTo errGACI
Dim rsQ As rdoResultset
Dim sValor() As String
Dim sDoc() As String
Dim iCont As Integer, iParcial As Integer
    
    db_GetQArtCobroUsadoEnInstalacion = 0
    If lArtCobro = 0 Then Exit Function
    
    Cons = "Select RInCobro, RInCantidad From Instalacion, RenglonInstalacion, Articulo " _
        & " Where InsTipoDocumento = " & prmTipoDoc & " And InsDocumento = " & prmDocumento _
        & " And ArtInstalacion = " & lArtCobro & " And InsAnulada Is Null And RInCobro Is Not Null" _
        & " And RInInstalacion = InsID And ArtID = RInArticulo"
    
    Set rsQ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsQ.EOF
        sValor = Split(rsQ(0), "|")
        
        'Sumo la cantidad que se pudo cobrar en la instalación.
        iParcial = sValor(0)
        
        If UBound(sValor) = 1 Then
            sDoc = Split(sValor(1), ";")
            For iCont = 0 To UBound(sDoc)
                'Tengo cantidad en otros documentos.
                iParcial = iParcial + Val(Mid(sDoc(iCont), 1, InStr(1, sDoc(iCont), ":", vbTextCompare) - 1))
            Next
        End If
        'La diferencia esta cobrada en el documento.
        iParcial = rsQ!RInCantidad - iParcial
        db_GetQArtCobroUsadoEnInstalacion = db_GetQArtCobroUsadoEnInstalacion + iParcial
        rsQ.MoveNext
    Loop
    rsQ.Close
    Exit Function
errGACI:
    clsGeneral.OcurrioError "Error al buscar la cantidad de artículos cobrados en instalaciones del documento.", Err.Description
End Function

Private Function f_OtroNecesitaCobro(ByVal iRow As Integer) As Integer
    'Retorno la fila
Dim iCont As Integer
    
    f_OtroNecesitaCobro = -1
    For iCont = 0 To vsArticulo.Rows - 1
        If iCont <> iRow Then
            If vsArticulo.Cell(flexcpData, iCont, 3) = vsArticulo.Cell(flexcpData, iRow, 3) And vsArticulo.Cell(flexcpData, iCont, 6) > 0 Then
                f_OtroNecesitaCobro = iCont
                Exit For
            End If
        End If
    Next iCont
    
End Function

Private Function fnc_ControlQMaxima(ByRef bSuceso As Boolean) As Boolean
'Retorno True para que siga grabando.
Dim rsQ As rdoResultset
Dim sQTop As String
Dim vQ() As String

    On Error GoTo errCQM
    bSuceso = False
    fnc_ControlQMaxima = True
    Cons = "Select IsNull(InsQMaxima, '0:0') from Instaladores Where InsCodigo = " & cInstalador.ItemData(cInstalador.ListIndex)
    Set rsQ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsQ.EOF Then
        sQTop = Trim(rsQ(0))
    Else
        rsQ.Close
        Exit Function
    End If
    rsQ.Close
    
    'Si es 0:0 me voy
    If sQTop = "0:0" Then Exit Function
    
    If InStr(1, sQTop, ":", vbTextCompare) = 0 Then sQTop = sQTop & ":0"
    
    vQ = Split(sQTop, ":")
    
    Dim bClose As Boolean
    bClose = False
    'Ahora busco el tipo de flete para el instalador.
    Cons = "Select TFlFechaAgeHab From TipoFlete Where TFlCodigo = (Select InsTipoFlete From Instaladores Where InsCodigo = " & cInstalador.ItemData(cInstalador.ListIndex) & ")"
    Set rsQ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsQ.EOF Then
        If Not IsNull(rsQ!TFlFechaAgeHab) Then
            If rsQ!TFlFechaAgeHab <= dpFecha.Value Then bClose = True
        End If
    End If
    rsQ.Close
    

    Cons = "Select Count(Distinct InsID) as QInst, IsNull(Sum(RInCantidad), 0) as QEquip " & _
                "From Instalacion " & _
                        " Left Outer Join Documento On DocCodigo = InsDocumento And InsTipoDocumento = 1 And DocAnulado = 0" & _
                " , RenglonInstalacion, Articulo " & _
                " Where InsFechaProm = '" & Format(dpFecha.Value, "yyyy/mm/dd") & "'" & _
                " And InsInstalador = " & cInstalador.ItemData(cInstalador.ListIndex) & _
                " And InsAnulada Is Null And InsLiquidacion Is Null And InsFechaRealizada Is Null " & _
                " And InsID = RInInstalacion And RInArticulo = ArtID"
    Set rsQ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsQ.EOF Then
       
        If Val(vQ(0)) <= rsQ(0) Or Val(vQ(1)) <= rsQ(1) Then
            
            'Cierro la agenda.
            If bClose Then loc_CierroAgenda
            
            If MsgBox("'" & cInstalador.Text & "' para la fecha seleccionada supera el máximo de instalaciones posibles." & vbCr & vbCr & _
                "Instalaciones asignadas: " & rsQ(0) & " (de " & vQ(0) & " posibles)" & vbCr & _
                "Artículos asignados: " & rsQ(1) & " (de " & vQ(1) & " posibles)" & _
                "Si contínua hay riesgo de no poder cumplir con la instalación." & vbCr & vbCr & "¿Desea de todas formas almacenar la instalación para la fecha seleccionada?", vbQuestion + vbYesNo + vbDefaultButton2, "Posible error") = vbNo Then
                
                fnc_ControlQMaxima = False
            Else
                bSuceso = True
            End If
        End If

        rsQ.Close
    End If
    Exit Function
errCQM:
    
    clsGeneral.OcurrioError "Error al calcular el máximo de instalaciones posibles.", Err.Description
End Function

Private Sub loc_CierroAgenda()
On Error GoTo errSF
Dim RsF As rdoResultset
    
Dim sMat As String
Dim iQ As Integer
Dim dCierre As Date
Dim douAgenda As Double, douHabilitado As Double
Dim vIndex() As String
Dim sIndex As String, sIndex2 As String
    
    If cInstalador.ListIndex = -1 Then Exit Sub
    ReDim vIndex(0)
    
    iQ = 0
    'Busco en la BD los indices que corresponden a la fecha seleccionada.
    Cons = "Select HFlIndice From HorarioFlete Where HFLDiaSemana = " & Weekday(dpFecha.Value)
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsF.EOF Then
        RsF.Close
        Exit Sub
    End If
    
    Do While Not RsF.EOF
        sIndex = sIndex & RsF("HFlIndice") & ","
        RsF.MoveNext
    Loop
    RsF.Close
    sIndex = "," & sIndex
    
    'Ahora busco el tipo de flete para el instalador.
    Cons = "Select * From TipoFlete Where TFlCodigo = (Select InsTipoFlete From Instaladores Where InsCodigo = " & cInstalador.ItemData(cInstalador.ListIndex) & ")"
    Set RsF = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If Not RsF.EOF Then
        If IsNull(RsF!TFlFechaAgeHab) Then dCierre = Date Else dCierre = RsF!TFlFechaAgeHab
        douAgenda = RsF!TFlAgenda
        If IsNull(RsF!TFlAgendaHabilitada) Then douHabilitado = -1 Else douHabilitado = RsF!TFlAgendaHabilitada
        
        If DateDiff("d", dCierre, Date) >= 7 Then
            'Como cerro hace + de una semana tomo la agenda normal.
            dCierre = Date
            sMat = superp_MatrizSuperposicion(douAgenda)
        Else
            If douHabilitado > 0 Then
                sMat = superp_MatrizSuperposicion(douHabilitado)
            Else
                sMat = superp_MatrizSuperposicion(douAgenda)
            End If
        End If
        
        douAgenda = 0
        'Con la Matriz lo que hago es recorrerla y quitarle los indices que consulte arriba.
        vIndex = Split(sMat, ",")
        
        'Recorro el array y si el indice no esta arriba --> lo agrego.
        For iQ = 0 To UBound(vIndex)
            If InStr(1, sIndex, "," & vIndex(iQ) & ",", vbTextCompare) = 0 Then
                douAgenda = douAgenda + superp_ValSuperposicion(CInt(vIndex(iQ)))
            End If
        Next
        
        RsF.Edit
        RsF!TFlFechaAgeHab = Format(dCierre, "yyyy/mm/dd")
        RsF!TFlAgendaHabilitada = douAgenda
        RsF.Update
        
    End If
    RsF.Close
    Exit Sub
errSF:
    clsGeneral.OcurrioError "Error al intentar cerrar la agenda del instalador para la fecha indicada.", Err.Description
End Sub

Private Sub loc_ShowTotalInstalaciones()
Dim rsQ As rdoResultset
Dim sQTop As String
Dim vQ() As String
On Error GoTo errSTI

    If cInstalador.ListIndex = -1 Then
        MsgBox "No hay instalador seleccionado.", vbInformation, "Atención"
        Exit Sub
    End If
    
    Cons = "Select IsNull(InsQMaxima, '0:0') from Instaladores Where InsCodigo = " & cInstalador.ItemData(cInstalador.ListIndex)
    Set rsQ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsQ.EOF Then
        sQTop = Trim(rsQ(0))
    Else
        rsQ.Close
        Exit Sub
    End If
    rsQ.Close
    
    'Si es 0:0 me voy
    If InStr(1, sQTop, ":", vbTextCompare) = 0 Then sQTop = sQTop & ":0"
    
    vQ = Split(sQTop, ":")
    
    Cons = "Select Count(Distinct InsID) as QInst, IsNull(Sum(RInCantidad), 0) as QEquip " & _
                "From Instalacion " & _
                        " Left Outer Join Documento On DocCodigo = InsDocumento And InsTipoDocumento = 1 And DocAnulado = 0" & _
                " , RenglonInstalacion, Articulo " & _
                " Where InsFechaProm = '" & Format(dpFecha.Value, "yyyymmdd") & "'" & _
                " And InsInstalador = " & cInstalador.ItemData(cInstalador.ListIndex) & _
                " And InsAnulada Is Null And InsLiquidacion Is Null And InsFechaRealizada Is Null " & _
                " And InsID = RInInstalacion And RInArticulo = ArtID"
    Set rsQ = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsQ.EOF Then
        'Cons = "Instalaciones asignadas: " & rsQ(0) & " de " & vQ(0) & " posibles"
        MsgBox "Para el día " & dpFecha.Value & " hay asignados " & rsQ(1) & " equipos de " & vQ(1) & " posibles.", IIf(rsQ(1) > vQ(1), vbExclamation, vbInformation), "Totales Asignados para la fecha"
    End If
    rsQ.Close
Exit Sub
errSTI:
    MsgBox ""
End Sub

Private Sub db_FindDireccionUltimosEnvios()
Dim rsDE As rdoResultset
On Error GoTo errDE
    
    Cons = "Select Top 15 EnvDireccion, CalCodigo, DirPuerta, EnvFechaPrometida as Fecha, rTrim(CalNombre) + ' ' + rTrim(DirPuerta) as Calle" & _
            " From Envio, Direccion, Calle " & _
            " Where EnvCliente = " & lIDCli & " And EnvDireccion = DirCodigo And DirCalle = CalCodigo Order By EnvCodigo Desc "
    
    Dim objLista As New clsListadeAyuda
    With objLista
        If .ActivarAyuda(cBase, Cons, 4500, 3, "Dirección últimos envíos") > 0 Then
            Me.Refresh
            cDireccion.ListIndex = -1
            BuscoCodigoEnCombo cDireccion, .RetornoDatoSeleccionado(0)
            If cDireccion.ListIndex = -1 Then
                cDireccion.AddItem .RetornoDatoSeleccionado(4)
                cDireccion.ItemData(cDireccion.NewIndex) = .RetornoDatoSeleccionado(0)
                cDireccion.ListIndex = cDireccion.NewIndex
            End If
        Else
            Me.Refresh
        End If
    End With
    Set objLista = Nothing
    
Exit Sub
errDE:
    clsGeneral.OcurrioError "Error al cargar las direcciones de los últimos envíos.", Err.Description
End Sub

