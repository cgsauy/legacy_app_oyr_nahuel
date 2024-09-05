VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Begin VB.Form frmRemito 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remitos de Mercadería"
   ClientHeight    =   5310
   ClientLeft      =   1995
   ClientTop       =   2130
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRemito.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8640
   Begin VSPrinter8LibCtl.VSPrinter vspPrinter 
      Height          =   3135
      Left            =   480
      TabIndex        =   30
      Top             =   1560
      Visible         =   0   'False
      Width           =   4695
      _cx             =   8281
      _cy             =   5530
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
      Zoom            =   14.6780303030303
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   8640
      _ExtentX        =   15240
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
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin AACombo99.AACombo cSucursal 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      BackColor       =   16777152
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
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   5055
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15216
            MinWidth        =   2
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox cIguales 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Hacer remitos &iguales (1 unidad por remito)."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4680
      Width           =   3975
   End
   Begin VB.TextBox tAmpliacion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      MaxLength       =   180
      TabIndex        =   9
      Top             =   2310
      Width           =   8175
   End
   Begin VB.TextBox tRemito 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   7560
      MaxLength       =   8
      TabIndex        =   7
      Text            =   "180023"
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox cDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox tNumero 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   2770
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "4821"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   5760
      MaxLength       =   3
      TabIndex        =   14
      Top             =   4680
      Width           =   615
   End
   Begin MSComctlLib.ListView lvVenta 
      Height          =   1815
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Artículo"
         Object.Width           =   7232
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Cantidad"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "A Retirar"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "En Remito"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Diferencia"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.ComboBox cMoneda 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6840
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   16
      Text            =   "cMoneda"
      Top             =   960
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
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
            Picture         =   "frmRemito.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemito.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemito.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemito.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemito.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemito.frx":0BA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSReport8LibCtl.VSReport vsrReport 
      Left            =   7920
      Top             =   4800
      _rv             =   800
      ReportName      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   0
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         Columns         =   1
         ColumnLayout    =   0
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   ""
         RecordSource    =   ""
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   0
      SectionCount    =   5
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   0
   End
   Begin VB.Label lPNC 
      BackStyle       =   0  'Transparent
      Caption         =   "Impresora:"
      Height          =   255
      Left            =   4440
      TabIndex        =   29
      Top             =   420
      Width           =   2415
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   960
      Left            =   120
      Picture         =   "frmRemito.frx":0EBE
      Top             =   405
      Width           =   3795
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Sucursal"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Lista"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Facturista:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   27
      Top             =   4740
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Núme&ro"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7560
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lCliente 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CARLOS EDUARDO, GUTIERREZ RAMELA"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4920
      TabIndex        =   25
      Top             =   1560
      UseMnemonic     =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      TabIndex        =   26
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Número"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Emisión"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   24
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lEmision 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20/20/00 20:00"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3720
      TabIndex        =   23
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Documento"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   8640
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "&Dígito Usuario:"
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   4740
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   15
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Ampliación:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2085
      Width           =   855
   End
   Begin VB.Label lOperacion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fac. Contado"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5520
      TabIndex        =   22
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Documento"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   21
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label labFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10-Dic-1998"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4440
      TabIndex        =   20
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "R.U.C. 21.025996.0012"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   18
      Top             =   120
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   630
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   8415
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   0
      X2              =   8640
      Y1              =   5040
      Y2              =   5040
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
      Begin VB.Menu MnuOpLinea1 
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
      Begin VB.Menu MnuReimprimirRemito 
         Caption         =   "Reimprimir remito"
      End
   End
   Begin VB.Menu MnuSalir 
      Caption         =   "&Salir"
      Begin VB.Menu MnuVolver 
         Caption         =   "&Del formulario"
         Shortcut        =   ^X
      End
      Begin VB.Menu MnuSalirL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrintConfig 
         Caption         =   "Configurar Impresora"
      End
   End
End
Attribute VB_Name = "frmRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prm_Documento As Long
Public prm_Remito As Long

Dim prmIDsFletes As String
Dim gSucesoUsr As Long, gSucesoDef As String

Dim aMsgError As String
Dim sNuevo As Boolean
Dim sModificar As Boolean

Dim prmFMRem As Date, prmFMDoc As Date
Dim aCodigo As Long

'RDO.------------------------------------------
Dim RsDoc As rdoResultset           'Documento
Dim rsRem As rdoResultset          'Remito
Dim rsRen As rdoResultset           'Renglon Remito

Dim jobnum As Integer

Dim aTexto As String
Dim itmX As ListItem

Private Sub cDocumento_GotFocus()
    Status.Panels(1).Text = "Seleccione el tipo de documento a buscar."
End Sub

Private Sub cDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tNumero
End Sub

Private Sub cIguales_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub

Private Sub cSucursal_GotFocus()
    cSucursal.SelStart = 0
    cSucursal.SelLength = Len(cSucursal.Text)
    Status.Panels(1).Text = " Seleccione la sucursal de emisión del documento."
End Sub

Private Sub cSucursal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cDocumento.SetFocus
End Sub

Private Sub cSucursal_LostFocus()
    cSucursal.SelLength = 0
End Sub

Private Sub Form_Activate()
    
    On Error Resume Next
    Screen.MousePointer = vbDefault
    DoEvents
    rsRem.Requery
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrLoad
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Height = 6000
    
    InicializoCrystalEngine
    
    FechaDelServidor

    sNuevo = False: sModificar = False
    Botones False, False, False, False, False, Toolbar1, Me
    
    labFecha.Caption = Format(gFechaServidor, "d-Mmm-yyyy")
    lOperacion.Caption = "REMITO"
    
    lPNC.Caption = "Imprimir en: " & paIRemitoN
    If Not paPrintEsXDefecto Then lPNC.ForeColor = &HC0&
    
    DeshabilitoIngreso
    LimpioFicha ""
    
    prmIDsFletes = CargoArticulosDeFlete
    
    z_CargoCombos
    
    'Inicializo el resultset de Remitos en Cero
    Cons = "Select * from Remito Where RemCodigo = 0"
    Set rsRem = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If prm_Remito > 0 Then
        tRemito.Text = prm_Remito
        CargoDatosRemito
    Else
        If prm_Documento > 0 Then CargoPorDocumento
    End If
    
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Err.Description
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Status.Panels(1).Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    crCierroTrabajo jobnum
    crCierroEngine
    
    rsRem.Close
    cBase.Close
    eBase.Close
    
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End

End Sub

Private Sub Label12_Click()
    Foco cMoneda
End Sub

Private Sub Label13_Click()
    Foco tUsuario
End Sub

Private Sub Label17_Click()
    Foco tNumero
End Sub

Private Sub Label3_Click()
    Foco tAmpliacion
End Sub

Private Sub Label5_Click()
    Foco tRemito
End Sub

Private Sub Label6_Click()
    cDocumento.SetFocus
End Sub

Private Sub lUsuario_Click()
On Error GoTo errU
Dim sinput As String
sinput = InputBox("¿Qué?")
If sinput = "printer" Then
    AccionImprimir (Val(tRemito.Text))
End If
Exit Sub
errU:
MsgBox Err.Description
End Sub

Private Sub lvVenta_GotFocus()

    If sNuevo Then
        Status.Panels(1).Text = "Lista de artículos ingresados. Para asignar/eliminar:  [+/-] por unidad - [Espacio/Supr] por renglón."
    ElseIf sModificar Then
        Status.Panels(1).Text = "Lista de artículos ingresados Para asignar/eliminar:  [+/-] por unidad."
    Else
        Status.Panels(1).Text = "Lista de artículos ingresados en el documento."
    End If
    
End Sub

Private Sub lvVenta_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrlvKD

    If Not sNuevo And Not sModificar Then Exit Sub
    
    If lvVenta.ListItems.Count > 0 Then
        '2=Cantidad, 3=AEntregar, 4=EnRemito, 5=Diferencia
        Select Case KeyCode
            Case vbKeySpace
                If sNuevo Then
                    If Val(lvVenta.SelectedItem.SubItems(3)) <> 0 Then
                        lvVenta.SelectedItem.SubItems(4) = lvVenta.SelectedItem.SubItems(3)
                        lvVenta.SelectedItem.SubItems(5) = "-" & lvVenta.SelectedItem.SubItems(3)
                    End If
                End If
                                
            Case vbKeyDelete
                If sNuevo Then
                    lvVenta.SelectedItem.SubItems(4) = "0"
                    lvVenta.SelectedItem.SubItems(5) = "0"
                End If
            
            Case vbKeySubtract
                If Val(lvVenta.SelectedItem.SubItems(4)) <> 0 Then
                    lvVenta.SelectedItem.SubItems(4) = Val(lvVenta.SelectedItem.SubItems(4)) - 1
                    lvVenta.SelectedItem.SubItems(5) = Val(lvVenta.SelectedItem.SubItems(5)) + 1
                    If sModificar Then lvVenta.SelectedItem.SubItems(3) = Val(lvVenta.SelectedItem.SubItems(3)) + 1
                End If
                
            Case vbKeyAdd
                If sNuevo Then
                    If Val(lvVenta.SelectedItem.SubItems(3)) <> Val(lvVenta.SelectedItem.SubItems(4)) Then
                        lvVenta.SelectedItem.SubItems(4) = Val(lvVenta.SelectedItem.SubItems(4)) + 1
                        lvVenta.SelectedItem.SubItems(5) = Val(lvVenta.SelectedItem.SubItems(5)) - 1
                    End If
                Else
                    If Val(lvVenta.SelectedItem.SubItems(3)) <> 0 Then
                        lvVenta.SelectedItem.SubItems(4) = Val(lvVenta.SelectedItem.SubItems(4)) + 1
                        lvVenta.SelectedItem.SubItems(5) = Val(lvVenta.SelectedItem.SubItems(5)) - 1
                        lvVenta.SelectedItem.SubItems(3) = Val(lvVenta.SelectedItem.SubItems(3)) - 1
                    End If
                End If
        End Select
    End If
    Exit Sub

ErrlvKD:
    clsGeneral.OcurrioError "Ocurrio un error inesperado."
End Sub

Private Sub lvVenta_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If cIguales.Enabled Then
            cIguales.SetFocus
        ElseIf tUsuario.Enabled Then
            tUsuario.SetFocus
        Else
            tNumero.SetFocus
        End If
    End If
    
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

Private Sub MnuReimprimirRemito_Click()
    If IsNumeric(tRemito.Text) Then AccionImprimir Val(tRemito.Text)
End Sub

Private Sub MnuVolver_Click()
    Unload Me
End Sub

Private Sub tAmpliacion_GotFocus()

    Status.Panels(1).Text = "Ingrese un texto ampliación para el remito."
    
End Sub

Private Sub tAmpliacion_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then lvVenta.SetFocus
    
End Sub

Private Sub tNumero_GotFocus()
    
    tNumero.SelStart = 0
    tNumero.SelLength = Len(tNumero.Text)
    Status.Panels(1).Text = "Ingrese el número de documento a buscar."
    
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Trim(tNumero.Text) = "" Or Trim(cDocumento.Text) = "" Then Exit Sub
        
        If cDocumento.ListIndex = -1 Or cSucursal.ListIndex = -1 Then
            MsgBox "Los datos ingresados para la búsqueda no son correctos.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        
        CargoDatosDocumento False

    End If
    
End Sub
Private Sub HabilitoIngreso()

    tRemito.Enabled = False
    tNumero.Enabled = False
    cDocumento.Enabled = False
    cSucursal.Enabled = False
    
    tRemito.BackColor = Inactivo
    tNumero.BackColor = Inactivo
    cDocumento.BackColor = Inactivo
    cSucursal.BackColor = Inactivo
    
    lEmision.BackColor = Inactivo
    cMoneda.BackColor = Inactivo
    lCliente.BackColor = Inactivo
    
    tAmpliacion.Enabled = True
    tAmpliacion.BackColor = Blanco
    lvVenta.BackColor = Blanco
    
    If sNuevo Then
        tRemito.Text = ""
        For Each itmX In lvVenta.ListItems
            If Val(itmX.SubItems(3)) > 1 Then
                cIguales.Enabled = True
                Exit For
            End If
        Next
    End If
    
    tUsuario.Enabled = True
    tUsuario.BackColor = Obligatorio
    DoEvents
    
End Sub

Private Sub DeshabilitoIngreso()
Dim clBusqueda As Long
    
    clBusqueda = &HFFFFC0
    
    tRemito.Enabled = True
    tNumero.Enabled = True
    cDocumento.Enabled = True
    cSucursal.Enabled = True
    
    lEmision.BackColor = Blanco
    cMoneda.BackColor = Blanco
    lCliente.BackColor = Blanco
    
    tRemito.BackColor = clBusqueda

    tNumero.BackColor = clBusqueda
    cDocumento.BackColor = clBusqueda
    cSucursal.BackColor = clBusqueda
    
    tAmpliacion.Enabled = False
    tAmpliacion.BackColor = Inactivo
    lvVenta.BackColor = Inactivo
    
    tUsuario.Enabled = False
    tUsuario.BackColor = Inactivo
    
    cIguales.Enabled = False
    DoEvents
    
End Sub

Private Sub AccionNuevo()

    If Not HayArticulos Then
        MsgBox "No hay artículos para los cuales hacer un remito.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    sNuevo = True
    HabilitoIngreso
    CargoDatosDocumento True
    
End Sub

Private Sub AccionModificar()

    If Not fnc_ValidoInstalacion(rsRem!RemCodigo) Then
        MsgBox "El remito está asociado a una Instalación." & vbCrLf & _
                    "No se puede modificar o eliminar el remito cuando tiene asociada una instalación.", vbInformation, "Remito Asociado a Instalación"
        Exit Sub
    End If

    sModificar = True
    HabilitoIngreso
    CargoDatosRemito

End Sub
Private Sub AccionCancelar()

    DeshabilitoIngreso
    
    If sNuevo Then
        tNumero.SetFocus
        sNuevo = False
        CargoDatosDocumento True
    End If
    
    If sModificar Then
        tRemito.SetFocus
        sModificar = False
        CargoDatosRemito
    End If
    
End Sub

Private Sub AccionGrabar()

Dim RemitosIguales As String

    aMsgError = ""
    If Not ValidoCampos Then Exit Sub

    If cIguales.Value = vbChecked Then          'Hacer Remitos IGUALES      -----------------------------------------------------------------
        RemitosIguales = GrabarRemitosIguales
        Do While RemitosIguales <> ""                   'Imprimo los Remitos
            AccionImprimir CLng(Mid(RemitosIguales, 1, InStr(RemitosIguales, ":") - 1))
            RemitosIguales = Trim(Mid(RemitosIguales, InStr(RemitosIguales, ":") + 1, Len(RemitosIguales)))
        Loop
        Exit Sub
    End If
    '--------------------------------------------------------------------------------------------------------------------------------------------------- !!
    
    If MsgBox("Confirma almacenar la información ingresada?" & vbCrLf & vbCrLf & "IMPRESORA: " & paIRemitoN, vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    
    On Error GoTo ErrBT
    FechaDelServidor
    Screen.MousePointer = 11
    
    If sNuevo Then      'NUEVO ARTICULO
        
        cBase.BeginTrans    'COMIENZO LA TRANSACCION-----------------------------------------------------------!!!!!!!!!!!!!
        On Error GoTo ErrET
        
        'Valido la modificacion del Documento----------------------------------------------------------------------
        Cons = "Select * from Documento Where DocCodigo = " & tNumero.Tag
        Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsDoc!DocFModificacion <> prmFMDoc Then
            aMsgError = "El documento ha sido modificado por otro usuario. La operación se cancelará."
            RsDoc.Close
            GoTo ErrET
        End If
        RsDoc.Close
        '------------------------------------------------------------------------------------------------------------------
        rsRem.Requery
        
        CargoCamposBDRemito
                
        cBase.CommitTrans   'FIN DE LA TRANSACCION-------------------------------------------------------------!!!!!!!!!!!!
        sNuevo = False
        
        Cons = "Select * from Remito Where RemCodigo = " & aCodigo
        Set rsRem = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
        'Impresion de REMITO
        AccionImprimir aCodigo
        
    Else                      'MODIFICACION DE LOS DATOS
        
        cBase.BeginTrans    'COMIENZO LA TRANSACCION-------------------------------------------------------!!!!!!!!!!!!
        On Error GoTo ErrET
        
        rsRem.Requery
        
        'Valido la modificacion del Documento----------------------------------------------------------------------
        Cons = "Select * from Documento where DocCodigo = " & tNumero.Tag
        Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If RsDoc!DocFModificacion <> prmFMDoc Then
            aMsgError = "El documento ha sido modificado por otro usuario. La operación se cancelará."
            RsDoc.Close
            GoTo ErrET
        End If
        RsDoc.Close
        '------------------------------------------------------------------------------------------------------------------
        'Valido Modificacion del Remito-------------------------------------------------------------------------
        If prmFMRem <> rsRem!RemModificado Then
            aMsgError = "La ficha ha sido modificada por otro usuario. Verifique los datos antes de grabar."
            GoTo ErrET
        End If
        '-----------------------------------------------------------------------------------------------
        
        aCodigo = rsRem!RemCodigo
        rsRem.Requery
        
        CargoCamposBDRemito
            
        sModificar = False
        cBase.CommitTrans   'FIN DE LA TRANSACCION-------------------------------
        rsRem.Requery
        
        'Impresion de REMITO
        Screen.MousePointer = 0
        If MsgBox("Desea imprimir una copia del remito modificado.", vbQuestion + vbYesNo, "Imprime Copia ?") = vbYes Then
            If fnc_PidoSuceso(aCodigo, Val(tNumero.Tag)) Then AccionImprimir aCodigo
        End If
        Screen.MousePointer = 11
        
    End If
    DeshabilitoIngreso
    CargoDatosDocumento True
    Screen.MousePointer = 0
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "No se pudo iniciar la transacción."
    rsRem.Requery
    Screen.MousePointer = vbDefault
    Exit Sub
ErrET:
    Resume ErrRoll
ErrRoll:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    If Trim(aMsgError) = "" Then aMsgError = "No se pudo almacenar la ficha del artículo, reintente."
    clsGeneral.OcurrioError aMsgError
    rsRem.Requery
    AccionCancelar
End Sub

Private Sub CargoCamposBDRemito()

    'Cargo campos BD: REMITO----------------------------------------------------------
    If sNuevo Then rsRem.AddNew Else rsRem.Edit

    rsRem!RemDocumento = Val(tNumero.Tag)
    If sNuevo Then rsRem!RemFecha = Format(gFechaServidor, sqlFormatoFH)
    
    rsRem!RemModificado = Format(gFechaServidor, sqlFormatoFH)
    If Trim(tAmpliacion.Text) <> "" Then rsRem!RemAmpliacion = Trim(tAmpliacion.Text) Else rsRem!RemAmpliacion = Null
    rsRem!RemUsuario = tUsuario.Tag
    
    rsRem.Update
    '-------------------------------------------------------------------------------------------
    
    If sNuevo Then            'Saco el codigo de remito------------------------------------
        Cons = "Select Max(RemCodigo) From Remito Where RemDocumento = " & Val(tNumero.Tag)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        aCodigo = RsAux(0)
        RsAux.Close
    Else
        aCodigo = rsRem!RemCodigo
    End If                          '----------------------------------------------------------------
    
    If sModificar Then
        'Borro los renglones del REMITO
        Cons = "Delete RenglonRemito Where RReRemito = " & aCodigo
        cBase.Execute Cons
    End If

    'Cargo campos BD: RENGLON-REMITO-------------------------------------------------
    Cons = "Select * from RenglonRemito Where RReRemito = " & aCodigo
    Set rsRen = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    For Each itmX In lvVenta.ListItems
        If Val(itmX.SubItems(4)) <> 0 Then
            rsRen.AddNew
            rsRen!RReRemito = aCodigo
            rsRen!RReArticulo = Right(itmX.Key, Len(itmX.Key) - 1)
            rsRen!RReCantidad = Val(itmX.SubItems(4))
            rsRen!RReAEntregar = Val(itmX.SubItems(4))
            rsRen.Update
        End If
    Next
    rsRen.Close
    '-----------------------------------------------------------------------------------------------
 
    'Actualizo Campos BD: RENGLON-DOCUMENTO----------------------------------------
    For Each itmX In lvVenta.ListItems
        If Val(itmX.SubItems(5)) <> 0 Then  'Hay Diferencia (fue modificado)
            Cons = "Select * from Renglon Where RenDocumento = " & Val(tNumero.Tag) _
                                                    & " And RenArticulo = " & Right(itmX.Key, Len(itmX.Key) - 1)
            Set rsRen = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            rsRen.Edit
            rsRen!RenARetirar = rsRen!RenARetirar + CCur(itmX.SubItems(5))
            rsRen.Update
            rsRen.Close
        End If
    Next
    '--------------------------------------------------------------------------------------------

    'Actualizo la fecha de modificacion del documento-------------------------------------------
    Cons = "Select * from Documento where DocCodigo = " & Val(tNumero.Tag)
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsDoc.Edit
    RsDoc!DocFModificacion = Format(gFechaServidor, sqlFormatoFH)
    RsDoc.Update
    RsDoc.Close
    '----------------------------------------------------------------------------------------------------
    
End Sub

Private Sub AccionEliminar()

    aMsgError = ""
    If Not fnc_ValidoInstalacion(rsRem!RemCodigo) Then
        MsgBox "El remito está asociado a una Instalación." & vbCrLf & _
                    "No se puede modificar o eliminar el remito cuando tiene asociada una instalación.", vbInformation, "Remito Asociado a Instalación"
        Exit Sub
    End If
    
    If MsgBox("Confirma eliminar el remito seleccionado ?", vbQuestion + vbYesNo, "ELIMINAR") = vbNo Then Exit Sub
    
    On Error GoTo ErrBT
    FechaDelServidor
    Screen.MousePointer = 11
    
    'Valido que no se hayan retirado articulos con el remito---------------------------------------------------
    Cons = "Select * From RenglonRemito" _
            & " Where  RReRemito = " & rsRem!RemCodigo _
            & " And RReCantidad <> RReAEntregar"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        Screen.MousePointer = 0
        MsgBox "Se han retirado artículos con el remito. No es posible eliminarlo ni modificarlo.", vbExclamation, "ATENCIÓN"
        RsAux.Close
        Exit Sub
    End If
    RsAux.Close
    '------------------------------------------------------------------------------------------------------------------
        
    'Llamo al registro del Suceso-------------------------------------------------------------
    Screen.MousePointer = 11
    Dim objSuceso As New clsSuceso
    objSuceso.ActivoFormulario CLng(tUsuario.Tag), "Anulación de Remitos", cBase
    Me.Refresh
    gSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
    gSucesoDef = objSuceso.RetornoValor(Defensa:=True)
    Set objSuceso = Nothing
    If gSucesoUsr = 0 Then Screen.MousePointer = 0: Exit Sub 'Abortó el ingreso del suceso
    '---------------------------------------------------------------------------------------------
        
    cBase.BeginTrans    'COMIENZO LA TRANSACCION------------------------------------------------------------------!!!!!!!!
    On Error GoTo ErrET
    rsRem.Requery
    
    'Valido la modificacion del Documento----------------------------------------------------------------------
    Cons = "Select * from Documento where DocCodigo = " & tNumero.Tag
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsDoc!DocFModificacion <> prmFMDoc Then
        aMsgError = "El documento ha sido modificado por otro usuario. La operación se cancelará."
        RsDoc.Close
        GoTo ErrET
    End If
    RsDoc.Close
    '------------------------------------------------------------------------------------------------------------------
    'Valido Modificaion del REMITO-------------------------------------------------------------------------------
    If prmFMRem <> rsRem!RemModificado Then
        aMsgError = "La ficha del remito sido modificada por otro usuario. Verifique los datos antes de eliminar."
        GoTo ErrET
    End If
    '------------------------------------------------------------------------------------------------------------------
    aCodigo = rsRem!RemCodigo
    
    'Actualizo Campos BD: RENGLON-DOCUMENTO----------------------------------------
    For Each itmX In lvVenta.ListItems
        If Val(itmX.SubItems(4)) <> 0 Then  'Hay Diferencia (fue modificado)
            Cons = "Select * from Renglon Where RenDocumento = " & tNumero.Tag _
                                                    & " And RenArticulo = " & Right(itmX.Key, Len(itmX.Key) - 1)
            Set rsRen = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            rsRen.Edit
            rsRen!RenARetirar = rsRen!RenARetirar + CCur(itmX.SubItems(4))
            rsRen.Update
            rsRen.Close
        End If
    Next
    '--------------------------------------------------------------------------------------------

    'Actualizo la fecha de modificacion del documento-------------------------------------------
    Cons = "Select * from Documento Where DocCodigo = " & tNumero.Tag
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsDoc.Edit
    RsDoc!DocFModificacion = Format(gFechaServidor, sqlFormatoFH)
    RsDoc.Update
    RsDoc.Close
    '----------------------------------------------------------------------------------------------------
    'Borro los renglones del REMITO Y El REMITO----------------------------------------
    Cons = "Delete RenglonRemito Where RReRemito = " & aCodigo
    cBase.Execute Cons
    
    rsRem.Delete
    
    'Inserto Suceso --------------------------------------------------------------------------
    'El documento es el Contado/Credito
    aTexto = "Remito Nº " & aCodigo & " (" & Trim(lEmision.Caption) & ")"
    clsGeneral.RegistroSuceso cBase, gFechaServidor, prmSucesoAnulacionDocs, paCodigoDeTerminal, gSucesoUsr, CLng(tNumero.Tag), Descripcion:=aTexto, Defensa:=Trim(gSucesoDef)
    
    cBase.CommitTrans   'FIN DE LA TRANSACCION---------------------------------------------------------------------!!!!!!!!!
    rsRem.Requery
    LimpioFicha ""
    Botones False, False, False, False, False, Toolbar1, Me
    Screen.MousePointer = 0
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "No se pudo iniciar la transacción."
    rsRem.Requery
    Screen.MousePointer = vbDefault
    Exit Sub
ErrET:
    Resume ErrRoll
ErrRoll:
    cBase.RollbackTrans
    If Trim(aMsgError) = "" Then aMsgError = "No se pudo eliminar el remito seleccionado. Reintente."
    clsGeneral.OcurrioError aMsgError
    rsRem.Requery
    Screen.MousePointer = vbDefault
End Sub
Private Sub LimpioFicha(Llamado As String)

    Select Case UCase(Llamado)
        Case "REMITO"
            tNumero.Text = ""
            cDocumento.ListIndex = -1
    
        Case "DOCUMENTO"
            tRemito.Text = ""
            
        Case Else
            tNumero.Text = ""
            cDocumento.ListIndex = -1
            tRemito.Text = ""
    End Select
    
    lEmision.Caption = ""
    lCliente.Caption = ""
    cMoneda.Text = ""
    
    tUsuario.Tag = 0
    tUsuario.Text = ""
    lUsuario.Caption = ""
    
    tAmpliacion.Text = ""
    cIguales.Value = vbUnchecked
    
    lvVenta.ListItems.Clear
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    
    Select Case Button.Key
        Case "nuevo": AccionNuevo
        Case "modificar": AccionModificar
        Case "eliminar":  AccionEliminar
        Case "grabar": AccionGrabar
        Case "cancelar": AccionCancelar
        Case "salir": Unload Me
    End Select

End Sub

Private Sub tRemito_GotFocus()

    tRemito.SelStart = 0
    tRemito.SelLength = Len(tRemito.Text)

    Status.Panels(1).Text = "Ingrese el número de remito a buscar."
    
End Sub

Private Sub tRemito_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Trim(tRemito.Text) <> "" And IsNumeric(tRemito.Text) Then
        CargoDatosRemito
    End If
    
End Sub

Private Sub tUsuario_GotFocus()

    tUsuario.SelStart = 0: tUsuario.SelLength = Len(tUsuario.Text)
    Status.Panels(1).Text = "Ingrese el dígito de usuario."

End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And IsNumeric(tUsuario.Text) Then
        CargoUsuario 0, Val(tUsuario.Text)
        If Val(tUsuario.Tag) <> 0 Then AccionGrabar
    End If
    
End Sub

Private Sub CargoUsuario(Usuario As Long, Digito As Long)

    If Usuario <> 0 Then
        Cons = "Select * from Usuario Where UsuCodigo = " & Usuario
    End If
    If Digito <> 0 Then
        Cons = "Select * from Usuario Where UsuDigito = " & Digito
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    If RsAux.EOF Then
        tUsuario.Tag = 0
        tUsuario.Text = ""
        lUsuario.Caption = ""
        MsgBox "No existe un usuario para el dígito ingresado.", vbExclamation, "ATENCIÓN"
    Else
        tUsuario.Tag = RsAux!UsuCodigo
        tUsuario.Text = RsAux!UsuDigito
        lUsuario.Caption = Trim(RsAux!UsuIdentificacion)
    End If
    RsAux.Close
    
End Sub

Private Sub CargoPorDocumento()
'Accedo para llenar los datos del documento.
On Error GoTo errCD

    Cons = "Select * From Documento Where DocCodigo = " & prm_Documento
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        BuscoCodigoEnCombo cSucursal, RsAux!DocSucursal
        BuscoCodigoEnCombo cDocumento, RsAux!DocTipo
        If cDocumento.ListIndex = -1 Then
            cSucursal.Text = ""
            RsAux.Close
            Exit Sub
        End If
        tNumero.Text = Trim(RsAux!DocNumero)
    Else
        RsAux.Close
        Exit Sub
    End If
    RsAux.Close
    On Error Resume Next
    CargoDatosDocumento False
    If MnuNuevo.Enabled Then AccionNuevo
    Exit Sub
errCD:
    clsGeneral.OcurrioError "Error al cargar los datos del documento.", Err.Description
End Sub

Private Function BuscoDocumento(ByVal filtroBuscar As String, ByVal tiposDocumento As String) As Long
On Error GoTo errBD
    
    Screen.MousePointer = 11
    Dim sSerie As String, iNumero As Long
    Dim iQ As Integer, iCodigo As Long
    Dim sQy As String
    
    If InStr(1, filtroBuscar, "D", vbTextCompare) > 1 Then
        iCodigo = Val(Mid(filtroBuscar, InStr(1, filtroBuscar, "D", vbTextCompare) + 1))
        sQy = " WHERE DocCodigo = " & iCodigo & " AND DocTipo IN (" & tiposDocumento & ") AND DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
    Else
        If InStr(filtroBuscar, "-") <> 0 Then
            sSerie = Mid(filtroBuscar, 1, InStr(filtroBuscar, "-") - 1)
            iNumero = Val(Mid(filtroBuscar, InStr(filtroBuscar, "-") + 1))
        Else
            filtroBuscar = Replace(filtroBuscar, " ", "")
            If IsNumeric(Mid(filtroBuscar, 2, 1)) Then
                sSerie = Mid(filtroBuscar, 1, 1)
                iNumero = Val(Mid(filtroBuscar, 2))
            Else
                sSerie = Mid(filtroBuscar, 1, 2)
                iNumero = Val(Mid(filtroBuscar, 3))
            End If
        End If
        sQy = " WHERE DocTipo IN (" & tiposDocumento & ") AND DocSerie = '" & sSerie & "' AND DocNumero = " & iNumero & " AND DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
    End If
    sQy = "SELECT DocCodigo, DocFecha as Fecha" & _
        ", rtrim(TDoNombre) Documento " & _
        ", rTrim(DocSerie) + '-' + rtrim(Convert(Varchar(6), DocNumero)) as Número" & _
        " FROM Documento (index = [iTipoSerieNumero]) INNER JOIN TipoDocumento ON DocTipo = TDoId " & sQy
    sQy = sQy & " Order by DocFecha DESC"
    
    iCodigo = 0
    Dim RsDoc As rdoResultset
    cBase.QueryTimeout = 30
    Set RsDoc = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    If Not RsDoc.EOF Then
        iCodigo = RsDoc("DocCodigo")
        iQ = 1
        RsDoc.MoveNext: If Not RsDoc.EOF Then iQ = 2
    End If
    RsDoc.Close
    
    Select Case iQ
        Case 2
            Dim miLDocs As New clsListadeAyuda
            iCodigo = miLDocs.ActivarAyuda(cBase, sQy, 6100, 1)
            Me.Refresh
            If iCodigo > 0 Then iCodigo = miLDocs.RetornoDatoSeleccionado(0)
            Set miLDocs = Nothing
    End Select
    BuscoDocumento = iCodigo
    Screen.MousePointer = 0
    Exit Function
errBD:
    MsgBox "Error al buscar el documento: " & Err.Description, vbCritical, "Buscar documento"
    Screen.MousePointer = 0
End Function



'---------------------------------------------------------------------------------------------------
'   Carga los Datos del Documento, Cliente y Los renglones de articulos
'   Funciona cuando la búsqueda se da por Serie, Numero y TipoDocumento
'---------------------------------------------------------------------------------------------------
Private Sub CargoDatosDocumento(ByVal YaTaCargado As Boolean)

    On Error GoTo errCargar
    Dim bExit As Boolean
    Dim idDoc As Long
    If YaTaCargado Then idDoc = Val(tNumero.Tag)
    
    Botones False, False, False, False, False, Toolbar1, Me
    Screen.MousePointer = 11
    LimpioFicha "documento"
    
    bExit = False
    If Not YaTaCargado Or idDoc = 0 Then
        idDoc = BuscoDocumento(tNumero.Text, cDocumento.ItemData(cDocumento.ListIndex))
        If idDoc = 0 Then
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    tNumero.Tag = 0
    
    'Saco los datos del Documento, Cliente
    Cons = "Select * from Documento, Cliente" _
            & " Where DocCodigo = " & idDoc _
            & " And DocCliente = CliCodigo"
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsDoc.EOF Then
        Screen.MousePointer = 0
        MsgBox "No existe un documento para los datos ingresados.", vbExclamation, "No Hay Datos"
        bExit = True
    End If
    
    If Not bExit Then
        'Verifico si el Documento no fue anulado (Papel)--------------------------------------
        If RsDoc!DocAnulado Then
            bExit = True: Screen.MousePointer = 0
            MsgBox "El documento ingresado figura como papel anulado.", vbExclamation, "Papel Anulado"
        End If
    End If
    
    If Not bExit Then
        prmFMDoc = RsDoc!DocFModificacion
        tNumero.Tag = RsDoc!DocCodigo
        tNumero.Text = Trim(RsDoc("DocSerie")) & "-" & RsDoc("DocNumero")
        lEmision.Caption = Format(RsDoc!DocFecha, "dd/mm/yy hh:mm")
        BuscoCodigoEnCombo cMoneda, RsDoc!DocMoneda
        
        'Cargo el Nombre del Cliente del Documento          -----------------------------------------------------------------------
        If RsDoc!CliTipo = 1 Then
            Cons = "Select * from CPersona Where CPeCliente = " & RsDoc!CliCodigo
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            lCliente.Caption = ArmoNombre(Format(RsAux!CPeApellido1, "#"), Format(RsAux!CPeApellido2, "#"), Format(RsAux!CPeNombre1, "#"), Format(RsAux!CPeNombre2, "#"))
            RsAux.Close
        Else
            Cons = "Select * from CEmpresa Where CEmCliente = " & RsDoc!CliCodigo
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            lCliente.Caption = Trim(RsAux!CEmFantasia)
            RsAux.Close
        End If
        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        'Cargo los articulos del documento
        Cons = "Select Renglon.*, ArtCodigo, CASE IsNull(AEsID ,0) WHEN 0 THEN '' ELSE  'AEsp:' + RTrim(Convert(varchar(10), AEsID)) + ' ' END + ISNULL(AEsNombre, ArtNombre) ArtNombre " _
                & "FROM Renglon INNER JOIN Articulo ON RenArticulo = ArtID " _
                & "LEFT OUTER JOIN ArticuloEspecifico ON RenDocumento = AEsDocumento And AEsTipoDocumento = 1 And RenArticulo = AEsArticulo " _
                & "WHERE RenDocumento = " & RsDoc!DocCodigo
                
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            If InStr(prmIDsFletes, RsAux!RenArticulo & ",") = 0 Then
                Set itmX = lvVenta.ListItems.Add(, "A" + Str(RsAux!RenArticulo), Format(RsAux!ArtCodigo, "(000,000)"))
                itmX.SubItems(1) = Trim(RsAux!ArtNombre)
                itmX.SubItems(2) = Trim(RsAux!RenCantidad)
                itmX.SubItems(3) = Trim(RsAux!RenARetirar)
                itmX.SubItems(4) = "0"
                itmX.SubItems(5) = "0"
            End If
            RsAux.MoveNext
        Loop
        RsAux.Close

    End If
    RsDoc.Close
    
    If sNuevo Then
        Botones False, False, False, True, True, Toolbar1, Me
    Else
        Botones True, False, False, False, False, Toolbar1, Me
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos."
End Sub

'---------------------------------------------------------------------------------------------------
'   Carga los Datos del Documento, Cliente y Los renglones de articulos y el REMITO
'   Funciona cuando la búsqueda se da por Numero de Remito
'---------------------------------------------------------------------------------------------------
Private Sub CargoDatosRemito()

Dim sSeRetiro As Boolean

    On Error GoTo errCargar
    Botones False, False, False, False, False, Toolbar1, Me
    Screen.MousePointer = 11
    sSeRetiro = False
    LimpioFicha "remito"
    
    'Saco los datos del Remito-----------------------------------------------------------------
    rsRem.Close
    Cons = "Select * from Remito Where RemCodigo = " & tRemito.Text
    Set rsRem = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If rsRem.EOF Then
        Screen.MousePointer = 0
        MsgBox "No existe un remito para la numeración ingresada.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    labFecha.Caption = Format(rsRem!RemFecha, "d-Mmm-yyyy")
    prmFMRem = rsRem!RemModificado
    If Not IsNull(rsRem!RemAmpliacion) Then tAmpliacion.Text = Trim(rsRem!RemAmpliacion)
    CargoUsuario rsRem!RemUsuario, 0
    
    'Saco los datos del Documento, Cliente
    Cons = "Select * from Documento, Cliente" _
            & " Where DocCodigo = " & rsRem!RemDocumento _
            & " And DocCliente = CliCodigo"
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsDoc.EOF Then
        Screen.MousePointer = 0
        MsgBox "No existe un documento asociado al remito.", vbExclamation, "ATENCIÓN"
    Else
        'Verifico si el Documento no fue anulado (Papel)--------------------------------------
        If RsDoc!DocAnulado Then
            Screen.MousePointer = 0
            MsgBox "El documento ingresado figura como papel anulado.", vbExclamation, "ATENCIÓN"
        Else
            If Not IsNull(RsDoc!DocSucursal) Then BuscoCodigoEnCombo cSucursal, RsDoc!DocSucursal
            prmFMDoc = RsDoc!DocFModificacion
            tNumero.Tag = RsDoc!DocCodigo
            tNumero.Text = Trim(RsDoc!DocSerie) & "-" & RsDoc!DocNumero
            BuscoCodigoEnCombo cDocumento, RsDoc!DocTipo
            lEmision.Caption = Format(RsDoc!DocFecha, "dd/mm/yy hh:mm")
            BuscoCodigoEnCombo cMoneda, RsDoc!DocMoneda
            
            'Cargo el Nombre del Cliente del Documento
            If RsDoc!CliTipo = 1 Then
                Cons = "Select * from CPersona Where CPeCliente = " & RsDoc!CliCodigo
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                lCliente.Caption = ArmoNombre(Format(RsAux!CPeApellido1, "#"), Format(RsAux!CPeApellido2, "#"), Format(RsAux!CPeNombre1, "#"), Format(RsAux!CPeNombre2, "#"))
                RsAux.Close
            Else
                Cons = "Select * from CEmpresa Where CEmCliente = " & RsDoc!CliCodigo
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                lCliente.Caption = Trim(RsAux!CEmFantasia)
                RsAux.Close
            End If
            
            'Cargo los articulos del documento
            Cons = "Select Renglon.*, ArtCodigo, CASE IsNull(AEsID ,0) WHEN 0 THEN '' ELSE  'AEsp:' + RTrim(Convert(varchar(10), AEsID)) + ' ' END + ISNULL(AEsNombre, ArtNombre) ArtNombre, RReCantidad, RReAEntregar" _
                    & " From Renglon INNER JOIN Articulo ON RenArticulo = ArtId LEFT OUTER JOIN RenglonRemito ON RenArticulo = RReArticulo " _
                    & " LEFT OUTER JOIN ArticuloEspecifico ON RenDocumento = AEsDocumento And AEsTipoDocumento = 1 And RenArticulo = AEsArticulo " _
                    & " Where RenDocumento = " & RsDoc!DocCodigo _
                    & " And RReRemito = " & rsRem!RemCodigo
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            Do While Not RsAux.EOF
                Set itmX = lvVenta.ListItems.Add(, "A" + Str(RsAux!RenArticulo), Format(RsAux!ArtCodigo, "(000,000)"))
                itmX.SubItems(1) = Trim(RsAux!ArtNombre)
                itmX.SubItems(2) = Trim(RsAux!RenCantidad)
                itmX.SubItems(3) = Trim(RsAux!RenARetirar)
                If Not IsNull(RsAux!RReCantidad) Then
                    itmX.SubItems(4) = RsAux!RReCantidad
                    If RsAux!RReCantidad <> RsAux!RReAEntregar Then sSeRetiro = True
                Else
                    itmX.SubItems(4) = "0"
                End If
                itmX.SubItems(5) = "0"
                RsAux.MoveNext
            Loop
            RsAux.Close
        End If
    End If
    RsDoc.Close
    
    If sSeRetiro Then
        Screen.MousePointer = 0
        MsgBox "Se han retirado artículos con el remito ingresado. No se podrán realizar modificaciones.", vbExclamation, "ATENCIÓN"
        Botones False, False, False, False, False, Toolbar1, Me
    Else
        If sModificar Then
            Botones False, False, False, True, True, Toolbar1, Me
        Else
            Botones False, True, True, False, False, Toolbar1, Me
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ValidoCampos()

Dim aCantidad As Long

    ValidoCampos = False
    
    aCantidad = 0
    For Each itmX In lvVenta.ListItems
        aCantidad = aCantidad + itmX.SubItems(4)
    Next
    
    If aCantidad = 0 Then
        MsgBox "No se han asignado artículos al remito.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    If tUsuario.Tag = 0 Then
        MsgBox "Debe ingresar el dígito de usuario para emitir el remito.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    ValidoCampos = True
    
End Function

Private Function HayArticulos() As Boolean

    HayArticulos = False
    For Each itmX In lvVenta.ListItems
        If Val(itmX.SubItems(3)) > 0 Then
            HayArticulos = True
            Exit Function
        End If
    Next
    
End Function

Private Function GrabarRemitosIguales() As String

Dim aRemitoInicial As Long
Dim aRemitos As String

    aMsgError = "": aRemitos = ""
        
    If MsgBox("Confirma realizar remitos por unidad para los artículos seleccionados." & vbCrLf & vbCrLf & "IMPRESORA: " & paIRemitoN, vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Function
    
    On Error GoTo ErrBT
    FechaDelServidor
    Screen.MousePointer = 11
    
    cBase.BeginTrans    'COMIENZO LA TRANSACCION-----------------------------------------------------------!!!!!!!!!!!!!
    On Error GoTo ErrET
    
    'Valido la modificacion del Documento----------------------------------------------------------------------
    Cons = "Select * from Documento Where DocCodigo = " & tNumero.Tag
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsDoc!DocFModificacion <> prmFMDoc Then
        aMsgError = "El documento ha sido modificado por otro usuario. " & vbCrLf & "La operación se cancelará."
        RsDoc.Close
        GoTo ErrET
    End If
    RsDoc.Close
    '------------------------------------------------------------------------------------------------------------------
    rsRem.Requery
    
    For Each itmX In lvVenta.ListItems
    
        For I = 1 To CLng(itmX.SubItems(4))
            'Cargo campos BD: REMITO----------------------------------------------------------
            rsRem.AddNew
            rsRem!RemDocumento = tNumero.Tag
            rsRem!RemFecha = Format(gFechaServidor, sqlFormatoFH)
            rsRem!RemModificado = Format(gFechaServidor, sqlFormatoFH)
            If Trim(tAmpliacion.Text) <> "" Then
                rsRem!RemAmpliacion = Trim(tAmpliacion.Text)
            Else
                rsRem!RemAmpliacion = Null
            End If
            rsRem!RemUsuario = tUsuario.Tag
            rsRem.Update
            '-------------------------------------------------------------------------------------------
        
            'Saco el Código del Remito--------------------------------------------------------------
            If I = 1 Then
                Cons = "Select Max(RemCodigo) From Remito"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                aRemitoInicial = RsAux(0)
                aCodigo = aRemitoInicial
                RsAux.Close
                '-------------------------------------------------------------------------------------------
            Else
                aCodigo = aRemitoInicial + I - 1
            End If
            aRemitos = aRemitos & aCodigo & ":"     'Guardo todos los codigos de remitos para imprimirlos
            
            Cons = "Insert into RenglonRemito (RReRemito, RReArticulo, RReCantidad, RReAEntregar) " _
                    & " Values(" & aCodigo & ", " & Right(itmX.Key, Len(itmX.Key) - 1) & ", 1, 1 )"
            cBase.Execute Cons
            
        Next I
    Next
    
    'Actualizo Campos BD: RENGLON-DOCUMENTO----------------------------------------
    For Each itmX In lvVenta.ListItems
        If Val(itmX.SubItems(5)) <> 0 Then  'Hay Diferencia (fue modificado)
            Cons = "Select * from Renglon Where RenDocumento = " & tNumero.Tag _
                                                    & " And RenArticulo = " & Right(itmX.Key, Len(itmX.Key) - 1)
            Set rsRen = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            rsRen.Edit
            rsRen!RenARetirar = rsRen!RenARetirar + CCur(itmX.SubItems(5))
            rsRen.Update
            rsRen.Close
        End If
    Next
    '--------------------------------------------------------------------------------------------

    'Actualizo la fecha de modificacion del documento-------------------------------------------
    Cons = "Select * from Documento where DocCodigo = " & tNumero.Tag
    Set RsDoc = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsDoc.Edit
    RsDoc!DocFModificacion = Format(gFechaServidor, sqlFormatoFH)
    RsDoc.Update
    RsDoc.Close
    '----------------------------------------------------------------------------------------------------

    cBase.CommitTrans   'FIN DE LA TRANSACCION-------------------------------------------------------------!!!!!!!!!!!!
    sNuevo = False

    rsRem.Requery

    DeshabilitoIngreso
    GrabarRemitosIguales = aRemitos         'Retorno los remitos grabados
    CargoDatosDocumento True
    Screen.MousePointer = 0
    Exit Function
    
ErrBT:
    clsGeneral.OcurrioError "No se pudo iniciar la transacción."
    rsRem.Requery
    Screen.MousePointer = vbDefault
    Exit Function
ErrET:
    Resume ErrRoll
ErrRoll:
    cBase.RollbackTrans
    Screen.MousePointer = vbDefault
    If Trim(aMsgError) = "" Then aMsgError = "No se pudo almacenar la ficha del artículo, reintente."
    clsGeneral.OcurrioError aMsgError
    rsRem.Requery
    AccionCancelar
End Function

Private Sub InicializoCrystalEngine()
    
    'Inicializa el Engine del Crystal y setea la impresora para el JOB
    On Error GoTo ErrCrystal
    If crAbroEngine = 0 Then
        MsgBox Trim(crMsgErr), vbCritical, "ERROR - Impresión"
        Exit Sub
    End If
    
    'Inicializo el Reporte y SubReportes
    jobnum = crAbroReporte(prmPathListados & "Remito.RPT")
    If jobnum = 0 Then GoTo ErrCrystal
    
    'Configuro la Impresora
    If Trim(Printer.DeviceName) <> Trim(paIRemitoN) Then SeteoImpresoraPorDefecto paIRemitoN
    If Not crSeteoImpresora(jobnum, Printer, paIRemitoB, paIRemitoP) Then GoTo ErrCrystal
    
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError Trim(crMsgErr) & vbCrLf & " No se podrán imprimir remitos."
End Sub

Sub ImprimoRemito(ByVal idRem As Long)
    
    On Error GoTo errImpresora
    Dim spaso As String
    spaso = "(1) "
    vspPrinter.Device = paIRemitoN
    spaso = "(2) "
    vspPrinter.PaperBin = paIRemitoB
    spaso = "(3) "
    On Error Resume Next
    vspPrinter.paperSize = paIRemitoP
    
errImpresora:
    'MsgBox spaso & "Error al setear la impresora: " & Err.Description & vbCrLf & "Impresora: " & paIRemitoN & vbCrLf & "Bandeja: " & paIRemitoB, vbCritical, "ATENCIÓN"
    On Error Resume Next
    
    With vsrReport
        
        .Clear                  ' clear any existing fields
        .FontName = "Tahoma"    ' set default font for all controls
        .FontSize = 8

        .Load prmPathListados & "Remito.xml", "RemitoRetiro"

        .DataSource.ConnectionString = cBase.Connect

        .DataSource.RecordSource = "SELECT RemCodigo Codigo, RemFecha Fecha, '" & UCase(prmNombreSucursal) & "' Sucursal, " & _
                    "'" & lCliente.Caption & "' Cliente, '" & Trim(cSucursal.Text) & " " & Trim(cDocumento.Text) & " " & Trim(tNumero.Text) & "' Documento " & _
                    ", '" & tUsuario.Text & "' Usuario, '" & CodigoDeBarras(TipoDocumento.Remito, idRem) & "' CodigoBarras, ArtCodigo CodigoArt, ArtNombre NombreArt, RReCantidad CantidadArt, RTRIM(RemAmpliacion) Comentario " & _
                    "FROM Remito INNER JOIN RenglonRemito ON RReRemito = RemCodigo INNER JOIN Articulo ON ArtId  = RReArticulo " & _
                    "WHERE RemCodigo = " & idRem

        .Render vspPrinter
    End With
    
    vspPrinter.PrintDoc False
    
End Sub

Private Sub AccionImprimir(Documento As Long)
On Error GoTo errAI
    
    ImprimoRemito Documento
    Exit Sub

errAI:
    clsGeneral.OcurrioError "Error al imprimir.", Err.Description
Exit Sub
On Error GoTo ErrCrystal
Dim result As Integer
Dim NombreFormula As String, CantForm As Integer

    Screen.MousePointer = 11
    crSeteoImpresora jobnum, Printer, paIRemitoB, paIRemitoP
    
    'Obtengo la cantidad de formulas que tiene el reporte.
    CantForm = crObtengoCantidadFormulasEnReporte(jobnum)
    If CantForm = -1 Then GoTo ErrCrystal
    
    For I = 0 To CantForm - 1
        NombreFormula = crObtengoNombreFormula(jobnum, CInt(I))
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "sucursal": result = crSeteoFormula(jobnum%, NombreFormula, "'SUCURSAL " & UCase(prmNombreSucursal) & "'")
            Case "documento"
                aTexto = Trim(cSucursal.Text) & " " & Trim(cDocumento.Text) & " " & Trim(tNumero.Text)
                result = crSeteoFormula(jobnum%, NombreFormula, "'" & aTexto & "'")
            
            Case "cliente": result = crSeteoFormula(jobnum%, NombreFormula, "'" & Trim(lCliente.Caption) & "'")
            
            Case "usuario": result = crSeteoFormula(jobnum%, NombreFormula, "'" & tUsuario.Text & "'")
                        
            Case "codigobarras": result = crSeteoFormula(jobnum%, NombreFormula, "'" & CodigoDeBarras(TipoDocumento.Remito, Documento) & "'")
                           
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    
    'Seteo la Query del reporte-----------------------------------------------------------------
'    Cons = "SELECT Remito.RemCodigo, Remito.RemFecha, Remito.RemAmpliacion, RenglonRemito.RReCantidad, Articulo.ArtCodigo , Articulo.ArtNombre" _
'           & " From { oj (" & paBD & ".dbo.Remito Remito " _
'                        & " INNER JOIN " & paBD & ".dbo.RenglonRemito RenglonRemito ON Remito.RemCodigo = RenglonRemito.RReRemito) " _
'                        & " INNER JOIN " & paBD & ".dbo.Articulo Articulo ON RenglonRemito.RReArticulo = Articulo.ArtId}" _
'                        & " LEFT OUTER JOIN " & paBD & ".dbo.ArticuloEspecifico ArticuloEspecifico ON RenDocumento = AEsDocumento And AEsTipoDocumento = 1 And RenArticulo = AEsArticulo " _
'           & " Where RemCodigo = " & Documento
    
    Cons = "SELECT Remito.RemCodigo, Remito.RemFecha, Remito.RemAmpliacion, RenglonRemito.RReCantidad, " & _
        " ArticuloEspecifico.AEsID, ArticuloEspecifico.AEsNombre, Articulo.ArtNombre, Articulo.ArtCodigo " & _
        "FROM { oj ((CGSA.dbo.Remito Remito INNER JOIN CGSA.dbo.RenglonRemito RenglonRemito ON Remito.RemCodigo = RenglonRemito.RReRemito) " & _
        "LEFT OUTER JOIN CGSA.dbo.ArticuloEspecifico ArticuloEspecifico ON Remito.RemDocumento = ArticuloEspecifico.AEsDocumento) " & _
        "INNER JOIN CGSA.dbo.Articulo Articulo ON RenglonRemito.RReArticulo = Articulo.ArtId} " & _
        "WHERE RemCodigo = " & Documento
        
    
    If crSeteoSqlQuery(jobnum%, Cons) = 0 Then GoTo ErrCrystal
    '-------------------------------------------------------------------------------------------------------------------------------------
    
    'If crMandoAPantalla(jobnum, "Factura Contado") = 0 Then GoTo ErrCrystal
    If crMandoAImpresora(jobnum, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(jobnum, True, False) Then GoTo ErrCrystal

    'crEsperoCierreReportePantalla
    Screen.MousePointer = 0
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr
End Sub

Private Function fnc_ValidoInstalacion(idRemito As Long) As Boolean
On Error GoTo errFnc
Dim rsVal As rdoResultset

    Cons = "Select * from Instalacion " & _
                " Where InsRemito = " & idRemito & " And InsAnulada Is Null"
                
    Set rsVal = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    fnc_ValidoInstalacion = rsVal.EOF
    rsVal.Close
    
errFnc:
End Function


'-------------------------------------------------------------------------------------------------------
'   Carga un string con todos los articulos que corresponden a los fletes.
'   Se utiliza en aquellos formularios que no filtren los fletes
'-------------------------------------------------------------------------------------------------------
Private Function CargoArticulosDeFlete() As String

Dim mRetValue As String
    On Error GoTo errCargar
    mRetValue = ""
    
    'Cargo los articulos a descartar-----------------------------------------------------------
    Cons = "Select Distinct(TFlArticulo) from TipoFlete Where TFlArticulo <> Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        mRetValue = mRetValue & RsAux!TFlArticulo & ","
        RsAux.MoveNext
    Loop
    RsAux.Close
    mRetValue = mRetValue & paArticuloPisoAgencia & "," & paArticuloDiferenciaEnvio & ","
    '----------------------------------------------------------------------------------------------
    CargoArticulosDeFlete = mRetValue
    Exit Function
    
errCargar:
    CargoArticulosDeFlete = mRetValue
End Function


Private Function ArmoNombre(Ape1 As String, Ape2 As String, Nom1 As String, Nom2 As String) As String

    ArmoNombre = Trim(Ape1) & " " & Trim(Ape2)
    ArmoNombre = Trim(ArmoNombre) & ", " & Trim(Nom1) & " " & Trim(Nom2)
    
End Function

Private Function z_CargoCombos()
On Error Resume Next

    'Cargo Sucursales---------------------------------------------------------------------------
    Cons = "Select SucCodigo, SucAbreviacion from Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cSucursal, ""
    BuscoCodigoEnCombo cSucursal, paCodigoDeSucursal
    '-----------------------------------------------------------------------------------------------
    
    'Cargo Monedas ------------------------
    Cons = "Select MonCodigo, MonSigno From Moneda  Order by MonSigno"
    CargoCombo Cons, cMoneda, ""
    
    'Cargo Documentos---------------------
    cDocumento.AddItem Trim("Contado"): cDocumento.ItemData(cDocumento.NewIndex) = TipoDocumento.Contado
    cDocumento.AddItem Trim("Crédito"): cDocumento.ItemData(cDocumento.NewIndex) = TipoDocumento.Credito
    
End Function

Private Sub MnuPrintConfig_Click()
On Error Resume Next
    
    prj_LoadConfigPrint True
    
    lPNC.Caption = "Imprimir en: " & paIRemitoN
    If Not paPrintEsXDefecto Then lPNC.ForeColor = &HC0& Else lPNC.ForeColor = vbBlack
    
End Sub

Private Function fnc_PidoSuceso(idRemito As Long, idDoc As Long) As Boolean

    On Error GoTo errFnc
    fnc_PidoSuceso = False
    Screen.MousePointer = 11
    
    Dim objSuceso As New clsSuceso
    Dim gSucesoUsr As Long, gSucesoDef As String
    
    objSuceso.ActivoFormulario paCodigoDeUsuario, "Reimpresión de Documentos", cBase
    gSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
    gSucesoDef = objSuceso.RetornoValor(Defensa:=True)
    Set objSuceso = Nothing
    Me.Refresh
    If gSucesoUsr = 0 Then Screen.MousePointer = 0: Exit Function 'Abortó el ingreso del suceso
    '---------------------------------------------------------------------------------------------
                
    Dim aNDocumento As String
    aNDocumento = "Modif. Remito " & idRemito
                
    clsGeneral.RegistroSuceso cBase, gFechaServidor, prmSucesoReimpresiones, paCodigoDeTerminal, gSucesoUsr, idDoc, _
                                   Descripcion:=aNDocumento, Defensa:=Trim(gSucesoDef)
        
    fnc_PidoSuceso = True
    Exit Function
    
errFnc:
End Function
