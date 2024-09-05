VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{1292AE18-2B08-4CE3-9F79-9CB06F26AB54}#1.7#0"; "orEMails.ocx"
Begin VB.Form frmCredito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitudes Condicionales"
   ClientHeight    =   6720
   ClientLeft      =   2430
   ClientTop       =   2715
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredito.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9120
   Begin VB.CheckBox chkRetiraAqui 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Retira aquí?"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   67
      Top             =   5460
      Width           =   1815
   End
   Begin VB.CheckBox chkVencidaConyuge 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Abonar vencidas conyuge $1,560.25"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   66
      Top             =   5820
      Width           =   3375
   End
   Begin VB.CheckBox chkVencidaTit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Abonar Vencidas $1,560.25"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   65
      Top             =   5820
      Width           =   2415
   End
   Begin VB.TextBox tGarantia 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   900
      TabIndex        =   7
      Top             =   2205
      Width           =   1455
   End
   Begin orEMails.ctrEMails cEMailsT 
      Height          =   315
      Left            =   1020
      TabIndex        =   56
      Top             =   1005
      Width           =   5895
      _ExtentX        =   5980
      _ExtentY        =   556
      BackColor       =   12640511
      ForeColor       =   0
   End
   Begin AACombo99.AACombo cPago 
      Height          =   315
      Left            =   3240
      TabIndex        =   25
      Top             =   4740
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
   End
   Begin VB.CommandButton bDevolver 
      Caption         =   "&Devolver"
      Height          =   375
      Left            =   6840
      TabIndex        =   60
      Top             =   6240
      Width           =   1000
   End
   Begin VB.CommandButton bEmpleo 
      Caption         =   "EMP (F7)"
      Height          =   375
      Left            =   1080
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton bTitulo 
      Caption         =   "TIT (F8)"
      Height          =   375
      Left            =   1920
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton bReferencia 
      Caption         =   "REF (F9)"
      Height          =   375
      Left            =   2760
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton bValidar 
      Caption         =   "&Facturar"
      Height          =   375
      Left            =   5760
      TabIndex        =   59
      Top             =   6240
      Width           =   1000
   End
   Begin VB.CommandButton bCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7920
      TabIndex        =   62
      Top             =   6240
      Width           =   1000
   End
   Begin VB.CommandButton bFicha 
      Caption         =   "CLI (F6)"
      Height          =   375
      Left            =   240
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   6240
      Width           =   795
   End
   Begin VB.CheckBox oDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   240
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   57
      Top             =   660
      Width           =   195
   End
   Begin VB.ComboBox cDireccion 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   660
      Width           =   1515
   End
   Begin VB.TextBox tComentarioR 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   4560
      MaxLength       =   15
      TabIndex        =   13
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox tFRetiro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   7800
      TabIndex        =   1
      Top             =   5145
      Width           =   1215
   End
   Begin ComctlLib.ListView lvVenta 
      Height          =   1575
      Left            =   120
      TabIndex        =   21
      Top             =   3120
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   2778
      SortKey         =   2
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   17
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Financiación"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cant."
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Artículo"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "I.V.A."
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Entrega"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cuota"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Sub Total"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Envío"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Contado"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   9
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Plan"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   10
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Facturado"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   11
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Envio"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   12
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Comentario"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   13
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "cCtdo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   14
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "idInstalador"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   15
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "qAEnviar"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   16
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "idespecifico"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.TextBox tComentario 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   70
      TabIndex        =   3
      Top             =   5475
      Width           =   5895
   End
   Begin VB.TextBox tArticulo 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1200
      MaxLength       =   60
      TabIndex        =   11
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox tCantidad 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   6000
      MaxLength       =   5
      TabIndex        =   15
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox tUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   7800
      MaxLength       =   3
      TabIndex        =   5
      Top             =   5760
      Width           =   675
   End
   Begin VB.TextBox tEntrega 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   6495
      MaxLength       =   12
      TabIndex        =   17
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox tEntregaT 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1200
      MaxLength       =   12
      TabIndex        =   23
      Top             =   4755
      Width           =   1335
   End
   Begin VB.ComboBox cMoneda 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8160
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   30
      TabStop         =   0   'False
      Text            =   "cMoneda"
      Top             =   420
      Width           =   735
   End
   Begin VB.TextBox tCondicion 
      Appearance      =   0  'Flat
      Height          =   555
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   1620
      Width           =   8895
   End
   Begin ComctlLib.ListView lEnvio 
      Height          =   1575
      Left            =   120
      TabIndex        =   45
      Top             =   3120
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Financiación"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Cant."
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Artículo"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "I.V.A."
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Sub Total"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Envíos"
         Object.Width           =   0
      EndProperty
   End
   Begin AACombo99.AACombo cPendiente 
      Height          =   315
      Left            =   1200
      TabIndex        =   27
      Top             =   5145
      Width           =   5895
      _ExtentX        =   10398
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
   Begin AACombo99.AACombo cCuota 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   2760
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
   Begin VSPrinter8LibCtl.VSPrinter vspPrinter 
      Height          =   2295
      Left            =   0
      TabIndex        =   68
      Top             =   3480
      Visible         =   0   'False
      Width           =   3495
      _cx             =   6165
      _cy             =   4048
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
      Zoom            =   9.375
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
   Begin VSReport8LibCtl.VSReport vsrReport 
      Left            =   0
      Top             =   0
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
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "e&Mails:"
      Height          =   255
      Left            =   480
      TabIndex        =   55
      Top             =   1080
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   120
      X2              =   9000
      Y1              =   6180
      Y2              =   6180
   End
   Begin VB.Label lUsuarioO 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S/D"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7200
      TabIndex        =   53
      Top             =   1380
      UseMnemonic     =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Realizada por:"
      Height          =   255
      Left            =   6120
      TabIndex        =   52
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comen&tario"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lDireccionN 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      Height          =   255
      Left            =   480
      TabIndex        =   51
      Top             =   690
      Width           =   815
   End
   Begin VB.Label lDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S/D"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2040
      TabIndex        =   50
      Top             =   690
      UseMnemonic     =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "RUC:"
      Height          =   255
      Left            =   2760
      TabIndex        =   49
      Top             =   180
      Width           =   855
   End
   Begin VB.Label lRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "21 378350 0011"
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
      Height          =   255
      Left            =   3240
      TabIndex        =   48
      Top             =   180
      UseMnemonic     =   0   'False
      Width           =   1995
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor:"
      Height          =   255
      Left            =   5160
      TabIndex        =   47
      Top             =   4785
      Width           =   735
   End
   Begin VB.Label lVendedor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Pendiente por:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6000
      TabIndex        =   46
      Top             =   4755
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Retira:"
      Height          =   255
      Left            =   7200
      TabIndex        =   0
      Top             =   5220
      Width           =   615
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "&Pendiente por:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5175
      Width           =   1215
   End
   Begin VB.Label lAccesoAListas 
      BackStyle       =   0  'Transparent
      Caption         =   "&LISTA"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label53 
      BackStyle       =   0  'Transparent
      Caption         =   "Co&mentarios:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5535
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Resuelta por:"
      Height          =   255
      Left            =   3240
      TabIndex        =   44
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label lUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S/D"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4320
      TabIndex        =   43
      Top             =   1380
      UseMnemonic     =   0   'False
      Width           =   1575
   End
   Begin VB.Image iNo 
      Height          =   480
      Left            =   0
      Picture         =   "frmCredito.frx":0442
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image iSi 
      Height          =   480
      Left            =   0
      Picture         =   "frmCredito.frx":0884
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image iCondicional 
      Height          =   480
      Left            =   0
      Picture         =   "frmCredito.frx":0B8E
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ca&nt."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F&inanciación"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   135
      TabIndex        =   8
      Top             =   2520
      Width           =   1080
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6480
      TabIndex        =   41
      Top             =   4755
      Width           =   1095
   End
   Begin VB.Label lblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuario:"
      Height          =   255
      Left            =   7200
      TabIndex        =   4
      Top             =   5820
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sub Total (F)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      TabIndex        =   18
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lSubTotalF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   7695
      TabIndex        =   19
      Top             =   2760
      Width           =   1320
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Entrega"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "E&ntrega:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4785
      Width           =   735
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "&Pago:"
      Height          =   255
      Left            =   2760
      TabIndex        =   24
      Top             =   4785
      Width           =   615
   End
   Begin VB.Label lCi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " 3.709.385-6"
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
      Height          =   255
      Left            =   1320
      TabIndex        =   40
      Top             =   180
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "C.I.:"
      Height          =   255
      Left            =   480
      TabIndex        =   39
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   38
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label lFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10-Dic-1998"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5880
      TabIndex        =   37
      Top             =   420
      Width           =   1335
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Solicitud"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   36
      Top             =   180
      Width           =   975
   End
   Begin VB.Label lCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "102565"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7200
      TabIndex        =   35
      Top             =   420
      Width           =   975
   End
   Begin VB.Label lNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
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
      Height          =   255
      Left            =   1320
      TabIndex        =   34
      Top             =   390
      UseMnemonic     =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   480
      TabIndex        =   33
      Top             =   390
      Width           =   855
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8160
      TabIndex        =   32
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "&Garantía:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2220
      Width           =   735
   End
   Begin VB.Label lGarantia 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rodriguez Fernandez, Rodrigo Bernardino"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2400
      TabIndex        =   31
      Top             =   2205
      UseMnemonic     =   0   'False
      Width           =   6615
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Condiciones para la Aprobación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   1380
      Width           =   8895
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   1245
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   8895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Artículo"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label lTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6720
      TabIndex        =   42
      Top             =   4755
      Width           =   2295
   End
End
Attribute VB_Name = "frmCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oCnfgCredito As New clsImpresoraTicketsCnfg
Dim oCnfgRecibo As New clsImpresoraTicketsCnfg

Dim oCliente As clsClienteCFE
Public prmIDSolicitud As Long

Dim itmX As ListItem
Dim itmA As ListItem, itmC As ListItem

Dim sFacturada As Boolean, sConEntrega As Boolean
Dim rsSol As rdoResultset

Dim aSolicitudEstado As Integer
Dim aIDServicio As Long

Dim aTexto As String, gTextoCondicion As String, aFletes As String

Dim aPlan As Long
Dim iAuxiliar As Currency

Dim gCliente As Long ', gCategoriaCliente As Long
Dim gDirFactura As Long

Dim aSerie As String, aNumero As Long
Dim aDocumentoRecibo As Long, aDocumentoFactura As Long

Dim aMaxDocumento As Long     'Maximo numero de Documento

Dim cCofis As Currency
Dim gSucesoUsr As Long, gSucesoDef As String, gSucesoDesc As String
Dim spv_Usuario As Long, spv_Defensa As String, spv_Autoriza As Long

Dim mMoneda As Long, mRedondeo As String

'Almaceno los datos de los créditos facturados  ------------------------ 06/01/2004
Private Type typCredito
    'idFactura As Long
    idPlan As Long
    CAE As clsCAEDocumento
    Credito As clsDocumentoCGSA
End Type
Private arrCreditos() As typCredito
Private arrIdx  As Integer

Private prmArticulosSINCofis As String

'Private Function ObtenerIDConyuge() As Long
'Dim rsCony As rdoResultset
'    Cons = "Select IsNull(CPeConyuge, 0) from CPersona Where CPeCliente = " & gCliente
'    Set rsCony = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
'    If Not rsCony.EOF Then ObtenerIDConyuge = rsCony(0)
'    rsCony.Close
'End Function

Private Function ObtenerImportesVencidas(ByVal iCliente As Long, ByRef idconyuge As Long, ByRef importeconyuge As Double) As Double
On Error GoTo errOIV
Dim rsV As rdoResultset
    ObtenerImportesVencidas = 0
    idconyuge = 0
    importeconyuge = 0
    'PROCEDURE prg_PagaVencidas @IdCli int, @Emitir bit = 0
    'Set rsV = cBase.OpenResultset("Exec prg_PagaVencidas " & iCliente & ", 0", rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, rsV, "Exec prg_PagaVencidas " & iCliente & ", 0, " & paCodigoDeSucursal, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
    If Not rsV.EOF Then
        ObtenerImportesVencidas = rsV(0)
        If rsV(2) > 0 Then
            idconyuge = rsV(2)
            If Not IsNull(rsV(1)) Then importeconyuge = rsV(1)
        End If
    End If
    rsV.Close
    Exit Function
errOIV:
    clsGeneral.OcurrioError "Error al buscar el importe de cuotas vencidas.", Err.Description, "Atención"
End Function

Private Function HayCuotasVencidas() As Boolean
'Agregamos este control ya que se facturo una cuota estando el documento anulado.
    
    If chkVencidaTit.value = 1 Or chkVencidaConyuge.value = 1 Then
    
        Dim douImporte As Double, douimporteconyuge As Double
        Dim idconyuge As Long
        douImporte = ObtenerImportesVencidas(gCliente, idconyuge, douimporteconyuge)
        
        If douImporte = 0 And chkVencidaTit.value = 1 Then
        
            MsgBox "El titular no posee créditos vencidos, no se emitiran recibos.", vbExclamation, "ATENCIÓN"
            chkVencidaTit.value = 0
        
        End If
        
        If douimporteconyuge = 0 And chkVencidaConyuge.value = 1 Then
            MsgBox "El conyuge no posee créditos vencidos, no se emitiran recibos.", vbExclamation, "ATENCIÓN"
            chkVencidaConyuge.value = 0
        End If
    
    End If
    HayCuotasVencidas = True
End Function

Private Sub CargarInfoVencidas()
    
    Dim douImporte As Double, douimporteconyuge As Double
    Dim idconyuge As Long
    douImporte = ObtenerImportesVencidas(gCliente, idconyuge, douimporteconyuge)
    chkVencidaConyuge.Tag = idconyuge
    
    'Si condición es pagar vencidas o tiene importe en vencidas lo prendo.
    chkVencidaTit.value = IIf(douImporte > 0 Or Val(chkVencidaTit.Tag) <> 0, 1, 0)
    'Si condición es pagar vencidas deshabilito.
    chkVencidaTit.Enabled = (Val(chkVencidaTit.Tag) = 0)
    'Si SP retorna > 1 es el importe que debe.
    If douImporte > 1 Then chkVencidaTit.Caption = "Abonar vencidas $" & Format(douImporte, "#,##0.00")
    
    If chkVencidaTit.Enabled And chkVencidaTit.value = 1 Then chkVencidaTit.ForeColor = &H80&
    
    'Si tiene conyuge.
    chkVencidaConyuge.Enabled = chkVencidaTit.Enabled
    If Val(Val(chkVencidaConyuge.Tag)) > 0 Then
        chkVencidaConyuge.value = IIf(douimporteconyuge > 0 Or Val(chkVencidaTit.Tag) <> 0, 1, 0)
        If douimporteconyuge > 1 Then chkVencidaConyuge.Caption = "Abonar vencidas conyuge $" & Format(douimporteconyuge, "#,##0.00")
        If chkVencidaConyuge.Enabled And chkVencidaConyuge.value = 1 Then chkVencidaConyuge.ForeColor = &H80&
    End If
    
End Sub

Private Sub LimpiarCtrlsVencidas()
    With chkVencidaConyuge
        .Caption = "Abonar vencidas conyuge"
        .Enabled = True
        .value = 0
        .Tag = 0
        .ForeColor = vbBlack
    End With
    With chkVencidaTit
        .Caption = "Abonar vencidas"
        .Enabled = True
        .value = 0
        .Tag = 0
        .ForeColor = vbBlack
    End With
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bDevolver_Click()
    
    If Val(tUsuario.Tag) = 0 Then
        MsgBox "Ingrese el dígito de usuario para realizar la operación.", vbExclamation, "ATENCIÓN"
        Foco tUsuario: Exit Sub
    End If
    
    If MsgBox("Confirma volver a procesar la solicitud.", vbQuestion + vbYesNo, "DEVOLVER") = vbNo Then Exit Sub
    AccionDevolver
    
End Sub

Private Sub AccionDevolver()

    On Error GoTo errorBT
    Screen.MousePointer = 11
    FechaDelServidor    'Saco la fecha del servidor
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    Cons = "Select * From Solicitud Where SolCodigo = " & prmIDSolicitud
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    RsAux.Edit
    
    RsAux!SolFecha = Format(gFechaServidor, sqlFormatoFH)
    RsAux!SolProceso = TipoResolucionSolicitud.Manual
    RsAux!SolEstado = EstadoSolicitud.ParaRetomar 'EstadoSolicitud.Pendiente
    
    If Trim(tComentario.Text) <> "" Then RsAux!SolComentarioS = Trim(tComentario.Text)
    RsAux!SolUsuarioS = tUsuario.Tag
    RsAux!SolFResolucion = Null
    RsAux!SolDevuelta = True
    RsAux!SolVisible = Null
    RsAux.Update
    RsAux.Close
    
    Screen.MousePointer = 0
    cBase.CommitTrans    'FIN DE TRANSACCION------------------------------------------
    
    NotificoCambioSignalR
    sFacturada = True       'Para poder salir sin preguntar
    Unload Me
    
    Screen.MousePointer = 0
    Exit Sub

errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bEmpleo_Click()

    If Not bEmpleo.Enabled Then Exit Sub
    On Error GoTo errIr
    Screen.MousePointer = 11
    
    Dim objCl As New clsCliente
    objCl.Empleos idCliente:=gCliente
    Me.Refresh
    Set objCl = Nothing
    
    Exit Sub
errIr:
    clsGeneral.OcurrioError "Ocurrió un error al acceder al formulario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bFicha_Click()

    If Not bFicha.Enabled Then Exit Sub
    On Error GoTo errIr
    Screen.MousePointer = 11
    Dim objCl As New clsCliente
    
    If bEmpleo.Enabled Then objCl.Personas idCliente:=gCliente Else objCl.Empresas idCliente:=gCliente
    Me.Refresh
    Set objCl = Nothing
    
    CargoDatosClienteTitular gCliente
    Screen.MousePointer = 0
    Exit Sub
errIr:
    clsGeneral.OcurrioError "Ocurrió un error al acceder al formulario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bReferencia_Click()

    If Not bReferencia.Enabled Then Exit Sub
    On Error GoTo errIr
    Screen.MousePointer = 11
    
    Dim objCl As New clsCliente
    objCl.Referencias idCliente:=gCliente
    Me.Refresh
    Set objCl = Nothing
    
    Exit Sub
errIr:
    clsGeneral.OcurrioError "Ocurrió un error al acceder al formulario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bTitulo_Click()

    If Not bTitulo.Enabled Then Exit Sub
    On Error GoTo errIr
    Screen.MousePointer = 11
    
    Dim objCl As New clsCliente
    objCl.Titulos idCliente:=gCliente
    Me.Refresh
    Set objCl = Nothing
    
    Exit Sub
errIr:
    clsGeneral.OcurrioError "Ocurrió un error al acceder al formulario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bValidar_Click()
    
    On Error GoTo errValidar
    
    'Controles eFactura.
    
    '23/10/2014 no dejamos facturar más empresas sin rut, separo condición para dar msg más claro.
    
    If oCliente.TipoCliente = TC_Empresa And oCliente.RUT = "" Then
        MsgBox "Las empresas deben aportar su número de RUT para que le podamos facturar.", vbCritical, "Validación eFactura"
        Exit Sub
    End If
    
    If (oCliente.TipoCliente = TC_Empresa) Or _
        (oCliente.TipoCliente = TC_Persona And (CCur(lTotal.Caption) > prmImporteConInfoCliente Or oCliente.RUT <> "")) Then
    
        If gDirFactura = 0 Then
            MsgBox "Debe ingresar una dirección del cliente.", vbExclamation, "Validación EFactura"
            Exit Sub
        End If
            
        If (oCliente.RUT = "" And oCliente.CI = "") Then
            MsgBox "Para facturar es necesario ingresar cédula o RUT.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        
    End If
    
    If (chkVencidaTit.value Or chkVencidaConyuge.value) And InStr(1, "," & Replace(paLocalesCobraVencidas, " ", "") & ",", "," & paCodigoDeSucursal) = 0 Then
        MsgBox "Su local no puede facturar cuotas vencidas.", vbExclamation, "Cuotas vencidas"
        Exit Sub
    End If
    
    'Si el artículo que se factura es refinanciación entonces solicitamos los documentos que se refinancian.
    Dim sRefinanciacion As String
    Dim iCasoRefinanciacion As Integer
    iCasoRefinanciacion = 0
    
    For Each itmX In lvVenta.ListItems
        If ArticuloDeLaClave(itmX.Key) = 2544 Then
            iCasoRefinanciacion = 1
            Exit For
        End If
    Next
    
    Dim sIDsRefinanciacion As String
    Dim sDocsRefinanciacion As String
    
    If iCasoRefinanciacion = 1 Then
        sRefinanciacion = InputBox("Ingrese el o los documentos que se están refinanciando." & vbCrLf & vbCrLf & "Ejemplo de ingreso: D123456 o D1234,D777 (cuando son más de 1 van separados por coma).", "Refinanciación", "Dxxxxx")
        If sRefinanciacion <> "SinGestor" And sRefinanciacion <> "" Then
            Set RsAux = cBase.OpenResultset("EXEC prg_RelacionaRefinanciaciones '" & sRefinanciacion & "', Null", rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                If RsAux(0) = 0 Then
                    RsAux.Close
                    MsgBox "No se encontró un documento que cumpla la condición de refinanciación para los datos ingresados.", vbExclamation, "Posible error"
                    Exit Sub
                End If
            End If
            Do While Not RsAux.EOF
                iCasoRefinanciacion = RsAux(0)
                If sDocsRefinanciacion <> "" Then
                    sDocsRefinanciacion = sDocsRefinanciacion & ", "
                    sIDsRefinanciacion = sIDsRefinanciacion & ","
                End If
                sDocsRefinanciacion = sDocsRefinanciacion & RsAux(2)
                sIDsRefinanciacion = sIDsRefinanciacion & RsAux(1)
                RsAux.MoveNext
            Loop
            RsAux.Close
            If iCasoRefinanciacion = -1 Then
                MsgBox "ATENCIÓN!!!" & vbCrLf & "Debe realizar la nota especial para el o los siguientes documentos: " & vbCrLf & vbCrLf & sDocsRefinanciacion, vbExclamation, "ATENCIÓN"
                Exit Sub
'            Else
'                If Replace(Replace(LCase(sDocsRefinanciacion), " ", ""), "-", "") <> Replace(Replace(LCase(sRefinanciacion), " ", ""), "-", "") Then
'                    sRefinanciacion   ACA ME FALTA CONSULTAR CON EL USUARIO.
'                End If
            End If
        ElseIf sRefinanciacion <> "SinGestor" And sRefinanciacion = "" Then
            MsgBox "Debe ingresar el documento al cual se refinancia.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
    End If
    
    'Valido si todos articulos estan habilitados para la venta----------------------------------------------
    gSucesoDesc = ""
    Dim bSalir As Boolean, bDeposito As Boolean
    bDeposito = False
    For Each itmX In lvVenta.ListItems
        If aIDServicio = 0 Then
            bSalir = False
            Cons = "Select * from Articulo Where ArtId = " & ArticuloDeLaClave(itmX.Key)
            'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
            If Not RsAux.EOF Then
                If IsNull(RsAux!ArtHabilitado) Or RsAux!ArtHabilitado <> "S" Then
                    
                    '21/1/2002 Si el Articulo esta deshabilitado, valido si pertenece a un combo
                    Dim bEsDeCombo As Boolean
                    bEsDeCombo = ArticuloDeCombo(RsAux!ArtID, PlanDeLaClave(itmX.Key))
                    If Not bEsDeCombo Then
                        gSucesoDesc = gSucesoDesc & Format(RsAux!ArtCodigo, "#,000,000") & ", "
                        If MsgBox("El artículo " & Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre) & ", no está habilitado para la venta." & vbCrLf & _
                                        "Desea continuar con la venta.", vbExclamation + vbYesNo, "Artículo No Habilitado") = vbNo Then bSalir = True
                    End If
                End If
                
                'Si hay para retirar y es en depósito consulto.
'                If Not IsNull(RsAux("ArtLocalRetira")) And paCodigoDeSucursal = 5 Then
'                    If (RsAux("ArtLocalRetira") = 6) And Trim(itmX.SubItems(11)) = "" Then
'                        'No hay envíos y retira en depósito
'                      bDeposito = True
'                    End If
'                End If
                
            End If
            RsAux.Close
            If bSalir Then Screen.MousePointer = 0: Exit Sub
        End If
    Next
    gSucesoUsr = 0
    If Trim(gSucesoDesc) <> "" Then
        gSucesoDesc = "Cód. " & Mid(gSucesoDesc, 1, Len(gSucesoDesc) - 2)
        Dim objSuceso As New clsSuceso
        objSuceso.ActivoFormulario paCodigoDeUsuario, "Artículos No Habilitados p/Venta", cBase
        Me.Refresh
        gSucesoUsr = objSuceso.RetornoValor(Usuario:=True)
        gSucesoDef = objSuceso.RetornoValor(Defensa:=True)
        Set objSuceso = Nothing
        If gSucesoUsr = 0 Then Screen.MousePointer = 0: Exit Sub 'Abortó el ingreso del suceso
    End If
    
    If Not ValidoVigenciaPrecios(prmIDSolicitud) Then         'Suceso para Precios Viejos
        Dim objSucesoPV As New clsSuceso
        objSucesoPV.TipoSuceso = prmSuc_ModificacionDePrecios
        objSucesoPV.ActivoFormulario paCodigoDeUsuario, "Factura con Precios No Vigentes", cBase
        
        Me.Refresh
        spv_Usuario = objSucesoPV.RetornoValor(Usuario:=True)
        spv_Defensa = objSucesoPV.RetornoValor(Defensa:=True)
        spv_Autoriza = objSucesoPV.Autoriza
        
        Set objSucesoPV = Nothing
        If spv_Usuario = 0 Then Screen.MousePointer = 0: Exit Sub 'Abortó el ingreso del suceso
    End If
    
    'Primero Valido la Entrega----------------------------------------------------------------------------
    If tEntregaT.Enabled Then
        If Not IsNumeric(tEntregaT.Text) Then
            MsgBox "El monto de entrega ingresado no es correcto.", vbCritical, "ATENCIÓN"
            Foco tEntregaT: Exit Sub
        End If
        If CCur(tEntregaT.Text) < CCur(tEntregaT.Tag) Then
            MsgBox "El monto de entrega no debe ser menor al original.", vbCritical, "ATENCIÓN"
            Foco tEntregaT: Exit Sub
        End If
    End If
    '---------------------------------------------------------------------------------------------------------
    If Not IsDate(tFRetiro.Text) Then
        MsgBox "La fecha de retiro de mercadería ingresada no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tFRetiro: Exit Sub
    End If
    
    If Val(tUsuario.Tag) = 0 Then
        MsgBox "Ingrese el dígito de usuario para realizar la operación.", vbExclamation, "ATENCIÓN"
        Foco tUsuario: Exit Sub
    End If

    If aSolicitudEstado <> EstadoSolicitud.Condicional And aSolicitudEstado <> EstadoSolicitud.Aprovada Then
        MsgBox "El estado de la solicitud no permite ser facturada.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If

    'Si paso la Validación controlo la direccion que factura-----------------------------------------------------------------
    If cDireccion.ListIndex <> -1 Then
        On Error Resume Next
        If gDirFactura <> cDireccion.ItemData(cDireccion.ListIndex) Then        'Cambio Dir Facutua
            If MsgBox("Ud. a cambiado la dirección con la que el cliente factura habitualmente." & vbCrLf & "Quiere que esta dirección quede por defecto para facturar.", vbQuestion + vbYesNo, "Dirección por Defecto al Facturar") = vbNo Then Exit Sub
            
            If cDireccion.ItemData(cDireccion.ListIndex) <> Val(cDireccion.Tag) Then        'Dir. selecc. <> a la Ppal.
                
                Dim rsAD As rdoResultset
                Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & gCliente & " And DAuDireccion = " & cDireccion.ItemData(cDireccion.ListIndex)
                Set rsAD = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rsAD.EOF Then
                    rsAD.Edit: rsAD!DAuFactura = True: rsAD.Update
                End If
                rsAD.Close
            End If
            
            If gDirFactura <> Val(cDireccion.Tag) Then      'La gDirFactura Anterior no era la ppal, la desmarco
                Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & gCliente & " And DAuDireccion = " & gDirFactura
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
    
    On Error GoTo errValidar
    BuscoPreciosContado
    
    HayCuotasVencidas
    
    'Suceso p/ facturacion de Articulos no Habilitados
    
    Dim resp As Byte
    Dim subresp As Byte
    resp = 255
    
'    If bDeposito And Not HayEnviosAProcesar Then
'        Dim fQVino As New frmEnQueVino
'        fQVino.Show vbModal
'        resp = fQVino.Respuesta
'        subresp = fQVino.SubRespuesta
'    End If
    
    If aSolicitudEstado = EstadoSolicitud.Condicional Then
        FechaDelServidor
        AccionValidar sIDsRefinanciacion, resp, subresp
        
    Else
    
        If aSolicitudEstado = EstadoSolicitud.Aprovada Then
            
            Dim mTPrinters As String
            mTPrinters = vbCrLf & vbCrLf & _
                                "Imprimir Créditos en: " & paICreditoN & vbCrLf & _
                                "Imprimir Conformes en: " & paIConformeN
            If MsgBox("La solicitud ha sido aprobada. " & "¿Desea facturarla? " & mTPrinters, vbQuestion + vbYesNo, "Facturar Solicitud") = vbYes Then
            
                If HayEnviosAProcesar Then
                    ProcesoEnvios
                    If MsgBox("¿Confirma realizar la emisión de las facturas?" & mTPrinters, vbQuestion + vbYesNo, "Emitir Facturas") = vbYes Then AccionFacturar sIDsRefinanciacion, resp, subresp
                Else
                    AccionFacturar sIDsRefinanciacion, resp, subresp
                End If
            
            End If
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub

errValidar:
    clsGeneral.OcurrioError "Ocurrió un error al validar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub cDireccion_Click()
On Error GoTo errCargar

    If cDireccion.ListIndex <> -1 Then
        Screen.MousePointer = 11
        lDireccion.Caption = ""
        lDireccion.Caption = clsGeneral.ArmoDireccionEnTexto(cBase, cDireccion.ItemData(cDireccion.ListIndex))
        If Not oCliente Is Nothing Then
            oCliente.Direccion.Domicilio = lDireccion.Caption
        End If
        Screen.MousePointer = 0
    End If

errCargar:
    Screen.MousePointer = 0
End Sub

Private Sub cCuota_Click()
    lSubTotalF.Caption = ""
    tEntrega.Text = ""
    If cCuota.Enabled Then aPlan = 0
End Sub

Private Sub cCuota_GotFocus()
    cCuota.SelStart = 0
    cCuota.SelLength = Len(cCuota.Text)
End Sub

Private Sub cCuota_KeyDown(KeyCode As Integer, Shift As Integer)

    lSubTotalF.Caption = ""
    tEntrega.Text = ""
    aPlan = 0
    
    If KeyCode = vbKeyEscape Then
        LimpioRenglon
        HabilitoRenglon False
        lvVenta.Enabled = True
        lvVenta.SetFocus
    End If
    
End Sub

Private Sub cCuota_KeyPress(KeyAscii As Integer)
    
    lSubTotalF.Caption = ""
    tEntrega.Text = ""
    aPlan = 0
    
    If KeyAscii = vbKeyReturn And cCuota.ListIndex <> -1 Then
        On Error GoTo ErrBAC
        Dim sOk As Boolean
        
        Screen.MousePointer = 11
        sOk = True
        'Busco los datos del tipo de cuota seleccionado y veo si puedo hacerlo
        Cons = "Select * from TipoCuota Where TCuCodigo = " & cCuota.ItemData(cCuota.ListIndex)
        'Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
        If RsAux!TCuCantidad > CLng(lvVenta.SelectedItem.Tag) Then
            Screen.MousePointer = 0
            MsgBox "No puede realizar la operación en un plan con mayor cantidad de cuotas.", vbExclamation, "ATENCIÓN"
            sOk = False
        Else
            If RsAux!TCuCantidad = CLng(lvVenta.SelectedItem.Tag) Then
                If Mid(lvVenta.SelectedItem.Key, 1, 1) <> "E" And Not IsNull(RsAux!TCuVencimientoE) Then
                    Screen.MousePointer = 0
                    MsgBox "No puede realizar la operación en un plan con mayor dificultad al original.", vbExclamation, "ATENCIÓN"
                    sOk = False
                End If
            End If
        End If
        
        If Not sOk Then
            RsAux.Close
            Exit Sub
        End If
        cCuota.Tag = RsAux!TCuCantidad
        If IsNull(RsAux!TCuVencimientoE) Then
            sConEntrega = False
            tEntrega.Enabled = False
            tEntrega.BackColor = Inactivo
            tEntrega.Text = ""
        Else
            sConEntrega = True
            tEntrega.Enabled = True
            tEntrega.BackColor = Obligatorio
            tEntrega.Text = lvVenta.SelectedItem.SubItems(4)
        End If
        RsAux.Close
    
        Screen.MousePointer = 11
        
        If sConEntrega Then
            'Saco el Precio contado del Articulo--------------------------------------------------------------------------
            Cons = "Select PViPrecio, PViPlan  From PrecioVigente" _
                    & " Where PViArticulo = " & tArticulo.Tag _
                    & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                    & " And PViTipoCuota = " & paTipoCuotaContado
            'Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
            lSubTotalF.Tag = RsAux!PViPrecio
            aPlan = RsAux!PViPlan
            RsAux.Close
            
        Else
            'Saco el valor de la cuota financiado
            Cons = "Select PViPrecio From PrecioVigente" _
                    & " Where PViArticulo = " & tArticulo.Tag _
                    & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                    & " And PViTipoCuota = " & cCuota.ItemData(cCuota.ListIndex)
            'Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If ObtenerResultSet(cBase, RsAux, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Sub
            If RsAux.EOF Then
                Screen.MousePointer = 0
                MsgBox "No se encotraron precios para la financiación seleccionada. Consulte", vbExclamation, "ATENCIÓN"
                RsAux.Close
                Exit Sub
            End If
            lSubTotalF.Tag = RsAux!PViPrecio               'Precio de Unitario Financiaciado
            If IsNumeric(tCantidad.Text) Then
                lSubTotalF.Caption = Format(RsAux!PViPrecio * CLng(tCantidad.Text), FormatoMonedaP)
            Else
                tCantidad.Text = ""
            End If
            RsAux.Close
        End If
        Screen.MousePointer = 0
        tCantidad.SetFocus
    End If
    Exit Sub

ErrBAC:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar los datos de financiacion."
End Sub

Private Sub cCuota_LostFocus()
    cCuota.SelLength = 0
End Sub

Private Sub cEMailsT_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If chkRetiraAqui.Visible Then
            chkRetiraAqui.SetFocus
        Else
            tUsuario.SetFocus
        End If
    End If
End Sub

Private Sub chkRetiraAqui_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then Foco tUsuario
End Sub

Private Sub cPago_GotFocus()
    cPago.SelStart = 0
    cPago.SelLength = Len(cPago.Text)
End Sub

Private Sub cPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cPendiente
End Sub

Private Sub cPago_LostFocus()
    cPago.SelLength = 0
End Sub

Private Sub cPendiente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tFRetiro
End Sub

Private Sub cPendiente_LostFocus()
    cPendiente.SelLength = 0
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF6: Call bFicha_Click
        Case vbKeyF7: If bEmpleo.Enabled Then Call bEmpleo_Click
        Case vbKeyF8: Call bTitulo_Click
        Case vbKeyF9: Call bReferencia_Click
        Case vbKeyF12
            If gCliente <> 0 Then EjecutarApp prmPathApp & "\Visualizacion de operaciones.exe", CStr(gCliente)
    End Select
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrLoad
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Height = 7110    '6730    '6480

    oCnfgCredito.CargarConfiguracion CGSA_TickeadoraCredito, "CuotasImpresora"
    If Val(oCnfgCredito.ImpresoraTickets) = 0 Then
        MsgBox "No tiene asignada una impresora para imprimir el ticket. CANCELE y en la grilla presione botón derecho con el mousse y seleccione ¿dónde imprimo eFactura?", vbExclamation, "ATENCIÓN"
    End If
    oCnfgRecibo.CargarConfiguracion "ImpresionDocumentos", "TicketCuota"

    SituacionRetiraAqui
    
    sFacturada = False
    aFletes = ""
    aIDServicio = 0
    
    tFRetiro.Text = Format(Now, "d-Mmm yyyy")
    tFRetiro.Tag = tFRetiro.Text
    
    cEMailsT.OpenControl cBase
    cEMailsT.IDUsuario = paCodigoDeUsuario
    
    LimpioFicha
    HabilitoRenglon False
    SetearLView lvValores.Grilla Or lvValores.FullRow, lvVenta
    SetearLView lvValores.Grilla Or lvValores.FullRow, lEnvio
    
    'Cargo las monedas ------------------------------------------------------------------------------------
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda, ""
    '-----------------------------------------------------------------------------------------------------------
    'Cargo las formas de pago-----------------------------------------------------------------------------
    cPago.AddItem "Efectivo": cPago.ItemData(cPago.NewIndex) = TipoPagoSolicitud.Efectivo
    cPago.AddItem "Cheques Diferidos": cPago.ItemData(cPago.NewIndex) = TipoPagoSolicitud.ChequeDiferido
    '-----------------------------------------------------------------------------------------------------------
    'Cargo los codigos de pendiente------------------------------------------------------------------------
    Cons = "Select PEnCodigo, PEnNombre From PendienteEntrega Order by PEnNombre"
    CargoCombo Cons, cPendiente, ""
    '-----------------------------------------------------------------------------------------------------------
    
    CargoSolicitud prmIDSolicitud
    
    CargoRenglonSolicitud prmIDSolicitud
    
    Select Case aSolicitudEstado
        Case EstadoSolicitud.Condicional
            Me.Caption = "Solicitud Condicional"
            iCondicional.Visible = True
        
        Case EstadoSolicitud.Aprovada
            Me.Caption = "Solicitud Aprobada"
            iSi.Visible = True
            
        Case EstadoSolicitud.Rechazada
            Me.Caption = "Solicitud Rechazada"
            iNo.Visible = True
    End Select
    
    If aSolicitudEstado = EstadoSolicitud.Rechazada Then
        If Trim(lUsuario.Caption) = "S/D" Then      'Solicitud  Devuelta
            cPendiente.Enabled = False: cPendiente.BackColor = Inactivo
            tFRetiro.Enabled = False: tFRetiro.BackColor = Inactivo
            tUsuario.Enabled = False: tUsuario.BackColor = Inactivo
            tComentario.Enabled = False: tComentario.BackColor = Inactivo
            bDevolver.Enabled = False
        End If
        
        lvVenta.Enabled = False
        cPago.Enabled = False: cPago.BackColor = Inactivo
        tEntregaT.Enabled = False: tEntregaT.BackColor = Inactivo
        bValidar.Enabled = False
    End If
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Error al cargar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error GoTo errSalir
    
    If sFacturada Then Exit Sub
    
    If aSolicitudEstado <> EstadoSolicitud.Rechazada Then
        If MsgBox("Confirma salir de la solicitud y no facturarla.", vbQuestion + vbYesNo, "SALIR") = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    'Bloqueo la solicitud y Actulizo el SolTipoResolucion a Manual
    Cons = "Select * from Solicitud Where SolCodigo = " & prmIDSolicitud
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        RsAux.Edit
        RsAux!SolProceso = TipoResolucionSolicitud.Manual
        RsAux.Update
    End If
    RsAux.Close
    
    'Hay que eliminar los articulos enviados
    For Each itmX In lvVenta.ListItems
        If Trim(itmX.SubItems(11)) <> "" Then BorroEnvios Trim(itmX.SubItems(11))
    Next
    Screen.MousePointer = 0
    Exit Sub

errSalir:
    Screen.MousePointer = 0
    Cancel = 1
    clsGeneral.OcurrioError "Ocurrió un error al modificar el estado de la solicitud."
End Sub

'-------------------------------------------------------------------------------------------------------------------
'   ABREVIACIONES:
'       FCO         -> Firmando con el cónyuge
'       FGA         -> Garantía predeterminada........FGA y codigo de cliente garante  (FGA10562)
'       GPR         -> Garantía propietaria
'       PRO         -> Titular propietario
'       CPL          -> Cambio de Plan...........CPL y codigo de plan (CPL8) ó (CPL1E1500) - con entrega
'       CHD         -> Pago con cheque diferido
'       MON        -> Por un valor máximo a XXXX ....... MON5600
'       COM        -> Exhibir Comprobante  ... COM y cod. comp V16000V28000 (Valor uno y/o valor 2 - montos mínimos)
'       RSU         -> Recibo de Sueldo RSU , cód. Moneda , Monto .... RSU1M6500  RSU $ Monto:6500
'       CLE          -> Con Clearing
'       CDE         -> Cancelar deuda pendiente
'       COP         -> Cancelar operaciones pendiente

'-------------------------------------------------------------------------------------------------------------------
Private Function InterpretoTexto(Condicion As String) As String

Dim sAuxiliar As String
Dim sCondicion As String
Dim sRetorno As String
Dim RsI As rdoResultset
Dim RsCI As rdoResultset

    On Error GoTo errInterpreto
    
    gTextoCondicion = Trim(Condicion)
    sRetorno = " ("
    Do While gTextoCondicion <> ""
        sAuxiliar = SacoCondicionDelTexto
        If sAuxiliar <> "" Then
            sCondicion = UCase(Mid(sAuxiliar, 1, 3))                'Las primeras 3 son la Condicion
            sAuxiliar = Trim(Mid(sAuxiliar, 4, Len(sAuxiliar)))    'El resto son los valores
            
            Select Case sCondicion
                Case "RSU"
                    'Saco la Moneda --------------------------------
                    Cons = "Select * from Moneda Where MonCodigo = " & Mid(sAuxiliar, 1, InStr(sAuxiliar, "M") - 1)
                    'Set RsI = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If ObtenerResultSet(cBase, RsI, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
                    sAuxiliar = Trim(Mid(sAuxiliar, InStr(sAuxiliar, "M") + 1, Len(sAuxiliar)))
                    sRetorno = sRetorno & "Recibo Sueldo x " & Trim(RsI!MonSigno) & " " & Format(sAuxiliar, "#,##0.00")
                    RsI.Close
                
                Case "CPL"
                    'Cambio de Plan --------------------------------
                    If InStr(sAuxiliar, "E") <> 0 Then
                        Cons = "Select * from TipoCuota Where TcuCodigo = " & Mid(sAuxiliar, 1, InStr(sAuxiliar, "E") - 1)
                        sAuxiliar = Trim(Mid(sAuxiliar, InStr(sAuxiliar, "E") + 1, Len(sAuxiliar)))
                    Else
                        Cons = "Select * from TipoCuota Where TCuCodigo = " & sAuxiliar
                        sAuxiliar = ""
                    End If
                    If ObtenerResultSet(cBase, RsI, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
                    'Set RsI = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    sRetorno = sRetorno & "Plan: " & Trim(RsI!TCuAbreviacion)
                    If sAuxiliar <> "" Then sRetorno = sRetorno & " Entrega: " & Format(sAuxiliar, "#,##0.00")
                    RsI.Close
                    
                Case "MON"
                    sRetorno = sRetorno & "Monto Solicitud: " & Format(sAuxiliar, "#,##0.00")
                
                Case "FGA"
                    Cons = "Select * from Cliente Where CliCodigo = " & sAuxiliar
                    'Set RsI = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If ObtenerResultSet(cBase, RsI, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
                    If RsI!CliTipo = 1 Then
                        sRetorno = sRetorno & "Garantía: " & clsGeneral.RetornoFormatoCedula(RsI!CliCIRuc)
                    Else
                        sRetorno = sRetorno & "Garantía: " & clsGeneral.RetornoFormatoRuc(RsI!CliCIRuc)
                    End If
                    RsI.Close
                    tGarantia.Enabled = True
                    tGarantia.BackColor = Obligatorio
                    lGarantia.BackColor = Obligatorio
                    
                Case "COM"      'PRESENTAR COMPROBANTE
                    Cons = "Select * from Comprobante Where ComCodigo = " & Mid(sAuxiliar, 1, InStr(sAuxiliar, "V1") - 1)
                    'Set RsI = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If ObtenerResultSet(cBase, RsI, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
                    sAuxiliar = Trim(Mid(sAuxiliar, InStr(sAuxiliar, "V1") + 2, Len(sAuxiliar)))
                    If InStr(sAuxiliar, "V2") <> 0 Then
                        
                        'Si el Valor es Cedula -> Hay que buscar el Cliente
                        If UCase(Trim(RsI!ComFormatoV1)) = "CEDULA" Then
                            Cons = "Select * from Cliente Where CliCodigo = " & Mid(sAuxiliar, 1, InStr(sAuxiliar, "V2") - 1)
                            Set RsCI = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                            Cons = RsCI!CliCIRuc
                            RsCI.Close
                        Else
                            Cons = Mid(sAuxiliar, 1, InStr(sAuxiliar, "V2") - 1)
                        End If
                        
                        sRetorno = sRetorno & Trim(RsI!ComNombre) & ": " & Trim(RsI!ComNombreValor1) & ": " & FormatoReferencia(Cons, RsI!ComFormatoV1)
                        
                        sAuxiliar = Trim(Mid(sAuxiliar, InStr(sAuxiliar, "V2") + 2, Len(sAuxiliar)))
                        If UCase(Trim(RsI!ComFormatoV2)) = "CEDULA" Then
                            Cons = "Select * from Cliente Where CliCodigo = " & sAuxiliar
                            Set RsCI = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                            sAuxiliar = Trim(RsCI!CliCIRuc)
                            RsCI.Close
                        End If
                        
                        sRetorno = sRetorno & " - " & Trim(RsI!ComNombreValor2) & ": " & FormatoReferencia(sAuxiliar, RsI!ComFormatoV1)
                        
                    Else
                    
                        'Si el Valor es Cedula -> Hay que buscar el Cliente
                        If UCase(Trim(RsI!ComFormatoV1)) = "CEDULA" Then
                            Cons = "Select * from Cliente Where CliCodigo = " & sAuxiliar
                            Set RsCI = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                            Cons = Trim(RsCI!CliCIRuc)
                            RsCI.Close
                        Else
                            Cons = sAuxiliar
                        End If
                        
                        sRetorno = sRetorno & Trim(RsI!ComNombre) & " " & Trim(RsI!ComNombreValor1) & ": " & FormatoReferencia(Cons, RsI!ComFormatoV1)
                    End If
                    RsI.Close
                    
                Case "FCO"
                    sRetorno = sRetorno & "Firma del Cónyuge"
                    tGarantia.Enabled = True
                    tGarantia.BackColor = Obligatorio
                    lGarantia.BackColor = Obligatorio
                
                Case "GPR"
                    sRetorno = sRetorno & "Garantía Propietaria"
                    tGarantia.Enabled = True
                    tGarantia.BackColor = Obligatorio
                    lGarantia.BackColor = Obligatorio
                    
                Case "PRO": sRetorno = sRetorno & "Exhibir Títulos"
                Case "CLE": sRetorno = sRetorno & "Hacer Clearing"
                Case "CHD": sRetorno = sRetorno & "Pago C/Cheque Dif."
                Case "CDE": sRetorno = sRetorno & "Pago Cuotas Vencidas"
                Case "COP": sRetorno = sRetorno & "Cancelar Créditos Vigentes"
            End Select
            sRetorno = sRetorno & ", "
        Else
            Exit Do
        End If
    Loop
    If sRetorno <> " (" Then
        sRetorno = Mid(sRetorno, 1, Len(sRetorno) - 2) & ")"
        InterpretoTexto = sRetorno
    Else
        InterpretoTexto = ""
    End If
    Exit Function

errInterpreto:
    clsGeneral.OcurrioError "Ocurrió un error al interpretar las condiciones de resolución."
End Function

Public Function CargoSolicitud(Codigo As Long)

Dim mTXT As String
Dim x_IDGarantia As Long

    Screen.MousePointer = 11
    
    x_IDGarantia = 0
    
    tGarantia.Tag = 0
    tCondicion.Tag = ""
    LimpiarCtrlsVencidas
    
    Cons = "Select TOP 1 Solicitud.*, SolicitudResolucion.*, ConNombre, IsNull(ConCancelarDeuda, 0) ConCancelarDeuda" & _
                " From Solicitud, SolicitudResolucion, CondicionResolucion" & _
                " Where SolCodigo = " & Codigo & _
                " And ResSolicitud = SolCodigo And ResCondicion = ConCodigo" & _
                " Order by ResNumero DESC"
    'Set rsSol = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If ObtenerResultSet(cBase, rsSol, Cons, "comercio") <> RAQ_SinError Then Screen.MousePointer = 0: Exit Function
    If Not rsSol.EOF Then
        aSolicitudEstado = rsSol!SolEstado
        gCliente = rsSol!SolCliente
        
        mTXT = Trim(rsSol!ConNombre)
        
        If Not IsNull(rsSol!ResTexto) Then
            tCondicion.Tag = Trim(rsSol!ResTexto)
            mTXT = mTXT & InterpretoTexto(rsSol!ResTexto)
        End If
        
        If Not IsNull(rsSol!ResComentario) Then mTXT = mTXT & IIf(mTXT = "", "", vbCrLf) & Trim(rsSol!ResComentario)
        tCondicion.Text = mTXT
        
        lCodigo.Caption = rsSol!SolCodigo
        lFecha.Caption = Format(rsSol!SolFecha, "d-Mmm-yy hh:mm")
        BuscoCodigoEnCombo cMoneda, rsSol!SolMoneda
        BuscoCodigoEnCombo cPago, rsSol!SolFormaPago
        If Not IsNull(rsSol!ResUsuario) Then lUsuario.Caption = z_BuscoUsuario(rsSol!ResUsuario, True) Else lUsuario.Caption = "S/D"
        If Not IsNull(rsSol!SolUsuarioS) Then lUsuarioO.Caption = z_BuscoUsuario(rsSol!SolUsuarioS, True) Else lUsuarioO.Caption = "S/D"
        
        lVendedor.Tag = ""
        If Not IsNull(rsSol!SolVendedor) Then
            lVendedor.Caption = z_BuscoUsuario(rsSol!SolVendedor, Digito:=True)
            lVendedor.Tag = rsSol!SolVendedor
        End If
        
        If Not IsNull(rsSol!SolIdServicio) Then aIDServicio = rsSol!SolIdServicio Else aIDServicio = 0
        
        CargoDatosClienteTitular rsSol!SolCliente
                
        'Cargo los datos del Cliente Garantia-----------------------------
        If Not IsNull(rsSol!SolGarantia) Then x_IDGarantia = rsSol!SolGarantia
        
        'Guardo si la condición es pagar vencidas o no.
        chkVencidaTit.Tag = IIf(rsSol("ConCancelarDeuda"), 1, 0)

    End If
    rsSol.Close
    
    'Veo condición de vencidas.
    CargarInfoVencidas
    
    If x_IDGarantia > 0 Then
        Cons = "Select * from Cliente " & _
                        " Left Outer Join CPersona On CliCodigo = CPeCliente " & _
                        " Left Outer Join CEmpresa On CliCodigo = CEmCliente " & _
                " Where CliCodigo = " & x_IDGarantia

        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If Not IsNull(RsAux!CliCIRuc) Then
                If RsAux!CliTipo = 1 Then tGarantia.Text = clsGeneral.RetornoFormatoCedula(RsAux!CliCIRuc)
                If RsAux!CliTipo = 2 Then tGarantia.Text = Trim(clsGeneral.RetornoFormatoRuc(RsAux!CliCIRuc))
            End If
            If RsAux!CliTipo = 1 Then
                lGarantia.Caption = " " & ArmoNombre(Format(RsAux!CPeApellido1, "#"), Format(RsAux!CPeApellido2, "#"), Format(RsAux!CPeNombre1, "#"), Format(RsAux!CPeNombre2, "#"))
            Else
                lGarantia.Caption = " " & Trim(RsAux!CEmFantasia)
            End If
            tGarantia.Tag = x_IDGarantia
        Else
            MsgBox "Esta solicitud tiene como garantía al cliente " & x_IDGarantia & ". Este cliente no existe en la base de datos." & vbCrLf & "Consultar con Carlos", vbExclamation, "Posible Error "
            tGarantia.Tag = 0
        End If
        RsAux.Close
        
    End If
    
    
    
    'cons = "Select * from Solicitud Where SolCodigo = " & Codigo
    'Set rsSol = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    'If Not rsSol.EOF Then
    '    aSolicitudEstado = rsSol!SolEstado
    '    If Not IsNull(rsSol!SolComentarioR) Then tCondicion.Text = tCondicion.Text & Trim(rsSol!SolComentarioR)
    '    gCliente = rsSol!SolCliente
'        lCodigo.Caption = rsSol!SolCodigo
'        lFecha.Caption = Format(rsSol!SolFecha, "d-Mmm-yy hh:mm")
'        BuscoCodigoEnCombo cMoneda, rsSol!SolMoneda
'        BuscoCodigoEnCombo cPago, rsSol!SolFormaPago
'        If Not IsNull(rsSol!SolUsuarioR) Then lUsuario.Caption = z_BuscoUsuario(rsSol!SolUsuarioR, True) Else lUsuario.Caption = "S/D"
'        If Not IsNull(rsSol!SolUsuarioS) Then lUsuarioO.Caption = z_BuscoUsuario(rsSol!SolUsuarioS, True) Else lUsuarioO.Caption = "S/D"
        
'        lVendedor.Tag = ""
'        If Not IsNull(rsSol!SolVendedor) Then
'            lVendedor.Caption = z_BuscoUsuario(rsSol!SolVendedor, Digito:=True)
'            lVendedor.Tag = rsSol!SolVendedor
'        End If
        
'        If Not IsNull(rsSol!SolIdServicio) Then aIDServicio = rsSol!SolIdServicio Else aIDServicio = 0
        
'        CargoDatosClienteTitular rsSol!SolCliente
                
        'Cargo los datos del Cliente Garantia-----------------------------
'        If Not IsNull(rsSol!SolGarantia) Then
'            cons = "Select * from Cliente, CPersona" _
                   & " Where CliCodigo = " & rsSol!SolGarantia _
                   & " And CliCodigo = CPeCliente "
                   
'            Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
'            If Not rsAux.EOF Then
'                If Not IsNull(rsAux!CliCIRuc) Then tGarantia.Text = clsGeneral.RetornoFormatoCedula(rsAux!CliCIRuc)
'                lGarantia.Caption = " " & ArmoNombre(Format(rsAux!CPeApellido1, "#"), Format(rsAux!CPeApellido2, "#"), Format(rsAux!CPeNombre1, "#"), Format(rsAux!CPeNombre2, "#"))
'                tGarantia.Tag = rsSol!SolGarantia
'            Else
'                MsgBox "Esta solicitud tiene como garantía al cliente " & rsSol!SolGarantia & ". Este cliente no existe en la base de datos." & vbCrLf & "Consultar con Carlos", vbExclamation, "Posible Error "
'                tGarantia.Tag = 0
'            End If
'            rsAux.Close
'
'        Else
'            tGarantia.Tag = 0
'        End If
        
'    End If
    
'    rsSol.Close
    
    mMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    mRedondeo = dis_arrMonedaProp(mMoneda, enuMoneda.pRedondeo)
    
    Screen.MousePointer = 0
    Exit Function
    
errSolicitud:
    clsGeneral.OcurrioError "Error al cargar los datos de la solicitud.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function CargoObjetoContacto() As clsClienteCFE
   
    Dim oContacto As New clsClienteCFE
    With oContacto
        .Codigo = RsAux("CliCodigo")
        .TipoCliente = RsAux("CliTipo")
        .CodigoDGICI = RsAux("PDDTipoDocIdentidad")
        If Not IsNull(RsAux!CliCategoria) Then .Categoria.Codigo = RsAux!CliCategoria
        Dim sNombre As String
        Select Case RsAux!CliTipo
            Case TipoCliente.Cliente
                If Not IsNull(RsAux!CliCIRuc) Then .CI = RsAux("CliCiRUC")
                
                sNombre = Trim(RsAux("CPeApellido1"))
                If Not IsNull(RsAux("CPeApellido2")) Then sNombre = sNombre & " " & Trim(RsAux!CPeApellido2)
                sNombre = sNombre & ", " & Trim(RsAux("CPeNombre1"))
                If Not IsNull(RsAux("CPeNombre2")) Then sNombre = sNombre & " " & Trim(RsAux!CPeNombre2)
                If Not IsNull(RsAux!CPERuc) Then .RUT = Trim(RsAux("CPeRUC"))
            
            Case TipoCliente.Empresa
                If Not IsNull(RsAux!CliCIRuc) Then
                    .RUT = Trim(RsAux("CliCIRUC"))
                End If
                
                If Not IsNull(RsAux!CEmNombre) Then
                    sNombre = Trim(RsAux!CEmNombre)
                Else
                    sNombre = Trim(RsAux!CEmFantasia)
                End If
        End Select
        .NombreCliente = sNombre
    End With
    Set CargoObjetoContacto = oContacto

End Function

Private Sub CargoDatosClienteTitular(ByVal idCliente As Long)
    
    cDireccion.Clear: cDireccion.Tag = 0
    gDirFactura = 0
        
    Set oCliente = Nothing
    Set oCliente = New clsClienteCFE
    
    'Cargo los datos del Cliente Titular-----------------------------
    Cons = "SELECT * FROM Cliente LEFT OUTER JOIN CPersona ON CliCodigo = CPeCliente " _
            & "LEFT OUTER JOIN CEmpresa ON CliCodigo = CEmCliente " _
            & "INNER JOIN PaisDelDocumento ON PDDId = CliPaisDelDocumento " _
           & " Where CliCodigo = " & idCliente
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Set oCliente = CargoObjetoContacto()
    
    If Not IsNull(RsAux!CliDireccion) Then
        cDireccion.AddItem "Dirección Principal": cDireccion.ItemData(cDireccion.NewIndex) = RsAux!CliDireccion
        cDireccion.Tag = RsAux!CliDireccion: gDirFactura = RsAux!CliDireccion
    End If
    RsAux.Close
    
    If oCliente.TipoCliente = TC_Empresa Then
        bEmpleo.Enabled = False
        oDireccion.value = vbChecked
    Else
        oDireccion.value = vbUnchecked
        If oCliente.CI <> "" Then lCi.Caption = clsGeneral.RetornoFormatoCedula(oCliente.CI)
    End If
    
    lNombre.Caption = oCliente.NombreCliente
    If oCliente.RUT <> "" Then
        Dim oValida As New clsValidaRUT
        If oValida.ValidarRUT(oCliente.RUT) Then
            lRuc.Caption = clsGeneral.RetornoFormatoRuc(oCliente.RUT)
            
            If oCliente.TipoCliente = TC_Persona Then
                Dim rsP As VbMsgBoxResult
                rsP = vbCancel
                Do While rsP = vbCancel
                    rsP = MsgBox("CLIENTE UNIPERSONAL" & vbCrLf & vbCrLf & "¿El cliente desea facturar con RUT?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "FACTURAR CON RUT")
                Loop
                If rsP = vbNo Then
                    oCliente.RUT = ""
                    lRuc.Caption = ""
                    rsP = vbCancel
                    Do While rsP = vbCancel
                        rsP = MsgBox("¿El cliente aún posee ese RUT?" & vbCrLf & vbCrLf & "Si responde NO el RUT se eliminará de la ficha del cliente.", vbQuestion + vbYesNoCancel + vbDefaultButton3, "RUT EN USO")
                    Loop
                    If rsP = vbNo Then
                        'Updateo la tabla CPERSONA y registro el suceso de cambio de RUT.
                        Cons = "UPDATE CPersona SET CPERuc = NULL WHERE CPeCliente = " & oCliente.Codigo
                        cBase.Execute Cons
                    End If
                End If
            End If
            
        Else
            MsgBox "El RUT que posee el cliente no es correcto debe validarlo y corregirlo.", vbExclamation, "ATENCIÓN"
            oCliente.RUT = ""
        End If
        Set oValida = Nothing
    End If
    
    CargoDireccionesAuxiliares oCliente.Codigo

    cEMailsT.CargarDatos oCliente.Codigo
    
End Sub

Private Function ObtenerMayorFechaRetiro() As Date
On Error GoTo errOFR
Dim itmA  As ListItem
Dim arrValor() As String
    ObtenerMayorFechaRetiro = DateSerial(2000, 1, 1)
    For Each itmA In lvVenta.ListItems
        arrValor = Split(itmA.SubItems(16), "|")
        If UBound(arrValor) >= 1 Then
            If CDate(arrValor(1)) > ObtenerMayorFechaRetiro Then
                ObtenerMayorFechaRetiro = CDate(arrValor(1))
            End If
        End If
    Next
    Exit Function
errOFR:
    clsGeneral.OcurrioError "Error al buscar la fecha disponible de los artículos, se pondrá por defecto la última encontrada.", Err.Description, "Buscar fecha disponible"
End Function


Private Sub CargoRenglonSolicitud(Codigo As Long)

Dim aTotal As Currency
Dim aEntrega As Currency
Dim RsIva As rdoResultset

Dim datDisponibleDesde As Date

    On Error GoTo errRenglon
    Screen.MousePointer = 11
    aTotal = 0
    aEntrega = 0
    lvVenta.ListItems.Clear
    Cons = "Select RenglonSolicitud.*, ArtID, ISNull(AEsNombre, ArtNombre) as ArtNombre, ArtInstalador, TCuAbreviacion, TCuCantidad, ArtDisponibleDesde, IsNull(ArtDemoraEntrega, 0) ArtDemoraEntrega, IsNull(AEsID, 0) AEsID" _
            & " From RenglonSolicitud " & _
                    " LEFT OUTER JOIN ArticuloEspecifico ON RSoSolicitud = AEsDocumento And AEsTipoDocumento = 2 And RSoArticulo = AEsArticulo, " & _
                " Articulo, TipoCuota" _
            & " Where RSoSolicitud = " & Codigo _
            & " And RSoArticulo = ArtID AND RSoDocumento IS NULL" _
            & " And RSoTipoCuota = TCuCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Do While Not RsAux.EOF
        'E ó N - Con entrega ó Normal  ..... P - Plan .....A - Artículo
        'E/N ...Plan ...Articulo..
        
        If Not IsNull(RsAux!RSoValorEntrega) Then
            Set itmX = lvVenta.ListItems.Add(, "EP" & RsAux!RSoTipoCuota & "A" & RsAux!RSoArticulo, Trim(RsAux!TCuAbreviacion))
        Else
            Set itmX = lvVenta.ListItems.Add(, "NP" & RsAux!RSoTipoCuota & "A" & RsAux!RSoArticulo, Trim(RsAux!TCuAbreviacion))
        End If
        
        itmX.Tag = RsAux!TCuCantidad    'Cantidad de Cuotas
        
        itmX.SubItems(1) = RsAux!RSoCantidad
        itmX.SubItems(2) = Trim(RsAux!ArtNombre)
        
        'Saco el I.V.A. del artículo--------------------------------------------------------------
        Cons = "Select IVAPorcentaje From ArticuloFacturacion, TipoIva " _
                & " Where AFaArticulo = " & RsAux!ArtID _
                & " And AFaIVA = IVACodigo"
        Set RsIva = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If RsIva.EOF Then
            itmX.SubItems(3) = "0.00"
        Else
            itmX.SubItems(3) = Format(RsIva!IVaPorcentaje, "#,##0.00")
        End If
        RsIva.Close '-------------------------------------------------------------------------------
        
        itmX.SubItems(7) = "No"
        
        If Not IsNull(RsAux!RSoValorEntrega) Then
            itmX.SubItems(4) = Format(RsAux!RSoValorEntrega, "#,##0.00")
            aEntrega = aEntrega + RsAux!RSoValorEntrega
        End If
        
        itmX.SubItems(5) = Format(RsAux!RSoValorCuota, "#,##0.00")
        
        If Not IsNull(RsAux!RSoValorEntrega) Then
            itmX.SubItems(6) = Format(RsAux!RSoValorEntrega + RsAux!RSoValorCuota * RsAux!TCuCantidad, "#,##0.00")
        Else
            itmX.SubItems(6) = Format(RsAux!RSoValorCuota * RsAux!TCuCantidad, "#,##0.00")
        End If
        
        itmX.SubItems(14) = IIf(Not IsNull(RsAux!ArtInstalador), RsAux!ArtInstalador, 0)
        
        '14/8/2008      cargo el id del artículo específico
        '               cargo a partir de que fecha está disponible el artículo.
        itmX.SubItems(16) = RsAux("AEsID")
    
        datDisponibleDesde = DateSerial(2000, 1, 1)
        If RsAux("AEsID") = 0 Then
            If Not IsNull(RsAux("ArtDisponibleDesde")) Or RsAux("ArtDemoraEntrega") > 0 Then
                datDisponibleDesde = Date
            End If
            If Not IsNull(RsAux("ArtDisponibleDesde")) Then
                If RsAux("ArtDisponibleDesde") > Date Then
                    datDisponibleDesde = RsAux("ArtDisponibleDesde")
                End If
            End If
            datDisponibleDesde = DateAdd("d", RsAux("ArtDemoraEntrega"), datDisponibleDesde)
        End If
        itmX.SubItems(16) = itmX.SubItems(16) & "|" & datDisponibleDesde
        aTotal = aTotal + CCur(itmX.SubItems(6))
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    tFRetiro.Text = ObtenerMayorFechaRetiro()
    If CDate(tFRetiro.Text) < Date Then
        tFRetiro.Text = Format(Date, "d-Mmm yyyy")
    Else
        tFRetiro.Text = Format(CDate(tFRetiro.Text), "d-Mmm yyyy")
    End If

    lTotal.Caption = Format(aTotal, "#,##0.00")
    tEntregaT.Text = Format(aEntrega, "#,##0.00")
    tEntregaT.Tag = aEntrega
    If aEntrega = 0 Then
        tEntregaT.Enabled = False
        tEntregaT.BackColor = Inactivo
    End If
    Screen.MousePointer = 0
    Exit Sub

errRenglon:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al cargar los articulos de la solicitud.", Err.Description
End Sub

Private Sub LimpioFicha()

    lCi.Caption = "S/D"
    lRuc.Caption = "S/D"
    lDireccion.Caption = "S/D"
    lNombre.Caption = ""
    tGarantia.Text = "" 'FormatoCedula
    lGarantia.Caption = ""
    
    lvVenta.ListItems.Clear
    tEntrega.Text = ""
    cPago.Text = ""
    
    lVendedor.Caption = ""
    
    cEMailsT.ClearObjects
    
End Sub

Private Sub AccionValidar(ByVal sIDsRefinanciacion As String, ByVal resp As Byte, ByVal subresp As Byte)

Dim sOk As Boolean
Dim sFacturar As Boolean
    
    On Error GoTo errValidar1

    sFacturar = False: sOk = False
    
    Screen.MousePointer = 11
    
    If Trim(tCondicion.Tag) <> "" Then
        sOk = ValidoCondicion(tCondicion.Tag)
        If sOk Then
            Screen.MousePointer = 0
            If MsgBox("La solicitud ha sido validada y puede proceder a facturarla." & vbCrLf & _
                           "¿Desea facturarla ahora?", vbQuestion + vbYesNo, "Facturar Solicitud") = vbYes Then
                sFacturar = True
            End If
        End If
    End If
    
    Screen.MousePointer = 0
    
    If Not sOk Then
        MsgBox "La solicitud no cumple con las condiciones para su autorización.", vbInformation, "ATENCIÓN"
    Else
        If sFacturar Then
            'Controlo que los documentos de paga vencidas no esten para anular.
            If HayEnviosAProcesar Then
                ProcesoEnvios
                If MsgBox("Confirma realizar la emisión de las facturas. ", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then AccionFacturar sIDsRefinanciacion, resp, subresp
            Else
                AccionFacturar sIDsRefinanciacion, resp, subresp
            End If
        End If
    End If
    Exit Sub
    
errValidar1:
    clsGeneral.OcurrioError "Error: " & Trim(Err.Description)
End Sub

Private Sub AccionFacturar(ByVal idsDocRefinanciados As String, ByVal respEQV As Byte, ByVal subRespEQV As Byte)
   
Dim aMsgError As String
Dim aPlanFacturando As Long     'Plan que estoy facturando
Dim itmF As ListItem                  'Items a factuar
Dim aFTotal As Currency            'Monto Total de la Factura
Dim aFEntrega As Currency        'Monto Total de la Entrega
Dim aFCuota As Currency           'Monto de una Cuota
Dim aFIva As Currency

Dim aFEnvio As Currency
Dim auxEnviosPagos As String        'Lista de envios pagos con la factura
Dim auxEnviosFactura As String    'Lista de envios de la factura (no importa la F. de Pago)

Dim aListaDeImpresion  As String             'lista de facturas para imprimir 1234E1200V49:124E0S: (factura E $entega V $envio S/N Envio:)
Dim aSTR As String

    'Tranco botón para doble enter o doble clic en el botón.
    If bValidar.Tag = "" And Me.Enabled Then
        bValidar.Tag = "1"
        Me.Enabled = False
    Else
        Exit Sub
    End If
           
    On Error GoTo errorBT
    
    aListaDeImpresion = ""
    Screen.MousePointer = 11
    FechaDelServidor    'Saco la fecha del servidor
    
    fnc_ValidoCantidadesAEnviar
    
    ReDim arrCreditos(0)
    arrCreditos(0).idPlan = 0
    
    'Verifico que otro usuario no haya agarrado la solicitud y emitio el documento.
    Cons = "SELECT SolProceso FROM Solicitud WHERE SolCodigo = " & prmIDSolicitud
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux("SolProceso") = TipoResolucionSolicitud.Facturada Then
        Screen.MousePointer = 0
        MsgBox "Atención!!!" & vbCrLf & vbCrLf & "El proceso de la solicitud es facturada, verifique por favor.", vbExclamation, "Solicitud facturada"
        RsAux.Close
        Exit Sub
    End If
    RsAux.Close
    
    
    If cDireccion.ListIndex > -1 Then
        Cons = "SELECT DepNombre, LocNombre " & _
            "FROM Direccion INNER JOIN Calle ON DirCalle = CalCodigo " & _
            "INNER JOIN Localidad ON CalLocalidad = LocCodigo " & _
            "INNER JOIN Departamento ON LocDepartamento = DepCodigo " & _
            "WHERE DirCodigo = " & cDireccion.ItemData(cDireccion.ListIndex)
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            oCliente.Direccion.Departamento = Trim(RsAux("DepNombre"))
            oCliente.Direccion.Localidad = Trim(RsAux("LocNombre"))
        End If
        RsAux.Close
    End If
    
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    'Si tiene activado el pago de vencidas lo envío.
    'OJO lo hago antes ya que el sp consulta vencidas a lo mejor insertando los nuevo los tome.
    If chkVencidaTit.value = 1 Then cBase.Execute "EXEC prg_PagaVencidas " & gCliente & ", 1, " & paCodigoDeSucursal
    If Val(chkVencidaConyuge.Tag) > 0 And chkVencidaConyuge.value = 1 Then cBase.Execute "EXEC prg_PagaVencidas " & Val(chkVencidaConyuge.Tag) & ", 1, " & paCodigoDeSucursal
        
    For Each itmX In lvVenta.ListItems
        If Trim(itmX.SubItems(10)) = "" Then        'Si no fue facturado
            aPlanFacturando = PlanDeLaClave(itmX.Key)
            aFTotal = 0: aFEntrega = 0: aFCuota = 0: aFIva = 0
            auxEnviosFactura = ""
            
            cCofis = 0
            
            For Each itmF In lvVenta.ListItems
                If Trim(itmF.SubItems(10)) = "" Then        'Si no fue facturado
                    If aPlanFacturando = PlanDeLaClave(itmF.Key) Then
                        itmF.SubItems(10) = "SI"
                        aFTotal = aFTotal + CCur(itmF.SubItems(6))
                        aFIva = aFIva + (CCur(itmF.SubItems(6)) - (CCur(itmF.SubItems(6)) / (1 + (CCur(itmF.SubItems(3)) / 100))))
                        aFCuota = aFCuota + CCur(itmF.SubItems(5))
                        If Trim(itmF.SubItems(4)) <> "" Then aFEntrega = aFEntrega + CCur(itmF.SubItems(4))
                        
                        'Inserto los envios que se hicieron con la factura (no importa la forma de pago)
                        If Trim(itmF.SubItems(11)) <> "" Then
                            If InStr(auxEnviosFactura, Trim(itmF.SubItems(11)) & ",") = 0 Then auxEnviosFactura = auxEnviosFactura & Trim(itmF.SubItems(11)) & ","
                        End If
                        
                    End If
                End If
            Next
            If auxEnviosFactura <> "" Then auxEnviosFactura = Mid(auxEnviosFactura, 1, Len(auxEnviosFactura) - 1) 'le saco la coma del final
            
            'Veo si Hay Envios para los articulos a facturar---------------------------------------
            aFEnvio = 0: auxEnviosPagos = ""
            For Each itmF In lEnvio.ListItems
                If aPlanFacturando = PlanDeLaClave(itmF.Key) Then
                    auxEnviosPagos = Trim(itmF.SubItems(5))     'Lista de Envios a Apdatear con la factura
                    aFTotal = aFTotal + CCur(itmF.SubItems(4))
                    aFIva = aFIva + (CCur(itmF.SubItems(4)) - (CCur(itmF.SubItems(4)) / (1 + (CCur(itmF.SubItems(3)) / 100))))
                    aFEnvio = aFEnvio + CCur(itmF.SubItems(4))
                End If
            Next
            GraboDatosTablas aFTotal, aFIva, aFEntrega, aFCuota, aPlanFacturando, aFEnvio, auxEnviosFactura, auxEnviosPagos
            
            'Cargo lista de documentos a imprimir y documentos facturados
            'Formato de Cada Item: 12300E1000V0.00S
            If Trim(auxEnviosFactura) = "" Then aSTR = "N" Else aSTR = "S"
            aListaDeImpresion = aListaDeImpresion & aDocumentoFactura & "E" & aFEntrega & "V" & aFEnvio & aSTR & ":"
        End If
    Next
    
    'Actualizo la solicitud como Facturada-----------------------------------------------------
    Cons = "Update Solicitud Set SolProceso = " & TipoResolucionSolicitud.Facturada & " Where SolCodigo = " & prmIDSolicitud
    cBase.Execute Cons
    '-----------------------------------------------------------------------------------------------
    
    'Actualizo el servico con el id de la factura-----------------------------------------------------
    If aIDServicio <> 0 Then
        Cons = "Update Servicio Set SerDocumento= " & aDocumentoFactura & " Where SerCodigo= " & aIDServicio
        cBase.Execute Cons
    End If
    '-----------------------------------------------------------------------------------------------
    If idsDocRefinanciados <> "" Then
        cBase.Execute "EXEC prg_RelacionaRefinanciaciones '" & idsDocRefinanciados & "', " & aDocumentoFactura
    End If
    
    If respEQV <> 255 Then
        Cons = "INSERT INTO EnQueVino VALUES (" & aDocumentoFactura & ", " & respEQV & ", " & IIf(subRespEQV > 0, subRespEQV, "Null") & ")"
        cBase.Execute Cons
    End If
    
    
    cBase.CommitTrans    'FIN DE TRANSACCION-------------------------------------------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

    On Error GoTo errorFIN
    sFacturada = True
    
    'Si paga con cheques diferidos llamo al ingreso------------------------------------------------------------------
    If cPago.ItemData(cPago.ListIndex) = TipoPagoSolicitud.ChequeDiferido Then
        EjecutarApp prmPathApp & "\AltaCheques.exe", CStr("T 1|D " & aDocumentoRecibo)
    End If
    '-----------------------------------------------------------------------------------------------------------------------
    
    AccionImprimir aListaDeImpresion
    
    'Veo si hay instalaciones para llamar al programa, Solo lo hago para la ultima factura. -----------------
    If fnc_HayArticulosDeInstalaciones(aDocumentoFactura) Then
        EjecutarApp prmPathApp & "\Instalaciones.exe", "doc:" & CStr(aDocumentoFactura)
    End If
    
    NotificoCambioSignalR
    Unload Me
    Screen.MousePointer = 0
    Exit Sub

errorFIN:
    bValidar.Tag = ""
    Me.Enabled = True
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al restaurar el formulario.", Err.Description
    Exit Sub

errorBT:
    bValidar.Tag = ""
    Me.Enabled = True
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación."
    Exit Sub

errorET:
    Resume ErrorRoll

ErrorRoll:
    cBase.RollbackTrans
    bValidar.Tag = ""
    Me.Enabled = True
    Screen.MousePointer = 0
    If Trim(aMsgError) = "" Then aMsgError = "No se ha podido realizar la transacción. Reintente la operación."
    clsGeneral.OcurrioError aMsgError
    'Hay que desmarcar todos los articulos que marque como facturados
    For Each itmX In lvVenta.ListItems: itmX.SubItems(10) = "": Next
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'   Total:  Total de la Factura (icluido envios - si hay)
'   IVA: Iva de la factura (icluido envios - si hay)
'   Entrega: Valor de la entrega
'   Cuota: Valor de la cuota
'   Plan: Codigo de plan que voy a facturar
'   Envio: Valor Total de los envios para la factura (a recargar en entrega o primer cuota)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GraboDatosTablas(Total As Currency, iva As Currency, Entrega As Currency, Cuota As Currency, Plan As Long, Envio As Currency, ListaEnvios As String, EnviosFacturados As String)
Dim aDistancia As Integer       'Distancia entre cuota y cuota
Dim aVencimientoE As Integer 'Dias de vencimiento de la entrega
Dim aVencimientoC As Integer 'Dias de vencimiento de la cuota
Dim aCantidadCuotas As Integer 'Cantidad de cuotas del plan
    
    'Saco las caracteristicas del Plan---------------------------------------------------------------
    aDistancia = -1
    aVencimientoE = -1
    aVencimientoC = -1
    Cons = "Select * from TipoCuota Where TCuCodigo = " & Plan
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    aCantidadCuotas = RsAux!TCuCantidad
    If Not IsNull(RsAux!TCuDistancia) Then aDistancia = RsAux!TCuDistancia
    If Not IsNull(RsAux!TCuVencimientoE) Then aVencimientoE = RsAux!TCuVencimientoE
    If Not IsNull(RsAux!TCuVencimientoC) Then aVencimientoC = RsAux!TCuVencimientoC
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------
    
    'Inserto los valores en la tabla Documento--------------------------------------------------------
    Dim oFact As typCredito
    oFact = GraboDatosTDocumento(Total, iva)
    aDocumentoFactura = oFact.Credito.Codigo
        
    arrIdx = 0
    If arrCreditos(0).idPlan <> 0 Then
        arrIdx = UBound(arrCreditos) + 1
        ReDim Preserve arrCreditos(arrIdx)
    End If
    With arrCreditos(arrIdx)
        .idPlan = Plan
        Set .CAE = oFact.CAE
        Set .Credito = oFact.Credito
    End With
    
    'Inserto los valores en la tabla Renglon------------------------------------------------------------
    GraboDatosTRenglonDocumento aDocumentoFactura, Plan, Envio
    Dim oRenglon As clsDocumentoRenglon
    Dim sQy As String
    Dim rsR As rdoResultset
    
    sQy = "SELECT AEsID, ArtID, ArtCodigo, ISNULL(AEsNombre, ArtNombre) ArtNombre, ArtTipo, IVAPorcentaje, RenCantidad, RenIVA, RenPrecio, RenDescripcion, RenARetirar " & _
        "FROM Renglon INNER JOIN Articulo ON RenArticulo = ArtID INNER JOIN ArticuloFacturacion ON AFAArticulo = ArtID " & _
        "INNER JOIN TipoIVA ON IvaCodigo = AFAIva " & _
        "LEFT OUTER JOIN ArticuloEspecifico ON AEsDocumento = RenDocumento AND AEsArticulo = RenArticulo AND AEsTipoDocumento = 1 " & _
        "WHERE RenDocumento = " & aDocumentoFactura
    Set rsR = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    Do While Not rsR.EOF
        Set oRenglon = New clsDocumentoRenglon
        oFact.Credito.Renglones.Add oRenglon
        With oRenglon
            .Articulo.ID = rsR("ArtID")
            If Not IsNull(rsR("AEsID")) Then
                .Articulo.IDEspecifico = rsR("AEsID")
            End If
            .Articulo.Codigo = rsR("ArtCodigo")
            .Articulo.Nombre = Trim(rsR("ArtNombre"))
            .Articulo.TipoIVA.Porcentaje = rsR("IVAPorcentaje")
            .Articulo.TipoArticulo = rsR("ArtTipo")
            .Cantidad = rsR("RenCantidad")
            .iva = rsR("RenIVA")
            .Precio = rsR("RenPrecio")
            .CantidadARetirar = rsR("RenARetirar")
            If Not IsNull(rsR("RenDescripcion")) Then .Descripcion = rsR("RenDescripcion") Else .Descripcion = ""
        End With
        rsR.MoveNext
    Loop
    rsR.Close
    
    'Si tiene prendida la señal de entrega en local.
    If chkRetiraAqui.value And chkRetiraAqui.Visible Then
        'prg_EntregaCliente_EntregoPorVenta @documento int, @local smallint
        'cBase.Execute "prg_EntregaCliente_EntregoPorVenta " & aDocumentoFactura & ", " & paCodigoDeSucursal
        cBase.Execute "prg_EntregaMercaderia_EntregoArticulo 9999, 9898, 0, 0, " & aDocumentoFactura & ", " & paCodigoDeSucursal & ", 0, 0"
    End If
    
    'Inserto los datos en tabla Credito---------------------------------------------------------------
    aMaxDocumento = GraboDatosTCredito(aDocumentoFactura, Plan, aVencimientoE, aVencimientoC, aCantidadCuotas, aDistancia, Total, Entrega, Cuota, Envio)
    
    'Actualizo RenglonSolicitud-----------------------------------------------------------------------
    Cons = " Update RenglonSolicitud Set RSoDocumento = " & aDocumentoFactura _
            & " Where RSoSolicitud = " & prmIDSolicitud _
            & " And RSoTipoCuota = " & Plan
    cBase.Execute Cons
    '-----------------------------------------------------------------------------------------------------
        
    'Cargo Cuotas Almacenadas
    GraboDatosTCreditoCuota aMaxDocumento, Envio, Cuota, Entrega, aCantidadCuotas, aVencimientoC, aVencimientoE, aDistancia
        
    'Si paga la Entrega o La primera Cuota hay que Hacer un Recibo--- Si Hay Envio También-------------------------
    If cPago.ItemData(cPago.ListIndex) = TipoPagoSolicitud.Efectivo Then
        If aVencimientoE = 0 Or aVencimientoC = 0 Then
            If aVencimientoE = 0 Then
                GraboPagoRecibo aDocumentoFactura, CLng(cMoneda.ItemData(cMoneda.ListIndex)), 0, aCantidadCuotas, Entrega + Envio
            Else
                GraboPagoRecibo aDocumentoFactura, CLng(cMoneda.ItemData(cMoneda.ListIndex)), 1, aCantidadCuotas, Cuota + Envio
            End If
            
        Else
            If Envio > 0 Then  'HAY ENVIOS
                'Si hay Envio inserto el pago Parcial del Envio a la primer cuota
                If aVencimientoE <> -1 Then 'Hay entrega, pago parcial (envio) de entrega
                    GraboPagoRecibo aDocumentoFactura, CLng(cMoneda.ItemData(cMoneda.ListIndex)), 0, aCantidadCuotas, Envio
                Else                                     'Hay cuota, pago parcial (envio) de cuota
                    GraboPagoRecibo aDocumentoFactura, CLng(cMoneda.ItemData(cMoneda.ListIndex)), 1, aCantidadCuotas, Envio
                End If
            End If
        End If
    Else
        'Pago con cheque diferido---------
        If aVencimientoE <> -1 Then 'Hay entrega
            GraboPagoRecibo aDocumentoFactura, CLng(cMoneda.ItemData(cMoneda.ListIndex)), 0, aCantidadCuotas, Total, True, Cuota, Envio, Entrega
        Else
            GraboPagoRecibo aDocumentoFactura, CLng(cMoneda.ItemData(cMoneda.ListIndex)), 1, aCantidadCuotas, Total, True, Cuota, Envio
        End If
    End If
    '-----------------------------------------------------------------------------------------------------
    'Actualizo los envios con los codigos de factura
    If ListaEnvios <> "" Then GraboDatosTEnvio aDocumentoFactura, ListaEnvios, EnviosFacturados
    If gSucesoUsr <> 0 Then
        clsGeneral.RegistroSuceso cBase, gFechaServidor, prmSuc_FacturaArticuloNOHabilitado, paCodigoDeTerminal, gSucesoUsr, aDocumentoFactura, _
                            Descripcion:=Trim(gSucesoDesc), Defensa:=Trim(gSucesoDef)
    End If
    If spv_Usuario <> 0 Then
        clsGeneral.RegistroSucesoAutorizado cBase, gFechaServidor, prmSuc_ModificacionDePrecios, paCodigoDeTerminal, spv_Usuario, aDocumentoFactura, _
                            Descripcion:="La vigencia de precios es mayor a la de Solicitud.", Defensa:=Trim(spv_Defensa), idAutoriza:=spv_Autoriza
    End If
    
    cBase.Execute "EXEC prg_PosInsertoDocumentosATickets '" & aDocumentoFactura & "', " & oCnfgCredito.ImpresoraTickets
    If EmitirCFE(oFact.Credito, oFact.CAE) <> "" Then RsAux.Edit
    
    
    
End Function

Private Function GraboDatosTDocumento(MTotal As Currency, MIva As Currency) As typCredito
    
    Dim aZonaDir As Long
    'Saco la zona de la direccion q factura-----------------------------
    aZonaDir = 0
    If cDireccion.ListIndex <> -1 Then
        aZonaDir = cDireccion.ItemData(cDireccion.ListIndex)
        aZonaDir = BuscoZonaDireccion(aZonaDir)
    End If
    '-----------------------------------------------------------------------

    Dim mFechaRetira As String
'    mFechaRetira = tFRetiro.Text
    'zGet_StringRetira  "", mFechaRetira, porPlan:=idPlan
    
    mFechaRetira = CDate(tFRetiro.Text)
    Dim dFRet As Date: dFRet = ObtenerMayorFechaRetiro
    If dFRet > DateSerial(2000, 1, 1) Then
        If CDate(mFechaRetira) > dFRet Then
            If Abs(DateDiff("d", CDate(mFechaRetira), dFRet)) > 59 Then
                mFechaRetira = mFechaRetira & " " & "01:59:00"
            Else
                mFechaRetira = mFechaRetira & " " & "01:" & Format(Abs(DateDiff("d", dFRet, CDate(mFechaRetira))), "00") & ":00"
            End If
        Else
            mFechaRetira = mFechaRetira & " " & "01:00:00"
        End If
    End If
        
    'Pido el numero de documento--------------------------

    Dim tipoCAE As Byte
    tipoCAE = IIf(oCliente.RUT <> "", 111, 101)
    Dim CAE As New clsCAEDocumento
    If Val(prmEFacturaProductivo) = 0 Then
        aTexto = NumeroDocumento(paDCredito)
        aSerie = Mid(aTexto, 1, 1)
        aNumero = CLng(Mid(aTexto, 2, Len(aTexto)))
        With CAE
            .Desde = 1
            .Hasta = 9999999
            .Serie = aSerie
            .Numero = aNumero
            .IdDGI = "901411"
            .TipoCFE = tipoCAE
            .Vencimiento = "31/12/" & CStr(Year(Date))
        End With
    Else
        Dim caeG As New clsCAEGenerador
        Set CAE = caeG.ObtenerNumeroCAEDocumento(cBase, tipoCAE, paCodigoDGI)
        Set caeG = Nothing
    End If
    
    'Inserto los datos en la tabla documento--------------------------------------------------------
    Cons = "INSERT INTO Documento" _
            & " (DocFecha, DocTipo, DocSerie, DocNumero, DocCliente, DocMoneda, DocTotal, DocIVA, DocAnulado, DocSucursal, DocUsuario, DocFModificacion, DocComentario, DocFRetira, DocPendiente, DocVendedor, DocZona, DocCofis) " _
            & " VALUES ('" & Format(gFechaServidor, sqlFormatoFH) & "', " _
            & TipoDocumento.Credito & ", " _
            & "'" & CAE.Serie & "', " _
            & CAE.Numero & ", " _
            & gCliente & ", " _
            & cMoneda.ItemData(cMoneda.ListIndex) & ", " _
            & MTotal & ", " _
            & MIva & ", " _
            & "0, " _
            & paCodigoDeSucursal & ", " _
            & tUsuario.Tag & ", " _
            & "'" & Format(gFechaServidor, sqlFormatoFH) & "', "
        
    If Trim(tComentario.Text) = "" Then Cons = Cons & "NULL, " Else:  Cons = Cons & "'" & Trim(tComentario.Text) & "', "
    
    Cons = Cons & "'" & Format(mFechaRetira, sqlFormatoFH) & "', "
    
    If cPendiente.ListIndex = -1 Then Cons = Cons & "NULL, " Else:  Cons = Cons & cPendiente.ItemData(cPendiente.ListIndex) & ", "
    
    If Trim(lVendedor.Tag) = "" Then Cons = Cons & "NULL," Else:  Cons = Cons & Trim(lVendedor.Tag) & ", "
    
    If aZonaDir = 0 Then Cons = Cons & "Null, " Else Cons = Cons & aZonaDir & ", "
    
    Cons = Cons & "Null"
    Cons = Cons & ")"
        
    cBase.Execute (Cons)
        
    Dim idDoc As Long
    'Saco el Maximo codigo de la tabla documento----------------------------------------------------
    Cons = "SELECT MAX(DocCodigo) From Documento" _
            & " WHERE DocTipo = " & TipoDocumento.Credito _
            & " AND DocSerie = '" & CAE.Serie & "'" _
            & " AND DocNumero = " & CAE.Numero _
            & " AND DocCliente = " & gCliente
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    idDoc = RsAux(0)
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------
    
    
    Dim doc As New clsDocumentoCGSA
    With doc
        Set .Cliente = oCliente
        .Codigo = idDoc
        .Emision = gFechaServidor
        .Tipo = TD_Credito
        .Numero = CAE.Numero
        .Serie = CAE.Serie
        .Moneda.Codigo = cMoneda.ItemData(cMoneda.ListIndex)
        .Total = CCur(MTotal)
        .iva = CCur(MIva)
        .sucursal = paCodigoDeSucursal
        .Digitador = tUsuario.Tag
        .Comentario = tComentario.Text
        .Zona = aZonaDir
        .FechaRetira = CDate(mFechaRetira)
        If cPendiente.ListIndex > -1 Then .Pendiente = cPendiente.ItemData(cPendiente.ListIndex)
        .Vendedor = Val(lVendedor.Tag)
    End With
    
    Set GraboDatosTDocumento.Credito = doc
    Set GraboDatosTDocumento.CAE = CAE
    
    
End Function

Private Function ObtenerIDEspecifico(ByVal clave As String) As Long
Dim arrvalores() As String
    ObtenerIDEspecifico = 0
    arrvalores = Split(clave, "|")
    If UBound(arrvalores) >= 0 Then
        If Val(arrvalores(0)) > 0 Then ObtenerIDEspecifico = Val(arrvalores(0))
    End If
End Function

Private Sub GraboDatosTRenglonDocumento(factura As Long, Plan As Long, MEnvio As Currency)
Dim itmA As ListItem
Dim rsEnv As rdoResultset
Dim aUnitario As Currency, aARetirar As Currency
Dim bEsFlete As Boolean

Dim acContado As Currency, aCofisX1 As Currency
Dim aUnitarioIVA As Currency

    If Trim(aFletes) = "" Then aFletes = CargoArticulosDeFlete
    
    'Artículos de la lista----------------------------------------------------------------------------------------------------
    For Each itmA In lvVenta.ListItems
        If Plan = PlanDeLaClave(itmA.Key) Then
            bEsFlete = False
            aUnitario = CCur(itmA.SubItems(6)) / CCur(itmA.SubItems(1))
            aUnitarioIVA = CCur(Format(aUnitario - (aUnitario / (1 + (CCur(itmA.SubItems(3)) / 100))), "0.00"))
            
            'Puede haber articulos que son de flete....hay que controlar Cant A Retirar y Stock
            If InStr(aFletes, ArticuloDeLaClave(itmA.Key) & ",") <> 0 Then bEsFlete = True
    
            'Hay que sacar cuantos de éstos artículos están para enviar ---> para descontar cant. ARetirar
            aARetirar = 0
            If aIDServicio = 0 Then
                If Not bEsFlete Then
                    aARetirar = CCur(itmA.SubItems(1))
                    If Trim(itmA.SubItems(11)) <> "" Then
                        Cons = " Select Sum(REvAEntregar) from RenglonEnvio " _
                                & " Where REvEnvio IN (" & itmA.SubItems(11) & ")" _
                                & " And REvArticulo = " & ArticuloDeLaClave(itmA.Key)
                        Set rsEnv = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                        If Not rsEnv.EOF Then aARetirar = aARetirar - rsEnv(0)
                        rsEnv.Close
                    End If
                End If
            End If
            
            aCofisX1 = 0
            
            Cons = "INSERT INTO Renglon (RenDocumento, RenArticulo, RenCantidad, RenPrecio, RenIVA, RenARetirar, RenDescripcion, RenCofis)" _
                    & " VALUES (" _
                    & factura & ", " _
                    & ArticuloDeLaClave(itmA.Key) & ", " _
                    & itmA.SubItems(1) & ", " _
                    & aUnitario & ", " _
                    & aUnitarioIVA & "," _
                    & aARetirar & ", "
            If Trim(itmA.SubItems(12)) <> "" Then Cons = Cons & "'" & Trim(itmA.SubItems(12)) & "', " Else Cons = Cons & " Null, "
            
            Cons = Cons & " Null)"
            cBase.Execute (Cons)
            
            '14/8/2008 asigno el documento al artículo específico.
            If ObtenerIDEspecifico(itmA.SubItems(16)) > 0 Then
                Cons = " Update ArticuloEspecifico " & _
                        " Set AEsTipoDocumento = 1, AEsDocumento = " & factura & _
                        ", AEsModificado = '" & Format(gFechaServidor, "yyyy/mm/dd hh:nn:ss") & "'" & _
                        " Where AEsID = " & ObtenerIDEspecifico(itmA.SubItems(16))
                cBase.Execute Cons
            End If
            
            'Bajo Mercadería del STOCK   - Articulo, Cant. ARetirar(x Cliente), Cant. ADomicilio, ....
            If Not bEsFlete And aIDServicio = 0 Then
                If Not fnc_EsDelTipoServicio(ArticuloDeLaClave(itmA.Key)) Then
                    MarcoStockVenta CLng(tUsuario.Tag), ArticuloDeLaClave(itmA.Key), aARetirar, CCur(itmA.SubItems(1)) - aARetirar, _
                                              0, TipoDocumento.Credito, factura, paCodigoDeSucursal
                End If
            End If
        End If
    Next
    
    'Artículos de envíos----------------------------------------------------------------------------------------------------
    Dim RsXX As rdoResultset
    For Each itmA In lEnvio.ListItems
        If Plan = PlanDeLaClave(itmA.Key) Then
            
            aUnitario = CCur(itmA.SubItems(4)) / CCur(itmA.SubItems(1))
            aUnitarioIVA = Format(aUnitario - (aUnitario / (1 + (CCur(itmA.SubItems(3)) / 100))), "0.00")
            
            aCofisX1 = 0
            
            Cons = " Select * from Renglon " _
                    & " Where RenDocumento = " & factura _
                    & " And RenArticulo = " & ArticuloDeLaClave(itmA.Key)
            Set RsXX = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If RsXX.EOF Then
                RsXX.AddNew
                RsXX!REnDocumento = factura
                RsXX!RenArticulo = ArticuloDeLaClave(itmA.Key)
                RsXX!RenCantidad = CCur(itmA.SubItems(1))
                RsXX!RenPrecio = aUnitario
                RsXX!RenIVA = aUnitarioIVA
                RsXX!RenARetirar = 0
                RsXX.Update
            
            Else
                RsXX.Edit
                RsXX!RenCantidad = RsXX!RenCantidad + CCur(itmA.SubItems(1))
                
                aUnitario = (RsXX!RenPrecio + aUnitario) / 2
                aUnitarioIVA = (RsXX!RenIVA + aUnitarioIVA) / 2
                If Not IsNull(RsXX!RenCofis) Then aCofisX1 = (RsXX!RenCofis + aCofisX1) / 2
                RsXX!RenPrecio = aUnitario
                RsXX!RenIVA = aUnitarioIVA
                RsXX.Update
            End If
            RsXX.Close
            
        End If
    Next
    
End Sub

Private Function GraboDatosTCredito(factura As Long, Plan As Long, VencimientoE As Integer, VencimientoC As Integer, CanCuotas As Integer, _
                        Distancia As Integer, MTotal As Currency, MEntrega As Currency, MCuota As Currency, MEnvio As Currency) As Long

Dim aVaDe As String, aProxVto As String, aCumplimineto As String, aFPago As String
Dim aSaldo As Currency
    
    'Saco los datos de Pago de Entrega y Cuotas--------------------------------------------------------------------------
    If VencimientoE <> -1 Then    'Hay Entrega -------------------------------!!!!!!!
        If VencimientoE = 0 Then     'La paga hoy
            aVaDe = "E"
            aFPago = Format(gFechaServidor, sqlFormatoFH)
            aSaldo = MTotal - MEntrega - MEnvio
            aProxVto = Format(ProximoVencimiento(gFechaServidor, gFechaServidor, VencimientoC), sqlFormatoFH) 'Prox. Vto es Cuota
            aCumplimineto = Trim(FormatoCumplimiento(CanCuotas + 1, gFechaServidor, ""))  '
        Else
            aVaDe = ""
            If MEnvio = 0 Then aFPago = "" Else: aFPago = Format(gFechaServidor, sqlFormatoFH)
            aSaldo = MTotal - MEnvio
            aProxVto = Format(gFechaServidor + VencimientoE, sqlFormatoFH)
            aCumplimineto = Trim(FormatoCumplimiento(CanCuotas + 1, gFechaServidor, "F"))
        End If
    Else                                    'No hay Entrega-------------------------------!!!!!!!!
        If VencimientoC = 0 Then  'La paga hoy
            aVaDe = "1"
            aFPago = Format(gFechaServidor, sqlFormatoFH)
            aSaldo = MTotal - MCuota - MEnvio
            aProxVto = Format(ProximoVencimiento(gFechaServidor, gFechaServidor, Distancia), sqlFormatoFH)
            aCumplimineto = Trim(FormatoCumplimiento(CanCuotas, gFechaServidor, ""))
        Else
            aVaDe = ""
            If MEnvio = 0 Then aFPago = "" Else: aFPago = Format(gFechaServidor, sqlFormatoFH)
            aSaldo = MTotal - MEnvio
            aProxVto = Format(gFechaServidor + VencimientoC, sqlFormatoFH)
            aCumplimineto = Trim(FormatoCumplimiento(CanCuotas, gFechaServidor, "F"))
        End If
    End If
    
    Dim idConforme As Long
    Dim sQy As String
    Dim rsConf As rdoResultset
    If cPago.ItemData(cPago.ListIndex) = TipoPagoSolicitud.Efectivo Then
        sQy = "SELECT * FROM Contador WHERE ConDocumento = 'Conforme(" & paNombreSucursal & ")'"
        Set rsConf = cBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
        If rsConf.EOF Then
            rsConf.AddNew
            idConforme = 1
            rsConf("ConDocumento") = "Conforme(" & paNombreSucursal & ")"
            rsConf("ConSerie") = "A"
        Else
            rsConf.Edit
            idConforme = rsConf("ConValor") + 1
        End If
        rsConf("ConValor") = idConforme
        rsConf.Update
        rsConf.Close
    End If
    
    'Inserto los datos en tabla Credito---------------------------------------------------------------
    Cons = " Insert Into Credito " _
        & " (CreFactura, CreTipoCuota, CreVaCuota, CreDeCuota, CreUltimoPago, CreSaldoFactura, CreProximoVto, " _
        & "  CreCumplimiento, CrePuntaje, CreCliente, CreGarantia, CreTipo, CreValorCuota, CreFormaPago, CreConforme)" _
        & " Values (" _
        & factura & ", " _
        & Plan & ", "
        
    If cPago.ItemData(cPago.ListIndex) = TipoPagoSolicitud.Efectivo Then    '-----------------------------------------
        'Pago EFECTIVO
        If Trim(aVaDe) <> "" Then Cons = Cons & "'" & aVaDe & "', " Else: Cons = Cons & "NULL, "
        Cons = Cons & "'" & CanCuotas & "', "
    
        If Trim(aFPago) <> "" Then Cons = Cons & "'" & aFPago & "', " Else: Cons = Cons & "NULL, "
        
        Cons = Cons & aSaldo & ", " _
        & "'" & aProxVto & "', " _
        & "'" & aCumplimineto & "', " _
        & "Null, "
    
    Else
        'Pago CHEQUE DIFERIDO
        Cons = Cons & "'" & CanCuotas & "', " _
                           & "'" & CanCuotas & "', " _
                           & "'" & Format(gFechaServidor, sqlFormatoFH) & "', " _
                           & "0, " _
                           & "Null, " _
                           & "'" & FormatoCumplimiento(Len(aCumplimineto), gFechaServidor, "P") & "', " _
                           & FormatoPuntaje(FormatoCumplimiento(Len(aCumplimineto), gFechaServidor, "P")) & ", "
    End If
    '--------------------------------------------

    Cons = Cons & gCliente & ", "
    
    If Trim(tGarantia.Tag) <> 0 Then Cons = Cons & tGarantia.Tag & ", " Else: Cons = Cons & "NULL, "

    'Tipo de Credito
    Cons = Cons & TipoCredito.Normal & ", "
    
    If MCuota > 0 Then Cons = Cons & MCuota & ", " Else: Cons = Cons & "Null, "
    
    Cons = Cons & cPago.ItemData(cPago.ListIndex) & ", " & IIf(idConforme > 0, idConforme, "NULL") & ")"        'Forma de Pago y # de conforme.
    cBase.Execute (Cons)
    '-----------------------------------------------------------------------------------------------------
    
    'Saco el Maximo codigo de la tabla Credito----------------------------------------------------
    Cons = "SELECT MAX(CreCodigo) From Credito" _
            & " WHERE CreFactura = " & factura
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    GraboDatosTCredito = RsAux(0)
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------

End Function

Private Sub GraboDatosTCreditoCuota(Credito As Long, ByVal MEnvio As Currency, MCuota As Currency, MEntrega As Currency, CanCuotas As Integer, _
                VencimientoC As Integer, VencimientoE As Integer, DistanciaC As Integer)

Dim aDesde As Integer
Dim aVencAnterior As Date
Dim aPrimerVencimiento As Date    'Si es diferido xq no es la misma fecha que la factura

Dim aCCuVencimiento As String, aCCuUltimoPago As String, aCCuSaldo As Currency, aCCuValor As Currency

    aDesde = 1
    If VencimientoE <> -1 Then aDesde = 0
    For i = aDesde To CanCuotas
        
        If i = 0 Then       'ENTREGA !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            aCCuValor = MEntrega + MEnvio
            If VencimientoE = 0 Then     'La paga hoy
                aCCuVencimiento = Format(gFechaServidor, sqlFormatoFH)
                aCCuUltimoPago = Format(gFechaServidor, sqlFormatoFH)
                aCCuSaldo = 0
            Else
                aCCuVencimiento = Format(gFechaServidor + VencimientoE, sqlFormatoFH)
                If MEnvio = 0 Then aCCuUltimoPago = "" Else aCCuUltimoPago = Format(gFechaServidor, sqlFormatoFH)
                aCCuSaldo = MEntrega
            End If
            'Como Hay entrega y ya procese el Monto del Envio --> lo pongo en 0 para q' no influya en las cuotas
            MEnvio = 0
            
        Else                    'CUOTAS !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            If i = 1 Then           'La primera Cuota
                aCCuValor = MCuota + MEnvio
                If VencimientoC = 0 Then     'La paga hoy - CCuVencimiento, CCuUltimoPago, CCuSaldo, CCuMora
                    aCCuVencimiento = Format(gFechaServidor, sqlFormatoFH)
                    aCCuUltimoPago = Format(gFechaServidor, sqlFormatoFH)
                    aCCuSaldo = 0
                    
                    aVencAnterior = gFechaServidor
                    aPrimerVencimiento = gFechaServidor
                Else
                    aCCuVencimiento = Format(ProximoVencimiento(gFechaServidor, gFechaServidor, VencimientoC), sqlFormatoFH)
                    
                    aCCuSaldo = MCuota
                    If MEnvio = 0 Then aCCuUltimoPago = "" Else: aCCuUltimoPago = Format(gFechaServidor, sqlFormatoFH)
                    
                    aVencAnterior = ProximoVencimiento(gFechaServidor, gFechaServidor, VencimientoC)
                    aPrimerVencimiento = aVencAnterior
                End If
                
            Else        'Las cuotas restantes
                aCCuValor = MCuota
                aCCuVencimiento = Format(ProximoVencimiento(aPrimerVencimiento, aVencAnterior, DistanciaC), sqlFormatoFH)
                aCCuUltimoPago = ""
                aVencAnterior = ProximoVencimiento(aPrimerVencimiento, aVencAnterior, DistanciaC)
                aCCuSaldo = MCuota
            End If
        End If
        
        Cons = "Insert into CreditoCuota (CCuCredito, CCuNumero, CCuValor, CCuVencimiento, CCuUltimoPago, CCuSaldo, CCuMora, CCuMoraACuenta)" _
                & " Values (" _
                & Credito & ", " _
                & i & ", " _
                & aCCuValor & ", " _
                & "'" & aCCuVencimiento & "', "
                
        'Voe cual es la forma de pago Efectivo o Ch. Diferido
        If cPago.ItemData(cPago.ListIndex) = TipoPagoSolicitud.Efectivo Then
            If Trim(aCCuUltimoPago) = "" Then
                Cons = Cons & "Null, "
            Else
                Cons = Cons & "'" & Trim(aCCuUltimoPago) & "', "
            End If
            
            Cons = Cons & aCCuSaldo & ", 0, "     'Saldo Y Mora
        Else
            'Paga con cheque diferido ---> Saldo en CERO y ultimo pago HOY (Supuesta Fecha Librado del CH.)
            Cons = Cons & "'" & Format(gFechaServidor, sqlFormatoFH) & "', 0, 0, "
        End If
        
        Cons = Cons & " 0)"     'Mora a Cuenta
        
        cBase.Execute Cons
        
    Next i
    
End Sub

Private Sub GraboDatosTEnvio(factura As Long, Envios As String, EnviosFacturados As String)

    'Actualiza los Envios con la factura pasada como parámetro
    Dim auxEnvios As String, auxEnviosF As String
    Dim aEnvio As Long, aHabilitado As Integer
    
    If cPendiente.ListIndex <> -1 Then aHabilitado = 0 Else aHabilitado = 1
    
    auxEnvios = Envios
    auxEnviosF = Trim(EnviosFacturados) & ","
    
    Do While auxEnvios <> ""
    
        If InStr(1, auxEnvios, ",") > 0 Then
            aEnvio = CLng(Left(auxEnvios, InStr(1, auxEnvios, ",") - 1))
            auxEnvios = Trim(Mid(auxEnvios, InStr(1, auxEnvios, ",") + 1, Len(auxEnvios)))
        Else
            aEnvio = CLng(auxEnvios)
            auxEnvios = ""
        End If
        
        If InStr(auxEnviosF, aEnvio & ",") <> 0 Then
            Cons = "Update Envio " _
                   & " Set EnvDocumento = " & factura & "," _
                   & " EnvDocumentoFactura = " & factura & ", " _
                   & " EnvUsuario = " & Val(tUsuario.Tag) _
                   & " Where EnvCodigo = " & aEnvio
        Else
            Cons = "Update Envio " _
                   & " Set EnvDocumento = " & factura & ", " _
                   & " EnvUsuario = " & Val(tUsuario.Tag) _
                   & " Where EnvCodigo = " & aEnvio
        End If
        cBase.Execute (Cons)
    Loop

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------
'   Parametros:
'       Factura:    Codigo del documento para el que se hace el pago
'       Moneda:    Codigo de Moneda del documento
'       Cuota:       Nro de cuota que se paga.
'       De:            Nro total de cuotas
'       Monto:      Valor total por el que se hace el recibo
'
'   Valores Opcionales:     (Para pagos con cheques diferidos)
'       PagoDiferido:  Indica si se paga con Ch. Diferido
'       ValorCuota:     Monto de cada cuota
'       ValorEnvio:     Monto del envio - para recargar a la primer cuota o Entrega
'       ValorEntrega: Monto de la Entrega
'----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GraboPagoRecibo(factura As Long, Moneda As Long, Cuota As Integer, De As Integer, Monto As Currency, _
            Optional PagoDiferido As Boolean = False, Optional ValorCuota As Currency = 0, Optional ValorEnvio As Currency = 0, _
            Optional ValorEntrega As Currency = 0)

    'Pido el Numero de Documento para hacer el RECIBO-------------
    aTexto = NumeroDocumento(paDRecibo)
    aSerie = Mid(aTexto, 1, 1)
    aNumero = CLng(Trim(Mid(aTexto, 2, Len(aTexto))))
    
    'Inserto campos en la tabla documento---------------------------------------------------------------------
    Cons = "INSERT INTO Documento (DocFecha, DocTipo, DocSerie, DocNumero, DocCliente, DocMoneda, DocTotal, DocIVA, DocAnulado, DocSucursal, DocUsuario, DocFModificacion) " _
            & "VALUES (" _
            & "'" & Format(gFechaServidor, sqlFormatoFH) & "', " _
            & TipoDocumento.ReciboDePago & ", " _
            & "'" & aSerie & "', " _
            & aNumero & ", " _
            & gCliente & ", " _
            & Moneda & ", " _
            & Monto & ", " _
            & "0, " _
            & "0, " _
            & paCodigoDeSucursal & ", " _
            & tUsuario.Tag & ", " _
            & "'" & Format(gFechaServidor, sqlFormatoFH) & "')"
    
    cBase.Execute (Cons)
    '---------------------------------------------------------------------------------------------------------------------
    'Saco el Numero de Recibo de Pago
    Cons = "SELECT MAX(DocCodigo) From Documento" _
            & " WHERE DocTipo = " & TipoDocumento.ReciboDePago _
            & " AND DocSerie = '" & aSerie & "'" _
            & " AND DocNumero = " & aNumero
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    aDocumentoRecibo = RsAux(0)
    RsAux.Close
    
    cBase.Execute "EXEC prg_PosInsertoDocumentosATickets '" & aDocumentoRecibo & "', " & oCnfgRecibo.ImpresoraTickets
       
    
    'Armo las Relaciones en la tabla DocumentoPago  -------------------------------------------------------------
    Cons = "Select * from DocumentoPago " & _
                " Where DPaDocASaldar = " & factura & _
                " And DPaDocQSalda = " & aDocumentoRecibo
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
    If Not PagoDiferido Then
        
        RsAux.AddNew
        RsAux!DPaDocASaldar = factura
        RsAux!DPaDocQSalda = aDocumentoRecibo
        RsAux!DPaCuota = Cuota
        RsAux!DPaDe = De
        RsAux!DPaAmortizacion = Monto
        RsAux.Update

    Else
        'Pago con Cheques --> el valor del envio se recarga a la 1 cta o entrega
        Dim ValorDePago As Currency
        For i = Cuota To De
            
            If i = Cuota Then
                If ValorEntrega <> 0 Then ValorDePago = ValorEntrega Else ValorDePago = ValorCuota
                ValorDePago = ValorDePago + ValorEnvio
            Else
                ValorDePago = ValorCuota
            End If
            
            RsAux.AddNew
            RsAux!DPaDocASaldar = factura
            RsAux!DPaDocQSalda = aDocumentoRecibo
            RsAux!DPaCuota = i
            RsAux!DPaDe = De
            RsAux!DPaAmortizacion = ValorDePago
            RsAux.Update
            
        Next i
    End If
    
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------
    
End Sub

Private Function ValidoCondicion(sTexto As String) As Boolean

On Error GoTo errValidar

Dim sAuxiliar As String
Dim sCondicion As String
Dim sRetorno As String

Dim sValor1 As String
Dim sValor2 As String
Dim rsVal As rdoResultset
    
    ValidoCondicion = True
    gTextoCondicion = sTexto
    Do While gTextoCondicion <> ""
        sAuxiliar = SacoCondicionDelTexto
        
        sCondicion = UCase(Mid(sAuxiliar, 1, 3))                'Las primeras 3 son la Condicion
        sAuxiliar = Trim(Mid(sAuxiliar, 4, Len(sAuxiliar)))    'El resto son los valores
        
        Select Case sCondicion
            Case "RSU"      'Recibo de Sueldo - %--------------------------------------------------------------------------------------
                sValor1 = Trim(Mid(sAuxiliar, InStr(sAuxiliar, "M") + 1, Len(sAuxiliar)))   'Saco el sueldo
                'De primera no controlo el valor del sueldo
                Cons = "Select * from Empleo, TipoIngreso" _
                        & " Where EmpCliente = " & gCliente _
                        & " And EmpMoneda = " & Mid(sAuxiliar, 1, InStr(sAuxiliar, "M") - 1) _
                        & " And EmpExhibido >= '" & Format(Date - paVaToleranciaDiasExh, sqlFormatoFH) & "'" _
                        & " And EmpTipoIngreso = TInCodigo"
                
                Set rsVal = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                ValidoCondicion = False
                Do While Not rsVal.EOF
                    If Not IsNull(rsVal!EmpIngreso) And Not IsNull(rsVal!TInPorcentaje) Then
                        '(Ingreso * Porc.) /100 >= Ingreso - Porc tolerancia
                        If ((CCur(rsVal!EmpIngreso) * CCur(rsVal!TInPorcentaje)) / 100) >= (CCur(sValor1) - CCur(sValor1) * (paVaToleranciaMonedaPorc / 100)) Then
                            ValidoCondicion = True
                            Exit Do
                        End If
                    End If
                    rsVal.MoveNext
                Loop
                rsVal.Close
                '----------------------------------------------------------------------------------------------------------------------------
                
            Case "CPL"      'Cambio de Plan -------------------------------------------------------------------------------------------
                'Veo si es con Entrega
                sValor1 = ""
                If InStr(sAuxiliar, "E") <> 0 Then
                    sValor1 = Trim(Mid(sAuxiliar, InStr(sAuxiliar, "E") + 1, Len(sAuxiliar)))       'Entrega
                    sAuxiliar = Mid(sAuxiliar, 1, InStr(sAuxiliar, "E") - 1)        'Codigo de Plan
                End If
                For Each itmX In lvVenta.ListItems
                    If Mid(itmX.Key, InStr(itmX.Key, "P") + 1, InStr(itmX.Key, "A") - InStr(itmX.Key, "P") - 1) <> sAuxiliar Then
                        ValidoCondicion = False
                        Exit For
                    End If
                Next
                If ValidoCondicion And sValor1 <> "" Then
                    If tEntregaT.Text <> "" Then
                        If CCur(tEntregaT.Text) < CCur(sValor1) Then ValidoCondicion = False
                    Else
                        ValidoCondicion = False
                    End If
                End If
                '----------------------------------------------------------------------------------------------------------------------------
                
            Case "MON"  'Valido el Monto de la Solicitud'-----------------------------------------------------------------------------
                'El monto debe ser menor al autorizado (el MONTO ES FINANCIADO)
                Dim aMonto As Currency: aMonto = 0
                For Each itmX In lvVenta.ListItems
                    Cons = "Select * from TipoCuota Where TCuCodigo = " & PlanDeLaClave(itmX.Key)
                    Set rsVal = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    'SubItems(4) = Entrega; SubItems(5) = Cuota
                    If Not IsNull(rsVal!TCuVencimientoE) Then 'Tiene Entrega
                        If rsVal!TCuVencimientoE = 0 Then aMonto = aMonto + CCur(itmX.SubItems(4))
                    Else
                        If Not IsNull(rsVal!TCuVencimientoC) Then 'Tiene Cuota
                            If rsVal!TCuVencimientoC = 0 Then aMonto = aMonto + CCur(itmX.SubItems(5))
                        End If
                    End If
                    rsVal.Close
                Next
                
                If CCur(sAuxiliar) < CCur(lTotal.Caption) - aMonto Then ValidoCondicion = False
                '----------------------------------------------------------------------------------------------------------------------------
            
            Case "FGA"      'Valido que firme con la garantía
                If CLng(tGarantia.Tag) <> CLng(sAuxiliar) Then ValidoCondicion = False
                '----------------------------------------------------------------------------------------------------------------------------
                
            Case "COM"      'PRESENTAR COMPROBANTE
                Dim FV1 As String
                Dim FV2 As String
                'Saco los formatos de datos para el comprobante----------------------------------------
                Cons = "Select * from Comprobante" _
                        & " Where  ComCodigo = " & Mid(sAuxiliar, 1, InStr(sAuxiliar, "V1") - 1)
                Set rsVal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                FV1 = UCase(Trim(rsVal!ComFormatoV1))
                FV2 = ""
                If Not IsNull(rsVal!ComFormatoV2) Then FV2 = UCase(Trim(rsVal!ComFormatoV2))
                rsVal.Close
                '-------------------------------------------------------------------------------------------------
                
                Cons = "Select * from ReferenciaCliente" _
                        & " Where RClCPersona = " & gCliente _
                        & " And RClComprobante = " & Mid(sAuxiliar, 1, InStr(sAuxiliar, "V1") - 1) _
                        & " And RClExhibido >= '" & Format(Date - paVaToleranciaDiasExh, sqlFormatoFH) & "'"
                        
                If InStr(sAuxiliar, "V2") <> 0 Then
                    sValor1 = Mid(sAuxiliar, 1, InStr(sAuxiliar, "V2") - 1)
                    sValor1 = Trim(Mid(sValor1, InStr(sValor1, "V1") + 2, Len(sValor1)))
                    sValor2 = Trim(Mid(sAuxiliar, InStr(sAuxiliar, "V2") + 2, Len(sAuxiliar)))
                Else
                    sValor1 = Trim(Mid(sAuxiliar, InStr(sAuxiliar, "V1") + 2, Len(sAuxiliar)))
                    sValor2 = ""
                End If
                
                If FV1 = "CEDULA" Then
                    'Como tengo el codigo del cliente -> saco la CI para el cliente
                    sAuxiliar = "Select * from Cliente Where CliCodigo = " & sValor1
                    Set rsVal = cBase.OpenResultset(sAuxiliar, rdOpenForwardOnly, rdConcurValues)
                    sValor1 = ""
                    If Not IsNull(rsVal!CliCIRuc) Then sValor1 = Trim(rsVal!CliCIRuc)
                    rsVal.Close
                End If
                If FV2 = "CEDULA" Then
                    'Como tengo el codigo del cliente -> saco la CI para el cliente
                    sAuxiliar = "Select * from Cliente Where CliCodigo = " & sValor2
                    Set rsVal = cBase.OpenResultset(sAuxiliar, rdOpenForwardOnly, rdConcurValues)
                    sValor2 = ""
                    If Not IsNull(rsVal!CliCIRuc) Then sValor2 = Trim(rsVal!CliCIRuc)
                    rsVal.Close
                End If
                
                If (FV1 <> "" And sValor1 = "") Or (FV2 <> "" And sValor2 = "") Then
                    ValidoCondicion = False
                    Exit Function
                End If
                
                If FV1 = "MONEDA" Then          '>  val - por.
                    If CCur(sValor1) <> 0 Then
                        Cons = Cons & " And Convert(real, RClValor1) >= " & (CLng(sValor1) - CLng(sValor1) * (paVaToleranciaMonedaPorc / 100))
                    End If
                Else
                    Cons = Cons & " And RClValor1 = '" & FormatoGrabarReferencia(sValor1, FV1) & "'"
                End If
                If FV2 <> "" Then
                    If FV2 = "MONEDA" Then
                        If CCur(sValor2) <> 0 Then
                            Cons = Cons & " And Convert(real, RClValor2) >= " & (CLng(sValor2) - CLng(sValor2) * (paVaToleranciaMonedaPorc / 100))
                        End If
                    Else
                        Cons = Cons & " And RClValor2 = '" & FormatoGrabarReferencia(sValor2, FV2) & "'"
                    End If
                End If
                Set rsVal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If rsVal.EOF Then ValidoCondicion = False
                rsVal.Close
                '----------------------------------------------------------------------------------------------------------------------------
                
            Case "FCO"      'Valido que tenga cónyuge y que la garantia sea el conyuge-------------------------------------------
                Cons = "Select * from CPersona Where CPeCliente = " & gCliente
                Set rsVal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                ValidoCondicion = False
                If Not rsVal.EOF Then
                    If Not IsNull(rsVal!CPeConyuge) Then
                        If CLng(tGarantia.Tag) = rsVal!CPeConyuge Then ValidoCondicion = True
                    End If
                End If
                rsVal.Close
                '----------------------------------------------------------------------------------------------------------------------------
                
            Case "GPR"  'Garantia Exhibió Titulos--------------------------------------------------------------------------------------
                'Primero valido que tenga la garantia
                If CLng(tGarantia.Tag) <> 0 Then
                    Cons = "Select * from Titulo" _
                            & " Where TitCliente = " & CLng(tGarantia.Tag) _
                            & " And TitExhibido >= '" & Format(Date - paVaToleranciaDiasExhTit, sqlFormatoFH) & "'"
                    Set rsVal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If rsVal.EOF Then ValidoCondicion = False
                    rsVal.Close
                Else
                    ValidoCondicion = False
                End If
                '----------------------------------------------------------------------------------------------------------------------------
            
            Case "PRO"      'Exhibió Titulos---------------------------------------------------------------------------------------------
                Cons = "Select * from Titulo" _
                        & " Where TitCliente = " & gCliente _
                        & " And TitExhibido >= '" & Format(Date - paVaToleranciaDiasExhTit, sqlFormatoFH) & "'"
                Set rsVal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If rsVal.EOF Then ValidoCondicion = False
                rsVal.Close
                '----------------------------------------------------------------------------------------------------------------------------
                
            Case "CLE"
                Dim bCle As Boolean
                bCle = True
                'Dias Clearing Hoy - 15 dias ----- Morosidades con saldo ó en 3 años ninguna morosidad
                '1) Debe haber clearing hechos hasta hace 15 dias
                Cons = "Select * from Clearing" _
                        & " Where CleCliente = " & gCliente _
                        & " And CleFecha >= '" & Format(gFechaServidor - 15, "mm/dd/yyyy") & "'"
                Set rsVal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If rsVal.EOF Then ValidoCondicion = False: bCle = False
                rsVal.Close
                
                If bCle Then
                    '2) No debe tener morosidades con saldo
                    Cons = "Select * from ClearingAntecedente" _
                            & " Where CAnCliente = " & gCliente _
                            & " And CAnSaldo > 0"
                    Set rsVal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If Not rsVal.EOF Then ValidoCondicion = False: bCle = False
                    rsVal.Close
                End If
                
                '3) No debe tener morosidades en los ultimos 3 años
                If bCle Then
                    Cons = "Select * from ClearingAntecedente" _
                            & " Where CAnCliente = " & gCliente _
                            & " And CAnFecha >= '" & Format(gFechaServidor - (365 * 3), "mm/dd/yyyy") & "'"
                    Set rsVal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If Not rsVal.EOF Then ValidoCondicion = False: bCle = False
                    rsVal.Close
                End If
                
                
            Case "CHD"      'Pago con Cheque Diferido--------------------------------------------------------------------------------
                If cPago.ListIndex = -1 Then
                    ValidoCondicion = False
                Else
                    If cPago.ItemData(cPago.ListIndex) <> TipoPagoSolicitud.ChequeDiferido Then ValidoCondicion = False
                End If
                '----------------------------------------------------------------------------------------------------------------------------
            
            Case "COP"      'Ccancelar Operaciones Pend.--------------------------------------------------------------------------------
                Cons = "Select * from Credito, Documento" _
                        & " Where CreCliente = " & gCliente _
                        & " And CreSaldoFactura > 0" _
                        & " And CreFactura = DocCodigo" _
                        & " And DocAnulado = 0"
                Set rsVal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not rsVal.EOF Then ValidoCondicion = False
                rsVal.Close
                '----------------------------------------------------------------------------------------------------------------------------
                
            Case "CDE"      'Cancelar Deuda Pendiente-------------------------------------------------------------------------------
                'Si tiene el clic no valido ya que si tiene las va a facturar.
                If chkVencidaTit.value = 0 Then
                    Cons = "Select * from Credito, Documento" _
                            & " Where CreCliente = " & gCliente _
                            & " And CreSaldoFactura > 0" _
                            & " And CreProximoVto < GetDate()" _
                            & " And CreFactura = DocCodigo" _
                            & " And DocAnulado = 0"
                    Set rsVal = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    If Not rsVal.EOF Then ValidoCondicion = False
                    rsVal.Close
                End If
                '----------------------------------------------------------------------------------------------------------------------------
                
        End Select
        If Not ValidoCondicion Then Exit Function
    Loop
    Exit Function
    
errValidar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al validar las condiciones.", Err.Description
    ValidoCondicion = False
End Function

'--------------------------------------------------------------------------------------------------------
'   Trabaja con una variable global gTextoCondicion en la que va retornando las siguientes
'   referencias, y retorna la primer condicion que encuentre en el texto.
'   Ejemplo:   Si recibe TextoCondicion = (GPR)(CPL10E1000)
'                   Retorna GPR, y en texto condicion queda (CPL10E1000)
'--------------------------------------------------------------------------------------------------------
Private Function SacoCondicionDelTexto() As String

    If InStr(gTextoCondicion, ")") > 0 And InStr(gTextoCondicion, "(") > 0 And InStr(gTextoCondicion, "(") < InStr(gTextoCondicion, ")") Then
        SacoCondicionDelTexto = Mid(gTextoCondicion, 2, InStr(gTextoCondicion, ")") - 2)       'Saco la condicion sin los parentesis
        gTextoCondicion = Trim(Mid(gTextoCondicion, InStr(gTextoCondicion, ")") + 1, Len(gTextoCondicion)))       'Resto del texto condicion
    ElseIf InStr(gTextoCondicion, ")") > 0 And InStr(gTextoCondicion, "(") = 0 Then
        gTextoCondicion = ""
    End If
End Function

Private Sub Label13_Click()
    Foco tUsuario
End Sub

'Private Sub Label12_Click()
'
'    If oCnfgPrint.Opcion = 0 Then
'        AccionImprimirConforme 7183958, 0, 0
'    Else
'        ImprimoConformeTickets 7183958
'    End If
'
'End Sub

Private Sub Label14_Click()
    Foco tEntrega
End Sub

Private Sub Label15_Click()
    Foco tEntregaT
End Sub

Private Sub Label16_Click()
    Foco tGarantia
End Sub

Private Sub Label17_Click()
    Foco cCuota
End Sub


Private Sub Label26_Click()
    Foco cPendiente
End Sub

Private Sub Label27_Click()
    Foco cPago
End Sub

Private Sub Label4_Click()
    Foco tFRetiro
End Sub

Private Sub Label53_Click()
    Foco tComentario
End Sub

Private Sub Label6_Click()
    Foco tCantidad
End Sub

Private Sub lEnvio_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF2 Then
        lvVenta.ZOrder 0
        lvVenta.SetFocus
    End If
    
End Sub

Private Sub lvVenta_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrlvKD
    
    If lvVenta.ListItems.Count > 0 Then
        Select Case KeyCode
            Case vbKeySpace    'EDITO EL RENGLON----------------------------------------------------------------------------------
                If aIDServicio <> 0 Then
                    MsgBox "Esta solicitud es para facturar un servicio." & Chr(vbKeyReturn) & "No podrá modificar los datos para facturarla.", vbExclamation, "Solicitud para Facturar Servicios"
                    Exit Sub
                End If
                EditoRenglon lvVenta.SelectedItem
                
            Case vbKeyDelete    'ELIMINO EL RENGLON----------------------------------------------------------------------------------
                If aIDServicio <> 0 Then
                    MsgBox "Esta solicitud es para facturar un servicio." & Chr(vbKeyReturn) & "No podrá modificar los datos para facturarla.", vbExclamation, "Solicitud para Facturar Servicios"
                    Exit Sub
                End If
                
                If lvVenta.SelectedItem.SubItems(11) <> "" Then
                    MsgBox "El artículo seleccionado está asignado a un envío. Para elimarlo acceda al envío.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                End If
                    
                If lvVenta.ListItems.Count = 1 Then
                    MsgBox "No se pueden eliminar todos los artículos solicitados.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                End If
                    
                If Trim(lvVenta.SelectedItem.SubItems(6)) <> "" Then
                    lTotal.Caption = Format(CCur(lTotal.Caption) - CCur(lvVenta.SelectedItem.SubItems(6)), FormatoMonedaP)
                End If
                
                'Si el que borro es plan con entrega REDISTRIBUYO------------------------------------
                If Left(lvVenta.SelectedItem.Key, 1) = "E" Then
                    lvVenta.ListItems.Remove lvVenta.SelectedItem.Index
                    If IsNumeric(tEntregaT.Text) And Trim(tEntregaT.Text) <> "" And lvVenta.ListItems.Count > 0 Then
                        DistribuirEntregas CCur(tEntregaT.Text)
                    Else
                        tEntregaT.Text = ""
                    End If
                Else
                    lvVenta.ListItems.Remove lvVenta.SelectedItem.Index
                End If
            
            Case vbKeyE         'CAMBIO ESTADO DE ENVIO S/N-----------------------------------------------------------
                
                If aIDServicio <> 0 Then
                    MsgBox "Esta solicitud es para facturar un servicio." & Chr(vbKeyReturn) & "En este tipo de factura no se permiten realizar envíos.", vbExclamation, "Solicitud para Facturar Servicios"
                    Exit Sub
                End If
                
                Dim mItm As ListItem                            'Esto es para que no me cambie el item cuando presiona E
                Set mItm = lvVenta.SelectedItem
                DoEvents
                Set lvVenta.SelectedItem = mItm: lvVenta.Refresh
                
                'Si el articulo es de flete no se puede editar
                If Trim(aFletes) = "" Then aFletes = CargoArticulosDeFlete
                If InStr(aFletes, ArticuloDeLaClave(mItm.Key) & ",") <> 0 Then
                    MsgBox "El artículo seleccionado es del tipo fletes. No se puede enviar.", vbExclamation, "ATENCIÓN"
                    Exit Sub
                End If
                
                If Trim(mItm.SubItems(7)) = "No" Then
                    aTexto = HayEnviosIngresados(PlanDeLaClave(mItm.Key))
                    If aTexto <> "" Then
                        If MsgBox("Hay envíos ingresados para el plan seleccionado, desea modificarlos.", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
                            mItm.SubItems(7) = "Si"
                            ModificarEnvio PlanDeLaClave(mItm.Key), aTexto
                        End If
                    Else
                        mItm.SubItems(7) = "Si"
                    End If
                Else
                    If mItm.SubItems(11) <> "" Then
                        If MsgBox("El artículo seleccionado está asignado a un envío. Desea ir al envío.", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
                            ModificarEnvio PlanDeLaClave(mItm.Key), mItm.SubItems(11)
                        End If
                    Else
                        mItm.SubItems(7) = "No"
                    End If
                End If  '------------------------------------------------------------------------------------------------------------

            Case vbKeyReturn: Foco tFRetiro 'If tEntregaT.Enabled Then tEntregaT.SetFocus Else cPago.SetFocus
            
            Case vbKeyF2: lEnvio.ZOrder 0: lEnvio.SetFocus
            
        End Select
    End If
    Exit Sub

ErrlvKD:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error inesperado."
End Sub

Private Sub EditoRenglon(mItm As ListItem)

Dim aFletes As String

    'Si el articulo es de flete no se puede editar
    ' ***25/03/2004 Es por las notas, ahora se están solicitando articulos de fletes, por lo tanto se financian !!
    '   Al hacer la nota consulto en la solicitud Para Sacar el Costo y no restarlo de la 1a Cuota Si no queda el la nota Pago -39 xej
    
    If Trim(aFletes) = "" Then aFletes = CargoArticulosDeFlete
    If InStr(aFletes, ArticuloDeLaClave(mItm.Key) & ",") <> 0 Then
        MsgBox "El artículo seleccionado es del tipo fletes. No se puede editar.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    If mItm.SubItems(11) <> "" Then
        MsgBox "El artículo seleccionado está asignado a un envío. No podrá modificarlo.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    'Saco el Codigo del Articulo
    tArticulo.Text = Trim(mItm.SubItems(2))
    tArticulo.Tag = ArticuloDeLaClave(mItm.Key)
    
    'Cargo los tipos de cuota para la moneda y articulo seleccionado---------------------------------------------
    aPlan = 0
    cCuota.Clear
    Cons = "Select TCuCodigo, TCuAbreviacion, PViPlan From PrecioVigente, TipoCuota" _
            & " Where PViArticulo = " & tArticulo.Tag _
            & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
            & " And PViTipoCuota = TCuCodigo " _
            & " And PViTipoCuota <> " & paTipoCuotaContado
    If oCliente.Categoria.Codigo = 0 Or oCliente.Categoria.Codigo = paCategoriaCliente Then Cons = Cons & " And TCuEspecial = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        aPlan = RsAux!PViPlan
        cCuota.AddItem Trim(RsAux!TCuAbreviacion)
        cCuota.ItemData(cCuota.NewIndex) = RsAux!TCuCodigo
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If aPlan <> 0 Then
        Cons = "Select TCuCodigo, TCuAbreviacion from Coeficiente, TipoCuota" _
                & " Where CoePlan = " & aPlan _
                & " And CoeMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                & " And CoeTipoCuota = TCuCodigo" _
                & " And TCuVencimientoE is Not NULL"
        If oCliente.Categoria.Codigo = 0 Or oCliente.Categoria.Codigo = paCategoriaCliente Then Cons = Cons & " And TCuEspecial = 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            cCuota.AddItem Trim(RsAux!TCuAbreviacion)
            cCuota.ItemData(cCuota.NewIndex) = RsAux!TCuCodigo
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    
    BuscoCodigoEnCombo cCuota, PlanDeLaClave(mItm.Key)
    cCuota.Tag = mItm.Tag           'Cantidad de Cuotas
    
    tComentarioR.Text = Trim(mItm.SubItems(12))
    tCantidad.Text = mItm.SubItems(1)
    lSubTotalF.Caption = mItm.SubItems(6)
    
    If Mid(mItm.Key, 1, 1) = "E" Then
        sConEntrega = True
        'Saco el Precio Contado------------------------
        Cons = "Select PViPrecio, PViPlan From PrecioVigente" _
            & " Where PViArticulo = " & tArticulo.Tag _
            & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
            & " And PViTipoCuota = " & paTipoCuotaContado
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then
            lSubTotalF.Tag = RsAux!PViPrecio
            aPlan = RsAux!PViPlan
        End If
        RsAux.Close
        tEntrega.Text = mItm.SubItems(4)
    Else
        sConEntrega = False
        lSubTotalF.Tag = CLng(mItm.SubItems(6)) / CLng(tCantidad.Text)
    End If
    '-----------------------------------------------------------------------------------------------------------
    If cCuota.ListIndex = -1 Then
        'Como no hay precios SOLO se puede editar para ingresar comentarios
        cCuota.AddItem Trim(mItm.Text)
        cCuota.ItemData(cCuota.NewIndex) = PlanDeLaClave(mItm.Key)
        BuscoCodigoEnCombo cCuota, PlanDeLaClave(mItm.Key)
        
        MsgBox "No se encontraron precios para el artículo." & Chr(vbKeyReturn) & "El renglón se editará para ingresar comentarios.", vbExclamation, "ATENCIÓN"
        HabilitoRenglon True, True
        Foco tComentarioR
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    HabilitoRenglon True
    lvVenta.Enabled = False
    Foco tCantidad
    Screen.MousePointer = 0
End Sub

Private Sub lvVenta_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub tCantidad_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        LimpioRenglon
        HabilitoRenglon False
        lvVenta.Enabled = True
        lvVenta.SetFocus
    End If
    
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And IsNumeric(tCantidad.Text) Then
        
        If CCur(tCantidad.Text) <= 0 Then
            MsgBox "La cantidad ingresada no es correcta.", vbExclamation, "ATENCIÓN"
            Exit Sub
        End If
        
        If CCur(tCantidad.Text) > CCur(lvVenta.SelectedItem.SubItems(1)) Then
            MsgBox "La cantidad de artículos no puede superar la cantidad origial.", vbExclamation, "ATENCIÓN"
            Foco tCantidad
            Exit Sub
        End If

        If Not sConEntrega Then
            'lSubTotalF.Tag = BuscoDescuentoCliente(CLng(tArticulo.Tag), gCategoriaCliente, CCur(lSubTotalF.Tag), CCur(tCantidad.Text), tArticulo.Text, CLng(cCuota.Tag))
            If lSubTotalF.Caption = "" Then
                MsgBox "Los datos de la financiación no se han cargado (SubTotal (F)).", vbExclamation, "Posible Error"
                cCuota.SetFocus: Screen.MousePointer = 0: Exit Sub
            End If
            lSubTotalF.Caption = Format(CCur(lSubTotalF.Tag) * CCur(tCantidad.Text), FormatoMonedaP)
            InsertoRenglon
        Else
            If IsNumeric(tEntrega.Text) And Trim(tEntrega.Text) <> "" Then
                'If CalculoPrecioConEntrega Then Foco tEntrega
                Foco tEntrega       'lo valido en la entrega
            Else
                If Trim(tEntrega.Text) = "" Then Foco tEntrega
            End If
        End If
    End If
    
End Sub

Private Function CalculoPrecioConEntrega() As Boolean

    On Error GoTo errPrecio
    CalculoPrecioConEntrega = False
    If aPlan = 0 Then
        MsgBox "Se debe seleccionar una financiación.", vbExclamation, "Fata Financiación"
        Foco cCuota: Exit Function
    End If
    
    Screen.MousePointer = 11
    
    'El valor de la cuota es el (Precio Contado - Entrega) * Coeficiente ----- Coeficiente (Plan, TCuota, Moneda)
    Cons = "Select * from Coeficiente, TipoCuota" _
        & " Where CoePlan = " & aPlan _
        & " And CoeTipoCuota = " & cCuota.ItemData(cCuota.ListIndex) _
        & " And CoeMoneda = " & mMoneda _
        & " And CoeTipocuota = TCuCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    'Calculo lo que queda por pagar ((P.contado * Cantidad)- Entega * Coeficiente)
    iAuxiliar = ((CCur(lSubTotalF.Tag) * CCur(tCantidad.Text)) - CCur(tEntrega.Text)) * RsAux!CoeCoeficiente
    
    'Veo si tiene descuento
    iAuxiliar = CCur(BuscoDescuentoCliente(CLng(tArticulo.Tag), oCliente.Categoria.Codigo, iAuxiliar, _
                             CCur(tCantidad.Text), tArticulo.Text, RsAux!TCuCodigo))
    
    'SubTotal = (Entrega + Las cuotas)
    Dim mSubTotal As Currency
    
    mSubTotal = Redondeo((iAuxiliar / RsAux!TCuCantidad), mRedondeo) * RsAux!TCuCantidad
    mSubTotal = mSubTotal + CCur(tEntrega.Text)
    lSubTotalF.Caption = Format(mSubTotal, "#,##0.00")
'    lSubTotalF.Caption = Format((Redondeo(iAuxiliar / RsAux!TCuCantidad) * RsAux!TCuCantidad) + CCur(tEntrega.Text), "#,##0.00")
    
    Screen.MousePointer = 0
    CalculoPrecioConEntrega = True
    Exit Function
    
errPrecio:
    clsGeneral.OcurrioError "Error al calcular el subtotal financiado.", Err.Description
    Screen.MousePointer = 0
End Function

Private Sub tComentario_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If chkRetiraAqui.Visible Then
            chkRetiraAqui.SetFocus
        Else
            Foco tUsuario
        End If
    End If
End Sub


Private Sub tComentarioR_GotFocus()
    tComentarioR.SelStart = 0
    tComentarioR.SelLength = Len(tComentarioR.Text)
End Sub

Private Sub tComentarioR_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        LimpioRenglon
        HabilitoRenglon False
        lvVenta.Enabled = True
        lvVenta.SetFocus
    End If
End Sub

Private Sub tComentarioR_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If tCantidad.Enabled Then
            Foco tCantidad
        Else
            lvVenta.SelectedItem.SubItems(12) = Trim(tComentarioR.Text)
            LimpioRenglon
            HabilitoRenglon False
            lvVenta.Enabled = True
            lvVenta.SetFocus
        End If
    End If
End Sub

Private Sub tEntrega_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        LimpioRenglon
        HabilitoRenglon False
        lvVenta.Enabled = True
        lvVenta.SetFocus
    End If
    
End Sub

Private Sub tEntrega_KeyPress(KeyAscii As Integer)

    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If Not IsNumeric(tEntrega.Text) Then
        MsgBox "El valor de entrega ingresado no es correcto.", vbCritical, "ATENCIÓN"
        Exit Sub
    End If
    
    'Valido los datos de cantidad-------------------------------------------------------------
    If Not IsNumeric(tCantidad.Text) Then
        MsgBox "La cantidad ingresada no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tCantidad
        Exit Sub
    End If
    If CCur(tCantidad.Text) <= 0 Then
        MsgBox "La cantidad ingresada no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tCantidad
        Exit Sub
    End If
    If CCur(tCantidad.Text) > CCur(lvVenta.SelectedItem.SubItems(1)) Then
        MsgBox "La cantidad de artículos no puede superar la cantidad origial.", vbExclamation, "ATENCIÓN"
        Foco tCantidad
        Exit Sub
    End If
    '------------------------------------------------------------------------------------------------
    
    tEntrega.Text = Redondeo(CCur(tEntrega.Text), mRedondeo)
    If CCur(tEntrega.Text) >= CCur(lSubTotalF.Tag) * CCur(tCantidad.Text) Then
        MsgBox "El valor de entrega no debe superar el precio contado del artículo.", vbExclamation, "ATENCIÓN"
    Else
        If CalculoPrecioConEntrega Then InsertoRenglon
    End If
    
End Sub

Private Sub tEntrega_LostFocus()

    If IsNumeric(tEntrega.Text) Then tEntrega.Text = Format(tEntrega.Text, FormatoMonedaP)
    
End Sub

Private Sub tEntregaT_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If Trim(tEntregaT.Text) <> "" Then
            If IsNumeric(tEntregaT.Text) Then
                 tEntregaT.Text = Redondeo(tEntregaT.Text, mRedondeo)
                If CCur(tEntregaT.Text) < CCur(tEntregaT.Tag) Then
                    MsgBox "El monto de entrega no puede ser menor al original.", vbCritical, "ATENCIÓN"
                    Exit Sub
                End If
                tEntregaT.Text = Format(tEntregaT.Text, "#,##0.00")
                DistribuirEntregas CCur(tEntregaT.Text)
            End If
            cPago.SetFocus
        End If
    End If
End Sub

Private Sub tFRetiro_Change()
    tFRetiro.Tag = ""
End Sub

Private Sub tFRetiro_GotFocus()
    tFRetiro.SelStart = 0
    tFRetiro.SelLength = Len(tFRetiro.Text)
End Sub

Private Sub tFRetiro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim fRetira As Date
        fRetira = ObtenerMayorFechaRetiro()
        If CDate(tFRetiro.Text) < fRetira Then
            MsgBox "La fecha de retiro es incorrecta, se sustiuirá.", vbExclamation, "Atención"
            tFRetiro.Text = Format(fRetira, "d-Mmm yyyy")
        End If
        Foco tComentario
    ElseIf UCase(Chr(KeyAscii)) = "E" Then
        On Error Resume Next
        KeyAscii = 0
        lvVenta.SetFocus
        lvVenta.SelectedItem = lvVenta.SelectedItem
        Exit Sub
    End If
End Sub

Private Sub tFRetiro_LostFocus()
    
    If IsDate(tFRetiro.Text) Then
        If Trim(tFRetiro.Tag) = "" Then tFRetiro.Text = Format(tFRetiro.Text, "d-Mmm yyyy")
    Else
        tFRetiro.Text = ""
    End If
    
End Sub

Private Sub tFRetiro_Validate(Cancel As Boolean)
    If IsDate(tFRetiro.Text) Then
        Dim fRetira As Date
        fRetira = ObtenerMayorFechaRetiro()
        If CDate(tFRetiro.Text) < fRetira Then
            MsgBox "La fecha de retiro es incorrecta, se sustiuirá.", vbExclamation, "Atención"
            tFRetiro.Text = Format(fRetira, "d-Mmm yyyy")
        End If
    Else
        Cancel = True
    End If
End Sub

Private Sub tGarantia_Change()
    tGarantia.Tag = 0
End Sub

Private Sub tGarantia_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And Trim(tGarantia.Tag) = 0 Then
        Screen.MousePointer = 11
        'Valido la Cédula ingresada----------
        If Trim(tGarantia.Text) <> "" Then
            If Not clsGeneral.CedulaValida(clsGeneral.QuitoFormatoCedula(tGarantia.Text)) Then
                Screen.MousePointer = vbDefault
                lGarantia.Caption = ""
                MsgBox "La cédula de identidad ingresada no es válida.", vbExclamation, "ATENCIÓN"
                Exit Sub
            End If
        End If
        'Busco el Cliente -----------------------
        If Trim(tGarantia.Text) <> "" Then
            Cons = "Select CliCodigo, CliCIRuc, CliCategoria, CPeNombre1, CPeNombre2, CPeApellido1, CPeAPellido2, CliDireccion, CPeRUC From Cliente, CPersona " _
                    & " Where CliCiRuc = '" & clsGeneral.QuitoFormatoCedula(tGarantia.Text) & "'" _
                    & " And CliTipo = " & TipoCliente.Cliente _
                    & " And CliCodigo = CPeCliente"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If RsAux.EOF Then
                Screen.MousePointer = 0
                MsgBox "No existe un cliente para la cédula de indentidad ingresada."
                            
            Else        'El cliente ingresado existe-------------------
                lGarantia.Caption = " " & ArmoNombre(Format(RsAux!CPeApellido1, "#"), Format(RsAux!CPeApellido2, "#"), Format(RsAux!CPeNombre1, "#"), Format(RsAux!CPeNombre2, "#"))
                tGarantia.Tag = RsAux!CliCodigo
                If lvVenta.Enabled Then
                    lvVenta.SetFocus
                Else
                    If tEntregaT.Enabled Then
                        Foco tEntrega
                    Else
                        If cPago.Enabled Then Foco cPago
                    End If
                End If
            End If
            RsAux.Close
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
    
errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al buscar la cédula de identidad ingresada."
End Sub

Private Sub LimpioRenglon()

    tArticulo.Text = ""
    tCantidad.Text = ""
    cCuota.Text = ""
    tEntrega.Text = ""
    lSubTotalF.Caption = ""
    tComentarioR.Text = ""
    
    cCuota.Tag = ""
    lSubTotalF.Tag = ""
    tEntrega.Tag = ""
    
End Sub

Private Sub InsertoRenglon()
    
    'E ó N - Con entrega ó Normal  ..... P - Plan .....A - Artículo
    'E/N ...Plan ...Articulo..
    
    On Error GoTo errInsertar
    Set itmX = lvVenta.SelectedItem
    If tEntrega.Enabled Then
        itmX.Key = "EP" & cCuota.ItemData(cCuota.ListIndex) & "A" & tArticulo.Tag
    Else
        itmX.Key = "NP" & cCuota.ItemData(cCuota.ListIndex) & "A" & tArticulo.Tag
    End If
    
    'lvVenta.Tag = cCuota.Tag        'Cantidad de Cuotas
    itmX.Tag = cCuota.Tag              'Esto lo anule el 22/01/03 por error en Cuotas
    
    itmX.Text = Trim(cCuota.Text)
    itmX.SubItems(1) = Trim(tCantidad.Text)
    itmX.SubItems(12) = Trim(tComentarioR.Text)
    itmX.SubItems(14) = zGet_InstaladorArticulo(Val(tArticulo.Tag))
    
    If tEntrega.Enabled Then
        'Entrega
        If Trim(tEntregaT.Text) = "" Or Not IsNumeric(tEntregaT.Text) Then tEntregaT.Text = "0"
        If CCur(tEntregaT.Text) = 0 Then
            tEntregaT.Text = Format(CCur(tEntregaT.Text) + CCur(tEntrega.Text), FormatoMonedaP)
        Else
            If Trim(itmX.SubItems(4)) <> "" Then
                tEntregaT.Text = Format(CCur(tEntregaT.Text) - CCur(itmX.SubItems(4)) + CCur(tEntrega.Text), FormatoMonedaP)
            Else
                tEntregaT.Text = Format(CCur(tEntregaT.Text) + CCur(tEntrega.Text), FormatoMonedaP)
            End If
        End If
        itmX.SubItems(4) = Format(tEntrega.Text, FormatoMonedaP)
        
        'Valor Cuota
        itmX.SubItems(5) = Format((CCur(lSubTotalF.Caption) - CCur(tEntrega.Text)) / CCur(cCuota.Tag), FormatoMonedaP)
    Else
        'Valor Cuota
        itmX.SubItems(5) = Format(CCur(lSubTotalF.Caption) / CCur(cCuota.Tag), FormatoMonedaP)
        itmX.SubItems(4) = ""
    End If
    
    'Totales
    lTotal.Caption = CCur(lTotal.Caption) - CCur(itmX.SubItems(6))
    itmX.SubItems(6) = Format(lSubTotalF.Caption, FormatoMonedaP)
    lTotal.Caption = Format(CCur(lTotal.Caption) + CCur(itmX.SubItems(6)), FormatoMonedaP)
    
    'Veo si quedan planes con entrega para deshabilitarla
    tEntregaT.Enabled = False
    tEntregaT.BackColor = Inactivo
    For Each itmX In lvVenta.ListItems
        If Mid(itmX.Key, 1, 1) = "E" Then
            tEntregaT.Enabled = True
            tEntregaT.BackColor = Obligatorio
            Exit For
        End If
    Next
    If Not tEntregaT.Enabled Then tEntregaT.Text = "0.00"
    '--------------------------------------------------------------
    LimpioRenglon
    HabilitoRenglon False
    lvVenta.Enabled = True: lvVenta.SetFocus
    Exit Sub
    
errInsertar:
    clsGeneral.OcurrioError "Verifique que el artículo no esté ingresado con la misma financiación.", Err.Description
    Screen.MousePointer = 0
End Sub

'----------------------------------------------------------------------------------------------------------------------------
Private Sub DistribuirEntregas(ValorAEntregar As Currency)

Dim sHay As Boolean
Dim iAuxiliar As Currency
Dim aTotal As Currency

    On Error GoTo errDistribuir
    sHay = False
    'Verifico si hay Cuotas con Entegas-------------------------------------------------------------------------
    For Each itmX In lvVenta.ListItems
        If Left(itmX.Key, 1) = "E" Then sHay = True: Exit For
    Next
    
    If Not sHay Then
        Screen.MousePointer = 0
        MsgBox "No hay financiaciones con entrega para realizar la distribución.", vbInformation, "ATENCIÓN"
        tEntregaT.Text = "": tEntregaT.Enabled = False: tEntregaT.BackColor = Inactivo
        Exit Sub
    End If
    '-----------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    'Limpio los campos para recalcular Y saco el Total de precios con entrega-----------
    For Each itmX In lvVenta.ListItems
        If Left(itmX.Key, 1) = "E" Then
            
            If Trim(itmX.SubItems(6)) <> "" Then
                lTotal.Caption = Format(CCur(lTotal.Caption) - CCur(itmX.SubItems(6)), FormatoMonedaP)
            End If
            
            'itmX.SubItems(4) = ""
            'itmX.SubItems(5) = ""
            'itmX.SubItems(6) = ""
            If Trim(itmX.SubItems(8)) = "" Then     'Precio Contado
                'Saco el Precio contado del Articulo--------------------------------------------------------------------------
                Cons = "Select PViPrecio, PViPlan From PrecioVigente" _
                        & " Where PViArticulo = " & ArticuloDeLaClave(itmX.Key) _
                        & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                        & " And PViTipoCuota = " & paTipoCuotaContado & " And PViHabilitado = 1"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not RsAux.EOF Then
                    itmX.SubItems(8) = Format(RsAux!PViPrecio, FormatoMonedaP)             'Contado x Cantidad
                    itmX.SubItems(9) = RsAux!PViPlan
                    RsAux.Close
                Else
                    RsAux.Close
                    'Si no lo tengo lo saco por cuentas
                    '(ValorCuota * CantCuotas) * Coeficiente = 1xx% ---> contado = 100 + Entrega
                    '1)  Saco el Coeficiente del plan por defecto
                    Cons = "Select * from Coeficiente, TipoCuota" _
                            & " Where CoePlan = " & paPlanPorDefecto _
                            & " And CoeTipoCuota = " & PlanDeLaClave(itmX.Key) _
                            & " And CoeMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                            & " And CoeTipocuota = TCuCodigo"
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    iAuxiliar = CCur(itmX.SubItems(5)) * RsAux!TCuCantidad * 100 / (RsAux!CoeCoeficiente * 100)
                    iAuxiliar = iAuxiliar + CCur(itmX.SubItems(4))      'Monto a financiar + Entrega = Contado
                    itmX.SubItems(8) = Redondeo(iAuxiliar, mRedondeo)            'Contado
                    itmX.SubItems(9) = paPlanPorDefecto
                    RsAux.Close
                End If
            End If
            itmX.SubItems(4) = ""
            itmX.SubItems(5) = ""
            itmX.SubItems(6) = ""
            
            aTotal = aTotal + (CCur(itmX.SubItems(8)) * CCur(itmX.SubItems(1)))         'Contado x Cantidad
        End If
    Next
    
    If aTotal <= ValorAEntregar Then
        Screen.MousePointer = 0
        MsgBox "El valor de entrega no debe superar los contados de los artículos." & Chr(vbKeyReturn) & "Verifique los datos.", vbExclamation, "ATENCIÓN"
        Foco tEntregaT
        Exit Sub
    End If
    '--------------------------------------------------------------------------------------------------
    
    Dim itmP As ListItem
    For Each itmX In lvVenta.ListItems
        If Left(itmX.Key, 1) = "E" Then
            'Veo Si ya hice la distribucion
            If Trim(itmX.SubItems(4)) = "" Then
                'Con el Total distriubuyo el porcentaje de la entrega
                For Each itmP In lvVenta.ListItems
                    If itmX.Text = itmP.Text Then   'El mismo Tipo de Cuota
                        itmP.SubItems(4) = Format(((CCur(itmP.SubItems(8)) * CCur(itmP.SubItems(1)) * 100) / aTotal) * ValorAEntregar / 100, "#,##0.00")
                        
                        'El valor de la cuota es el (Precio Contado - Entrega) * Coeficiente ----- Coeficiente (Plan, TCuota, Moneda)
                        Cons = "Select * from Coeficiente, TipoCuota" _
                            & " Where CoePlan = " & itmP.SubItems(9) _
                            & " And CoeTipoCuota = " & Mid(itmP.Key, 3, InStr(itmP.Key, "A") - 3) _
                            & " And CoeMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                            & " And CoeTipocuota = TCuCodigo"
                        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                        'Calculo lo que queda por pagar ((P.contado * Cantidad)- Entega * Coeficiente)
                        iAuxiliar = ((CCur(itmP.SubItems(8)) * CCur(itmP.SubItems(1))) - CCur(itmP.SubItems(4))) * RsAux!CoeCoeficiente
                        
                        'Veo si tiene descuento
                        iAuxiliar = CCur(BuscoDescuentoCliente(ArticuloDeLaClave(itmP.Key), oCliente.Categoria.Codigo, iAuxiliar, _
                                                 CCur(itmP.SubItems(1)), itmP.SubItems(2), PlanDeLaClave(itmP.Key)))
                        
                        'Valor de Cada Cuota
                        itmP.SubItems(5) = Format(Redondeo(iAuxiliar / RsAux!TCuCantidad, mRedondeo), "#,##0.00")
                        'SubTotal = (Entrega + Las cuotas)
                        itmP.SubItems(6) = Format((CCur(itmP.SubItems(5)) * RsAux!TCuCantidad) + CCur(itmP.SubItems(4)), "#,##0.00")
                        
                        lTotal.Caption = Format(CCur(lTotal.Caption) + CCur(itmP.SubItems(6)), FormatoMonedaP)
                        RsAux.Close
                    End If
                Next
            End If
        End If
    Next
    Screen.MousePointer = 0
    Exit Sub

errDistribuir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la distribución de la entrega."
End Sub

'--------------------------------------------------------------------------------------------------------------------------------------------------
'   Parametros:
'       CodArticulo = id de Articulo.
'       CodCatCliente = categoria de cliente
'       Unitario = Precio del articulo para aplicar el dto.
'       Cantidad = cantidad de articulos para validar la minima.
'       Articulo = texto del articulo para dar mensaje.
'       Plan = id de plan
'
'   RETORNA = precio unitario con o sin descuentos
'--------------------------------------------------------------------------------------------------------------------------------------------------
Private Function BuscoDescuentoCliente(CodArticulo As Long, CodCatCliente As Long, Unitario As Currency, Cantidad As Currency, _
                                                          Articulo As String, Plan As Long) As String

Dim RsBDC As rdoResultset
Dim aRetorno As String

    On Error GoTo errDescuento
    aRetorno = Redondeo(Unitario, mRedondeo)
    
    If CodCatCliente > 0 Then
    
        Cons = "Select CDTPorcentaje, AFaCantidadD From ArticuloFacturacion, CategoriaDescuento" _
                & " Where AFaArticulo = " & CodArticulo _
                & " And AFaCategoriaD = CDtCatArticulo " _
                & " And CDtCatCliente = " & CodCatCliente _
                & " And CDtCatPlazo = " & Plan
            
        Set RsBDC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        If Not RsBDC.EOF Then
            If Not IsNull(RsBDC!AFaCantidadD) Then
                If RsBDC!AFaCantidadD <= Cantidad Then
                    aRetorno = Redondeo(Unitario - (Unitario * RsBDC(0)) / 100, mRedondeo)
                Else
                    If MsgBox("La cantidad no llega a la mímima (" & RsBDC!AFaCantidadD & ") para aplicar el descuento. " & Chr(vbKeyReturn) _
                                & "Desea aplicar el descuento correspondiente.", vbQuestion + vbYesNo, Trim(Articulo)) = vbYes Then
                        aRetorno = Redondeo(Unitario - (Unitario * RsBDC(0)) / 100, mRedondeo)
                    End If
                End If
            End If
        End If
        
        RsBDC.Close
    End If
    
    BuscoDescuentoCliente = aRetorno
    Exit Function

errDescuento:
    clsGeneral.OcurrioError "Error al procesar los descuentos.", Err.Description
    BuscoDescuentoCliente = aRetorno
    Screen.MousePointer = 0
End Function

Private Sub HabilitoRenglon(Valor As Boolean, Optional SoloComentario As Boolean = False)


    tComentarioR.Enabled = Valor
    If Valor Then tComentarioR.BackColor = Blanco Else tComentarioR.BackColor = Inactivo
    If SoloComentario Then Exit Sub
    
    tCantidad.Enabled = Valor
    cCuota.Enabled = Valor
        
    If Not Valor Then
        tEntrega.BackColor = Inactivo
        tEntrega.Enabled = False
    Else
        If sConEntrega Then
            If Valor Then
                tEntrega.BackColor = Obligatorio
                tEntrega.Enabled = True
            Else
                tEntrega.BackColor = Inactivo
                tEntrega.Enabled = False
            End If
        End If
    End If
    
    If Valor Then
        cCuota.BackColor = Obligatorio
        tCantidad.BackColor = Obligatorio
    Else
        cCuota.BackColor = Inactivo
        tCantidad.BackColor = Inactivo
    End If
        
End Sub

Private Function PlanDeLaClave(clave As String) As Long
    PlanDeLaClave = CLng(Mid(clave, 3, InStr(clave, "A") - 3))
End Function
Private Function ArticuloDeLaClave(clave As String) As Long
    ArticuloDeLaClave = CLng(Trim(Mid(clave, InStr(clave, "A") + 1, Len(clave))))
End Function

Private Sub tUsuario_Change()

    tUsuario.Tag = 0
    
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn And IsNumeric(tUsuario.Text) Then
        tUsuario.Tag = z_BuscoUsuarioDigito(Val(tUsuario.Text), Codigo:=True)
        If Val(tUsuario.Tag) = 0 Then
            tUsuario.Text = "": tUsuario.Tag = 0
            Exit Sub
        End If
        If Val(tUsuario.Tag) <> 0 Then
            If bValidar.Enabled Then bValidar.SetFocus Else: If bDevolver.Enabled Then bDevolver.SetFocus
        End If
    End If

End Sub

Function ProximoVencimiento(FechaFactura As Date, FechaDesde As Date, DistanciaEntreCuotas As Integer) As String

Dim aAnio As String
Dim aMesSiguiente As Integer

    If DistanciaEntreCuotas <> 30 Then
        ProximoVencimiento = FechaDesde + DistanciaEntreCuotas
    Else
        'Al ser 30 dias hay un calculo especial
        aAnio = Year(FechaDesde)
        If Month(FechaDesde) = 12 Then aAnio = Year(FechaDesde) + 1
        
        Select Case Day(FechaFactura)
            Case Is < 29
                If Month(FechaDesde) <> 12 Then aMesSiguiente = Month(FechaDesde) + 1 Else aMesSiguiente = 1
                ProximoVencimiento = Day(FechaFactura) & "/" & aMesSiguiente & "/" & aAnio
            
            Case 29, 30, 31
                Dim aDia As Integer
                aDia = Day(FechaFactura)
                'Si para el mismo mes el 30/31 es fecha y es distinto de la fecha desde va el 30/31
                If IsDate(aDia & "/" & Month(FechaDesde) & "/" & Year(FechaDesde)) And _
                   Format(aDia & "/" & Month(FechaDesde) & "/" & Year(FechaDesde), "dd/mm/yyyy") <> Format(FechaDesde, "dd/mm/yyyy") Then
                        ProximoVencimiento = aDia & "/" & Month(FechaDesde) & "/" & Year(FechaDesde)
                Else
                    'Si es día en el mes siguiente
                    
                    If Month(FechaDesde) <> 12 Then aMesSiguiente = Month(FechaDesde) + 1 Else aMesSiguiente = 1
                    
                    If IsDate(aDia & "/" & aMesSiguiente & "/" & aAnio) Then
                        ProximoVencimiento = aDia & "/" & aMesSiguiente & "/" & aAnio
                    Else
                        'Al ultimo dia del Mes Sig. le sumo la diferencia de dias entre DiaFactura y UltimoDiaMes
                        ProximoVencimiento = UltimoDia("1/" & aMesSiguiente & "/" & aAnio) + (Day(FechaFactura) - Day(UltimoDia("1/" & aMesSiguiente & "/" & aAnio)))
                    End If
                End If
                        
        End Select
    End If

End Function

'------------------------------------------------------------------------------------------------------------------------------
'   Esta rutina pide los envíos en forma consecutiva ---> Para cada plan llama a la pantalla de envio, carga
'   los datos y vuelve a llamar si hay otro plan con envíos.  (no sirve para modificar!!!!).
'------------------------------------------------------------------------------------------------------------------------------
Private Function ProcesoEnvios() As Boolean

Dim aPlanEnvio As Long
Dim idTabla As Integer
Dim aEnvios As String

    For Each itmC In lvVenta.ListItems
        If Trim(itmC.SubItems(11)) = "" And UCase(itmC.SubItems(7)) = "SI" Then        'Si no fue enviado
            aPlanEnvio = PlanDeLaClave(itmC.Key)
            
            idTabla = CargoAuxiliarEnvio(aPlanEnvio)
            
            If idTabla = 0 Then
                clsGeneral.OcurrioError "Ocurrió un error al cargar la tabla auxiliar de envíos. Vuelva a intentar la operación", Err.Description
                Screen.MousePointer = 0: Exit Function
            End If
            
            'Llamo al formulario de Envíos---------------------------------------------------------------------------
            Dim objEnvio As New clsEnvio
            objEnvio.NuevoEnvio cBase, "0", idTabla, gCliente, cMoneda.ItemData(cMoneda.ListIndex), TipoEnvio.Entrega
            Me.Refresh
            aEnvios = objEnvio.RetornoEnvios
            Set objEnvio = Nothing
            Me.Refresh
            If aEnvios = vbNullString Then aEnvios = "0"
                        
            If aEnvios <> "0" Then
                If Not InsertoRenglonEnvio(aEnvios, aPlanEnvio, itmC.Text) Then Exit Function
            End If
            
            'A los articulos q' envie les pongo el codigo (donde dice E - subitms(11))-------------
            For Each itmA In lvVenta.ListItems
                If Trim(itmA.SubItems(11)) = "E" Then itmA.SubItems(11) = aEnvios
            Next    '---------------------------------------------------------------------------------------
                        
        End If
    Next
    
    'A los articulos q' envie y el envio me devolvio cero los limpio -------------
    For Each itmA In lvVenta.ListItems
        If Trim(itmA.SubItems(11)) = "0" Then itmA.SubItems(11) = ""
    Next    '---------------------------------------------------------------------------------------
    
End Function

Private Function CargoAuxiliarEnvio(PlanEnvio As Long, Optional Nuevo As Boolean = True) As Integer

Dim IdAux As Integer
Dim itm2 As ListItem

    Screen.MousePointer = 11
    CargoAuxiliarEnvio = 0
    On Error GoTo ErrBT
    
    IdAux = Autonumerico(TAutonumerico.AuxiliarEnvio)
    
    cBase.BeginTrans
    On Error GoTo ErrRB
    
    'Cargo la tabla auxiliar con los articulos a enviar para el Plan Seleccionado------------------------------------------------------------------------
    For Each itmA In lvVenta.ListItems
        If Nuevo Then
            If Trim(itmA.SubItems(11)) = "" And UCase(itmA.SubItems(7)) = "SI" And PlanEnvio = PlanDeLaClave(itmA.Key) Then        'Si no fue enviado
                itmA.SubItems(11) = "E"
                
                'EAuID  EAuArticulo EAuCantidad
                Cons = "Insert into EnvioAuxiliar (EAuID, EAuArticulo, EAuCantidad) Values (" & _
                                                                IdAux & ", " & ArticuloDeLaClave(itmA.Key) & ", " & itmA.SubItems(1) & ")"
                cBase.Execute (Cons)
            End If
            
        Else
            If PlanDeLaClave(itmA.Key) = PlanEnvio And UCase(itmA.SubItems(7)) = "SI" Then
                itmA.SubItems(11) = "E"
                'EAuID  EAuArticulo EAuCantidad
                Cons = "Insert into EnvioAuxiliar (EAuID, EAuArticulo, EAuCantidad) Values (" & _
                                                                IdAux & ", " & ArticuloDeLaClave(itmA.Key) & ", " & itmA.SubItems(1) & ")"
                cBase.Execute (Cons)
            End If
        End If
    Next
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    cBase.CommitTrans
    CargoAuxiliarEnvio = IdAux
    Screen.MousePointer = 0
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

Private Function InsertoRenglonEnvio(ListaEnvios As String, PlanDelEnvio As Long, NombrePlan As String) As Boolean

Dim aEnvio As Long               'Codigo del Envio a Procesar
Dim Inserte As Boolean          'Señal para saber si ya inserte un envio para ese plan y codigo
Dim RsAux2 As rdoResultset
Dim auxstr As String

Dim gEnviosFacturados As String     'Devuelve los envios facturados
    
    InsertoRenglonEnvio = False
    auxstr = ListaEnvios
    gEnviosFacturados = ""
    
    Do While auxstr <> ""
        If InStr(1, auxstr, ",") > 0 Then
            aEnvio = CLng(Left(auxstr, InStr(1, auxstr, ",") - 1))
            auxstr = Trim(Mid(auxstr, InStr(1, auxstr, ",") + 1, Len(auxstr)))
        Else
            aEnvio = CLng(auxstr)
            auxstr = ""
        End If
        
        Cons = "Select * From Envio  Where EnvCodigo = " & aEnvio
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        If Not RsAux.EOF Then
            'Si lo factura ahora lo cargo.-----------------------------
            If RsAux!EnvFormaPago = TipoPagoEnvio.PagaAhora Then
                gEnviosFacturados = gEnviosFacturados & RsAux!EnvCodigo & ","
                
                Cons = "Select ArtId, ArtNombre, IVaPorcentaje From TipoFlete, Articulo, ArticuloFacturacion, TipoIva" _
                    & " Where TFlCodigo = " & RsAux!EnvTipoFlete _
                    & " And ArtID = TFlArticulo" _
                    & " And ArtId = AFaArticulo " _
                    & " And AFaIva = IvaCodigo"
                Set RsAux2 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                
                If RsAux2.EOF Then
                    MsgBox "Error crítico, se almacenó un valor de tipo de flete incorrecto.", vbCritical, "ATENCIÓN"
                    RsAux2.Close
                    Exit Function
                End If
                
                'Busco si ya hay un articulo del envio para el mismo Plan
                Inserte = False
                For Each itmA In lEnvio.ListItems
                    'La clave tiene que empezar con X
                    If Mid(itmA.Key, 1, 1) = "X" And PlanDelEnvio = PlanDeLaClave(itmA.Key) Then
                        If RsAux2!ArtID = ArticuloDeLaClave(itmA.Key) Then
                            'Ya existe el articulo del envio (sumo cantidad)
                            itmA.SubItems(1) = CLng(itmA.SubItems(1)) + 1
                            itmA.SubItems(4) = Format(CCur(itmA.SubItems(4)) + RsAux!EnvValorFlete, "#,##0.00")
    
                            RsAux2.Close
                            Inserte = True
                            Exit For
                        End If
                    End If
                Next
                If Not Inserte Then
                    Set itmA = lEnvio.ListItems.Add(, "XP" & PlanDelEnvio & "A" & RsAux2!ArtID, Trim(NombrePlan))
                    itmA.SubItems(1) = "1"
                    itmA.SubItems(2) = Trim(RsAux2!ArtNombre)
                    itmA.SubItems(3) = Format(RsAux2!IVaPorcentaje, "#,##0.00")
                    itmA.SubItems(4) = Format(RsAux!EnvValorFlete, "#,##0.00")
                    
                    itmA.SubItems(5) = ListaEnvios
                    RsAux2.Close
                End If
                '----------------------------------------------------------------------------------------Fin Tipo Flete
    
                'Valores de piso.------------------------------------------------------------------------------------------
                If Not IsNull(RsAux!EnvValorPiso) Then
                
                    Cons = "Select ArtID, ArtNombre, IVAPorcentaje From Articulo, ArticuloFacturacion, TipoIva" _
                        & " Where ArtId = " & paArticuloPisoAgencia _
                        & " And ArtID = AFaArticulo And AFaArticulo = " & paArticuloPisoAgencia _
                        & " And AFaIVA = IVACodigo"
                    Set RsAux2 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                
                    If RsAux2.EOF Then
                        MsgBox "Error crítico: existe un costo de piso y no existe un artículo asociado a su facturación.", vbCritical, "ATENCIÓN"
                        RsAux2.Close
                        Exit Function
                    End If
                    
                    Inserte = False
                    For Each itmA In lEnvio.ListItems
                        'La clave tiene que empezar con X
                        If Mid(itmA.Key, 1, 1) = "X" And PlanDelEnvio = PlanDeLaClave(itmA.Key) Then
                            If RsAux2!ArtID = ArticuloDeLaClave(itmA.Key) Then
                                'Ya existe el articulo del envio (sumo cantidad)
                                itmA.SubItems(1) = CLng(itmA.SubItems(1)) + 1
                                itmA.SubItems(4) = Format(CCur(itmA.SubItems(4)) + RsAux!EnvValorPiso, "#,##0.00")
                                
                                RsAux2.Close
                                Inserte = True
                                Exit For
                            End If
                        End If
                    Next
                    
                    If Not Inserte Then
                        Set itmA = lEnvio.ListItems.Add(, "XP" & PlanDelEnvio & "A" & RsAux2!ArtID, Trim(NombrePlan))
                        itmA.SubItems(1) = "1"
                        itmA.SubItems(2) = Trim(RsAux2!ArtNombre)
                        itmA.SubItems(3) = Format(RsAux2!IVaPorcentaje, "#,##0.00")
                        itmA.SubItems(4) = Format(RsAux!EnvValorPiso, "#,##0.00")
                        
                        itmA.SubItems(5) = ListaEnvios
                        RsAux2.Close
                        
                    End If
                End If
                '----------------------------------------------------------------------------------------Fin Valor Piso
            End If      'Fin lo Paga AHORA
            
        End If  'RS Vacio
    
    Loop
    If gEnviosFacturados <> "" Then
        gEnviosFacturados = Mid(gEnviosFacturados, 1, Len(gEnviosFacturados) - 1)
        
        'A Todos los envios que inserte le pongo en el itmx de envios solo los facturados
        For Each itmA In lEnvio.ListItems
            If PlanDeLaClave(itmA.Key) = PlanDelEnvio Then itmA.SubItems(5) = gEnviosFacturados
        Next
    End If
    
    InsertoRenglonEnvio = True
    
End Function

Private Sub BorroEnvios(sEnvios As String)

Dim itmE As ListItem
Dim auxEnvios As String
Dim aEnvio As Long
    
    auxEnvios = sEnvios
    
    On Error GoTo errorBT
    Screen.MousePointer = 11
    cBase.BeginTrans    'COMIENZO TRANSACCION------------------------------------------
    On Error GoTo errorET
    
    Do While auxEnvios <> ""
    
        If InStr(1, auxEnvios, ",") > 0 Then
            aEnvio = CLng(Left(auxEnvios, InStr(1, auxEnvios, ",") - 1))
            auxEnvios = Trim(Mid(auxEnvios, InStr(1, auxEnvios, ",") + 1, Len(auxEnvios)))
        Else
            aEnvio = CLng(auxEnvios)
            auxEnvios = ""
        End If
                
        Cons = "DELETE RenglonEnvio Where REvEnvio = " & aEnvio
        cBase.Execute (Cons)
        
        Cons = "DELETE Envio Where EnvCodigo = " & aEnvio
        cBase.Execute (Cons)
    Loop
    
    For Each itmE In lvVenta.ListItems
        If Trim(itmE.SubItems(11)) = sEnvios Then itmE.SubItems(11) = ""
    Next
    
    cBase.CommitTrans    'FIN DE TRANSACCION------------------------------------------
    Exit Sub

errorBT:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción para borrar los envíos realizados."
    Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "No se han podido borrar los envíos realizados. Códigos: " & sEnvios
    For Each itmE In lvVenta.ListItems
        If Trim(itmE.SubItems(11)) = sEnvios Then itmE.SubItems(11) = ""
    Next
End Sub

Private Function HayEnviosAProcesar() As Boolean

Dim itmE As ListItem
    
    HayEnviosAProcesar = False
    For Each itmE In lvVenta.ListItems
        If UCase(itmE.SubItems(7)) = "SI" And Trim(itmE.SubItems(11)) = "" Then
            HayEnviosAProcesar = True
            Exit Function
        End If
    Next
    
End Function

'-------------------------------------------------------------------------------
'   Verifica si hay envios ingresados para el plan Plan
'-------------------------------------------------------------------------------
Private Function HayEnviosIngresados(Plan As Long) As String

Dim itmE As ListItem
    
    HayEnviosIngresados = ""
    For Each itmE In lvVenta.ListItems
        If PlanDeLaClave(itmE.Key) = Plan And Trim(itmE.SubItems(11)) <> "" Then
            HayEnviosIngresados = itmE.SubItems(11)
            Exit Function
        End If
    Next
    
End Function

Private Sub ModificarEnvio(PlanAModificar As Long, ItmEnvios As String)

Dim itm2 As ListItem
Dim aNombrePlan As String
Dim aEnvios As String
Dim idTabla As Integer

    On Error GoTo errModificar
    aNombrePlan = Trim(lvVenta.SelectedItem.Text)
    idTabla = CargoAuxiliarEnvio(PlanAModificar, False)
            
    If idTabla = 0 Then
        clsGeneral.OcurrioError "Ocurrió un error al cargar la tabla auxiliar de envíos. Vuelva a intentar la operación", Err.Description
        Screen.MousePointer = 0: Exit Sub
    End If
    
    'Borro el/los envios con el plan que edito---------------------------
    i = 1
    Do While i <= lEnvio.ListItems.Count
        If PlanAModificar = PlanDeLaClave(lEnvio.ListItems(i).Key) Then lEnvio.ListItems.Remove i Else i = i + 1
    Loop
    '------------------------------------------------------------------------
    
    'Llamo al formulario de Envíos---------------------------------------------------------------------------
    Dim objEnvio As New clsEnvio
    objEnvio.NuevoEnvio cBase, ItmEnvios, idTabla, gCliente, cMoneda.ItemData(cMoneda.ListIndex), TipoEnvio.Entrega
    Me.Refresh
    aEnvios = objEnvio.RetornoEnvios
    Set objEnvio = Nothing
    If aEnvios = vbNullString Then aEnvios = "0"
            
    Me.Refresh
    
    If aEnvios <> "0" Then
        If Not InsertoRenglonEnvio(Trim(aEnvios), PlanAModificar, aNombrePlan) Then Exit Sub
        'A los articulos q' envie les pongo el codigo (donde dice E - subitms(11))-------------
        For Each itmA In lvVenta.ListItems
            If Trim(itmA.SubItems(11)) = "E" Then 'itmA.SubItems(11) = MaEnvio.pSeleccionado
                'Verifico si el articulo está en el envio pSeleccionado
                Cons = "Select * from RenglonEnvio " _
                       & " Where REvEnvio In (" & aEnvios & ")" _
                       & " And REvArticulo = " & ArticuloDeLaClave(itmA.Key)
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                
                If RsAux.EOF Then
                    itmA.SubItems(7) = "No": itmA.SubItems(11) = ""
                Else
                    itmA.SubItems(11) = aEnvios
                End If
                RsAux.Close
            End If
        Next
    Else
        'Como se borraron los envíos los marco en blanco
        For Each itmA In lvVenta.ListItems
            If Trim(itmA.SubItems(11)) = "E" Then itmA.SubItems(11) = ""
        Next
    End If
    '--------------------------------------------------------------------------------------------------------------------------
    Exit Sub
    
errModificar:
    clsGeneral.OcurrioError "Error al modificar los envíos.", Err.Description
    Screen.MousePointer = 0
End Sub

'--------------------------------------------------------------------------------------------------
'   En la lista deben venir todos los documentos a imprimir con el siguiente formato.
'   E= Entrega, V= Envio.
'   Formato: Nro de Factura + "E" + $entrega + "V" + $envio + ":"    --> 1234E1200V49:
'--------------------------------------------------------------------------------------------------

Private Sub AccionImprimir(Lista As String)

On Error GoTo ErrCrystal
Dim iFactura As Long         'Factura para imprimir
Dim iEnvio As Currency      'Monto del Envio para imprimir
Dim iEntrega As Currency    'Monto de la entrega
Dim FormulasCr As Integer, FormulasCo As Integer   'Cantidad Formulas Cr-Credito Co-Conforme
Dim bHayEnvios As Boolean

Dim sErr As String

    Screen.MousePointer = 11
    
    sErr = "SETEO Config"
    If ChangeCnfgPrint Then prj_LoadConfigPrint bShowFrm:=False
    '----------------------------------------------------------------------------------------------------------------------------
    'Configuro la Impresora
    If Trim(Printer.DeviceName) <> Trim(paICreditoN) Then SeteoImpresoraPorDefecto paICreditoN
'    If Not crSeteoImpresora(iJobCre, Printer, paICreditoB) Then GoTo ErrCrystal
        
    'Configuro la Impresora
'    If Trim(Printer.DeviceName) <> Trim(paIConformeN) Then SeteoImpresoraPorDefecto paIConformeN
 '   If Not crSeteoImpresora(iJobCon, Printer, paIConformeB, paIConformeP) Then GoTo ErrCrystal
    '----------------------------------------------------------------------------------------------------------------------------
        
    
    'Obtengo la cantidad de formulas que tiene el reporte.
'    FormulasCr = crObtengoCantidadFormulasEnReporte(iJobCre)
'    If FormulasCr = -1 Then GoTo ErrCrystal
'    FormulasCo = crObtengoCantidadFormulasEnReporte(iJobCon)
'    If FormulasCo = -1 Then GoTo ErrCrystal
    
    Do While Trim(Lista) <> ""
    
        'Saco la factura y el $$ del envio para la factura.-----------------
        iFactura = Mid(Lista, 1, InStr(Lista, "E") - 1)
        iEntrega = Mid(Lista, InStr(Lista, "E") + 1, InStr(Lista, "V") - InStr(Lista, "E") - 1)
        iEnvio = Mid(Lista, InStr(Lista, "V") + 1, InStr(Lista, ":") - 1 - InStr(Lista, "V") - 1)
        If Mid(Lista, InStr(Lista, ":") - 1, 1) = "S" Then bHayEnvios = True Else bHayEnvios = False
        Lista = Trim(Mid(Lista, InStr(Lista, ":") + 1, Len(Lista)))
        
        
'        AccionImprimirFactura_VSReport iFactura, iEntrega, iEnvio
        
        'Si es con cheques no se imprime conforme
        'Si Hay Entrega Van los Envios al conforme si no no
        If iEntrega > 0 Then iEntrega = iEntrega + iEnvio
        If cPago.ItemData(cPago.ListIndex) = TipoPagoSolicitud.Efectivo Then
            If oCnfgPrint.Opcion = 0 Then
                AccionImprimirConforme iFactura, iEntrega, FormulasCo
            Else
                ImprimoConformeTickets iFactura
            End If
        End If
        
    Loop
    Screen.MousePointer = 0
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    If crMsgErr <> "" Then clsGeneral.OcurrioError crMsgErr Else clsGeneral.OcurrioError "No se pudo realizar la impresión. " & Trim(Err.Description)
End Sub

Private Function LlevaConforme(ByVal factura As Long) As Boolean
On Error GoTo errlC
    LlevaConforme = True
    Dim rsC As rdoResultset
    Cons = "Select * from Credito, TipoCuota " _
             & " Where CreFactura = " & factura _
             & " And CreTipoCuota = TCuCodigo"
     Set rsC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    '   6/6/2003        Valido Si el TipoDeCuota lleva Conforme ---------------------------------------
    If Not (rsC!TCuLlevaConforme) Then LlevaConforme = False
    rsC.Close
Exit Function
errlC:
End Function

Private Sub ImprimoConformeTickets(ByVal factura As Long)

    If Not LlevaConforme(factura) Then Exit Sub
    
    With vsrReport
        .Clear                  ' clear any existing fields
        .FontName = "Tahoma"    ' set default font for all controls
        .FontSize = 8
        
        .Load gPathListados & "Conforme.xml", "Conforme"
    
        .DataSource.ConnectionString = cBase.Connect
        .DataSource.RecordSource = "prg_Conformes_Impresion (" & factura & ")"
        
        vspPrinter.Device = oCnfgPrint.ImpresoraTickets
        
        .Render vspPrinter
        
    End With
    vspPrinter.PrintDoc False
    DoEvents
    'Hago una pausa ya que dicen que la impresora se tranca en un 2do conforme.
    Dim iQ As Integer, iNada As Integer
    For iQ = 0 To 30000
        iNada = iQ
    Next
    DoEvents

End Sub

Private Sub AccionImprimirFactura_VSReport(ByVal factura As Long, MEntrega As Currency, MEnvio As Currency)
Dim result As Integer, JobSRep1 As Integer, JobSRep2 As Integer
Dim NombreFormula As String, Cont As Integer
Dim RsAuxC As rdoResultset
Dim sTexto As String, sConCheques As Boolean
Dim idClienteCredito As Long 'Creo esto para solucionar los casos que imprime el documento de otro cliente

On Error GoTo ErrCrystal
Dim sErr As String
    sErr = "Paso 1"

    'Consulta para sacar los datos del credito------------------------------------------------------------
    Cons = " Select * from Credito, TipoCuota " _
             & " Where CreFactura = " & factura _
             & " And CreTipoCuota = TCuCodigo"
    Set RsAuxC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
     'VaCuota "" = Diferido, pago envio ...si hay, "E" = Entrega, "1...N" Cuota.
    If RsAuxC!CreFormaPago = TipoPagoSolicitud.ChequeDiferido Then sConCheques = True Else sConCheques = False
    idClienteCredito = RsAuxC("CreCliente")
    
    'Si es con Diferidos saco valor de las cuotas que no vencen el mismo dia que la factura ---------------------------
    Dim mTCheques As Currency
    mTCheques = 0
    If sConCheques Then
        sErr = "Paso 2"
        Cons = "Select * From CreditoCuota " & _
                    " Where CCuCredito = " & RsAuxC!CreCodigo
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            If Format(gFechaServidor, "yyyy/mm/dd") <> Format(RsAux!CCuVencimiento, "yyyy/mm/dd") Then
                mTCheques = mTCheques + RsAux!CCuValor
            End If
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    '--------------------------------------------------------------------------------------------------------------------------------
    sErr = "Paso 3"
    Dim oCliFactura As clsClienteCFE
    Cons = "Select * from Cliente, CPersona, CEmpresa, PaisDelDocumento" _
           & " " _
           & " Where CliCodigo = " & idClienteCredito _
           & " And CliCodigo *= CPeCliente " _
           & " And CliCodigo *= CEmCliente AND PDDId = CliPaisDelDocumento"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Set oCliFactura = CargoObjetoContacto()
    RsAux.Close
    
    sErr = "Paso 4"
    Dim oImprimo As clsImpresionCredito
    Set oImprimo = New clsImpresionCredito
    oImprimo.DondeImprimo.Bandeja = paICreditoB
    oImprimo.DondeImprimo.Impresora = paICreditoN
    oImprimo.DondeImprimo.Papel = 1
    oImprimo.PathReportes = gPathListados
    oImprimo.StringConnect = miConexion.TextoConexion("Comercio")
    
    sErr = "Paso 5"
    With oImprimo
        .field_ClienteCedula = Trim(clsGeneral.RetornoFormatoCedula(oCliFactura.CI))
        .field_ClienteDireccion = Trim(lDireccion.Caption)
        
        If tGarantia.Text <> "" Then
            .field_ClienteGarantia = tGarantia.Text & " " & Trim(lGarantia.Caption)
        Else
            .field_ClienteGarantia = ""
        End If
        If oDireccion.Visible And oDireccion.value = vbChecked Then
            .field_ClienteNombre = Trim(oCliFactura.NombreCliente) & " (" & Trim(cDireccion.Text) & ")"
        Else
            .field_ClienteNombre = Trim(oCliFactura.NombreCliente)
        End If

        
        sTexto = Trim(RsAuxC!TCuAbreviacion) & " - "
        If Not IsNull(RsAuxC!TCuVencimientoE) And IsNumeric(Label1.Tag) Then
            MEntrega = CCur(Label1.Tag) - (RsAuxC!TCuCantidad * RsAuxC!CreValorCuota)
            If MEntrega > 0 Then sTexto = sTexto & "Ent.: " & Format(MEntrega, FormatoMonedaP) & " "
        End If
        sTexto = sTexto & RsAuxC!TCuCantidad & " x " & Format(RsAuxC!CreValorCuota, FormatoMonedaP)
        .field_CreditoFinanciacion = sTexto
        
        If Not IsNull(RsAuxC!CreProximoVto) Then
            .field_CreditoProxVto = Format(RsAuxC!CreProximoVto, "d Mmm yyyy")
        Else
            .field_CreditoProxVto = ""
        End If
        
        .field_MonedaNombre = BuscoNombreMoneda(cMoneda.ItemData(cMoneda.ListIndex))
        .field_MonedaSimbolo = Trim(cMoneda.Text)
        
        .field_NombreDocumento = paDCredito
        .field_NombreRecibo = paDRecibo
        
        If Trim(RsAuxC!CreVaCuota) <> "" Then
            .field_ReciboInfoCuota = "Cuotas: " & Trim(RsAuxC!CreVaCuota) & " de " & Trim(RsAuxC!CreDeCuota)
        Else
            .field_ReciboInfoCuota = ""
        End If
        
        If Not sConCheques Then
            .field_ReciboImporteFlete = Format(MEnvio, FormatoMonedaP)
        Else
            .field_ReciboImporteFlete = Format(mTCheques, FormatoMonedaP)
        End If
        
        If Not sConCheques Then
            .field_ReciboInfoTFlete = "Flete:"
        Else
            .field_ReciboInfoTFlete = "Ch. diferidos total:"
        End If
        
        If Not IsNull(RsAuxC!CreProximoVto) Then
            .field_ReciboInfoVto = "Próximo vencimiento: " & Format(RsAuxC!CreProximoVto, "d Mmm yyyy")
        Else
            .field_ReciboInfoVto = ""
        End If
        
        If oCliFactura.RUT <> "" Then
            .field_RUT = clsGeneral.RetornoFormatoRuc(oCliFactura.RUT)
        Else
            .field_RUT = ""
        End If
        
        zGet_StringRetira sTexto, "", porFactura:=factura
        .field_TextoRetira = sTexto
        
        .field_Vendedor = lVendedor.Caption
        .field_Digitador = tUsuario.Text
        
        sErr = "Paso envío a componente"
        .ImprimoFacturaContado_VSReport factura
    End With
    Screen.MousePointer = 0
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al imprimir, paso: " & sErr, Err.Description, "Error al imprimir"
End Sub


Private Sub AccionImprimirFactura(ByVal factura As Long, MEntrega As Currency, MEnvio As Currency, Formulas As Integer, HayEnvios As Boolean)

Dim result As Integer, JobSRep1 As Integer, JobSRep2 As Integer
Dim NombreFormula As String, Cont As Integer
Dim RsAuxC As rdoResultset
Dim sTexto As String, sConCheques As Boolean
Dim idClienteCredito As Long 'Creo esto para solucionar los casos que imprime el documento de otro cliente

    'Consulta para sacar los datos del credito------------------------------------------------------------
    Cons = " Select * from Credito, TipoCuota " _
             & " Where CreFactura = " & factura _
             & " And CreTipoCuota = TCuCodigo"
    Set RsAuxC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
     'VaCuota "" = Diferido, pago envio ...si hay, "E" = Entrega, "1...N" Cuota.
    If RsAuxC!CreFormaPago = TipoPagoSolicitud.ChequeDiferido Then sConCheques = True Else sConCheques = False
    idClienteCredito = RsAuxC("CreCliente")
    
    'Si es con Diferidos saco valor de las cuotas que no vencen el mismo dia que la factura ---------------------------
    Dim mTCheques As Currency
    mTCheques = 0
    If sConCheques Then
        Cons = "Select * From CreditoCuota " & _
                    " Where CCuCredito = " & RsAuxC!CreCodigo
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        Do While Not RsAux.EOF
            If Format(gFechaServidor, "yyyy/mm/dd") <> Format(RsAux!CCuVencimiento, "yyyy/mm/dd") Then
                mTCheques = mTCheques + RsAux!CCuValor
            End If
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    '--------------------------------------------------------------------------------------------------------------------------------
    
    Dim oCliFactura As clsClienteCFE
    Cons = "Select * from Cliente, CPersona, CEmpresa, PaisDelDocumento" _
           & " " _
           & " Where CliCodigo = " & idClienteCredito _
           & " And CliCodigo *= CPeCliente " _
           & " And CliCodigo *= CEmCliente AND PDDId = CliPaisDelDocumento"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Set oCliFactura = CargoObjetoContacto()
    RsAux.Close
    
    'Cargo Propiedades para el reporte Credito --------------------------------------------------------------------------------
    For Cont = 0 To Formulas - 1
        NombreFormula = crObtengoNombreFormula(iJobCre, Cont)
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "nombredocumento": result = crSeteoFormula(iJobCre%, NombreFormula, "'" & paDCredito & "'")
            Case "cliente":
                If oDireccion.Visible And oDireccion.value = vbChecked Then
                    'Result = crSeteoFormula(iJobCre%, NombreFormula, "'" & Trim(lNombre.Caption) & " (" & Trim(cDireccion.Text) & ")" & "'")
                    result = crSeteoFormula(iJobCre%, NombreFormula, "'" & Trim(oCliFactura.NombreCliente) & " (" & Trim(cDireccion.Text) & ")" & "'")
                Else
                    result = crSeteoFormula(iJobCre%, NombreFormula, "'" & Trim(oCliFactura.NombreCliente) & "'")
                End If
                
            Case "cedula": result = crSeteoFormula(iJobCre%, NombreFormula, "'" & Trim(clsGeneral.RetornoFormatoCedula(oCliFactura.CI)) & "'")  'lCi.Caption
            Case "direccion": result = crSeteoFormula(iJobCre%, NombreFormula, "'" & Trim(lDireccion.Caption) & "'")
            
            Case "ruc":
'                If Trim(lRuc.Caption) <> "S/D" Then sTexto = Trim(clsGeneral.RetornoFormatoRuc(lRuc.Caption)) Else sTexto = ""
                If oCliFactura.RUT <> "" Then
                    sTexto = clsGeneral.RetornoFormatoRuc(oCliFactura.RUT)
                Else
                    sTexto = ""
                End If
                result = crSeteoFormula(iJobCre%, NombreFormula, "'" & sTexto & "'")
                
            Case "codigobarras": result = crSeteoFormula(iJobCre%, NombreFormula, "'" & CodigoDeBarras(TipoDocumento.Credito, factura) & "'")
            
            Case "signomoneda": result = crSeteoFormula(iJobCre%, NombreFormula, "'" & Trim(cMoneda.Text) & "'")
            Case "nombremoneda": result = crSeteoFormula(iJobCre%, NombreFormula, "'(" & BuscoNombreMoneda(cMoneda.ItemData(cMoneda.ListIndex)) & ")'")
            
            Case "usuario": result = crSeteoFormula(iJobCre%, NombreFormula, "'" & Trim(tUsuario.Text) & "'")
            Case "vendedor": result = crSeteoFormula(iJobCre%, NombreFormula, "'" & Trim(lVendedor.Caption) & "'")
            
            Case "garantia"
                If tGarantia.Text <> "" Then sTexto = tGarantia.Text & " " & Trim(lGarantia.Caption) Else sTexto = ""
                result = crSeteoFormula(iJobCre%, NombreFormula, "'" & sTexto & "'")
                
            Case "textoretira"
                
                zGet_StringRetira sTexto, "", porFactura:=factura
                'If Not HayEnvios Then
                '    sTexto = "RETIRA "
                '    If Format(tFRetiro.Text, "dd/mm/yyyy") = Format(gFechaServidor, "dd/mm/yyyy") Then sTexto = sTexto & "HOY" Else sTexto = sTexto & Format(tFRetiro.Text, "dd/mm/yyyy")
                'Else
                '    sTexto = "HAY ENVIOS DE MERCADERIA"
                'End If
                result = crSeteoFormula(iJobCre%, NombreFormula, "'" & sTexto & "'")
            
            Case "nombrerecibo": result = crSeteoFormula(iJobCre%, NombreFormula, "'" & paDRecibo & "'")
            
            '---------------------------------------------------------------------------------------------------------------------------------------------------------
            Case "recibotcuota"
                If Not sConCheques Then aTexto = "Cuotas:" Else aTexto = "Al DIA:"
                result = crSeteoFormula(iJobCre%, NombreFormula, "'" & aTexto & "'")
                
            Case "recibotflete"
                If Not sConCheques Then aTexto = "Flete:" Else aTexto = "Ch. Diferidos Total:"
                result = crSeteoFormula(iJobCre%, NombreFormula, "'" & aTexto & "'")
                
            Case "reciboflete":
                'Result = crSeteoFormula(iJobCre%, NombreFormula, "'" & Format(MEnvio, FormatoMonedaP) & "'")
                If Not sConCheques Then
                    sTexto = Format(MEnvio, FormatoMonedaP)
                Else
                    sTexto = Format(mTCheques, FormatoMonedaP)
                End If
                result = crSeteoFormula(iJobCre%, NombreFormula, "'" & sTexto & "'")
            
            Case "recibocuota"
                If Not IsNull(RsAuxC!CreVaCuota) Then
                    If Trim(RsAuxC!CreVaCuota) <> "" Then
                        sTexto = Trim(RsAuxC!CreVaCuota) & " de " & Trim(RsAuxC!CreDeCuota)
                        result = crSeteoFormula(iJobCre%, NombreFormula, "'" & sTexto & "'")
                    End If
                End If
            '---------------------------------------------------------------------------------------------------------------------------------------------------------
                
            Case "financiacion"
                sTexto = Trim(RsAuxC!TCuAbreviacion) & " - "
                If MEntrega > 0 Then sTexto = sTexto & "Ent.: " & Format(MEntrega, FormatoMonedaP) & " "
                sTexto = sTexto & RsAuxC!TCuCantidad & " x " & Format(RsAuxC!CreValorCuota, FormatoMonedaP)
                result = crSeteoFormula(iJobCre%, NombreFormula, "'" & sTexto & "'")
                    
            Case "proximovto"
                If Not IsNull(RsAuxC!CreProximoVto) Then sTexto = Format(RsAuxC!CreProximoVto, "d Mmm yyyy") Else sTexto = ""
                result = crSeteoFormula(iJobCre%, NombreFormula, "'" & sTexto & "'")
                
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    RsAuxC.Close
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Documento.DocCodigo, Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal," _
            & " Documento.DocIVA, Documento.DocVendedor, Recibo.DocSerie , Recibo.DocNumero, Recibo.DocTotal, Documento.DocCofis" _
            & " From " _
            & " { oj (" & paBD & ".dbo.Documento Documento " _
                        & " LEFT OUTER JOIN " & paBD & ".dbo.DocumentoPago DocumentoPago ON  Documento.DocCodigo = DocumentoPago.DPaDocASaldar)" _
                        & " LEFT OUTER JOIN " & paBD & ".dbo.Documento Recibo ON  DocumentoPago.DPaDocQSalda = Recibo.DocCodigo}" _
            & " Where Documento.DocCodigo = " & factura

    If sConCheques Then
        Cons = Cons & " Group by Documento.DocCodigo, Documento.DocFecha, Documento.DocSerie, Documento.DocNumero, Documento.DocTotal, Documento.DocIVA, Documento.DocVendedor, Recibo.DocSerie , Recibo.DocNumero, Recibo.DocTotal, Documento.DocComentario, Documento.DocCofis "
    End If
    
    If crSeteoSqlQuery(iJobCre%, Cons) = 0 Then GoTo ErrCrystal
        
    'Subreporte srCredito.rpt  y srCredito.rpt - 01-----------------------------------------------------------------------------
    JobSRep1 = crAbroSubreporte(iJobCre, "srCredito.rpt")
    If JobSRep1 = 0 Then GoTo ErrCrystal
    
    Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion," _
            & " From { oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
                           & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId}"

'    Cons = "SELECT Renglon.RenDocumento, Renglon.RenCantidad, Renglon.RenPrecio, Renglon.RenDescripcion" _
'      & " From ({ oj " & paBD & ".dbo.Renglon Renglon INNER JOIN " _
'                     & paBD & ".dbo.Articulo Articulo ON Renglon.RenArticulo = Articulo.ArtId} Left Outer Join " _
'                     & paBD & ".dbo.ArticuloEspecifico On AEsTipoDocumento = 1 And AEsDocumento = RenDocumento And AEsArticulo = RenArticulo)"

    If crSeteoSqlQuery(JobSRep1, Cons) = 0 Then GoTo ErrCrystal

    JobSRep2 = crAbroSubreporte(iJobCre, "srCredito.rpt - 01")
    If JobSRep2 = 0 Then GoTo ErrCrystal
    If crSeteoSqlQuery(JobSRep2, Cons) = 0 Then GoTo ErrCrystal
    '-------------------------------------------------------------------------------------------------------------------------------------
    
    'If crMandoAPantalla(iJobCre, "Factura Credito") = 0 Then GoTo ErrCrystal
    If crMandoAImpresora(iJobCre, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(iJobCre, True, False) Then GoTo ErrCrystal
    
    If Not crCierroSubReporte(JobSRep1) Then GoTo ErrCrystal
    If Not crCierroSubReporte(JobSRep2) Then GoTo ErrCrystal
        
'    crEsperoCierreReportePantalla
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr
    On Error Resume Next
    Screen.MousePointer = 11
    crCierroSubReporte JobSRep1
    crCierroSubReporte JobSRep2
End Sub

Private Sub AccionImprimirConforme(factura As Long, MEntrega As Currency, Formulas As Integer)

'   Caso de Entregas:   Va la entrega Con Envio (no importa).
'   Caso Sin Entrega:   Va el valor de la cuota (se descarta el envio).     15/jul (Con Carlos)

Dim result As Integer
Dim NombreFormula As String, Cont As Integer
Dim sTexto As String, sMoneda As String, sAux As String
Dim RsAuxC As rdoResultset

    'Consulta para sacar los datos del credito------------------------------------------------------------
     Cons = "Select * from Credito, TipoCuota " _
             & " Where CreFactura = " & factura _
             & " And CreTipoCuota = TCuCodigo"
     Set RsAuxC = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
     'VaCuota "" = Diferido, pago envio ...si hay, "E" = Entrega, "1...N" Cuota.
    
    '   6/6/2003        Valido Si el TipoDeCuota lleva Conforme ---------------------------------------
    If Not (RsAuxC!TCuLlevaConforme) Then
        RsAuxC.Close
        Exit Sub
    End If
    '----------------------------------------------------------------------------------------------------------
    
    sMoneda = LCase(BuscoNombreMoneda(cMoneda.ItemData(cMoneda.ListIndex)))
    
    'Cargo Propiedades para el reporte Credito --------------------------------------------------------------------------------
    For Cont = 0 To Formulas - 1
        NombreFormula = crObtengoNombreFormula(iJobCon, Cont)
        
        Select Case LCase(NombreFormula)
            Case "": GoTo ErrCrystal
            Case "sucursal": result = crSeteoFormula(iJobCon%, NombreFormula, "'SUCURSAL " & UCase(paNombreSucursal) & "'")
            Case "titular"
                sTexto = ""
                If Trim(lCi.Caption) <> "S/D" Then sTexto = Trim(lCi.Caption) Else If Trim(lRuc.Caption) <> "S/D" Then sTexto = Trim(lRuc.Caption)
                sTexto = Trim(sTexto & " " & Trim(lNombre.Caption))
                result = crSeteoFormula(iJobCon%, NombreFormula, "'" & sTexto & "'")
            
            Case "garantia"
                sTexto = ""
                If Trim(lGarantia.Caption) <> "" Then sTexto = Trim(tGarantia.Text) & " " & Trim(lGarantia.Caption)
                result = crSeteoFormula(iJobCon%, NombreFormula, "'" & sTexto & "'")
            
            Case "texto1"       'Texto del Conforme Totoal $
                sAux = ImporteATexto((RsAuxC!CreValorCuota * RsAuxC!TCuCantidad) + MEntrega)
                sTexto = " por la suma de " & sMoneda & " " & UCase(sAux) & ", "
                sTexto = sTexto & "que pagaremos en forma indivisible y solidariamente a CARLOS GUTIERREZ S.A. o a su orden en "
                
                result = crSeteoFormula(iJobCon%, NombreFormula, "'" & sTexto & "'")
                
            Case "texto2"       'Texto del Conforme Entrega
                sTexto = ""
                If Not IsNull(RsAuxC!TCuVencimientoE) Then
                    sTexto = sTexto & "una entrega de " & sMoneda & " " & UCase(ImporteATexto(MEntrega)) & " con vencimiento el "
                    'Vencimiento de entrega
                    sAux = gFechaServidor + RsAuxC!TCuVencimientoE
                    sAux = " " & Format(sAux, "d") & " de " & Format(sAux, "Mmmm") & " de " & Format(sAux, "yyyy")
                    sTexto = sTexto & sAux & " y "
                End If
                
                result = crSeteoFormula(iJobCon%, NombreFormula, "'" & sTexto & "'")
            
            Case "texto3"       'Texto del Conforme Cuotas
                sAux = ImporteATexto(RsAuxC!CreValorCuota)
                'sTexto = RsAuxC!TCuCantidad & " cuotas consecutivas de " & sMoneda & " " & UCase(sAux) _
                            & " exigibles cada " & RsAuxC!TCuDistancia & " días a partir del "
                sTexto = RsAuxC!TCuCantidad & " cuotas consecutivas de " & sMoneda & " " & UCase(sAux) _
                            & " exigibles cada " & RsAuxC!TCuDistancia & " días, venciendo la primera el "
                            
                'Primer Vencimiento de cuotas
                sAux = gFechaServidor + RsAuxC!TCuVencimientoC
                sAux = Format(sAux, "d") & " de " & Format(sAux, "Mmmm") & " de " & Format(sAux, "yyyy")
                sTexto = sTexto & " " & sAux
                                
                result = crSeteoFormula(iJobCon%, NombreFormula, "'" & sTexto & "'")
                                      
            Case "fecha"
                sTexto = "Montevideo, " & Format(gFechaServidor, "d") & " de " & Format(gFechaServidor, "Mmmm") & " de " & Format(gFechaServidor, "yyyy")
                result = crSeteoFormula(iJobCon%, NombreFormula, "'" & sTexto & "'")
                        
            Case Else: result = 1
        End Select
        If result = 0 Then GoTo ErrCrystal
    Next
    '--------------------------------------------------------------------------------------------------------------------------------------------
    RsAuxC.Close
    
    'Seteo la Query del reporte-----------------------------------------------------------------
    Cons = "SELECT Documento.DocSerie, Documento.DocNumero " _
            & " From " & paBD & ".dbo.Documento Documento " _
            & " Where Documento.DocCodigo = " & factura

    If crSeteoSqlQuery(iJobCon%, Cons) = 0 Then GoTo ErrCrystal
   
    'If crMandoAPantalla(iJobCon, "Conforme") = 0 Then GoTo ErrCrystal
    If crMandoAImpresora(iJobCon, 1) = 0 Then GoTo ErrCrystal
    If Not crInicioImpresion(iJobCon, True, False) Then GoTo ErrCrystal
        
    'crEsperoCierreReportePantalla
    Exit Sub

ErrCrystal:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError crMsgErr, Err.Description
    Screen.MousePointer = 11
End Sub

Private Sub CargoDireccionesAuxiliares(aIdCliente As Long)

    On Error GoTo errCDA
    Dim rsDA As rdoResultset
    
    'Direcciones Auxiliares-----------------------------------------------------------------------
    Cons = "Select * from DireccionAuxiliar Where DAuCliente = " & aIdCliente & " Order by DAuNombre"
    Set rsDA = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsDA.EOF Then
        Do While Not rsDA.EOF
            cDireccion.AddItem Trim(rsDA!DAuNombre)
            cDireccion.ItemData(cDireccion.NewIndex) = rsDA!DAuDireccion
            If rsDA!DAuFactura Then gDirFactura = rsDA!DAuDireccion
            rsDA.MoveNext
        Loop
        
        If cDireccion.ListCount > 1 Then cDireccion.BackColor = vbWhite
    End If
    rsDA.Close
    
    If Val(cDireccion.Tag) = 0 And cDireccion.ListCount > 0 And gDirFactura = 0 Then
        cDireccion.Text = cDireccion.List(0)
    Else
        If gDirFactura <> 0 Then BuscoCodigoEnCombo cDireccion, gDirFactura
    End If
    
    Dim bVisible As Boolean     'Cdo. no se activa el form siempre queda en falso
    If cDireccion.ListCount > 1 Then bVisible = True Else bVisible = False
    cDireccion.Visible = bVisible: cDireccion.Refresh: oDireccion.Visible = bVisible
    
    If bVisible Then
        lDireccion.Left = cDireccion.Left + cDireccion.Width + 40
    Else
        lDireccion.Left = lDireccionN.Left + lDireccionN.Width + 40
    End If
    
errCDA:
End Sub

Private Function BuscoZonaDireccion(lngIDDir As Long) As Long

    On Error GoTo errBZ
    Dim rsZona As rdoResultset, strCons As String
    
    BuscoZonaDireccion = 0
    
    strCons = "Select CZoZona from Direccion, CalleZona" _
            & " Where DirCodigo = " & lngIDDir _
            & " And CZoCalle = DirCalle " _
            & " And CZoDesde <= DirPuerta " & " And CZoHasta >= DirPuerta"
    Set rsZona = cBase.OpenResultset(strCons, rdOpenDynamic, rdConcurValues)
    
    If Not rsZona.EOF Then
        If Not IsNull(rsZona(0)) Then BuscoZonaDireccion = rsZona(0)
    End If
    rsZona.Close

errBZ:
End Function

Private Function BuscoPreciosContado() As Boolean

    On Error GoTo errCtdos
    BuscoPreciosContado = False
    
    Screen.MousePointer = 11
    Dim iAuxiliar As Currency
    'Busco los precios contado de los articulos para dps calcular el cofis
    For Each itmX In lvVenta.ListItems
        If Trim(itmX.SubItems(13)) = "" Then     'Precio Contado
            'Saco el Precio contado del Articulo--------------------------------------------------------------------------
            Cons = "Select PViPrecio, PViPlan From PrecioVigente" _
                    & " Where PViArticulo = " & ArticuloDeLaClave(itmX.Key) _
                    & " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                    & " And PViTipoCuota = " & paTipoCuotaContado & " And PViHabilitado = 1"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not RsAux.EOF Then
                itmX.SubItems(13) = Format(RsAux!PViPrecio, FormatoMonedaP)
                RsAux.Close
            Else
                RsAux.Close
                'Si no lo tengo lo saco por cuentas
                '(ValorCuota * CantCuotas) * Coeficiente = 1xx% ---> contado = 100 + Entrega
                '1)  Saco el Coeficiente del plan por defecto
                Cons = "Select * from Coeficiente, TipoCuota" _
                        & " Where CoePlan = " & paPlanPorDefecto _
                        & " And CoeTipoCuota = " & PlanDeLaClave(itmX.Key) _
                        & " And CoeMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                        & " And CoeTipocuota = TCuCodigo"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not RsAux.EOF Then
                    iAuxiliar = CCur(itmX.SubItems(5)) * RsAux!TCuCantidad * 100 / (RsAux!CoeCoeficiente * 100)
                End If
                itmX.SubItems(13) = Redondeo(iAuxiliar, mRedondeo)            'Contado
                RsAux.Close
            End If
        End If
    Next
    
    BuscoPreciosContado = True
    Screen.MousePointer = 0
    Exit Function

errCtdos:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al buscar los precios contado (COFIS).", Err.Description
End Function

Private Function fnc_EsDelTipoServicio(idArticulo As Long) As Boolean
    fnc_EsDelTipoServicio = False
    Dim miRs As rdoResultset
    Cons = "Select ArtTipo from Articulo Where ArtID = " & idArticulo
    Set miRs = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    fnc_EsDelTipoServicio = EsTipoDeServicio(miRs("ArtTipo"))
    miRs.Close
End Function

Private Function ArticuloDeCombo(idArticulo As Long, idPlan As Long) As Boolean
On Error GoTo errCmbo
Dim miRs As rdoResultset, rsVal As rdoResultset
Dim miSQL As String
    
Dim bEsta As Boolean

Dim miIX As ListItem

    ArticuloDeCombo = False
    
    miSQL = "Select * from PresupuestoArticulo Where PArArticulo = " & idArticulo
    Set miRs = cBase.OpenResultset(miSQL, rdOpenDynamic, rdConcurValues)
    Do While Not miRs.EOF
        bEsta = False
        miSQL = "Select * from PresupuestoArticulo Where PArPresupuesto = " & miRs!PArPresupuesto
        Set rsVal = cBase.OpenResultset(miSQL, rdOpenDynamic, rdConcurValues)
        Do While Not rsVal.EOF
            bEsta = False
            For Each miIX In lvVenta.ListItems
                If ArticuloDeLaClave(miIX.Key) = rsVal!PArArticulo And idPlan = PlanDeLaClave(miIX.Key) Then
                    bEsta = True: Exit For
                End If
            Next
            
            If Not bEsta Then Exit Do
            rsVal.MoveNext
        Loop
        rsVal.Close
        
        If bEsta Then
            ArticuloDeCombo = True
            Exit Do
        End If
        
        miRs.MoveNext
    Loop
    miRs.Close

errCmbo:
End Function


Function BuscoNombreMoneda(Codigo As Long) As String

    On Error GoTo ErrBU
    Dim Rs As rdoResultset
    BuscoNombreMoneda = ""

    Cons = "SELECT * FROM Moneda WHERE MonCodigo = " & Codigo
    Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not Rs.EOF Then BuscoNombreMoneda = Trim(Rs!MonNombre)
    Rs.Close
    Exit Function
    
ErrBU:
End Function


Private Function ValidoVigenciaPrecios(mIDSolicitud As Long) As Boolean
On Error GoTo errVig

    ValidoVigenciaPrecios = True
    Dim mFSol As String
    mFSol = Format(CDate(lFecha.Caption), "yyyy/mm/dd")
    If mFSol = Format(Now, "yyyy/mm/dd") Then Exit Function
    
    Cons = "Select SolFecha, Max(PViVigencia) as SolVigencia" & _
            " From Solicitud, RenglonSolicitud, PrecioVigente " & _
            " Where RSoSolicitud = " & mIDSolicitud & _
            " And SolCodigo = RSoSolicitud" & _
            " And RSoArticulo = PViArticulo " & _
            " And PViTipoCuota = " & paTipoCuotaContado & _
            " And PViMoneda = SolMoneda " & _
            " Group by SolFecha"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux!SolVigencia > RsAux!SolFecha Then ValidoVigenciaPrecios = False
    End If
    RsAux.Close
    
    Exit Function
errVig:
End Function

Private Function zGet_InstaladorArticulo(mIDArticulo As Long) As Long
On Error GoTo errFnc
Dim rsIns As rdoResultset
    
    zGet_InstaladorArticulo = 0
    
    Cons = "Select * From Articulo " & _
               " Where ArtID = " & mIDArticulo & _
               " And ArtInstalador > 0 "
               
    Set rsIns = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsIns.EOF Then zGet_InstaladorArticulo = rsIns!ArtInstalador
    rsIns.Close

errFnc:
End Function

Private Sub zGet_StringRetira(sTextoImp As String, sFechaP As String, Optional porPlan As Long = 0, Optional porFactura As Long = 0)
On Error GoTo errFnc

Dim iQitem As Currency
Dim iQEnv As Currency
Dim iQInst As Currency

    sTextoImp = ""
    sFechaP = tFRetiro.Text
        
    'Si hay envio o hay instal. evitar poner palabra retira hoy.
    If porPlan = 0 Then
        For arrIdx = LBound(arrCreditos) To UBound(arrCreditos)
            If arrCreditos(arrIdx).Credito.Codigo = porFactura Then
                porPlan = arrCreditos(arrIdx).idPlan
                Exit For
            End If
        Next
    End If
    
    If porPlan = 0 Then Exit Sub
    
    'Busco cantidades.
    zGet_QItmesEnvioInstala porPlan, iQitem, iQEnv, iQInst
    
    If iQEnv > 0 Then
        'Tengo envíos
        If iQEnv = iQitem Then
            sTextoImp = IIf(iQEnv = 1, "Se Envía", "Se Envía Todo")
            sFechaP = DateAdd("d", 1, Date) & " 09:00:00"
        Else
            If iQInst = 0 Then
                'No tengo instalación.
                sTextoImp = "Se Envían " & iQEnv & " de " & iQitem
                sFechaP = tFRetiro.Text
                If Trim(tFRetiro.Tag) <> "" Then sFechaP = sFechaP & " 23:00:00"
            Else
                'Tengo instalaciones
                sTextoImp = "Se Envían " & iQEnv & " de " & iQitem & ". V.Inst."
                'Unifico generico mañana a las 10
                sFechaP = DateAdd("d", 1, Date) & " 10:00:00"
            End If
        End If
    Else
        If iQInst > 0 Then
            If iQInst <> iQitem Then
                sTextoImp = "Retira " & IIf(CDate(tFRetiro.Text) = Date, "Hoy", tFRetiro.Text) & ". V.Inst."
                sFechaP = tFRetiro.Text
                If Trim(tFRetiro.Tag) <> "" Then sFechaP = sFechaP & " 23:00:00"
            Else
                sTextoImp = "Ver Instalación."
                sFechaP = tFRetiro.Text
                If CDate(sFechaP) <> Date Then
                    sFechaP = sFechaP & " 23:00:00"
                Else
                    sFechaP = DateAdd("d", 1, Date) & " 10:00:00"
                End If
            End If
        Else
            'Caso todo a retirar.
            sTextoImp = "Retira " & IIf(CDate(tFRetiro.Text) = Date, "Hoy", tFRetiro.Text)
            sFechaP = tFRetiro.Text
            If Trim(tFRetiro.Tag) <> "" Then sFechaP = sFechaP & " 23:00:00"
        End If
    End If

errFnc:
End Sub

Private Sub zGet_QItmesEnvioInstala(iIDPlan As Long, retItems As Currency, retEnvia As Currency, retInstala As Currency)

On Error GoTo errGetQ
'Recorro la lista y retorno el total de artículos y el total que se envía

Dim iQEnvia As Currency

    retItems = 0: retEnvia = 0: retInstala = 0
    
    Dim item_QV As ListItem
    
    For Each item_QV In lvVenta.ListItems
        If iIDPlan = PlanDeLaClave(item_QV.Key) Then
        
            If Not fnc_EsDelTipoServicio(ArticuloDeLaClave(item_QV.Key)) Then
                If Not InStr(aFletes, ArticuloDeLaClave(item_QV.Key) & ",") <> 0 Then
                    
                    retItems = retItems + CCur(item_QV.SubItems(1))
                    retEnvia = retEnvia + CCur(item_QV.SubItems(15))
                    
                    If Val(item_QV.SubItems(14)) > 0 Then retInstala = retInstala + (CCur(item_QV.SubItems(1)) - CCur(item_QV.SubItems(15)))
                    
                End If
            End If
        End If
    Next
    
errGetQ:
End Sub

Private Function fnc_ValidoCantidadesAEnviar()
On Error GoTo errValidoQ

Dim item_QV As ListItem
Dim mIDArticulo As Long, mQAEnviar As Currency

Dim xIDArticulos As String
    
    prmArticulosSINCofis = "": xIDArticulos = ""
    
    For Each item_QV In lvVenta.ListItems
        mQAEnviar = 0
        mIDArticulo = ArticuloDeLaClave(item_QV.Key)
        xIDArticulos = xIDArticulos & IIf(xIDArticulos = "", "", ",") & mIDArticulo
        
        If Trim(item_QV.SubItems(11)) <> "" Then
        
            If Not fnc_EsDelTipoServicio(ArticuloDeLaClave(item_QV.Key)) Then
                If Not InStr(aFletes, ArticuloDeLaClave(item_QV.Key) & ",") <> 0 Then
            
                    Cons = "Select Sum(RevCantidad) from RenglonEnvio" & _
                            " Where RevArticulo = " & mIDArticulo & _
                            " And RevEnvio IN (" & item_QV.SubItems(11) & ")"
                
                    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If Not RsAux.EOF Then
                        If Not IsNull(RsAux(0)) Then mQAEnviar = RsAux(0)
                    End If
                    RsAux.Close
                    
                End If
            End If
        End If
        
        item_QV.SubItems(15) = mQAEnviar
    Next
    
    If xIDArticulos <> "" Then
        '2) Para Todos los artículos que facturo veo cuales NO llevan Cofis
        Cons = "Select ArtID from Articulo Where ArtID IN (" & xIDArticulos & ")" & " And ArtTipo IN (" & prmTipoArtSinCofis & ")"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            prmArticulosSINCofis = prmArticulosSINCofis & IIf(prmArticulosSINCofis = "", "", ",") & RsAux!ArtID
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    
    
    Exit Function
    
errValidoQ:
    clsGeneral.OcurrioError "Error al validar las cantidades a enviar. " & vbCrLf & "El texto retira pude no ser el correcto.", Err.Description
End Function

Sub SituacionRetiraAqui()
Dim lRet As String
    lRet = GetSetting("Ventas", "Config", "RetiraAqui", "0")
    If Val(lRet) = 1 Then
        lblUsuario.Top = 5820
        tUsuario.Top = 5760
    Else
        lblUsuario.Top = 5535
        tUsuario.Top = 5475
    End If
    chkRetiraAqui.Visible = (Val(lRet) = 1)
End Sub

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
    Open "\\ibm3200\oyr\efactura\logEFactura.txt" For Append As #1
    Print #1, Now & Space(5) & "Terminal: " & miConexion.NombreTerminal & Space(5); "CRÉDITO" & Space(5) & Texto
    Close #1
    Exit Sub
errEAL:
End Sub
