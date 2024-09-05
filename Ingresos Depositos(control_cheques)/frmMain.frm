VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F93D243E-5C15-11D5-A90D-000021860458}#10.0#0"; "orFecha.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H8000000B&
   Caption         =   "Listado - Control de Cheques"
   ClientHeight    =   6960
   ClientLeft      =   1965
   ClientTop       =   1800
   ClientWidth     =   10665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   10665
   Begin VSFlex6DAOCtl.vsFlexGrid vsRebotes 
      Height          =   3255
      Left            =   1200
      TabIndex        =   25
      Top             =   2880
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5741
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
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   12
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
      OutlineBar      =   1
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
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   3495
      Left            =   4320
      TabIndex        =   21
      Top             =   2400
      Width           =   5655
      _Version        =   196608
      _ExtentX        =   9975
      _ExtentY        =   6165
      _StockProps     =   229
      BorderStyle     =   1
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsDepositos 
      Height          =   3255
      Left            =   5520
      TabIndex        =   24
      Top             =   3300
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5741
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
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   12
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
      OutlineBar      =   1
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   6705
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   18283
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   420
      ScaleHeight     =   435
      ScaleWidth      =   7275
      TabIndex        =   5
      Top             =   6060
      Width           =   7335
      Begin VB.CommandButton bAyuda 
         Height          =   310
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Ayuda."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmMain.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmMain.frx":067C
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmMain.frx":097E
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmMain.frx":0CC0
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4560
         Picture         =   "frmMain.frx":0FC2
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3480
         Picture         =   "frmMain.frx":1388
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmMain.frx":148A
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2340
         Picture         =   "frmMain.frx":16C4
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2700
         Picture         =   "frmMain.frx":17AE
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   3840
         Picture         =   "frmMain.frx":1898
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmMain.frx":1D12
         Height          =   310
         Left            =   4200
         Picture         =   "frmMain.frx":1E5C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6120
         TabIndex        =   18
         Top             =   120
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   476
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame frame1 
      Caption         =   "Filtros de Consulta"
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   10035
      Begin orctFecha.orFecha tDesde 
         Height          =   285
         Left            =   780
         TabIndex        =   1
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
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
         Object.Width           =   1095
         EnabledMes      =   -1  'True
         EnabledAño      =   -1  'True
         EnabledPrimerUltimoDia=   -1  'True
         FechaFormato    =   "dd/mm/yyyy"
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   2820
         TabIndex        =   3
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   180
         TabIndex        =   0
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Moneda:"
         Height          =   255
         Left            =   2100
         TabIndex        =   2
         Top             =   285
         Width           =   795
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsIngresos 
      Height          =   3255
      Left            =   3120
      TabIndex        =   19
      Top             =   1860
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5741
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
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   4
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   12
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
      OutlineBar      =   1
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
   Begin ComctlLib.TabStrip tabOpciones 
      Height          =   4455
      Left            =   60
      TabIndex        =   23
      Top             =   840
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7858
      TabWidthStyle   =   2
      TabFixedWidth   =   3175
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Cheques Ingresados"
            Key             =   "ingresos"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Depósitos del día"
            Key             =   "depositos"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cheques &Rebotados"
            Key             =   "rebotes"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList img1 
      Left            =   8580
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":238E
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":26A8
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuAcciones 
      Caption         =   "MnuAcciones"
      Visible         =   0   'False
      Begin VB.Menu MnuAccTitulo 
         Caption         =   "Ir a ..."
      End
      Begin VB.Menu MnuAccL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSeguimiento 
         Caption         =   "Seguimiento de Cheques"
      End
      Begin VB.Menu MnuDetalleF 
         Caption         =   "Detalle del Documento"
      End
      Begin VB.Menu MnuDeudaCheques 
         Caption         =   "Deuda en Cheques"
      End
      Begin VB.Menu MnuAccL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVOpeCCheque 
         Caption         =   "Operaciones Cliente Cheque"
      End
      Begin VB.Menu MnuVOpeCDoc 
         Caption         =   "Operaciones Cliente Doc."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim aValor As Long
Dim mTexto As String

Private Sub bAyuda_Click()
On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    Cons = "Select * from Aplicacion Where AplNombre = '" & prmKeyApp & "'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux!AplHelp) Then aFile = Trim(RsAux!AplHelp)
    RsAux.Close
    
    If aFile <> "" Then EjecutarApp aFile
    
    Screen.MousePointer = 0
    Exit Sub
    
errHelp:
    clsGeneral.OcurrioError "Error al activar el archivo de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bConsultar_Click()

    'Valido el ingreso de los campos para realizar la consulta.------------------------------------------
    If tDesde.FechaValor = "" Then
        MsgBox "La fecha 'desde' no es correcta. Verifique", vbExclamation, "Posible Error"
        Foco tDesde: Exit Sub
    End If
    
    AccionConsultar
    If vsIngresos.Rows > 1 Then vsIngresos.SetFocus
    
End Sub

Private Sub cMoneda_GotFocus()
    cMoneda.SelStart = 0: cMoneda.SelLength = Len(cMoneda.Text)
    Status.Panels(1).Text = "Moneda a filtrar."
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub Form_Activate()
       Screen.MousePointer = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyE: AccionConsultar
            Case vbKeyQ: AccionLimpiar
            Case vbKeyI: AccionImprimir
            
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    ObtengoSeteoForm Me, , , 10800
    
    Screen.MousePointer = 11
    InicializoGrilla
    PropiedadesImpresion
    
    AccionLimpiar
    
    'Cargo las monedas en el combo-------------------------
    Cons = "Select MonCodigo, MonNombre from Moneda Order by MonNombre"
    CargoCombo Cons, cMoneda
    BuscoCodigoEnCombo cMoneda, prmMonedaPesos
    
    picBotones.BorderStyle = 0
    vsIngresos.ZOrder 0
    vsListado.Visible = False
    vsDepositos.Visible = False
    
    bCancelar.Picture = img1.ListImages("salir").ExtractIcon
    bAyuda.Picture = img1.ListImages("help").ExtractIcon
   
End Sub


Private Sub AccionConsultar()

    If Not IsDate(tDesde.FechaValor) Then
        MsgBox "Debe ingresar la fecha para realizar la consulta de datos.", vbExclamation, "Falta filtro Fecha"
        tDesde.SetFocus: Exit Sub
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda para realizar la consulta de datos.", vbExclamation, "Falta filtro Moneda"
        cMoneda.SetFocus: Exit Sub
    End If
    
    Select Case LCase(tabOpciones.SelectedItem.Key)
        Case "ingresos": ConsultoIngresados
        Case "depositos": ConsultoDepositados
        Case "rebotes": ConsultoRebotes
    End Select
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Frame1
        .Left = 60: .Top = 60
        .Width = Me.ScaleWidth - (.Left * 2)
    End With
    
    With tabOpciones
        .Width = Frame1.Width
        .Left = Frame1.Left
        .Top = Frame1.Top + Frame1.Height + 80
        .Height = Me.ScaleHeight - .Top - picBotones.Height - Status.Height
    End With
    
    With vsIngresos
        .Left = tabOpciones.ClientLeft
        .Width = tabOpciones.ClientWidth
        .Top = tabOpciones.ClientTop
        .Height = tabOpciones.ClientHeight
    End With
    
    With vsRebotes
        .Left = vsIngresos.Left: .Top = vsIngresos.Top
        .Width = vsIngresos.Width: .Height = vsIngresos.Height
    End With
    
    If LCase(tabOpciones.SelectedItem.Key) = "depositos" Then zSizeListado Else zSizeListado (True)
    
    With picBotones
        .Top = Me.ScaleHeight - .Height - Status.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 100
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
End Sub

Private Sub Label1_Click()
    Foco tDesde
End Sub

Private Sub Label2_Click()
    Foco cMoneda
End Sub

Private Sub MnuDetalleF_Click()
    EjecutarApp prmPathApp & "Detalle de Factura.exe ", MnuDetalleF.Tag
End Sub

Private Sub MnuDeudaCheques_Click()
    EjecutarApp prmPathApp & "Deuda en Cheques.exe ", MnuDeudaCheques.Tag
End Sub

Private Sub MnuSeguimiento_Click()
    EjecutarApp prmPathApp & "SeguimientoCheques.exe ", MnuSeguimiento.Tag
End Sub

Private Sub MnuVOpeCCheque_Click()
    EjecutarApp prmPathApp & "Visualizacion de Operaciones.exe ", MnuVOpeCCheque.Tag
End Sub

Private Sub MnuVOpeCDoc_Click()
    EjecutarApp prmPathApp & "Visualizacion de Operaciones.exe ", MnuVOpeCDoc.Tag
End Sub

Private Sub tabOpciones_Click()
    
    On Error Resume Next
    
    Select Case tabOpciones.SelectedItem.Key
        Case "ingresos"
            vsIngresos.ZOrder 0
            vsListado.Visible = False
            
        Case "depositos":
            Screen.MousePointer = 11
            vsListado.StartDoc: vsListado.EndDoc
            zSizeListado
            vsListado.Visible = True: vsListado.ZOrder 0
            Screen.MousePointer = 0
            
        Case "rebotes"
            vsRebotes.ZOrder 0
            vsListado.Visible = False
    End Select
    
End Sub

Private Sub zSizeListado(Optional bGlobal As Boolean = False)
On Error Resume Next
    With vsListado
        If bGlobal Then
            .Width = tabOpciones.Width: .Left = tabOpciones.Left
            .Top = tabOpciones.Top: .Height = tabOpciones.Height
        Else
            .Left = tabOpciones.ClientLeft
            .Width = tabOpciones.ClientWidth
            .Top = tabOpciones.ClientTop
            .Height = tabOpciones.ClientHeight
        End If
    
    End With
End Sub

Private Sub tDesde_GotFocus()
    tDesde.SelStart = 0: tDesde.SelLength = Len(tDesde.FechaText)
    Status.Panels(1).Text = "Ingrese el rango de fechas para realizar la consulta."
End Sub

Private Sub tDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And tDesde.FechaValor <> "" Then cMoneda.SetFocus
End Sub

Private Sub tDesde_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConfigurar_Click()
    AccionConfigurar
End Sub

Private Sub bImprimir_Click()
    AccionImprimir True
End Sub

Private Sub bNoFiltros_Click()
    AccionLimpiar
End Sub

Private Sub bPrimero_Click()
    IrAPagina vsListado, 1
End Sub

Private Sub bSiguiente_Click()
    IrAPagina vsListado, vsListado.PreviewPage + 1
End Sub

Private Sub bUltima_Click()
    IrAPagina vsListado, vsListado.PageCount
End Sub

Private Sub bZMas_Click()
    Zoom vsListado, vsListado.Zoom + 5
End Sub

Private Sub bZMenos_Click()
    Zoom vsListado, vsListado.Zoom - 5
End Sub


Private Sub chVista_Click()

    Select Case LCase(tabOpciones.SelectedItem.Key)
        Case "ingresos"
            If chVista.Value = 0 Then
                vsIngresos.ZOrder 0: vsListado.Visible = False
                
            Else
                AccionImprimir
                vsListado.ZOrder 0: vsListado.Visible = True
            End If
        
        Case "rebotes"
            If chVista.Value = 0 Then
                vsRebotes.ZOrder 0: vsListado.Visible = False
            Else
                AccionImprimir
                vsListado.ZOrder 0: vsListado.Visible = True
            End If
    End Select
    
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    
    Select Case LCase(tabOpciones.SelectedItem.Key)
        Case "ingresos"
        
            'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
            Screen.MousePointer = 11
            With vsListado
                .StartDoc
                If .Error Then
                    MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                    Screen.MousePointer = 0: Exit Sub
                End If
            End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
            EncabezadoListado vsListado, "Cheques Ingresados al " & tDesde.FechaValor, True
            vsListado.FileName = "Cheques Ingresados"
            
            vsListado.Paragraph = cMoneda.Text
            vsIngresos.ExtendLastCol = False
            vsListado.RenderControl = vsIngresos.hwnd
            vsIngresos.ExtendLastCol = True
            
            vsListado.EndDoc
            vsListado.Refresh
            
        Case "rebotes"
        
            'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
            Screen.MousePointer = 11
            With vsListado
                .StartDoc
                If .Error Then
                    MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                    Screen.MousePointer = 0: Exit Sub
                End If
            End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
            EncabezadoListado vsListado, "Cheques Rebotados", True
            vsListado.FileName = "Cheques Rebotados"
            
            vsListado.Paragraph = cMoneda.Text
            vsRebotes.ExtendLastCol = False
            vsListado.RenderControl = vsRebotes.hwnd
            vsRebotes.ExtendLastCol = True
            
            vsListado.EndDoc
            vsListado.Refresh
        
    
    End Select
    
    If Imprimir Then
        frmSetup.pControl = vsListado
        frmSetup.Show vbModal, Me
        Me.Refresh
        If frmSetup.pOK Then vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    clsGeneral.OcurrioError "Error al realizar la impresión", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub


Private Sub InicializoGrilla()

    On Error Resume Next
         
    With vsIngresos
        .Cols = 1: .Rows = 1
        .FormatString = "ID Cheque|Hora|<Banco Emisor|<Vence|<Cheque|<Librado|>Importe|<|ID_Doc|Cli_Cheque|Cli_Doc|<Depositado"
        .ColWidth(1) = 650: .ColWidth(2) = 3400
        .ColWidth(3) = 1100: .ColWidth(4) = 1100: .ColWidth(5) = 1000: .ColWidth(6) = 1000
        .ColWidth(7) = 1000:: .ColWidth(11) = 1200
        .ColHidden(0) = True: .ColHidden(8) = True: .ColHidden(9) = True: .ColHidden(10) = True
        
        .WordWrap = False
        .ExtendLastCol = True
        .SubtotalPosition = flexSTBelow
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarNone
    End With
    
    With vsRebotes
        .Cols = 1: .Rows = 1
        .FormatString = "ID Cheque|Ingresado|<Banco Emisor|<Vence|<Cheque|<Librado|>Importe|<|ID_Doc|Cli_Cheque|Cli_Doc"
        .ColWidth(1) = 1400: .ColWidth(2) = 3400
        .ColWidth(3) = 1100: .ColWidth(4) = 1100: .ColWidth(5) = 1000: .ColWidth(6) = 1000
        .ColWidth(7) = 50:: .ColWidth(11) = 1200
        .ColHidden(0) = True: .ColHidden(8) = True: .ColHidden(9) = True: .ColHidden(10) = True
        
        .WordWrap = False
        .ExtendLastCol = True
        .SubtotalPosition = flexSTBelow
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarNone
    End With
    
    With vsDepositos
        .Cols = 1: .Rows = 1
        .FormatString = "ID Cheque|<Banco|<Sucursal|<Cheque|>Importe|"
        .ColWidth(1) = 4000
        .ColWidth(2) = 2100: .ColWidth(3) = 1100: .ColWidth(4) = 1600
        .ColHidden(0) = True
        
        .WordWrap = False
        .ExtendLastCol = False
        .SubtotalPosition = flexSTBelow
        .MergeCells = flexMergeSpill
        
    End With
    
End Sub

Private Sub AccionLimpiar()
    On Error Resume Next
    tDesde.FechaValor = Now
    cMoneda.Text = ""
End Sub

Private Sub PropiedadesImpresion()

  On Error Resume Next
  With vsListado
        .PaperSize = vbPRPSLetter
        .PhysicalPage = True
        .Orientation = orPortrait
        
        .PreviewMode = pmScreen
        
        .PreviewPage = 1
        .FontName = "Tahoma": .FontSize = 10: .FontBold = False: .FontItalic = False
        .Zoom = 100
        .MarginBottom = 500: .MarginTop = 750
        .MarginRight = 450: .MarginLeft = 450
        
    End With

End Sub

Private Sub vsIngresos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo errBD
    With vsIngresos
        If .Rows = 1 Then Exit Sub
        If Button <> vbRightButton Then Exit Sub
        .Select .MouseRow, 1
        If Val(.Cell(flexcpData, .Row, 0)) = 0 Then Exit Sub
        'ID Cheque|<Banco Emisor|<Vence|<Cheque|<Librado|>Importe|<|ID_Doc|Cli_Cheque|Cli_Doc
        
        MnuDetalleF.Enabled = Val(.Cell(flexcpText, .Row, 8)) <> 0
                
        MnuDetalleF.Tag = .Cell(flexcpText, .Row, 8)
        MnuDeudaCheques.Tag = .Cell(flexcpText, .Row, 9)
        MnuSeguimiento.Tag = .Cell(flexcpData, .Row, 0)
        MnuVOpeCCheque.Tag = .Cell(flexcpText, .Row, 9)
        MnuVOpeCDoc.Tag = .Cell(flexcpText, .Row, 10)
        
        PopupMenu MnuAcciones, , , , MnuAccTitulo
    End With
    
errBD:
End Sub


Private Sub ConsultoIngresados()

On Error GoTo errCargar
Dim aQ As Long

    Screen.MousePointer = 11
    
    vsIngresos.Rows = 1: vsIngresos.Refresh
    
    'Query COUNT(*) ---------------------------------------------------------------------------------------------------------
    aQ = 0
    Cons = "Select Count(*) from ChequeDiferido" & _
                " Where CDiIngresado Between '" & Format(tDesde.FechaValor & " 00:00", "mm/dd/yyyy hh:mm") & "'" & _
                                                " And '" & Format(tDesde.FechaValor & " 23:59:59", "mm/dd/yyyy hh:mm") & "'" & _
                " And CDiMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then aQ = RsAux(0)
    RsAux.Close
    
    pbProgreso.Value = 0
    If aQ > 0 Then pbProgreso.Max = aQ
    '------------------------------------------------------------------------------------------------------------------------------
    
    If aQ = 0 Then
        MsgBox "No hay cheques para los filtros ingresados.", vbInformation, "No hay datos"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    'Query Con DATOS---------------------------------------------------------------------------------------------------------
    Cons = "Select * From ChequeDiferido, BancoSSFF, SucursalDeBanco" _
            & " Where CDiIngresado Between '" & Format(tDesde.FechaValor & " 00:00", "mm/dd/yyyy hh:mm") & "'" & _
                                                " And '" & Format(tDesde.FechaValor & " 23:59:59", "mm/dd/yyyy hh:mm") & "'" _
            & " And CDiMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
            & " And CDiSucursal = SBaCodigo" _
            & " And SBaBanco = BanCodigo" _
            & " Order by CDiIngresado"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Dim mTipoCheque As String: mTipoCheque = ""
    Dim bAddHAlDia As Boolean, bAddHDiferido As Boolean
    bAddHAlDia = False: bAddHDiferido = False
    
    Dim mDAlDia As Currency, mDDiferido As Currency
    Dim mDQAlDia As Currency, mDQDiferido As Currency
    mDAlDia = 0: mDDiferido = 0
    mDQAlDia = 0: mDQDiferido = 0
    
    vsIngresos.Redraw = False
    
    Do While Not RsAux.EOF
        'ID Cheque|<Banco Emisor|<Vence|<Cheque|<Librado|>Importe|<|ID_Doc|Cli_Cheque|Cli_Doc
        With vsIngresos
            
            If IsNull(RsAux!CDiVencimiento) Then
                mTipoCheque = "Al Día"
                If Not bAddHAlDia Then
                    .AddItem CStr(mTipoCheque): bAddHAlDia = True
                    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Azul
                    .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
                End If
            Else
                mTipoCheque = "Diferidos"
                If Not bAddHDiferido Then
                    .AddItem CStr(mTipoCheque): bAddHDiferido = True
                    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Azul
                    .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
                End If
            End If
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = CStr(mTipoCheque)
            aValor = RsAux!CDiCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!CDiIngresado, "hh:mm")
            
            mTexto = Format(RsAux!BanCodigoB, "00") & "-" & Format(RsAux!SBaCodigoS, "000") & "    " & _
                            Trim(RsAux!BanNombre) & " (" & Trim(RsAux!SBaNombre) & ")"
            .Cell(flexcpText, .Rows - 1, 2) = mTexto
            
            If Not IsNull(RsAux!CDiVencimiento) Then .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!CDiVencimiento, "dd/mm/yy")
            
            .Cell(flexcpText, .Rows - 1, 4) = Trim(Trim(RsAux!CDiSerie) & " " & Trim(RsAux!CDiNumero))
            .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!CDiLibrado, "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!CDiImporte, "#,##0.00")
            
            If Not IsNull(RsAux!CDiRebotado) Then .Cell(flexcpText, .Rows - 1, 7) = "Rebotado"
            If Not IsNull(RsAux!CDiEliminado) Then
                .Cell(flexcpText, .Rows - 1, 7) = "Eliminado": .Cell(flexcpForeColor, .Rows - 1, 7) = Colores.Rojo
                .Cell(flexcpText, .Rows - 1, 6) = ""        'Corrigo el importe x la suma final
            End If
            
            If Not IsNull(RsAux!CDiDocumento) Then .Cell(flexcpText, .Rows - 1, 8) = RsAux!CDiDocumento
            If Not IsNull(RsAux!CDiCliente) Then .Cell(flexcpText, .Rows - 1, 9) = RsAux!CDiCliente
            If Not IsNull(RsAux!CDiClienteFactura) Then .Cell(flexcpText, .Rows - 1, 10) = RsAux!CDiClienteFactura
            
            If Not IsNull(RsAux!CDiCobrado) And IsNull(RsAux!CDiEliminado) Then
                .Cell(flexcpText, .Rows - 1, 11) = Format(RsAux!CDiCobrado, "dd/mm/yy hh:mm")
                If IsNull(RsAux!CDiVencimiento) Then
                    mDAlDia = mDAlDia + RsAux!CDiImporte
                    mDQAlDia = mDQAlDia + 1
                Else
                    mDDiferido = mDDiferido + RsAux!CDiImporte
                    mDQDiferido = mDQDiferido + 1
                End If
                
            Else
                .Cell(flexcpText, .Rows - 1, 11) = " "
            End If
            
        End With
        
        pbProgreso.Value = pbProgreso.Value + 1
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If vsIngresos.Rows > 1 Then
        With vsIngresos
            .ColSort(0) = flexSortGenericAscending
            .ColSort(1) = flexSortGenericAscending
            .Select 0, 0, 0, 1
            .Sort = flexSortUseColSort
            
            .Subtotal flexSTCount, 0, 4, "0", Colores.Gris, Colores.Azul, True, "%s"
            .Subtotal flexSTSum, 0, 6
            
            .AddItem ""
            .Subtotal flexSTCount, -1, 4, "0", Colores.Azul, Colores.Blanco, True, "Totales"
            .Subtotal flexSTSum, -1, 6
            
            'Recorro para agregar suma de depositos
            Dim mMonto As Currency, mQTotal As Integer
            
            Dim mRow As Integer
            'mRows = vsIngresos.Rows
            I = 0
            For mRow = 1 To vsIngresos.Rows - 1
                I = I + 1
                If .IsSubtotal(I) Then
                    Select Case LCase(Trim(.Cell(flexcpText, I, 0)))
                        Case "al día"
                            mMonto = .Cell(flexcpValue, I, 6)
                            mQTotal = .Cell(flexcpValue, I, 4)
                            
                            .AddItem "Al Dia Depositados", I
                            .Cell(flexcpText, I, 6) = Format(mDAlDia, "#,##0.00")
                            .Cell(flexcpText, I, 4) = Format(mDQAlDia, "#,##0")
                            .Cell(flexcpBackColor, I, 0, I, .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, I, 0, I, .Cols - 1) = Colores.Azul
                            I = I + 1
                            .AddItem "Al Dia Sin Depositar", I
                            .Cell(flexcpText, I, 6) = Format(mMonto - mDAlDia, "#,##0.00")
                            .Cell(flexcpText, I, 4) = Format(mQTotal - mDQAlDia, "#,##0")
                            .Cell(flexcpBackColor, I, 0, I, .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, I, 0, I, .Cols - 1) = Colores.Azul
                            I = I + 1
                            
                             I = I + 1: .AddItem "", I
                            
                        Case "diferidos"
                            mMonto = .Cell(flexcpValue, I, 6)
                            mQTotal = .Cell(flexcpValue, I, 4)
                            
                            .AddItem "Diferidos Depositados", I
                            .Cell(flexcpText, I, 6) = Format(mDDiferido, "#,##0.00")
                            .Cell(flexcpText, I, 4) = Format(mDQDiferido, "#,##0")
                            .Cell(flexcpBackColor, I, 0, I, .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, I, 0, I, .Cols - 1) = Colores.Azul
                            I = I + 1
                            .AddItem "Diferidos Sin Depositar", I
                            .Cell(flexcpText, I, 6) = Format(mMonto - mDDiferido, "#,##0.00")
                            .Cell(flexcpText, I, 4) = Format(mQTotal - mDQDiferido, "#,##0")
                            .Cell(flexcpBackColor, I, 0, I, .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, I, 0, I, .Cols - 1) = Colores.Azul
                            I = I + 1
                            
                        Case "totales"
                            mMonto = .Cell(flexcpValue, I, 6)
                            mQTotal = .Cell(flexcpValue, I, 4)
                            
                            .AddItem "Total Depositados", I
                            .Cell(flexcpText, I, 6) = Format(mDDiferido + mDAlDia, "#,##0.00")
                            .Cell(flexcpText, I, 4) = Format(mDQDiferido + mDQAlDia, "#,##0")
                            .Cell(flexcpBackColor, I, 0, I, .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, I, 0, I, .Cols - 1) = Colores.Azul
                            I = I + 1
                            .AddItem "Total Sin Depositar", I
                            .Cell(flexcpText, I, 6) = Format(mMonto - (mDDiferido + mDAlDia), "#,##0.00")
                            .Cell(flexcpText, I, 4) = Format(mQTotal - (mDQDiferido + mDQAlDia), "#,##0")
                            .Cell(flexcpBackColor, I, 0, I, .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, I, 0, I, .Cols - 1) = Colores.Azul
                            I = I + 1
                    End Select
                End If
            Next
                    
        End With
    End If
    pbProgreso.Value = 0
    vsIngresos.Redraw = True
    Screen.MousePointer = 0

    Exit Sub
errCargar:
    vsIngresos.Redraw = True
    clsGeneral.OcurrioError "Error al cargar los cheques ingresados.", Err.Description
    Screen.MousePointer = 0
End Sub



Private Sub ConsultoDepositados()
On Error GoTo errCargar
    
Dim aQ As Long

    Screen.MousePointer = 11
    
    vsDepositos.Rows = 1
    
    'Query COUNT(*) ---------------------------------------------------------------------------------------------------------
    aQ = 0
    Cons = "Select Count(*) from ChequeDiferido" & _
                " Where CDiCobrado Between '" & Format(tDesde.FechaValor & " 00:00", "mm/dd/yyyy hh:mm") & "'" & _
                                                " And '" & Format(tDesde.FechaValor & " 23:59:59", "mm/dd/yyyy hh:mm") & "'" & _
                " And CDiMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then aQ = RsAux(0)
    RsAux.Close
    
    pbProgreso.Value = 0
    If aQ > 0 Then pbProgreso.Max = aQ
    '------------------------------------------------------------------------------------------------------------------------------
    
    If aQ = 0 Then
        MsgBox "No hay cheques para los filtros ingresados.", vbInformation, "No hay datos"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    Dim mHeaderBanco As String
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    With vsListado
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
    EncabezadoListado vsListado, "Depósitos de Cheques al " & tDesde.FechaValor, True
    vsListado.FileName = "Depósitos de Cheques"
        
    'Query Con DATOS---------------------------------------------------------------------------------------------------------
     Cons = "Select CDiCodigo, CDiSerie, CDiNumero, CDiImporte, CDiVencimiento, CDiIngresado, " & _
                            " NomBcoDep =  BancoDeposito.BanNombre, CodBco = BancoDeposito.BanCodigo, " & _
                            " SucBco = SucursalDeBanco.SBaNombre, NroCuenta = SucursalDeBanco.SBaCuentaCGSA, SucCod = SucursalDeBanco.SBaCodigo, " & _
                            " Banco = BancoSSFF.BanNombre, Sucursal = SucursalCheque.SBaNombre " & _
                " From ChequeDiferido, BancoSSFF, SucursalDeBanco SucursalCheque, SucursalDeBanco, BancoSSFF BancoDeposito " & _
                " Where CDiCobrado Between '" & Format(tDesde.FechaValor & " 00:00", "mm/dd/yyyy hh:mm") & "'" & _
                                                " And '" & Format(tDesde.FechaValor & " 23:59:59", "mm/dd/yyyy hh:mm") & "'" & _
                " And CDiMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) & _
                " And CDiBanco = BancoSSFF.BanCodigo " & _
                " And CDiSucursal = SucursalCheque.SBaCodigo " & _
                " And CDiDepositado = SucursalDeBanco.SBaCodigo" & _
                " And SucursalDeBanco.SBaBanco = BancoDeposito.BanCodigo" & _
                " Order by  CodBco, SucCod, CDiIngresado"
                '" Order by CodBco, SucCod, CDiSerie, CDiNumero"
    
    cBase.QueryTimeout = 30
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Dim mTipo As String
    Dim bAlDia As Boolean, bDiferidos As Boolean
    Dim mIDCarga As Long: mIDCarga = 0
    Dim bPie As Boolean
    Dim mQPorTabla As Integer
    
    'vsDepositos.OutlineCol = 0
    
    mQPorTabla = 0
    Do While Not RsAux.EOF
        mQPorTabla = mQPorTabla + 1
        If mIDCarga <> RsAux!SucCod Or mQPorTabla = 1 Then
            If mIDCarga <> 0 Then vsListado.NewPage
            mIDCarga = RsAux!SucCod
            vsDepositos.Rows = 1
            bAlDia = False: bDiferidos = False
            mHeaderBanco = Trim(RsAux!NomBcoDep) & "  (" & Trim(RsAux!SucBco) & ")"
            If Not IsNull(RsAux!NroCuenta) Then mHeaderBanco = mHeaderBanco & Space(16) & "Cuenta Nº: " & RsAux!NroCuenta
        End If
            
        With vsDepositos
            If IsNull(RsAux!CDiVencimiento) Then
                mTipo = "Cheques Al Día"
                If Not bAlDia Then .AddItem mTipo: .Cell(flexcpFontBold, .Rows - 1, 0) = True
                bAlDia = True
            Else
                mTipo = "Cheques Diferidos"
                If Not bDiferidos Then .AddItem mTipo: .Cell(flexcpFontBold, .Rows - 1, 0) = True
                bDiferidos = True
            End If
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = mTipo 'rsAux!CDiCodigo
            
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!CDiIngresado, "dd/mm/yy hh:mm") & "      " & Trim(RsAux!Banco)
            .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!Sucursal)
            .Cell(flexcpText, .Rows - 1, 3) = Trim(Trim(RsAux!CDiSerie) & " " & Trim(RsAux!CDiNumero))
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!CDiImporte, "#,##0.00")
            
        End With
        
        pbProgreso.Value = pbProgreso.Value + 1
        RsAux.MoveNext
        
        bPie = False
        If RsAux.EOF Then
            bPie = True
        Else
            If RsAux!SucCod <> mIDCarga Then bPie = True
        End If
        
        If bPie Or mQPorTabla = 30 Then
            mQPorTabla = 0
            If vsDepositos.Rows > 1 Then
                
                With vsDepositos
                    .ColSort(0) = flexSortGenericAscending
                    .Sort = flexSortUseColSort
                    .Subtotal flexSTCount, 0, 3, "0", Colores.Azul, Colores.Blanco, True, "%s"  '"Totales"
                    .Subtotal flexSTSum, 0, 4
                    
                    .AddItem ""
                    .Subtotal flexSTCount, -1, 3, "0", Colores.Azul, Colores.Blanco, True, "Totales"
                    .Subtotal flexSTSum, -1, 4
                End With
                
                vsListado.Paragraph = mHeaderBanco
                vsListado.Paragraph = ""
                vsListado.Paragraph = cMoneda.Text
                
                vsListado.RenderControl = vsDepositos.hwnd
                vsDepositos.Row = 1
            End If
        End If
    Loop
    RsAux.Close
    
    vsListado.EndDoc
    vsListado.Refresh
    'vsListado.Visible = True: vsListado.ZOrder 0
    pbProgreso.Value = 0
    Screen.MousePointer = 0

    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los cheques depositados.", Err.Description
    Screen.MousePointer = 0
End Sub



Private Sub ConsultoRebotes()

On Error GoTo errCargar
Dim aQ As Long

    Screen.MousePointer = 11
    
    vsRebotes.Rows = 1: vsRebotes.Refresh
    
    'Query COUNT(*) ---------------------------------------------------------------------------------------------------------
    aQ = 0
    Cons = "Select Count(*) from ChequeDiferido" & _
                " Where CDiRebotado Is Not Null " & _
                " And CDiMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) & _
                " And CDiEliminado is Null"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then aQ = RsAux(0)
    RsAux.Close
    
    pbProgreso.Value = 0
    If aQ > 0 Then pbProgreso.Max = aQ
    '------------------------------------------------------------------------------------------------------------------------------
    
    If aQ = 0 Then
        MsgBox "No hay cheques para los filtros ingresados.", vbInformation, "No hay datos"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    'Query Con DATOS---------------------------------------------------------------------------------------------------------
    Cons = "Select * From ChequeDiferido, BancoSSFF, SucursalDeBanco" _
            & " Where CDiRebotado Is Not Null " _
            & " And CDiMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
            & " And CDiSucursal = SBaCodigo" _
            & " And SBaBanco = BanCodigo" _
            & " And CDiEliminado is Null" _
            & " Order by CDiIngresado"
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Dim mTipoCheque As String: mTipoCheque = ""
    Dim bAddHAlDia As Boolean, bAddHDiferido As Boolean
    bAddHAlDia = False: bAddHDiferido = False
    
    Dim mDAlDia As Currency, mDDiferido As Currency
    Dim mDQAlDia As Currency, mDQDiferido As Currency
    mDAlDia = 0: mDDiferido = 0
    mDQAlDia = 0: mDQDiferido = 0
    
    vsRebotes.Redraw = False
    
    Do While Not RsAux.EOF
        'ID Cheque|<Banco Emisor|<Vence|<Cheque|<Librado|>Importe|<|ID_Doc|Cli_Cheque|Cli_Doc
        With vsRebotes
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = CStr(RsAux!CDiCodigo)
            aValor = RsAux!CDiCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!CDiIngresado, "dd/mm/yyyy hh:mm")
            
            mTexto = Format(RsAux!BanCodigoB, "00") & "-" & Format(RsAux!SBaCodigoS, "000") & "    " & _
                            Trim(RsAux!BanNombre) & " (" & Trim(RsAux!SBaNombre) & ")"
            .Cell(flexcpText, .Rows - 1, 2) = mTexto
            
            If Not IsNull(RsAux!CDiVencimiento) Then .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!CDiVencimiento, "dd/mm/yy")
            
            .Cell(flexcpText, .Rows - 1, 4) = Trim(Trim(RsAux!CDiSerie) & " " & Trim(RsAux!CDiNumero))
            .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!CDiLibrado, "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!CDiImporte, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 7) = " "
            
            
            If Not IsNull(RsAux!CDiDocumento) Then .Cell(flexcpText, .Rows - 1, 8) = RsAux!CDiDocumento
            If Not IsNull(RsAux!CDiCliente) Then .Cell(flexcpText, .Rows - 1, 9) = RsAux!CDiCliente
            If Not IsNull(RsAux!CDiClienteFactura) Then .Cell(flexcpText, .Rows - 1, 10) = RsAux!CDiClienteFactura
                       
        End With
        
        pbProgreso.Value = pbProgreso.Value + 1
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If vsRebotes.Rows > 1 Then
        With vsRebotes
            .ColSort(0) = flexSortGenericAscending
            .ColSort(1) = flexSortGenericAscending
            .Select 0, 0, 0, 1
            .Sort = flexSortUseColSort
            
            .AddItem ""
            .Subtotal flexSTCount, -1, 4, "0", Colores.Azul, Colores.Blanco, True, "Totales"
            .Subtotal flexSTSum, -1, 6
        End With
    End If
    
    pbProgreso.Value = 0
    vsRebotes.Redraw = True
    Screen.MousePointer = 0

    Exit Sub
errCargar:
    vsRebotes.Redraw = True
    clsGeneral.OcurrioError "Error al cargar los cheques rebotados.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub vsRebotes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error GoTo errBD
 
    With vsRebotes
        If .Rows = 1 Then Exit Sub
        If Button <> vbRightButton Then Exit Sub
        .Select .MouseRow, 1
        If Val(.Cell(flexcpData, .Row, 0)) = 0 Then Exit Sub
        'ID Cheque|<Banco Emisor|<Vence|<Cheque|<Librado|>Importe|<|ID_Doc|Cli_Cheque|Cli_Doc
        
        MnuDetalleF.Enabled = Val(.Cell(flexcpText, .Row, 8)) <> 0
                
        MnuDetalleF.Tag = .Cell(flexcpText, .Row, 8)
        MnuDeudaCheques.Tag = .Cell(flexcpText, .Row, 9)
        MnuSeguimiento.Tag = .Cell(flexcpData, .Row, 0)
        MnuVOpeCCheque.Tag = .Cell(flexcpText, .Row, 9)
        MnuVOpeCDoc.Tag = .Cell(flexcpText, .Row, 10)
        
        PopupMenu MnuAcciones, , , , MnuAccTitulo
    End With
    
errBD:
End Sub

