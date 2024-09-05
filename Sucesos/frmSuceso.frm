VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F93D243E-5C15-11D5-A90D-000021860458}#10.0#0"; "orFecha.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmSuceso 
   BackColor       =   &H8000000B&
   Caption         =   "Visualización de Sucesos"
   ClientHeight    =   6960
   ClientLeft      =   1710
   ClientTop       =   2385
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
   Icon            =   "frmSuceso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   10665
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   39
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
            Object.Width           =   18389
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   3495
      Left            =   4320
      TabIndex        =   36
      Top             =   2400
      Width           =   5655
      _Version        =   196608
      _ExtentX        =   9975
      _ExtentY        =   6165
      _StockProps     =   229
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
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   420
      ScaleHeight     =   435
      ScaleWidth      =   7275
      TabIndex        =   20
      Top             =   6060
      Width           =   7335
      Begin VB.CommandButton bAyuda 
         Height          =   310
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Ayuda."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmSuceso.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmSuceso.frx":067C
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmSuceso.frx":097E
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmSuceso.frx":0CC0
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4560
         Picture         =   "frmSuceso.frx":0FC2
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3480
         Picture         =   "frmSuceso.frx":1388
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmSuceso.frx":148A
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2340
         Picture         =   "frmSuceso.frx":16C4
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2700
         Picture         =   "frmSuceso.frx":17AE
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   3840
         Picture         =   "frmSuceso.frx":1898
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmSuceso.frx":1D12
         Height          =   310
         Left            =   4200
         Picture         =   "frmSuceso.frx":1E5C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6120
         TabIndex        =   33
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
      Height          =   1335
      Left            =   60
      TabIndex        =   19
      Top             =   60
      Width           =   10035
      Begin VB.TextBox tUsuario 
         Height          =   300
         Left            =   7860
         TabIndex        =   9
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox tDocumento 
         Height          =   300
         Left            =   1140
         TabIndex        =   15
         Top             =   960
         Width           =   915
      End
      Begin VB.TextBox tDescripcion 
         Height          =   300
         Left            =   5220
         TabIndex        =   13
         Top             =   600
         Width           =   3795
      End
      Begin orctFecha.orFecha tDesde 
         Height          =   285
         Left            =   1140
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
      Begin MSMask.MaskEdBox tHora1 
         Height          =   285
         Left            =   5220
         TabIndex        =   5
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "99:99"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tHora2 
         Height          =   285
         Left            =   6300
         TabIndex        =   7
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "99:99"
         PromptChar      =   "_"
      End
      Begin AACombo99.AACombo cSuceso 
         Height          =   315
         Left            =   1140
         TabIndex        =   11
         Top             =   600
         Width           =   3015
         _ExtentX        =   5318
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
      Begin orctFecha.orFecha tHasta 
         Height          =   285
         Left            =   2700
         TabIndex        =   3
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
      Begin MSMask.MaskEdBox tCi 
         Height          =   300
         Left            =   5220
         TabIndex        =   17
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   12582912
         PromptInclude   =   0   'False
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "#.###.###-#"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox tRuc 
         Height          =   300
         Left            =   5220
         TabIndex        =   18
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   12582912
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "## ### ### ####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuario:"
         Height          =   255
         Left            =   7200
         TabIndex        =   8
         Top             =   285
         Width           =   615
      End
      Begin VB.Label lDoc 
         BackColor       =   &H00808080&
         Caption         =   "Label9"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2160
         TabIndex        =   38
         Top             =   975
         Width           =   1995
      End
      Begin VB.Label lCliente 
         BackColor       =   &H00808080&
         Caption         =   "Label9"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6660
         TabIndex        =   37
         Top             =   975
         Width           =   3195
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&CI/Ruc:"
         Height          =   255
         Left            =   4320
         TabIndex        =   16
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Docu&mento:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descri&pción:"
         Height          =   255
         Left            =   4260
         TabIndex        =   12
         Top             =   640
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   260
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo &Suceso:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   640
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Hora:"
         Height          =   255
         Left            =   4680
         TabIndex        =   4
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Al"
         Height          =   255
         Left            =   2460
         TabIndex        =   2
         Top             =   255
         Width           =   255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         Height          =   255
         Left            =   6060
         TabIndex        =   6
         Top             =   255
         Width           =   135
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   3255
      Left            =   3120
      TabIndex        =   34
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
            Picture         =   "frmSuceso.frx":238E
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSuceso.frx":26A8
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuBDerecho 
      Caption         =   "BotonDerecho"
      Visible         =   0   'False
      Begin VB.Menu MnuIrA 
         Caption         =   "Ir a ..."
      End
      Begin VB.Menu MnuVerL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSuceso 
         Caption         =   "Detalle del Suceso"
      End
      Begin VB.Menu MnuFactura 
         Caption         =   "Detalle de Factura"
      End
      Begin VB.Menu MnuCliente 
         Caption         =   "&Visualización de Operaciones"
      End
      Begin VB.Menu MnuComentarios 
         Caption         =   "Agregar Comentarios al Cliente"
      End
      Begin VB.Menu MnuSuCliente 
         Caption         =   "Ver Sucesos del Cliente"
      End
   End
End
Attribute VB_Name = "frmSuceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim aValor As Long

Private Sub bAyuda_Click()
On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    cons = "Select * from Aplicacion Where AplNombre = 'Sucesos'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux!AplHelp) Then aFile = Trim(rsAux!AplHelp)
    rsAux.Close
    
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
    If tHasta.FechaValor = "" Then
        MsgBox "La fecha 'hasta' no es correcta. Verifique", vbExclamation, "Posible Error"
        Foco tHasta: Exit Sub
    End If
    If CDate(tDesde.FechaValor) > CDate(tHasta.FechaValor) Then
        MsgBox "El período de fechas ingresado no es correcto.", vbExclamation, "Posbile Error"
        Foco tDesde: Exit Sub
    End If
    
    If Not IsDate(tHora1.Text) Then
        MsgBox "El rango de horas ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tHora1: Exit Sub
    End If
    If Not IsDate(tHora2.Text) Then
        MsgBox "El rango de horas ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tHora2: Exit Sub
    End If
    If tDesde.FechaValor = tHasta.FechaValor Then
        If CDate(tHora1.Text) > CDate(tHora2.Text) Then
            MsgBox "El rango de horas ingresado no es correcto.", vbExclamation, "Posbile Error"
            Foco tHora2: Exit Sub
        End If
    End If
    '------------------------------------------------------------------------------------------------------------
    
    AccionConsultar
    If vsConsulta.Rows > 1 Then vsConsulta.SetFocus
    
End Sub

Private Sub cSuceso_GotFocus()
    cSuceso.SelStart = 0: cSuceso.SelLength = Len(cSuceso.Text)
    Status.Panels(1).Text = "Tipo de sucesos a filtrar."
End Sub

Private Sub cSuceso_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDescripcion
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
    cons = "Select TSuCodigoSistema, TSuNombre from TipoSuceso Order by TSuNombre"
    CargoCombo cons, cSuceso
    
    picBotones.BorderStyle = 0
    vsConsulta.ZOrder 0
    
    bCancelar.Picture = img1.ListImages("salir").ExtractIcon
    bAyuda.Picture = img1.ListImages("help").ExtractIcon
   
End Sub

Public Function gbl_Consulto(fDesde As String, fHasta As String, idSuceso As Long)
On Error GoTo errGbl

Dim bQuery As Boolean: bQuery = False
    
    If Trim(fDesde) <> "" Then
        If IsDate(fDesde) Then tDesde.FechaValor = CDate(fDesde): bQuery = True
    End If
    If Trim(fHasta) <> "" Then
        If IsDate(fHasta) Then tHasta.FechaValor = CDate(fHasta): bQuery = True
    End If
    
    If idSuceso <> 0 Then
        BuscoCodigoEnCombo cSuceso, idSuceso
        If cSuceso.ListIndex <> -1 Then bQuery = True
    End If
    
    If bQuery Then
        Me.Show: DoEvents
        AccionConsultar
    End If
errGbl:
End Function

Private Sub AccionConsultar()
On Error GoTo errPago
    
Dim Fecha1 As String, Fecha2 As String
Dim aQ As Long

    Screen.MousePointer = 11
    
    Fecha1 = Format(tDesde.FechaValor, "mm/dd/yyyy") & " " & tHora1.Text
    Fecha2 = Format(tHasta.FechaValor, "mm/dd/yyyy") & " " & tHora2.Text
    vsConsulta.Rows = 1
    
    'Query COUNT(*) ---------------------------------------------------------------------------------------------------------
    aQ = 0
    cons = "Select Count(*) from Suceso" _
            & " Where SucFecha Between '" & Fecha1 & "' And '" & Fecha2 & "'"
    If cSuceso.ListIndex <> -1 Then cons = cons & " And SucTipo = " & cSuceso.ItemData(cSuceso.ListIndex)
    If Trim(tDescripcion.Text) <> "" Then cons = cons & " And SucDescripcion like '" & Trim(tDescripcion.Text) & "%'"
    If Val(lCliente.Tag) <> 0 Then cons = cons & " And SucCliente = " & Val(lCliente.Tag)
    If Val(lDoc.Tag) <> 0 Then cons = cons & " And SucDocumento = " & Val(lDoc.Tag)
    If Val(tUsuario.Tag) <> 0 Then cons = cons & " And SucUsuario = " & Val(tUsuario.Tag)
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux(0)) Then aQ = rsAux(0)
    rsAux.Close
    
    pbProgreso.Value = 0
    If aQ > 0 Then pbProgreso.Max = aQ
    '------------------------------------------------------------------------------------------------------------------------------
    
    If aQ = 0 Then
        MsgBox "No hay sucesos para los filtros ingresados.", vbInformation, "No hay datos"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    'Query Con DATOS---------------------------------------------------------------------------------------------------------
    cons = "Select * from Suceso, Usuario" _
            & " Where SucFecha Between '" & Fecha1 & "' And '" & Fecha2 & "'"
         
    If cSuceso.ListIndex <> -1 Then cons = cons & " And SucTipo = " & cSuceso.ItemData(cSuceso.ListIndex)
    If Trim(tDescripcion.Text) <> "" Then cons = cons & " And SucDescripcion like '" & Trim(tDescripcion.Text) & "%'"
    If Val(lCliente.Tag) <> 0 Then cons = cons & " And SucCliente = " & Val(lCliente.Tag)
    If Val(lDoc.Tag) <> 0 Then cons = cons & " And SucDocumento = " & Val(lDoc.Tag)
    If Val(tUsuario.Tag) <> 0 Then cons = cons & " And SucUsuario = " & Val(tUsuario.Tag)
    cons = cons _
            & " And SucUsuario = UsuCodigo" _
            & " Order by SucFecha DESC"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not rsAux.EOF
        'Filtro la Hora del Suceso
        If Format(rsAux!SucFecha, "hh:mm") >= tHora1.Text And Format(rsAux!SucFecha, "hh:mm") <= tHora2.Text Then
            
            With vsConsulta
                .AddItem CStr(rsAux!SucCodigo)
                .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!SucFecha, "dd/mm hh:mm")
                
                If Not IsNull(rsAux!SucDocumento) Then aValor = rsAux!SucDocumento Else aValor = 0
                .Cell(flexcpData, .Rows - 1, 0) = aValor
                
                If Not IsNull(rsAux!SucCliente) Then aValor = rsAux!SucCliente Else aValor = 0
                .Cell(flexcpData, .Rows - 1, 1) = aValor

                
                
                If Not IsNull(rsAux!SucDescripcion) Then .Cell(flexcpText, .Rows - 1, 2) = Trim(rsAux!SucDescripcion)
                .Cell(flexcpText, .Rows - 1, 3) = Trim(rsAux!UsuIdentificacion)
                        
                If Not IsNull(rsAux!SucDefensa) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(rsAux!SucDefensa)
                
            End With
        End If
        
        pbProgreso.Value = pbProgreso.Value + 1
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    If vsConsulta.Rows = 1 Then MsgBox "No hay sucesos para los filtros ingresados.", vbInformation, "No hay datos"
    
    vsConsulta.AutoSizeMode = flexAutoSizeRowHeight
    vsConsulta.AutoSize 2, , False
    Exit Sub
    
errPago:
    clsGeneral.OcurrioError "Error al cargar los sucesos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With frame1
        .Left = 60: .Top = 60
        .Width = Me.ScaleWidth - (.Left * 2)
    End With
    lCliente.Width = frame1.Width - lCliente.Left - 200
    
    With vsConsulta
        .Width = frame1.Width: .Left = frame1.Left
        .Top = frame1.Top + frame1.Height + 80
        .Height = Me.ScaleHeight - .Top - picBotones.Height - Status.Height
    End With
    
    With vsListado
        .Width = vsConsulta.Width: .Left = vsConsulta.Left
        .Top = vsConsulta.Top: .Height = vsConsulta.Height
    End With
    
    With picBotones
        .Top = Me.ScaleHeight - .Height - Status.Height
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 100
    
    With vsConsulta
        Dim aSize As Currency
        For I = 0 To .Cols - 3: aSize = aSize + .ColWidth(I): Next I
        .ColWidth(.Cols - 2) = .Width - (aSize + .ColWidth(.Cols - 1) + 300)
        
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 2, , False
    End With
    
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
    Foco cSuceso
End Sub

Private Sub Label3_Click()
    tHasta.SetFocus
End Sub

Private Sub Label4_Click()
    tHora1.SetFocus
End Sub

Private Sub Label5_Click()
    tHora2.SetFocus
End Sub

Private Sub MnuCliente_Click()
On Error GoTo errCliente
    Screen.MousePointer = 11
    
    aValor = vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)
    EjecutarApp prmPathApp & "Visualizacion de Operaciones.exe", CStr(aValor)
    
    Screen.MousePointer = 0
    Exit Sub
    
errCliente:
    clsGeneral.OcurrioError "Error al acceder a la ficha del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuComentarios_Click()
On Error GoTo errCliente
    
    aValor = vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)
    If aValor = 0 Then Exit Sub
    Screen.MousePointer = 11

    Dim miCliente As New clsCliente
    miCliente.Comentarios aValor
    
    Me.Refresh
    Set miCliente = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
errCliente:
    clsGeneral.OcurrioError "Error al acceder a los comentarios.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuFactura_Click()
    On Error GoTo errFactura
        
    aValor = vsConsulta.Cell(flexcpData, vsConsulta.Row, 0)
    EjecutarApp prmPathApp & "Detalle de Factura", CStr(aValor)
    
errFactura:
End Sub

Private Sub MnuSuceso_Click()
    Call vsConsulta_DblClick
End Sub

Private Sub MnuSuCliente_Click()
On Error GoTo errFactura
        
    aValor = vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)
    If aValor = 0 Then Exit Sub
    EjecutarApp prmPathApp & "Suceso_Cliente", CStr(aValor)
    
errFactura:
End Sub

Private Sub tCi_Change()
    lCliente.Caption = "": lCliente.Tag = 0
End Sub

Private Sub tCi_GotFocus()
    tCi.SelStart = 0: tCi.SelLength = (Len(tCi.FormattedText))
    Status.Panels(1).Text = "Ingrese el cliente a filtrar.  [F2]- Cambia CI/Ruc.   [F4]- Buscar."
End Sub

Private Sub tCi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2, vbKeyR, vbKeyE, vbKeyC: tCi.Visible = False: tRuc.Visible = True: tRuc.SetFocus: lCliente.Tag = 0: lCliente.Caption = ""
        
        Case vbKeyReturn
                If Val(lCliente.Tag) = 0 And Trim(tCi.Text) <> "" Then
                    If Len(tCi.Text) < 7 Then Exit Sub
                    If Len(tCi.Text) = 7 Then tCi.Text = clsGeneral.AgregoDigitoControlCI(Trim(tCi.Text))
                    BuscarCliente miCi:=Trim(tCi.Text)
                Else
                    bConsultar.SetFocus
                End If
        
        Case vbKeyF4: BuscarClientes TipoCliente.Cliente
    End Select
    
End Sub

Private Sub BuscarClientes(aTipoCliente As Integer)
    
    On Error GoTo errCargar
    Screen.MousePointer = 11
    Dim objBuscar As New clsBuscarCliente
    Dim aTipo As Integer, aCliente As Long
    
    If aTipoCliente = TipoCliente.Cliente Then objBuscar.ActivoFormularioBuscarClientes txtConexion, Persona:=True
    If aTipoCliente = TipoCliente.Empresa Then objBuscar.ActivoFormularioBuscarClientes txtConexion, Empresa:=True
    Me.Refresh
    aTipo = objBuscar.BCTipoClienteSeleccionado
    aCliente = objBuscar.BCClienteSeleccionado
    Set objBuscar = Nothing
    
    If aCliente <> 0 Then BuscarCliente miId:=aCliente, miTipo:=aTipo
    
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tCi_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub tDescripcion_GotFocus()
    tDescripcion.SelStart = 0: tDescripcion.SelLength = Len(tDescripcion)
    Status.Panels(1).Text = "Descripción de los sucesos a filtrar."
End Sub

Private Sub tDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tDocumento.SetFocus
End Sub

Private Sub tDesde_GotFocus()
    tDesde.SelStart = 0: tDesde.SelLength = Len(tDesde.FechaText)
    Status.Panels(1).Text = "Ingrese el rango de fechas para realizar la consulta."
End Sub

Private Sub tDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And tDesde.FechaValor <> "" Then tHasta.SetFocus
End Sub

Private Sub tDesde_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub tDocumento_Change()
    lDoc.Caption = "": lDoc.Tag = 0
End Sub

Private Sub tDocumento_GotFocus()
    tDocumento.SelStart = 0: tDocumento.SelLength = (Len(tDocumento.Text))
    Status.Panels(1).Text = "Ingrese el núemero de documento para filtrar los datos."
End Sub

Private Sub tDocumento_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(lDoc.Tag) <> 0 Then
            If tCi.Visible Then tCi.SetFocus Else tRuc.SetFocus
            Exit Sub
        End If
        If Trim(tDocumento.Text) = "" Then
            If tCi.Visible Then tCi.SetFocus Else tRuc.SetFocus
            Exit Sub
        End If
        
        If Not IsNumeric(tDocumento.Text) Then Exit Sub
        
        Dim adQ As Integer, adCodigo As Long, adTexto As String
        
        Screen.MousePointer = 11
        adQ = 0
        cons = "Select * from Documento Where DocNumero = " & Val(tDocumento.Text) '& _
                   " And DocNumero = " & Val(tDocumento.Text)
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            adCodigo = rsAux!DocCodigo
            adTexto = Documento(rsAux!DocTipo, rsAux!DocSerie, rsAux!DocNumero)
            adQ = 1
            rsAux.MoveNext: If Not rsAux.EOF Then adQ = 2
        End If
        rsAux.Close
        
        Select Case adQ
            Case 2
                Dim miLDocs As New clsListadeAyuda
                cons = "Select DocCodigo, DocFecha as Fecha, DocSerie as Serie, Convert(char(7),DocNumero) as Numero " & _
                           " from Documento Where DocNumero = " & Val(tDocumento.Text)
                miLDocs.ActivoListaAyuda cons, False, txtConexion, 4100
                Me.Refresh
                adCodigo = miLDocs.ValorSeleccionado
                Set miLDocs = Nothing
                
                If adCodigo > 0 Then
                    cons = "Select * from Documento Where DocCodigo = " & adCodigo
                    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    If Not rsAux.EOF Then
                        adTexto = Documento(rsAux!DocTipo, rsAux!DocSerie, rsAux!DocNumero)
                    End If
                    rsAux.Close
                End If
        End Select
        
        If adCodigo > 0 Then
            lDoc.Tag = adCodigo: lDoc.Caption = adTexto
        Else
            lDoc.Caption = " No Existe !!"
        End If
        
        If Val(lDoc.Tag) <> 0 Then If tCi.Visible Then tCi.SetFocus Else tRuc.SetFocus
        
        Screen.MousePointer = 0
    End If
    
    Exit Sub
errDoc:
    clsGeneral.OcurrioError "Error al buscar el documento.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function Documento(Tipo As Integer, Serie As String, Numero As Long) As String

    Select Case Tipo
        Case 1: Documento = "Ctdo. "
        Case 2: Documento = "Créd. "
        Case 3: Documento = "N/Dev. "
        Case 4: Documento = "N/Créd. "
        Case 5: Documento = "Recibo "
        Case 10: Documento = "N/Esp. "
    End Select
    
    Documento = Documento & Trim(Serie) & " " & Numero

End Function

Private Sub tDocumento_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub tHasta_GotFocus()
    tHasta.SelStart = 0: tHasta.SelLength = Len(tHasta.FechaText)
    Status.Panels(1).Text = "Ingrese el rango de fechas para realizar la consulta."
End Sub

Private Sub tHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And tHasta.FechaValor <> "" Then tHora1.SetFocus
End Sub

Private Sub tHasta_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub tHora1_GotFocus()
    tHora1.SelStart = 0: tHora1.SelLength = 5
    Status.Panels(1).Text = "Ingrese el rango horario para buscar los sucesos."
End Sub

Private Sub tHora1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsDate(tHora1.Text) Then tHora2.SetFocus
End Sub

Private Sub tHora1_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub tHora2_GotFocus()
    tHora2.SelStart = 0: tHora2.SelLength = 5
    Status.Panels(1).Text = "Ingrese el rango horario para buscar los sucesos."
End Sub

Private Sub tHora2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsDate(tHora2.Text) Then Foco tUsuario
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
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
    Else
        AccionImprimir
        vsListado.ZOrder 0
    End If
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    If vsConsulta.Rows > 1 Then
        'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
        Screen.MousePointer = 11
        With vsListado
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        EncabezadoListado vsListado, "Consulta de Sucesos", False
        vsListado.FileName = "Consulta de Sucesos"
        
        vsConsulta.ExtendLastCol = False
        vsListado.RenderControl = vsConsulta.hwnd
        vsConsulta.ExtendLastCol = True
        
        vsListado.EndDoc
        vsListado.Refresh
    End If
    
    If Imprimir Then
        frmSetup.pControl = vsListado
        frmSetup.Show vbModal, Me
        Me.Refresh
        If frmSetup.pOK Then vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub


Private Sub InicializoGrilla()

    On Error Resume Next
         
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = "<ID Suceso|<Fecha|<Descripción|<Usuario|<Defensa|"
        .ColWidth(1) = 1100
        .ColWidth(2) = 3750: .ColWidth(3) = 1100: .ColWidth(4) = 3000
        
        .WordWrap = True
        .ExtendLastCol = True
        
        .WordWrap = True
        .ColAlignment(0) = flexAlignLeftTop: .ColAlignment(1) = flexAlignLeftTop: .ColAlignment(2) = flexAlignLeftTop: .ColAlignment(3) = flexAlignLeftTop
        .ColAlignment(4) = flexAlignLeftTop
    End With
    
End Sub

Private Sub tHora2_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub tRuc_Change()
lCliente.Caption = "": lCliente.Tag = 0
End Sub

Private Sub tRuc_GotFocus()
    tRuc.SelStart = 0: tRuc.SelLength = (Len(tRuc.FormattedText))
    Status.Panels(1).Text = "Ingrese el cliente a filtrar.  [F2]- Cambia CI/Ruc.   [F4]- Buscar."
End Sub

Private Sub tRuc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2, vbKeyP, vbKeyC
                tCi.Visible = True: tRuc.Visible = False: tCi.SetFocus:: lCliente.Tag = 0: lCliente.Caption = ""
                
        Case vbKeyReturn
                If Val(lCliente.Tag) = 0 And Trim(tRuc.Text) <> "" Then
                    BuscarCliente miRuc:=Trim(tRuc.Text)
                Else
                    bConsultar.SetFocus
                End If
                
        Case vbKeyF4: BuscarClientes TipoCliente.Empresa
    End Select
End Sub

Private Sub tRuc_LostFocus()
    Status.Panels(1).Text = ""
End Sub

Private Sub tUsuario_Change()
    tUsuario.Tag = 0
End Sub

Private Sub tUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errBuscar
    
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(tUsuario.Text) = "" Or Val(tUsuario.Tag) <> 0 Then Foco cSuceso: Exit Sub
            
            Screen.MousePointer = 11
            If IsNumeric(tUsuario.Text) Then
                cons = "Select * from Usuario Where UsuDigito = " & Trim(tUsuario.Text)
            Else
                cons = "Select UsuCodigo, UsuIdentificacion as 'Identificación', Convert(Char(6), UsuDigito) as 'Dígito' " & _
                           " from Usuario Where UsuIdentificacion like  '" & Trim(tUsuario.Text) & "%'" & _
                           " Order by UsuIdentificacion"
            End If
            
            Dim aId As Long, aQ As Integer, aTSel As String
            aId = 0: aQ = 0
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then
                aId = rsAux!UsuCodigo: aTSel = Trim(rsAux(1))
                aQ = 1: rsAux.MoveNext
                If Not rsAux.EOF Then aQ = 2
            End If
            rsAux.Close
            
            Select Case aQ
                Case 0: MsgBox "No hay usuarios para el dígito/identificación ingresada.", vbExclamation, "No hay datos."
                
                Case 2:
                        aId = 0
                        Dim miLista As New clsListadeAyuda
                        miLista.ActivoListaAyuda cons, False, txtConexion, 4100
                        Me.Refresh
                        If miLista.ValorSeleccionado > 0 Then
                            aId = miLista.ValorSeleccionado
                            aTSel = Trim(miLista.ItemSeleccionado)
                        End If
                        Set miLista = Nothing
            End Select
            
            If aId > 0 Then
                tUsuario.Text = aTSel: tUsuario.Tag = aId
                Foco cSuceso
            End If
            
            Screen.MousePointer = 0
    End Select
    Exit Sub
    
errBuscar:
    clsGeneral.OcurrioError "Error al buscar el usuario.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub vsConsulta_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    On Error Resume Next
    
    Dim aSize As Currency
    With vsConsulta
        'For I = 0 To .Cols - 3: aSize = aSize + .ColWidth(I): Next I
        '.ColWidth(.Cols - 2) = .Width - (aSize + .ColWidth(.Cols - 1) + 200)
        
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 2, , False
    End With
    
End Sub

Private Sub AccionLimpiar()
    On Error Resume Next
    tDesde.FechaValor = Now
    tHasta.FechaValor = Now
    tHora1.Text = "00:00": tHora2.Text = "23:59"
    cSuceso.Text = ""
    tDescripcion.Text = ""
    tDocumento.Text = "": lDoc.Caption = "": lDoc.Tag = 0
    
    tCi.Text = "": tRuc.Text = ""
    lCliente.Tag = 0: lCliente.Caption = ""
    tCi.Visible = True: tRuc.Visible = False
    
End Sub

Private Sub PropiedadesImpresion()

  On Error Resume Next
  With vsListado
        .PaperSize = vbPRPSLetter
        .PhysicalPage = True
        .Orientation = orPortrait
        
        .PreviewMode = pmPrinter
        
        .PreviewPage = 1
        .FontName = "Tahoma": .FontSize = 10: .FontBold = False: .FontItalic = False
        .Zoom = 100
        .MarginBottom = 500: .MarginTop = 500
        .MarginRight = 350: .MarginLeft = 350
    End With

End Sub

Private Sub vsConsulta_DblClick()
    If vsConsulta.Rows = 1 Then Exit Sub
    Screen.MousePointer = 11
    
    frmDetalle.prm_Suceso = vsConsulta.Cell(flexcpValue, vsConsulta.Row, 0)
    frmDetalle.Show vbModal, Me
    Me.Refresh
    
    Screen.MousePointer = 0
    
End Sub

Private Sub vsConsulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo errBD
    With vsConsulta
        If .Rows = 1 Then Exit Sub
        If Button <> vbRightButton Then Exit Sub
        
        If .Cell(flexcpData, .Row, 0) = 0 Then MnuFactura.Enabled = False Else MnuFactura.Enabled = True
        If .Cell(flexcpData, .Row, 1) = 0 Then MnuCliente.Enabled = False Else MnuCliente.Enabled = True
        MnuComentarios.Enabled = MnuCliente.Enabled
        MnuSuCliente.Enabled = MnuCliente.Enabled
        
        PopupMenu MnuBDerecho, , , , MnuIrA
    End With
    
errBD:
End Sub

Private Sub BuscarCliente(Optional miCi As String = "", Optional miRuc As String = "", Optional miId As Long = 0, Optional miTipo As Integer = 0)
    
    If miCi <> "" Then
        cons = "Select Cliente.*, (RTrim(CPeNombre1) + ' ' + RTrim(isnull(CPeNombre2, '')) + ' ' + RTrim(CPeApellido1) + ' ' + RTrim(isnull(CPeApellido2, '')))  as Nombre" & _
                   " From CPersona, Cliente" & _
                   " Where CliCodigo = CPeCliente And CliCiRuc = '" & Trim(miCi) & "'"
    End If
    
    If miRuc <> "" Then
        cons = "Select Cliente.*, (RTrim(CEmNombre) + ' (' + RTrim(isnull(CEmFantasia, '')) + ')')  as Nombre" & _
                   " From CEmpresa, Cliente" & _
                   " Where CliCodigo = CEmCliente And CliCiRuc = '" & Trim(miRuc) & "'"
    End If
    
    If miId <> 0 And miTipo <> 0 Then
        If miTipo = TipoCliente.Cliente Then
             cons = "Select Cliente.*, (RTrim(CPeNombre1) + ' ' + RTrim(isnull(CPeNombre2, '')) + ' ' + RTrim(CPeApellido1) + ' ' + RTrim(isnull(CPeApellido2, '')))  as Nombre" & _
                        " From CPersona, Cliente Where CliCodigo = CPeCliente And CliCodigo = " & miId
        End If
        If miTipo = TipoCliente.Empresa Then
            cons = "Select Cliente.*, (RTrim(CEmNombre) + ' (' + RTrim(isnull(CEmFantasia, '')) + ')')  as Nombre" & _
                      " From CEmpresa, Cliente Where CliCodigo = CEmCliente And CliCodigo = " & miId
        End If
    End If
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If rsAux!CliTipo = 1 Then
            If Not IsNull(rsAux!CliCiRuc) Then tCi.Text = Trim(rsAux!CliCiRuc)
            tRuc.Visible = False: tCi.Visible = True: tCi.SetFocus
        End If
        If rsAux!CliTipo = 2 Then
            If Not IsNull(rsAux!CliCiRuc) Then tRuc.Text = Trim(rsAux!CliCiRuc)
            tCi.Visible = False: tRuc.Visible = True: tRuc.SetFocus
        End If
        
        lCliente.Caption = Trim(rsAux!Nombre)
        lCliente.Tag = rsAux!CliCodigo
    Else
        lCliente.Caption = " No Existe !!"
    End If
    rsAux.Close
End Sub
