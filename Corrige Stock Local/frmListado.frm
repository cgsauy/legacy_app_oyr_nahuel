VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListado 
   Caption         =   "Corregir Stock Total"
   ClientHeight    =   7530
   ClientLeft      =   1950
   ClientTop       =   2070
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   10020
   Begin MSComctlLib.ImageList imgList 
      Left            =   180
      Top             =   5820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":0442
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4035
      Left            =   1980
      TabIndex        =   0
      Top             =   1320
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   7117
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
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
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
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11415
      _Version        =   196608
      _ExtentX        =   20135
      _ExtentY        =   7858
      _StockProps     =   229
      BorderStyle     =   1
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
      PreviewMode     =   1
      Zoom            =   70
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   11595
      TabIndex        =   3
      Top             =   6720
      Width           =   11655
      Begin VB.CommandButton bHelp 
         Height          =   310
         Left            =   5340
         Picture         =   "frmListado.frx":075C
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Ayuda"
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0A66
         Height          =   310
         Left            =   4500
         Picture         =   "frmListado.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4140
         Picture         =   "frmListado.frx":109A
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   3180
         Picture         =   "frmListado.frx":1514
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2820
         Picture         =   "frmListado.frx":15FE
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   2220
         Picture         =   "frmListado.frx":16E8
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3780
         Picture         =   "frmListado.frx":1922
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4860
         Picture         =   "frmListado.frx":1A24
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5700
         Picture         =   "frmListado.frx":1DEA
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":1EEC
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmListado.frx":21EE
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1860
         Picture         =   "frmListado.frx":2530
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmListado.frx":2832
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin MSComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6180
         TabIndex        =   16
         Top             =   135
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lQ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   150
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7275
      Width           =   10020
      _ExtentX        =   17674
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
            Object.Width           =   9472
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuPopUp 
      Caption         =   "MnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu MnuAccesos 
         Caption         =   "Menú Accesos"
      End
      Begin VB.Menu MnuAccL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuStockTotal 
         Caption         =   "Stock &Total"
      End
      Begin VB.Menu MnuStockBalance 
         Caption         =   "Stock &Balance"
      End
      Begin VB.Menu MnuPlBalance 
         Caption         =   "Ver &Plantilla Balance"
      End
      Begin VB.Menu MnuCorregirUno 
         Caption         =   "&Corregir Artículo"
      End
      Begin VB.Menu MnuCorregirUnoDel 
         Caption         =   "Corregir Artículo y &Eliminar"
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "Eliminar sin Corregir"
      End
      Begin VB.Menu MnuAccL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMaDiferencias 
         Caption         =   "Marcar Dif < a ..."
      End
      Begin VB.Menu MnuCorregirTodos 
         Caption         =   "Corregir Todos"
      End
      Begin VB.Menu MnuAccL3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVerTodos 
         Caption         =   "Ver Todos"
      End
      Begin VB.Menu MnuEstado 
         Caption         =   "MnuEstado"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs1 As rdoResultset
Private aTexto As String

Dim bCargarImpresion As Boolean
Dim aQEstados As Integer

Dim aRow As Long, aValor As Long

Private Type typStock
    Articulo As Long
    Stock As String
End Type

Private Type typLocal
    LocalTipo As Integer
    LocalName As String
    Local As Long
    Estado As Integer
    EstadoName As String
    Q As Long
End Type

Private m_Error As String

Private mColLIFO As Integer

Private Sub AccionLimpiar()
    
    vsConsulta.Rows = 2
    lQ.Caption = ""
    
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub bHelp_Click()
    AccionMenuHelp
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

Private Sub bConfigurar_Click()
    AccionConfigurar
End Sub

Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub

Private Sub chVista_Click()
    
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
    Else
        AccionImprimir
        vsListado.ZOrder 0
    End If

End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()

    On Error GoTo ErrLoad
    FechaDelServidor
    bHelp.Picture = imgList.ListImages("help").ExtractIcon
    
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    
    AccionLimpiar
    
    InicializoGrillas
    
    bCargarImpresion = True
    With vsListado
        .PhysicalPage = True
        .PaperSize = 1
        .Orientation = orLandscape ' orPortrait
        .Zoom = 100
        .MarginLeft = 900: .MarginRight = 350
        .MarginBottom = 750: .MarginTop = 750
    End With
    
    Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
                    
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()
 On Error Resume Next
    aQEstados = 0
    Dim aENames As String: aENames = ""
        
    cons = "Select * from EstadoMercaderia"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        aENames = aENames & rsAux!EsMCodigo & ":" & rsAux!EsMAbreviacion
        
        If aQEstados > 0 Then Load MnuEstado(aQEstados)
        With MnuEstado(aQEstados)
            .Caption = "Ver " & Trim(rsAux!EsMNombre)
            .Tag = rsAux!EsMCodigo
            .Visible = True
        End With
        
        aQEstados = aQEstados + 1
        rsAux.MoveNext
        If Not rsAux.EOF Then aENames = aENames & "|"
    Loop
    rsAux.Close
    
    
    With vsConsulta
        .Editable = True
        .OutlineBar = flexOutlineBarNone ' = flexOutlineBarComplete
        .OutlineCol = 0
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        
        .Rows = 1
        .Cols = aQEstados * 3 + 5 + 1
        .ColWidth(.Cols - 1) = 100
        '.Cell(flexcpText, .Rows - 1, 0) = "chk"
        
        .Cell(flexcpText, .Rows - 1, 2) = "Sistema (Locales) " & PropiedadesConnect(cBase.Connect, Database:=True)
        .Cell(flexcpText, .Rows - 1, 2 + aQEstados + 1) = "Se Contó (Balance) " & PropiedadesConnect(cBase.Connect, Database:=True)
        .Cell(flexcpText, .Rows - 1, 2 + (aQEstados + 1) * 2) = "Faltan (Conteo - Sistema)"
                        
        .ColWidth(0) = 300
        .ColWidth(1) = 2800
        
        Dim arrEstados() As String, arrData() As String, Idx As Integer
        arrEstados = Split(aENames, "|")
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = "Artículos"
        Dim mAncho As Long
        
        For Idx = LBound(arrEstados) To UBound(arrEstados)
            
            arrData = Split(arrEstados(Idx), ":")
            If Val(arrData(0)) = paEstadoArticuloEntrega Then mAncho = 750 Else mAncho = 550
            .Cell(flexcpText, .Rows - 1, 2 + Idx) = Trim(arrData(1))
            .Cell(flexcpData, .Rows - 1, 2 + Idx) = arrData(0)
            .ColWidth(2 + Idx) = mAncho
            .Cell(flexcpBackColor, 0, 2 + Idx, .Rows - 1) = Colores.clNaranja
            
            .Cell(flexcpText, .Rows - 1, 2 + aQEstados + 1 + Idx) = Trim(arrData(1))
            .Cell(flexcpData, .Rows - 1, 2 + aQEstados + 1 + Idx) = arrData(0)
            .ColWidth(2 + aQEstados + 1 + Idx) = mAncho
            .Cell(flexcpBackColor, 0, 2 + aQEstados + 1 + Idx, .Rows - 1) = Colores.Gris
            
            .Cell(flexcpText, .Rows - 1, 2 + (aQEstados + 1) * 2 + Idx) = Trim(arrData(1))
            .Cell(flexcpData, .Rows - 1, 2 + (aQEstados + 1) * 2 + Idx) = arrData(0)
            .ColWidth(2 + (aQEstados + 1) * 2 + Idx) = mAncho
            .Cell(flexcpBackColor, 0, 2 + (aQEstados + 1) * 2 + Idx, .Rows - 1) = Colores.Blanco
            
            If Idx = UBound(arrEstados) Then
                .Cell(flexcpText, .Rows - 1, 2 + Idx + 1) = "Total"
                .Cell(flexcpData, .Rows - 1, 2 + Idx + 1) = -1
                .Cell(flexcpBackColor, 0, 2 + Idx + 1, .Rows - 1) = Colores.clNaranja
                
                .Cell(flexcpText, .Rows - 1, 2 + aQEstados + 1 + Idx + 1) = "Total"
                .Cell(flexcpData, .Rows - 1, 2 + aQEstados + 1 + Idx + 1) = -1
                .Cell(flexcpBackColor, 0, 2 + aQEstados + 1 + Idx + 1, .Rows - 1) = Colores.Gris
                
                .Cell(flexcpText, .Rows - 1, 2 + (aQEstados + 1) * 2 + Idx + 1) = "Total"
                .Cell(flexcpData, .Rows - 1, 2 + (aQEstados + 1) * 2 + Idx + 1) = -1
                .Cell(flexcpBackColor, 0, 2 + (aQEstados + 1) * 2 + Idx + 1, .Rows - 1, .Cols - 1) = Colores.Blanco
            End If
        Next
        
        
        '.Cell(flexcpBackColor, 0, 0, .Rows - 1, 1) = .BackColorFixed
        '.Cell(flexcpBackColor, 0, 2, .Rows - 1, 6) = Colores.clNaranja
        '.Cell(flexcpBackColor, 0, 7, .Rows - 1, 11) = Colores.Gris
        '.Cell(flexcpBackColor, 0, 12, .Rows - 1, 12) = Colores.Blanco
        '.Cell(flexcpBackColor, 0, 13, .Rows - 1, 13) = Colores.clCeleste
        '.Cell(flexcpBackColor, 0, 14, .Rows - 1, 14) = Colores.Obligatorio
        '.Cell(flexcpBackColor, 0, 15, .Rows - 1, 16) = Colores.Blanco
        
                
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridFlatHorz
        .FixedRows = 2
        .WordWrap = False: .MergeCells = flexMergeSpill
        
        .Cols = .Cols + 1
        mColLIFO = .Cols - 1
        .Cell(flexcpText, 0, mColLIFO) = PropiedadesConnect(cBaseMov.Connect, Database:=True)
        .Cell(flexcpText, .Rows - 1, .Cols - 1) = "LIFO"
        .Cell(flexcpData, .Rows - 1, .Cols - 1) = 0
        .ColWidth(.Cols - 1) = mAncho
        .ColAlignment(.Cols - 1) = flexAlignRightCenter
        
        
        .Cols = .Cols + 1
        .Cell(flexcpText, .Rows - 1, .Cols - 1) = "Sist-LIFO"
        .Cell(flexcpData, .Rows - 1, .Cols - 1) = 0
        .ColWidth(.Cols - 1) = 825
        .ColAlignment(.Cols - 1) = flexAlignRightCenter
        
        .ExtendLastCol = False
    End With
    
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            
            Case vbKeyE: AccionConsultar
            
            Case vbKeyP: IrAPagina vsListado, 1
            Case vbKeyA: IrAPagina vsListado, vsListado.PreviewPage - 1
            Case vbKeyS: IrAPagina vsListado, vsListado.PreviewPage + 1
            Case vbKeyU: IrAPagina vsListado, vsListado.PageCount
            
            Case vbKeyAdd: Zoom vsListado, vsListado.Zoom + 5
            Case vbKeySubtract: Zoom vsListado, vsListado.Zoom - 5
            
            Case vbKeyQ: AccionLimpiar
            Case vbKeyI: AccionImprimir True
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            Case vbKeyC: AccionConfigurar
            
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11

    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
    'fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Left = 60
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = vsListado.Height
    vsConsulta.Left = vsListado.Left
    
    picBotones.Width = vsListado.Width
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 150
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexionBDMov
    
    CierroConexion
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub AccionConsultar()
On Error GoTo errConsultar
    
'    If Not ValidoCampos Then Exit Sub
    Screen.MousePointer = 11
    

    bCargarImpresion = True
    lQ.Tag = 0: lQ.Caption = ""
    
    vsConsulta.Rows = vsConsulta.FixedRows
    vsConsulta.Refresh
    vsConsulta.Redraw = False
    
    CargoTodosLosArticulos
    
    CargoStockConteo
    
    CargoStockTotal
    
    CargoStockLIFO
    
    CargoDifrenciasStock
    With vsConsulta
        .Select 1, 1, 1, 0
        .Sort = flexSortGenericAscending
        .Redraw = True
    End With
    
    EliminoArticulosArchivo
    
    pbProgreso.Value = 0
    
    'CargoStockLIFO
    
    EliminoArticulosEnCero
    Status.Panels("help").Text = ""

    MarcoArticulosConDifCero

    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    vsConsulta.Redraw = True
    clsGeneral.OcurrioError "Error al realizar la consulta de datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoStockConteo()

    Status.Panels("help").Text = "Cargando Artículos del Conteo ...": Status.Refresh
    
    Dim aQ As Long
    aQ = 0
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    cons = "Select Count(Distinct(BMRArticulo)) from BMRenglon, BMLocal" & _
               " Where BMRIDBML = BMLID  And BMLCodigo = 1"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux(0)) Then aQ = rsAux(0)
    rsAux.Close
    
    If aQ = 0 Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        Screen.MousePointer = 0:  Exit Sub
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------
    
    pbProgreso.Min = 0: pbProgreso.Value = 0
    pbProgreso.Max = aQ
    
    'Cargo Artículos del STOCK CONTEO  ---------------------------------------------------------------------------------------------------------
    cons = "Select ArtID, ArtCodigo, ArtNombre, BMREstado, Sum(BMRCantidad) as Q" & _
                " From BMRenglon, Articulo, BMLocal" & _
                " Where BMRArticulo = ArtID " & _
                " And BMRIDBML = BMLID  And BMLCodigo = 1" ' & _
                " And ArtTipo <> " & paTipoArticuloServicio
    
    cons = cons & " Group by ArtID, ArtCodigo, ArtNombre, BMREstado " & _
                " Order by ArtID"
                
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Dim aID As Long: aID = 0
    Dim aST As Long: aST = 0
    Dim mCol As Integer
    With vsConsulta
        Do While Not rsAux.EOF
            
            If aID <> rsAux!ArtID Then
                aID = rsAux!ArtID
                pbProgreso.Value = pbProgreso.Value + 1
    
                aST = 0
'                .AddItem ""
'                aRow = .Rows - 1
'                .Cell(flexcpChecked, aRow, 0) = flexUnchecked
'                .Cell(flexcpText, aRow, 1) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
'                aValor = rsAux!ArtID: .Cell(flexcpData, aRow, 1) = aValor
                
                aRow = RowEnLista(aID)  'Cambio 21/9/02
                If aRow = 0 Then MsgBox "Fila 0: ID Artículo:" & aID
                
            End If
            
            mCol = ColEstado(2, rsAux!BMREstado)
            .Cell(flexcpText, aRow, mCol) = Format(rsAux!Q, "#,##0")
                        
            aST = aST + rsAux!Q
            mCol = ColEstado(2, -1)
            .Cell(flexcpText, aRow, mCol) = Format(aST, "#,##0")
            
            rsAux.MoveNext
        Loop
        rsAux.Close
    
        .Cell(flexcpBackColor, 2, 2 + aQEstados + 1, .Rows - 1, 2 + (aQEstados * 2) + 1) = Colores.Gris
    End With
    
End Sub

Private Sub CargoStockTotal()

    Status.Panels("help").Text = "Consultando datos Stock Total ...": Status.Refresh
    'Cargo Artículos del STOCK ACTUAL   ---------------------------------------------------------------------------------------------------------
    'cons = "Select Count(Distinct(StLArticulo)) from StockLocal" & _
               " Where StLArticulo Not In (Select ArtId from Articulo Where ArtTipo = " & paTipoArticuloServicio & ")"
    Dim aQ As Long, arrArticulos() As typStock
    aQ = 0
    cons = "Select Count(Distinct(StLArticulo)) from StockLocal" '& _
               " Where StLArticulo  In (Select BMRArticulo from BMRenglon, BMLocal Where BMRIDBML = BMLID  And BMLCodigo = 1 )"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux(0)) Then aQ = rsAux(0)
    rsAux.Close
    If aQ = 0 Then Exit Sub
    
    ReDim arrArticulos(aQ - 1)
    Dim Idx As Long, aID As Long
    
    pbProgreso.Min = 0: pbProgreso.Value = 0
    pbProgreso.Max = aQ
    
    cons = "Select ArtID, ArtCodigo, ArtNombre, StLEstado, Sum(StLCantidad) as Q " & _
               " From StockLocal, Articulo" & _
               " Where StLArticulo = ArtID " '& _
               " AND StLArticulo  In (Select BMRArticulo from BMRenglon, BMLocal Where BMRIDBML = BMLID  And BMLCodigo = 1 )"
               '" And ArtTipo <> " & paTipoArticuloServicio
            
    cons = cons & " Group by ArtID, ArtCodigo, ArtNombre, StLEstado" & _
                          " Order by ArtID"
                
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    aID = 0
    Do While Not rsAux.EOF
        If aID <> rsAux!ArtID Then
            aRow = 0
            If aID <> 0 Then Idx = Idx + 1
            aID = rsAux!ArtID
            arrArticulos(Idx).Articulo = aID
            pbProgreso.Value = pbProgreso.Value + 1
        End If
        
        With vsConsulta
            arrArticulos(Idx).Stock = Trim(arrArticulos(Idx).Stock) & rsAux!StLEstado & ":" & rsAux!Q & "|"
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    pbProgreso.Value = 0
    
    Status.Panels("help").Text = "Procesando Stock Total ...": Status.Refresh
    
    'Cargo los datos de la tabla --------------------------------------------------------------------------------
    Dim mRow As Long, idxE As Integer, mEstado As Long, mQ As Long, aST As Long
    Dim arrEstados() As String, arrData() As String
    
    For Idx = LBound(arrArticulos) To UBound(arrArticulos)
        aST = 0
        With arrArticulos(Idx)
            mRow = RowEnLista(.Articulo)
            If mRow = 0 Then MsgBox "Fila 0: ID Artículo:" & .Articulo & vbCrLf & "Stock: " & .Stock
            arrEstados = Split(.Stock, "|")
        End With
        
        For idxE = LBound(arrEstados) To UBound(arrEstados) - 1
            arrData = Split(arrEstados(idxE), ":")
            mEstado = arrData(0)
            mQ = arrData(1)
            
            With vsConsulta
            .Cell(flexcpText, mRow, ColEstado(1, mEstado)) = Format(mQ, "#,##0")
            
            aST = aST + mQ
            .Cell(flexcpText, mRow, ColEstado(1, -1)) = Format(aST, "#,##0")
            End With
        Next
        pbProgreso.Value = pbProgreso.Value + 1
    Next
    
    With vsConsulta
        .Cell(flexcpBackColor, 2, 2, .Rows - 1, 2 + aQEstados) = Colores.clNaranja
    End With
    
    ReDim arrArticulos(0)
    
    pbProgreso.Value = 0
    
End Sub

Private Sub CargoDifrenciasStock()
    
    Dim aCol As Integer
    Dim aQM As Long
    
    With vsConsulta
        For I = 2 To .Rows - 1
            If Trim(.Cell(flexcpText, I, 2)) = "" Then .Cell(flexcpText, I, 2) = " "
            For aCol = 2 To 2 + aQEstados
                aQM = 0
                aQM = .Cell(flexcpValue, I, aCol)
                aQM = .Cell(flexcpValue, I, aCol + aQEstados + 1) - aQM
                
                .Cell(flexcpText, I, aCol + (aQEstados + 1) * 2) = Format(aQM, "#,##0")
                
                If (aCol - 2) = aQEstados Then
                    aQM = .Cell(flexcpValue, I, aCol)
                    aQM = aQM - .Cell(flexcpValue, I, mColLIFO)
                    .Cell(flexcpText, I, mColLIFO + 1) = Format(aQM, "#,##0")
                End If
            Next
        Next
        .Cell(flexcpBackColor, 2, 2 + (aQEstados * 2) + 2, .Rows - 1, 2 + (aQEstados * 3) + 2 + 1) = Colores.Blanco
    End With
    
    
End Sub

Private Sub MarcoArticulosConDifCero()
' Cuando lista marcar con gris los q no tienen diferencia entre Total sist, total cdo y lifo.

    Dim aCol As Integer
    Dim aQM As Long
    
    Dim mColTotal As Integer
    mColTotal = 2 + (aQEstados + 1) * 2 + aQEstados
    
    With vsConsulta
        For I = 2 To .Rows - 1
            If .Cell(flexcpValue, I, mColTotal) = .Cell(flexcpValue, I, mColLIFO + 1) And .Cell(flexcpValue, I, mColTotal) = 0 Then
                .Cell(flexcpBackColor, I, 0, , .Cols - 1) = Colores.GrisOscuro
            End If
        Next
        '.Cell(flexcpBackColor, 2, 2 + (aQEstados * 2) + 2, .Rows - 1, 2 + (aQEstados * 3) + 2 + 1) = Colores.Blanco
    End With

End Sub

Private Sub MnuCorregirTodos_Click()
On Error GoTo errDif
    
    If MsgBox("Confirma corregir el stock de todos los artículos seleccionados.", vbQuestion + vbYesNo + vbDefaultButton2, "Corregir Todos") = vbNo Then Exit Sub
    
    Dim I As Long
    Screen.MousePointer = 11
    With vsConsulta
        pbProgreso.Value = 0
        pbProgreso.Max = .Rows - 2
        For I = 2 To .Rows - 1
            If .Cell(flexcpChecked, I, 0) = flexChecked Then
                gb_CorrigoStock .Cell(flexcpData, I, 1), I, withMsg:=False, bDeAUno:=False
            End If
            pbProgreso.Value = pbProgreso.Value + 1
            If (I Mod 5) = 0 Then Me.Refresh
        Next
        
        'Dim xItem As Long
        'xItem = 2
        'For I = 2 To .Rows - 1
        '    If .Cell(flexcpChecked, xItem, 0) = flexChecked Then
        '        .RemoveItem xItem
        '    Else
        '        xItem = xItem + 1
        '    End If
        'Next
        
        pbProgreso.Value = 0
    End With
    Screen.MousePointer = 0
    Exit Sub

errDif:
    clsGeneral.OcurrioError "Error al realizar la corrección global.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuCorregirUno_Click()
    On Error GoTo errDif
    
    gb_CorrigoStock vsConsulta.Cell(flexcpData, vsConsulta.Row, 1), vsConsulta.Row, , bDeAUno:=True
    
     Exit Sub

errDif:
    clsGeneral.OcurrioError "Error al realizar la corrección.", Err.Description
    If Trim(m_Error) <> "" Then MsgBox m_Error, vbCritical, "Datos Error"
    Screen.MousePointer = 0
End Sub

Private Sub MnuCorregirUnoDel_Click()
On Error GoTo errDif
    
    gb_CorrigoStock vsConsulta.Cell(flexcpData, vsConsulta.Row, 1), vsConsulta.Row, , bDeAUno:=True, bDelete:=True
    
     Exit Sub

errDif:
    clsGeneral.OcurrioError "Error al realizar la corrección.", Err.Description
    If Trim(m_Error) <> "" Then MsgBox m_Error, vbCritical, "Datos Error"
    Screen.MousePointer = 0
End Sub

Private Sub MnuDelete_Click()

    If MsgBox("Confirma eliminar el artículo sin corregirlo ?", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar SIN Corregir !!") = vbNo Then Exit Sub
    On Error GoTo errDel
    Screen.MousePointer = 11
    
    Dim mId As Long, idRow As Long
    
    idRow = vsConsulta.Row
    mId = vsConsulta.Cell(flexcpData, idRow, 1)
    
    fileAddArticulo mId
    vsConsulta.RemoveItem idRow
    Screen.MousePointer = 0
    Exit Sub
    
errDel:
    clsGeneral.OcurrioError "Error al eliminar el artículo sin corregir.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuEstado_Click(Index As Integer)
    VerEstados Val(MnuEstado(Index).Tag)
End Sub

Private Sub MnuMaDiferencias_Click()
    On Error GoTo errDif
    Dim bRet As Long
    
    bRet = Val(InputBox("Ingrese la diferencia máxima a considerar." & vbCrLf & vbCrLf & _
                            "(*) Selecciona los artículos con diferencias de stock, en valor absoluto, menores al número ingresado.", "Marcar Diferencias Menores A ..."))
    
    Screen.MousePointer = 11
    Dim aCol As Integer
    Dim aQM As Long
    Dim bOk As Boolean
    With vsConsulta
        For I = 2 To .Rows - 1
            bOk = True
            For aCol = 2 + (aQEstados + 1) * 2 To .Cols - 3
                If Abs(.Cell(flexcpValue, I, aCol)) >= bRet Then
                    bOk = False
                    Exit For
                End If
            Next
            If bOk Then .Cell(flexcpChecked, I, 0) = flexChecked
        Next
    End With
    Screen.MousePointer = 0
    Exit Sub

errDif:
    clsGeneral.OcurrioError "Error al marcar las diferencias.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuPlBalance_Click()
On Error Resume Next
    Dim aIdReg As String
    
    aIdReg = vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)
    If aIdReg = "" Then Exit Sub
    
    Screen.MousePointer = 11
    
    cons = "Select ArtCodigo from Articulo Where ArtId = " & aIdReg
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then aIdReg = rsAux!ArtCodigo
    rsAux.Close
    
    EjecutarApp prmPathApp & "\appExploreMsg.exe ", prmPlBalance & ":" & aIdReg
    
    Screen.MousePointer = 0
End Sub

Private Sub MnuStockBalance_Click()
On Error Resume Next
    With vsConsulta
        EjecutarApp prmPathApp & "\Stock Balance.exe", .Cell(flexcpData, .Row, 1)
    End With
    
End Sub

Private Sub MnuStockTotal_Click()
On Error Resume Next
    With vsConsulta
        EjecutarApp prmPathApp & "\Stock Total.exe", .Cell(flexcpData, .Row, 1)
    End With
            
End Sub

Private Sub MnuVerTodos_Click()
    VerEstados 0
End Sub

Private Sub vsConsulta_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
    If vsConsulta.Rows = 2 Then Cancel = True
End Sub

Private Sub vsConsulta_DblClick()
    
    With vsConsulta
        If Not (.Rows > .FixedRows) Then Exit Sub
        Select Case .Col
            Case 1: gb_CorrigoStock .Cell(flexcpData, .Row, 1), .Row, bDeAUno:=True
            Case 2 To 2 + aQEstados: EjecutarApp prmPathApp & "\Stock Total.exe", .Cell(flexcpData, .Row, 1)
            
            Case 2 + aQEstados + 1 To 2 + (aQEstados * 2 + 1): EjecutarApp prmPathApp & "\Stock Balance.exe", .Cell(flexcpData, .Row, 1)
        End Select
        
    End With
    
End Sub

Private Sub vsConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If KeyCode = 93 Then
        With vsConsulta
            If Not (.Rows > .FixedRows) Then Exit Sub
            Dim xY As Single, xX As Single
            xY = .Top + .RowHeight(0) * .Row
            xX = .Left + 500
            PopupMenu MnuPopUp, , xX, xY, MnuAccesos
            
        End With
    End If
    
End Sub

Private Sub vsConsulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
        With vsConsulta
            If Not (.Rows > .FixedRows) Then Exit Sub
            PopupMenu MnuPopUp, , , , MnuAccesos
            
        End With
    End If
    
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    If bCargarImpresion Then
        If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub
        With vsListado
            .Columns = 1
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        aTexto = "Corregir Stock Total"
        EncabezadoListado vsListado, aTexto, False
        vsListado.FileName = "Corregir Stock Total"
         
        With vsConsulta
            '.Redraw = False
            '.FontSize = 6
            'AnchoEncabezado Impresora:=True
            vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
            'AnchoEncabezado Pantalla:=True
            '.FontSize = 8
            '.Redraw = True
        End With
        
        vsListado.EndDoc
        bCargarImpresion = False
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
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al realizar la impresión", Err.Description
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

Private Sub vsListado_NewPage()
    lQ.Tag = Val(lQ.Tag) + 1
    lQ.Caption = lQ.Tag: lQ.Refresh
End Sub

Private Function RowEnLista(idArticulo As Long) As Long
    RowEnLista = 0
    Dim I As Long
    
    With vsConsulta
        For I = 2 To .Rows - 1
            If idArticulo = .Cell(flexcpData, I, 1) Then
                RowEnLista = I: Exit For
            End If
        Next
    End With
    
End Function

Private Function ColEstado(IdTipo As Integer, idEstado As Long) As Integer

    ColEstado = 0
    Dim I As Long, aDesde As Integer
    Select Case IdTipo
        Case 1: aDesde = 2
        Case Else: aDesde = 2 + (aQEstados * (IdTipo - 1)) + 1
    End Select
    
    With vsConsulta
        For I = aDesde To .Cols - 1
            If idEstado = .Cell(flexcpData, 1, I) Then
                ColEstado = I: Exit For
            End If
        Next
    End With

End Function

Private Sub VerEstados(idEstado As Long)

    With vsConsulta
        For I = 2 To .Cols - 2
            If idEstado <> 0 Then
                If .Cell(flexcpData, 1, I) = idEstado Or .Cell(flexcpData, 1, I) = -1 Then
                    .ColHidden(I) = False
                Else
                    .ColHidden(I) = True
                End If
            Else
                .ColHidden(I) = False
            End If
            
        Next
        .WordWrap = False: .MergeCells = flexMergeSpill
    End With
    
End Sub

Private Sub gb_CorrigoStock(idArticulo As Long, idRow As Long, Optional withMsg As Boolean = True, _
            Optional bDeAUno As Boolean = False, Optional bDelete As Boolean = False)

Dim arrSTConteo() As typLocal
Dim Idx As Long

    If Not withMsg And Not bDeAUno Then
        Status.Panels("help").Text = "Consultando stock a " & vsConsulta.Cell(flexcpText, idRow, 1) & " ..."
        Status.Refresh
    End If
    
    Idx = 0
    ReDim Preserve arrSTConteo(Idx)
    
    cons = "Select LocNombre, LocTipo, BMLLocal, BMREstado, EsMNombre, Sum(BMRCantidad) as Q" & _
               " From BMRenglon, BMLocal, Local, EstadoMercaderia " & _
               " Where BMRIDBML = BMLID " & _
               " And BMLLocal = LocCodigo " & _
               " And BMREstado = EsMCodigo " & _
               " And BMLCodigo = 1" & _
               " And BMRArticulo = " & idArticulo & _
               " Group by LocNombre, LocTipo, BMLLocal, BMREstado, EsMNombre " & _
               " Order by BMLLocal"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        ReDim Preserve arrSTConteo(Idx)
        With arrSTConteo(Idx)
            .LocalTipo = rsAux!LocTipo
            .LocalName = Trim(rsAux!LocNombre)
            .Local = rsAux!BMLLocal
            .Estado = rsAux!BMREstado
            .EstadoName = Trim(rsAux!EsMNombre)
            .Q = rsAux!Q
        End With
        
        Idx = Idx + 1
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    Dim bOk As Boolean
    Idx = 0
    cons = "Select LocNombre, LocTipo, StLLocal, StLEstado, EsMNombre, Sum(StLCantidad) as Q " & _
               " From StockLocal, Local, EstadoMercaderia " & _
               " Where StLArticulo = " & idArticulo & _
               " And StLLocal = LocCodigo " & _
               " And StLEstado = EsMCodigo" & _
               " Group by LocNombre, LocTipo, StLLocal, StLEstado, EsMNombre " & _
               " Order by StLLocal"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        bOk = False
        For I = LBound(arrSTConteo) To UBound(arrSTConteo)
            With arrSTConteo(I)
                If .Local = rsAux!StlLocal And .Estado = rsAux!StLEstado Then
                    .Q = .Q - rsAux!Q
                    bOk = True
                End If
            End With
        Next
        If Not bOk Then
            Idx = UBound(arrSTConteo) + 1
            ReDim Preserve arrSTConteo(Idx)
            With arrSTConteo(Idx)
                .LocalTipo = rsAux!LocTipo
                .LocalName = Trim(rsAux!LocNombre)
                .Local = rsAux!StlLocal
                .Estado = rsAux!StLEstado
                .EstadoName = Trim(rsAux!EsMNombre)
                .Q = 0 - rsAux!Q
            End With
        End If
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    Dim bCorregir As Boolean: bCorregir = False
    If withMsg Then
    
        Dim aMsg As String: aMsg = ""
        For I = LBound(arrSTConteo) To UBound(arrSTConteo)
            With arrSTConteo(I)
                 If .Q <> 0 Then
                    
                    Select Case .Q
                        Case Is < 0: aMsg = aMsg & " Sacar " & vbTab
                        Case Is > 0:  aMsg = aMsg & " Agregar " & vbTab
                    End Select
               
                    aMsg = aMsg & Abs(.Q) & " " & vbTab & .EstadoName & "   a " & .LocalName & vbCrLf
                End If
                
            End With
        Next
    
        If Trim(aMsg) = "" Then
            MsgBox "Artículo " & vsConsulta.Cell(flexcpText, idRow, 1) & vbCrLf & _
                        "El stock está OK !!", vbInformation, "Movimientos " & vsConsulta.Cell(flexcpText, idRow, 1)
            If bDelete Then
                fileAddArticulo idArticulo
                vsConsulta.RemoveItem idRow
            End If
        Else
            If MsgBox(aMsg & vbCrLf & _
                        "Quiere corregir el Stock ahora ? ", vbInformation + vbYesNo + vbDefaultButton2, "Movimientos " & vsConsulta.Cell(flexcpText, idRow, 1)) = vbYes Then bCorregir = True
        End If
    
    Else
        bCorregir = True
    End If
    
    m_Error = ""
    If bCorregir Then
        Status.Panels("help").Text = "Corrigiendo stock a " & vsConsulta.Cell(flexcpText, idRow, 1) & " ..."
        Status.Refresh
        FechaDelServidor
        
        Dim mAB As Integer
        For I = LBound(arrSTConteo) To UBound(arrSTConteo)
            With arrSTConteo(I)
                 If .Q <> 0 Then
                    
                    Select Case .Q
                        Case Is < 0: mAB = -1
                        Case Is > 0:  mAB = 1
                    End Select
                    
                    m_Error = "Balance: Tabla StockLocal "
                    MarcoMovimientoStockFisicoEnLocal .LocalTipo, .Local, idArticulo, Abs(.Q), .Estado, mAB
                    m_Error = m_Error & vbTab & "OK" & vbCrLf
                    
                    m_Error = m_Error & "Balance: Tabla MovimientoStockFisico "
                    MarcoMovimientoStockFisico paCodigoDeUsuario, .LocalTipo, .Local, idArticulo, Abs(.Q), .Estado, mAB, 25
                    m_Error = m_Error & vbTab & "OK" & vbCrLf
                    
                    m_Error = m_Error & "Balance: Tabla StockTotal "
                    MarcoMovimientoStockTotal idArticulo, TipoEstadoMercaderia.Fisico, .Estado, Abs(.Q), mAB
                    m_Error = m_Error & vbTab & "OK" & vbCrLf
                    
                    If bHayBDMov Then
                        m_Error = m_Error & "Comercio: Tabla StockLocal "
                        bdMov_MarcoMovimientoStockFisicoEnLocal .LocalTipo, .Local, idArticulo, Abs(.Q), .Estado, mAB
                        m_Error = m_Error & vbTab & "OK" & vbCrLf
                        
                        m_Error = m_Error & "Comercio: Tabla MovimientoStockFisico "
                        bdMov_MarcoMovimientoStockFisico paCodigoDeUsuario, .LocalTipo, .Local, idArticulo, Abs(.Q), .Estado, mAB, 25
                        m_Error = m_Error & vbTab & "OK" & vbCrLf
                        
                        m_Error = m_Error & "Comercio: Tabla StockTotal "
                        bdMov_MarcoMovimientoStockTotal idArticulo, TipoEstadoMercaderia.Fisico, .Estado, Abs(.Q), mAB
                        m_Error = m_Error & vbTab & "OK" & vbCrLf
                    End If
                    
                End If
            End With
        Next
    
        Status.Panels("help").Text = ""
        Status.Refresh
        If bDeAUno And bDelete Then
            fileAddArticulo idArticulo
            vsConsulta.RemoveItem idRow
        End If
    End If
    
End Sub

Private Sub CargoStockLIFO()
On Error GoTo errCLifo
    
    Status.Panels("help").Text = "Consultando datos LIFO ...": Status.Refresh
    Dim aQ As Long, arrArticulos() As typStock
    aQ = 0
    
    cons = "Select Count(Distinct(ComArticulo))" & _
                " From CMCompra "
    
    Set rsAux = cBaseMov.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux(0)) Then aQ = rsAux(0)
    rsAux.Close
    If aQ = 0 Then Exit Sub
    
    ReDim arrArticulos(aQ - 1)
    Dim Idx As Long, aID As Long
    
    pbProgreso.Min = 0: pbProgreso.Value = 0
    pbProgreso.Max = aQ
    
    cons = "Select ComArticulo, Sum(ComCantidad) as Q" & _
                " From CMCOmpra " & _
                " Group by ComArticulo"
                
    Set rsAux = cBaseMov.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Idx = 0
    Do While Not rsAux.EOF
        
        arrArticulos(Idx).Articulo = rsAux!ComArticulo
        arrArticulos(Idx).Stock = rsAux!Q
        
        pbProgreso.Value = pbProgreso.Value + 1
        Idx = Idx + 1
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    pbProgreso.Value = 0
    
    Status.Panels("help").Text = "Procesando Stock LIFO ...": Status.Refresh
    
    'Cargo los datos de la tabla --------------------------------------------------------------------------------
    Dim mRow As Long
        
    For Idx = LBound(arrArticulos) To UBound(arrArticulos)
        mRow = RowEnLista(arrArticulos(Idx).Articulo)
        
        If mRow > 0 Then
            With vsConsulta
                .Cell(flexcpText, mRow, mColLIFO) = Format(arrArticulos(Idx).Stock, "#,##0")
            End With
        End If
        pbProgreso.Value = pbProgreso.Value + 1
    Next
    
'    With vsConsulta
'        .Cell(flexcpBackColor, 2, 2, .Rows - 1, 2 + aQEstados) = Colores.clNaranja
'    End With
    
    ReDim arrArticulos(0)
    pbProgreso.Value = 0
    Exit Sub
    
errCLifo:
    clsGeneral.OcurrioError "Error al cargar datos del LIFO.", Err.Description
End Sub


Private Sub EliminoArticulosArchivo()
    If Trim(prmFileName) = "" Then Exit Sub
    
    Status.Panels("help").Text = "Eliminando artículos corregidos ...": Status.Refresh
    On Error GoTo errEliminar
    Dim bError As Boolean, myText As String
    
    myText = CreateTextFromFile(prmFileName, bError)
    Dim arrFil() As String, uB As Long
    arrFil = Split(myText, vbCrLf)
    
    uB = UBound(arrFil)
    Dim iK As Long
    
    For I = LBound(arrFil) To uB
        If Val(arrFil(I)) <> 0 Then
            For iK = 2 To vsConsulta.Rows - 1
                If vsConsulta.Cell(flexcpData, iK, 1) = Val(arrFil(I)) Then
                    vsConsulta.RemoveItem iK
                    Exit For
                End If
            Next
        End If
    Next
    Status.Panels("help").Text = "": Status.Refresh
    Exit Sub

errEliminar:
    clsGeneral.OcurrioError "Error al eliminar los artículos del archivo.", Err.Description
End Sub

Private Sub EliminoArticulosEnCero()
    
    Status.Panels("help").Text = "Eliminando artículos sin Datos ...": Status.Refresh
    On Error GoTo errEliminar
    Dim bQueda As Boolean
    Dim mIdx As Long, iK As Long
    
    mIdx = 2
    For I = 2 To vsConsulta.Rows - 1
    
        bQueda = False
        For iK = 2 To vsConsulta.Cols - 1
            If vsConsulta.Cell(flexcpValue, mIdx, iK) <> 0 Then
                bQueda = True
                Exit For
            End If
        Next
        
        If Not bQueda Then
            vsConsulta.RemoveItem mIdx
        Else
            mIdx = mIdx + 1
        End If
        
    Next
    Status.Panels("help").Text = "": Status.Refresh
    Exit Sub

errEliminar:
    clsGeneral.OcurrioError "Error al eliminar los artículos sin datos.", Err.Description
End Sub

Private Sub CargoTodosLosArticulos()

    Status.Panels("help").Text = "Cargando Artículos ...": Status.Refresh
    
    Dim aQ As Long
    aQ = 0
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    cons = "Select Count(ArtID) from Articulo"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux(0)) Then aQ = rsAux(0)
    rsAux.Close
    
    If aQ = 0 Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        Screen.MousePointer = 0:  Exit Sub
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------
    
    pbProgreso.Min = 0: pbProgreso.Value = 0
    pbProgreso.Max = aQ
    
    'Cargo Artículos del STOCK CONTEO  ---------------------------------------------------------------------------------------------------------
    cons = "Select ArtID, ArtCodigo, ArtNombre from Articulo"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    With vsConsulta
        Do While Not rsAux.EOF
            pbProgreso.Value = pbProgreso.Value + 1


            .AddItem ""
            aRow = .Rows - 1

            .Cell(flexcpChecked, aRow, 0) = flexUnchecked
                
            .Cell(flexcpText, aRow, 1) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
            aValor = rsAux!ArtID: .Cell(flexcpData, aRow, 1) = aValor
            
            rsAux.MoveNext
        Loop
        rsAux.Close
    
    End With
    
    pbProgreso.Value = 0
    
End Sub


Private Sub AccionMenuHelp()
    On Error GoTo errHelp
    Screen.MousePointer = 11
    
    Dim aFile As String
    cons = "Select * from Aplicacion Where AplNombre = 'Diferencias Balance'"
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

