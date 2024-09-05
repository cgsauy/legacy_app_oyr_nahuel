VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCompra 
   Caption         =   "Compra de Mercadería"
   ClientHeight    =   8340
   ClientLeft      =   1290
   ClientTop       =   1920
   ClientWidth     =   12780
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCompra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   12780
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   9555
      TabIndex        =   12
      Top             =   5880
      Width           =   9615
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmCompra.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmCompra.frx":052C
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Visible         =   0   'False
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmCompra.frx":0616
         Height          =   310
         Left            =   4440
         Picture         =   "frmCompra.frx":0718
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bExportar 
         Height          =   310
         Left            =   5280
         Picture         =   "frmCompra.frx":0C4A
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Exportar"
         Top             =   120
         Visible         =   0   'False
         Width           =   310
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   4320
      TabIndex        =   9
      Top             =   1560
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7646
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
      HighLight       =   2
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   8085
      Width           =   12780
      _ExtentX        =   22543
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "sucursal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   11774
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   960
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   11115
      Begin VB.TextBox tRubro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8100
         TabIndex        =   7
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox tSubRubro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4860
         TabIndex        =   5
         Top             =   240
         Width           =   2595
      End
      Begin MSComCtl2.DTPicker tDesde 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   23855105
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker tHasta 
         Height          =   315
         Left            =   2580
         TabIndex        =   3
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   23855105
         CurrentDate     =   37543
      End
      Begin VB.Label Label1 
         Caption         =   "&Rubro:"
         Height          =   255
         Left            =   7560
         TabIndex        =   6
         Top             =   315
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "&Sub Rubro:"
         Height          =   255
         Left            =   4020
         TabIndex        =   4
         Top             =   315
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   315
         Width           =   615
      End
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4455
      Left            =   60
      TabIndex        =   25
      Top             =   1500
      Width           =   11175
      _Version        =   196608
      _ExtentX        =   19711
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
      Zoom            =   100
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   10080
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompra.frx":0F54
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompra.frx":126E
            Key             =   "help"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompra.frx":1380
            Key             =   "consultar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompra.frx":14DA
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompra.frx":1634
            Key             =   "print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompra.frx":178E
            Key             =   "limpiar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompra.frx":18A0
            Key             =   "vista2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompra.frx":19FA
            Key             =   "vista1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompra.frx":1B54
            Key             =   "move2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompra.frx":1CAE
            Key             =   "move3"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompra.frx":1E08
            Key             =   "move4"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompra.frx":1F62
            Key             =   "move1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompra.frx":20BC
            Key             =   "configprint"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsAux As rdoResultset
Private aTexto As String

Dim mSQL As String
Dim aTiposDocs As String

Private Sub AccionLimpiar()
    tDesde.Value = Format(PrimerDia(Now), "dd/mm/yyyy")
    tHasta.Value = Format(UltimoDia(Now), "dd/mm/yyyy")
    tSubRubro.Text = ""
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub
Private Sub bImprimir_Click()
    AccionImprimir Imprimir:=True
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
    Me.Caption = "Listado de Comprobantes"
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    InicializoGrillas
    AccionLimpiar
    
    'Cargo los datos-------------------------------------------------------------------------------------------------
    tDesde.Value = Format(PrimerDia(Now), "dd/mm/yyyy")
    tHasta.Value = Format(UltimoDia(Now), "dd/mm/yyyy")

    
    PropiedadesImpresion
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = "Compra|Fecha|Proveedor|<Factura|>Importe|>Cofis|>I.V.A.|>Total|>Total M/E|"
            
        .WordWrap = False
        .ColWidth(0) = 750: .ColWidth(1) = 900: .ColWidth(2) = 2700: .ColWidth(3) = 900: _
        .ColWidth(4) = 1500: .ColWidth(5) = 900: .ColWidth(6) = 1200: .ColWidth(7) = 1400
        .ColWidth(8) = 1500
    End With
      
    With img1
        bConsultar.Picture = .ListImages("consultar").ExtractIcon
        bPrimero.Picture = .ListImages("move1").ExtractIcon
        bAnterior.Picture = .ListImages("move2").ExtractIcon
        bSiguiente.Picture = .ListImages("move3").ExtractIcon
        bUltima.Picture = .ListImages("move4").ExtractIcon
        
        bImprimir.Picture = .ListImages("print").ExtractIcon
        bConfigurar.Picture = .ListImages("configprint").ExtractIcon
        
        bNoFiltros.Picture = .ListImages("limpiar").ExtractIcon
        bCancelar.Picture = .ListImages("salir").ExtractIcon
        chVista.Picture = .ListImages("vista1").ExtractIcon
        chVista.DownPicture = .ListImages("vista2").ExtractIcon
        
    End With
      
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

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11
    picBotones.BorderStyle = vbFlat
    fFiltros.Width = Me.Width - (fFiltros.Left * 2.5)
    
    picBotones.Top = Me.ScaleHeight - (Status.Height + picBotones.Height + 90)
    
    With vsConsulta
        .Left = fFiltros.Left
        .Top = fFiltros.Top + fFiltros.Height + 50
        .Height = Me.ScaleHeight - (.Top + Status.Height + picBotones.Height + 120)
        .Width = fFiltros.Width
    
        vsListado.Move .Left, .Top, .Width, .Height
    End With
    
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub AccionConsultar()

Dim mTotal As Currency, mIva As Currency, mCofis As Currency
Dim aCompra As Long
Dim mImporteME As Currency

    On Error GoTo errConsultar
    If Not ValidoFiltros Then Exit Sub
            
    Screen.MousePointer = 11
    vsConsulta.Rows = vsConsulta.FixedRows

    aTiposDocs = TipoDocumento.CompraContado & ", " & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraRecibo & ", " _
                    & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraNotaDevolucion & ", " _
                    & TipoDocumento.CompraEntradaCaja & ", " & TipoDocumento.CompraSalidaCaja

    
    mSQL = " Select * from Compra " & _
                    " Left Outer Join ProveedorCliente On ComProveedor = PClCodigo" & _
                " Where ComFecha Between '" & Format(tDesde.Value, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Value, "mm/dd/yyyy 23:59:59") & "'"
    
    mSQL = mSQL & " And ComTipoDocumento In (" & aTiposDocs & ")"
                
    If Val(tSubRubro.Tag) <> 0 Or Val(tRubro.Tag) <> 0 Then
        mSQL = mSQL & " And ComCodigo IN (" & _
                                    " Select GSrIDCompra from GastoSubrubro, SubRubro " & _
                                    " Where GSrIDSubrubro  = SRuID " & _
                                     IIf(Val(tSubRubro.Tag) <> 0, " And SRuID = " & Val(tSubRubro.Tag), "") & _
                                     IIf(Val(tRubro.Tag) <> 0, " And SRuRubro = " & Val(tRubro.Tag), "") & _
                                     ")"
                                        
    End If

    mSQL = mSQL & " Order by ComFecha"
    '------------------------------------------------------------------------------------------------------------------------------------------------------------
    Set rsAux = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        vsConsulta.Rows = 1
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        rsAux.Close: Screen.MousePointer = 0: Exit Sub
    End If
    
    aCompra = 0
    mTotal = 0: mIva = 0: mCofis = 0: mImporteME = 0
    Do While Not rsAux.EOF
        '"Compra|Fecha|Proveedor|<Factura|>Importe|>Cofis|>I.V.A.|>Total|>Total M/E|"
            
        With vsConsulta
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!ComCodigo, "#,###,##0")
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!ComFecha, "dd/mm/yy")
            If Not IsNull(rsAux!PClNombre) Then .Cell(flexcpText, .Rows - 1, 2) = Trim(rsAux!PClNombre)
                
            If Not IsNull(rsAux!ComSerie) Then aTexto = Trim(rsAux!ComSerie) & " " Else aTexto = ""
            If Not IsNull(rsAux!ComNumero) Then aTexto = aTexto & Trim(rsAux!ComNumero)
            .Cell(flexcpText, .Rows - 1, 3) = Trim(aTexto)
                
            If rsAux!ComMoneda <> paMonedaPesos Then
                mImporteME = rsAux!ComImporte
                If Not IsNull(rsAux!ComIva) Then mImporteME = mImporteME + rsAux!ComIva
                If Not IsNull(rsAux!ComCofis) Then mImporteME = mImporteME + rsAux!ComCofis
                
                mTotal = Format(rsAux!ComImporte * rsAux!ComTC, FormatoMonedaP)
                If Not IsNull(rsAux!ComIva) Then mIva = rsAux!ComIva * rsAux!ComTC
                If Not IsNull(rsAux!ComCofis) Then mCofis = rsAux!ComCofis * rsAux!ComTC
                
            Else
                mTotal = rsAux!ComImporte
                If Not IsNull(rsAux!ComIva) Then mIva = rsAux!ComIva
                If Not IsNull(rsAux!ComCofis) Then mCofis = rsAux!ComCofis
            End If
            
            If rsAux!ComTipoDocumento = TipoDocumento.CompraNotaCredito Or rsAux!ComTipoDocumento = TipoDocumento.CompraNotaDevolucion Or _
                rsAux!ComTipoDocumento = TipoDocumento.CompraEntradaCaja Then
                .Cell(flexcpText, .Rows - 1, 4) = Format(.Cell(flexcpValue, .Rows - 1, 4) * -1, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 5) = Format(.Cell(flexcpValue, .Rows - 1, 5) * -1, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpValue, .Rows - 1, 6) * -1, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 7) = Format(.Cell(flexcpValue, .Rows - 1, 7) * -1, FormatoMonedaP)
                .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.RojoClaro
            End If

            .Cell(flexcpText, .Rows - 1, 4) = Format(mTotal, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 5) = Format(mCofis, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = Format(mIva, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 7) = Format(mTotal + mIva + mCofis, FormatoMonedaP)
            
            If mImporteME <> 0 Then .Cell(flexcpText, .Rows - 1, 8) = Format(mImporteME, FormatoMonedaP)
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    If vsConsulta.Rows > vsConsulta.FixedRows Then
    With vsConsulta
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 4, , &HC00000, vbWhite, True, " "
        .Subtotal flexSTSum, -1, 5: .Subtotal flexSTSum, -1, 6: .Subtotal flexSTSum, -1, 7
        .Subtotal flexSTCount, -1, 0, "#,##0"
    End With
    End If
    
    
    Screen.MousePointer = 0
    
    Exit Sub

errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Function ValidoFiltros() As Boolean

    ValidoFiltros = False
   
    If Not IsDate(tDesde.Value) Then
        MsgBox "Debe ingresar una fecha válida para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    If Not IsDate(tHasta.Value) Then
        MsgBox "Debe ingresar una fecha válida para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco tHasta: Exit Function
    End If
    
    If CDate(tDesde.Value) > CDate(tHasta.Value) Then
        MsgBox "El rango de fechas para realizar la consulta no es correcto..", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    ValidoFiltros = True
    
End Function

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
Dim J As Integer

    If vsConsulta.Rows = 1 Then
        MsgBox "No hay datos en la lista para realizar la impresión.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    On Error GoTo errPrint
    Screen.MousePointer = 11
    
    With vsListado
        .Orientation = orPortrait '  orLandscape
                
        .Preview = True
        .StartDoc
                
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN"
            Screen.MousePointer = vbDefault: Exit Sub
        End If
        
        Dim mTitulo As String
        mTitulo = "Comprobantes desde " & Trim(tDesde.Value) & " al " & Trim(tHasta.Value)
        If Val(tRubro.Tag) <> 0 Then mTitulo = mTitulo & " - " & tRubro.Text
        If Val(tSubRubro.Tag) <> 0 Then mTitulo = mTitulo & " - " & tSubRubro.Text
        
    
        EncabezadoListado vsListado, mTitulo, False
        .FileName = "Comprobantes"
        
        vsConsulta.ExtendLastCol = False: .RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        
        .EndDoc
    End With
    
    If Imprimir Then
        'frmSetup.pControl = vsListado
        'frmSetup.Show vbModal, Me
        'Me.Refresh
'        If frmSetup.pOK Then vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
        vsListado.PrintDoc True, 1, vsListado.PageCount
        'If Not vsListado.PrintDialog(pdPrinterSetup) Then Screen.MousePointer = 0: Exit Sub
    End If

    Screen.MousePointer = 0
    Exit Sub

errPrint:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión. ", Err.Description
End Sub


Private Sub Label4_Click()
    Foco tHasta
End Sub

Private Sub tDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsDate(tDesde.Value) Then tHasta.Value = Format(UltimoDia(tDesde.Value), "dd/mm/yyyy")
        Foco tHasta
    End If
End Sub

Private Sub tHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco tSubRubro
End Sub

Private Sub tRubro_Change()
    tRubro.Tag = 0
End Sub

Private Sub tRubro_GotFocus()
    tRubro.SelStart = 0: tRubro.SelLength = Len(tRubro.Text)
End Sub

Private Sub tRubro_KeyPress(KeyAscii As Integer)
On Error GoTo errBS
    
    If KeyAscii = vbKeyReturn Then
        
        If Trim(tRubro.Text) <> "" And Val(tRubro.Tag) = 0 Then
            ing_BuscoSubrubro tRubro, 0
            Exit Sub
        End If
                
        Foco bConsultar
    End If
    
    Exit Sub
errBS:
    clsGeneral.OcurrioError "Error al buscar el subrubro.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub PropiedadesImpresion()

    With vsListado
        .PhysicalPage = True
        .PaperSize = vbPRPSLetter
        .Orientation = orLandscape
        .PreviewMode = pmPrinter
        .PreviewPage = 1
        .Zoom = 100
        .MarginLeft = 450: .MarginRight = 350
        .MarginBottom = 750: .MarginTop = 750
    End With
    
End Sub

Private Sub tSubRubro_Change()
    tSubRubro.Tag = 0
End Sub

Private Sub tSubRubro_GotFocus()
    tSubRubro.SelStart = 0: tSubRubro.SelLength = Len(tSubRubro.Text)
End Sub

Private Sub tSubRubro_KeyPress(KeyAscii As Integer)
On Error GoTo errBS
    
    If KeyAscii = vbKeyReturn Then
        
        If Trim(tSubRubro.Text) <> "" And Val(tSubRubro.Tag) = 0 Then
            ing_BuscoSubrubro tSubRubro, 1
            Exit Sub
        End If
                
        Foco tRubro
    End If
    
    Exit Sub
errBS:
    clsGeneral.OcurrioError "Error al buscar el subrubro.", Err.Description
    Screen.MousePointer = 0
End Sub

Public Function ing_BuscoSubrubro(mControlSR As TextBox, mCual As Byte) As Boolean
On Error GoTo errBS
'mCual = 0-Rubro; 1-SubRubo


    ing_BuscoSubrubro = False
    Dim aQ As Integer, aID As Long, aTexto As String
    aQ = 0: aID = 0
    
    mControlSR.Text = Replace(RTrim(mControlSR.Text), " ", "%")
    
    Select Case mCual
    Case 0
        cons = "Select RubID, RubNombre as 'Rubro', RubCodigo as 'Cód. Rubro'" _
                    & " from  Rubro " _
                    & " Where RubNombre like '" & Trim(mControlSR.Text) & "%'" _
                    & " Order by RubNombre"

    Case 1
        cons = "Select SRuID, SRuNombre as 'SubRubro', SRuCodigo as 'Cód. SR', RubNombre as 'Rubro', RubCodigo as 'Cód. Rubro'" _
                & " from SubRubro, Rubro " _
                & " Where SRuNombre like '" & Trim(mControlSR.Text) & "%'" _
                & " And SRuRubro = RubID " _
                & " Order by SRuNombre"
    End Select
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        aQ = 1: aID = rsAux(0): aTexto = Trim(rsAux(1))
        rsAux.MoveNext
        If Not rsAux.EOF Then
            aQ = 2: aID = 0
        End If
    End If
    rsAux.Close
    
    Select Case aQ
        Case 0: MsgBox "No existen datos para el texto ingresado.", vbExclamation, "No hay datos"
        
        Case 1:
                mControlSR.Text = aTexto: mControlSR.Tag = aID
        
        Case 2:
                Dim aLista As New clsListadeAyuda
                aID = aLista.ActivarAyuda(cBase, cons, 5500, 1, "Ayuda de datos")
                
                If aID <> 0 Then
                    aTexto = Trim(aLista.RetornoDatoSeleccionado(1))
                    aID = aLista.RetornoDatoSeleccionado(0)
                End If
                
                Set aLista = Nothing
    End Select
    
    If aID <> 0 Then
        Select Case mCual
            Case 0: cons = "Select RubID, RubNombre from Rubro Where RubID = " & aID
            Case 1: cons = "Select SRuID, SRuNombre from Subrubro Where SRuID = " & aID
        End Select
        
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsAux.EOF Then
            mControlSR.Text = Trim(rsAux(1))
            mControlSR.Tag = rsAux(rsAux(0))
            ing_BuscoSubrubro = True
        End If
        rsAux.Close
        
    End If
    
    Screen.MousePointer = 0
    Exit Function

errBS:
    clsGeneral.OcurrioError "Error al buscar el subrubro.", Err.Description
    Screen.MousePointer = 0
End Function

