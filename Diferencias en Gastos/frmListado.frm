VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Begin VB.Form frmListado 
   Caption         =   "Diferencias en Gastos"
   ClientHeight    =   7065
   ClientLeft      =   1500
   ClientTop       =   2715
   ClientWidth     =   11190
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
   ScaleHeight     =   7065
   ScaleWidth      =   11190
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   5115
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   10695
      _Version        =   196608
      _ExtentX        =   18865
      _ExtentY        =   9022
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
      PreviewMode     =   1
      Zoom            =   60
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
      ForeColor       =   &H00000080&
      Height          =   660
      Left            =   120
      TabIndex        =   19
      Top             =   60
      Width           =   11055
      Begin VB.CheckBox chIncluir 
         Caption         =   "&Incluir Gastos sin Pagos"
         Height          =   195
         Left            =   4320
         TabIndex        =   22
         Top             =   300
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.TextBox tHasta 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         MaxLength       =   12
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox tDesde 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   10995
      TabIndex        =   17
      Top             =   6240
      Width           =   11055
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmListado.frx":0EF0
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmListado.frx":0FDA
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmListado.frx":10C4
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "frmListado.frx":12FE
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Picture         =   "frmListado.frx":1400
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5640
         Picture         =   "frmListado.frx":17C6
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":18C8
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmListado.frx":1BCA
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmListado.frx":1F0C
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmListado.frx":220E
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin MSComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6240
         TabIndex        =   21
         Top             =   120
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   6810
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
            Object.Width           =   11536
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   3255
      Left            =   1380
      TabIndex        =   4
      Top             =   1080
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
   Begin MSComctlLib.ImageList img1 
      Left            =   6960
      Top             =   3720
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
            Picture         =   "frmListado.frx":2448
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":2762
            Key             =   "help"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":2874
            Key             =   "consultar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":29CE
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":2B28
            Key             =   "print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":2C82
            Key             =   "limpiar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":2D94
            Key             =   "vista2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":2EEE
            Key             =   "vista1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":3048
            Key             =   "move2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":31A2
            Key             =   "move3"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":32FC
            Key             =   "move4"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":3456
            Key             =   "move1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":35B0
            Key             =   "configprint"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsAux As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean


Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConfigurar_Click()
    AccionConfigurar
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
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

Private Sub chIncluir_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
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

Private Sub Form_Load()
    
    On Error Resume Next
    FechaDelServidor
    ObtengoSeteoForm Me, 1000, 500, 3840, 4230
    pbProgreso.Value = 0
    
    With vsListado
        .Orientation = orPortrait
        .PaperSize = 1
        .MarginRight = 350: .MarginLeft = 350
        .MarginBottom = 750: .MarginTop = 750
        .Zoom = 100
    End With
    
    InicializoGrilla
    vsConsulta.ZOrder 0
    
    picBotones.BorderStyle = vbBSNone
    
    tDesde.Text = Format(PrimerDia(Now), "dd/mm/yyyy")
    tHasta.Text = Format(UltimoDia(Now), "dd/mm/yyyy")
    
    bCargarImpresion = True
    
    With img1
        bConsultar.Picture = .ListImages("consultar").ExtractIcon
        bImprimir.Picture = .ListImages("print").ExtractIcon
        bNoFiltros.Picture = .ListImages("limpiar").ExtractIcon
        bCancelar.Picture = .ListImages("salir").ExtractIcon
        chVista.Picture = .ListImages("vista1").ExtractIcon
        chVista.DownPicture = .ListImages("vista2").ExtractIcon
        bPrimero.Picture = .ListImages("move1").ExtractIcon
        bAnterior.Picture = .ListImages("move2").ExtractIcon
        bSiguiente.Picture = .ListImages("move3").ExtractIcon
        bUltima.Picture = .ListImages("move4").ExtractIcon
        bConfigurar.Picture = .ListImages("configprint").ExtractIcon
    End With
    
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = fFiltros.Width
    
    picBotones.Width = vsListado.Width
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 50
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = vsListado.Height
    vsConsulta.Left = vsListado.Left
    
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

Private Sub tDesde_GotFocus()
    With tDesde: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    If bCargarImpresion Then
        With vsListado
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        aTexto = "Diferencias en Gastos - del " & Trim(tDesde.Text) & " al " & Trim(tHasta.Text)
        EncabezadoListado vsListado, aTexto, True
        vsListado.FileName = "Diferencias en Gastos"
        
        vsConsulta.ExtendLastCol = False
        vsListado.RenderControl = vsConsulta.hwnd
        vsConsulta.ExtendLastCol = True
        
        vsListado.EndDoc
        'bCargarImpresion = False
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
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub

Private Sub tDesde_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If IsDate(tDesde.Text) Then tHasta.Text = Format(UltimoDia(CDate(tDesde.Text)), "dd/mm/yyyy")
        Foco tHasta
    End If
    
End Sub

Private Sub tDesde_LostFocus()
    If IsDate(tDesde.Text) Then tDesde.Text = Format(tDesde.Text, "dd/mm/yyyy")
End Sub


Private Sub tHasta_GotFocus()
    With tHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chIncluir.SetFocus
End Sub

Private Sub tHasta_LostFocus()
    If IsDate(tHasta.Text) Then tHasta.Text = Format(tHasta.Text, "dd/mm/yyyy")
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionConsultar()
 
Dim rs1 As rdoResultset
Dim aTotal As Currency
Dim aMonCodigo As Long, aMonSigno As String

    On Error GoTo ErrCDML
    bCargarImpresion = True
    
    If Not ValidoDatos Then Exit Sub
    
    Screen.MousePointer = 11
    chVista.Value = vbUnchecked
    
    vsConsulta.Rows = 1: vsConsulta.Refresh
    
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    pbProgreso.Value = 0
    cons = "Select Count(*) from  Compra " _
                                    & " left Outer Join MovimientoDisponibilidad On MDiIdCompra = ComCodigo" _
           & " Where ComFecha Between '" & Format(tDesde.Text, sqlFormatoF) & "' And '" & Format(tHasta.Text, sqlFormatoF) & "'" _
           & " And ComTipoDocumento Not In (" & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraNotaCredito & ")"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If rsAux(0) = 0 Then
            MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
            rsAux.Close: Screen.MousePointer = 0: Exit Sub
        End If
        pbProgreso.Max = rsAux(0)
    End If
    rsAux.Close
    '-----------------------------------------------------------------------------------------------------------------
    
    cons = "Select * from  Compra " _
                                    & " left Outer Join MovimientoDisponibilidad On MDiIdCompra = ComCodigo" _
           & " Where ComFecha Between '" & Format(tDesde.Text, sqlFormatoF) & "' And '" & Format(tHasta.Text, sqlFormatoF) & "'" _
           & " And ComTipoDocumento Not In (" & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraNotaCredito & ")"
    cons = cons & " Order by ComCodigo"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If rsAux.EOF Then Screen.MousePointer = 0: rsAux.Close: Exit Sub
        
    'Preparo Query para sacar el Rubro del Subrubro----------------------
    Dim qyPago As rdoQuery
    cons = "Select Compra = Sum(MDRImporteCompra), Pesos = Sum(MDRImportePesos) " _
           & " From MovimientoDisponibilidadRenglon Where MDRIdMovimiento = ?"
    Set qyPago = cBase.CreateQuery("", cons)
    
    aMonCodigo = 0
    With vsConsulta
    .Redraw = False
    Do While Not rsAux.EOF
        pbProgreso.Value = pbProgreso.Value + 1
        
        '       id_Gasto|Fecha|>Moneda|>Total Gasto|>T/C|>Total (G) $|>Total (G) U$S|>Pagos $|>Pagos U$S
        
        If aMonCodigo <> rsAux!ComMoneda Then
            aMonCodigo = rsAux!ComMoneda
            cons = "Select * from Moneda Where MonCodigo = " & aMonCodigo
            Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rs1.EOF Then aMonSigno = Trim(rs1!MonSigno)
            rs1.Close
        End If
            
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!ComCodigo, "#,000")
        .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!ComFecha, "dd/mm/yy")
        .Cell(flexcpText, .Rows - 1, 2) = aMonSigno
         If rsAux!ComMoneda <> paMonedaPesos Then .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!ComTC, "#.000")
        
        aTotal = rsAux!ComImporte
        If Not IsNull(rsAux!ComIVa) Then aTotal = aTotal + rsAux!ComIVa
        If Not IsNull(rsAux!ComCofis) Then aTotal = aTotal + rsAux!ComCofis
        aTotal = Abs(aTotal)
        .Cell(flexcpText, .Rows - 1, 3) = Format(aTotal, FormatoMonedaP)
        
        If rsAux!ComMoneda = paMonedaPesos Then
            .Cell(flexcpText, .Rows - 1, 5) = Format(aTotal, FormatoMonedaP)
            '.Cell(flexcpText, .Rows - 1, 6) = Format(aTotal / .Cell(flexcpValue, .Rows - 1, 4), FormatoMonedaP)
        Else
            .Cell(flexcpText, .Rows - 1, 6) = Format(aTotal, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 5) = Format(aTotal * .Cell(flexcpValue, .Rows - 1, 4), FormatoMonedaP)
        End If
        
        If Not IsNull(rsAux!MDiID) Then
            qyPago.rdoParameters(0) = rsAux!MDiID
            Set rs1 = qyPago.OpenResultset(rdOpenDynamic, rdConcurValues)
            If Not rs1.EOF Then
                .Cell(flexcpText, .Rows - 1, 7) = Format(rs1!Pesos, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 8) = Format(rs1!Compra, FormatoMonedaP)
            End If
            rs1.Close
        End If
        
        If IsNull(rsAux!MDiID) And chIncluir.Value = vbUnchecked Then
            .RemoveItem .Rows - 1
        Else
            'Si coinciden los importe remuevo el renglon
            If .Cell(flexcpValue, .Rows - 1, 5) = .Cell(flexcpValue, .Rows - 1, 7) Then
                .RemoveItem .Rows - 1
            Else
                .Cell(flexcpText, .Rows - 1, 9) = Format(.Cell(flexcpValue, .Rows - 1, 5) - .Cell(flexcpValue, .Rows - 1, 7), FormatoMonedaP)
            End If
        End If
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    qyPago.Close
    
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 5, , Colores.Rojo, Colores.Blanco, True, "Diferencias"
    .Subtotal flexSTSum, -1, 7
    .Subtotal flexSTSum, -1, 9
    
    End With
    
    pbProgreso.Value = 0
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
    If vsConsulta.Rows = 1 Then MsgBox "No se encontraron diferencias en el registro de los gastos. " & Chr(vbKeyReturn) & "(*) Los cáclulos se realizan contra los pagos ingresados.", vbInformation, "No Hay Diferencias"
    Exit Sub
    
ErrCDML:
    clsGeneral.OcurrioError "Error al cargar los datos.", Err.Description
    vsConsulta.Redraw = True
    pbProgreso.Value = 0
    Screen.MousePointer = 0
End Sub

Private Sub AccionLimpiar()
    tDesde.Text = "": tHasta.Text = ""
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

Private Function ValidoDatos() As Boolean
    On Error Resume Next
    ValidoDatos = False
    
    If Not IsDate(tDesde.Text) Then
        MsgBox "La fecha ingresada para consultar no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    If Not IsDate(tHasta.Text) Then
        MsgBox "La fecha ingresada para consultar no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tHasta: Exit Function
    End If
    
    If CDate(tDesde.Text) > CDate(tHasta.Text) Then
        MsgBox "El rango de fechas ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tHasta: Exit Function
    End If

    ValidoDatos = True
    
End Function

Private Sub InicializoGrilla()

    On Error Resume Next
    
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = "id_Gasto|Fecha|<Moneda|>Total Gasto|>T/C|>Total (G) $|>Total (G) U$S|>Pagos $|>Pagos (Mon. Gasto)|>Diferencia|"
        .ColWidth(0) = 750: .ColWidth(1) = 750
        .ColWidth(3) = 1400: .ColWidth(4) = 750: .ColWidth(5) = 1400
        .ColWidth(6) = 1400: .ColWidth(7) = 1400: .ColWidth(8) = 1600
        .ColWidth(9) = 1100
        .WordWrap = False
        .MergeCells = flexMergeSpill
    End With
      
End Sub


