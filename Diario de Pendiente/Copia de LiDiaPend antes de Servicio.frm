VERSION 5.00
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACOMBO.OCX"
Begin VB.Form frmListado 
   Caption         =   "Diario de Pendiente Contado"
   ClientHeight    =   3870
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LiDiaPend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrilla 
      Height          =   1095
      Left            =   2160
      TabIndex        =   18
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1931
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
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   1935
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   1931
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
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   9855
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   8760
         TabIndex        =   24
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
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
      Begin VB.TextBox tHasta 
         Height          =   285
         Left            =   6720
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox tDesde 
         Height          =   285
         Left            =   4800
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin AACombo99.AACombo cSucursal 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   240
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
      Begin VB.Label Label4 
         Caption         =   "&Moneda:"
         Height          =   255
         Left            =   8040
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   6120
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   4200
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Sucursal:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   6075
      TabIndex        =   15
      Top             =   3000
      Width           =   6135
      Begin VB.CheckBox chVista 
         DownPicture     =   "LiDiaPend.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "LiDiaPend.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "LiDiaPend.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "LiDiaPend.frx":0F38
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "LiDiaPend.frx":1022
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "LiDiaPend.frx":110C
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "LiDiaPend.frx":1346
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Picture         =   "LiDiaPend.frx":1448
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Limpiar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5400
         Picture         =   "LiDiaPend.frx":180E
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "LiDiaPend.frx":1910
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "LiDiaPend.frx":1C12
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "LiDiaPend.frx":1F54
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
         Picture         =   "LiDiaPend.frx":2256
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   3615
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12091
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cambios
    ' 13/4 le agregue a la consulta que las ventas teléfonicas tengan envreclamocobro.

Option Explicit

Private strEncabezado As String, strFormato As String
Private aTexto As String
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
    AccionImprimir
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

Private Sub cMoneda_GotFocus()
    With cMoneda: .SelStart = 0: .SelLength = Len(.Text): End With
    Ayuda "Seleccione la moneda con que se facturó."
End Sub
Private Sub cMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco bConsultar
End Sub
Private Sub cMoneda_LostFocus()
    With cMoneda: .SelStart = 0: End With
    Ayuda ""
End Sub

Private Sub cSucursal_GotFocus()
    With cSucursal
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda " Seleccione una sucursal."
End Sub
Private Sub cSucursal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDesde
End Sub
Private Sub cSucursal_LostFocus()
    cSucursal.SelStart = 0: Ayuda ""
End Sub

Private Sub chVista_Click()
    If chVista.Value = 0 Then
        'Selecciono Grilla
        vsGrilla.ZOrder 0
    Else
        'Selecciono Listado
        vsGrilla.ExtendLastCol = False
        With vsListado
            .StartDoc
            EncabezadoListado vsListado, "Diario de Pendiente Contados desde " & tDesde.Text & " hasta " & tHasta.Text, True
            .RenderControl = vsGrilla.hWnd
            .EndDoc
        End With
        vsListado.ZOrder 0
        vsGrilla.ExtendLastCol = True
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
            Case vbKeyI: AccionImprimir
            Case vbKeyC: AccionConfigurar
            
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad

    ObtengoSeteoForm Me, 1000, 500, 3840, 4230
    picBotones.BorderStyle = vbBSNone
    
    PropiedadesImpresion
    LimpioGrilla
    
    'Cargo Sucursales.-------------------------------
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal " _
        & " Order By SucAbreviacion "
    CargoCombo Cons, cSucursal
    cSucursal.AddItem "Todos"
    cSucursal.ItemData(cSucursal.NewIndex) = 0
    'Por defecto pongo todos.
    For I = 0 To cSucursal.ListCount - 1
        If cSucursal.ItemData(I) = 0 Then cSucursal.ListIndex = I: Exit For
    Next
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda
    BuscoCodigoEnCombo cMoneda, CLng(paMonedaPesos)
    '--------------------------------------------------------------
    tDesde.Text = Format(Date, FormatoFP)
    tHasta.Text = Format(Date, FormatoFP)
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error inesperado al cargar el formulario.", Err.Description
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = fFiltros.Width
    
    vsGrilla.Width = vsListado.Width
    vsGrilla.Height = vsListado.Height
    vsGrilla.Left = vsListado.Left
    
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
    Foco cSucursal
End Sub

Private Sub AccionImprimir()
    On Error GoTo ErrImprimir
    
    vsGrilla.ExtendLastCol = False
    With vsListado
        .StartDoc
        .filename = "DiarioPendiente"
        EncabezadoListado vsListado, "Diario de Pendiente Contados desde " & tDesde.Text & " hasta " & tHasta.Text, True
        .RenderControl = vsGrilla.hWnd
        .EndDoc
        
    End With
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    Me.Refresh
    If Not frmSetup.pOK Then Exit Sub
    vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
    
    vsGrilla.ExtendLastCol = True
    Exit Sub
    
ErrImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub

Private Sub PropiedadesImpresion()
  
  With vsListado
        .PaperSize = vbPRPSLetter
        .PhysicalPage = True
        .Orientation = orPortrait
        .PreviewMode = pmPrinter
        .PreviewPage = 1
        .FontName = "Tahoma": .FontSize = 10: .FontBold = False: .FontItalic = False
        .TextAlign = 0: .PageBorder = 3
        .Columns = 1
        .TableBorder = tbBoxRows
        .Zoom = 60
    End With

End Sub

Private Sub Label2_Click()
    Foco tDesde
End Sub

Private Sub Label3_Click()
    Foco tHasta
End Sub

Private Sub tDesde_GotFocus()
    With tDesde
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese desde que fecha desea consultar."
End Sub
Private Sub tDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(tDesde.Text) Then
            Foco tHasta
        Else
            MsgBox "La fecha desde no es correcta.", vbExclamation, "ATENCIÓN"
        End If
    End If
End Sub
Private Sub tDesde_LostFocus()
    tDesde.SelStart = 0: Ayuda ""
    If IsDate(tDesde.Text) Then tDesde.Text = Format(tDesde.Text, FormatoFP)
End Sub
Private Sub tHasta_GotFocus()
    With tHasta
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese hasta que fecha desea consultar."
End Sub
Private Sub tHasta_KeyPress(KeyAscii As Integer)
    If vbKeyReturn = KeyAscii Then Foco cMoneda
End Sub
Private Sub tHasta_LostFocus()
    Ayuda ""
    If IsDate(tHasta.Text) Then tHasta.Text = Format(tHasta.Text, FormatoFP)
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionConsultar()

    On Error GoTo ErrCDML
    
    If Not ValidoDatos Then Exit Sub
    
    Screen.MousePointer = 11
    vsGrilla.ZOrder 0
    LimpioGrilla
    chVista.Value = 0
    
    'Saco las ventas telefónicas.
    Cons = "Select Documento.*, Renglon.*, Articulo.*, EnvCodigo, SucAbreviacion, CamNombre " _
        & " From VentaTelefonica, Documento, Renglon, Envio, Articulo, Sucursal, Camion" _
        & " Where DocTipo = " & TipoDocumento.Contado _
        & " And DocFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59:59") & "'" _
        & " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
    
    If cSucursal.ItemData(cSucursal.ListIndex) > 0 Then Cons = Cons & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
            
    Cons = Cons _
        & " And DocAnulado = 0" _
        & " And EnvEstado IN ( " & EstadoEnvio.Impreso & ", " & EstadoEnvio.Entregado & ") And EnvReclamoCobro <> 0" _
        & " And DocSucursal = SucCodigo And EnvLiquidacion = Null And DocCodigo = VTeDocumento And DocCodigo = EnvDocumento And EnvCamion = CamCodigo" _
        & " And DocCodigo = RenDocumento And RenArticulo = ArtId " _

    'Le uno todas las facturas que se hicieron para cobrar el Flete.
    Cons = Cons & " UNION ALL " _
        & "Select Documento.*, Renglon.*, Articulo.*, EnvCodigo, SucAbreviacion, CamNombre From Documento, Renglon, Envio, Articulo, Sucursal, Camion " _
        & " Where DocTipo = " & TipoDocumento.Contado _
        & " And DocFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59:59") & "'" _
        & " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
    
    If cSucursal.ItemData(cSucursal.ListIndex) > 0 Then Cons = Cons & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
    
    Cons = Cons _
        & " And DocAnulado = 0" _
        & " And EnvEstado IN ( " & EstadoEnvio.Impreso & ", " & EstadoEnvio.Entregado & ")" _
        & " And EnvLiquidacion = Null  And EnvFormaPago = " & TipoPagoEnvio.PagaDomicilio _
        & " And DocSucursal = SucCodigo And DocCodigo = EnvDocumentoFactura And EnvDocumento <> EnvDocumentoFactura And EnvCamion = CamCodigo" _
        & " And DocCodigo = RenDocumento And RenArticulo = ArtId"
    
    'Le uno las diferencias de Envios.
    Cons = Cons & " UNION ALL " _
        & " Select Documento.*, Renglon.*, Articulo.*, EnvCodigo, SucAbreviacion, CamNombre From Documento, Renglon, Envio, Articulo, DiferenciaEnvio, Sucursal, Camion " _
        & " Where DocTipo = " & TipoDocumento.Contado _
        & " And DocFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59:59") & "'" _
        & " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
    
    If cSucursal.ItemData(cSucursal.ListIndex) > 0 Then Cons = Cons & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
    
    Cons = Cons _
        & " And DocAnulado = 0" _
        & " And EnvEstado IN ( " & EstadoEnvio.Impreso & ", " & EstadoEnvio.Entregado & ")" _
        & " And EnvLiquidacion = Null " _
        & " And DEvFormaPago = " & TipoPagoEnvio.PagaDomicilio _
        & " And DocSucursal = SucCodigo And DocCodigo = DEvDocumento And EnvCodigo = DEvEnvio And EnvCamion = CamCodigo" _
        & " And DocCodigo = RenDocumento And RenArticulo = ArtId" _
        & " Order by DocSucursal, DocFecha"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If Not RsAux.EOF Then
        CargoDatos
        RsAux.Close
    Else
        RsAux.Close
        MsgBox "No hay datos a desplegar.", vbExclamation, "ATENCIÓN"
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrCDML:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos.", Err.Description
End Sub
Private Function ExisteEnGrilla(CodDoc As Long, IDArt As Long)
    ExisteEnGrilla = False
    With vsGrilla
        For I = 0 To .Rows - 1
            If .Cell(flexcpData, I, 0) = CodDoc And .Cell(flexcpData, I, 1) = IDArt Then ExisteEnGrilla = True: Exit Function
        Next I
    End With
End Function
Private Sub CargoDatos()
Dim aDoc As Long, aArt As Long
    Do While Not RsAux.EOF
        
        aDoc = RsAux!DocCodigo
        aArt = RsAux!ArtID
        'Esto es una chinura era la forma más rápida para no cambiar las consultas.
        If Not ExisteEnGrilla(aDoc, aArt) Then
            With vsGrilla
                .AddItem ""
                'Esto es una chinura era la forma más rápida para no cambiar las consultas.
                .Cell(flexcpData, .Rows - 1, 0) = aDoc
                .Cell(flexcpData, .Rows - 1, 1) = aArt
                .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!SucAbreviacion)
                .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!DocFecha, "dd/mm/yy hh:mm")
                .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!DocSerie) & " " & Trim(RsAux!DocNumero)
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!ArtCodigo, "#,000,000")
                .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 5) = RsAux!RenCantidad
                .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!RenCantidad * (RsAux!RenPrecio), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 7) = Trim(RsAux!CamNombre)
                .Cell(flexcpText, .Rows - 1, 8) = RsAux!EnvCodigo
            End With
        End If
        RsAux.MoveNext
    Loop
    With vsGrilla
        .Subtotal flexSTClear
        .Subtotal flexSTSum, 0, 6, , Obligatorio, , True
        .Subtotal flexSTSum, -1, 6, , Obligatorio, Rojo, True, "Total"
     End With
    
End Sub

Private Sub AccionLimpiar()
On Error Resume Next
    LimpioGrilla
    chVista.Value = 0
    With vsListado
        .StartDoc: .EndDoc
    End With
End Sub

Private Sub AccionConfigurar()
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
End Sub

Private Sub LimpioGrilla()
    With vsGrilla
        .Redraw = False
        .ExtendLastCol = True
        .Clear
        .Rows = 1
        .Cols = 1
        .FormatString = "Sucursal|<Hora|<Factura|>Código|<Nombre|>Cant|>Contado|<Camión|<Envío|"
        .ColWidth(0) = 110: .ColWidth(1) = 1250: .ColWidth(2) = 1000: .ColWidth(3) = 1000: .ColWidth(4) = 2800: .ColWidth(5) = 600: .ColWidth(6) = 1200: .ColWidth(7) = 1300: .ColWidth(8) = 2000: .ColWidth(9) = 20
        .MergeCells = flexMergeSpill
        '.MergeCol(0) = True
        .OutlineBar = flexOutlineBarSimple
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        .AllowUserResizing = flexResizeColumns
        .Redraw = True
    End With
End Sub
Private Sub Ayuda(strTexto As String)
    Status.Panels(3).Text = strTexto
End Sub
Private Function ValidoDatos() As Boolean
    ValidoDatos = False
    If cSucursal.ListIndex = -1 Then
        MsgBox "No selecciono una sucursal válida.", vbExclamation, "ATENCIÓN"
        cSucursal.SetFocus: Exit Function
    End If
    If Not IsDate(tDesde.Text) Then
        MsgBox "No ingreso una fecha válida, verifique.", vbExclamation, "ATENCIÓN"
        tDesde.SetFocus: Exit Function
    End If
    If Not IsDate(tHasta.Text) Then
        MsgBox "No ingreso una fecha válida, verifique.", vbExclamation, "ATENCIÓN"
        tHasta.SetFocus: Exit Function
    End If
    If CDate(tDesde.Text) > CDate(tHasta.Text) Then
        MsgBox "No se ingreso un rango de fechas válido, verifique.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    If cMoneda.ListIndex = -1 Then
        MsgBox "Seleccione una moneda.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
    ValidoDatos = True
End Function

