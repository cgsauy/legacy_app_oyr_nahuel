VERSION 5.00
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListado 
   Caption         =   "Cobranza de Cuotas"
   ClientHeight    =   7530
   ClientLeft      =   2130
   ClientTop       =   1875
   ClientWidth     =   10830
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
   ScaleWidth      =   10830
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   1200
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8LCtl.VSFlexGrid vsConsulta 
      Height          =   2295
      Left            =   2880
      TabIndex        =   26
      Top             =   3960
      Width           =   3615
      _cx             =   6376
      _cy             =   4048
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
      AutoSearchDelay =   2
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
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4455
      Left            =   120
      TabIndex        =   10
      Top             =   720
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
      TabIndex        =   11
      Top             =   6720
      Width           =   11655
      Begin VB.CommandButton bExcel 
         Height          =   310
         Left            =   3480
         Picture         =   "frmListado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Exportar a excel"
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0784
         Height          =   310
         Left            =   4560
         Picture         =   "frmListado.frx":0886
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4200
         Picture         =   "frmListado.frx":0DB8
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmListado.frx":1232
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmListado.frx":131C
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmListado.frx":1406
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3840
         Picture         =   "frmListado.frx":1640
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4920
         Picture         =   "frmListado.frx":1742
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5520
         Picture         =   "frmListado.frx":1B08
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":1C0A
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmListado.frx":1F0C
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmListado.frx":224E
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmListado.frx":2550
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6120
         TabIndex        =   24
         Top             =   135
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   476
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   7275
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
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
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10874
            Key             =   ""
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
      Height          =   660
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   10335
      Begin VB.TextBox txtHasta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin AACombo99.AACombo cSucursal 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
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
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   8400
         TabIndex        =   7
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&a:"
         Height          =   255
         Left            =   5760
         TabIndex        =   4
         Top             =   285
         Width           =   255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Moneda:"
         Height          =   255
         Left            =   7560
         TabIndex        =   6
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Sucursal:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   3720
         TabIndex        =   2
         Top             =   285
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'.FormatString = "<Sucursal|<Fecha|<Factura|Titular|Van|>Valor Cuota|>Valor Cuota|>Mora|>Mora|>Total|>Total|<Recibo"
Private Enum colGrid
    Sucursal = 0
    Fecha
    Factura
    Titular
    Van
    Valor1
    Valor2
    Mora1
    Mora2
    Total1
    Total2
    Recibo
End Enum

Private RsAux As rdoResultset, rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean

Private Sub AccionLimpiar()
    cSucursal.Text = "": cMoneda.Text = ""
    tFecha.Text = ""
    vsConsulta.Rows = 1
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub bExcel_Click()
On Error GoTo errBE
    Dim sFile As String
    sFile = fnc_Browse(1, Replace(Me.Caption, "/", "-") & ".xls", "Exportar a excel")
    If sFile = "" Then Exit Sub
    vsConsulta.SaveGrid sFile, flexFileExcel, SaveExcelSettings.flexXLSaveFixedRows Or SaveExcelSettings.flexXLSaveRaw
errBE:
End Sub

Private Function fnc_Browse(ByVal xToFile As Byte, ByVal sFileN As String, ByVal sDialogT As String, Optional bShowSave As Boolean = True) As String
On Error GoTo errCancel
fnc_Browse = ""
 
    'Inicializo INITDIR
'    fnc_ValDirectory
            
    With cdFile
        .CancelError = True
        .DialogTitle = sDialogT
    'Var global
        '.InitDir = mExportDir
        If bShowSave Then .FileName = sFileN
        .Flags = cdlOFNOverwritePrompt Or cdlOFNCreatePrompt Or cdlOFNPathMustExist
        Select Case xToFile '1-Excel;   2-csv;  3=html
            Case 1: .Filter = "Libro de Microsoft Excel|*.xls"
            Case 2: .Filter = "Archivo de texto (csv)|*.csv"
            Case 3: .Filter = "Archivo html (*.htm)|*.htm"""
        End Select
        If bShowSave Then
            .ShowSave
        Else
            .ShowOpen
        End If
    End With
    fnc_Browse = cdFile.FileName
errCancel:
End Function


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
    Me.Refresh

End Sub


Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub cSucursal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tFecha
End Sub

Private Sub Label1_Click()
    Foco cSucursal
End Sub

Private Sub Label2_Click()
    Foco cMoneda
End Sub

Private Sub tFecha_GotFocus()
    With tFecha: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsDate(tFecha.Text) Then Foco txtHasta
End Sub

Private Sub Label5_Click()
    Foco tFecha
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()

    On Error GoTo ErrLoad
    
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    InicializoGrillas
    AccionLimpiar
    
    'Cargo las sucursales y las monedas en loc combos-------------------------------------------------------------------------
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Where SucDcontado <> null Or SucDCredito <> Null"
    CargoCombo Cons, cSucursal
    
    Cons = "Select MonCodigo, MonSigno From Moneda"
    CargoCombo Cons, cMoneda
    BuscoCodigoEnCombo cMoneda, paMonedaFacturacion
    
    FechaDelServidor
    tFecha.Text = Format(gFechaServidor, "dd/mm/yyyy")
    txtHasta.Text = tFecha.Text
    '----------------------------------------------------------------------------------------------------------------------------------
    
    bCargarImpresion = True
    vsListado.PaperSize = 1
    vsListado.MarginRight = 350
    vsListado.Orientation = orPortrait
    vsListado.Zoom = 100
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    vsListado.MarginTop = 500
    vsListado.MarginBottom = 600
    
    With vsConsulta
        
        .OutlineBar = flexOutlineBarNone ' flexOutlineBarComplete
        .OutlineCol = 0
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<Sucursal|<Fecha|<Factura|Titular|Van|>Valor Cuota|>Valor Cuota|>Mora|>Mora|>Total|>Total|<Recibo"
            
        .WordWrap = False
        .ColHidden(colGrid.Sucursal) = True
        .ColWidth(colGrid.Factura) = 850: .ColWidth(3) = 0: .ColWidth(4) = 500
        .ColDataType(colGrid.Van) = flexDTString
        .ColDataType(5) = flexDTCurrency
        .ColWidth(5) = 1000: .ColWidth(7) = 850: .ColWidth(9) = 1100
        .ColWidth(6) = 0: .ColWidth(8) = 0: .ColWidth(10) = 0
        .ColWidth(10) = 850
        .MergeCells = flexMergeSpill
        
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
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = fFiltros.Width
    vsListado.Left = fFiltros.Left
    
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
    
    CierroConexion
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub AccionConsultar()

Dim aIDSucursal As Long, aTxtSucursal As String
Dim rs1 As rdoResultset
Dim mAportes As Currency
Dim idx As Integer, arrSucs() As String, arrData() As String

    On Error GoTo errConsultar
    If Not ValidoCampos Then Exit Sub
    
    Screen.MousePointer = 11
    vsConsulta.ZOrder 0
    Me.Refresh
    bCargarImpresion = True
    
    aIDSucursal = 0
    mAportes = 0
    
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    Cons = "Select Count(*) From Documento Left Outer Join DocumentoPago On DocCodigo = DPaDocQSalda" _
            & " Where DocTipo = " & TipoDocumento.ReciboDePago _
            & " And DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(txtHasta.Text, "mm/dd/yyyy 23:59:59") & "'"
        
            If cSucursal.ListIndex <> -1 Then Cons = Cons & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
            Cons = Cons & " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux(0) = 0 Then
            MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
            RsAux.Close: Screen.MousePointer = 0: vsConsulta.Rows = 1: Exit Sub
        End If
        pbProgreso.Max = RsAux(0)
    End If
    RsAux.Close
    '-------------------------------------------------------------------------------------------------------------------------------------------
    
    Cons = "Select Recibo.*, DocumentoPago.*, Factura.DocSerie FacSerie, Factura.DocNumero FacNumero, Factura.DocTipo FacTipo " _
            & " From Documento Recibo (index = iTipoFechaSucursalMoneda) " _
                    & " Left Outer Join DocumentoPago ON Recibo.DocCodigo = DPaDocQSalda " _
                            & " Left Outer Join Documento Factura ON DPaDocASaldar = Factura.DocCodigo " _
            & " Where Recibo.DocTipo = " & TipoDocumento.ReciboDePago _
            & " And Recibo.DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tFecha.Text, "mm/dd/yyyy 23:59:59") & "'"

    If cSucursal.ListIndex <> -1 Then Cons = Cons & " And Recibo.DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
            
    Cons = Cons & " And Recibo.DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
    Cons = Cons & " Order by Recibo.DocSucursal, Recibo.DocSerie, Recibo.DocNumero"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        Screen.MousePointer = 0: vsConsulta.Rows = 1: Exit Sub
    End If
    
    vsConsulta.Rows = 1: vsConsulta.Refresh
    Do While Not RsAux.EOF
        pbProgreso.Value = pbProgreso.Value + 1
        
        If IsNull(RsAux("FacTipo").Value) Or Not (RsAux("FacTipo").Value = TipoDocumento.NotaDebito) Then
        
            If aIDSucursal <> RsAux!DocSucursal Then
                aIDSucursal = RsAux!DocSucursal
                Cons = "Select * from Sucursal Where SucCodigo = " & RsAux!DocSucursal
                Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not rs1.EOF Then aTxtSucursal = Trim(rs1!SucAbreviacion) Else aTxtSucursal = ""
                rs1.Close
            End If
            
            With vsConsulta
                .AddItem aTxtSucursal
                
                .Cell(flexcpText, .Rows - 1, colGrid.Fecha) = Format(RsAux("DocFecha"), "dd/mm/yyyy")
                .Cell(flexcpText, .Rows - 1, colGrid.Factura) = Trim(RsAux!FacSerie) & " " & Trim(RsAux!FacNumero)
                            
                If Not IsNull(RsAux!DPaCuota) Then      'Van/Son--------------
                    If RsAux!DPaCuota = 0 Then aTexto = "E" Else aTexto = RsAux!DPaCuota
                    aTexto = aTexto & "/" & RsAux!DPaDe
                    .Cell(flexcpText, .Rows - 1, colGrid.Van) = aTexto
                    
                    If Not IsNull(RsAux!DPaAmortizacion) Then .Cell(flexcpText, .Rows - 1, colGrid.Valor1) = Format(RsAux!DPaAmortizacion, FormatoMonedaP)
                    If Not IsNull(RsAux!DPaMora) Then .Cell(flexcpText, .Rows - 1, colGrid.Mora1) = Format(RsAux!DPaMora, FormatoMonedaP) Else .Cell(flexcpText, .Rows - 1, colGrid.Mora1) = "0.00"
                
                Else        'Aportes a Cuentas
                    .Cell(flexcpText, .Rows - 1, 1) = "APORTE"
                    .Cell(flexcpText, .Rows - 1, colGrid.Valor1) = Format(RsAux!DocTotal, FormatoMonedaP)
                    .Cell(flexcpText, .Rows - 1, colGrid.Mora1) = "0.00"
                    If Not RsAux!DocAnulado Then mAportes = mAportes + .Cell(flexcpValue, .Rows - 1, colGrid.Valor1)
                End If
                
                .Cell(flexcpText, .Rows - 1, colGrid.Total1) = Format(.Cell(flexcpValue, .Rows - 1, colGrid.Valor1) + .Cell(flexcpValue, .Rows - 1, colGrid.Mora1), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, colGrid.Recibo) = Trim(RsAux!DocSerie) & " " & Trim(RsAux!DocNumero)     'Nro de Recibo
                
                If RsAux!DocAnulado Then
                    .Cell(flexcpText, .Rows - 1, colGrid.Valor2) = "0.00": .Cell(flexcpText, .Rows - 1, colGrid.Mora2) = "0.00"
                    .Cell(flexcpText, .Rows - 1, colGrid.Total2) = "0.00"
                    .Cell(flexcpBackColor, .Rows - 1, 1, , .Cols - 1) = vbButtonFace 'Colores.Gris
                Else
                    .Cell(flexcpText, .Rows - 1, colGrid.Valor2) = .Cell(flexcpText, .Rows - 1, colGrid.Valor1)
                    .Cell(flexcpText, .Rows - 1, colGrid.Mora2) = .Cell(flexcpText, .Rows - 1, colGrid.Mora1)
                    .Cell(flexcpText, .Rows - 1, colGrid.Total2) = .Cell(flexcpText, .Rows - 1, colGrid.Total1)
                End If
                
            End With

        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If vsConsulta.Rows > vsConsulta.FixedRows Then
    With vsConsulta
        .Subtotal flexSTSum, 0, colGrid.Valor2, , Colores.Rojo, Colores.Blanco, False, "%s"
        .Subtotal flexSTSum, 0, colGrid.Mora2:
        .Subtotal flexSTSum, 0, colGrid.Total2
        .Subtotal flexSTCount, 0, colGrid.Recibo, 0
        .AddItem ""
        .Subtotal flexSTSum, -1, colGrid.Valor2, , Colores.Rojo, Colores.Blanco, False, "Total Cobranzas"
        .Subtotal flexSTSum, -1, colGrid.Mora1
        .Subtotal flexSTSum, -1, colGrid.Total2
        .Subtotal flexSTCount, -1, colGrid.Recibo, 0
        
        idx = -1
        For I = 1 To .Rows - 1
            If .IsSubtotal(I) Then
                .Cell(flexcpText, I, colGrid.Valor1) = .Cell(flexcpText, I, colGrid.Valor2)
                .Cell(flexcpText, I, colGrid.Mora1) = .Cell(flexcpText, I, colGrid.Mora2)
                .Cell(flexcpText, I, colGrid.Total1) = .Cell(flexcpText, I, colGrid.Total2)
                
                idx = idx + 1
                ReDim Preserve arrSucs(idx)
                arrSucs(idx) = .Cell(flexcpText, I, colGrid.Fecha) & "|" & .Cell(flexcpText, I, colGrid.Valor1) & "|" & .Cell(flexcpText, I, colGrid.Valor2) & "|" & .Cell(flexcpText, I, colGrid.Total1)
            End If
        Next
        
        
        If mAportes <> 0 Then
            idx = idx + 1
            ReDim Preserve arrSucs(idx)
            arrSucs(idx) = arrSucs(idx - 1)
            
            arrSucs(idx - 1) = "Total Señas|" & Format(mAportes, FormatoMonedaP) & "||" & Format(mAportes, FormatoMonedaP)
            
            arrData = Split(arrSucs(idx), "|")
            arrData(3) = Format(arrData(3) - mAportes, FormatoMonedaP)
            arrData(1) = Format(arrData(1) - mAportes, FormatoMonedaP)
            arrSucs(idx) = Join(arrData, "|")
            
        End If
        
        If idx <> -1 Then
            .AddItem ""
            For idx = LBound(arrSucs) To UBound(arrSucs)
                arrData = Split(arrSucs(idx), "|")
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, colGrid.Fecha) = arrData(0)
                .Cell(flexcpText, .Rows - 1, colGrid.Valor1) = arrData(1)
                .Cell(flexcpText, .Rows - 1, colGrid.Mora1) = arrData(2)
                .Cell(flexcpText, .Rows - 1, colGrid.Total1) = arrData(3)
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.clCeleste ' &H800000    'Colores.Azul
                '.Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
            Next
        End If
        
        'Veo Si hay Creditos A PERDIDA----------------------------------------------------------------------------------------------
        Cons = "Select Sum(DPAAmortizacion) Suma, Count(*)  Cantidad from Documento, DocumentoPago, Credito " _
                & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFecha.Text, "mm/dd/yyyy") & " 23:59'"
        
        If cSucursal.ListIndex <> -1 Then Cons = Cons & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
    
        Cons = Cons & " And DocTipo = " & TipoDocumento.ReciboDePago _
                            & " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
                            & " And DocAnulado = 0" _
                            & " And DocCodigo = DPaDocQSalda And DPaDocASaldar = CreFactura" _
                            & " And CreTipo = " & TipoCredito.Incobrable
                
        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        If Not RsAux.EOF Then
            If Not IsNull(RsAux!Cantidad) Then
                If RsAux!Cantidad <> 0 Then
                    .AddItem "Amortización de Créd. A Pérdida"
                    .Cell(flexcpText, .Rows - 1, colGrid.Recibo) = RsAux!Cantidad
                    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.clCeleste   'Colores.Rojo
                    '.Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
                    If Not IsNull(RsAux!Suma) Then .Cell(flexcpText, .Rows - 1, colGrid.Mora2) = Format(RsAux!Suma, FormatoMonedaP)
                End If
            End If
        End If
        RsAux.Close
        '---------------------------------------------------------------------------------------------------------------------------------
    End With
    
    End If
    
    pbProgreso.Value = 0: Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
End Sub


Private Sub txtHasta_GotFocus()
    With txtHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And IsDate(txtHasta.Text) Then cMoneda.SetFocus
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
            .StartDoc
            vsListado.Columns = 2
            
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        EncabezadoListado vsListado, "Cobranza de Cuotas en " & Trim(cMoneda.Text) & " - Al " & Trim(tFecha.Text), False
        vsListado.FileName = "Cobranza de Cuotas"
        
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
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

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

Private Function ValidoCampos() As Boolean
    
    ValidoCampos = False
    
    If Not IsDate(tFecha.Text) Then
        MsgBox "Debe ingresar una fecha para realizar la consulta de datos.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    
    If Not IsDate(txtHasta.Text) Then
        MsgBox "Debe ingresar la fecha hasta para realizar la consulta de datos.", vbExclamation, "ATENCIÓN"
        Foco txtHasta: Exit Function
    End If
    
    If CDate(txtHasta.Text) < CDate(tFecha.Text) Then
        MsgBox "El período ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda para realizar la consulta de datos.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
    
    ValidoCampos = True
    
End Function
