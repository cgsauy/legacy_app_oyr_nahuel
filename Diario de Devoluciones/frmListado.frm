VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmListado 
   Caption         =   "Diario de Devoluciones"
   ClientHeight    =   7530
   ClientLeft      =   1470
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   4440
      TabIndex        =   8
      Top             =   1800
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
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4455
      Left            =   120
      TabIndex        =   11
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
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   11595
      TabIndex        =   12
      Top             =   6720
      Width           =   11655
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5400
         Picture         =   "frmListado.frx":17C6
         Style           =   1  'Graphical
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   265
         Left            =   6000
         TabIndex        =   25
         Top             =   140
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   476
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
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
            Object.Width           =   10980
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
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   10335
      Begin AACombo99.AACombo cSucursal 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3420
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   1035
      End
      Begin AACombo99.AACombo cDocumento 
         Height          =   315
         Left            =   5640
         TabIndex        =   5
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
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
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   8160
         TabIndex        =   7
         Top             =   240
         Width           =   915
         _ExtentX        =   1614
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
         Caption         =   "&Sucursal:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   280
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Documento:"
         Height          =   255
         Left            =   4620
         TabIndex        =   4
         Top             =   280
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   2820
         TabIndex        =   2
         Top             =   280
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Moneda:"
         Height          =   255
         Left            =   7500
         TabIndex        =   6
         Top             =   285
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsAux As rdoResultset, rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean

Private Sub AccionLimpiar()
    tFecha.Text = ""
    cSucursal.Text = "": cMoneda.Text = "": cDocumento.Text = ""
    vsConsulta.Rows = 1
End Sub

Private Sub bCancelar_Click()
    Unload Me
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

Private Sub bConfigurar_Click()
    AccionConfigurar
End Sub

Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub

Private Sub cDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cMoneda
End Sub

Private Sub chVista_Click()
    
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
    Else
        AccionImprimir
        vsListado.ZOrder 0
    End If

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
    Foco tFecha
End Sub

Private Sub Label3_Click()
    Foco cDocumento
End Sub

Private Sub Label4_Click()
    Foco cMoneda
End Sub

Private Sub tFecha_GotFocus()
    With tFecha: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cDocumento
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
    
    CargoDatosCombos
    
    FechaDelServidor
    tFecha.Text = Format(gFechaServidor, "dd/mm/yyyy")

    bCargarImpresion = True
    With vsListado
        .PaperSize = 1
        .Orientation = orPortrait
        .Zoom = 100
        .MarginLeft = 900: .MarginRight = 350
        .MarginBottom = 750: .MarginTop = 750
    End With
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub CargoDatosCombos()

    On Error Resume Next
    cons = "Select SucCodigo, SucAbreviacion From Sucursal " _
        & " Where SucDcontado <> Null Or SucDCredito <> Null"
    CargoCombo cons, cSucursal, ""
    
    cDocumento.Clear
    cDocumento.AddItem "Notas de Devolución": cDocumento.ItemData(cDocumento.NewIndex) = TipoDocumento.NotaDevolucion
    cDocumento.AddItem "Notas de Crédito": cDocumento.ItemData(cDocumento.NewIndex) = TipoDocumento.NotaCredito
    cDocumento.AddItem "Notas Especiales": cDocumento.ItemData(cDocumento.NewIndex) = TipoDocumento.NotaEspecial
    
    cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo cons, cMoneda, ""
    BuscoCodigoEnCombo cMoneda, paMonedaFacturacion

End Sub


Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .OutlineBar = flexOutlineBarNone ' flexOutlineBarComplete
        .OutlineCol = 0
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<Sucursal|<Tipo|<Factura|Artículo|>Q|>Neto|Neto (A)|>I.V.A.|>I.V.A. (A)|>Subtotal|>Subtotal (A)|>Total|>Total (A)"
        
        .ColWidth(0) = 0: .ColWidth(1) = 0: .ColWidth(2) = 900: .ColWidth(3) = 3600: .ColWidth(4) = 500
        .ColWidth(5) = 1300: .ColWidth(6) = 0
        .ColWidth(7) = 1200: .ColWidth(8) = 0
        .ColWidth(9) = 1400: .ColWidth(10) = 0
        .ColWidth(11) = 1400: .ColWidth(12) = 0
            
        .WordWrap = False
        .MergeCells = flexMergeSpill
        .MergeCol(0) = True
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

    On Error GoTo errConsultar
    If Not ValidoCampos Then Exit Sub
    Screen.MousePointer = 11
    bCargarImpresion = True
    
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    cons = "Select Count(*) From Documento, Renglon"
    
    If cDocumento.ListIndex <> -1 Then
        Select Case cDocumento.ItemData(cDocumento.ListIndex)
            Case TipoDocumento.NotaCredito: cons = cons & " Where DocTipo = " & TipoDocumento.NotaCredito
            Case TipoDocumento.NotaDevolucion: cons = cons & " Where DocTipo = " & TipoDocumento.NotaDevolucion
            Case TipoDocumento.NotaEspecial: cons = cons & " Where DocTipo = " & TipoDocumento.NotaEspecial
        End Select
    Else
        cons = cons & " Where DocTipo In( " & TipoDocumento.NotaCredito & ", " & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")"
    End If
    
    cons = cons & " And DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tFecha.Text, "mm/dd/yyyy 23:59:59") & "'" _
                       & " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
    If cSucursal.ListIndex <> -1 Then cons = cons & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
    
    cons = cons & " And DocCodigo = RenDocumento "
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If rsAux(0) = 0 Then
            MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
            rsAux.Close: Screen.MousePointer = 0: InicializoGrillas: Exit Sub
        Else
            pbProgreso.Max = rsAux(0)
        End If
    End If
    rsAux.Close
    '-------------------------------------------------------------------------------------------------------------------------------------------
    
    cons = "Select * From Documento, Renglon, Articulo"
    
    If cDocumento.ListIndex <> -1 Then
        Select Case cDocumento.ItemData(cDocumento.ListIndex)
            Case TipoDocumento.NotaCredito: cons = cons & " Where DocTipo = " & TipoDocumento.NotaCredito
            Case TipoDocumento.NotaDevolucion: cons = cons & " Where DocTipo = " & TipoDocumento.NotaDevolucion
            Case TipoDocumento.NotaEspecial: cons = cons & " Where DocTipo = " & TipoDocumento.NotaEspecial
        End Select
    Else
        cons = cons & " Where DocTipo In( " & TipoDocumento.NotaCredito & ", " & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")"
    End If
    
    cons = cons & " And DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tFecha.Text, "mm/dd/yyyy 23:59:59") & "'" _
                       & " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
    If cSucursal.ListIndex <> -1 Then cons = cons & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
    
    cons = cons & " And DocCodigo = RenDocumento " _
                       & " And RenArticulo = ArtId" _
                       & " Order by DocSucursal, DocTipo, DocSerie, DocNumero"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        rsAux.Close: Screen.MousePointer = 0: InicializoGrillas: Exit Sub
    End If
    
    Dim aIdDocumento As Long: aIdDocumento = 0: Dim aIdTipo As Integer: aIdTipo = 0
    Dim aIdSucursal As Long: aIdSucursal = 0
    Dim aTxtSucursal As String: aTxtSucursal = "": Dim aTxtTipo As String: aTxtTipo = ""
    With vsConsulta
        .Rows = 1
        Do While Not rsAux.EOF
            pbProgreso.Value = pbProgreso.Value + 1
            
            If aIdSucursal <> rsAux!DocSucursal Then      '--------------------------------------------------------------
                
                aIdSucursal = rsAux!DocSucursal
                
                cons = "Select * from Sucursal Where SucCodigo = " & aIdSucursal
                Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                If Not rsAux.EOF Then aTxtSucursal = Trim(rs1!SucAbreviacion)
                rs1.Close
                
                aIdTipo = rsAux!DocTipo
                aTxtTipo = RetornoNombreDocumento(rsAux!DocTipo)
                
                If .Rows > 1 Then
                    .AddItem aTxtSucursal
                    .Cell(flexcpText, .Rows - 1, 1) = aTxtSucursal & " (" & aTxtTipo & ")"
                    .Cell(flexcpText, .Rows - 1, 2) = " "
                End If
                
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = aTxtSucursal
                .Cell(flexcpText, .Rows - 1, 1) = aTxtSucursal & " (" & aTxtTipo & ")"
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.GrisOscuro ': .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
                '.Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
            End If  '---------------------------------------------------------------------------------------------------------
            
            If aIdTipo <> rsAux!DocTipo Then
                aIdTipo = rsAux!DocTipo
                aTxtTipo = RetornoNombreDocumento(rsAux!DocTipo)
                If .Rows > 1 Then
                    .AddItem aTxtSucursal
                    .Cell(flexcpText, .Rows - 1, 1) = aTxtSucursal & " (" & aTxtTipo & ")"
                    .Cell(flexcpText, .Rows - 1, 2) = " "
                End If
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = aTxtSucursal
                .Cell(flexcpText, .Rows - 1, 1) = aTxtSucursal & " (" & aTxtTipo & ")"
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.GrisOscuro ': .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
                '.Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
            End If
                
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Trim(aTxtSucursal)
            .Cell(flexcpText, .Rows - 1, 1) = Trim(aTxtSucursal) & " (" & aTxtTipo & ")"
            .Cell(flexcpText, .Rows - 1, 2) = " "
            .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 4) = rsAux!RenCantidad
            .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!RenCantidad * (rsAux!RenPrecio - Format(rsAux!RenIva, FormatoMonedaP)), FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 7) = Format(Format(rsAux!RenIva, FormatoMonedaP) * rsAux!RenCantidad, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 9) = Format(rsAux!RenPrecio * rsAux!RenCantidad, FormatoMonedaP)
            If aIdDocumento <> rsAux!DocCodigo Then
                aIdDocumento = rsAux!DocCodigo
                .Cell(flexcpText, .Rows - 1, 2) = Trim(rsAux!DocSerie) & " " & Format(rsAux!DocNumero, "##000000")
                .Cell(flexcpText, .Rows - 1, 11) = Format(rsAux!DocTotal, FormatoMonedaP)
            End If
            
            If rsAux!DocAnulado Then
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris
            Else
                .Cell(flexcpText, .Rows - 1, 6) = .Cell(flexcpText, .Rows - 1, 5)
                .Cell(flexcpText, .Rows - 1, 8) = .Cell(flexcpText, .Rows - 1, 7)
                .Cell(flexcpText, .Rows - 1, 10) = .Cell(flexcpText, .Rows - 1, 9)
                .Cell(flexcpText, .Rows - 1, 12) = .Cell(flexcpText, .Rows - 1, 11)
            End If
            
            rsAux.MoveNext
        Loop
        rsAux.Close
        
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, 1, 6, , Colores.GrisOscuro, , False, "Subtotal %s"
        .Subtotal flexSTSum, 1, 8: .Subtotal flexSTSum, 1, 10: .Subtotal flexSTSum, 1, 12
        
        .Subtotal flexSTSum, 0, 6, , Colores.Rojo, Colores.Blanco, True, "Total %s"
        .Subtotal flexSTSum, 0, 8: .Subtotal flexSTSum, 0, 10: .Subtotal flexSTSum, 0, 12
        
        .AddItem ""
        .Subtotal flexSTSum, -1, 6, , Colores.Rojo, Colores.Blanco, True, "Total de Devoluciones"
        .Subtotal flexSTSum, -1, 8: .Subtotal flexSTSum, -1, 10: .Subtotal flexSTSum, -1, 12
        
        For I = 1 To .Rows - 1
            If .IsSubtotal(I) Then
                .Cell(flexcpText, I, 5) = .Cell(flexcpText, I, 6)
                .Cell(flexcpText, I, 7) = .Cell(flexcpText, I, 8)
                .Cell(flexcpText, I, 9) = .Cell(flexcpText, I, 10)
                .Cell(flexcpText, I, 11) = .Cell(flexcpText, I, 12)
            End If
        Next
    End With
    
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Function ValidoCampos() As Boolean
    On Error Resume Next
    ValidoCampos = False
    
    If Not IsDate(tFecha.Text) Then
        MsgBox "Debe ingresar la fecha para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
    ValidoCampos = True
    
End Function

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
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
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        aTexto = "Diario de Devoluciones " & Trim(cDocumento.Text) & " -  al  " & Trim(tFecha.Text) & "  (" & Trim(cMoneda.Text) & ")"
        EncabezadoListado vsListado, aTexto, False
        vsListado.FileName = "Diario de Devoluciones"
         
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        
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
    clsGeneral.OcurrioError "Error al realizar la impresión", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

