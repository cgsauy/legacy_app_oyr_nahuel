VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmListado 
   Caption         =   "Asientos Varios"
   ClientHeight    =   7590
   ClientLeft      =   1860
   ClientTop       =   1905
   ClientWidth     =   11880
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
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   1995
      Left            =   120
      TabIndex        =   19
      Top             =   720
      Width           =   11415
      _Version        =   196608
      _ExtentX        =   20135
      _ExtentY        =   3519
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
      Zoom            =   100
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   2775
      Left            =   4680
      TabIndex        =   16
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4895
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
      GridLines       =   0
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
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   11475
      TabIndex        =   20
      Top             =   6720
      Width           =   11535
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   12
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   11
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
         Left            =   5400
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
         TabIndex        =   4
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   5880
         TabIndex        =   22
         Top             =   120
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
      TabIndex        =   18
      Top             =   7335
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12832
            TextSave        =   ""
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
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   11175
      Begin VB.CheckBox opNetear 
         Caption         =   "&Netear Asientos"
         Height          =   195
         Left            =   4500
         TabIndex        =   23
         Top             =   280
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox tFHasta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Text            =   "28/12/2000"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Text            =   "28/12/2000"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   255
         Width           =   735
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsMovimiento 
      Height          =   2415
      Left            =   840
      TabIndex        =   21
      Top             =   4200
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4260
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
      GridLines       =   0
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

Private Type typRenglon
    NombreRubro As String
    Importe As Currency
End Type

Dim arrAsientos() As typRenglon

Private Sub AccionLimpiar()
    tFecha.Text = "": tFHasta.Text = ""
    vsConsulta.Rows = 1: vsMovimiento.Rows = 1
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


Private Sub chVista_Click()
    
    If chVista.Value = 0 Then
        'vsConsulta.ZOrder 0
        vsListado.Visible = False
    Else
        AccionImprimir
        vsListado.ZOrder 0
        vsListado.Visible = True
    End If

End Sub

Private Sub Label2_Click()
    Foco tFHasta
End Sub

Private Sub opNetear_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub tFecha_GotFocus()
    With tFecha: .SelStart = 0: .SelLength = Len(.Text): End With
    Ayuda "Ingrese una fecha de compra."
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tFHasta
End Sub
Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then
        tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
        If Not IsDate(tFHasta.Text) Then tFHasta.Text = tFecha.Text
    End If
    Ayuda ""
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
    pbProgreso.Value = 0
    
    InicializoGrillas
    AccionLimpiar
    bCargarImpresion = True
    
    CargoConstantesSubrubros
    
    vsListado.Orientation = orPortrait
    vsListado.Visible = False
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    vsListado.MarginLeft = 1000
    With vsConsulta
        .ColSel = 0
        .ColSort(0) = flexSortStringAscending
        .Sort = flexSortUseColSort
        
        .Cols = 1: .Rows = 1:
        .FormatString = "Rubro Orden|<Rubro|>Debe $|>Debe M/E|>Haber $|>Haber M/E|"
            
        .WordWrap = False
        .ColWidth(0) = 0
        .ColWidth(1) = 3500: .ColWidth(2) = 1300: .ColWidth(3) = 1300: .ColWidth(4) = 1300:: .ColWidth(5) = 1300
    End With
    
    With vsMovimiento
        .Cols = 1: .Rows = 1:
        .FormatString = "<Movimiento|<Concepto|>Debe $|>Haber $|"
            
        .WordWrap = False
        .ColWidth(0) = 3500: .ColWidth(1) = 4000: .ColWidth(2) = 1300: .ColWidth(3) = 1300
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
    
    picBotones.Width = vsListado.Width
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 150
    
    Dim Altura As Currency
    Altura = vsListado.Height
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Left = vsListado.Left
    If vsMovimiento.Visible Then vsConsulta.Height = (Altura / 5) * 4 Else vsConsulta.Height = Altura
    
    vsMovimiento.Width = vsListado.Width
    vsMovimiento.Left = vsListado.Left
    vsMovimiento.Top = vsConsulta.Top + vsConsulta.Height + 40
    vsMovimiento.Height = (Altura / 5) - 40
    
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Private Sub AccionConsultar()
Dim aSR As Long, aTipoM As Long
Dim aSROrden As String

    If Not ValidoCampos Then Exit Sub
    
    On Error GoTo errConsultar
    Screen.MousePointer = 11
    bCargarImpresion = True
    vsConsulta.Rows = 1: vsMovimiento.Rows = 1
    vsConsulta.Refresh: vsMovimiento.Refresh
    
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    cons = " Select Count(Distinct(MDiTipo)) from MovimientoDisponibilidad " _
            & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDiIDCompra Is Null "
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If rsAux(0) = 0 Then
            MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
            rsAux.Close: Screen.MousePointer = 0: vsConsulta.Rows = 1: Exit Sub
        End If
        pbProgreso.Max = rsAux(0)
    End If
    rsAux.Close
    '-----------------------------------------------------------------------------------------------------------------
    
    cons = " Select TMDNombre, MDiTipo, DH = 'Haber', Importe = Sum(MDrImportePesos), IOriginal = Sum(MDrHaber)" _
            & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, TipoMovDisponibilidad " _
            & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDiID = MDRIDMovimiento " _
            & " And MDiIDCompra Is Null " _
            & " And MDiTipo = TMDCodigo " _
            & " And MDRHaber <> Null " _
            & " Group by TMDNombre, MDiTipo"
    
    cons = cons & " UNION ALL "
    
    cons = cons & " Select TMDNombre, MDiTipo, DH = 'Debe', Importe = Sum(MDrImportePesos), IOriginal = Sum(MDrHaber)" _
                        & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, TipoMovDisponibilidad " _
                        & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
                        & " And MDiID = MDRIDMovimiento " _
                        & " And MDiIDCompra Is Null " _
                        & " And MDiTipo = TMDCodigo " _
                        & " And MDRDebe <> Null " _
                        & " Group by TMDNombre, MDiTipo"
            
    cons = cons & " Order by TMDNombre "
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        rsAux.Close: Screen.MousePointer = 0: Exit Sub
    End If
    
    vsConsulta.Redraw = False
    aTipoM = 0
    pbProgreso.Value = 0
    Do While Not rsAux.EOF
        aSR = 0
        With vsConsulta
            
            Select Case rsAux!MDiTipo
                Case paMCNotaCredito, paMCAnulacion: aSR = paSubrubroDeudoresPorVenta
                Case paMCChequeDiferido: aSR = paSubrubroCDAlCobro
                Case paMCVtaTelefonica: aSR = paSubrubroVtasTelACobrar
                Case paMCLiquidacionCamionero: aSR = paSubrubroCobranzaVtasTel
                Case paMCSenias: aSR = paSubrubroSeniasRecibidas
                
                Case Else
                    'Consulto en el movimiento si es del tipo tranferencia
                    cons = "Select * from TipoMovDisponibilidad Where TMDCodigo = " & rsAux!MDiTipo
                    Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    If Not rs1.EOF Then
                        If Not IsNull(rs1!TMDTransferencia) Then If rs1!TMDTransferencia = 1 Then aSR = -1
                    End If
                    rs1.Close
            End Select
            
            If aSR <> 0 Then
            
                If aTipoM <> rsAux!MDiTipo Then
                    pbProgreso.Value = pbProgreso.Value + 1
                    aTipoM = rsAux!MDiTipo
                    
                    .AddItem aTipoM 'Nombre del Movimiento
                    .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!TMDNombre)
                    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
                    
                    CargoConceptos aTipoM   'Hay que cargar contra que conceptos van
                    
                End If
                
                If aSR <> -1 Then       'NO ES TRANSFERENCIA
                    .AddItem ""
                    aTexto = RetornoConstanteSubrubro(aSR)
                    .Cell(flexcpText, .Rows - 1, 0) = aTipoM
                    .Cell(flexcpText, .Rows - 1, 1) = aTexto
                    
                    Select Case LCase(rsAux!DH)     'Van al reves por los tipos de movimientos
                        Case "debe"
                                    .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Importe, FormatoMonedaP)
                                    If rsAux!IOriginal <> rsAux!Importe Then .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!IOriginal, FormatoMonedaP)
                        Case "haber":
                                    .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!Importe, FormatoMonedaP)
                                    If rsAux!IOriginal <> rsAux!Importe Then .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!IOriginal, FormatoMonedaP)
                    End Select
                End If
            Else
                If aTipoM <> rsAux!MDiTipo Then
                    pbProgreso.Value = pbProgreso.Value + 1
                    aTipoM = rsAux!MDiTipo
                    If aTipoM <> paMCIngresosOperativos Then CargoListaMovimiento aTipoM
                End If
            End If
            
            rsAux.MoveNext
        End With
    Loop
    rsAux.Close
    
    With vsConsulta
        If vsConsulta.Rows > 1 Then
            .Select 1, 0, 1, 1
            .Sort = flexSortGenericDescending
        End If
    End With
    
    EliminoAsientosDobles   'Recorro para eliminar asientos dobles--------------------------------
    
    CargoOtrosAsientos
    CargoAsientosDeVentas
    
    If vsMovimiento.Rows > 1 Then
        With vsMovimiento
            .SubtotalPosition = flexSTBelow
            .Subtotal flexSTSum, -1, 2, , Colores.Inactivo, , , "Total"
            .Subtotal flexSTSum, -1, 3
        End With
        vsConsulta.Height = (vsListado.Height / 5) * 4: vsMovimiento.Visible = True: vsMovimiento.Refresh
    Else
        vsConsulta.Height = vsListado.Height: vsMovimiento.Visible = False
    End If
    
    vsConsulta.Redraw = True
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    clsGeneral.OcurrioError "Error al realizar la consulta de datos." & vbCrLf & cons, Err.Description
    vsConsulta.Redraw = True: Screen.MousePointer = 0
End Sub

Private Sub CargoConceptos(aIDTipoM As Long, Optional Transferencia As Boolean = False)

    cons = " Select SRuCodigo, SRuNombre, DH = 'Haber', Importe = Sum(MDrImportePesos), IOriginal = Sum(MDRHaber)"
    cons = cons & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, Disponibilidad, SubRubro " _
                        & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
                        & " And MDiID = MDRIDMovimiento " _
                        & " And MDiIDCompra Is Null " _
                        & " And MDiTipo = " & aIDTipoM _
                        & " And MDRIDDisponibilidad = DisID " _
                        & " And DisIDSubrubro = SRuID" _
                        & " And MDRHaber <> Null " _
                        & " Group by SRuCodigo, SRuNombre"
    
    cons = cons & " UNION ALL "
    
    cons = cons & " Select SRuCodigo, SRuNombre, DH = 'Debe', Importe = Sum(MDrImportePesos), IOriginal = Sum(MDRDebe) "
    cons = cons & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, Disponibilidad, SubRubro " _
                        & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
                        & " And MDiID = MDRIDMovimiento " _
                        & " And MDiIDCompra Is Null " _
                        & " And MDiTipo = " & aIDTipoM _
                        & " And MDRIDDisponibilidad = DisID " _
                        & " And DisIDSubrubro = SRuID" _
                        & " And MDRDebe <> Null " _
                        & " Group by SRuCodigo, SRuNombre"
            
    Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rs1.EOF
        
        With vsConsulta
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = aIDTipoM
             aTexto = Format(rs1!SRuCodigo, "000000000") & " " & Trim(rs1!SRuNombre)
            .Cell(flexcpText, .Rows - 1, 1) = aTexto
            
            Select Case LCase(rs1!DH)
                Case "debe"
                    .Cell(flexcpText, .Rows - 1, 2) = Format(rs1!Importe, FormatoMonedaP)
                    If rs1!IOriginal <> rs1!Importe Then .Cell(flexcpText, .Rows - 1, 3) = Format(rs1!IOriginal, FormatoMonedaP)
                
                Case "haber"
                    .Cell(flexcpText, .Rows - 1, 4) = Format(rs1!Importe, FormatoMonedaP)
                    If rs1!IOriginal <> rs1!Importe Then .Cell(flexcpText, .Rows - 1, 5) = Format(rs1!IOriginal, FormatoMonedaP)
            End Select
        End With
        
        rs1.MoveNext
    Loop
    rs1.Close
    
End Sub

Private Sub CargoAsientosDeVentas()
    Dim aCofis As Currency
    Dim aSRCaja As String
    
    AsientoVentasContado
    
    
    'Saco el subrubro de la disponibilidad (Caja)  para hacer el otro asiento--------------------------------------------
    cons = "Select * from Disponibilidad, Subrubro " _
           & " Where DisID = " & paDisponibilidad _
           & " And DisIDSubRubro = SRuId"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        aSRCaja = Format(rsAux!SRuCodigo, "000000000") & " " & Trim(rsAux!SRuNombre)
    Else
        aSRCaja = "CAJA"
    End If
    rsAux.Close
    '------------------------------------------------------------------------------------------------------------------------------------
    
    GoTo etCredito
    
    ReDim arrAsientos(0)
    'ASIENTO DE VENTAS CONTADO PESOS--------------------------------------------------------------------------------------------------------
    cons = "Select Sum(DocTotal) Total, Sum(DocIva) Iva, Sum(DocCofis) Cofis from Documento " _
            & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
            & " And DocTipo = " & TipoDocumento.Contado _
            & " And DocAnulado = 0" _
            & " And DocMoneda = " & paMonedaPesos
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!Total) Then
        If rsAux!Total <> 0 Then
            If Not IsNull(rsAux!Cofis) Then aCofis = rsAux!Cofis Else aCofis = 0
            With vsConsulta
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = "Ventas Contado (Moneda Nacional)"
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVContado)
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!Total, FormatoMonedaP)
        
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentas)
            .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Total - rsAux!Iva - aCofis, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubRubroCofis)
            .Cell(flexcpText, .Rows - 1, 4) = Format(aCofis, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIVA)
            .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Iva, FormatoMonedaP)
            
            
            With arrAsientos(0)
                .NombreRubro = aSRCaja
                .Importe = Format(rsAux!Total, FormatoMonedaP)
            End With
            
            End With
        End If
        End If
    End If
    rsAux.Close
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------
    'meVentasContado
    
    With vsConsulta         'ASIENTO RESUMEN DE VENTAS CONTADO
        Dim mTotalAsiento As Currency
        mTotalAsiento = 0
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = "Ventas Contado"
        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
        
        For I = LBound(arrAsientos) To UBound(arrAsientos)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = arrAsientos(I).NombreRubro 'aSRCaja
            .Cell(flexcpText, .Rows - 1, 2) = Format(arrAsientos(I).Importe, FormatoMonedaP)
            mTotalAsiento = mTotalAsiento + arrAsientos(I).Importe
        Next
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVContado)
        .Cell(flexcpText, .Rows - 1, 4) = Format(mTotalAsiento, FormatoMonedaP)
    End With
    
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------

etCredito:

    'ASIENTO DE VENTAS CREDITO--------------------------------------------------------------------------------------------------------
    cons = "Select Sum(DocTotal) Total, Sum(DocIva) IVA, Sum(DocCofis) Cofis from Documento " _
            & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
            & " And DocTipo = " & TipoDocumento.Credito _
            & " And DocAnulado = 0"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!Total) Then
        If rsAux!Total <> 0 Then
            If Not IsNull(rsAux!Cofis) Then aCofis = rsAux!Cofis Else aCofis = 0
            With vsConsulta
            .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Ventas Crédito": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroDeudoresPorVenta)
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!Total, FormatoMonedaP)
        
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentas)
            .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Total - rsAux!Iva - aCofis, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubRubroCofis)
            .Cell(flexcpText, .Rows - 1, 4) = Format(aCofis, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIVA)
            .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Iva, FormatoMonedaP)
            End With
        End If
        End If
    End If
    rsAux.Close
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'ASIENTO COBRANZA DE MOROSIDADES Y CTAS--------------------------------------------------------------------------------------------------------
    Dim aMoraTotal As Currency
    cons = "Select Sum(DPaMora) Total from Documento, DocumentoPago " _
            & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
            & " And DocTipo = " & TipoDocumento.ReciboDePago _
            & " And DocAnulado = 0" _
            & " And DocCodigo = DPaDocQSalda and DPaMora > 0"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!Total) Then aMoraTotal = rsAux!Total
    End If
    rsAux.Close
    
    'A la Mora, Le sumo la cobranza de creditos a Perdida (Amortizacion --> La Mora ya esta)
    cons = "Select Sum(DPAAmortizacion) Suma from Documento, DocumentoPago, Credito " _
                & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
                & " And DocTipo = " & TipoDocumento.ReciboDePago _
                & " And DocAnulado = 0" _
                & " And DocCodigo = DPaDocQSalda And DPaDocASaldar = CreFactura" _
                & " And CreTipo = " & TipoCredito.Incobrable
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!Suma) Then aMoraTotal = aMoraTotal + rsAux!Suma
    End If
    rsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------------
    
    cons = "Select Sum(DocTotal) Total, Sum(DocIva) IVA from Documento " _
            & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
            & " And DocTipo = " & TipoDocumento.ReciboDePago _
            & " And DocAnulado = 0"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!Total) Then
        If rsAux!Total <> 0 Then
            With vsConsulta
            If aMoraTotal <> 0 Then
                'MORAS
                .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Cobranza de Morosidades": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
                
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = aSRCaja
                .Cell(flexcpText, .Rows - 1, 2) = Format(aMoraTotal, FormatoMonedaP)
            
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIngresosVarios)
                .Cell(flexcpText, .Rows - 1, 4) = Format(aMoraTotal - rsAux!Iva, FormatoMonedaP)   'En recibos el IVA es sobre la MORA
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIVA)
                .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Iva, FormatoMonedaP)
            End If
            
            'Primero saco las ctas por senias
            Dim rs1 As rdoResultset, aSenias As Currency: aSenias = 0
            cons = "Select Sum(DocTotal) Total from Documento Left Outer Join DocumentoPago on DocCodigo = DPaDocQSalda " _
                    & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
                    & " And DocTipo = " & TipoDocumento.ReciboDePago _
                    & " And DocAnulado = 0 And DPaDocQSalda is null And DPaDocASaldar is null"
            Set rs1 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not rs1.EOF Then
                If Not IsNull(rs1!Total) Then aSenias = rs1!Total
            End If
            rs1.Close
            
            'CUOTAS
            .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Cobranza de Cuotas": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = aSRCaja
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!Total - aMoraTotal, FormatoMonedaP)
            
            If rsAux!Total - aSenias > 0 Then   'hay pago de ctas
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroDeudoresPorVenta)
                .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Total - aMoraTotal - aSenias, FormatoMonedaP)
            End If
            
            If aSenias > 0 Then   'hay senias
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroSeniasRecibidas)
                .Cell(flexcpText, .Rows - 1, 4) = Format(aSenias, FormatoMonedaP)
            End If
            End With
        End If
        End If
    End If
    rsAux.Close
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------
    'ASIENTO DE NOTAS DEV Y ESPECIALES-------------------------------------------------------------------------------------------------------------
    cons = "Select Sum(DocTotal) Total, Sum(DocIva) IVA, Sum(DocCofis) Cofis from Documento " _
            & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
            & " And DocTipo IN (" & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")" _
            & " And DocAnulado = 0"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!Total) Then
        If rsAux!Total <> 0 Then
            If Not IsNull(rsAux!Cofis) Then aCofis = rsAux!Cofis Else aCofis = 0
            With vsConsulta
            .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Notas de Devolución y Notas Especiales": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentas)
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!Total - rsAux!Iva - aCofis, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubRubroCofis)
            .Cell(flexcpText, .Rows - 1, 2) = Format(aCofis, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIVA)
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!Iva, FormatoMonedaP)
                    
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVContado)
            .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Total, FormatoMonedaP)
            
            .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Notas de Devolución y Notas Especiales": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVContado)
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!Total, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = aSRCaja
            .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Total, FormatoMonedaP)
            End With
        End If
        End If
    End If
    rsAux.Close
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------
    'ASIENTO DE NOTAS DE CREDITO-------------------------------------------------------------------------------------------------------------
    cons = "Select Sum(DocTotal) Total, Sum(DocIva) IVA,  Sum(DocCofis) Cofis from Documento " _
            & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
            & " And DocTipo = " & TipoDocumento.NotaCredito _
            & " And DocAnulado = 0"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux!Total) Then
        If rsAux!Total <> 0 Then
            If Not IsNull(rsAux!Cofis) Then aCofis = rsAux!Cofis Else aCofis = 0
            With vsConsulta
            .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Notas de Crédito": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentas)
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!Total - rsAux!Iva - aCofis, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubRubroCofis)
            .Cell(flexcpText, .Rows - 1, 2) = Format(aCofis, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIVA)
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!Iva, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroDeudoresPorVenta)
            .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Total, FormatoMonedaP)
            
            End With
        End If
        End If
    End If
    rsAux.Close
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub AsientoVentasContado()

Dim rsSuc As rdoResultset
Dim rsData As rdoResultset
Dim idX As Integer
Dim mSubRubro As String
    
    ReDim arrAsientos(0)
    idX = 0
    
    Dim mSucursal As Long
    Dim mIDDisponibilidad As Long
    Dim mTotal As Currency, mIva As Currency, mCofis As Currency
    
    cons = "Select * from Sucursal Where SucDisponibilidad Is Not Null"
    Set rsSuc = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsSuc.EOF
        mSucursal = rsSuc!SucCodigo
        Ayuda "Cargando Ventas Contado M/N (" & Trim(rsSuc!SucAbreviacion) & ")..."
        
        'ASIENTO DE VENTAS CONTADO PESOS--------------------------------------------------------------------------------------------------------
        cons = "Select Sum(DocTotal) Total, Sum(DocIva) Iva, Sum(DocCofis) Cofis from Documento " _
                & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
                & " And DocTipo = " & TipoDocumento.Contado _
                & " And DocAnulado = 0" _
                & " And DocMoneda = " & paMonedaPesos _
                & " And DocSucursal = " & mSucursal
                
        Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        If Not rsAux.EOF Then
            If Not IsNull(rsAux!Total) Then
            If rsAux!Total <> 0 Then
            
                If Not IsNull(rsAux!Cofis) Then mCofis = mCofis + rsAux!Cofis
                mTotal = mTotal + Format(rsAux!Total, FormatoMonedaP)
                mIva = mIva + Format(rsAux!Iva, FormatoMonedaP)
                
                mIDDisponibilidad = dis_DisponibilidadPara(mSucursal, CLng(paMonedaPesos))
                
                'Saco el subrubro de la disponibilidad para hacer el asiento final --------------------------------------------
                cons = "Select * from Disponibilidad, Subrubro " _
                       & " Where DisID = " & mIDDisponibilidad _
                       & " And DisIDSubRubro = SRuId"
                Set rsData = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                If Not rsData.EOF Then
                    mSubRubro = Format(rsData!SRuCodigo, "000000000") & " " & Trim(rsData!SRuNombre)
                Else
                    mSubRubro = "CAJA M/N"
                End If
                rsData.Close
                '------------------------------------------------------------------------------------------------------------------------------------
                
                'Si el Subrubro está en el array --> Acumulo los datos, si no lo agrego
                If idX = 0 Then
                    With arrAsientos(0)
                        .NombreRubro = mSubRubro
                        .Importe = Format(rsAux!Total, FormatoMonedaP)
                    End With
                Else
                    Dim bOK As Boolean: bOK = False
                    For idX = LBound(arrAsientos) To UBound(arrAsientos)
                        If arrAsientos(idX).NombreRubro = mSubRubro Then
                            arrAsientos(idX).Importe = arrAsientos(idX).Importe + Format(rsAux!Total, FormatoMonedaP)
                            bOK = True
                            Exit For
                        End If
                    Next
                    If Not bOK Then
                        ReDim Preserve arrAsientos(UBound(arrAsientos) + 1)
                        With arrAsientos(UBound(arrAsientos))
                            .NombreRubro = mSubRubro
                            .Importe = Format(rsAux!Total, FormatoMonedaP)
                        End With
                    End If
                End If
                
                idX = 1
                
            End If
            
            End If
        End If
        rsAux.Close
            
        rsSuc.MoveNext
    Loop
    rsSuc.Close
    
    'ASIENTO DE VENTAS CONTADO EN M/N
    With vsConsulta
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = "Ventas Contado (Moneda Nacional)"
        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVContado)
        .Cell(flexcpText, .Rows - 1, 2) = Format(mTotal, FormatoMonedaP)
    
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentas)
        .Cell(flexcpText, .Rows - 1, 4) = Format(mTotal - mIva - mCofis, FormatoMonedaP)
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubRubroCofis)
        .Cell(flexcpText, .Rows - 1, 4) = Format(mCofis, FormatoMonedaP)
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIVA)
        .Cell(flexcpText, .Rows - 1, 4) = Format(mIva, FormatoMonedaP)
    End With
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Cargo Asientos de VENTAS CONTADO EN M/E    --------------------------------------------------------------------------------------------------
    Dim rsMon As rdoResultset
    Dim mName As String, mCodigo As Long
    
    cons = "Select * from Moneda Where MonCodigo <> " & paMonedaPesos '& " And MonCoeficienteMora Is Not Null "
    Set rsMon = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsMon.EOF
        mName = Trim(rsMon!MonNombre)
        mCodigo = rsMon!MonCodigo
        
        meVentasContado mCodigo, mName
        
        rsMon.MoveNext
    Loop
    rsMon.Close
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------
    
    With vsConsulta         'ASIENTO RESUMEN DE VENTAS CONTADO (TODAS LAS MONEDAS CONTRA VENTAS)
        Dim mTotalAsiento As Currency
        mTotalAsiento = 0
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = "Ventas Contado"
        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
        
        For I = LBound(arrAsientos) To UBound(arrAsientos)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = arrAsientos(I).NombreRubro
            .Cell(flexcpText, .Rows - 1, 2) = Format(arrAsientos(I).Importe, FormatoMonedaP)
            mTotalAsiento = mTotalAsiento + arrAsientos(I).Importe
        Next
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVContado)
        .Cell(flexcpText, .Rows - 1, 4) = Format(mTotalAsiento, FormatoMonedaP)
    End With
    
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------
    
End Sub

Private Sub CargoOtrosAsientos()
    
    Dim aTC As Currency
            
    'ASIENTO DE CHEQUES DIF A PAGAR (LOS Q VENCIERON)------------------------------------------------------------------------------------------------
    'Cargo los Rubros de la Dispnibilidad con los Cheuqes Diferidos a Pagar (Van al reves por q son la contrapartida de los siguientes)
    cons = "Select DisMoneda, SRuCodigo, SRuNombre, Importe = Sum(MDRImportePesos), DH = 'Haber', IOriginal = Sum(MDRHaber) " _
            & " From Cheque, MovimientoDisponibilidadRenglon, Disponibilidad, SubRubro" _
            & " Where CheID = MDRIDCheque " _
            & " And MDRIdDisponibilidad = DisID " _
            & " And DisIDSRCheque = SRuID " _
            & " And CheVencimiento Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDRHaber Is Not Null " _
            & " Group by DisMoneda, SRuCodigo, SRuNombre" _
                    & " UNION ALL " _
            & " Select DisMoneda, SRuCodigo, SRuNombre, Importe = Sum(MDRImportePesos), DH = 'Debe', IOriginal = Sum(MDRDebe)" _
            & " From Cheque, MovimientoDisponibilidadRenglon, Disponibilidad, SubRubro" _
            & " Where CheID = MDRIDCheque " _
            & " And MDRIdDisponibilidad = DisID " _
            & " And DisIDSRCheque = SRuID " _
            & " And CheVencimiento Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDRDebe Is Not Null " _
            & " Group by DisMoneda, SRuCodigo, SRuNombre"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        With vsConsulta
        .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Vto. de Cheques Diferidos a Pagar": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
        Do While Not rsAux.EOF
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!SRuCodigo, "000000000") & " " & Trim(rsAux!SRuNombre)
    
            'TC del ultimo dia del mes anterior
            aTC = TasadeCambio(rsAux!DisMoneda, paMonedaPesos, UltimoDia(DateAdd("m", -1, CDate(tFecha.Text))))
            
            Select Case LCase(rsAux!DH)
                Case "haber"
                    .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!Importe, FormatoMonedaP)
                    If rsAux!IOriginal <> rsAux!Importe Then
                        .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!IOriginal, FormatoMonedaP)
                        .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!IOriginal * aTC, FormatoMonedaP)
                    End If
                
                Case "debe"
                    .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Importe, FormatoMonedaP)
                    If rsAux!IOriginal <> rsAux!Importe Then
                        .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!IOriginal, FormatoMonedaP)
                        .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!IOriginal * aTC, FormatoMonedaP)
                    End If
            End Select
            rsAux.MoveNext
        Loop
        End With
    End If
    rsAux.Close
    
    cons = "Select DisMoneda, SRuCodigo, SRuNombre, Importe = Sum(MDRImportePesos), DH = 'Haber', IOriginal = Sum(MDRHaber) " _
            & " From Cheque, MovimientoDisponibilidadRenglon, Disponibilidad, SubRubro" _
            & " Where CheID = MDRIDCheque " _
            & " And MDRIdDisponibilidad = DisID " _
            & " And DisIDSubrubro = SRuID " _
            & " And CheVencimiento Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDRHaber Is Not Null " _
            & " Group by DisMoneda, SRuCodigo, SRuNombre" _
                    & " UNION ALL " _
            & " Select DisMoneda, SRuCodigo, SRuNombre, Importe = Sum(MDRImportePesos), DH = 'Debe', IOriginal = Sum(MDRDebe)" _
            & " From Cheque, MovimientoDisponibilidadRenglon, Disponibilidad, SubRubro" _
            & " Where CheID = MDRIDCheque " _
            & " And MDRIdDisponibilidad = DisID " _
            & " And DisIDSubrubro = SRuID " _
            & " And CheVencimiento Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDRDebe Is Not Null " _
            & " Group by DisMoneda, SRuCodigo, SRuNombre"
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then
        With vsConsulta
        Do While Not rsAux.EOF
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!SRuCodigo, "000000000") & " " & Trim(rsAux!SRuNombre)
            
            'TC del ultimo dia del mes anterior
            aTC = TasadeCambio(rsAux!DisMoneda, paMonedaPesos, UltimoDia(DateAdd("m", -1, CDate(tFecha.Text))))
            
            Select Case LCase(rsAux!DH)
                Case "debe"
                    .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!Importe, FormatoMonedaP)
                    If rsAux!IOriginal <> rsAux!Importe Then
                        .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!IOriginal, FormatoMonedaP)
                        .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!IOriginal * aTC, FormatoMonedaP)
                    End If
                
                Case "haber"
                    .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!Importe, FormatoMonedaP)
                    If rsAux!IOriginal <> rsAux!Importe Then
                        .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!IOriginal, FormatoMonedaP)
                        .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!IOriginal * aTC, FormatoMonedaP)
                    End If
            End Select
            rsAux.MoveNext
        Loop
        End With
    End If
    rsAux.Close
    
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub CargoListaMovimiento(aIDTipoM As Long)

    cons = " Select TMDNombre, MDiComentario, DH = 'Haber', Importe = Sum(MDrImportePesos)" _
            & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, TipoMovDisponibilidad " _
            & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDiTipo = " & aIDTipoM _
            & " And MDiID = MDRIDMovimiento " _
            & " And MDiIDCompra Is Null " _
            & " And MDiTipo = TMDCodigo " _
            & " And MDRHaber <> Null " _
            & " Group by TMDNombre, MDiComentario"
    
    cons = cons & " UNION ALL "
    
    cons = cons & " Select TMDNombre, MDiComentario, DH = 'Debe', Importe = Sum(MDrImportePesos)" _
                        & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, TipoMovDisponibilidad " _
                        & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
                        & " And MDiTipo = " & aIDTipoM _
                        & " And MDiID = MDRIDMovimiento " _
                        & " And MDiIDCompra Is Null " _
                        & " And MDiTipo = TMDCodigo " _
                        & " And MDRDebe <> Null " _
                        & " Group by TMDNombre, MDiComentario"
    
    Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rs1.EOF
        
        With vsMovimiento
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = rs1!TMDNombre
            If Not IsNull(rs1!MDiComentario) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(rs1!MDiComentario)
            
            Select Case LCase(rs1!DH)
                Case "debe": .Cell(flexcpText, .Rows - 1, 2) = Format(rs1!Importe, FormatoMonedaP)
                Case "haber": .Cell(flexcpText, .Rows - 1, 3) = Format(rs1!Importe, FormatoMonedaP)
            End Select
        End With
        
        rs1.MoveNext
    Loop
    rs1.Close
    
End Sub

Private Sub tFHasta_GotFocus()
    With tFHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tFHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then opNetear.SetFocus
End Sub

Private Sub tFHasta_LostFocus()
    If IsDate(tFHasta.Text) Then tFHasta.Text = Format(tFHasta.Text, "dd/mm/yyyy")
End Sub

Private Sub vsConsulta_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)

    vsConsulta.Row = vsConsulta.MouseRow
    
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
        
        EncabezadoListado vsListado, "Asientos Varios- Del " & Trim(tFecha.Text) & " al " & Trim(tFHasta.Text), False
        vsListado.FileName = "Listado Asientos Varios"
            
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        
        If vsMovimiento.Rows > 1 Then
            vsListado.NewPage
            vsListado.Paragraph = "": vsListado.Paragraph = "Movimientos sin Asientos"
            vsMovimiento.ExtendLastCol = False: vsListado.RenderControl = vsMovimiento.hwnd: vsMovimiento.ExtendLastCol = True
        End If
        
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
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionConfigurar()
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
End Sub

Private Sub Ayuda(strTexto As String)
    Status.Panels(4).Text = strTexto
    Status.Refresh
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Function ValidoCampos() As Boolean

    ValidoCampos = False
    
    If Not IsDate(tFecha.Text) And IsDate(tFHasta.Text) Then
        MsgBox "Ingrese la fecha desde.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    If IsDate(tFecha.Text) And Not IsDate(tFHasta.Text) Then
        If Trim(tFHasta.Text) = "" Then
            tFHasta.Text = tFecha.Text
        Else
            MsgBox "La fecha hasta no es correcta.", vbExclamation, "ATENCIÓN"
            Foco tFHasta: Exit Function
        End If
    End If
    If IsDate(tFecha.Text) And IsDate(tFHasta.Text) Then
        If CDate(tFecha.Text) > CDate(tFHasta.Text) Then
            MsgBox "Los rangos de fecha no son correctos.", vbExclamation, "ATENCIÓN"
            Foco tFecha: Exit Function
        End If
    End If
        
    ValidoCampos = True
    
End Function

Private Sub EliminoAsientosDobles()

    On Error GoTo errEliminar
    With vsConsulta
        If opNetear.Value = vbChecked Then
        For I = 1 To .Rows - 1
            If I + 1 <= .Rows - 1 Then
                If .Cell(flexcpText, I, 0) = .Cell(flexcpText, I + 1, 0) And (.Cell(flexcpText, I, 2) <> "" Or .Cell(flexcpText, I, 4) <> "") Then
                    If .Cell(flexcpText, I, 1) = .Cell(flexcpText, I + 1, 1) Then
                        
                        If .Cell(flexcpText, I, 2) <> "" Then
                            Select Case .Cell(flexcpValue, I, 2)
                                Case Is > .Cell(flexcpValue, I + 1, 4)
                                            .Cell(flexcpText, I, 2) = Format(.Cell(flexcpValue, I, 2) - .Cell(flexcpValue, I + 1, 4), FormatoMonedaP)
                                            If .Cell(flexcpText, I, 3) <> "" Then .Cell(flexcpText, I, 3) = Format(.Cell(flexcpValue, I, 3) - .Cell(flexcpValue, I + 1, 5), FormatoMonedaP)
                                            .RemoveItem I + 1
                                
                                Case Is < .Cell(flexcpValue, I + 1, 4)
                                            .Cell(flexcpText, I + 1, 4) = Format(.Cell(flexcpValue, I + 1, 4) - .Cell(flexcpValue, I, 2), FormatoMonedaP)
                                            If .Cell(flexcpText, I + 1, 5) <> "" Then .Cell(flexcpText, I + 1, 5) = Format(.Cell(flexcpValue, I + 1, 5) - .Cell(flexcpValue, I, 3), FormatoMonedaP)
                                            .RemoveItem I
                                
                                Case Is = .Cell(flexcpValue, I + 1, 4): .RemoveItem I: .RemoveItem I
                            End Select
                        Else
                            Select Case .Cell(flexcpValue, I, 4)
                                Case Is > .Cell(flexcpValue, I + 1, 2)
                                            .Cell(flexcpText, I, 4) = Format(.Cell(flexcpValue, I, 4) - .Cell(flexcpValue, I + 1, 2), FormatoMonedaP)
                                            If .Cell(flexcpText, I, 5) <> "" Then .Cell(flexcpText, I, 5) = Format(.Cell(flexcpValue, I, 5) - .Cell(flexcpValue, I + 1, 3), FormatoMonedaP)
                                            .RemoveItem I + 1
                                
                                Case Is < .Cell(flexcpValue, I + 1, 2)
                                            .Cell(flexcpText, I + 1, 2) = Format(.Cell(flexcpValue, I + 1, 2) - .Cell(flexcpValue, I, 4), FormatoMonedaP)
                                            If .Cell(flexcpText, I + 1, 3) <> "" Then .Cell(flexcpText, I + 1, 3) = Format(.Cell(flexcpValue, I + 1, 3) - .Cell(flexcpValue, I, 5), FormatoMonedaP)
                                            .RemoveItem I
                                
                                Case Is = .Cell(flexcpValue, I + 1, 2): .RemoveItem I: .RemoveItem I
                            End Select
                        End If
                    End If
                End If
            End If
        Next
        
        For I = 1 To .Rows - 1
            If I + 1 <= .Rows - 1 Then
                If .Cell(flexcpBackColor, I, 0) = Colores.Inactivo And .Cell(flexcpBackColor, I + 1, 0) = Colores.Inactivo Then
                    .RemoveItem I
                End If
            End If
        Next
        If .Cell(flexcpBackColor, .Rows - 1, 0) = Colores.Inactivo Then .RemoveItem .Rows - 1
        End If
        
        'Vuelvo a Ordenar para que me quede DEBE/HABER
        For I = 1 To .Rows - 1
            If .Cell(flexcpBackColor, I, 0) = Colores.Inactivo Then
                .Cell(flexcpText, I, 0) = .Cell(flexcpText, I, 0) & "2"
            Else
                If .Cell(flexcpText, I, 2) <> "" Then .Cell(flexcpText, I, 0) = .Cell(flexcpText, I, 0) & "2" Else .Cell(flexcpText, I, 0) = .Cell(flexcpText, I, 0) & "1"
            End If
        Next
        If vsConsulta.Rows > 1 Then
            .Select 1, 0, 1, 1
            .Sort = flexSortGenericDescending
        End If
    End With
    Exit Sub

errEliminar:
    clsGeneral.OcurrioError "Ocurrió un error al eliminar los asientos dobles.", Err.Description
End Sub

Private Function meVentasContado(mIDMoneda As Long, mNameMoneda As String)

Dim rsSuc As rdoResultset
Dim rsData As rdoResultset
Dim mSubRubro As String
    
    Dim mSucursal As Long
    Dim mIDDisponibilidad As Long
    Dim mTotal As Currency, mIva As Currency, mCofis As Currency
    
    mTotal = 0: mIva = 0: mCofis = 0
    cons = "Select * from Sucursal Where SucDisponibilidadME Is Not Null"
    Set rsSuc = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsSuc.EOF
        mSucursal = rsSuc!SucCodigo
        
        Ayuda "Cargando Ventas Contado " & mNameMoneda & " (" & Trim(rsSuc!SucAbreviacion) & ")..."
        'ASIENTO DE VENTAS CONTADO PESOS--------------------------------------------------------------------------------------------------------
        cons = "Select Sum(DocTotal) Total, Sum(DocIva) Iva, Sum(DocCofis) Cofis from Documento " _
                & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
                & " And DocTipo = " & TipoDocumento.Contado _
                & " And DocAnulado = 0" _
                & " And DocMoneda = " & mIDMoneda _
                & " And DocSucursal = " & mSucursal
                
        Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
        If Not rsAux.EOF Then
            If Not IsNull(rsAux!Total) Then
            If rsAux!Total <> 0 Then
            
                If Not IsNull(rsAux!Cofis) Then mCofis = mCofis + rsAux!Cofis
                mTotal = mTotal + Format(rsAux!Total, FormatoMonedaP)
                mIva = mIva + Format(rsAux!Iva, FormatoMonedaP)
                
                mIDDisponibilidad = dis_DisponibilidadPara(mSucursal, mIDMoneda)
                
                'Saco el subrubro de la disponibilidad para hacer el asiento final --------------------------------------------
                cons = "Select * from Disponibilidad, Subrubro " _
                       & " Where DisID = " & mIDDisponibilidad _
                       & " And DisIDSubRubro = SRuId"
                Set rsData = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                If Not rsData.EOF Then
                    mSubRubro = Format(rsData!SRuCodigo, "000000000") & " " & Trim(rsData!SRuNombre)
                Else
                    mSubRubro = "CAJA " & mNameMoneda
                End If
                rsData.Close
                '------------------------------------------------------------------------------------------------------------------------------------
                
                'Si el Subrubro está en el array --> Acumulo los datos, si no lo agrego
                Dim bOK As Boolean: bOK = False
                For idX = LBound(arrAsientos) To UBound(arrAsientos)
                    If arrAsientos(idX).NombreRubro = mSubRubro Then
                        arrAsientos(idX).Importe = arrAsientos(idX).Importe + Format(rsAux!Total, FormatoMonedaP)
                        bOK = True
                        Exit For
                    End If
                Next
                If Not bOK Then
                    ReDim Preserve arrAsientos(UBound(arrAsientos) + 1)
                    With arrAsientos(UBound(arrAsientos))
                        .NombreRubro = mSubRubro
                        .Importe = Format(rsAux!Total, FormatoMonedaP)
                    End With
                End If
                
            End If
            
            End If
        End If
        rsAux.Close
            
        rsSuc.MoveNext
    Loop
    rsSuc.Close
    
    If mTotal <> 0 Then     'ASIENTO DE VENTAS CONTADO EN M/E
        With vsConsulta
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = "Ventas Contado (" & mNameMoneda & ")"
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVContado)
            .Cell(flexcpText, .Rows - 1, 2) = Format(mTotal, FormatoMonedaP)
        
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentas)
            .Cell(flexcpText, .Rows - 1, 4) = Format(mTotal - mIva - mCofis, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubRubroCofis)
            .Cell(flexcpText, .Rows - 1, 4) = Format(mCofis, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIVA)
            .Cell(flexcpText, .Rows - 1, 4) = Format(mIva, FormatoMonedaP)
        End With
    End If
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------
    
End Function
