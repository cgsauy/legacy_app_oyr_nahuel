VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmListado 
   Caption         =   "Asientos de Gastos a Crédito"
   ClientHeight    =   7590
   ClientLeft      =   1815
   ClientTop       =   2055
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   4680
      TabIndex        =   19
      Top             =   960
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
      TabIndex        =   22
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
      Zoom            =   100
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   9555
      TabIndex        =   23
      Top             =   6720
      Width           =   9615
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   15
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   14
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
         TabIndex        =   17
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
         TabIndex        =   18
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
         TabIndex        =   7
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6000
         TabIndex        =   24
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
      TabIndex        =   21
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
            Object.Width           =   12753
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
      TabIndex        =   20
      Top             =   0
      Width           =   11175
      Begin VB.CheckBox chComentarios 
         Caption         =   "&Ver Comentarios"
         Height          =   195
         Left            =   6720
         TabIndex        =   6
         Top             =   300
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
      Begin AACombo99.AACombo cTipoListado 
         Height          =   315
         Left            =   4740
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
      Begin VB.Label Label1 
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   4260
         TabIndex        =   4
         Top             =   240
         Width           =   615
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
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RsAux As rdoResultset, Rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean
Dim txtAcreedoresVarios As String

Private Sub AccionLimpiar()
    tFecha.Text = "": tFHasta.Text = ""
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

Private Sub chComentarios_Click()
    If chComentarios.Value = vbChecked Then
        vsConsulta.ColWidth(6) = 1600
        vsConsulta.Cell(flexcpText, 0, 6) = "Comentarios"
    Else
        vsConsulta.ColWidth(6) = 0
        vsConsulta.Cell(flexcpText, 0, 6) = ""
    End If
End Sub

Private Sub cTipoListado_GotFocus()
    With cTipoListado: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub cTipoListado_KeyPress(KeyAscii As Integer)
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

Private Sub Label2_Click()
    Foco tFHasta
End Sub

Private Sub tFecha_GotFocus()
    With tFecha: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tFHasta
End Sub
Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then
        tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
        If Not IsDate(tFHasta.Text) Then tFHasta.Text = tFecha.Text
    End If
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
    bCargarImpresion = True
    
    CargoConstantesSubrubros
    Cons = "Select * from Subrubro Where SRuID = " & paSubrubroAcreedoresVarios
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then txtAcreedoresVarios = Format(RsAux!SRuCodigo, "000000000") & " " & Trim(RsAux!SRuNombre)
    RsAux.Close
    
    cTipoListado.AddItem "Ingresos": cTipoListado.ItemData(cTipoListado.NewIndex) = 0
    cTipoListado.AddItem "Egresos": cTipoListado.ItemData(cTipoListado.NewIndex) = 1
    
    vsListado.Orientation = orPortrait
    vsListado.MarginBottom = 750: vsListado.MarginTop = 750
    
    pbProgreso.Value = 0
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    vsListado.MarginRight = 300
    With vsConsulta
        .OutlineBar = flexOutlineBarSimple
        '.OutlineBar = flexOutlineBarComplete 'flexOutlineBarNone
        .OutlineCol = 0
        .MultiTotals = True
        .ColSel = 0
        .ColSort(0) = flexSortStringAscending
        .Sort = flexSortUseColSort
        
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<Rubro|<SubRubro|>Importe $|>I.V.A. $|>Total $|>Total M/E|<|"
            
        .WordWrap = False
        .ColWidth(0) = 2300: .ColWidth(1) = 2150: .ColWidth(2) = 1550: .ColWidth(3) = 1400: .ColWidth(4) = 1550: .ColWidth(5) = 1400
        .ColWidth(6) = 0
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
    
    picBotones.Width = vsListado.Width
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 150
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = vsListado.Height
    vsConsulta.Left = vsListado.Left
    
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
Dim IdCompra As Long
Dim IvaCompra As Currency, IvaGastos As Currency
Dim RsRub As rdoResultset

    If Not ValidoCampos Then Exit Sub
    
    On Error GoTo errConsultar
    Screen.MousePointer = 11

    bCargarImpresion = True
    
    vsConsulta.Rows = 1: vsConsulta.Refresh
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    Cons = " Select Count(*) from Compra, GastoSubRubro " _
            & " Where ComFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And ComCodigo = GSrIDCompra"
    
    Select Case cTipoListado.ItemData(cTipoListado.ListIndex)
        Case 0: Cons = Cons & " And ComTipoDocumento = " & TipoDocumento.CompraNotaCredito        '0- Ingresos
        Case 1: Cons = Cons & " And ComTipoDocumento = " & TipoDocumento.CompraCredito              '1- Egresos
    End Select
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux(0) = 0 Then
            MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
            RsAux.Close: Screen.MousePointer = 0: Exit Sub
        End If
        pbProgreso.Max = RsAux(0)
    End If
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------------------

    Cons = "Select * From Compra, GastoSubRubro, SubRubro, Rubro" _
            & " Where ComFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And ComCodigo = GSrIDCompra" _
            & " And GSrIDSubRubro = SRuID " _
            & " And SRuRubro = RubID"
    
    Select Case cTipoListado.ItemData(cTipoListado.ListIndex)
        Case 0: Cons = Cons & " And ComTipoDocumento = " & TipoDocumento.CompraNotaCredito        '0- Ingresos
        Case 1: Cons = Cons & " And ComTipoDocumento = " & TipoDocumento.CompraCredito              '1- Egresos
    End Select
    
    Cons = Cons & " Order by ComCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        RsAux.Close: Screen.MousePointer = 0: InicializoGrillas: Exit Sub
    End If
    
    IdCompra = 0
    IvaCompra = 0: IvaGastos = 0
    
    vsConsulta.Rows = 1: vsConsulta.Redraw = False
    Do While Not RsAux.EOF
        pbProgreso.Value = pbProgreso.Value + 1
        With vsConsulta
            .AddItem ""
            'En el data del item 0 pongo si el rubro expande o no
            If RsAux!RubExpandir Then .Cell(flexcpData, .Rows - 1, 0) = 1 Else .Cell(flexcpData, .Rows - 1, 0) = 0
            
            .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!RubCodigo, "000000000") & " " & Trim(RsAux!RubNombre)
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!SRuCodigo, "000000000") & " " & Trim(RsAux!SRuNombre)
            
            If Not IsNull(RsAux!ComComentario) Then .Cell(flexcpText, .Rows - 1, 6) = Trim(RsAux!ComComentario)
            .Cell(flexcpText, .Rows - 1, 7) = " "
            
            If IdCompra <> RsAux!ComCodigo Then
                
                If Not IsNull(RsAux!ComIVA) Then
                    If RsAux!ComMoneda = paMonedaPesos Then
                         .Cell(flexcpText, .Rows - 1, 3) = Format(Abs(RsAux!ComIVA), FormatoMonedaP)
                         .Cell(flexcpText, .Rows - 1, 2) = Format(Abs(RsAux!ComImporte), FormatoMonedaP)
                    Else
                        .Cell(flexcpText, .Rows - 1, 3) = Format(Abs(RsAux!ComIVA) * RsAux!ComTC, FormatoMonedaP)
                        .Cell(flexcpText, .Rows - 1, 2) = Format(Abs(RsAux!ComImporte) * RsAux!ComTC, FormatoMonedaP)
                        .Cell(flexcpText, .Rows - 1, 5) = Format(Abs(RsAux!ComIVA), FormatoMonedaP)
                    End If
                    
                    If paSubrubroCompraMercaderia = RsAux!SRuID Then
                        IvaCompra = IvaCompra + .Cell(flexcpText, .Rows - 1, 3)
                    Else
                        IvaGastos = IvaGastos + .Cell(flexcpText, .Rows - 1, 3)
                    End If
                End If
            End If
            If Not IsNull(RsAux!GSrIDSubRubro) Then
                If RsAux!ComMoneda = paMonedaPesos Then
                    .Cell(flexcpText, .Rows - 1, 2) = Format(Abs(RsAux!GSrImporte), FormatoMonedaP)
                Else
                    .Cell(flexcpText, .Rows - 1, 2) = Format(Abs(RsAux!GSrImporte) * RsAux!ComTC, FormatoMonedaP)
                    .Cell(flexcpText, .Rows - 1, 5) = Format(.Cell(flexcpValue, .Rows - 1, 5) + Abs(RsAux!GSrImporte), FormatoMonedaP)
                End If
            End If
            .Cell(flexcpText, .Rows - 1, 4) = Format(.Cell(flexcpValue, .Rows - 1, 2) + .Cell(flexcpValue, .Rows - 1, 3), FormatoMonedaP)
            
            IdCompra = RsAux!ComCodigo
            
            RsAux.MoveNext
        End With
    Loop
    RsAux.Close
    
    Dim aTotal As Currency
    With vsConsulta
        .Select 1, 0, 1, 1
        .Sort = flexSortGenericAscending
        
        'Totales para la Columna 1
        .Subtotal flexSTSum, 1, 2, , , , False, "%s"
        .Subtotal flexSTSum, 1, 3: .Subtotal flexSTSum, 1, 4: .Subtotal flexSTSum, 1, 5
        .Cell(flexcpForeColor, 1, 0, .Rows - 1, .Cols - 1) = Colores.Azul
        
        'Totales para la Columna 0
        .Subtotal flexSTSum, 0, 2, , , , , "%s"
        .Subtotal flexSTSum, 0, 3: .Subtotal flexSTSum, 0, 4: .Subtotal flexSTSum, 0, 5

        'Total de todos los Renglones
        .Subtotal flexSTSum, -1, 2, , Colores.Obligatorio, &H80&, True, "Total"
        .Subtotal flexSTSum, -1, 3: .Subtotal flexSTSum, -1, 4: .Subtotal flexSTSum, -1, 5
        aTotal = .Cell(flexcpValue, .Rows - 1, 4)
        
        If IvaCompra <> 0 Then
            .AddItem "I.V.A. Compras"
            .Cell(flexcpText, .Rows - 1, 3) = Format(IvaCompra, FormatoMonedaP)
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
        End If
        If IvaGastos <> 0 Then
            .AddItem "I.V.A. Gastos"
            .Cell(flexcpText, .Rows - 1, 3) = Format(IvaGastos, FormatoMonedaP)
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
        End If
    End With
    
    AgrupoCamposEnGrilla
        
    'Cargo contra que concepto Van---------------------------------------------------------------------------
    With vsConsulta
        .AddItem "": .AddItem ""
        Select Case cTipoListado.ItemData(cTipoListado.ListIndex)
            Case 0: .Cell(flexcpText, .Rows - 1, 1) = "Concepto al DEBE"        '0- Ingresos
            Case 1: .Cell(flexcpText, .Rows - 1, 1) = "Concepto al HABER"      '1- Egresos
        End Select
        .Cell(flexcpBackColor, .Rows - 1, 1, , .Cols - 1) = Colores.Azul: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco: .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
    
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = txtAcreedoresVarios
        .Cell(flexcpText, .Rows - 1, 4) = Format(aTotal, FormatoMonedaP)
    End With
    '------------------------------------------------------------------------------------------------------------------------------------------------------
    
    vsConsulta.Redraw = True
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True: Screen.MousePointer = 0
End Sub

Private Sub tFHasta_GotFocus()
    With tFHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tFHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cTipoListado
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
        
        vsListado.PaperSize = 1
        If chComentarios.Value = vbUnchecked Then
            vsListado.Orientation = orPortrait
        Else
            vsListado.Orientation = orLandscape
            vsConsulta.ColWidth(6) = 3900
        End If
        
        With vsListado
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        EncabezadoListado vsListado, "Asientos de Gastos a Crédito (" & Trim(cTipoListado.Text) & ") - Del " & Trim(tFecha.Text) & " al " & Trim(tFHasta.Text), False
        vsListado.filename = "Listado Asientos de Gastos a Credito"
            
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hWnd: vsConsulta.ExtendLastCol = True
        
        vsListado.EndDoc
        'bCargarImpresion = False
        If chComentarios.Value = vbChecked Then vsConsulta.ColWidth(6) = 1400
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

Private Sub AgrupoCamposEnGrilla()

On Error GoTo ErrACEG
Dim aData As Boolean

    With vsConsulta
        For I = 1 To .Rows - 1
            If .IsSubtotal(I) Then
                If aData Then   'Hay que expandir la rama de los subrubros
                    Select Case .RowOutlineLevel(I)
                        Case 0: .IsCollapsed(I) = flexOutlineExpanded
                        Case 1: .IsCollapsed(I) = flexOutlineCollapsed
                    End Select
                Else
                    If .RowOutlineLevel(I) <> -1 Then .IsCollapsed(I) = flexOutlineCollapsed
                End If
            Else
                If .Cell(flexcpData, I, 0) = 1 Then aData = True Else aData = False
            End If
            
        Next I
    End With
    
    Exit Sub
ErrACEG:
End Sub

Private Sub vsConsulta_Collapsed()
    
    If vsConsulta.RowOutlineLevel(vsConsulta.Row) <> 1 Then Exit Sub
    
    With vsConsulta
        If .IsCollapsed(.Row) Then
            .Cell(flexcpForeColor, .Row, 0, , .Cols - 1) = Colores.Azul
        Else
            .Cell(flexcpForeColor, .Row, 0, , .Cols - 1) = vbBlack
        End If
    End With
    
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
        
    If cTipoListado.ListIndex = -1 Then
        MsgBox "Debe seleccioar el tipo de movimientos a listr (Ingresos/Egresos).", vbExclamation, "ATENCIÓN"
        Foco cTipoListado: Exit Function
    End If
    ValidoCampos = True
    
End Function
