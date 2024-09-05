VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmListado 
   Caption         =   "Asientos Varios"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   450
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
   StartUpPosition =   3  'Windows Default
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
            Object.Width           =   12753
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

Private RsAux As rdoResultset, Rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean

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
        .FormatString = "Rubro Orden|<Rubro|>Debe $|Rubro|>Haber $|"
            
        .WordWrap = False
        .ColWidth(0) = 0
        .ColWidth(1) = 3500: .ColWidth(2) = 1300: .ColWidth(3) = 3500: .ColWidth(4) = 1300
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
    vsConsulta.Height = (Altura / 5) * 4
    
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
    Cons = " Select Count(Distinct(MDiTipo)) from MovimientoDisponibilidad " _
            & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDiIDCompra Is Null "
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux(0) = 0 Then
            MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
            RsAux.Close: Screen.MousePointer = 0: vsConsulta.Rows = 1: Exit Sub
        End If
        pbProgreso.Max = RsAux(0)
    End If
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------------------
     
    
    Cons = " Select TMDNombre, MDiTipo, DH = 'Haber', Importe = Sum(MDrImportePesos)" _
            & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, TipoMovDisponibilidad " _
            & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDiID = MDRIDMovimiento " _
            & " And MDiIDCompra Is Null " _
            & " And MDiTipo = TMDCodigo " _
            & " And MDRHaber <> Null " _
            & " Group by TMDNombre, MDiTipo"
    
    Cons = Cons & " UNION ALL "
    
    Cons = Cons & " Select TMDNombre, MDiTipo, DH = 'Debe', Importe = Sum(MDrImportePesos)" _
                        & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, TipoMovDisponibilidad " _
                        & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
                        & " And MDiID = MDRIDMovimiento " _
                        & " And MDiIDCompra Is Null " _
                        & " And MDiTipo = TMDCodigo " _
                        & " And MDRDebe <> Null " _
                        & " Group by TMDNombre, MDiTipo"
            
    Cons = Cons & " Order by TMDNombre "
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        RsAux.Close: Screen.MousePointer = 0: Exit Sub
    End If
    
    vsConsulta.Redraw = False
    aTipoM = 0
    pbProgreso.Value = 0
    Do While Not RsAux.EOF
        aSR = 0
        With vsConsulta
            
            Select Case RsAux!MDiTipo
                Case paMCNotaCredito, paMCAnulacion: aSR = paSubrubroDeudoresPorVenta
                Case paMCChequeDiferido: aSR = paSubrubroCDAlCobro
                Case paMCVtaTelefonica: aSR = paSubrubroVtasTelACobrar
                Case paMCLiquidacionCamionero: aSR = paSubrubroCobranzaVtasTel
                Case paMCSenias: aSR = paSubrubroSeniasRecibidas
                
                Case Else
                    'Consulto en el movimiento si es del tipo tranferencia
                    Cons = "Select * from TipoMovDisponibilidad Where TMDCodigo = " & RsAux!MDiTipo
                    Set Rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If Not Rs1.EOF Then
                        If Not IsNull(Rs1!TMDTransferencia) Then If Rs1!TMDTransferencia = 1 Then aSR = -1
                    End If
                    Rs1.Close
            End Select
            
            If aSR <> 0 Then
            
                If aTipoM <> RsAux!MDiTipo Then
                    pbProgreso.Value = pbProgreso.Value + 1
                    aTipoM = RsAux!MDiTipo
                    
                    .AddItem aTipoM   'Nombre del Movimiento
                    .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!TMDNombre)
                    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
                    
                    CargoConceptos aTipoM   'Hay que cargar contra que conceptos van
                    
                End If
                
                If aSR <> -1 Then       'NO ES TRANSFERENCIA
                    .AddItem ""
                    aTexto = RetornoConstanteSubrubro(aSR)
                    .Cell(flexcpText, .Rows - 1, 0) = aTipoM
                    
                    Select Case LCase(RsAux!DH)     'Van al reves por los tipos de movimientos
                        Case "debe"
                            .Cell(flexcpText, .Rows - 1, 3) = aTexto
                            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!Importe, FormatoMonedaP)
                        
                        Case "haber"
                            .Cell(flexcpText, .Rows - 1, 1) = aTexto
                            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!Importe, FormatoMonedaP)
                    End Select
                End If
            Else
                If aTipoM <> RsAux!MDiTipo Then
                    pbProgreso.Value = pbProgreso.Value + 1
                    aTipoM = RsAux!MDiTipo
                    If aTipoM <> paMCIngresosOperativos Then CargoListaMovimiento aTipoM
                End If
            End If
            
            RsAux.MoveNext
        End With
    Loop
    RsAux.Close
    
    With vsConsulta
        If vsConsulta.Rows > 1 Then
            .Select 1, 0, 1, 1
            .Sort = flexSortGenericDescending
        End If
    End With
    
    CargoAsientosDeVentas
    
    If vsMovimiento.Rows > 1 Then
        With vsMovimiento
            .SubtotalPosition = flexSTBelow
            .Subtotal flexSTSum, -1, 2, , Colores.Inactivo, , , "Total"
            .Subtotal flexSTSum, -1, 3
        End With
        vsConsulta.Height = (vsListado.Height / 5) * 4
    Else
        vsConsulta.Height = vsListado.Height
    End If
    
    vsConsulta.Redraw = True
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True: Screen.MousePointer = 0
End Sub

Private Sub CargoConceptos(aIDTipoM As Long, Optional Transferencia As Boolean = False)

    Cons = " Select SRuCodigo, SRuNombre, DH = 'Haber', Importe = Sum(MDrImportePesos)"
    Cons = Cons & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, Disponibilidad, SubRubro " _
                        & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
                        & " And MDiID = MDRIDMovimiento " _
                        & " And MDiIDCompra Is Null " _
                        & " And MDiTipo = " & aIDTipoM _
                        & " And MDRIDDisponibilidad = DisID " _
                        & " And DisIDSubrubro = SRuID" _
                        & " And MDRHaber <> Null " _
                        & " Group by SRuCodigo, SRuNombre"
    
    Cons = Cons & " UNION ALL "
    
    Cons = Cons & " Select SRuCodigo, SRuNombre, DH = 'Debe', Importe = Sum(MDrImportePesos)"
    Cons = Cons & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, Disponibilidad, SubRubro " _
                        & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
                        & " And MDiID = MDRIDMovimiento " _
                        & " And MDiIDCompra Is Null " _
                        & " And MDiTipo = " & aIDTipoM _
                        & " And MDRIDDisponibilidad = DisID " _
                        & " And DisIDSubrubro = SRuID" _
                        & " And MDRDebe <> Null " _
                        & " Group by SRuCodigo, SRuNombre"
            
    Set Rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not Rs1.EOF
        
        With vsConsulta
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = aIDTipoM
             aTexto = Format(Rs1!SRuCodigo, "000000000") & " " & Trim(Rs1!SRuNombre)
            
            Select Case LCase(Rs1!DH)
                Case "debe"
                    .Cell(flexcpText, .Rows - 1, 1) = aTexto
                    .Cell(flexcpText, .Rows - 1, 2) = Format(Rs1!Importe, FormatoMonedaP)
                
                Case "haber"
                    .Cell(flexcpText, .Rows - 1, 3) = aTexto
                    .Cell(flexcpText, .Rows - 1, 4) = Format(Rs1!Importe, FormatoMonedaP)
            End Select
        End With
        
        Rs1.MoveNext
    Loop
    Rs1.Close
    
End Sub

Private Sub CargoAsientosDeVentas()

    Dim aSRCaja As String
    'Saco el subrubro de la disponibilidad (Caja)  para hacer el otro asiento--------------------------------------------
    Cons = "Select * from Disponibilidad, Subrubro " _
           & " Where DisID = " & paDisponibilidad _
           & " And DisIDSubRubro = SRuId"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        aSRCaja = Format(RsAux!SRuCodigo, "000000000") & " " & Trim(RsAux!SRuNombre)
    Else
        aSRCaja = "CAJA"
    End If
    RsAux.Close
    '------------------------------------------------------------------------------------------------------------------------------------
    
    'ASIENTO DE VENTAS CONTADO--------------------------------------------------------------------------------------------------------
    Cons = "Select Sum(DocTotal) Total, Sum(DocIva) IVA from Documento " _
            & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
            & " And DocTipo = " & TipoDocumento.Contado _
            & " And DocAnulado = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        With vsConsulta
            .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Ventas Contado": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVContado)
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!Total, FormatoMonedaP)
        
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 3) = RetornoConstanteSubrubro(paSubrubroVentas)
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!Total - RsAux!Iva, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 3) = RetornoConstanteSubrubro(paSubrubroIVA)
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!Iva, FormatoMonedaP)
                    
            .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Ventas Contado": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = aSRCaja
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!Total, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 3) = RetornoConstanteSubrubro(paSubrubroVContado)
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!Total, FormatoMonedaP)
        End With
    End If
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'ASIENTO DE VENTAS CREDITO--------------------------------------------------------------------------------------------------------
    Cons = "Select Sum(DocTotal) Total, Sum(DocIva) IVA from Documento " _
            & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
            & " And DocTipo = " & TipoDocumento.Credito _
            & " And DocAnulado = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        With vsConsulta
            .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Ventas Crédito": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroDeudoresPorVenta)
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!Total, FormatoMonedaP)
        
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 3) = RetornoConstanteSubrubro(paSubrubroVentas)
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!Total - RsAux!Iva, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 3) = RetornoConstanteSubrubro(paSubrubroIVA)
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!Iva, FormatoMonedaP)
        End With
    End If
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'ASIENTO COBRANZA DE MOROSIDADES Y CTAS--------------------------------------------------------------------------------------------------------
    Dim aMoraTotal As Currency
    Cons = "Select Sum(DPaMora) Total from Documento, DocumentoPago " _
            & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
            & " And DocTipo = " & TipoDocumento.ReciboDePago _
            & " And DocAnulado = 0" _
            & " And DocCodigo = DPaDocQSalda and DPaMora > 0"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        If Not IsNull(RsAux!Total) Then aMoraTotal = RsAux!Total
    End If
    RsAux.Close
    
    Cons = "Select Sum(DocTotal) Total, Sum(DocIva) IVA from Documento " _
            & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
            & " And DocTipo = " & TipoDocumento.ReciboDePago _
            & " And DocAnulado = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        With vsConsulta
            'MORAS
            .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Cobranza de Morosidades": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = aSRCaja
            .Cell(flexcpText, .Rows - 1, 2) = Format(aMoraTotal, FormatoMonedaP)
        
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 3) = RetornoConstanteSubrubro(paSubrubroIngresosVarios)
            .Cell(flexcpText, .Rows - 1, 4) = Format(aMoraTotal - RsAux!Iva, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 3) = RetornoConstanteSubrubro(paSubrubroIVA)
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!Iva, FormatoMonedaP)
            
            'CUOTAS
            .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Cobranza de Cuotas": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = aSRCaja
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!Total - aMoraTotal, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 3) = RetornoConstanteSubrubro(paSubrubroDeudoresPorVenta)
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!Total - aMoraTotal, FormatoMonedaP)
        End With
    End If
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------
    'ASIENTO DE NOTAS DEV Y ESPECIALES-------------------------------------------------------------------------------------------------------------
    Cons = "Select Sum(DocTotal) Total, Sum(DocIva) IVA from Documento " _
            & " Where DocFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & " 00:00' And '" & Format(tFHasta.Text, "mm/dd/yyyy") & " 23:59'" _
            & " And DocTipo IN (" & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")" _
            & " And DocAnulado = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        With vsConsulta
            .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Notas de Devolución y Notas Especiales": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVentas)
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!Total - RsAux!Iva, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroIVA)
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!Iva, FormatoMonedaP)
                    
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 3) = RetornoConstanteSubrubro(paSubrubroVContado)
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!Total, FormatoMonedaP)
            
            .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = "Notas de Devolución y Notas Especiales": .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = RetornoConstanteSubrubro(paSubrubroVContado)
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!Total, FormatoMonedaP)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 3) = aSRCaja
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!Total, FormatoMonedaP)

        End With
    End If
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------


End Sub


Private Sub CargoListaMovimiento(aIDTipoM As Long)

    Cons = " Select TMDNombre, MDiComentario, DH = 'Haber', Importe = Sum(MDrImportePesos)" _
            & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, TipoMovDisponibilidad " _
            & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDiTipo = " & aIDTipoM _
            & " And MDiID = MDRIDMovimiento " _
            & " And MDiIDCompra Is Null " _
            & " And MDiTipo = TMDCodigo " _
            & " And MDRHaber <> Null " _
            & " Group by TMDNombre, MDiComentario"
    
    Cons = Cons & " UNION ALL "
    
    Cons = Cons & " Select TMDNombre, MDiComentario, DH = 'Debe', Importe = Sum(MDrImportePesos)" _
                        & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, TipoMovDisponibilidad " _
                        & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
                        & " And MDiTipo = " & aIDTipoM _
                        & " And MDiID = MDRIDMovimiento " _
                        & " And MDiIDCompra Is Null " _
                        & " And MDiTipo = TMDCodigo " _
                        & " And MDRDebe <> Null " _
                        & " Group by TMDNombre, MDiComentario"
    
    Set Rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not Rs1.EOF
        
        With vsMovimiento
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Rs1!TMDNombre
            If Not IsNull(Rs1!MDiComentario) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(Rs1!MDiComentario)
            
            Select Case LCase(Rs1!DH)
                Case "debe": .Cell(flexcpText, .Rows - 1, 2) = Format(Rs1!Importe, FormatoMonedaP)
                Case "haber": .Cell(flexcpText, .Rows - 1, 3) = Format(Rs1!Importe, FormatoMonedaP)
            End Select
        End With
        
        Rs1.MoveNext
    Loop
    Rs1.Close
    
End Sub

Private Sub tFHasta_GotFocus()
    With tFHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tFHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
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
        vsListado.filename = "Listado Asientos Varios"
            
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hWnd: vsConsulta.ExtendLastCol = True
        
        If vsMovimiento.Rows > 1 Then
            vsListado.NewPage
            vsListado.Paragraph = "": vsListado.Paragraph = "Movimientos sin Asientos"
            vsMovimiento.ExtendLastCol = False: vsListado.RenderControl = vsMovimiento.hWnd: vsMovimiento.ExtendLastCol = True
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
