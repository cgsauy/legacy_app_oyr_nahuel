VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACOMBO.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form CoComprobante 
   Caption         =   "Consulta de Comprobantes"
   ClientHeight    =   8595
   ClientLeft      =   2070
   ClientTop       =   1545
   ClientWidth     =   11100
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000C0&
   Icon            =   "CoComprobante.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11100
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   6615
      Left            =   240
      TabIndex        =   15
      Top             =   1200
      Width           =   9855
      _Version        =   196608
      _ExtentX        =   17383
      _ExtentY        =   11668
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
      Zoom            =   70
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   7275
      TabIndex        =   16
      Top             =   7920
      Width           =   7335
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "CoComprobante.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "CoComprobante.frx":067C
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "CoComprobante.frx":097E
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "CoComprobante.frx":0CC0
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5640
         Picture         =   "CoComprobante.frx":0FC2
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Picture         =   "CoComprobante.frx":10C4
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "CoComprobante.frx":148A
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "CoComprobante.frx":158C
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "CoComprobante.frx":17C6
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "CoComprobante.frx":18B0
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "CoComprobante.frx":199A
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "CoComprobante.frx":1E14
         Height          =   310
         Left            =   4440
         Picture         =   "CoComprobante.frx":1F5E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   310
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4455
      Left            =   960
      TabIndex        =   12
      Top             =   1440
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7858
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
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   12
      FixedRows       =   0
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
      TabIndex        =   14
      Top             =   8340
      Width           =   11100
      _ExtentX        =   19579
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
            Object.Width           =   11377
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
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   9615
      Begin VB.OptionButton OpComprobante 
         Caption         =   "Entradas y &Salidas"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Tag             =   "Entradas y Salidas"
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chDetalle 
         Caption         =   "&Detalles del Gasto"
         Height          =   255
         Left            =   5040
         TabIndex        =   9
         Top             =   700
         Width           =   1695
      End
      Begin VB.TextBox tMes 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7620
         MaxLength       =   9
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton OpComprobante 
         Caption         =   "&Boletas y Recibos"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Tag             =   "Boletas y Recibos"
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton OpComprobante 
         Caption         =   "&Facturas y Notas"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Tag             =   "Facturas y Notas"
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
      Begin AACombo99.AACombo cProveedor 
         Height          =   315
         Left            =   2940
         TabIndex        =   4
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
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
      Begin AACombo99.AACombo cRubro 
         Height          =   315
         Left            =   2940
         TabIndex        =   8
         Top             =   660
         Width           =   2055
         _ExtentX        =   3625
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
      Begin AACombo99.AACombo cSubRubro 
         Height          =   315
         Left            =   7620
         TabIndex        =   11
         Top             =   660
         Width           =   2055
         _ExtentX        =   3625
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
      Begin VB.Label Label3 
         Caption         =   "&Sub Rubro:"
         Height          =   255
         Left            =   6780
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "&Rubro:"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "&Mes:"
         Height          =   255
         Left            =   6780
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "&Empresa:"
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   285
         Width           =   855
      End
   End
End
Attribute VB_Name = "CoComprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RsConsulta As rdoResultset
Dim aFormato As String, aTituloTabla As String, aComentario As String
Dim aTexto As String

Private Sub AccionLimpiar()
    cProveedor.Text = ""
    cRubro.Text = ""
    tMes.Text = ""
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

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
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


Private Sub cSubRubro_GotFocus()
    With cSubRubro
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cSubRubro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub


Private Sub chDetalle_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cSubRubro
End Sub

Private Sub chVista_Click()
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
    Else
        AccionImprimir
        vsListado.ZOrder 0
    End If
End Sub

Private Sub cProveedor_GotFocus()
On Error GoTo ErrProv
    With cProveedor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    If cProveedor.ListCount = 0 Then
        Cons = "Select PClCodigo, PClFantasia From ProveedorCliente Order by PClFantasia"
        CargoCombo Cons, cProveedor
    End If
    Exit Sub
ErrProv:
    clsGeneral.OcurrioError "Ocurrio un error inesperado.", Trim(Err.Description)
End Sub

Private Sub cProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tMes
End Sub
Private Sub cProveedor_LostFocus()
    cProveedor.SelStart = 0
End Sub
Private Sub cRubro_Click()
    cSubRubro.Clear
    If cRubro.ListIndex > -1 Then
        If paRubroImportaciones = cRubro.ItemData(cRubro.ListIndex) Then
            chDetalle.Enabled = True
            chDetalle.Value = 0
        Else
            chDetalle.Enabled = False
            chDetalle.Value = 0
        End If
    End If
End Sub

Private Sub cRubro_Change()
    cSubRubro.Clear
    If cRubro.ListIndex > -1 Then
        If paRubroImportaciones = cRubro.ItemData(cRubro.ListIndex) Then
            chDetalle.Enabled = True
            chDetalle.Value = 0
        Else
            chDetalle.Enabled = False
            chDetalle.Value = 0
        End If
    Else
        chDetalle.Enabled = False
        chDetalle.Value = 0
    End If
End Sub

Private Sub cRubro_GotFocus()
On Error GoTo ErrProv
    With cRubro
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    If cRubro.ListCount = 0 Then
        Screen.MousePointer = 11
        Cons = "Select RubID, RubNombre From Rubro Order by RubNombre"
        CargoCombo Cons, cRubro
        Screen.MousePointer = 0
    End If
    Exit Sub
ErrProv:
    clsGeneral.OcurrioError "Ocurrio un error inesperado.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub cRubro_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        CargoSubRubros
        If chDetalle.Enabled Then chDetalle.SetFocus Else cSubRubro.SetFocus
    End If
End Sub

Private Sub cRubro_LostFocus()
    If cRubro.ListIndex > -1 Then
        If paRubroImportaciones = cRubro.ItemData(cRubro.ListIndex) Then
            chDetalle.Enabled = True
        Else
            chDetalle.Enabled = False
            chDetalle.Value = 0
        End If
    End If
    cRubro.SelStart = 0
End Sub

Private Sub Label1_Click()
    Foco tMes
End Sub

Private Sub Label5_Click()
    Foco cProveedor
End Sub
Private Sub Form_Activate()
    Screen.MousePointer = 0: Me.Refresh
End Sub
Private Sub Form_Load()

    On Error GoTo ErrLoad
    
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    LimpioGrilla
    FechaDelServidor
    gFechaServidor = Format(gFechaServidor, FormatoFP)
    AccionLimpiar
    chDetalle.Enabled = False
    vsConsulta.ZOrder 0
    vsConsulta.WordWrap = False
    vsListado.Zoom = 100
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrio un error al inicializar el formulario.", Trim(Err.Description)
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
    
    fFiltros.Left = 120
    vsListado.Left = fFiltros.Left
    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = fFiltros.Width
    
    picBotones.Width = vsListado.Width
    picBotones.BorderStyle = 0
        
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
    
    If Not VerificoFiltros Then Exit Sub
    LimpioGrilla
    
    On Error GoTo errConsultar
    Screen.MousePointer = 11
    ArmoConsulta
    Set RsConsulta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsConsulta.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        Screen.MousePointer = 0: Exit Sub
    Else
        CargoLista
    End If
    RsConsulta.Close
    Screen.MousePointer = 0
    Exit Sub

errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub
Private Sub ArmoConsulta()
'----------------------------------------------------------------------------------------------------
'Si el rubro no es importaciones leo tabla Compra y GastoSubRubro.
'Si es importaciones busco directamente en GastoImportacion.
'----------------------------------------------------------------------------------------------------
On Error GoTo ErrArmoCons
        
    If chDetalle.Value = 1 Then
        
        'Union de folder.
        Cons = "Select ComProveedor, ComCodigo, ComFecha, SRuNombre, SRuCodigo, CarCodigo,  ComTipoDocumento, ComSerie, ComNumero, ComImporte, ComIva, ComMoneda, ComTC, PClFantasia" _
            & " From Compra (Index = iComFecha), ProveedorCliente, GastoImportacion, SubRubro, Carpeta " _
            & " Where ComFecha Between '" & Format("1-" & tMes.Text & " 00:00:00", sqlFormatoFH) & "'" _
            & " And '" & Format(UltimoDia(CDate("1-" & tMes.Text)) & " 23:59:59", sqlFormatoFH) & "'" _
            & " And SRuRubro = " & cRubro.ItemData(cRubro.ListIndex) _
            & " And GImNivelFolder = " & Folder.cFCarpeta & " And GImFolder = CarID " _
            & " And ComCodigo = GImIDCompra And ComProveedor = PClCodigo And SRuID = GimIDSubRubro"
        If cProveedor.ListIndex > -1 Then Cons = Cons & " And ComProveedor = " & cProveedor.ItemData(cProveedor.ListIndex)
        
        If cSubRubro.ListIndex > -1 Then Cons = Cons & " And SRuID = " & cSubRubro.ItemData(cSubRubro.ListIndex)
        If OpComprobante(0).Value Then Cons = Cons & " And ComTipoDocumento IN(" & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraCredito & ")"
        If OpComprobante(1).Value Then Cons = Cons & " And ComTipoDocumento IN(" & TipoDocumento.CompraNotaDevolucion & ", " & TipoDocumento.Compracontado & ", " & TipoDocumento.CompraRecibo & ")"
        If OpComprobante(2).Value Then Cons = Cons & " And ComTipoDocumento IN(" & TipoDocumento.CompraSalidaCaja & ", " & TipoDocumento.CompraEntradaCaja & ")"
        
        Cons = Cons & " Union All " _
            & "Select ComProveedor, ComCodigo, ComFecha, SRuNombre, SRuCodigo, CarCodigo,  ComTipoDocumento, ComSerie, ComNumero, ComImporte, ComIva, ComMoneda, ComTC, PClFantasia" _
            & " From Compra (Index = iComFecha), ProveedorCliente, GastoImportacion, SubRubro, Carpeta, Embarque " _
            & " Where ComFecha Between '" & Format("1-" & tMes.Text & " 00:00:00", sqlFormatoFH) & "'" _
            & " And '" & Format(UltimoDia(CDate("1-" & tMes.Text)) & " 23:59:59", sqlFormatoFH) & "'" _
            & " And SRuRubro = " & cRubro.ItemData(cRubro.ListIndex) _
            & " And GImNivelFolder = " & Folder.cFEmbarque & " And GImFolder = EmbID And EmbCarpeta = CarID" _
            & " And ComCodigo = GImIDCompra And ComProveedor = PClCodigo And SRuID = GimIDSubRubro"
        If cProveedor.ListIndex > -1 Then Cons = Cons & " And ComProveedor = " & cProveedor.ItemData(cProveedor.ListIndex)
        
        If cSubRubro.ListIndex > -1 Then Cons = Cons & " And SRuID = " & cSubRubro.ItemData(cSubRubro.ListIndex)
        If OpComprobante(0).Value Then Cons = Cons & " And ComTipoDocumento IN(" & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraCredito & ")"
        If OpComprobante(1).Value Then Cons = Cons & " And ComTipoDocumento IN(" & TipoDocumento.CompraNotaDevolucion & ", " & TipoDocumento.Compracontado & ", " & TipoDocumento.CompraRecibo & ")"
        If OpComprobante(2).Value Then Cons = Cons & " And ComTipoDocumento IN(" & TipoDocumento.CompraSalidaCaja & ", " & TipoDocumento.CompraEntradaCaja & ")"
        
        Cons = Cons & " Union All " _
            & "Select ComProveedor, ComCodigo, ComFecha, SRuNombre, SRuCodigo, CarCodigo,  ComTipoDocumento, ComSerie, ComNumero, ComImporte, ComIva, ComMoneda, ComTC, PClFantasia" _
            & " From Compra (Index = iComFecha), ProveedorCliente, GastoImportacion, SubRubro, Carpeta, Embarque, SubCarpeta " _
            & " Where ComFecha Between '" & Format("1-" & tMes.Text & " 00:00:00", sqlFormatoFH) & "'" _
            & " And '" & Format(UltimoDia(CDate("1-" & tMes.Text)) & " 23:59:59", sqlFormatoFH) & "'"
        
        If cSubRubro.ListIndex > -1 Then Cons = Cons & " And SRuID = " & cSubRubro.ItemData(cSubRubro.ListIndex)
        If OpComprobante(0).Value Then Cons = Cons & " And ComTipoDocumento IN(" & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraCredito & ")"
        If OpComprobante(1).Value Then Cons = Cons & " And ComTipoDocumento IN(" & TipoDocumento.CompraNotaDevolucion & ", " & TipoDocumento.Compracontado & ", " & TipoDocumento.CompraRecibo & ")"
        If OpComprobante(2).Value Then Cons = Cons & " And ComTipoDocumento IN(" & TipoDocumento.CompraSalidaCaja & ", " & TipoDocumento.CompraEntradaCaja & ")"
        
        Cons = Cons & " And ComCodigo = GImIDCompra And SRuRubro = " & cRubro.ItemData(cRubro.ListIndex) _
            & " And GImNivelFolder = " & Folder.cFSubCarpeta & " And GImFolder = SubID And SubEmbarque = EmbID And EmbCarpeta = CarID" _
            & " And ComProveedor = PClCodigo And SRuID = GimIDSubRubro"
        
        'El de proveedor esta despues del endif.
    Else
        
        Cons = "Select ComProveedor, ComCodigo, CarCodigo = '',ComFecha, SRuNombre = '' , SRuCodigo = '' , ComTipoDocumento, ComSerie, ComNumero, ComImporte, ComIva, ComMoneda, ComTC, PClFantasia" _
            & " From Compra (Index = iComFecha) , ProveedorCliente " _
            & " Where ComFecha Between '" & Format("1-" & tMes.Text & " 00:00:00", sqlFormatoFH) & "'" _
            & " And '" & Format(UltimoDia(CDate("1-" & tMes.Text)) & " 23:59:59", sqlFormatoFH) & "'"
        
        If OpComprobante(0).Value Then Cons = Cons & " And ComTipoDocumento IN(" & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraCredito & ")"
        If OpComprobante(1).Value Then Cons = Cons & " And ComTipoDocumento IN(" & TipoDocumento.CompraNotaDevolucion & ", " & TipoDocumento.Compracontado & ", " & TipoDocumento.CompraRecibo & ")"
        If OpComprobante(2).Value Then Cons = Cons & " And ComTipoDocumento IN(" & TipoDocumento.CompraSalidaCaja & ", " & TipoDocumento.CompraEntradaCaja & ")"
        
        Cons = Cons & " And ComProveedor = PClCodigo "
        If cRubro.ListIndex > -1 Then
            Cons = Cons & " And ComCodigo IN (Select  GSrIDCompra From  GastoSubRubro, SubRubro Where SRuID = GSrIDSubRubro" _
                & " And SRuRubro = " & cRubro.ItemData(cRubro.ListIndex)
            If cSubRubro.ListIndex > -1 Then Cons = Cons & " And SRuID = " & cSubRubro.ItemData(cSubRubro.ListIndex)
            Cons = Cons & ")"
        End If
    End If
    
    If cProveedor.ListIndex > -1 Then Cons = Cons & " And ComProveedor = " & cProveedor.ItemData(cProveedor.ListIndex)
    Cons = Cons & " Order by ComProveedor, ComFecha, ComCodigo"
    
    Exit Sub
ErrArmoCons:
    clsGeneral.OcurrioError "Ocurrio un error al armar la consulta.", Trim(Err.Description)
End Sub

Private Sub CargoLista()
Dim tSerie As String
Dim CodCompra As Long, CodProveedor As Long
    
    On Error GoTo ErrInesperado
    
    vsConsulta.Rows = 1
    vsConsulta.Refresh
    vsConsulta.Redraw = False
    CodCompra = 0: CodProveedor = 0
    Screen.MousePointer = 11
    Do While Not RsConsulta.EOF
        'Inserto en la grilla.------------------------------------------------
        With vsConsulta
            If RsConsulta!ComProveedor <> CodProveedor Then
                CodProveedor = RsConsulta!ComProveedor
                .AddItem ""
                .Cell(flexcpText, vsConsulta.Rows - 1, 0) = Trim(RsConsulta!PClFantasia)
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = inactivo
            End If
            .AddItem ""
            .Cell(flexcpText, vsConsulta.Rows - 1, 0) = Trim(RsConsulta!PClFantasia)
            .Cell(flexcpText, vsConsulta.Rows - 1, 1) = Format(RsConsulta!ComFecha, "dd/mm/yy")
            .Cell(flexcpText, vsConsulta.Rows - 1, 2) = Trim(RsConsulta!SRuNombre)
            .Cell(flexcpText, vsConsulta.Rows - 1, 3) = Trim(RsConsulta!CarCodigo)
            If CodCompra <> RsConsulta!ComCodigo Then
                If Not IsNull(RsConsulta!ComSerie) Then tSerie = RsConsulta!ComSerie Else tSerie = ""
                .Cell(flexcpText, vsConsulta.Rows - 1, 4) = Trim(RetornoNombreDocumento(RsConsulta!ComTipoDocumento, True)) & " " & tSerie & " " & RsConsulta!ComNumero
                If RsConsulta!ComMoneda = paMonedaPesos Then
                    .Cell(flexcpText, vsConsulta.Rows - 1, 5) = Format(RsConsulta!ComImporte, FormatoMonedaP)
                    If Not IsNull(RsConsulta!ComIVA) Then .Cell(flexcpText, vsConsulta.Rows - 1, 6) = Format(RsConsulta!ComIVA, FormatoMonedaP)
                    .Cell(flexcpText, vsConsulta.Rows - 1, 7) = Format(.Cell(flexcpValue, vsConsulta.Rows - 1, 6) + .Cell(flexcpValue, vsConsulta.Rows - 1, 5), FormatoMonedaP)
                    .Cell(flexcpText, vsConsulta.Rows - 1, 8) = ""
                    .Cell(flexcpText, vsConsulta.Rows - 1, 9) = ""
                    .Cell(flexcpText, vsConsulta.Rows - 1, 10) = ""
                Else
                    .Cell(flexcpText, vsConsulta.Rows - 1, 5) = Format(RsConsulta!ComImporte * RsConsulta!ComTC, FormatoMonedaP)
                    .Cell(flexcpText, vsConsulta.Rows - 1, 6) = Format(RsConsulta!ComIVA * RsConsulta!ComTC, FormatoMonedaP)
                    .Cell(flexcpText, vsConsulta.Rows - 1, 7) = Format(.Cell(flexcpValue, vsConsulta.Rows - 1, 6) + .Cell(flexcpValue, vsConsulta.Rows - 1, 5), FormatoMonedaP)
                    .Cell(flexcpText, vsConsulta.Rows - 1, 8) = Format(RsConsulta!ComImporte, FormatoMonedaP)
                    .Cell(flexcpText, vsConsulta.Rows - 1, 9) = Format(RsConsulta!ComIVA, FormatoMonedaP)
                    .Cell(flexcpText, vsConsulta.Rows - 1, 10) = Format(.Cell(flexcpValue, vsConsulta.Rows - 1, 8) + .Cell(flexcpValue, vsConsulta.Rows - 1, 9), FormatoMonedaP)
                End If
                CodCompra = RsConsulta!ComCodigo
            End If
            .Cell(flexcpText, vsConsulta.Rows - 1, 11) = Format(RsConsulta!ComTC, FormatoMonedaP)
        End With
        RsConsulta.MoveNext
    Loop
    vsConsulta.Subtotal flexSTClear
    With vsConsulta
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, 0, 5, , &HFFFFC0, , , "%s"
        .Subtotal flexSTSum, 0, 6, , , , True
        .Subtotal flexSTSum, 0, 7, , , , True
        .Subtotal flexSTSum, 0, 8, , , , True
        .Subtotal flexSTSum, 0, 9, , , , True
        .Subtotal flexSTSum, 0, 10, , , , True
        
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 5, , Colores.Obligatorio, &H80&, True, "Total", , True
        .Subtotal flexSTSum, -1, 6, , Colores.Obligatorio, , True, "Total", , True
        .Subtotal flexSTSum, -1, 7, , Colores.Obligatorio, , True, "Total", , True
     End With
     
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
    Exit Sub
    
ErrInesperado:
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de stock.", Err.Description
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
Dim J As Integer
Dim AnchoAnt As Double

    If vsConsulta.Rows = 1 Then
        MsgBox "No hay datos en la lista para realizar la impresión.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    On Error GoTo errPrint
    Screen.MousePointer = 11
    aTituloTabla = "": aComentario = ""
    
    With vsListado
        If chDetalle.Value = vbUnchecked Then .Orientation = orPortrait Else .Orientation = orLandscape
        .StartDoc
                
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN"
            Screen.MousePointer = vbDefault: Exit Sub
        End If
        
        aTexto = ""
        If cProveedor.ListIndex <> -1 Then aTexto = Trim(cProveedor.Text) & " / "
        If OpComprobante(0).Value Then aTexto = aTexto & OpComprobante(0).Tag
        If OpComprobante(1).Value Then aTexto = aTexto & OpComprobante(1).Tag
        If OpComprobante(2).Value Then aTexto = aTexto & OpComprobante(2).Tag
        If cRubro.ListIndex <> -1 Then aTexto = aTexto & " / Rubro: " & Trim(cRubro.Text)
        
        EncabezadoListado vsListado, aTexto & " / " & Trim(tMes.Text), False
        
        .MarginLeft = 350: .MarginRight = 350
        .filename = "Consulta de Comprobantes"
        .FontSize = 8: .FontBold = False
        
        vsConsulta.ExtendLastCol = False
        AnchoAnt = vsConsulta.ColWidth(0)
        .RenderControl = vsConsulta.hWnd
        vsConsulta.ColWidth(0) = AnchoAnt
        vsConsulta.ExtendLastCol = True
        
        .EndDoc
    End With
    
    If Imprimir Then
        frmSetup.pControl = vsListado
        frmSetup.Show vbModal, Me
        Me.Refresh
        If frmSetup.pOK Then vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
    End If
    
    
    Screen.MousePointer = 0
    Exit Sub

errPrint:
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión. ", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ArmoFormulaFiltros() As String

Dim aRetorno As String

    On Error Resume Next
    aRetorno = ""
    
    If cProveedor.ListIndex <> -1 Then aRetorno = aRetorno & " Prov.: " & cProveedor.Text & ", "
    
    aRetorno = Mid(aRetorno, 1, Len(aRetorno) - 2)
    ArmoFormulaFiltros = aRetorno

End Function

Private Function VerificoFiltros() As Boolean

    VerificoFiltros = False
    
    If cProveedor.Text <> "" And cProveedor.ListIndex = -1 Then
        MsgBox "El proveedor de artículos no es correcto.", vbExclamation, "ATENCIÓN"
        Foco cProveedor: Exit Function
    End If
    If cRubro.Text <> "" And cRubro.ListIndex = -1 Then
        MsgBox "El Rubro no es correcto.", vbExclamation, "ATENCIÓN"
        Foco cRubro: Exit Function
    End If
    
    If Not IsDate(tMes.Text) Then
        MsgBox "Es necesario que ingrese un mes válido de consulta.", vbExclamation, "ATENCIÓN"
        Foco tMes: Exit Function
    End If
    
    VerificoFiltros = True

End Function

Private Sub tMes_GotFocus()
    With tMes
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub tMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cRubro
End Sub
Private Sub tMes_LostFocus()
    If IsDate(tMes.Text) Then tMes.Text = Format(tMes.Text, "Mmm-yyyy") Else tMes.Text = ""
End Sub

Private Sub vsConsulta_LostFocus()
    vsConsulta.Select 0, 0, 0, 0
End Sub
Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub
Private Sub LimpioGrilla()
    
    With vsConsulta
        .Redraw = False
        .Clear
        .Cols = 13
        .ExtendLastCol = True
        .ScrollTrack = False
        .ScrollTips = True
        .GridLines = flexGridNone
        .Rows = 1
        .FixedRows = 1
        .FormatString = "|^Fecha|<Gasto|<Carpeta|<Comprobante|>Importe $|>I.V.A. $|>Total $|>Importe U$S|>I.V.A. U$S|>Total U$S|T.C."
        .ColWidth(0) = 110
        .ColWidth(1) = 750
        .ColWidth(2) = 1700
        .ColWidth(3) = 680
        .ColWidth(4) = 1500
        .ColWidth(5) = 1300
        .ColDataType(5) = flexDTCurrency
        .ColWidth(6) = 1050
        .ColDataType(6) = flexDTCurrency
        .ColWidth(7) = 1400
        .ColDataType(7) = flexDTCurrency
        .ColWidth(8) = 1300
        .ColDataType(8) = flexDTCurrency
        .ColWidth(9) = 1050
        .ColDataType(9) = flexDTCurrency
        .ColWidth(10) = 1400
        .ColDataType(10) = flexDTCurrency
        .ColWidth(11) = 750
        .ColWidth(12) = 50
        .AllowUserResizing = flexResizeColumns
        If chDetalle.Value Then
            .ColHidden(2) = False
            .ColHidden(3) = False
        Else
            .ColHidden(3) = True
            .ColHidden(2) = True
        End If
        .MergeCells = flexMergeSpill
        .Redraw = True
    End With
    
End Sub

Private Sub CargoSubRubros()
On Error GoTo ErrCSR
    cSubRubro.Clear
    If cRubro.ListIndex = -1 Then Exit Sub
    Screen.MousePointer = 11
    Cons = "Select SRuID, SRuNombre From SubRubro Where SRuRubro = " & cRubro.ItemData(cRubro.ListIndex) & " Order by SRuNombre"
    CargoCombo Cons, cSubRubro
    Screen.MousePointer = 0
    Exit Sub
ErrCSR:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los subrubros.", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub
