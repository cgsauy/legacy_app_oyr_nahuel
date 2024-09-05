VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form frmListado 
   Caption         =   "Recepción de Traslados de Servicios"
   ClientHeight    =   6285
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7800
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
   ScaleHeight     =   6285
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4455
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7858
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
      AllowBigSelection=   0   'False
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
      Left            =   60
      TabIndex        =   20
      Top             =   720
      Width           =   7335
      _Version        =   196608
      _ExtentX        =   12938
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
      Left            =   50
      ScaleHeight     =   435
      ScaleWidth      =   6075
      TabIndex        =   21
      Top             =   5400
      Width           =   6135
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4140
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bGrabar 
         Height          =   310
         Left            =   3420
         Picture         =   "frmListado.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Grabar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2760
         Picture         =   "frmListado.frx":0B78
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2400
         Picture         =   "frmListado.frx":0C62
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
         Picture         =   "frmListado.frx":0D4C
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
         Picture         =   "frmListado.frx":0F86
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   4800
         Picture         =   "frmListado.frx":1088
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
         Picture         =   "frmListado.frx":118A
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
         Picture         =   "frmListado.frx":148C
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
         Picture         =   "frmListado.frx":17CE
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
         Picture         =   "frmListado.frx":1AD0
         Style           =   1  'Graphical
         TabIndex        =   8
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
      TabIndex        =   19
      Top             =   6030
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10663
            TextSave        =   ""
            Key             =   "msg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Picture         =   "frmListado.frx":1D0A
            TextSave        =   ""
            Key             =   "printer"
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
      Height          =   675
      Left            =   50
      TabIndex        =   18
      Top             =   0
      Width           =   7755
      Begin VB.TextBox tCodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   780
         MaxLength       =   5
         TabIndex        =   1
         Top             =   240
         Width           =   555
      End
      Begin AACombo99.AACombo cIngreso 
         Height          =   315
         Left            =   5820
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
      Begin AACombo99.AACombo cCamion 
         Height          =   315
         Left            =   3120
         TabIndex        =   3
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
         BackStyle       =   0  'Transparent
         Caption         =   "&Código:"
         Height          =   255
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "&Origen en:"
         Height          =   195
         Left            =   4980
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Camión que &Traslada:"
         Height          =   255
         Left            =   1500
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin ComctlLib.ImageList ImgList 
      Left            =   7200
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmListado.frx":1E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmListado.frx":2136
            Key             =   "retiro"
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

Private Enum EstadoS
    Anulado = 0
    Visita = 1
    Retiro = 2
    Taller = 3
    Entrega = 4
    Cumplido = 5
End Enum

Private aTexto As String

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub bImprimir_Click()
    AccionImprimir True
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


Private Sub cCamion_GotFocus()
    With cCamion
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione un camión y presione F1 (ver pendientes)"
End Sub

Private Sub cCamion_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn: Foco cIngreso
        Case vbKeyF1
            If cCamion.ListIndex > -1 Then
                Cons = "Select Distinct(TseCodigo), Traslado = TSeCodigo, Destino = LocNombre From Servicio, Taller, TrasladoServicio, Local " _
                    & " Where SerEstadoServicio = " & EstadoS.Taller _
                    & " And TalIngresoCamion = " & cCamion.ItemData(cCamion.ListIndex) _
                    & " And SerCodigo = TalServicio And SerLocalReparacion = LocCodigo" _
                    & " And TalFIngresoRealizado Is Not Null And TalFIngresoRecepcion = Null And SerProducto = TSeProducto"
                Dim objLista As New clsListadeAyuda
                objLista.ActivoListaAyuda Cons, False, txtConexion, 3000
                Set objLista = Nothing
            End If
    End Select
End Sub

Private Sub cCamion_LostFocus()
    cCamion.SelStart = 0
End Sub

Private Sub cIngreso_GotFocus()
    With cIngreso
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cIngreso_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub
Private Sub cIngreso_LostFocus()
    cIngreso.SelStart = 0
End Sub


Private Sub chVista_Click()
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
        Me.Refresh
    Else
        AccionImprimir
        vsListado.ZOrder 0
        Me.Refresh
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad

    Status.Panels("printer").Text = paPrintCartaD
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    InicializoGrillas
    FechaDelServidor
    
    Cons = "Select CamCodigo, CamNombre From Camion Order by CamNombre"
    CargoCombo Cons, cCamion
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Order by SucAbreviacion"
    CargoCombo Cons, cIngreso
    
    vsListado.Orientation = orPortrait: vsListado.PaperSize = 1
    PrueboBandejaImpresora
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub
Private Sub InicializoGrillas()
    On Error Resume Next
    With vsConsulta
        .Redraw = False
        .WordWrap = False
        .Cols = 1: .Rows = 1
        .FormatString = "Selección|>Servicio|<Producto|<Ingresó|<Local |<Camion|"
        .ColWidth(2) = 3900: .ColWidth(3) = 1200: .ColWidth(4) = 1200: .ColWidth(5) = 1200: .ColWidth(6) = 10
        .ColDataType(0) = flexDTBoolean
        .Redraw = True
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
            
            Case vbKeyI: AccionImprimir True
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            
            Case vbKeyG: AccionGrabar
            
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
Dim rs As rdoResultset
    On Error GoTo errConsultar
    Screen.MousePointer = 11
    chVista.Value = 0
    vsConsulta.Rows = 1
    vsConsulta.Refresh
    vsConsulta.Redraw = False
    CargoServicios
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True
End Sub

Private Sub Label2_Click()
    Foco cCamion
End Sub
Private Sub Label3_Click()
    Foco cIngreso
End Sub

Private Sub Status_PanelClick(ByVal Panel As ComctlLib.Panel)
    If "printer" = Panel.Key Then
        prj_GetPrinter True
        Panel.Text = paPrintCartaD
    End If
End Sub

Private Sub tCodigo_Change()
    vsConsulta.Rows = 1
End Sub

Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCodigo.Text) Then AccionConsultar: bConsultar.SetFocus Else Foco cCamion
    End If
End Sub

Private Sub vsConsulta_DblClick()
    If vsConsulta.Row >= 1 Then
        If vsConsulta.Cell(flexcpChecked, vsConsulta.Row, 0) = flexChecked Then
            vsConsulta.Cell(flexcpChecked, vsConsulta.Row, 0) = flexUnchecked
        Else
            vsConsulta.Cell(flexcpChecked, vsConsulta.Row, 0) = flexChecked
        End If
    End If
End Sub

Private Sub vsConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        
        Case vbKeySpace
            If vsConsulta.Row >= 1 Then
                If vsConsulta.Cell(flexcpChecked, vsConsulta.Row, 0) = flexChecked Then
                    vsConsulta.Cell(flexcpChecked, vsConsulta.Row, 0) = flexUnchecked
                Else
                    vsConsulta.Cell(flexcpChecked, vsConsulta.Row, 0) = flexChecked
                End If
            End If
    End Select
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub

    SeteoImpresoraPorDefecto paPrintCartaD
    With vsListado
        .Device = paPrintCartaD
        .Orientation = orPortrait
        .PaperSize = paPrintCartaPaperSize
        .PaperBin = paPrintCartaB         'Bandeja por defecto.
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    
    If cCamion.ListIndex = -1 Then
        EncabezadoListado vsListado, "Recepción de Traslados de Servicios ", False
    Else
        EncabezadoListado vsListado, "Recepción de Traslados de Servicios, Camión: " & Trim(cCamion.Text), False
    End If
    vsListado.FileName = "Recepción de Traslados de Servicios"
    vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
    vsListado.EndDoc
    
    If Imprimir Then vsListado.PrintDoc
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub

Private Sub Ayuda(strTexto As String)
    Status.Panels("msg").Text = strTexto
End Sub

Private Sub CargoServicios()
Dim aModificacion As String, aValor As Long
    
    If Not IsNumeric(tCodigo.Text) Then MsgBox "Debe ingresar el código de traslado.", vbExclamation, "ATENCIÓN": Exit Sub
    
    'Cargo Impresos.
    Cons = "Select * From Servicio, Taller, Producto, Articulo, Sucursal, Camion, TrasladoServicio " _
        & " Where SerEstadoServicio = " & EstadoS.Taller _
        & " And SerLocalReparacion = " & paCodigoDeSucursal _
        & " And TalFIngresoRealizado Is Not Null And TalFIngresoRecepcion is Null "
    
    If IsNumeric(tCodigo.Text) Then Cons = Cons & " And TSeCodigo = " & CLng(tCodigo.Text)
    If cIngreso.ListIndex > -1 Then Cons = Cons & " And SerLocalIngreso = " & cIngreso.ItemData(cIngreso.ListIndex)
    
    Cons = Cons & " And SerLocalIngreso <> SerLocalReparacion And SerCodigo = TalServicio And SerProducto = ProCodigo And ProArticulo = ArtID " _
        & " And SerLocalIngreso = SucCodigo And TalIngresoCamion  =  CamCodigo And ProCodigo = TSeProducto "
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsAux.EOF
    
        With vsConsulta
            .AddItem ""
            
            aModificacion = RsAux!SerModificacion: .Cell(flexcpData, .Rows - 1, 0) = aModificacion
            aModificacion = RsAux!TalModificacion: .Cell(flexcpData, .Rows - 1, 1) = aModificacion
            .Cell(flexcpData, .Rows - 1, 2) = 1 'Son para reparar.
            
            If RsAux!ProCliente = paClienteEmpresa Then .Cell(flexcpData, .Rows - 1, 3) = paClienteEmpresa Else .Cell(flexcpData, .Rows - 1, 3) = 0
            aValor = RsAux!ProArticulo: .Cell(flexcpData, .Rows - 1, 4) = aValor
            aValor = RsAux!CamCodigo: .Cell(flexcpData, .Rows - 1, 5) = aValor
            
            .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked
            .Cell(flexcpText, .Rows - 1, 1) = RsAux!SerCodigo
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ProCodigo, "(#,000)") & " " & Trim(RsAux!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!SerFecha, FormatoFP)
            .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!SucAbreviacion)
            .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!CamNombre)
            
        End With
        RsAux.MoveNext
        
    Loop
    RsAux.Close
    
    'Cargo los que son para entregar.
    Cons = "Select * From Servicio, Taller, Producto, Articulo, Sucursal, Camion, TrasladoServicio " _
        & " Where SerEstadoServicio = " & EstadoS.Taller _
        & " And TalLocalAlCliente = " & paCodigoDeSucursal _
        & " And TalFIngresoRealizado Is Not Null And TalFSalidaRealizado <> Null And TalFSalidaRecepcion = Null "
    
    If IsNumeric(tCodigo.Text) Then Cons = Cons & " And TSeCodigo = " & CLng(tCodigo.Text)
    If cIngreso.ListIndex > -1 Then Cons = Cons & " And SerLocalReparacion = " & cIngreso.ItemData(cIngreso.ListIndex)
    
    'And SerLocalIngreso <> SerLocalReparacion
    Cons = Cons & " And SerCodigo = TalServicio And SerProducto = ProCodigo And ProArticulo = ArtID " _
        & " And SerLocalIngreso = SucCodigo And TalSalidaCamion  =  CamCodigo And ProCodigo = TSeProducto "
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsAux.EOF
    
        With vsConsulta
            .AddItem ""
            
            aModificacion = RsAux!SerModificacion: .Cell(flexcpData, .Rows - 1, 0) = aModificacion
            aModificacion = RsAux!TalModificacion: .Cell(flexcpData, .Rows - 1, 1) = aModificacion
            .Cell(flexcpData, .Rows - 1, 2) = 2 'Son para reparar.
            If RsAux!ProCliente = paClienteEmpresa Then .Cell(flexcpData, .Rows - 1, 3) = paClienteEmpresa Else .Cell(flexcpData, .Rows - 1, 3) = 0
            aValor = RsAux!ProArticulo: .Cell(flexcpData, .Rows - 1, 4) = aValor
            aValor = RsAux!CamCodigo: .Cell(flexcpData, .Rows - 1, 5) = aValor
            
            .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked
            .Cell(flexcpText, .Rows - 1, 1) = RsAux!SerCodigo
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ProCodigo, "(#,000)") & " " & Trim(RsAux!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!SerFecha, FormatoFP)
            .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!SucAbreviacion)
            .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!CamNombre)
            
        End With
        RsAux.MoveNext
        
    Loop
    RsAux.Close
    
End Sub

Private Sub AccionGrabar()
Dim Msg As String, Usuario As String
    
    If vsConsulta.Rows = 1 Then MsgBox "No hay datos en la consulta.", vbExclamation, "ATENCIÓN": Exit Sub
    
    If MsgBox("¿Confirma grabar la recepción del traslado?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    
    Usuario = ""
    Usuario = InputBox("Ingrese su digito de usuario.", "Grabar Traslado")
    If Not IsNumeric(Usuario) Then Exit Sub
    
    Usuario = BuscoUsuarioDigito(CLng(Usuario), True)
    
    If Val(Usuario) = 0 Then MsgBox "Usuario incorrecto.", vbExclamation, "ATENCIÓN": Exit Sub
    
    On Error GoTo ErrBT
    Screen.MousePointer = 11
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo ErrRB
    With vsConsulta
        For I = 1 To .Rows - 1
            If .Cell(flexcpChecked, I, 0) = flexChecked Then
                Msg = ""
                'TABLA SERVICIO
                Cons = "Select * From Servicio Where SerCodigo = " & Val(.Cell(flexcpText, I, 1))
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If RsAux.EOF Then
                    Msg = "Otra terminal eliminó el servicio = " & Val(.Cell(flexcpText, I, 1))
                    RsAux.Close: RsAux.Edit 'Provoco error.
                Else
                    If RsAux!SerModificacion = CDate(.Cell(flexcpData, I, 0)) Then
                        RsAux.Edit
                        RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
                        RsAux.Update
                        RsAux.Close
                    Else
                        Msg = "Otra terminal modifico el servicio = " & Val(.Cell(flexcpText, I, 1))
                        RsAux.Close: RsAux.Edit 'Provoco error.
                    End If
                End If
                
                'Tabla Taller
                Cons = "Select * From Taller Where TalServicio = " & Val(.Cell(flexcpText, I, 1))
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If RsAux.EOF Then
                    Msg = "Otra terminal eliminó los datos de Taller."
                    RsAux.Close: RsAux.Edit 'Provoco error.
                Else
                    If RsAux!TalModificacion = CDate(.Cell(flexcpData, I, 1)) Then
                        RsAux.Edit
                        RsAux!TalModificacion = Format(gFechaServidor, sqlFormatoFH)
                        If Val(.Cell(flexcpData, I, 2)) = 1 Then
                            RsAux!TalFIngresoRecepcion = Format(gFechaServidor, sqlFormatoFH)
                        Else
                            RsAux!TalFSalidaRecepcion = Format(gFechaServidor, sqlFormatoFH)
                        End If
                        RsAux.Update
                        RsAux.Close
                        'Si el artículo es de carlos gutierrez marco el traslado en el stock.
                        If CLng(.Cell(flexcpData, I, 3)) > 0 Then HagoCambioDeEstado CLng(.Cell(flexcpData, I, 4)), Val(.Cell(flexcpText, I, 1)), CLng(Usuario), -1, CLng(.Cell(flexcpData, I, 5))
                        
                    Else
                        Msg = "Otra terminal modificó los datos de taller."
                        RsAux.Close: RsAux.Edit 'Provoco error.
                    End If
                End If
            End If
        Next I
    End With
    cBase.CommitTrans
    If MsgBox("¿Desea imprimir los traslados recepcionados?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
        'Como voy a imprimir eliminó los que no están asignados.
        With vsConsulta
            I = 1
            Do While I <= .Rows - 1
                If .Cell(flexcpChecked, I, 0) = flexUnchecked Then .RemoveItem I: I = I - 1
                I = I + 1
            Loop
        End With
        AccionImprimir True
    End If
    AccionConsultar
    Screen.MousePointer = 0
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrRB:
    Resume ErrVA
ErrVA:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar almacenar la información." & Chr(13) & Msg, Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub HagoCambioDeEstado(IDArticulo As Long, IdServicio As Long, IdUsuario As Long, AltaBajaLocal As Integer, IdCamion As Long)
'Si Altabajalocal = -1 entonces le doy de baja al local sino le doy de baja al camion .
    'Cedo el artículo al camión.
    MarcoMovimientoStockFisico IdUsuario, TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, AltaBajaLocal * -1, TipoDocumento.ServicioCambioEstado, IdServicio
    MarcoMovimientoStockFisico IdUsuario, TipoLocal.Camion, IdCamion, IDArticulo, 1, paEstadoARecuperar, AltaBajaLocal, TipoDocumento.ServicioCambioEstado, IdServicio
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Deposito, paCodigoDeSucursal, IDArticulo, 1, paEstadoARecuperar, AltaBajaLocal * -1
    MarcoMovimientoStockFisicoEnLocal TipoLocal.Camion, IdCamion, IDArticulo, 1, paEstadoARecuperar, AltaBajaLocal
End Sub

Private Sub PrueboBandejaImpresora()
On Error GoTo ErrPBI
    
    With vsListado
        .PaperSize = 1  'Hoja carta
        .Orientation = orPortrait
    End With
'        .Device = paICartaN
'        If .Device <> paICartaN Then MsgBox "Ud no tiene instalada la impresora para hoja blanca. Avise al administrador.", vbExclamation, "ATENCIÒN"
'        If .PaperBins(paICartaB) Then .PaperBin = paICartaB Else MsgBox "Esta mal definida la bandeja de hoja blanca en su sucursal, comuniquele al administrador.", vbInformation, "ATENCIÓN": paICartaB = .PaperBin
'    End With
    Exit Sub
ErrPBI:
    clsGeneral.OcurrioError "Ocurrio un error al setear la impresora, consulte con el administrador de impresión este problema.", Err.Description
End Sub


