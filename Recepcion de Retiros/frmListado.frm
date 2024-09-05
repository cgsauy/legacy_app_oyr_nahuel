VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form frmListado 
   Caption         =   "Recepción de Retiros"
   ClientHeight    =   6285
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11730
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
   ScaleWidth      =   11730
   StartUpPosition =   3  'Windows Default
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4455
      Left            =   240
      TabIndex        =   2
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
      TabIndex        =   16
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
      TabIndex        =   17
      Top             =   5400
      Width           =   6135
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":030A
         Height          =   310
         Left            =   4140
         Picture         =   "frmListado.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bGrabar 
         Height          =   310
         Left            =   3420
         Picture         =   "frmListado.frx":093E
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Grabar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2760
         Picture         =   "frmListado.frx":0A40
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2400
         Picture         =   "frmListado.frx":0B2A
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmListado.frx":0C14
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3780
         Picture         =   "frmListado.frx":0E4E
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   4800
         Picture         =   "frmListado.frx":0F50
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
         Picture         =   "frmListado.frx":1052
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmListado.frx":1354
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmListado.frx":1696
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmListado.frx":1998
         Style           =   1  'Graphical
         TabIndex        =   4
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
      TabIndex        =   15
      Top             =   6030
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   20161
            TextSave        =   ""
            Key             =   "msg"
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
      TabIndex        =   14
      Top             =   0
      Width           =   10215
      Begin VB.CommandButton butRetiroTraslada 
         Caption         =   "Recepciono Retiro para trasladar"
         Height          =   315
         Left            =   7320
         TabIndex        =   19
         ToolTipText     =   "El producto se repara en el local ya asignado"
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton butForzado 
         Caption         =   "Recepciono Retiro en mi local"
         Height          =   315
         Left            =   4440
         TabIndex        =   18
         ToolTipText     =   "El producto se repara en mi sucursal"
         Top             =   240
         Width           =   2655
      End
      Begin AACombo99.AACombo cCamion 
         Height          =   315
         Left            =   1740
         TabIndex        =   1
         Top             =   240
         Width           =   1875
         _ExtentX        =   3307
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Camión que Retira:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
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
            Picture         =   "frmListado.frx":1BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmListado.frx":1EEC
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
Private aTexto As String

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar True
End Sub
Private Sub bGrabar_Click()
    AccionGrabar 0
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

Private Sub butForzado_Click()
On Error GoTo errBF
    Dim idServicio As Long, fEdit As Date
    Cons = InputBox("Ingrese el código del servicio", "Recepcionar retiro")
    If IsNumeric(Cons) Then
        Cons = "SELECT SerCodigo, SerModificacion FROM Servicio INNER JOIN ServicioVisita ON SerCodigo = VisServicio And VisFImpresion IS NOT Null And VisSinEfecto = 0" _
            & " Where SerCodigo = " & Cons & " AND SerEstadoServicio = " & EstadoS.Retiro _
            & " And VisTipo = " & TipoServicio.Retiro & " And VisCamion IS NOT NULL " _
            & " And VisFImpresion IS NOT Null And VisSinEfecto = 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            idServicio = RsAux("SerCodigo")
            fEdit = RsAux("SerModificacion")
        Else
            MsgBox "No hay un retiro para recepcionar para ese código.", vbExclamation, "ATENCIÓN"
        End If
        RsAux.Close
        If idServicio > 0 Then
            AccionGrabar idServicio, fEdit, True
        End If
    Else
        MsgBox "No ingresó un número.", vbCritical, "ATENCIÓN"
    End If
    Exit Sub
errBF:
    clsGeneral.OcurrioError "Error al intentar recepcionar el servicio.", Err.Description
End Sub

Private Sub butRetiroTraslada_Click()
On Error GoTo errBF
    Dim idServicio As Long, fEdit As Date
    Cons = InputBox("Ingrese el código del servicio", "Recepcionar retiro")
    If IsNumeric(Cons) Then
        Cons = "SELECT SerCodigo, SerModificacion FROM Servicio INNER JOIN ServicioVisita ON SerCodigo = VisServicio And VisFImpresion IS NOT Null And VisSinEfecto = 0" _
            & " Where SerCodigo = " & Cons & " AND SerEstadoServicio = " & EstadoS.Retiro _
            & " And VisTipo = " & TipoServicio.Retiro & " And VisCamion IS NOT NULL " _
            & " And VisFImpresion IS NOT Null And VisSinEfecto = 0"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            idServicio = RsAux("SerCodigo")
            fEdit = RsAux("SerModificacion")
        Else
            MsgBox "No hay un retiro para recepcionar para ese código.", vbExclamation, "ATENCIÓN"
        End If
        RsAux.Close
        If idServicio > 0 Then
            AccionGrabar idServicio, fEdit, False
        End If
    Else
        MsgBox "No ingresó un número.", vbCritical, "ATENCIÓN"
    End If
    Exit Sub
errBF:
    clsGeneral.OcurrioError "Error al intentar recepcionar el servicio.", Err.Description
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
Private Sub cCamion_Click()
    vsConsulta.Rows = 1
    chVista.Value = 0
    Me.Refresh
End Sub

Private Sub cCamion_Change()
    vsConsulta.Rows = 1
    chVista.Value = 0
    Me.Refresh
End Sub

Private Sub cCamion_GotFocus()
    With cCamion
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cCamion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub
Private Sub cCamion_LostFocus()
    cCamion.SelStart = 0
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
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    InicializoGrillas
    Cons = "Select CamCodigo, CamNombre From Camion Order by CamNombre"
    CargoCombo Cons, cCamion
    vsListado.Orientation = orPortrait: vsListado.PaperSize = 1
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
        .FormatString = "Selección|>Servicio|<Producto|<Dirección|"
        .ColWidth(2) = 3500: .ColWidth(3) = 3100: .ColWidth(4) = 10
        .ColDataType(0) = flexDTBoolean
        .Redraw = True
    End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyE: AccionConsultar True
            
            Case vbKeyP: IrAPagina vsListado, 1
            Case vbKeyA: IrAPagina vsListado, vsListado.PreviewPage - 1
            Case vbKeyS: IrAPagina vsListado, vsListado.PreviewPage + 1
            Case vbKeyU: IrAPagina vsListado, vsListado.PageCount
            
            Case vbKeyAdd: Zoom vsListado, vsListado.Zoom + 5
            Case vbKeySubtract: Zoom vsListado, vsListado.Zoom - 5
            
            Case vbKeyI: AccionImprimir True
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            
            Case vbKeyG: AccionGrabar 0
            
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

Private Sub AccionConsultar(sMsg As Boolean)
    On Error GoTo errConsultar
    If cCamion.ListIndex = -1 Then MsgBox "Seleccione el camión que trae los retiros.", vbExclamation, "ATENCIÓN": Exit Sub
    Screen.MousePointer = 11
    chVista.Value = 0
    vsConsulta.Rows = 1
    vsConsulta.Refresh
    vsConsulta.Redraw = False
    CargoServicios
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
    If sMsg And vsConsulta.Rows = 1 Then MsgBox "No hay datos para el filtro ingresado.", vbExclamation, "ATENCIÓN"
    Exit Sub
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True
End Sub

Private Sub Label2_Click()
    Foco cCamion
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

    With vsListado
        .Device = paICartaN
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    EncabezadoListado vsListado, "Recepción de Retiros, Camión: " & Trim(cCamion.Text), False
    vsListado.FileName = "Recepción de Retiros"
    vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
    vsListado.EndDoc
    
    If Imprimir Then
        With vsListado
            .Device = paICartaN
            .PaperBin = 2
            .PrintDoc
        End With
    End If
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
Dim aModificacion As String
    
    'Cargo Impresos.
    Cons = "Select * From Servicio, ServicioVisita, Producto, Articulo, Cliente " _
        & " Where SerEstadoServicio = " & EstadoS.Retiro _
        & " And SerLocalReparacion = " & paCodigoDeSucursal _
        & " And VisTipo = " & TipoServicio.Retiro & " And VisCamion = " & cCamion.ItemData(cCamion.ListIndex) _
        & " And VisFImpresion IS NOT Null And VisSinEfecto = 0" _
        & " And SerCodigo = VisServicio And SerProducto = ProCodigo And ProArticulo = ArtID And ProCliente = CliCodigo "
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsAux.EOF
    
        With vsConsulta
            .AddItem ""
            
            aModificacion = RsAux!SerModificacion: .Cell(flexcpData, .Rows - 1, 0) = aModificacion
            aModificacion = RsAux!VisFModificacion: .Cell(flexcpData, .Rows - 1, 1) = aModificacion
            
            .Cell(flexcpChecked, .Rows - 1, 0) = flexChecked
            .Cell(flexcpText, .Rows - 1, 1) = RsAux!SerCodigo
            .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ProCodigo, "(#,000)") & " " & Trim(RsAux!ArtNombre)
            If Not IsNull(RsAux!ProDireccion) Then
                .Cell(flexcpText, .Rows - 1, 3) = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!ProDireccion, False, True, True)
            Else
                .Cell(flexcpText, .Rows - 1, 3) = clsGeneral.ArmoDireccionEnTexto(cBase, RsAux!CliDireccion, False, True, True)
            End If
            
        End With
        RsAux.MoveNext
        
    Loop
    RsAux.Close
    
End Sub
Private Sub GraboDatos(ByVal idServicio As Long, ByVal fechaEdit As Date, ByVal forzarLocalReparacion As Boolean, ByVal user As Integer)
Dim idLRep As Long
    Cons = "Select * From Servicio Where SerCodigo = " & idServicio
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        'msg = "Otra terminal elimino el servicio = " & Val(.Cell(flexcpText, I, 1))
        RsAux.Close: RsAux.Edit 'Provoco error.
    Else
        If RsAux!SerModificacion = fechaEdit Then
            RsAux.Edit
            RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
            RsAux!SerEstadoServicio = EstadoS.Taller
            idLRep = RsAux("SerLocalReparacion")
            If forzarLocalReparacion Then
                RsAux("SerLocalReparacion") = paCodigoDeSucursal
            End If
            'le agrego el local de ingreso de forma que se pueda trasladar.
            RsAux!SerLocalIngreso = paCodigoDeSucursal
            RsAux.Update
            RsAux.Close
        Else
            'msg = "Otra terminal modifico el servicio = " & Val(.Cell(flexcpText, I, 1))
            RsAux.Close: RsAux.Edit 'Provoco error.
        End If
    End If
    
    If idLRep = paCodigoDeSucursal Or forzarLocalReparacion Then
        Cons = "Insert Into Taller(TalServicio, TalFIngresoRealizado, TalFIngresoRecepcion,TalModificacion, TalUsuario) Values (" _
            & idServicio & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
            & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', " & user & ")"
        cBase.Execute (Cons)
    End If
    
End Sub

Private Sub AccionGrabar(ByVal ServicioAMiLocal As Long, Optional fEdit As Date, Optional CambiarAMiLocal As Boolean)
Dim msg As String, Usuario As String
    
    If ServicioAMiLocal = 0 Then
        If vsConsulta.Rows = 1 Then MsgBox "No hay datos en la consulta.", vbExclamation, "ATENCIÓN": Exit Sub
        If cCamion.ListIndex = -1 Then MsgBox "Seleccione el camión que traslada los artículos.", vbInformation, "ATENCIÓN": Foco cCamion: Exit Sub
        If MsgBox("¿Confirma recepcionar los retiros seleccionados?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    Else
        If MsgBox("¿Confirma recepcionar el retiro del servicio " & ServicioAMiLocal & " ?", vbQuestion + vbYesNo, "ATENCIÓN") = vbNo Then Exit Sub
    End If
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
    If ServicioAMiLocal = 0 Then
        With vsConsulta
            For I = 1 To .Rows - 1
                If .Cell(flexcpChecked, I, 0) = flexChecked Then
                    
                    msg = ""
                    GraboDatos Val(.Cell(flexcpText, I, 1)), CDate(.Cell(flexcpData, I, 0)), False, Usuario
    '                'TABLA SERVICIO
    '                Cons = "Select * From Servicio Where SerCodigo = " & Val(.Cell(flexcpText, I, 1))
    '                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    '                If RsAux.EOF Then
    '                    msg = "Otra terminal elimino el servicio = " & Val(.Cell(flexcpText, I, 1))
    '                    RsAux.Close: RsAux.Edit 'Provoco error.
    '                Else
    '                    If RsAux!SerModificacion = CDate(.Cell(flexcpData, I, 0)) Then
    '                        RsAux.Edit
    '                        RsAux!SerModificacion = Format(gFechaServidor, sqlFormatoFH)
    '                        RsAux!SerEstadoServicio = EstadoS.Taller
    '                        If ServicioAMiLocal Then
    '                            RsAux("SerLocalReparacion") = paCodigoDeSucursal
    '                        End If
    '                        RsAux.Update
    '                        RsAux.Close
    '                    Else
    '                        msg = "Otra terminal modifico el servicio = " & Val(.Cell(flexcpText, I, 1))
    '                        RsAux.Close: RsAux.Edit 'Provoco error.
    '                    End If
    '                End If
    '                Cons = "Insert Into Taller(TalServicio, TalFIngresoRealizado, TalFIngresoRecepcion,TalModificacion, TalUsuario) Values (" _
    '                    & Val(.Cell(flexcpText, I, 1)) & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', '" & Format(gFechaServidor, sqlFormatoFH) & "'" _
    '                    & ", '" & Format(gFechaServidor, sqlFormatoFH) & "', " & Usuario & ")"
    '                cBase.Execute (Cons)
                End If
            Next I
            
        End With
    Else
        GraboDatos ServicioAMiLocal, fEdit, CambiarAMiLocal, Usuario
    End If
    cBase.CommitTrans
    If ServicioAMiLocal = 0 Then
        If MsgBox("¿Desea imprimir los traslados?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
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
    End If
    AccionConsultar False
    Screen.MousePointer = 0
    Exit Sub
    
ErrBT:
    clsGeneral.OcurrioError "Ocurrió un error al intentar iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrRB:
    Resume ErrVA
ErrVA:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrió un error al intentar almacenar la información." & Chr(13) & msg, Err.Description
    Screen.MousePointer = 0
End Sub
