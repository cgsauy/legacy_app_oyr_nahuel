VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Begin VB.Form frmListado 
   BackColor       =   &H00C0C000&
   Caption         =   "Stock Total por Artículo de Existencia"
   ClientHeight    =   7530
   ClientLeft      =   420
   ClientTop       =   1785
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
   ScaleHeight     =   7530
   ScaleWidth      =   11880
   Begin VB.Frame labSeparador 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   5880
      Width           =   4575
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Distribución en Locales"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   3255
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLocal 
      Height          =   4335
      Left            =   5520
      TabIndex        =   3
      Top             =   1080
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   840
      TabIndex        =   2
      Top             =   1080
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
      TabIndex        =   18
      Top             =   720
      Width           =   11415
      _Version        =   196608
      _ExtentX        =   20135
      _ExtentY        =   7858
      _StockProps     =   229
      BackColor       =   -2147483633
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
      BackColor       =   -2147483633
   End
   Begin VB.PictureBox picBotones 
      BackColor       =   &H00C0C000&
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   6915
      TabIndex        =   19
      Top             =   6720
      Width           =   6975
      Begin VB.CommandButton bConexion 
         Caption         =   "CB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   6180
         Picture         =   "frmListado.frx":08CA
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Cambiar Base de Datos."
         Top             =   120
         Width           =   375
      End
      Begin VB.CheckBox chAgrupo 
         DownPicture     =   "frmListado.frx":09CC
         Height          =   310
         Left            =   4800
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmListado.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Abrir y cerrar datos. [Ctrl+O]"
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0BC0
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0CC2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Vista impresión o Grilla.[Ctrl+L]"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":11F4
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
         Picture         =   "frmListado.frx":166E
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
         Picture         =   "frmListado.frx":1758
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
         Picture         =   "frmListado.frx":1842
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
         Picture         =   "frmListado.frx":1A7C
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   5280
         Picture         =   "frmListado.frx":1B7E
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5760
         Picture         =   "frmListado.frx":1F44
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
         Picture         =   "frmListado.frx":2046
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
         Picture         =   "frmListado.frx":2348
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
         Picture         =   "frmListado.frx":268A
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
         Picture         =   "frmListado.frx":298C
         Style           =   1  'Graphical
         TabIndex        =   5
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
      TabIndex        =   17
      Top             =   7275
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   2805
            MinWidth        =   2805
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   15055
            Key             =   "msg"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      BackColor       =   &H00C0C000&
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
      TabIndex        =   16
      Top             =   0
      Width           =   10335
      Begin VB.TextBox tArticulo 
         Height          =   285
         Left            =   960
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   300
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private aTexto As String
Private bCargarImpresion As Boolean

Private Sub AccionLimpiar()
    tArticulo.Text = "": tArticulo.Tag = "0"
    vsConsulta.Rows = 1: vsLocal.Rows = 1
End Sub
Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConexion_Click()
Dim newB As String
    
    On Error GoTo errCh
    
    If Not miconexion.AccesoAlMenu("Cambiar_Conexion") Then Exit Sub
    
    newB = InputBox("Ingrese el texto del login para la nueva conexión" & vbCrLf & _
                "Id de aplicación en archivo de conexiones.", "Cambio de Base de Datos")
    
    If Trim(newB) = "" Then Exit Sub
    If MsgBox("Está seguro de cambiar la base de datos al login " & newB, vbQuestion + vbYesNo + vbDefaultButton2, "Cambiar Base") = vbNo Then Exit Sub
    
    'Limpio la ficha
    AccionLimpiar
    
    newB = miconexion.TextoConexion(newB)
    If Trim(newB) = "" Then Exit Sub
    
    Screen.MousePointer = 11
    On Error Resume Next
    cBase.Close
    On Error GoTo errCh
    Set cBase = Nothing
    Set cBase = eBase.OpenConnection("", rdDriverNoPrompt, , newB)
    
    Status.Panels("bd").Text = "Base de Datos: " & miconexion.RetornoPropiedad(bdb:=True)
    
    Screen.MousePointer = 0
    MsgBox "Ahora está trabajanbo en la nueva base de datos.", vbExclamation, "Base Cambiada OK"
    Exit Sub
    
errCh:
    clsGeneral.OcurrioError "Error de Conexión." & vbCrLf & " La conexión está en estado de error, conectese a una base de datos.", Err.Description
    Screen.MousePointer = 0
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

Private Sub cGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub chAgrupo_Click()
    If chAgrupo.Value = 0 Then
        'Abierto
        AgrupoCamposEnGrilla False
    Else
        'Cerrado
        AgrupoCamposEnGrilla True
    End If
End Sub

Private Sub chVista_Click()
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
        vsLocal.ZOrder 0
        labSeparador.ZOrder 0
    Else
        AccionImprimir
        vsListado.ZOrder 0
    End If
    Me.Refresh
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
    labSeparador.ZOrder 0
    vsListado.Orientation = orPortrait
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub
Private Sub InicializoGrillas()
    On Error Resume Next
    With vsConsulta
        .Redraw = False
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        .WordWrap = False
        .Cols = 1: .Rows = 1
        .FormatString = "<Artículo|<Estado|>Disponible|>No Disponible|>Total|"
        .ColWidth(0) = 3100: .ColWidth(1) = 1700: .ColWidth(3) = 1100: .ColWidth(4) = 1000: .ColWidth(5) = 15
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(0) = True
        .Redraw = True
    End With
    With vsLocal
        .Redraw = False
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        .WordWrap = False
        .Cols = 1: .Rows = 1
        .FormatString = "<Local|<Estado|>Cantidad|"
        .ColWidth(0) = 3100: .ColWidth(1) = 1700: .ColWidth(2) = 1000: .ColWidth(3) = 15
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(0) = True
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
            
            Case vbKeyQ: AccionLimpiar
            Case vbKeyI: AccionImprimir True
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            Case vbKeyO: If chAgrupo.Value = 0 Then chAgrupo.Value = 1 Else chAgrupo.Value = 0
            
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
    vsConsulta.Height = (vsListado.Height / 2) - (labSeparador.Height * 2)
    vsConsulta.Left = vsListado.Left
    labSeparador.Top = vsConsulta.Top + vsConsulta.Height
    labSeparador.Left = vsConsulta.Left
    labSeparador.Width = vsConsulta.Width
    vsLocal.Left = vsConsulta.Left
    vsLocal.Height = vsListado.Height - (vsConsulta.Height + labSeparador.Height)
    vsLocal.Width = vsListado.Width
    vsLocal.Top = labSeparador.Top + labSeparador.Height
    Me.Refresh
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    Set miconexion = Nothing
    End
    
End Sub

Private Sub AccionConsultar()
    On Error GoTo errConsultar
    Screen.MousePointer = 11
    bCargarImpresion = True
    vsConsulta.Rows = 1: vsLocal.Rows = 1
    CargoStock
    Foco tArticulo
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Sub Label3_Click()
    Foco tArticulo
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = "0"
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.Panels(3).Text = "Ingrese el artículo a consultar."
End Sub
Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrAP
    Screen.MousePointer = 11
    If KeyAscii = vbKeyReturn Then
        If Trim(tArticulo.Text) <> "" Then
            If IsNumeric(tArticulo.Text) Then
                BuscoArticuloPorCodigo tArticulo.Text
            Else
                BuscoArticuloPorNombre tArticulo.Text
            End If
            If Val(tArticulo.Tag) > 0 Then
                vsConsulta.Rows = 1: vsLocal.Rows = 1
                bConsultar.SetFocus
            End If
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrAP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub tArticulo_LostFocus()
    Ayuda ""
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
        
        EncabezadoListado vsListado, "Stock Total de Existencia al " & Format(Date, FormatoFP), False
        vsListado.FileName = "Consulta de Stock de Existencia"
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        vsListado.Paragraph = ""
        vsListado.FontBold = True
        vsListado.Paragraph = "Distribución en Locales"
        vsListado.FontBold = False
        vsListado.Paragraph = ""
        vsLocal.ExtendLastCol = False: vsListado.RenderControl = vsLocal.hwnd: vsLocal.ExtendLastCol = True
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
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub

Private Sub AccionConfigurar()
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
End Sub

Private Sub Ayuda(strTexto As String)
    Status.Panels("msg").Text = strTexto
End Sub
Private Sub CargoStock()
Dim Rs As rdoResultset
Dim CantExtra As Currency
On Error GoTo ErrCS
    If tArticulo.Tag = "0" Then Exit Sub
    
    Screen.MousePointer = 11
    
    'cons = "Select BMRCantidad = Sum(BMRCantidad), BMREstado, EsMAbreviacion, ArtID,  ArtCodigo, ArtNombre, EsMBajaStockTotal  From BMLocal, BMRenglon, Articulo,EstadoMercaderia " _
        & " Where BMRArticulo = " & tArticulo.Tag _
        & " And BMLCodigo =  1" _
        & " And BMLId = BMRIdBML And BMRArticulo = ArtID And BMREstado = EsMCodigo " _
        & " Group by BMREstado, EsMAbreviacion, ArtID, ArtCodigo, ArtNombre, EsMBajaStockTotal "

    cons = "Select Sum(BMRCantidad) as Q, BMREstado, EsMAbreviacion, ArtID,  ArtCodigo, ArtNombre, EsMBajaStockTotal  " _
        & " From BMLocal, BMRenglon, Articulo,EstadoMercaderia, GrupoUnoConteoBalance " _
        & " Where BMRArticulo = " & tArticulo.Tag _
        & " And BMLId = BMRIdBML And BMRArticulo = ArtID And BMREstado = EsMCodigo " _
        & " And BMLLocal = GUnSucursal And BMLArea = GUnQue And BMLCodigo=GUnGrupo" _
        & " Group by BMREstado, EsMAbreviacion, ArtID, ArtCodigo, ArtNombre, EsMBajaStockTotal "

        
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurReadOnly)
    If rsAux.EOF Then
        rsAux.Close
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCIÓN"
    Else
        Do While Not rsAux.EOF
            CantExtra = 0
            If rsAux!EsMBajaStockTotal = 0 Then
                CantExtra = CantidadNoDisponible(rsAux!ArtID, rsAux!BMREstado)
                If rsAux!Q <> 0 Or CantExtra <> 0 Then InsertoFila vsConsulta, Format(rsAux!ArtCodigo, "#,000,000") & " " & Trim(rsAux!ArtNombre), Trim(rsAux!EsMAbreviacion), rsAux!Q - CantExtra, CantExtra
            Else
                InsertoFila vsConsulta, Format(rsAux!ArtCodigo, "#,000,000") & " " & Trim(rsAux!ArtNombre), Trim(rsAux!EsMAbreviacion), 0, rsAux!BMRCantidad
            End If
            rsAux.MoveNext
        Loop
        rsAux.Close
        
        'cons = "Select BMRCantidad, BMREstado, EsMAbreviacion, LocNombre From BMLocal, BMRenglon, EstadoMercaderia, Local " _
            & " Where BMRArticulo = " & tArticulo.Tag & " And BMLCodigo = 1 " _
            & " And BMLId = BMRIdBML And BMREstado = EsMCodigo And BMLLocal = LocCodigo " _
            & " Order by LocNombre "
            
        cons = "Select BMRCantidad, BMREstado, EsMAbreviacion, LocNombre From BMLocal, BMRenglon, EstadoMercaderia, Local, GrupoUnoConteoBalance " _
            & " Where BMRArticulo = " & tArticulo.Tag _
            & " And BMLId = BMRIdBML And BMREstado = EsMCodigo And BMLLocal = LocCodigo " _
            & " And BMLLocal = GUnSucursal And BMLArea like GUnQue And BMLCodigo=GUnGrupo" _
            & " Order by LocNombre "
        
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurReadOnly)
        Do While Not rsAux.EOF
            InsertoFila vsLocal, rsAux!LocNombre, rsAux!EsMAbreviacion, rsAux!BMRCantidad, 0, True
            rsAux.MoveNext
        Loop
        rsAux.Close
        
        With vsConsulta
            .Subtotal flexSTSum, -1, 2, "#,##0", Obligatorio, Rojo, True, "Total"
            .Subtotal flexSTSum, -1, 3, "#,##0", Obligatorio, Rojo, True, ""
            .Subtotal flexSTSum, -1, 4, "#,##0", Obligatorio, Rojo, True, ""
        End With
        With vsLocal
            If .Rows > 1 Then .Select 1, 0, 1, 1
            .Sort = flexSortGenericAscending
            .Subtotal flexSTSum, 0, 2, "#,##0", Inactivo, Rojo, False, "%s"
            .Subtotal flexSTSum, -1, 2, "#,##0", Obligatorio, Rojo, True, "Total"
        End With
        If chAgrupo.Value = 1 Then AgrupoCamposEnGrilla True
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrCS:
    clsGeneral.OcurrioError "Ocurrio un error al cargar el stock.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function CantidadNoDisponible(IdArticulo As Long, IdEstado As Integer) As Currency
Dim RsStL As rdoResultset
    cons = "Select Sum(BMRCantidad) From BMLocal, BMRenglon " _
        & " Where BMRArticulo = " & IdArticulo _
        & " And BMREstado = " & IdEstado _
        & " And BMLId = BMRIdBML And BMLLocal IN (Select SucCodigo From Sucursal Where SucExtras = 1)"
    Set RsStL = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not IsNull(RsStL(0)) Then CantidadNoDisponible = RsStL(0) Else CantidadNoDisponible = 0
    RsStL.Close
End Function
Private Sub InsertoFila(Grilla As vsFlexGrid, CampoCero As String, Estado As String, QDisponible As Currency, QNoDisponible As Currency, Optional EstFisico As Boolean = True)
    With Grilla
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = CampoCero
        .Cell(flexcpText, .Rows - 1, 1) = Trim(Estado)
        .Cell(flexcpText, .Rows - 1, 2) = Format(QDisponible, "#,##0")
        If .Cols > 4 Then
            .Cell(flexcpText, .Rows - 1, 3) = Format(QNoDisponible, "#,##0")
            .Cell(flexcpText, .Rows - 1, 4) = Format(QNoDisponible + QDisponible, "#,##0")
            .Cell(flexcpFontBold, .Rows - 1, 4) = True
        Else
            .Cell(flexcpFontBold, .Rows - 1, 3) = True
        End If
        If Not EstFisico Then .Cell(flexcpForeColor, .Rows - 1, 1) = vbHighlight
    End With
End Sub

Private Sub BuscoArticuloPorCodigo(CodArticulo As Long)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
    
    Screen.MousePointer = 11
    cons = "Select * From Articulo Where ArtCodigo = " & CodArticulo
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurReadOnly)
    
    If rsAux.EOF Then
        rsAux.Close
        tArticulo.Tag = "0"
        MsgBox "No existe un artículo que posea ese código.", vbExclamation, "ATENCIÓN"
    Else
        tArticulo.Text = Format(rsAux!ArtCodigo, "#,000,000") & " " & Trim(rsAux!ArtNombre)
        tArticulo.Tag = rsAux!ArtID
        rsAux.Close
    End If
    Screen.MousePointer = 0

End Sub

Private Sub BuscoArticuloPorNombre(NomArticulo As String)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
Dim Resultado As Long

    Screen.MousePointer = 11
    
    cons = "Select ArtId, Código = ArtCodigo, Nombre = ArtNombre from Articulo" _
        & " Where ArtNombre LIKE '" & NomArticulo & "%'" _
        & " Order By ArtNombre"
            
    Dim LiAyuda As New clsListadeAyuda
    LiAyuda.ActivoListaAyuda cons, False, cBase.Connect
    Screen.MousePointer = 11
    If LiAyuda.ItemSeleccionado <> "" Then
        Resultado = LiAyuda.ItemSeleccionado
    Else
        Resultado = 0
    End If
    If Resultado > 0 Then BuscoArticuloPorCodigo Resultado
    Set LiAyuda = Nothing       'Destruyo la clase.
    Screen.MousePointer = 0
    
End Sub

Private Sub AgrupoCamposEnGrilla(Cierro As Boolean)
On Error GoTo ErrACEG
    With vsLocal
        .Redraw = False
        For I = 1 To .Rows - 1
            If Cierro Then
                If .IsSubtotal(I) And .RowOutlineLevel(I) = 0 Then .IsCollapsed(I) = flexOutlineCollapsed
            Else
                If .IsSubtotal(I) And .RowOutlineLevel(I) = 0 Then .IsCollapsed(I) = flexOutlineExpanded
            End If
        Next I
        .Redraw = True
    End With
    
    Exit Sub
ErrACEG:
    vsConsulta.Redraw = True
End Sub

