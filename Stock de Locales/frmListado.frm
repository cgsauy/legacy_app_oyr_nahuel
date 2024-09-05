VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmListado 
   Caption         =   "Consulta de Stock en Locales"
   ClientHeight    =   7530
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
   ScaleHeight     =   7530
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   180
      TabIndex        =   12
      Top             =   1200
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
      Left            =   60
      TabIndex        =   28
      Top             =   1200
      Width           =   10575
      _Version        =   196608
      _ExtentX        =   18653
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
      ScaleWidth      =   6675
      TabIndex        =   29
      Top             =   6720
      Width           =   6735
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
         Left            =   6300
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Cambiar base de datos."
         Top             =   120
         Width           =   375
      End
      Begin VB.CheckBox chAgrupo 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4800
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmListado.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Abrir y cerrar datos. [Ctrl+O]"
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0636
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0738
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Vista. [Ctrl+L]"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0C6A
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmListado.frx":10E4
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmListado.frx":11CE
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmListado.frx":12B8
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "frmListado.frx":14F2
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   5160
         Picture         =   "frmListado.frx":15F4
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5880
         Picture         =   "frmListado.frx":19BA
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":1ABC
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmListado.frx":1DBE
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
         Picture         =   "frmListado.frx":2100
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmListado.frx":2402
         Style           =   1  'Graphical
         TabIndex        =   14
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
      TabIndex        =   27
      Top             =   7275
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
            Object.Width           =   2805
            MinWidth        =   2805
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12488
            Key             =   "msg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Key             =   "fecha"
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
      Left            =   60
      TabIndex        =   26
      Top             =   0
      Width           =   11175
      Begin AACombo99.AACombo cEstadoArticulo 
         Height          =   315
         Left            =   4140
         TabIndex        =   8
         Top             =   600
         Width           =   1035
         _ExtentX        =   1826
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
      Begin VB.CheckBox chPorArticulo 
         Caption         =   "Orden &Por Artículos"
         Height          =   255
         Left            =   7620
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox tArticulo 
         Height          =   285
         Left            =   840
         MaxLength       =   60
         TabIndex        =   7
         Top             =   600
         Width           =   3255
      End
      Begin AACombo99.AACombo cTipo 
         Height          =   315
         Left            =   840
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
      Begin AACombo99.AACombo cGrupo 
         Height          =   315
         Left            =   5820
         TabIndex        =   5
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
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
      Begin AACombo99.AACombo cMarca 
         Height          =   315
         Left            =   3360
         TabIndex        =   3
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
      Begin AACombo99.AACombo cLocal 
         Height          =   315
         Left            =   5820
         TabIndex        =   10
         Top             =   600
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Local:"
         Height          =   255
         Left            =   5280
         TabIndex        =   9
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Marca:"
         Height          =   255
         Left            =   2820
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Grupo:"
         Height          =   255
         Left            =   5280
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Rs1 As rdoResultset
Private aTexto As String
Private bCargarImpresion As Boolean

Private Sub AccionLimpiar()
    cTipo.Text = ""
    cGrupo.Text = ""
    cMarca.Text = ""
    cEstadoArticulo.Text = ""
    tArticulo.Text = "": tArticulo.Tag = "0"
    cLocal.Text = ""
    chPorArticulo.Value = 0
    vsConsulta.Rows = 1
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
    cTipo.Clear: cGrupo.Clear: cMarca.Clear
    
    newB = miconexion.TextoConexion(newB)
    If Trim(newB) = "" Then Exit Sub
    
    Screen.MousePointer = 11
    On Error Resume Next
    cBase.Close
    On Error GoTo errCh
    Set cBase = Nothing
    Set cBase = eBase.OpenConnection("", rdDriverNoPrompt, , newB)
    
    Status.Panels("bd").Text = "Base de Datos: " & miconexion.RetornoPropiedad(bdb:=True)
    CargoDatos
    
    Screen.MousePointer = 0
    MsgBox "Ahora está trabajanbo en la nueva base de datos.", vbExclamation, "Base Cambiada OK"
    Exit Sub
    
errCh:
    clsGeneral.OcurrioError "Error de Conexión", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub
Private Sub bImprimir_Click()
    AccionImprimir True
End Sub
Private Sub bNoFiltros_Click()
    cTipo.Text = ""
    cGrupo.Text = ""
    cMarca.Text = ""
    cEstadoArticulo.Text = ""
    tArticulo.Text = "": tArticulo.Tag = "0"
    cLocal.Text = ""
    chPorArticulo.Value = 0
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

Private Sub cEstadoArticulo_GotFocus()
    With cEstadoArticulo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Indique que estado desea consultar. [Blanco = Todos]"
End Sub
Private Sub cEstadoArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cLocal
End Sub
Private Sub cEstadoArticulo_LostFocus()
    Ayuda ""
End Sub

Private Sub cGrupo_GotFocus()
    With cGrupo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione un grupo de artículo a consultar."
End Sub
Private Sub cGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tArticulo
End Sub
Private Sub cGrupo_LostFocus()
    Ayuda ""
End Sub

Private Sub cLocal_GotFocus()
    With cLocal
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione un local a consultar. [Blanco = Todos]"
End Sub
Private Sub cLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chPorArticulo.SetFocus
End Sub
Private Sub cLocal_LostFocus()
    Ayuda ""
End Sub

Private Sub cMarca_GotFocus()
    With cMarca
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione una marca de artículo a consultar."
End Sub

Private Sub cMarca_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cGrupo
End Sub

Private Sub cMarca_LostFocus()
    Ayuda ""
End Sub

Private Sub cTipo_GotFocus()
    With cTipo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda " Seleccione el tipo de artículo a consultar"
End Sub

Private Sub cTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cMarca
End Sub

Private Sub cTipo_LostFocus()
    Ayuda ""
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

Private Sub chPorArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco bConsultar
End Sub

Private Sub chVista_Click()
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
    Else
        AccionImprimir
        vsListado.ZOrder 0
    End If
End Sub

Private Sub Label1_Click()
    Foco cTipo
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    InicializoGrillas chPorArticulo.Value
    AccionLimpiar
    CargoDatos
    cEstadoArticulo.Clear
    cEstadoArticulo.AddItem "En Desuso"
    cEstadoArticulo.ItemData(cEstadoArticulo.NewIndex) = 0
    cEstadoArticulo.AddItem "En Uso"
    cEstadoArticulo.ItemData(cEstadoArticulo.NewIndex) = 1
    
    bCargarImpresion = True
    vsListado.Orientation = orPortrait
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub
Private Sub InicializoGrillas(PorArticulo As Boolean)
    On Error Resume Next
    With vsConsulta
        .Redraw = False
        .OutlineBar = flexOutlineBarComplete
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        .WordWrap = False
        .Cols = 1: .Rows = 1
        If PorArticulo Then
            .FormatString = "<Artículo|<Local|<Estado|>Cantidad|Contenedores|"
            .ColWidth(0) = 3900: .ColWidth(1) = 1700: .ColWidth(2) = 1700: .ColWidth(3) = 1050: .ColWidth(4) = 1000
        Else
            .FormatString = "<Local|<Artículo|<Estado|>Cantidad|Contenedores|"
            .ColWidth(0) = 1700: .ColWidth(1) = 3900: .ColWidth(2) = 1050: .ColWidth(3) = 1050: .ColWidth(4) = 1000
        End If
        .MergeCells = flexMergeSpill
        .MergeCol(0) = True: .MergeCol(1) = True
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
    vsConsulta.Height = vsListado.Height
    vsConsulta.Left = vsListado.Left
    
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
Dim Rs As rdoResultset
    On Error GoTo errConsultar
    cBase.QueryTimeout = 90
    Screen.MousePointer = 11
    bCargarImpresion = True
    InicializoGrillas chPorArticulo.Value
    CargoStock
    cBase.QueryTimeout = 15
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Sub Label2_Click()
    Foco cGrupo
End Sub

Private Sub Label3_Click()
    Foco cMarca
End Sub

Private Sub Label4_Click()
    Foco cLocal
End Sub

Private Sub Label5_Click()
    Foco tArticulo
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = "0"
End Sub
Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrAP
    If KeyAscii = vbKeyReturn Then
        Screen.MousePointer = 11
        If Trim(tArticulo.Text) <> "" Then
            If Val(tArticulo.Tag) <> 0 Then Foco cEstadoArticulo: Screen.MousePointer = 0: Exit Sub
            If IsNumeric(tArticulo.Text) Then
                BuscoArticuloPorCodigo tArticulo.Text
            Else
                BuscoArticuloPorNombre tArticulo.Text
            End If
            If Val(tArticulo.Tag) > 0 Then Foco cEstadoArticulo
        Else
            Foco cEstadoArticulo
        End If
        Screen.MousePointer = 0
    End If
    Exit Sub
ErrAP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
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
        
        EncabezadoListado vsListado, "Consulta de Stock en Local al " & Format(Date, FormatoFP), False
        vsListado.FileName = "Consulta de Stock en Local"
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
Dim QAFacturarR As Currency
Dim QAFacturarE As Currency
Dim bConten As Boolean
On Error GoTo ErrCS
    Screen.MousePointer = 11
    ArmoConsultaTotal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCIÓN"
    Else
        Do While Not RsAux.EOF
            If paEstadoArticuloEntrega = RsAux!StLEstado Then bConten = True Else bConten = False
            If RsAux!StLCantidad <> 0 Then
                If chPorArticulo.Value = 0 Then
                    InsertoFila vsConsulta, RsAux!LocNombre, Format(RsAux!ArtCodigo, "#,000,000") & " " & Trim(RsAux!ArtNombre), Trim(RsAux!EsMAbreviacion), RsAux!StLCantidad, bConten
                Else
                    InsertoFila vsConsulta, Format(RsAux!ArtCodigo, "#,000,000") & " " & Trim(RsAux!ArtNombre), RsAux!LocNombre, Trim(RsAux!EsMAbreviacion), RsAux!StLCantidad, bConten
                End If
            End If
            RsAux.MoveNext
        Loop
        RsAux.Close
        With vsConsulta
            .Subtotal flexSTSum, 0, 3, , Inactivo, Rojo, False, "%s"
            .Subtotal flexSTSum, 0, 4, , Inactivo, Rojo, False, "%s"
            .Subtotal flexSTSum, -1, 3, , Obligatorio, Rojo, True, "Total"
            .Subtotal flexSTSum, -1, 4, , Obligatorio, Rojo, True, "Total"
        End With
        If chAgrupo.Value = 1 Then AgrupoCamposEnGrilla True
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrCS:
    clsGeneral.OcurrioError "Ocurrio un error al cargar el stock.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub ArmoConsultaTotal()
    
    'Saco todos los artículos y su stock Total.---------------------------
    'Hago una unión por tipo de estado de mercadería.
    Cons = "Select ArtID, ArtCodigo, ArtNombre,  StLEstado, StLCantidad, EsMAbreviacion, LocNombre, ArtVolumen From Articulo, StockLocal, EstadoMercaderia, Local " _
        & " Where ArtID = StLArticulo And StLEstado = EsMCodigo " _
        & " And StLLocal = LocCodigo "
        
    If cEstadoArticulo.ListIndex <> -1 Then Cons = Cons & " And ArtEnUso = " & cEstadoArticulo.ItemData(cEstadoArticulo.ListIndex)
        
    If cTipo.ListIndex > -1 Then Cons = Cons & " And ArtTipo = " & cTipo.ItemData(cTipo.ListIndex)
    If cMarca.ListIndex > -1 Then Cons = Cons & " And ArtMarca = " & cMarca.ItemData(cMarca.ListIndex)
    If cGrupo.ListIndex > -1 Then
        Cons = Cons & " And ArtID IN (" _
                & " Select AGrArticulo from ArticuloGrupo" _
                & " Where AGrGrupo = " & cGrupo.ItemData(cGrupo.ListIndex) & ")"
    End If
    
    If Val(tArticulo.Tag) > 0 Then Cons = Cons & " And StLArticulo = " & tArticulo.Tag
    If cLocal.ListIndex > -1 Then Cons = Cons & " And StLLocal = " & cLocal.ItemData(cLocal.ListIndex)
    
    
    If chPorArticulo.Value = 0 Then
        Cons = Cons & " Order by StLLocal,  ArtCodigo,  StLEstado"
    Else
        Cons = Cons & " Order by ArtCodigo,  StLLocal, StLEstado"
    End If
        
End Sub

Private Sub InsertoFila(Grilla As vsFlexGrid, CampoCero As String, CampoUno As String, Estado As String, QTotal As Currency, bContenedor As Boolean)
    With Grilla
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = CampoCero
        .Cell(flexcpText, .Rows - 1, 1) = CampoUno
        .Cell(flexcpText, .Rows - 1, 2) = Trim(Estado)
        .Cell(flexcpText, .Rows - 1, 3) = Format(QTotal, FormatoMonedaP)
        .Cell(flexcpFontBold, .Rows - 1, 3) = True
        If bContenedor Then
            If Not IsNull(RsAux!ArtVolumen) Then .Cell(flexcpText, .Rows - 1, 4) = Format((QTotal * RsAux!ArtVolumen) / 60000, "###0.00")
        End If
    End With
End Sub

Private Function StockAFacturarRetira(Articulo As Long) As Long
On Error GoTo ErrSAFR
    Cons = "Select Sum(RVTARetirar) From VentaTelefonica, RenglonVtaTelefonica " _
            & " Where VTeTipo = " & TipoDocumento.ContadoDomicilio _
            & " And VTeDocumento = Null " _
            & " And RVTArticulo = " & Articulo _
            & "And VTeCodigo = RVTVentaTelefonica"
    Set Rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If IsNull(Rs1(0)) Then StockAFacturarRetira = 0 Else StockAFacturarRetira = Rs1(0)
    Rs1.Close
    Exit Function
ErrSAFR:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el stock a facturar que se retira.", Err.Description
    StockAFacturarRetira = 0
End Function

Private Function StockAFacturarEnvia(Articulo As Long) As Long
On Error GoTo ErrSAFE
    Cons = "Select Sum(REvAEntregar) From Envio, RenglonEnvio" _
        & " Where EnvTipo = " & TipoEnvio.Cobranza _
        & " And EnvEstado NOT IN (" & EstadoEnvio.Anulado & " , " & EstadoEnvio.Entregado & " ," & EstadoEnvio.Rebotado & ")" _
        & " And REvArticulo = " & Articulo _
        & " And REvAEntregar > 0" _
        & "And EnvCodigo = REvEnvio"
    Set Rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If IsNull(Rs1(0)) Then StockAFacturarEnvia = 0 Else StockAFacturarEnvia = Rs1(0)
    Rs1.Close
    Exit Function
ErrSAFE:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el stock a facturar que se envía.", Err.Description
    StockAFacturarEnvia = 0
End Function

Private Sub AgrupoCamposEnGrilla(Cierro As Boolean)
On Error GoTo ErrACEG
    With vsConsulta
        .Redraw = False
        For I = 1 To .Rows - 1
            'If .IsSubtotal(I) And .RowOutlineLevel(I) = 0 Then .IsCollapsed(I) = flexOutlineCollapsed
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


Private Sub BuscoArticuloPorCodigo(CodArticulo As Long)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
    
    Screen.MousePointer = 11
    Cons = "Select * From Articulo Where ArtCodigo = " & CodArticulo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If RsAux.EOF Then
        RsAux.Close
        tArticulo.Tag = "0"
        MsgBox "No existe un artículo que posea ese código.", vbExclamation, "ATENCIÓN"
    Else
        tArticulo.Text = Format(RsAux!ArtCodigo, "#,000,000") & " " & Trim(RsAux!ArtNombre)
        tArticulo.Tag = RsAux!ArtID
        RsAux.Close
    End If
    Screen.MousePointer = 0

End Sub

Private Sub BuscoArticuloPorNombre(NomArticulo As String)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
Dim Resultado As Long

    Screen.MousePointer = 11
    Cons = "Select Código = ArtCodigo, Nombre = ArtNombre from Articulo" _
        & " Where ArtNombre LIKE '" & Replace(NomArticulo, " ", "%") & "%'" _
        & " Order By ArtNombre"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No existe un nombre de artículo con esas características.", vbInformation, "ATENCIÓN"
    Else
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.MoveFirst
            Resultado = RsAux(0)
            RsAux.Close
        Else
            RsAux.Close
            Dim LiAyuda As New clsListadeAyuda
            If LiAyuda.ActivarAyuda(cBase, Cons, Titulo:="Buscar Artículo") > 0 Then
                Resultado = LiAyuda.RetornoDatoSeleccionado(0)
            Else
                Resultado = 0
            End If
            Set LiAyuda = Nothing       'Destruyo la clase.
        End If
        If Resultado > 0 Then BuscoArticuloPorCodigo Resultado
    End If
    Screen.MousePointer = 0
    
End Sub


Private Sub CargoDatos()
On Error Resume Next

    CargoParametro
    
    Cons = "Select TipCodigo, TipNombre From Tipo Order by TipNombre"
    CargoCombo Cons, cTipo
    Cons = "Select GruCodigo, GruNombre From Grupo Order by GruNombre"
    CargoCombo Cons, cGrupo
    Cons = "Select MarCodigo, MarNombre From Marca Order by MarNombre"
    CargoCombo Cons, cMarca
    Cons = "Select LocCodigo, LocNombre From Local Order by LocNombre"
    CargoCombo Cons, cLocal
    
End Sub

Private Sub CargoParametro()
On Error GoTo errCP
    Cons = "Select * from Parametro Where ParNombre like '%EstadoArticuloEntrega%'"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        paEstadoArticuloEntrega = RsAux!ParValor
    Else
        paEstadoArticuloEntrega = 0
    End If
    RsAux.Close
    Exit Sub
errCP:
    clsGeneral.OcurrioError "Ocurrió el siguiente error al cargar el parámetro estado Sano.", Err.Description
End Sub
