VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form CoBuscarGasto 
   Caption         =   "Consulta Buscar Gastos"
   ClientHeight    =   7575
   ClientLeft      =   1950
   ClientTop       =   2220
   ClientWidth     =   10050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CoBuscarGasto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   10050
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   3495
      Left            =   480
      TabIndex        =   26
      Top             =   2580
      Width           =   8895
      _Version        =   196608
      _ExtentX        =   15690
      _ExtentY        =   6165
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
      PreviewMode     =   1
      Zoom            =   70
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      BorderStyle     =   0  'None
      Height          =   425
      Left            =   420
      ScaleHeight     =   420
      ScaleWidth      =   8475
      TabIndex        =   25
      Top             =   6060
      Width           =   8475
      Begin VB.PictureBox picBotonesL 
         Height          =   375
         Left            =   720
         ScaleHeight     =   315
         ScaleWidth      =   1455
         TabIndex        =   31
         Top             =   30
         Width           =   1515
         Begin VB.CommandButton bPrimeroL 
            Height          =   310
            Left            =   0
            Picture         =   "CoBuscarGasto.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Ir a la primer página."
            Top             =   0
            Width           =   310
         End
         Begin VB.CommandButton bSiguienteL 
            Height          =   310
            Left            =   720
            Picture         =   "CoBuscarGasto.frx":067C
            Style           =   1  'Graphical
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Ir a la siguiente página."
            Top             =   0
            Width           =   310
         End
         Begin VB.CommandButton bAnteriorL 
            Height          =   310
            Left            =   360
            Picture         =   "CoBuscarGasto.frx":097E
            Style           =   1  'Graphical
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Ir a la página anterior."
            Top             =   0
            Width           =   310
         End
         Begin VB.CommandButton bUltimaL 
            Height          =   310
            Left            =   1080
            Picture         =   "CoBuscarGasto.frx":0CC0
            Style           =   1  'Graphical
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Ir a la última página."
            Top             =   0
            Width           =   310
         End
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2340
         Picture         =   "CoBuscarGasto.frx":0EFA
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2700
         Picture         =   "CoBuscarGasto.frx":0FE4
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "CoBuscarGasto.frx":10CE
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   50
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "CoBuscarGasto.frx":1548
         Height          =   310
         Left            =   4440
         Picture         =   "CoBuscarGasto.frx":164A
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "CoBuscarGasto.frx":1B7C
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "CoBuscarGasto.frx":1DB6
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "CoBuscarGasto.frx":20B8
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "CoBuscarGasto.frx":23FA
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   6060
         Picture         =   "CoBuscarGasto.frx":26FC
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Picture         =   "CoBuscarGasto.frx":27FE
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "CoBuscarGasto.frx":2BC4
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   50
         Width           =   310
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   5040
      TabIndex        =   15
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
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
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
      TabIndex        =   24
      Top             =   7320
      Width           =   10050
      _ExtentX        =   17727
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
            Object.Width           =   9525
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
      Height          =   1335
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   9135
      Begin VB.TextBox tSerie 
         Height          =   315
         Left            =   6840
         MaxLength       =   2
         TabIndex        =   9
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox tComprobante 
         Height          =   315
         Left            =   7320
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin AACombo99.AACombo cComprobante 
         Height          =   315
         Left            =   5040
         TabIndex        =   8
         Top             =   600
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
         Height          =   285
         Left            =   5040
         TabIndex        =   14
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox tImporte 
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   600
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
         Text            =   ""
      End
      Begin AACombo99.AACombo cTipo 
         Height          =   315
         Left            =   5040
         TabIndex        =   3
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
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
      Begin AACombo99.AACombo cProveedor 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
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
      Begin AACombo99.AACombo cCarpeta 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label Label6 
         Caption         =   "Comprobante:"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Importe:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "&Carpeta:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   4440
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Menu MnuAccesos 
      Caption         =   "Accesos"
      Visible         =   0   'False
      Begin VB.Menu MnuverEmbarque 
         Caption         =   "&Ver Embarque"
      End
      Begin VB.Menu MnuVerGastos 
         Caption         =   "Ver &Gastos"
      End
      Begin VB.Menu MnuVerEvolucion 
         Caption         =   "Ver &Evolución del Costeo"
      End
   End
End
Attribute VB_Name = "CoBuscarGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public clsGeneral As New clsLibGeneral

Private Parametro As String

Const PorPantalla = 400
Private PosLista As Integer

Private RsConsulta As rdoResultset, RsVtas As rdoResultset

Private aFormato As String, aTituloTabla As String, aComentario As String, aFechas  As String
Private aTexto As String, aImporte As String
Private strMonDolar As String, strMonPesos As String

Private Sub AccionSiguiente()

    On Error GoTo ErrAA
    If Not bSiguiente.Enabled Then Exit Sub
    
    If Not RsConsulta.EOF Then
        bAnterior.Enabled = True: bPrimero.Enabled = True
        
        PosLista = PosLista + (vsConsulta.Rows - 1)
        CargoLista
    Else
        MsgBox "Se ha llegado al final de la consulta, no hay más datos a desplegar.", vbInformation, "ATENCIÓN"
    End If
    
    vsConsulta.SetFocus
    Exit Sub
    
ErrAA:
    clsGeneral.OcurrioError "Ocurrió un error inesperado. " & Err.Description
End Sub

Private Sub AccionAnterior()

Dim UltimaPosicion As Long

    On Error GoTo ErrAA
    If Not bAnterior.Enabled Then Exit Sub
    
    If RsConsulta.EOF And vsConsulta.Rows - 1 > 0 And RsConsulta.AbsolutePosition = -1 Then
        RelojA
        RsConsulta.MoveLast
        RelojD
        UltimaPosicion = PosLista + vsConsulta.Rows
        RsConsulta.MoveNext
        RelojD
    Else
        UltimaPosicion = PosLista + vsConsulta.Rows
    End If
    
    If UltimaPosicion - (vsConsulta.Rows - 1) - PorPantalla >= 1 Then
        
        If UltimaPosicion - (vsConsulta.Rows - 1) - PorPantalla = 1 Then bAnterior.Enabled = False: bPrimero.Enabled = False
        bSiguiente.Enabled = True
        
        RsConsulta.Move UltimaPosicion - (vsConsulta.Rows - 1) - PorPantalla, 1
        CargoLista
        PosLista = PosLista - (vsConsulta.Rows - 1)
        RelojD
    Else
        MsgBox "Se ha llegado al principio de la consulta.", vbInformation, "ATENCIÓN"
    End If
    vsConsulta.SetFocus
    Exit Sub
    
ErrAA:
    RelojD
    clsGeneral.OcurrioError "Ocurrió un error inesperado. " & Err.Description
End Sub

Private Sub AccionPrimero()
    
    If Not bPrimero.Enabled Then Exit Sub
    
    PosLista = 0
    RelojA
    On Error Resume Next
    RsConsulta.MoveFirst
    On Error GoTo ErrAA
    CargoLista
    vsConsulta.SetFocus
    bSiguiente.Enabled = True
    bPrimero.Enabled = False: bAnterior.Enabled = False
    RelojD
    Exit Sub
ErrAA:
    RelojD
    clsGeneral.OcurrioError "Ocurrió un error inesperado. " & Err.Description
End Sub
Private Sub AccionLimpiar()
    
    cTipo.Text = ""
    cProveedor.Text = ""
    cMoneda.Text = ""
    tImporte.Text = ""
    cComprobante.Text = ""
    tSerie.Text = ""
    tComprobante.Text = ""
    tFecha.Text = ""
    cCarpeta.Text = ""
    
End Sub
Private Sub bAnterior_Click()
    AccionAnterior
End Sub
Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConfigurar_Click()
    AccionConfigurar
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
    AccionPrimero
End Sub

Private Sub bPrimeroL_Click()
    IrAPagina vsListado, 1
End Sub

Private Sub bSiguienteL_Click()
    IrAPagina vsListado, vsListado.PreviewPage + 1
End Sub

Private Sub bAnteriorL_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub

Private Sub bUltimaL_Click()
    IrAPagina vsListado, vsListado.PageCount
End Sub

Private Sub bSiguiente_Click()
    AccionSiguiente
End Sub

Private Sub bZMas_Click()
    Zoom vsListado, vsListado.Zoom + 5
End Sub

Private Sub bZMenos_Click()
    Zoom vsListado, vsListado.Zoom - 5
End Sub

Private Sub cCarpeta_GotFocus()
    RelojA
    cCarpeta.SelStart = 0
    cCarpeta.SelLength = Len(cCarpeta.Text)
    CargoComboFolder
    RelojD
End Sub
Private Sub cCarpeta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco tFecha
End Sub
Private Sub cCarpeta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cCarpeta.ListCount = 0 Then CargoComboFolder
End Sub

Private Sub cComprobante_GotFocus()
    With cComprobante
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cComprobante_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tSerie
End Sub

Private Sub cComprobante_LostFocus()
    cComprobante.SelStart = 0
End Sub

Private Sub chVista_Click()
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
        picBotonesL.Visible = False
    Else
        AccionImprimir
        vsListado.ZOrder 0
        picBotonesL.Visible = True
    End If
End Sub

Private Sub cMoneda_GotFocus()
    With cMoneda
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tImporte.SetFocus
End Sub

Private Sub cMoneda_LostFocus()
    cMoneda.SelStart = 0
End Sub

Private Sub cProveedor_GotFocus()
    With cProveedor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    If cProveedor.ListCount = 0 Then
        RelojA
        Cons = "Select PClCodigo, PClFantasia From ProveedorCliente Order by PClFantasia"
        CargoCombo Cons, cProveedor
        RelojD
    End If
End Sub

Private Sub cProveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cTipo
End Sub

Private Sub cProveedor_LostFocus()
    cProveedor.SelStart = 0
End Sub

Private Sub cTipo_GotFocus()
    With cTipo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    RelojA
    If cTipo.ListCount = 0 Then
        Cons = "Select SRuID, SRuNombre From SubRubro" _
            & " Where SRuRubro = " & paRubroImportaciones _
            & "Order by SRuNombre"
        CargoCombo Cons, cTipo
    End If
    RelojD
End Sub

Private Sub cTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco cCarpeta
End Sub

Private Sub cTipo_LostFocus()
    cTipo.SelStart = 0
End Sub

Private Sub Label2_Click()
    Foco cMoneda
End Sub

Private Sub Label3_Click()
    Foco tFecha
End Sub

Private Sub Label4_Click()
    Foco cCarpeta
End Sub

Private Sub Label5_Click()
    Foco cProveedor
End Sub

Private Sub Form_Activate()
    RelojD
    Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad

    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    
    With vsListado
        .PaperSize = vbPRPSLetter
        .PhysicalPage = True
        .Orientation = orPortrait
        .Zoom = 100
    End With
    vsConsulta.ZOrder 0
    picBotonesL.Visible = False: picBotonesL.BorderStyle = 0
    
    bPrimero.Enabled = False: bSiguiente.Enabled = False: bAnterior.Enabled = False
    LimpioGrilla
    
    Cons = "Select * from Articulo Where ArtID = 0"
    Set RsConsulta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    CargoDocumentos
    strMonDolar = "": strMonPesos = ""
    
    Cons = "Select MonCodigo, MonSigno From Moneda " _
            & " Where MonCodigo IN (" & paMonedaPesos & ", " & paMonedaDolar & ")" _
            & " Order by MonSigno"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsAux.EOF
        cMoneda.AddItem RsAux!MonSigno
        cMoneda.ItemData(cMoneda.NewIndex) = RsAux!MonCodigo
        If RsAux!MonCodigo = paMonedaPesos Then strMonPesos = RsAux!MonSigno Else strMonDolar = RsAux!MonSigno
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    FechaDelServidor
    gFechaServidor = Format(gFechaServidor, FormatoFP)
    AccionLimpiar
    
    '---------------------------------------------------------------------------------------------------
    '  Para seleccionar una carpeta hay que pasar el CODIGO y no el ID.
    '---------------------------------------------------------------------------------------------------
    If Trim(Command()) <> "" Then
        CargoComboFolder
        Foco cCarpeta
        cCarpeta.Text = CStr(Command())
        If cCarpeta.ListIndex <> -1 Then AccionConsultar
    End If
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyE: AccionConsultar
            
            Case vbKeyP: AccionPrimero
            Case vbKeyA: AccionAnterior
            Case vbKeyS: AccionSiguiente
            
            Case vbKeyQ: AccionLimpiar
            Case vbKeyI: AccionImprimir
            
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    RelojA
    picBotones.BorderStyle = vbFlat
    picBotones.Top = Me.ScaleHeight - (picBotones.Height + Status.Height + 40)
    fFiltros.Width = Me.Width - (fFiltros.Left * 2.5)
    
    vsConsulta.Left = fFiltros.Left
    vsConsulta.Top = fFiltros.Top + fFiltros.Height + 50
    vsConsulta.Height = Me.ScaleHeight - (vsConsulta.Top + picBotones.Height + Status.Height + 90)
    vsConsulta.Width = fFiltros.Width
    
    vsListado.Top = vsConsulta.Top: vsListado.Left = vsConsulta.Left
    vsListado.Height = vsConsulta.Height: vsListado.Width = vsConsulta.Width
    
    RelojD
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    RsConsulta.Close
    CierroConexion
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub Label1_Click()
    Foco cTipo
End Sub

Private Sub AccionConsultar()
    
    If Not VerificoFiltros Then Exit Sub
    
    vsConsulta.Rows = 1
    
    'Cierro el cursor.---------------------------------
    On Error Resume Next: RsConsulta.Close
    
    On Error GoTo errConsultar
    
    RelojA
    PosLista = 0
    ArmoConsulta
    Set RsConsulta = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsConsulta.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        bPrimero.Enabled = False: bAnterior.Enabled = False: bSiguiente.Enabled = False
        RelojD
        Exit Sub
    End If

    CargoLista
    If RsConsulta.EOF Then bPrimero.Enabled = False: bAnterior.Enabled = False: bSiguiente.Enabled = False
    
    RelojD
    Exit Sub

errConsultar:
    RelojD
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Sub ArmoConsulta()
    '----------------------------------------------------------------------------------------------
    'La consulta esta realizada por tres uniones.
    'Se buscan los gastos para la Carpeta
    'Union
    'los gastos de los Embarques de esa carpeta
    'Union
    'Los gastos de las subcarpetas de la misma.
    '----------------------------------------------------------------------------------------------

    'Consulta para Carpetas.-------------------------------------------------------------
    Cons = "Select CarCodigo, EmbCodigo = '', SubCodigo = '', GImFolder, GImNivelFolder,  GImImporte, SRuNombre, ComMoneda, ComFecha, ComTipoDocumento, ComSerie, ComNumero, PClFantasia, ComTC, ComCodigo " _
        & " From Compra, GastoImportacion, SubRubro, ProveedorCliente, Carpeta " _
        & " Where GImNivelFolder = " & Folder.cFCarpeta & " And GImFolder = CarID "
    If cCarpeta.ListIndex > -1 Then Cons = Cons & " And GImFolder = " & cCarpeta.ItemData(cCarpeta.ListIndex)
    Cons = Cons & FiltrosComunes
    'Hago SubConsulta con los artículos de la carpeta.
    If cTipo.ListIndex > -1 Then Cons = Cons & " And GImIDSubRubro = " & cTipo.ItemData(cTipo.ListIndex)
    Cons = Cons & " And ComCodigo = GImIDCompra And PClCodigo = ComProveedor And GImIDSubRubro = SRuID"
    '----------------------------------------------------------------------------------------------
    Cons = Cons & " Union All "
    'Consulta para Embarques.-------------------------------------------------------------
    Cons = Cons & "Select CarCodigo, EmbCodigo, SubCodigo = '', GImFolder, GImNivelFolder,  GImImporte, SRuNombre, ComMoneda, ComFecha, ComTipoDocumento, ComSerie, ComNumero, PClFantasia, ComTC, ComCodigo  From Compra, GastoImportacion, SubRubro, ProveedorCliente, Carpeta, Embarque " _
        & " Where GImNivelFolder = " & Folder.cFEmbarque & " And GImFolder = EmbID And CarID = EmbCarpeta"
    If cCarpeta.ListIndex > -1 Then Cons = Cons & " And EmbCarpeta = " & cCarpeta.ItemData(cCarpeta.ListIndex)
    Cons = Cons & FiltrosComunes
    'Hago SubConsulta con los artículos de la carpeta.
    If cTipo.ListIndex > -1 Then Cons = Cons & " And GImIDSubRubro = " & cTipo.ItemData(cTipo.ListIndex)
    Cons = Cons & " And ComCodigo = GImIDCompra And PClCodigo = ComProveedor And GImIDSubRubro = SRuID"
    '----------------------------------------------------------------------------------------------
    Cons = Cons & " Union All "
    'Consulta para SubCarpetas.-------------------------------------------------------------
    Cons = Cons & "Select CarCodigo, EmbCodigo, SubCodigo ,GImFolder, GImNivelFolder,  GImImporte, SRuNombre, ComMoneda, ComFecha, ComTipoDocumento, ComSerie, ComNumero, PClFantasia, ComTC, ComCodigo  From Compra, GastoImportacion, SubRubro, ProveedorCliente, Carpeta, Embarque, SubCarpeta " _
        & " Where GImNivelFolder = " & Folder.cFSubCarpeta & " And GImFolder = SubID And CarID = EmbCarpeta And EmbID = SubEmbarque "
    If cCarpeta.ListIndex > -1 Then Cons = Cons & " And EmbCarpeta = " & cCarpeta.ItemData(cCarpeta.ListIndex)
    Cons = Cons & FiltrosComunes
    'Hago SubConsulta con los artículos de la carpeta.
    If cTipo.ListIndex > -1 Then Cons = Cons & " And GImIDSubRubro = " & cTipo.ItemData(cTipo.ListIndex)
    Cons = Cons & " And ComCodigo = GImIDCompra And PClCodigo = ComProveedor And GImIDSubRubro = SRuID"
    '----------------------------------------------------------------------------------------------
    Cons = Cons & " Order by CarCodigo, EmbCodigo, SubCodigo"
    
End Sub
Private Function FiltrosComunes() As String
Dim strCons As String
    strCons = ""
    If cProveedor.ListIndex > -1 Then strCons = strCons & " And ComProveedor = " & cProveedor.ItemData(cProveedor.ListIndex)
    If cComprobante.ListIndex > -1 Then
        strCons = strCons & " And ComTipoDocumento = " & cComprobante.ItemData(cComprobante.ListIndex)
        If Trim(tSerie.Text) <> "" Then strCons = strCons & " And ComSerie = '" & tSerie.Text & "'"
        strCons = strCons & " And ComNumero = " & tComprobante.Text
    End If
    'Fecha de ingresado el gasto.-------------------------------
    If aFechas <> "" Then strCons = strCons & ConsultaDeFecha(" And", "ComFecha", Trim(tFecha.Text))
    FiltrosComunes = strCons
End Function

Private Sub CargoLista()
On Error GoTo ErrInesperado
Dim strAux As String
Dim aImporte1 As Currency, aImporte2 As Currency
Dim sOk As Boolean
Dim aTotalPesos As Currency, aTotalDolar As Currency

    vsConsulta.Redraw = False
    vsConsulta.Rows = 1
    aTotalPesos = 0
    aTotalDolar = 0
    
    aImporte1 = 0
    aImporte2 = 9999999999999#
    'Controlo el filtro del importe
    If Trim(tImporte.Text) <> "" Then
        If Mid(aImporte, 1, 1) = ">" Or Mid(aImporte, 1, 1) = "<" Then
             If Mid(aImporte, 1, 1) = ">" Then
                aImporte1 = CCur(Mid(aImporte, 2, Len(aImporte)))
             Else
                aImporte2 = CCur(Mid(aImporte, 2, Len(aImporte)))
             End If
        Else
            aImporte1 = CCur(Mid(aImporte, 1, InStr(aImporte, "-") - 1))
            aImporte2 = CCur(Mid(aImporte, InStr(aImporte, "-") + 1, Len(aImporte)))
        End If
    End If
    
    RelojA
    Do While Not RsConsulta.EOF And vsConsulta.Rows - 1 < PorPantalla
        sOk = True
        If Trim(tImporte.Text) <> "" Then
            sOk = False
            If RsConsulta!ComMoneda = paMonedaPesos Then     'Gasto en pesos
                If cMoneda.ItemData(cMoneda.ListIndex) = paMonedaPesos Then
                    If RsConsulta!GImImporte > aImporte1 And RsConsulta!GImImporte < aImporte2 Then
                        sOk = True
                    End If
                Else
                    If (RsConsulta!GImImporte / RsConsulta!ComTC) > aImporte1 And (RsConsulta!GImImporte / RsConsulta!ComTC) < aImporte2 Then
                        sOk = True
                    End If
                End If
            Else                                            'Gasto en dolares
                If cMoneda.ItemData(cMoneda.ListIndex) = paMonedaDolar Then
                    If RsConsulta!GImImporte > aImporte1 And RsConsulta!GImImporte < aImporte2 Then sOk = True
                Else
                    If (RsConsulta!GImImporte * RsConsulta!ComTC) > aImporte1 And (RsConsulta!GImImporte * RsConsulta!ComTC) < aImporte2 Then sOk = True
                End If
            End If
        End If
        
        If sOk Then
        
            'Inserto en la grilla.------------------------------------------------
            vsConsulta.AddItem "", vsConsulta.Rows
            With vsConsulta
                .Cell(flexcpText, vsConsulta.Rows - 1, 0) = RsConsulta!GImFolder
                .Cell(flexcpText, vsConsulta.Rows - 1, 1) = RsConsulta!GImNivelFolder
                
                Select Case RsConsulta!GImNivelFolder
                    Case Folder.cFCarpeta: .Cell(flexcpText, vsConsulta.Rows - 1, 2) = RsConsulta!CarCodigo
                    Case Folder.cFEmbarque: .Cell(flexcpText, vsConsulta.Rows - 1, 2) = RsConsulta!CarCodigo & "." & Trim(RsConsulta!EmbCodigo)
                    Case Folder.cFSubCarpeta: .Cell(flexcpText, vsConsulta.Rows - 1, 2) = RsConsulta!CarCodigo & "." & Trim(RsConsulta!EmbCodigo) & "/" & RsConsulta!SubCodigo
                End Select
                
                .Cell(flexcpText, vsConsulta.Rows - 1, 3) = Format(RsConsulta!ComFecha, "dd/mm/yyyy")
                .Cell(flexcpText, vsConsulta.Rows - 1, 4) = Trim(RsConsulta!PClFantasia)
                strAux = RetornoNombreDocumento(RsConsulta!ComTipoDocumento, True) & ": "
                If Not IsNull(RsConsulta!ComSerie) Then strAux = strAux & Trim(RsConsulta!ComSerie)
                If Not IsNull(RsConsulta!ComNumero) Then strAux = strAux & RsConsulta!ComNumero
                .Cell(flexcpText, vsConsulta.Rows - 1, 5) = strAux
                
                If RsConsulta!ComMoneda = paMonedaPesos Then strAux = strMonPesos Else strAux = strMonDolar
                .Cell(flexcpText, vsConsulta.Rows - 1, 6) = strAux
                
                .Cell(flexcpText, vsConsulta.Rows - 1, 7) = Trim(RsConsulta!SRuNombre)
                If RsConsulta!ComMoneda = paMonedaPesos Then
                    .Cell(flexcpText, vsConsulta.Rows - 1, 8) = Format(RsConsulta!GImImporte, FormatoMonedaP)
                    aTotalPesos = aTotalPesos + RsConsulta!GImImporte
                    .Cell(flexcpText, vsConsulta.Rows - 1, 9) = Format(RsConsulta!GImImporte / RsConsulta!ComTC, FormatoMonedaP)
                    aTotalDolar = aTotalDolar + (RsConsulta!GImImporte / RsConsulta!ComTC)
                Else
                    aTotalPesos = aTotalPesos + (RsConsulta!GImImporte * RsConsulta!ComTC)
                    .Cell(flexcpText, vsConsulta.Rows - 1, 8) = Format(RsConsulta!GImImporte * RsConsulta!ComTC, FormatoMonedaP)
                    .Cell(flexcpText, vsConsulta.Rows - 1, 9) = Format(RsConsulta!GImImporte, FormatoMonedaP)
                    aTotalDolar = aTotalDolar + RsConsulta!GImImporte
                End If
                .Cell(flexcpText, vsConsulta.Rows - 1, 10) = Format(RsConsulta!ComTC, "#,##0.000")
                If Not IsNull(RsConsulta!ComCodigo) Then .Cell(flexcpText, vsConsulta.Rows - 1, 11) = RsConsulta!ComCodigo
            End With
        End If
        RsConsulta.MoveNext
    Loop
    
    vsConsulta.AddItem ""
    vsConsulta.Cell(flexcpBackColor, vsConsulta.Rows - 1, 0, vsConsulta.Rows - 1, 10) = Colores.Rojo
    vsConsulta.Cell(flexcpForeColor, vsConsulta.Rows - 1, 8, vsConsulta.Rows - 1, 9) = Colores.Blanco
    vsConsulta.Cell(flexcpText, vsConsulta.Rows - 1, 8) = Format(aTotalPesos, FormatoMonedaP)
    vsConsulta.Cell(flexcpText, vsConsulta.Rows - 1, 9) = Format(aTotalDolar, FormatoMonedaP)
    
    If RsConsulta.EOF Then bSiguiente.Enabled = False Else bSiguiente.Enabled = True
    vsConsulta.Redraw = True
    RelojD
    Exit Sub
    
ErrInesperado:
    vsConsulta.Redraw = True
    RelojD
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de stock.", Err.Description
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    If vsConsulta.Rows = 1 Then Exit Sub
    Screen.MousePointer = 11
    
    With vsListado
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    
    EncabezadoListado vsListado, "Importaciones - Consulta de Gastos", False
    vsListado.filename = "Consulta de Gastos"
    
    vsListado.Paragraph = "Filtros: " & ArmoFormulaFiltros
    
    vsConsulta.ExtendLastCol = False
    vsListado.RenderControl = vsConsulta.hWnd
    vsConsulta.ExtendLastCol = True
    
    vsListado.EndDoc
    
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

Private Function ArmoFormulaFiltros() As String
Dim aRetorno As String

    On Error Resume Next
    aRetorno = ""
    
    If cTipo.ListIndex <> -1 Then aRetorno = aRetorno & " Tipo: " & cTipo.Text & ", "
    
    If cProveedor.ListIndex <> -1 Then aRetorno = aRetorno & " Prov.: " & cProveedor.Text & ", "
    If Trim(tFecha.Text) <> "" Then aRetorno = aRetorno & "(" & Trim(tFecha.Text) & "), "
    
    aRetorno = Mid(aRetorno, 1, Len(aRetorno) - 2)
    ArmoFormulaFiltros = aRetorno

End Function

Private Function VerificoFiltros() As Boolean

    VerificoFiltros = False
    
    If cProveedor.Text <> "" And cProveedor.ListIndex = -1 Then
        MsgBox "El proveedor de artículos no es correcto.", vbExclamation, "ATENCIÓN"
        Foco cProveedor: Exit Function
    End If
    
    If cTipo.Text <> "" And cTipo.ListIndex = -1 Then
        MsgBox "El tipo de gasto seleccionado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco cTipo: Exit Function
    End If
    
    If cMoneda.Text <> "" And cMoneda.ListIndex = -1 Then
        MsgBox "La moneda no es correcta.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
    
    If cMoneda.ListIndex = -1 And Trim(tImporte.Text) <> "" Or cMoneda.ListIndex > -1 And Trim(tImporte.Text) = "" Then
        MsgBox "Los valores ingresados no son correctos.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
    
    aImporte = ""
    If Trim(tImporte.Text) <> "" Then
        aImporte = ValidoImporte(tImporte.Text)
        If aImporte = "" Then
            MsgBox "El importe ingresado para aplicar al filtro no es correcto." & Chr(13) _
                & " > xxx, < xxx ó ExxYzz", vbExclamation, "ATENCIÓN"
            Foco tImporte: Exit Function
        End If
    End If
    
    If cComprobante.Text <> "" And cComprobante.ListIndex = -1 Then
        MsgBox "El comprobante seleccionado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco cComprobante: Exit Function
    End If
    If cComprobante.ListIndex = -1 And Trim(tComprobante.Text) <> "" Or cComprobante.ListIndex > -1 And Trim(tComprobante.Text) = "" Then
        MsgBox "Los valores ingresados no son correctos.", vbExclamation, "ATENCIÓN"
        Foco cComprobante: Exit Function
    End If
    
    If cCarpeta.Text <> "" And cCarpeta.ListIndex = -1 Then
        MsgBox "La carpeta ingresada no es correcta.", vbExclamation, "ATENCIÓN"
        Foco cCarpeta: Exit Function
    End If
    
    aFechas = ""
    If Trim(tFecha.Text) <> "" Then
        aFechas = ValidoPeriodoFechas(Trim(tFecha.Text))
        If aFechas = "" Then
            MsgBox "Hay errores en el formato de fechas ingresado, o no es correcto.", vbExclamation, "ATENCIÓN"
            Foco tFecha: Exit Function
        End If
        aFechas = Trim(tFecha.Text)
    End If
    VerificoFiltros = True

End Function

Private Sub Label6_Click()
    Foco cComprobante
End Sub
Private Sub MnuverEmbarque_Click()
On Error GoTo errApp
    Dim RetVal
    Screen.MousePointer = 11
    Parametro = vsConsulta.Cell(flexcpText, vsConsulta.Row, 0)
    RetVal = Shell(App.Path & "\Embarque " & Parametro, 1)
    Screen.MousePointer = 0
    Exit Sub
errApp:
    clsGeneral.OcurrioError "Ocurrió un error al ejecutar la aplicación EMBARQUE. ", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub MnuVerEvolucion_Click()
On Error GoTo errApp
    Dim RetVal
    Screen.MousePointer = 11
    Parametro = vsConsulta.Cell(flexcpText, vsConsulta.Row, 1)
    Parametro = Parametro & vsConsulta.Cell(flexcpText, vsConsulta.Row, 0)
    RetVal = Shell(App.Path & "\Costeo de Carpetas " & Parametro, 1)
    Screen.MousePointer = 0
    Exit Sub
errApp:
    clsGeneral.OcurrioError "Ocurrió un error al ejecutar la aplicación Evolución del costeo. ", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub MnuVerGastos_Click()
On Error GoTo errApp
    Dim RetVal
    Screen.MousePointer = 11
    Parametro = vsConsulta.Cell(flexcpText, vsConsulta.Row, 11)
    RetVal = Shell(App.Path & "\Ingreso de Gastos " & Parametro, 1)
    Screen.MousePointer = 0
    Exit Sub
errApp:
    clsGeneral.OcurrioError "Ocurrió un error al ejecutar la aplicación Gastos. ", Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub tComprobante_GotFocus()
    With tComprobante
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub tComprobante_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cCarpeta
End Sub
Private Sub tFecha_GotFocus()
    With tFecha
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub
Private Sub tImporte_GotFocus()
    With tImporte
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub tImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cComprobante
End Sub
Private Sub tImporte_LostFocus()
    If IsNumeric(tImporte.Text) Then tImporte.Text = Format(tImporte.Text, FormatoMonedaP)
End Sub
Private Sub tSerie_GotFocus()
    With tSerie
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub tSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tComprobante
End Sub
Private Sub vsConsulta_Click()
    If vsConsulta.MouseRow = 0 Then
        vsConsulta.ColSel = vsConsulta.MouseCol
        If vsConsulta.ColSort(vsConsulta.MouseCol) = flexSortGenericAscending Then
            vsConsulta.ColSort(vsConsulta.MouseCol) = flexSortGenericDescending
        Else
            vsConsulta.ColSort(vsConsulta.MouseCol) = flexSortGenericAscending
        End If
        vsConsulta.Sort = flexSortUseColSort
    End If
End Sub

Private Sub vsConsulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And vsConsulta.Rows > 1 Then PopupMenu MnuAccesos, X:=X + vsConsulta.Left, Y:=Y + vsConsulta.Top
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub
Private Sub LimpioGrilla()

    With vsConsulta
        .Redraw = False
        .ExtendLastCol = True
        .Clear
        .Editable = False
        .Rows = 1
        .Cols = 12
        .FormatString = "ID|Tipo|Folder|<Fecha|Proveedor|Comprobante|Moneda|Sub Rubro|>Pesos|>Dolares|>T.C.|IDCompra"
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 800
        .ColWidth(3) = 1050
        .ColWidth(4) = 1800
        .ColWidth(5) = 1800
        .ColWidth(6) = 650
        .ColWidth(7) = 1650
        .ColWidth(8) = 1300
        .ColWidth(9) = 1300
        .ColWidth(10) = 650
        .ColHidden(11) = True: .ColHidden(0) = True: .ColHidden(1) = True
        
        .ColAlignment(2) = flexAlignLeftTop
        .AllowUserResizing = flexResizeColumns
        .Redraw = True
    End With

End Sub
Private Sub RelojA()
    Screen.MousePointer = 11
End Sub
Private Sub RelojD()
    Screen.MousePointer = 0
End Sub
Private Sub CargoComboFolder()
    On Error GoTo ErrCCF

    If cCarpeta.ListCount > 0 Then Exit Sub
    Screen.MousePointer = 11
    Cons = "Select CarID, CarCodigo from Carpeta "
    CargoCombo Cons, cCarpeta
    Screen.MousePointer = 0
    Exit Sub
    
ErrCCF:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los folder."
End Sub

Private Sub CargoDocumentos()
    
    'Cargo los valores para los comprobantes de pago
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.Compracontado)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.Compracontado
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraCredito)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraCredito
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaCredito)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraNotaCredito
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaDevolucion)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraNotaDevolucion
    cComprobante.AddItem RetornoNombreDocumento(TipoDocumento.CompraRecibo)
    cComprobante.ItemData(cComprobante.NewIndex) = TipoDocumento.CompraRecibo

End Sub
Private Function ValidoImporte(Cadena As String)
    
Dim aS1 As String
Dim aS2 As String

    ValidoImporte = ""
    Cadena = UCase(Cadena)
    
    If Mid(Cadena, 1, 1) = ">" Or Mid(Cadena, 1, 1) = "<" Then
        If IsNumeric(Mid(Cadena, 2, Len(Cadena))) Then
            ValidoImporte = Mid(Cadena, 1, 1) & Mid(Cadena, 2, Len(Cadena))
            Exit Function
        End If
    End If
    
    If Mid(Cadena, 1, 1) = "E" Then
        If InStr(Cadena, "Y") <> 0 Then
            If IsNumeric(Mid(Cadena, 2, InStr(Cadena, "Y") - 2)) Then
                aS1 = Mid(Cadena, 2, InStr(Cadena, "Y") - 2)
                If IsNumeric(Mid(Cadena, InStr(Cadena, "Y") + 1, Len(Cadena))) Then
                    aS2 = Mid(Cadena, InStr(Cadena, "Y") + 1, Len(Cadena))
                    ValidoImporte = aS1 & "- " & aS2
                    Exit Function
                End If
            End If
        End If
    End If
    
End Function

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    Me.Refresh
    
End Sub
