VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMovFisico 
   Caption         =   "Consulta de Movimientos Físicos"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9825
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
   ScaleHeight     =   6795
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   7680
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8LCtl.VSFlexGrid vsConsulta 
      Height          =   3615
      Left            =   0
      TabIndex        =   35
      Top             =   1440
      Width           =   3855
      _cx             =   6800
      _cy             =   6376
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
      AutoSearchDelay =   2
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
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4095
      Left            =   60
      TabIndex        =   33
      Top             =   1380
      Width           =   11415
      _Version        =   196608
      _ExtentX        =   20135
      _ExtentY        =   7223
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
      ScaleWidth      =   6915
      TabIndex        =   34
      Top             =   6000
      Width           =   6975
      Begin VB.CommandButton butExcel 
         Height          =   310
         Left            =   3480
         Picture         =   "frmListado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Exportar a excel"
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0784
         Height          =   310
         Left            =   4560
         Picture         =   "frmListado.frx":0886
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4200
         Picture         =   "frmListado.frx":0DB8
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmListado.frx":1232
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmListado.frx":131C
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmListado.frx":1406
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3840
         Picture         =   "frmListado.frx":1640
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4920
         Picture         =   "frmListado.frx":1742
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5520
         Picture         =   "frmListado.frx":1B08
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":1C0A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmListado.frx":1F0C
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmListado.frx":224E
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmListado.frx":2550
         Style           =   1  'Graphical
         TabIndex        =   20
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
      TabIndex        =   32
      Top             =   6540
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
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
            AutoSize        =   1
            Object.Width           =   11668
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
      Height          =   1335
      Left            =   50
      TabIndex        =   31
      Top             =   0
      Width           =   9615
      Begin VB.TextBox tDocumento 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8340
         TabIndex        =   18
         Top             =   960
         Width           =   855
      End
      Begin AACombo99.AACombo cTipoDocumento 
         Height          =   315
         Left            =   6120
         TabIndex        =   17
         Top             =   960
         Width           =   2235
         _ExtentX        =   3942
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
      Begin VB.TextBox tCantidad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4380
         MaxLength       =   10
         TabIndex        =   15
         Top             =   960
         Width           =   915
      End
      Begin VB.TextBox tHasta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2580
         MaxLength       =   12
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox tDesde 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   780
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin AACombo99.AACombo cTipo 
         Height          =   315
         Left            =   780
         TabIndex        =   9
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   780
         MaxLength       =   50
         TabIndex        =   13
         Top             =   960
         Width           =   3015
      End
      Begin AACombo99.AACombo cGrupo 
         Height          =   315
         Left            =   4380
         TabIndex        =   11
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
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
         Left            =   4380
         TabIndex        =   5
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
      Begin AACombo99.AACombo cEstado 
         Height          =   315
         Left            =   7200
         TabIndex        =   7
         Top             =   240
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
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo &Mov.:"
         Height          =   255
         Left            =   5340
         TabIndex        =   16
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Cant.:"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "&Estado:"
         Height          =   255
         Left            =   6540
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Local:"
         Height          =   255
         Left            =   3840
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   1980
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Grupo:"
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Menu MnuAccesos 
      Caption         =   "Accesos"
      Visible         =   0   'False
      Begin VB.Menu MnuAccMenues 
         Caption         =   "Menú Accesos"
      End
      Begin VB.Menu MnuAccLinea 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAccDetalleFactura 
         Caption         =   "Detalle de &Factura"
      End
      Begin VB.Menu MnuAccDetalleOperacion 
         Caption         =   "Detalle de &Operación"
      End
   End
End
Attribute VB_Name = "frmMovFisico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Rs1 As rdoResultset
Private aTexto As String
Private bCargarImpresion As Boolean

Public prmArticulo As Long

Private Sub AccionLimpiar()
    cLocal.Text = ""
    cEstado.Text = ""
    cTipo.Text = ""
    cGrupo.Text = ""
    tArticulo.Text = "": tArticulo.Tag = ""
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

Private Sub butExcel_Click()
On Error GoTo errBE
    Dim sFile As String
    sFile = fnc_Browse(1, Replace(Me.Caption, "/", "-") & ".xls", "Exportar a excel")
    If sFile = "" Then Exit Sub
    vsConsulta.SaveGrid sFile, flexFileExcel, SaveExcelSettings.flexXLSaveFixedRows Or SaveExcelSettings.flexXLSaveRaw
errBE:
End Sub

Private Function fnc_Browse(ByVal xToFile As Byte, ByVal sFileN As String, ByVal sDialogT As String, Optional bShowSave As Boolean = True) As String
On Error GoTo errCancel
fnc_Browse = ""
 
    'Inicializo INITDIR
'    fnc_ValDirectory
            
    With cdFile
        .CancelError = True
        .DialogTitle = sDialogT
    'Var global
        '.InitDir = mExportDir
        If bShowSave Then .FileName = sFileN
        .Flags = cdlOFNOverwritePrompt Or cdlOFNCreatePrompt Or cdlOFNPathMustExist
        Select Case xToFile '1-Excel;   2-csv;  3=html
            Case 1: .Filter = "Libro de Microsoft Excel|*.xls"
            Case 2: .Filter = "Archivo de texto (csv)|*.csv"
            Case 3: .Filter = "Archivo html (*.htm)|*.htm"""
        End Select
        If bShowSave Then
            .ShowSave
        Else
            .ShowOpen
        End If
    End With
    fnc_Browse = cdFile.FileName
errCancel:
End Function

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

Private Sub cEstado_GotFocus()
    With cEstado
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda " Seleccione el estado físico del artículo a consultar. [Blanco = Todos]"
End Sub
Private Sub cEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cTipo
End Sub
Private Sub cEstado_LostFocus()
    cEstado.SelStart = 0
    Ayuda ""
End Sub

Private Sub cGrupo_GotFocus()
    With cGrupo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda " Seleccione un grupo de artículo a consultar"
End Sub
Private Sub cGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tArticulo
End Sub
Private Sub cGrupo_LostFocus()
    Ayuda ""
    cGrupo.SelStart = 0
End Sub

Private Sub cLocal_GotFocus()
    With cLocal
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda " Seleccione un local donde se dieron los movimientos. [Blanco = Todos] "
End Sub
Private Sub cLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cEstado
End Sub
Private Sub cLocal_LostFocus()
    cLocal.SelStart = 0
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
    If KeyAscii = vbKeyReturn Then Foco cGrupo
End Sub
Private Sub cTipo_LostFocus()
    cTipo.SelStart = 0
    Ayuda ""
End Sub

Private Sub cTipoDocumento_GotFocus()
    With cTipoDocumento
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda " Seleccione el tipo de documento que realizó el movimiento."
End Sub

Private Sub cTipoDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDocumento
End Sub

Private Sub cTipoDocumento_LostFocus()
    Ayuda ""
End Sub

Private Sub chVista_Click()
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
        Me.Refresh
    Else
        vsListado.ZOrder 0
        Me.Refresh
        AccionImprimir
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
    InicializoGrillas
    AccionLimpiar
    CargoTiposDeDocumento
    Cons = "Select TipCodigo, TipNombre From Tipo Order by TipNombre"
    CargoCombo Cons, cTipo
    Cons = "Select GruCodigo, GruNombre From Grupo Order by GruNombre"
    CargoCombo Cons, cGrupo
    Cons = "Select LocCodigo, LocNombre From Local Order By LocNombre"
    CargoCombo Cons, cLocal
    'cLocal.AddItem "(Todos)": cLocal.ItemData(cLocal.NewIndex) = 0
    Cons = "Select EsMCodigo, EsMAbreviacion From EstadoMercaderia Order by EsMAbreviacion"
    CargoCombo Cons, cEstado
    'cEstado.AddItem "(Todos)": cEstado.ItemData(cEstado.NewIndex) = 0
    FechaDelServidor
    tDesde.Text = Format(gFechaServidor, FormatoFP)
    tHasta.Text = tDesde.Text
    bCargarImpresion = True
    vsListado.Orientation = orPortrait
    If prmArticulo > 0 Then BuscoArticuloPorID prmArticulo
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub
Private Sub InicializoGrillas()
    On Error Resume Next
    With vsConsulta
        .Redraw = False
        .MultiTotals = False
        .WordWrap = False
        .Cols = 1: .Rows = 1
        .FixedCols = 0
        .FormatString = "<Fecha|<Artículo|<Estado|>Q|<Local|<Tipo|"
        .ColWidth(0) = 1750: .ColWidth(1) = 3250: .ColWidth(2) = 1000: .ColWidth(3) = 500: .ColWidth(4) = 1600: .ColWidth(5) = 1600
        .ColWidth(6) = 10
        .ExtendLastCol = True
        '.MergeCells = flexMergeRestrictColumns
        '.MergeCol(0) = True
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
    If Not IsDate(tDesde.Text) Then MsgBox "Debe ingresar una fecha desde.", vbInformation, "ATENCIÓN": Foco tDesde: Exit Sub
    If Not IsDate(tHasta.Text) Then MsgBox "Debe ingresar una fecha hasta.", vbInformation, "ATENCIÓN": Foco tHasta: Exit Sub
    If CDate(tDesde.Text) > CDate(tHasta.Text) Then MsgBox "La fecha Hasta es menor que la fecha Desde.", vbInformation, "ATENCIÓN": Foco tDesde: Exit Sub
    'Valido campo tcantidad
    If Trim(tCantidad.Text) <> "" Then
        If Not IsNumeric(tCantidad.Text) Then
            If Mid(Trim(tCantidad.Text), 1, 1) <> "<" And Mid(Trim(tCantidad.Text), 1, 1) <> ">" Then
                MsgBox "El formato de cantidad debe ser: un nro.; ó >#; ó < #.", vbExclamation, "ATENCIÓN"
                Foco tCantidad: Exit Sub
            Else
                If Not IsNumeric(Mid(Trim(tCantidad.Text), 2, Len(Trim(tCantidad.Text)))) Then
                    MsgBox "Formato incorrecto en la cantidad.", vbExclamation, "ATENCIÓN"
                    Foco tCantidad: Exit Sub
                End If
            End If
        End If
    End If
    'Valido Nro. de Documento
    tDocumento.Tag = "0"
    If Trim(tDocumento.Text) <> "" Then
        If Not ValidoDocumento Then Exit Sub
    End If
    Screen.MousePointer = 11
    bCargarImpresion = True
    vsConsulta.Rows = 1
    vsConsulta.Redraw = False
    CargoMovimientos Val(tDocumento.Tag)
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True
End Sub

Private Sub Label2_Click()
    Foco cGrupo
End Sub

Private Sub Label3_Click()
    Foco tArticulo
End Sub

Private Sub Label4_Click()
    Foco tDesde
End Sub

Private Sub Label5_Click()
    Foco tHasta
End Sub

Private Sub Label6_Click()
    Foco cLocal
End Sub

Private Sub Label7_Click()
    Foco cEstado
End Sub

Private Sub Label8_Click()
    Foco tCantidad
End Sub

Private Sub Label9_Click()
    Foco cTipoDocumento
End Sub

Private Sub MnuAccDetalleFactura_Click()
    EjecutarApp App.Path & "\Detalle de Factura.exe", vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)
End Sub

Private Sub MnuAccDetalleOperacion_Click()
    EjecutarApp App.Path & "\Detalle de Operaciones.exe", vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = "0"
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.Panels(3).Text = "Ingrese el mes y año de liquidación a consultar."
End Sub
Private Sub tArticulo_KeyPress(KeyAscii As Integer)
On Error GoTo ErrAP
    If KeyAscii = vbKeyReturn Then
        If Trim(tArticulo.Text) <> "" Then
            If Val(tArticulo.Tag) <> 0 Then Foco tCantidad: Exit Sub
            Screen.MousePointer = 11
            If IsNumeric(tArticulo.Text) Then
                BuscoArticuloPorCodigo tArticulo.Text
            Else
                BuscoArticuloPorNombre tArticulo.Text
            End If
            If Val(tArticulo.Tag) > 0 Then Foco tCantidad
        Else
            Foco tCantidad
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

Private Sub tCantidad_GotFocus()
    With tCantidad
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.Panels(3).Text = "Ingrese la cantidad de artículos que realizaron el movimiento. [Formatos <#,>#,#]"
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cTipoDocumento
End Sub

Private Sub tDesde_GotFocus()
    With tDesde
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda " Ingrese la fecha desde a consultar."
End Sub
Private Sub tDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tHasta
End Sub
Private Sub tDesde_LostFocus()
    Ayuda ""
    If IsDate(tDesde.Text) Then tDesde.Text = Format(tDesde.Text, FormatoFP) Else tDesde.Text = ""
End Sub

Private Sub tDocumento_GotFocus()
    With tDocumento
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese serie y nro. del documento(ctdo., crédito y remito) o el código del tipo de documento."
End Sub

Private Sub tDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub tDocumento_LostFocus()
    Ayuda ""
End Sub

Private Sub tHasta_GotFocus()
    With tHasta
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda " Ingrese la fecha hasta a consultar."
End Sub
Private Sub tHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cLocal
End Sub
Private Sub tHasta_LostFocus()
    Ayuda ""
    If IsDate(tHasta.Text) Then tHasta.Text = Format(tHasta.Text, FormatoFP) Else tHasta.Text = ""
End Sub

Private Sub vsConsulta_DblClick()
    If vsConsulta.Rows = 1 Then Exit Sub
    DeMovimientoStock.pTipoMovimiento = TipoEstadoMercaderia.Fisico
    DeMovimientoStock.pLista = vsConsulta
    DeMovimientoStock.Show vbModal, Me
End Sub

Private Sub vsConsulta_GotFocus()
    Ayuda "Doble Click = Detalle de Movimientos; Botón derecho acceso a Menúes."
End Sub

Private Sub vsConsulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        With vsConsulta
            If .Cell(flexcpData, .Row, 2) <> 0 Then
                MnuAccDetalleOperacion.Enabled = False
                MnuAccDetalleFactura.Enabled = False
                If .Cell(flexcpData, .Row, 2) = TipoDocumento.Credito Or .Cell(flexcpData, .Row, 2) = TipoDocumento.Contado Or .Cell(flexcpData, .Row, 2) = TipoDocumento.NotaCredito _
                    Or .Cell(flexcpData, .Row, 2) = TipoDocumento.NotaDevolucion Or .Cell(flexcpData, .Row, 2) = TipoDocumento.Remito Then
                    MnuAccDetalleOperacion.Enabled = True
                    MnuAccDetalleFactura.Enabled = True
                    PopupMenu MnuAccesos
                End If
            End If
        End With
    End If
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
        
        EncabezadoListado vsListado, "Consulta de Movimientos Físicos del " & Format(tDesde.Text, FormatoFP) & " al " & Format(tHasta.Text, FormatoFP), False
        vsListado.FileName = "Consulta de Movimientos Físicos"
        
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
Private Sub CargoMovimientos(IDDocumento As Long)
Dim Rs As rdoResultset
On Error GoTo ErrCS
Dim aCod As Long
    
    Screen.MousePointer = 11
    Cons = "Select * From MovimientoStockFisico "
    If cLocal.ListIndex > -1 Then
        Cons = Cons & ", Local "
    Else
        Cons = Cons & "Left Outer Join Local ON LocCodigo = MSFLocal "
    End If
    Cons = Cons & ", Articulo, EstadoMercaderia " _
        & " Where MSFFecha >= '" & Format(tDesde.Text & " 00:00:00", sqlFormatoFH) & "'" _
        & " And MSFFecha <= '" & Format(tHasta.Text & " 23:59:59", sqlFormatoFH) & "'" _
    
    If cTipoDocumento.ListIndex > -1 Then Cons = Cons & " And MSFTipoDocumento = " & cTipoDocumento.ItemData(cTipoDocumento.ListIndex)
    
    If IDDocumento > 0 Then Cons = Cons & " And MSFDocumento = " & IDDocumento
    'Uno Estado
    If cEstado.ListIndex > -1 Then Cons = Cons & " And MSFEstado = " & cEstado.ItemData(cEstado.ListIndex)
    Cons = Cons & " And MSFEstado = EsMCodigo"
    
    'Uno Artículo
    If Val(tArticulo.Tag) > 0 Then Cons = Cons & " And MSFArticulo = " & tArticulo.Tag
    Cons = Cons & " And MSFArticulo = ArtID"
    
    'Filtro de Tipo de Artículo
    If cTipo.ListIndex > -1 Then Cons = Cons & " And ArtTipo = " & cTipo.ItemData(cTipo.ListIndex)
    
    'Si hay local lo uno.
    If cLocal.ListIndex > -1 Then
        Cons = Cons & " And MSFLocal = " & cLocal.ItemData(cLocal.ListIndex) _
            & " And MSFLocal = LocCodigo"
    End If
    
    If Trim(tCantidad.Text) <> "" Then
        If Not IsNumeric(tCantidad.Text) Then
            If Mid(Trim(tCantidad.Text), 1, 1) = "<" Or Mid(Trim(tCantidad.Text), 1, 1) = ">" Then Cons = Cons & " And MSFCantidad " & tCantidad.Text
        Else
            Cons = Cons & " And MSFCantidad = " & tCantidad.Text
        End If
    End If
    
    If cGrupo.ListIndex > -1 Then
        Cons = Cons & " And ArtID IN (" _
                & " Select AGrArticulo from ArticuloGrupo" _
                & " Where AGrGrupo = " & cGrupo.ItemData(cGrupo.ListIndex) & ")"
    End If
    
    Cons = Cons & " Order By MSFFecha "
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCIÓN"
    Else
        Do While Not RsAux.EOF
            With vsConsulta
                .AddItem ""
                aCod = RsAux!MSFCodigo
                .Cell(flexcpData, .Rows - 1, 0) = aCod
                If Not IsNull(RsAux!MSFDocumento) Then aCod = RsAux!MSFDocumento Else aCod = 0
                .Cell(flexcpData, .Rows - 1, 1) = aCod
                If Not IsNull(RsAux!MSFTipoDocumento) Then aCod = RsAux!MSFTipoDocumento Else aCod = 0
                .Cell(flexcpData, .Rows - 1, 2) = aCod
                .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!MSFFecha, FormatoFHP)
                .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "#,000,000") & " " & Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!EsMAbreviacion)
                .Cell(flexcpText, .Rows - 1, 3) = RsAux!MSFCantidad
                If Not IsNull(RsAux!LocNombre) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!LocNombre)
                If Not IsNull(RsAux!MSFTipoDocumento) Then .Cell(flexcpText, .Rows - 1, 5) = RetornoNombreDocumento(RsAux!MSFTipoDocumento)
            End With
            RsAux.MoveNext
        Loop
        RsAux.Close
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrCS:
    clsGeneral.OcurrioError "Ocurrio un error al cargar los movimientos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub BuscoArticuloPorID(ByVal IDArticulo As Long)
On Error GoTo errBA
    
    Screen.MousePointer = 11
    Cons = "Select * From Articulo Where ArtID = " & IDArticulo
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
    Exit Sub
errBA:
    clsGeneral.OcurrioError "Error al buscar el artículo por ID.", Err.Description, "Buscar artículo"
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


Private Sub CargoTiposDeDocumento()
    
    With cTipoDocumento
        .Clear
        .AddItem RetornoNombreDocumento(TipoDocumento.ArregloStock): .ItemData(.NewIndex) = TipoDocumento.ArregloStock
        .AddItem RetornoNombreDocumento(TipoDocumento.CambioEstadoMercaderia): .ItemData(.NewIndex) = TipoDocumento.CambioEstadoMercaderia
        .AddItem RetornoNombreDocumento(TipoDocumento.CompraCarpeta): .ItemData(.NewIndex) = TipoDocumento.CompraCarpeta
'        .AddItem RetornoNombreDocumento(TipoDocumento.CompraCarta): .ItemData(.NewIndex) = TipoDocumento.CompraCarta
'        .AddItem RetornoNombreDocumento(TipoDocumento.Compracontado): .ItemData(.NewIndex) = TipoDocumento.Compracontado
'        .AddItem RetornoNombreDocumento(TipoDocumento.CompraCredito): .ItemData(.NewIndex) = TipoDocumento.CompraCredito
'        .AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaCredito): .ItemData(.NewIndex) = TipoDocumento.CompraNotaCredito
'        .AddItem RetornoNombreDocumento(TipoDocumento.CompraNotaDevolucion): .ItemData(.NewIndex) = TipoDocumento.CompraNotaDevolucion
        .AddItem RetornoNombreDocumento(TipoDocumento.Contado): .ItemData(.NewIndex) = TipoDocumento.Contado
        .AddItem RetornoNombreDocumento(TipoDocumento.ContadoDomicilio): .ItemData(.NewIndex) = TipoDocumento.ContadoDomicilio
        .AddItem RetornoNombreDocumento(TipoDocumento.Credito): .ItemData(.NewIndex) = TipoDocumento.Credito
        .AddItem RetornoNombreDocumento(TipoDocumento.Devolucion): .ItemData(.NewIndex) = TipoDocumento.Devolucion
        .AddItem RetornoNombreDocumento(TipoDocumento.Envios): .ItemData(.NewIndex) = TipoDocumento.Envios
        .AddItem RetornoNombreDocumento(TipoDocumento.IngresoMercaderiaEspecial): .ItemData(.NewIndex) = TipoDocumento.IngresoMercaderiaEspecial
        .AddItem RetornoNombreDocumento(TipoDocumento.NotaCredito): .ItemData(.NewIndex) = TipoDocumento.NotaCredito
        .AddItem RetornoNombreDocumento(TipoDocumento.NotaDevolucion): .ItemData(.NewIndex) = TipoDocumento.NotaDevolucion
        .AddItem RetornoNombreDocumento(TipoDocumento.NotaEspecial): .ItemData(.NewIndex) = TipoDocumento.NotaEspecial
        .AddItem RetornoNombreDocumento(TipoDocumento.Remito): .ItemData(.NewIndex) = TipoDocumento.Remito
        .AddItem RetornoNombreDocumento(TipoDocumento.Servicio): .ItemData(.NewIndex) = TipoDocumento.Servicio
        .AddItem RetornoNombreDocumento(TipoDocumento.ServicioCambioEstado): .ItemData(.NewIndex) = TipoDocumento.ServicioCambioEstado
        .AddItem RetornoNombreDocumento(TipoDocumento.ServicioDomicilio): .ItemData(.NewIndex) = TipoDocumento.ServicioDomicilio
        .AddItem RetornoNombreDocumento(TipoDocumento.Traslados): .ItemData(.NewIndex) = TipoDocumento.Traslados
    End With
End Sub
Private Function ValidoDocumento() As Boolean
On Error GoTo ErrVD
    ValidoDocumento = True
    If cTipoDocumento.ListIndex > -1 Then
        Select Case cTipoDocumento.ItemData(cTipoDocumento.ListIndex)
            Case TipoDocumento.Contado, TipoDocumento.Credito, TipoDocumento.NotaCredito, TipoDocumento.NotaDevolucion, TipoDocumento.NotaEspecial
                Dim strSerie As String
                strSerie = Mid(tDocumento.Text, 1, 1)
                If Trim(strSerie) = "" Then
                    MsgBox "Debe ingresar el formato (Z ######) donde (Z) es la serie del documento y el resto es el nro. de documento.", vbInformation, "ATENCIÓN": ValidoDocumento = False: Exit Function
                Else
                    If IsNumeric(strSerie) Then MsgBox "Debe ingresar el formato (Z ######) donde (Z) es una letra que indica la serie del documento y el resto es el nro. de documento.", vbInformation, "ATENCIÓN": ValidoDocumento = False: Exit Function
                End If
                If Not IsNumeric(Trim(Mid(tDocumento.Text, 2, Len(tDocumento.Text)))) Then MsgBox "Debe ingresar el formato (Z ######) donde (Z) es la serie del documento y el resto es el nro. de documento.", vbInformation, "ATENCIÓN": ValidoDocumento = False: Exit Function
                Cons = "Select * From Documento Where DocTipo = " & cTipoDocumento.ItemData(cTipoDocumento.ListIndex) _
                    & " And DocSerie = '" & strSerie & "' And DocNumero = " & Trim(Mid(tDocumento.Text, 2, Len(tDocumento.Text)))
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then
                    tDocumento.Tag = RsAux!DocCodigo
                    RsAux.Close
                Else
                    RsAux.Close
                    MsgBox "No se encontro un documento con tipo = " & Trim(cTipoDocumento.Text) & ", Serie = " & strSerie & " y Nro. = " & Trim(Mid(tDocumento.Text, 2, Len(tDocumento.Text))), vbInformation, "ATENCIÓN": ValidoDocumento = False: Exit Function
                End If
            Case Else
                If Not IsNumeric(tDocumento.Text) Then
                    MsgBox "El código de documento tiene que ser un nro. positvo para ese tipo de documento.", vbInformation, "ATENCIÓN": tDocumento.Text = "": ValidoDocumento = False
                Else
                    'Si me puso algo raro sale por el error.
                    tDocumento.Text = Abs(CLng(tDocumento.Text))
                    tDocumento.Tag = tDocumento.Text
                End If
        End Select
    Else
        If Not IsNumeric(tDocumento.Text) Then
            MsgBox "El código de documento tiene que ser un nro. positvo.", vbInformation, "ATENCIÓN": tDocumento.Text = "": ValidoDocumento = False
        Else
            'Si me puso algo raro sale por el error.
            tDocumento.Text = Abs(CLng(tDocumento.Text))
            tDocumento.Tag = tDocumento.Text
        End If
    End If
    Exit Function
ErrVD:
    clsGeneral.OcurrioError "Ocurrio un error al validar el código de documento.", Err.Description
End Function
