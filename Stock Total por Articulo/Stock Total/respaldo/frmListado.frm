VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Begin VB.Form frmListado 
   Caption         =   "Consulta de Stock Total por Artículo"
   ClientHeight    =   6690
   ClientLeft      =   1905
   ClientTop       =   2760
   ClientWidth     =   7575
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   7575
   Begin VB.PictureBox picHorizontal 
      BackColor       =   &H8000000D&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   2655
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLocal 
      Height          =   4335
      Left            =   5520
      TabIndex        =   5
      Top             =   1560
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7646
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
      TabIndex        =   4
      Top             =   1920
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7646
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
      TabIndex        =   8
      Top             =   1920
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
      Zoom            =   70
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   6435
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2805
            MinWidth        =   2805
            Key             =   "bd"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10028
            Key             =   "msg"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      BorderStyle     =   0  'None
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
      Height          =   975
      Left            =   60
      TabIndex        =   6
      Top             =   480
      Width           =   4935
      Begin VB.CommandButton bModificar 
         Caption         =   "&Modificar"
         Height          =   285
         Left            =   3480
         TabIndex        =   3
         Top             =   660
         Width           =   975
      End
      Begin VB.CheckBox chEnUso 
         Caption         =   "En &Uso"
         Height          =   195
         Left            =   2160
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chHabilitado 
         Caption         =   "&Habilitado"
         Height          =   195
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox tArticulo 
         Height          =   285
         Left            =   840
         MaxLength       =   50
         TabIndex        =   0
         Top             =   0
         Width           =   3615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Width           =   855
      End
      Begin VB.Label labUltimaCompra 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   3615
      End
   End
   Begin MSComctlLib.ImageList imgExplore 
      Left            =   8880
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":030A
            Key             =   "next"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":0626
            Key             =   "back"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":0942
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":0A9E
            Key             =   "plantilla"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1378
            Key             =   "help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1694
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":19B0
            Key             =   "print"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1AC2
            Key             =   "printcnfg"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1BD4
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1EEE
            Key             =   "form"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":2340
            Key             =   "changedb"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":2792
            Key             =   "firstpage"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":2BE4
            Key             =   "previouspage"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":3036
            Key             =   "nextpage"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":3488
            Key             =   "lastpage"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbExplorer 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bback"
            Object.ToolTipText     =   "Anterior (Ctrl+A)."
            Style           =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bnext"
            Object.ToolTipText     =   "Siguiente (Ctrl+S)."
            Style           =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
            Object.ToolTipText     =   "Actualizar (Ctrl+Z)."
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "plantilla"
            Object.ToolTipText     =   "Plantillas interactivas"
            Style           =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   300
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Imprimir [Ctrl+I]"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "preview"
            Object.ToolTipText     =   "Preview [Ctrl+P]"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "firstpage"
            Object.ToolTipText     =   "Primera página."
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "previouspage"
            Object.ToolTipText     =   "Página anterior."
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nextpage"
            Object.ToolTipText     =   "Página Siguiente."
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "lastpage"
            Object.ToolTipText     =   "Última página."
         EndProperty
      EndProperty
   End
   Begin VB.Image imgHorizontal 
      Height          =   45
      Left            =   120
      MousePointer    =   7  'Size N S
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Menu MnuOpcion 
      Caption         =   "&Archivo"
      Begin VB.Menu MnuOpBack 
         Caption         =   "&Anterior"
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuOpNext 
         Caption         =   "&Siguiente"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuOpRefrescar 
         Caption         =   "&Refrescar"
         Shortcut        =   ^Z
      End
      Begin VB.Menu MnuOpLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuOpConfPage 
         Caption         =   "&Configurar Página"
      End
      Begin VB.Menu MnuOpPreview 
         Caption         =   "&Preview"
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuOpLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpChangeDB 
         Caption         =   "Cambiar &Base de Datos"
      End
      Begin VB.Menu MnuOpLinea3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOpSalir 
         Caption         =   "Sa&lir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuIrA 
      Caption         =   "&Ir a"
      Begin VB.Menu MnuIrAStockLocal 
         Caption         =   "&Sotck en Locales"
      End
      Begin VB.Menu MnuIrALinea1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIrAMovimiento 
         Caption         =   "&Movimientos"
         Begin VB.Menu MnuIrAMovFisico 
            Caption         =   "&Movimientos físicos"
         End
         Begin VB.Menu MnuIrAControlMovFis 
            Caption         =   "&Control de Mov. Físicos"
         End
         Begin VB.Menu MnuIrAGenHisStock 
            Caption         =   "Genero Historico Stock"
         End
         Begin VB.Menu MnuIrAMovLinea 
            Caption         =   "-"
         End
         Begin VB.Menu MnuIrAMovVirt 
            Caption         =   "Movimientos &Virtuales"
         End
      End
      Begin VB.Menu MnuIrAArreglos 
         Caption         =   "&Arreglos"
         Begin VB.Menu MnuIrAPendRetiro 
            Caption         =   "&Pendientes de Retiro"
         End
         Begin VB.Menu MnuIrAArrIngEspecial 
            Caption         =   "&Ingreso Especial"
         End
         Begin VB.Menu MnuIrATraslEspecial 
            Caption         =   "&Traslado Especial"
         End
         Begin VB.Menu MnuIrAArrLinea 
            Caption         =   "-"
         End
         Begin VB.Menu MnuIrACorrStockVir 
            Caption         =   "Corrijo Stock &Virtual"
         End
         Begin VB.Menu MnuIrAArrArregloStk 
            Caption         =   "&Arreglo Stock"
         End
         Begin VB.Menu MnuIrAVerifStock 
            Caption         =   "&Verifico Stock"
         End
      End
   End
   Begin VB.Menu MnuPlantillas 
      Caption         =   "&Plantillas"
      Begin VB.Menu MnuPlaIndex 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lIDArticulo As Long

Private bSizeAjuste As Boolean
Private Rs1 As rdoResultset
Private aTexto As String
Private bCargarImpresion As Boolean

Public Sub SetArticuloParmetro(ByVal lArtParam As Long)
    If lArtParam > 0 Then
        BuscoArticuloPorID lArtParam, False, False
        If Val(tArticulo.Tag) > 0 Then AccionConsultar
    End If
End Sub

Private Sub AccionLimpiar()
    tArticulo.Text = "": tArticulo.Tag = "0"
    vsConsulta.Rows = 1: vsLocal.Rows = 1
End Sub

Private Sub bModificar_Click()

    If chEnUso.Enabled Then
    
        Select Case MsgBox("¿Confirma modificar la ficha del artículo?", vbYesNo, "ATENCIÓN")
            Case vbYes
                FechaDelServidor
'                On Error GoTo errBM
                If bModificar.Tag = "" Then
                    bModificar.Tag = InputBox("Ingrese su dígito de Usuario", "Usuario")
                    If bModificar.Tag = "" Then Exit Sub
                    bModificar.Tag = BuscoUsuarioDigito(Val(bModificar.Tag), True)
                    If bModificar.Tag = "0" Then bModificar.Tag = "": Exit Sub
                End If
                
                Cons = "Select * From Articulo Where ArtID = " & Val(tArticulo.Tag)
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                RsAux.Edit
                If chEnUso.Value = 1 Then
                    RsAux!ArtEnUso = True
                Else
                    RsAux!ArtEnUso = False
                End If
                If chHabilitado.Value = 1 Then
                    RsAux("ArtHabilitado") = "S"
                Else
                    RsAux("ArtHabilitado") = "N"
                End If
                RsAux("ArtModificado") = Format(gFechaServidor, "mm/dd/yyyy hh:nn:ss")
                RsAux("ArtUsuModificacion") = Val(bModificar.Tag)
                RsAux.Update
                RsAux.Close
                chEnUso.Enabled = False: chHabilitado.Enabled = False
            
            Case vbNo
                chEnUso.Enabled = False: chHabilitado.Enabled = False
                chEnUso.Value = Val(chEnUso.Tag)
                chHabilitado.Value = Val(chHabilitado.Tag)
        End Select
        bModificar.Caption = "&Modificar"
        
    Else
        chEnUso.Enabled = True
        chHabilitado.Enabled = True
        bModificar.Caption = "A&plicar"
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    
    ReDim arrClientes(25)
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    InicializoGrillas
    AccionLimpiar
    bCargarImpresion = True
    vsListado.Orientation = orPortrait
    BorroTodo
    CargoMenuPlantillas
    Status.Panels("bd").Text = "BD: " & miConexion.RetornoPropiedad(bdb:=True) & "  "
    Status.Tag = miConexion.RetornoPropiedad(bdb:=True)
    MenuExplorer
    InicializoToolbar
    
    MenuPlantilla False
    
    vsListado.Visible = False
    'AccionPreview
    fFiltros.Top = 480
    imgHorizontal.Top = Me.ScaleHeight / 2
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
        .FormatString = "<Estado|>Disponible|>No Disponible|>Total|"
        .ColWidth(0) = 1700: .ColWidth(3) = 1000: .ColWidth(4) = 15
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
        .ColWidth(0) = 2100: .ColWidth(1) = 1500: .ColWidth(2) = 1000: .ColWidth(3) = 15
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(0) = True
        .Redraw = True
    End With
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11
    If Me.Height < fFiltros.Top + fFiltros.Height + 1700 Then
        Me.Height = fFiltros.Top + fFiltros.Height + 1700
    End If
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    
    vsListado.Move fFiltros.Left, fFiltros.Top + fFiltros.Height + 20, fFiltros.Width, _
        Me.ScaleHeight - (vsListado.Top + Status.Height + 20)
    
    With imgHorizontal
        .Left = vsListado.Left
        .Width = vsListado.Width
        If .Top > Me.ScaleHeight Then .Top = Me.ScaleHeight - 800
        If .Top < fFiltros.Height + tbExplorer.Height + 300 Then .Top = fFiltros.Height + tbExplorer.Height + 300
        If Not vsListado.Visible Then
            .ZOrder 0
        End If
    End With
    
    vsConsulta.Move vsListado.Left, vsListado.Top, vsListado.Width, imgHorizontal.Top - vsListado.Top
    
    With vsLocal
        .Left = vsConsulta.Left
        .Width = vsListado.Width
        .Top = imgHorizontal.Top + imgHorizontal.Height
        .Height = vsListado.Top + vsListado.Height - .Top - 20
    End With
'    Me.Refresh
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    Set cBase = Nothing
    Set eBase = Nothing
    End
    
End Sub

Private Sub AccionConsultar()
Dim Rs As rdoResultset
    On Error GoTo errConsultar
    Screen.MousePointer = 11
    chEnUso.Enabled = False: chHabilitado.Enabled = False
    chEnUso.Value = Val(chEnUso.Tag)
    chHabilitado.Value = Val(chHabilitado.Tag)
    bModificar.Caption = "&Modificar"
    bCargarImpresion = True
    vsConsulta.Rows = 1: vsLocal.Rows = 1
    CargoStock
    Me.Refresh
    ButtonRegistros vsListado.Visible
    If tbExplorer.Buttons("plantilla").ButtonMenus.Count > 0 Then MenuPlantilla True
    Foco tArticulo
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Sub imgHorizontal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bSizeAjuste = True
    With picHorizontal
        .Move imgHorizontal.Left, imgHorizontal.Top, imgHorizontal.Width, imgHorizontal.Height
        .Visible = True
        .ZOrder 0
    End With
End Sub

Private Sub imgHorizontal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bSizeAjuste Then
        If Y < picHorizontal.Top Then
            picHorizontal.Move vsConsulta.Left, imgHorizontal.Top + Y, vsConsulta.Width
        Else
            picHorizontal.Move vsConsulta.Left, imgHorizontal.Top - Y, vsConsulta.Width
        End If
    End If
End Sub

Private Sub imgHorizontal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If picHorizontal.Top < fFiltros.Height + 800 Then picHorizontal.Top = fFiltros.Height + 800
    If picHorizontal.Top + 800 > Me.ScaleHeight Then picHorizontal.Top = Me.ScaleHeight - 800
    imgHorizontal.Move 0, picHorizontal.Top
    
    picHorizontal.Visible = False
    bSizeAjuste = False
    Call Form_Resize
End Sub

Private Sub Label3_Click()
    Foco tArticulo
End Sub

Private Sub MnuIrAArrArregloStk_Click()
    EjecutarApp App.Path & "\Arreglo Stock.exe"
End Sub

Private Sub MnuIrAArrIngEspecial_Click()
    EjecutarApp App.Path & "\Ingreso MercaderiaE.exe"
End Sub

Private Sub MnuIrAControlMovFis_Click()
    EjecutarApp App.Path & "\Control Movimiento Fisico.exe"
End Sub

Private Sub MnuIrACorrStockVir_Click()
    EjecutarApp App.Path & "\Corrijo Stock Virtual.exe"
End Sub

Private Sub MnuIrAGenHisStock_Click()
    EjecutarApp App.Path & "\Genero Historico Stock.exe"
End Sub

Private Sub MnuIrAMovFisico_Click()
    EjecutarApp App.Path & "\Movimientos Fisicos.exe"
End Sub

Private Sub MnuIrAMovVirt_Click()
    EjecutarApp App.Path & "\Movimientos Virtuales.exe"
End Sub

Private Sub MnuIrAPendRetiro_Click()
    EjecutarApp App.Path & "\Pendientes de Retiro.exe"
End Sub

Private Sub MnuIrAStockLocal_Click()
    EjecutarApp App.Path & "\Stock de locales.exe"
End Sub

Private Sub MnuIrATraslEspecial_Click()
    EjecutarApp App.Path & "\Traslado_Mercaderia_Especial.exe"
End Sub

Private Sub MnuIrAVerifStock_Click()
    EjecutarApp App.Path & "\Verificacion de Stock.exe"
End Sub

Private Sub MnuOpBack_Click()
    Call tbExplorer_ButtonMenuClick(tbExplorer.Buttons("bback").ButtonMenus.Item(1))
End Sub

Private Sub MnuOpChangeDB_Click()
Dim newB As String
    
    On Error GoTo errCh
    
    If Not miConexion.AccesoAlMenu("Cambiar_Conexion") Then Exit Sub
    
    newB = InputBox("Ingrese el texto del login para la nueva conexión" & vbCrLf & _
                "Id de aplicación en archivo de conexiones.", "Cambio de Base de Datos")
    
    If Trim(newB) = "" Then Exit Sub
    If MsgBox("Está seguro de cambiar la base de datos al login " & newB, vbQuestion + vbYesNo + vbDefaultButton2, "Cambiar Base") = vbNo Then Exit Sub
    
    'Limpio la ficha
    AccionLimpiar
    
    newB = miConexion.TextoConexion(newB)
    If Trim(newB) = "" Then Exit Sub
    
    If Status.Tag = miConexion.RetornoPropiedad(bdb:=True) Then
        Me.BackColor = vbButtonFace
        fFiltros.BackColor = vbButtonFace
        chEnUso.BackColor = vbButtonFace
        chHabilitado.BackColor = vbButtonFace
    Else
        Me.BackColor = &HC0C000
        fFiltros.BackColor = &HC0C000
        chEnUso.BackColor = &HC0C000
        chHabilitado.BackColor = &HC0C000
    End If
    Screen.MousePointer = 11
    On Error Resume Next
    cBase.Close
    On Error GoTo errCh
    Set cBase = Nothing
    If InicioConexionBD(newB) Then
        Status.Panels("bd").Text = "BD: " & miConexion.RetornoPropiedad(bdb:=True) & "  "
    Else
        Status.Panels("bd").Text = "BD: EN ERROR  "
    End If
    
    Screen.MousePointer = 0
    MsgBox "Ahora está trabajanbo en la nueva base de datos.", vbExclamation, "Base Cambiada OK"
    Exit Sub
    
errCh:
    Status.Panels("bd").Text = "BD: EN ERROR  "
    clsGeneral.OcurrioError "Error de Conexión." & vbCrLf & " La conexión está en estado de error, conectese a una base de datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub MnuOpConfPage_Click()
    AccionConfigurar
End Sub

Private Sub MnuOpImprimir_Click()
    AccionImprimir True
End Sub

Private Sub MnuOpNext_Click()
    Call tbExplorer_ButtonMenuClick(tbExplorer.Buttons("bnext").ButtonMenus.Item(1))
End Sub

Private Sub MnuOpPreview_Click()
    AccionPreview
End Sub

Private Sub MnuOpRefrescar_Click()
    AccionRefrescar
End Sub

Private Sub MnuOpSalir_Click()
    Unload Me
End Sub

Private Sub MnuPlaIndex_Click(Index As Integer)
    If Val(tArticulo.Tag) <> 0 Then
        EjecutarApp App.Path & "\appExploreMsg.exe ", Val(MnuPlaIndex(Index).Tag) & ":" & tArticulo.Tag
    End If
End Sub

Private Sub tArticulo_Change()
    If Val(tArticulo.Tag) > 0 Then
        BorroTodo
        MenuPlantilla False
    End If
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Status.Panels(2).Text = "Ingrese el artículo a consultar."
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrAP
    Screen.MousePointer = 11
    
    If KeyAscii = vbKeyReturn Then
        If Val(tArticulo.Tag) <> 0 Then
            Screen.MousePointer = 0: Exit Sub
        End If
        labUltimaCompra.Caption = ""
        vsConsulta.Rows = 1: vsLocal.Rows = 1
        If Trim(tArticulo.Text) <> "" Then
            If IsNumeric(tArticulo.Text) Then
                BuscoArticuloPorCodigo tArticulo.Text
            Else
                BuscoArticuloPorNombre tArticulo.Text
            End If
            If Val(tArticulo.Tag) > 0 Then AccionConsultar
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
        
        EncabezadoListado vsListado, "Consulta de Stock Total al " & Format(Date, FormatoFP), False
        vsListado.FileName = "Consulta de Stock Total"
        vsListado.FontBold = True
        vsListado.Paragraph = "Artículo: " & tArticulo.Text
        vsListado.Paragraph = ""
        vsListado.FontBold = False
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
Dim QAFacturarR As Long
Dim QAFacturarE As Long, CantExtra As Long
On Error GoTo ErrCS
    
    If tArticulo.Tag = "0" Then Exit Sub
    Screen.MousePointer = 11
    ArmoConsultaTotal
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If RsAux.EOF Then
        RsAux.Close
        Cons = "Select LocNombre, StLCantidad, EsMAbreviacion From Local, StockLocal, EstadoMercaderia " _
            & " Where StlArticulo = " & CLng(tArticulo.Tag) _
            & " And StlCantidad <> 0 And StLLocal = LocCodigo And StLEstado = EsMCodigo"
        Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        Do While Not Rs.EOF
            InsertoFilaLocal Trim(Rs!LocNombre), Trim(Rs!EsMAbreviacion), Rs!StLCantidad
            Rs.MoveNext
        Loop
        Rs.Close
        If vsLocal.Rows = 1 Then
            MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCIÓN"
        End If
        
    Else
        
        QAFacturarR = StockAFacturarRetira(CLng(tArticulo.Tag))
        QAFacturarE = StockAFacturarEnvia(CLng(tArticulo.Tag))
        If QAFacturarR > 0 Then InsertoFilaConsulta "A Facturar Retirar", QAFacturarR, 0, False
        If QAFacturarE > 0 Then InsertoFilaConsulta "A Facturar Envia", QAFacturarE, 0, False
        
        Do While Not RsAux.EOF
            If RsAux!StTTipoEstado = TipoEstadoMercaderia.Fisico Then
                If RsAux!EsMBajaStockTotal = 0 Then
                    CantExtra = CantidadNoDisponible(CLng(tArticulo.Tag), RsAux!Estado)
                    If RsAux!StTCantidad <> 0 Or CantExtra <> 0 Then InsertoFilaConsulta Trim(RsAux!EsMAbreviacion), RsAux!StTCantidad - CantExtra, CantExtra
                End If
                Cons = "Select LocNombre, StLCantidad From Local, StockLocal " _
                    & " Where StlArticulo = " & CLng(tArticulo.Tag) _
                    & " And StLEstado = " & RsAux!Estado _
                    & " And StlCantidad <> 0 And StLLocal = LocCodigo"
                Set Rs = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
                Do While Not Rs.EOF
                    InsertoFilaLocal Trim(Rs!LocNombre), Trim(RsAux!EsMAbreviacion), Rs!StLCantidad
                    Rs.MoveNext
                Loop
                Rs.Close
            Else
                'Estados Virtuales
                Select Case RsAux!Estado
                    Case TipoMovimientoEstado.AEntregar
                        If RsAux!StTCantidad - QAFacturarE <> 0 Then InsertoFilaConsulta RetornoEstadoVirtual(RsAux!Estado), RsAux!StTCantidad - QAFacturarE, 0, False
                    Case TipoMovimientoEstado.ARetirar
                        If RsAux!StTCantidad - QAFacturarR <> 0 Then InsertoFilaConsulta RetornoEstadoVirtual(RsAux!Estado), RsAux!StTCantidad - QAFacturarR, 0, False
                    Case TipoMovimientoEstado.Reserva
                        If RsAux!StTCantidad <> 0 Then InsertoFilaConsulta RetornoEstadoVirtual(RsAux!Estado), RsAux!StTCantidad, 0, False
                End Select
            End If
            RsAux.MoveNext
        Loop
        RsAux.Close
        CargoFisicosNoDisponibles
    End If
    If vsConsulta.Rows > 1 Then
        With vsConsulta
            .Subtotal flexSTSum, -1, 1, "#,##0", Obligatorio, Rojo, True, "Total"
            .Subtotal flexSTSum, -1, 2, "#,##0", Obligatorio, Rojo, True, ""
            .Subtotal flexSTSum, -1, 3, "#,##0", Obligatorio, Rojo, True, ""
        End With
        With vsLocal
            If .Rows > 1 Then .Select 1, 0, 1, 1
            .Sort = flexSortGenericAscending
            .Subtotal flexSTSum, 0, 2, "#,##0", Inactivo, Rojo, False, "%s"
            .Subtotal flexSTSum, -1, 2, "#,##0", Obligatorio, Rojo, True, "Total"
        End With
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrCS:
    clsGeneral.OcurrioError "Ocurrio un error al cargar el stock.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub CargoFisicosNoDisponibles()
    
    Cons = "Select ArtID, ArtCodigo, ArtNombre, StTTipoEstado, StTEstado, StTCantidad, EsMAbreviacion From Articulo, StockTotal, EstadoMercaderia " _
        & " Where ArtHabilitado = 'S' And StTTipoEstado = " & TipoEstadoMercaderia.Fisico _
        & " And ArtID = " & tArticulo.Tag _
        & " And ArtID = StTArticulo And StTEstado = EsMCodigo And EsMBajaStockTotal = 1"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    Do While Not RsAux.EOF
        If RsAux!StTCantidad <> 0 Then InsertoFilaConsulta Trim(RsAux!EsMAbreviacion), 0, RsAux!StTCantidad
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Sub
Private Sub ArmoConsultaTotal()
    
    'Saco todos los artículos y su stock Total.---------------------------
    'Hago una unión por tipo de estado de mercadería. ArtID, ArtCodigo, ArtNombre,
    Cons = "Select  IsNull(StTTipoEstado, 1) as StTTipoEstado, EsMCodigo as Estado, StTCantidad, EsMAbreviacion, EsMBajaStockTotal " _
        & " From EstadoMercaderia " _
                & " Left Outer Join StockTotal On STTEstado = EsMCodigo And StTArticulo = " & tArticulo.Tag _
                & " And StTTipoEstado = " & TipoEstadoMercaderia.Fisico _
        & " Union" _
        & " Select StTTipoEstado, StTEstado as Estado, StTCantidad, EsMAbreviacion = '', EsMBajaStockTotal = 0 From StockTotal" _
        & " Where StTArticulo = " & tArticulo.Tag _
        & " And StTTipoEstado = " & TipoEstadoMercaderia.Virtual _
        & " Order by StTTipoEstado DESC"
        
End Sub
Private Function CantidadNoDisponible(IDArticulo As Long, IdEstado As Integer) As Currency
Dim RsStL As rdoResultset
    Cons = "Select Sum(StLCantidad) From StockLocal " _
        & " Where StlArticulo = " & IDArticulo _
        & " And StLEstado = " & IdEstado _
        & " And StLLocal IN (Select SucCodigo From Sucursal Where SucExtras = 1)"
    Set RsStL = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not IsNull(RsStL(0)) Then CantidadNoDisponible = RsStL(0) Else CantidadNoDisponible = 0
    RsStL.Close
End Function
Private Sub InsertoFilaConsulta(ByVal sEstado As String, lDisponible As Long, lNoDisponible As Long, Optional EstFisico As Boolean = True)
    With vsConsulta
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = Trim(sEstado)
        .Cell(flexcpText, .Rows - 1, 1) = Format(lDisponible, "#,##0")
        .Cell(flexcpText, .Rows - 1, 2) = Format(lNoDisponible, "#,##0")
        .Cell(flexcpText, .Rows - 1, 3) = Format(lNoDisponible + lDisponible, "#,##0")
        .Cell(flexcpFontBold, .Rows - 1, 3) = True
        If Not EstFisico Then .Cell(flexcpForeColor, .Rows - 1, 0) = vbHighlight
    End With
End Sub
Private Sub InsertoFilaLocal(ByVal sLocal As String, ByVal sEstado As String, ByVal lCant As Long)
    With vsLocal
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = sLocal
        .Cell(flexcpText, .Rows - 1, 1) = Trim(sEstado)
        .Cell(flexcpText, .Rows - 1, 2) = Format(lCant, "#,##0")
        .Cell(flexcpFontBold, .Rows - 1, 3) = True
    End With
End Sub

Private Sub BuscoArticuloPorID(ByVal IDArticulo As Long, ByVal bNext As Boolean, ByVal bPrevious As Boolean)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
Dim bSeImporta As Boolean

    Screen.MousePointer = 11
    If Not (bNext Or bPrevious) Then
        Cons = "Select * From Articulo Where ArtID = " & IDArticulo
    Else
        Cons = "Select Top 1 art2.* From Articulo art1, Articulo art2 Where art1.ArtID = " & IDArticulo _
            & " And art2.ArtEnUso = 1 And art1.ArtCodigo "
        If bNext Then
            Cons = Cons & " < art2.ArtCodigo Order By art2.ArtCodigo Asc"
        Else
            Cons = Cons & " > art2.ArtCodigo Order By art2.ArtCodigo Desc"
        End If
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    BorroTodo
    If RsAux.EOF Then
        RsAux.Close
        tArticulo.Tag = "0"
        If bNext Or bPrevious Then BuscoArticuloPorID IDArticulo, False, False
    Else
        tArticulo.Text = Format(RsAux!ArtCodigo, "#,000,000") & " " & Trim(RsAux!ArtNombre)
        tArticulo.Tag = RsAux!artid
        bSeImporta = RsAux("ArtSeImporta")
        If RsAux("ArtEnUso") Then chEnUso.Value = 1: chEnUso.Tag = "1"
        If Not IsNull(RsAux("ArtHabilitado")) Then
            If UCase(RsAux("ArtHabilitado")) = "S" Then
                chHabilitado.Value = 1
                chHabilitado.Tag = "1"
            End If
        End If
        RsAux.Close
        lIDArticulo = Val(tArticulo.Tag)
        EstadoObjetos True
        If bSeImporta Then
            labUltimaCompra.Caption = "Última Compra: " & RetornoUltimaCompra(tArticulo.Tag)
        Else
            labUltimaCompra.Caption = "Última Compra: " & CompraNoImportacion(tArticulo.Tag)
        End If
        arr_AddItem lIDArticulo, Trim(tArticulo.Text)
        MenuExplorer
    End If
    Screen.MousePointer = 0
End Sub

Private Sub BuscoArticuloPorCodigo(IDArticulo As Long)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
Dim bSeImporta As Boolean

    Screen.MousePointer = 11
    Cons = "Select * From Articulo Where ArtCodigo = " & IDArticulo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    BorroTodo
    
    If RsAux.EOF Then
        RsAux.Close
        tArticulo.Tag = "0"
        MsgBox "No existe un artículo que posea ese código.", vbExclamation, "ATENCIÓN"
    Else
        tArticulo.Text = Format(RsAux!ArtCodigo, "#,000,000") & " " & Trim(RsAux!ArtNombre)
        tArticulo.Tag = RsAux!artid
        bSeImporta = RsAux("ArtSeImporta")
        
        If RsAux("ArtEnUso") Then chEnUso.Value = 1: chEnUso.Tag = "1"
        If Not IsNull(RsAux("ArtHabilitado")) Then
            If UCase(RsAux("ArtHabilitado")) = "S" Then
                chHabilitado.Value = 1
                chHabilitado.Tag = "1"
            End If
        End If
        RsAux.Close
        
        EstadoObjetos True
        
        If bSeImporta Then
            labUltimaCompra.Caption = "Última Compra: " & RetornoUltimaCompra(tArticulo.Tag)
        Else
            labUltimaCompra.Caption = "Última Compra: " & CompraNoImportacion(tArticulo.Tag)
        End If
        lIDArticulo = Val(tArticulo.Tag)
        arr_AddItem lIDArticulo, Trim(tArticulo.Text)
        MenuExplorer
    End If
    Screen.MousePointer = 0

End Sub

Private Sub BuscoArticuloPorNombre(NomArticulo As String)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
Dim Resultado As Long

    Screen.MousePointer = 11
    Resultado = 0
    Cons = "Select ArtId, Código = ArtCodigo, Nombre = ArtNombre from Articulo" _
        & " Where ArtNombre LIKE '" & Replace(NomArticulo, " ", "%") & "%'" _
        & " Order By ArtNombre"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No hay datos para el filtro ingresado.", vbInformation, "ATENCIÓN"
    Else
        Resultado = RsAux(1)
        RsAux.MoveNext
        If RsAux.EOF Then
            RsAux.Close
        Else
            RsAux.Close
            Resultado = 0
            Dim LiAyuda As New clsListadeAyuda
            If LiAyuda.ActivarAyuda(cBase, Cons, 5000, 1, "Lista de Artículos") Then
                Resultado = LiAyuda.RetornoDatoSeleccionado(1)
            Else
                Resultado = 0
            End If
            Set LiAyuda = Nothing       'Destruyo la clase.
        End If
    End If
    Screen.MousePointer = 11
    If Resultado > 0 Then BuscoArticuloPorCodigo Resultado
    Screen.MousePointer = 0
    
End Sub

Private Function StockAFacturarRetira(Articulo As Long) As Long
On Error GoTo ErrSAFR
    Cons = "Select Sum(RVTARetirar) From VentaTelefonica, RenglonVtaTelefonica " _
            & " Where VTeTipo = " & TipoDocumento.ContadoDomicilio _
            & " And VTeDocumento = Null And VTeAnulado Is Null" _
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

Private Function RetornoUltimaCompra(lArt As Long) As String
On Error GoTo ErrRUC
    
    RetornoUltimaCompra = ""
    Cons = "Select * From Embarque, ArticuloFolder" _
        & " Where AFoArticulo = " & lArt & " And AFoTipo = 2" _
        & " And AFoCodigo = EmbID And EmbFLocal Is Not Null" _
        & " And EmbFArribo = (" _
            & " Select Max(EmbFArribo) From Embarque, ArticuloFolder" _
            & " Where AFoArticulo = " & lArt & " And AFoTipo = 2" _
            & " And AFoCodigo = EmbID And EmbFLocal Is Not Null)" _

    Set Rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly)
    
    If Not Rs1.EOF Then
        RetornoUltimaCompra = Format(Rs1("EmbFArribo"), "d/mm/yy") & "  Cantidad: " & Rs1("AFoCantidad")
    End If
    Rs1.Close
    
    Exit Function
ErrRUC:
    clsGeneral.OcurrioError "Ocurrio un error al buscar la última fecha de compra.", Err.Description
End Function

Private Function CompraNoImportacion(ByVal lArt As Long) As String
Dim lCant As Long

    CompraNoImportacion = ""
    'Saco la Fecha y el Documento de compra.
    Cons = " Select * from Compra, CompraRenglon" _
        & " Where CReArticulo = " & lArt _
        & " And ComTipoDocumento In (" & TipoDocumento.Compracontado & ", " & TipoDocumento.CompraCredito & ")" _
        & " And ComCodigo = CReCompra And ComFecha = (" _
            & " Select Max(ComFecha) from Compra, CompraRenglon" _
            & " Where CReArticulo = " & lArt _
            & " And ComTipoDocumento In (" & TipoDocumento.Compracontado & ", " & TipoDocumento.CompraCredito & ")" _
            & " And ComCodigo = CReCompra )"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        CompraNoImportacion = Format(RsAux("ComFecha"), "d/mm/yy") & "  Cantidad: " & RsAux("CReCantidad")
    End If
    RsAux.Close
        
End Function

Private Sub BorroTodo()
    labUltimaCompra.Caption = ""
    tArticulo.Tag = "0"
    vsConsulta.Rows = 1: vsLocal.Rows = 1
    chEnUso.Tag = ""
    chHabilitado.Tag = ""
    EstadoObjetos False
    If Not vsListado.Visible Then
        With tbExplorer
            .Buttons("nextpage").Enabled = False
            .Buttons("previouspage").Enabled = False
            .Refresh
        End With
    End If
End Sub

Private Sub EstadoObjetos(ByVal bEst As Boolean)
    
    If Not bEst Then
        chEnUso.Value = 0
        chHabilitado.Value = 0
        chEnUso.Enabled = bEst
        chHabilitado.Enabled = bEst
        bModificar.Caption = "&Modificar"
    End If
    bModificar.Enabled = bEst
    
End Sub

Private Sub CargoMenuPlantillas()
On Error GoTo errStart
Dim iCont As Integer
    Screen.MousePointer = 11
    tbExplorer.Buttons("plantilla").ButtonMenus.Clear
    Cons = "Select PlaCodigo, PlaNombre from Plantilla Where PlaCodigo IN (" & prmPlantillasArtStock & ") Order by PlaNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        MnuPlantillas.Visible = False
    End If
    iCont = 0
    Do While Not RsAux.EOF
        With tbExplorer.Buttons("plantilla").ButtonMenus
            .Add
            .Item(.Count).Tag = CStr(RsAux!PlaCodigo)
            .Item(.Count).Text = Trim(RsAux!PlaNombre)
        End With
        If iCont <> 0 Then
            Load MnuPlaIndex(iCont)
        End If
        MnuPlaIndex(iCont).Caption = Trim(RsAux!PlaNombre)
        MnuPlaIndex(iCont).Tag = RsAux!PlaCodigo
        iCont = iCont + 1
        RsAux.MoveNext
    Loop
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
    
errStart:
    clsGeneral.OcurrioError "Error al cargar las plantillas.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tbExplorer_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    
    Select Case Button.Key
        Case "bback": Call tbExplorer_ButtonMenuClick(tbExplorer.Buttons("bback").ButtonMenus.Item(1))
        Case "bnext": Call tbExplorer_ButtonMenuClick(tbExplorer.Buttons("bnext").ButtonMenus.Item(1))
        
        Case "refresh": AccionRefrescar
        
        Case "preview": AccionPreview
        
        Case "print": AccionImprimir True
        
        Case "firstpage": IrAPagina vsListado, 1
        Case "previouspage"
            If vsListado.Visible Then
                IrAPagina vsListado, vsListado.PreviewPage - 1
            Else
                GetArticuloRegistro False
            End If
        Case "nextpage"
            If vsListado.Visible Then
                IrAPagina vsListado, vsListado.PreviewPage + 1
            Else
                'Voy al primer código mayor.
                GetArticuloRegistro True
            End If
        Case "lastpage": IrAPagina vsListado, vsListado.PageCount
    End Select
    
End Sub

Private Sub tbExplorer_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim aIdx As Long
    DoEvents
    
    Select Case ButtonMenu.Parent.Key
        Case "bback", "bnext"
                    aIdx = Val(ButtonMenu.Tag)
                    BuscoArticuloPorID arrClientes(aIdx).Codigo, False, False
                    AccionConsultar
                    
        Case "plantilla"
            If Val(tArticulo.Tag) <> 0 Then EjecutarApp App.Path & "\appExploreMsg.exe ", Val(ButtonMenu.Tag) & ":" & tArticulo.Tag
    End Select
    
End Sub

Private Sub MenuExplorer()
    On Error GoTo errMnu
    Dim aIdx As Integer, miX As Long

    miX = arrcli_Item(lIDArticulo)
    
    tbExplorer.Buttons("bback").ButtonMenus.Clear
    tbExplorer.Buttons("bnext").ButtonMenus.Clear
    
    For aIdx = LBound(arrClientes) To UBound(arrClientes)
        If arrClientes(aIdx).Codigo <> lIDArticulo And arrClientes(aIdx).Codigo <> 0 Then
            If aIdx < miX Then
                With tbExplorer.Buttons("bback").ButtonMenus
                    .Add Index:=1
                    .Item(1).Tag = CStr(aIdx)
                    .Item(1).Text = Trim(arrClientes(aIdx).Nombre)
                End With
                
            Else
                With tbExplorer.Buttons("bnext").ButtonMenus
                    .Add 'Index:=1
                    .Item(.Count).Tag = CStr(aIdx)
                    .Item(.Count).Text = Trim(arrClientes(aIdx).Nombre)
                End With
            End If
        End If
    Next
    
    If tbExplorer.Buttons("bback").ButtonMenus.Count = 0 Then tbExplorer.Buttons("bback").Enabled = False Else tbExplorer.Buttons("bback").Enabled = True
    If tbExplorer.Buttons("bnext").ButtonMenus.Count = 0 Then tbExplorer.Buttons("bnext").Enabled = False Else tbExplorer.Buttons("bnext").Enabled = True
    
    If lIDArticulo = 0 Then tbExplorer.Buttons("refresh").Enabled = False Else tbExplorer.Buttons("refresh").Enabled = True
errMnu:
End Sub

Private Sub InicializoToolbar()
On Error Resume Next
    
    Set tbExplorer.ImageList = imgExplore
    tbExplorer.Buttons("bback").Image = "back"
    tbExplorer.Buttons("bnext").Image = "next"
    tbExplorer.Buttons("refresh").Image = "refresh"
    tbExplorer.Buttons("print").Image = "print"
    tbExplorer.Buttons("preview").Image = "preview"
    tbExplorer.Buttons("plantilla").Image = "plantilla"

    tbExplorer.Buttons("firstpage").Image = "firstpage"
    tbExplorer.Buttons("previouspage").Image = "previouspage"
    tbExplorer.Buttons("nextpage").Image = "nextpage"
    tbExplorer.Buttons("lastpage").Image = "lastpage"
        
    With tbExplorer
        .Buttons("firstpage").Visible = False
        .Buttons("lastpage").Visible = False
        
        .Buttons("previouspage").ToolTipText = "Primer código de artículo menor."
        .Buttons("nextpage").ToolTipText = "Primer código de artículo mayor."
        
        .Buttons("nextpage").Enabled = False
        .Buttons("previouspage").Enabled = False
        
    End With
    
End Sub

Private Sub ButtonRegistros(ByVal bPage As Boolean)
    
    With tbExplorer
        If .Buttons("firstpage").Visible <> bPage Then .Buttons("firstpage").Visible = bPage
        If .Buttons("lastpage").Visible <> bPage Then .Buttons("lastpage").Visible = bPage
        
        If bPage Then
            '.Buttons("firstpage").ToolTipText = "Primera página."
            '.Buttons("lastpage").ToolTipText = "Última página."
            .Buttons("previouspage").ToolTipText = "Página anterior."
            .Buttons("nextpage").ToolTipText = "Página siguiente."
        Else
            .Buttons("previouspage").ToolTipText = "Primer código de artículo menor."
            .Buttons("nextpage").ToolTipText = "Primer código de artículo mayor."
            
            If Val(tArticulo.Tag) = 0 Then
                If .Buttons("previouspage").Enabled = True Then .Buttons("previouspage").Enabled = False
                If .Buttons("nextpage").Enabled = True Then .Buttons("nextpage").Enabled = False
            Else
                If .Buttons("previouspage").Enabled = False Then .Buttons("previouspage").Enabled = True
                If .Buttons("nextpage").Enabled = False Then .Buttons("nextpage").Enabled = True
            End If
        End If
        .Refresh
    End With
    
End Sub

Private Sub AccionPreview()
    
    If Not vsListado.Visible Then
        ButtonRegistros True
        AccionImprimir
        
        vsConsulta.Visible = False
        vsLocal.Visible = False
        imgHorizontal.Visible = False
        
        vsListado.Visible = True
        vsListado.ZOrder 0
        tbExplorer.Buttons("preview").Value = tbrPressed
        MnuOpPreview.Checked = True
    Else
        ButtonRegistros False
        vsConsulta.ZOrder 0
        vsLocal.ZOrder 0
        vsListado.Visible = False
        picHorizontal.ZOrder 0
        tbExplorer.Buttons("preview").Value = tbrUnpressed
        MnuOpPreview.Checked = False
        vsConsulta.Visible = True
        vsLocal.Visible = True
        imgHorizontal.Visible = True
    End If
    Me.Refresh
    
End Sub

Private Sub AccionRefrescar()
    If lIDArticulo <> 0 Then
        BuscoArticuloPorID lIDArticulo, False, False
        AccionConsultar
    End If
End Sub

Private Sub GetArticuloRegistro(ByVal bNext As Boolean)
    BuscoArticuloPorID lIDArticulo, bNext, Not bNext
    If Val(tArticulo.Tag) > 0 Then AccionConsultar
End Sub

Private Sub MenuPlantilla(ByVal bEnabled As Boolean)
    
    tbExplorer.Buttons("plantilla").Enabled = bEnabled
    MnuPlantillas.Enabled = bEnabled
    
End Sub
