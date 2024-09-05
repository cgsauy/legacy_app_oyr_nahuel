VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACOMBO.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Begin VB.Form frmListado 
   BackColor       =   &H00C0C000&
   Caption         =   "Ingreso de Existencia"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8400
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
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   3075
      Left            =   60
      TabIndex        =   15
      Top             =   1800
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5424
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
      SelectionMode   =   0
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
      Height          =   3135
      Left            =   60
      TabIndex        =   30
      Top             =   1800
      Width           =   6075
      _Version        =   196608
      _ExtentX        =   10716
      _ExtentY        =   5530
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
      Align           =   2  'Align Bottom
      BackColor       =   &H00C0C000&
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   8340
      TabIndex        =   31
      Top             =   6105
      Width           =   8400
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   900
         Picture         =   "frmListado.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Cancelar. [Ctrl+C]"
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bGrabar 
         Height          =   310
         Left            =   540
         Picture         =   "frmListado.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Grabar. [Ctrl+G]"
         Top             =   60
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0ACE
         Height          =   310
         Left            =   4020
         Picture         =   "frmListado.frx":0BD0
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   3180
         Picture         =   "frmListado.frx":1102
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2820
         Picture         =   "frmListado.frx":11EC
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   2400
         Picture         =   "frmListado.frx":12D6
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3660
         Picture         =   "frmListado.frx":1510
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bSalir 
         Height          =   310
         Left            =   4560
         Picture         =   "frmListado.frx":1612
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bNuevo 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":1714
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Nuevo. [Ctrl+N]"
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1680
         Picture         =   "frmListado.frx":1816
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   2040
         Picture         =   "frmListado.frx":1B58
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   60
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   1320
         Picture         =   "frmListado.frx":1E5A
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   60
         Width           =   310
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   6540
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   14314
            Key             =   "msg"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      BackColor       =   &H00C0C000&
      Caption         =   "Ingreso de datos."
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
      Height          =   1695
      Left            =   60
      TabIndex        =   28
      Top             =   0
      Width           =   8115
      Begin VB.TextBox tCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   5040
         MaxLength       =   5
         TabIndex        =   5
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox tArea 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   3
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox tComentario 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         MaxLength       =   100
         TabIndex        =   12
         Top             =   960
         Width           =   6915
      End
      Begin VB.TextBox tUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   2
         TabIndex        =   14
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox tCantidad 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4140
         MaxLength       =   6
         TabIndex        =   8
         Top             =   600
         Width           =   555
      End
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   600
         Width           =   3195
      End
      Begin AACombo99.AACombo cEstado 
         Height          =   315
         Left            =   5400
         TabIndex        =   10
         Top             =   600
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
      Begin AACombo99.AACombo cLocal 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
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
      Begin AACombo99.AACombo cCaso 
         Height          =   315
         Left            =   7080
         TabIndex        =   32
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
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
         BackStyle       =   0  'Transparent
         Caption         =   "N° Ingreso:"
         Height          =   255
         Left            =   4140
         TabIndex        =   4
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Á&rea:"
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Co&mentario:"
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuario:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Estado:"
         Height          =   255
         Left            =   4800
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Local:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   495
      End
   End
   Begin ComctlLib.ImageList ImgList 
      Left            =   5820
      Top             =   180
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
            Picture         =   "frmListado.frx":2094
            Key             =   "Alta"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmListado.frx":23AE
            Key             =   "Baja"
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
    AccionCancelar
End Sub

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub bImprimir_Click()
    AccionImprimir True
End Sub

Private Sub bNuevo_Click()
    AccionNuevo
End Sub

Private Sub bPrimero_Click()
    IrAPagina vsListado, 1
End Sub

Private Sub bSalir_Click()
    Unload Me
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
Private Sub cCaso_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cCaso.ListIndex > -1 Then InsertoRenglon
End Sub

Private Sub cEstado_GotFocus()
    If cEstado.Text = "" Then BuscoCodigoEnCombo cEstado, CLng(paEstadoArticuloEntrega)
    With cEstado
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione el estado de la mercadería."
End Sub
Private Sub cEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cCaso
End Sub
Private Sub cEstado_LostFocus()
    cEstado.SelStart = 0: Ayuda ""
End Sub

Private Sub cLocal_Change()
    vsConsulta.Rows = 1
End Sub
Private Sub cLocal_GotFocus()
    With cLocal
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Seleccione el local donde ingresará la mercadería."
End Sub
Private Sub cLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tArea
End Sub
Private Sub cLocal_LostFocus()
    cLocal.SelStart = 0
    Ayuda ""
End Sub

Private Sub chVista_Click()
    If chVista.Value = 0 Then
        vsConsulta.ZOrder 0
        Me.Refresh
    Else
        AccionImprimir False
        vsListado.ZOrder 0
        Me.Refresh
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0: Me.Refresh
End Sub
Private Sub Form_Load()
On Error GoTo ErrLoad
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    FechaDelServidor
    Cons = "Select LocCodigo, LocNombre From Local Order by LocNombre"
    CargoCombo Cons, cLocal
    Cons = "Select EsMCodigo, EsMAbreviacion From EstadoMercaderia Order by EsMAbreviacion"
    CargoCombo Cons, cEstado
    cCaso.Clear
    cCaso.AddItem "Alta": cCaso.ItemData(cCaso.NewIndex) = 1
    cCaso.AddItem "Baja": cCaso.ItemData(cCaso.NewIndex) = 0
    AccionCancelar
    'Hoja Carta
    vsListado.Orientation = orPortrait: vsListado.PaperSize = 1
    InicializoGrillas
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub
Private Sub InicializoGrillas()
    On Error Resume Next
    With vsConsulta
        .Editable = True
        .Redraw = False
        .WordWrap = False
        .Cols = 1: .Rows = 1
        .FormatString = "Tipo|>Cantidad|<Artículo|<Estado|"
        .ColWidth(0) = 800: .ColWidth(1) = 800: .ColWidth(2) = 3100: .ColWidth(3) = 800
        .Redraw = True
    End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            
            Case vbKeyP: IrAPagina vsListado, 1
            Case vbKeyA: IrAPagina vsListado, vsListado.PreviewPage - 1
            Case vbKeyS: IrAPagina vsListado, vsListado.PreviewPage + 1
            Case vbKeyU: IrAPagina vsListado, vsListado.PageCount
            
            Case vbKeyAdd: Zoom vsListado, vsListado.Zoom + 5
            Case vbKeySubtract: Zoom vsListado, vsListado.Zoom - 5
            
            Case vbKeyI: AccionImprimir True
                
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            Case vbKeyN: AccionNuevo
            Case vbKeyC: AccionCancelar
            Case vbKeyG: AccionGrabar
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11

    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + 70 + picBotones.Height)
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

Private Sub Label1_Click()
    Foco cLocal
End Sub

Private Sub Label2_Click()
    Foco cEstado
End Sub

Private Sub Label6_Click()
    Foco tCodigo
End Sub

Private Sub Label7_Click()
    Foco tComentario
End Sub
Private Sub Label8_Click()
    Foco tUsuario
End Sub


Private Sub tArea_GotFocus()
    With tArea
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese el área del local."
End Sub

Private Sub tArea_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tCodigo
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = 0
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese el mes y año de liquidación a consultar."
End Sub
Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrAP
    Screen.MousePointer = 11
    If KeyAscii = vbKeyReturn Then
        If Trim(tArticulo.Text) <> "" Then
            If Val(tArticulo.Tag) <> 0 Then Foco tCantidad: Exit Sub
            If IsNumeric(tArticulo.Text) Then
                BuscoArticuloPorCodigo tArticulo.Text
            Else
                BuscoArticuloPorNombre tArticulo.Text
            End If
            If Val(tArticulo.Tag) > 0 Then Foco tCantidad
        Else
            If vsConsulta.Rows > 1 Then Foco tComentario
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
    With tCantidad: .SelStart = 0: .SelLength = Len(.Text): End With
    Ayuda "Ingrese la cantidad de artículos."
End Sub

Private Sub tCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tCantidad.Text) Then If CLng(tCantidad.Text) > 0 Then Foco cEstado
    End If
End Sub

Private Sub tCodigo_GotFocus()
    With tCodigo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese el nro. de ingreso."
End Sub

Private Sub tCodigo_KeyPress(KeyAscii As Integer)
Dim aValor As Long

    If KeyAscii = vbKeyReturn Then
        If cLocal.ListIndex > -1 And Trim(tArea.Text) <> "" Then
            
            If vsConsulta.Rows > 1 And bNuevo.Enabled = True Then MsgBox "Ud. modifico el código de ingreso, se eliminaran los datos existentes.", vbExclamation, "ATENCIÓN"
        
            Cons = "Select * From BMLocal, BMRenglon, Articulo, EstadoMercaderia " _
                & " Where BMLLocal = " & cLocal.ItemData(cLocal.ListIndex) _
                & " And BMLArea = '" & tArea.Text & "' And BMLCodigo = " & CInt(tCodigo.Text) _
                & " And BMLId = BMRIdBML And BMRArticulo = ArtID And BMREstado = ESMCodigo"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            Do While Not RsAux.EOF
                With vsConsulta
                    If RsAux!BMRcantidad < 0 Then
                        .AddItem "Baja"
                        .Cell(flexcpData, .Rows - 1, 0) = 0
                        .Cell(flexcpPicture, .Rows - 1, 0) = ImgList.ListImages("Baja").ExtractIcon
                    Else
                        .AddItem "Alta"
                        .Cell(flexcpData, .Rows - 1, 0) = 1
                        .Cell(flexcpPicture, .Rows - 1, 0) = ImgList.ListImages("Alta").ExtractIcon
                    End If
                    aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 1) = aValor
                    aValor = RsAux!ESMCodigo: .Cell(flexcpData, .Rows - 1, 2) = aValor
                    .Cell(flexcpText, .Rows - 1, 1) = RsAux!BMRcantidad
                    .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!ArtNombre)
                    .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!ESMNombre)
                    .Cell(flexcpData, .Rows - 1, 3) = 1     'Me digo que estaba en el stock.
                End With
                RsAux.MoveNext
            Loop
            RsAux.Close
            Foco tArticulo
        End If
    End If
End Sub

Private Sub tComentario_GotFocus()
    tComentario.SelStart = 0
    tComentario.SelLength = Len(tComentario.Text)
    Ayuda " Ingrese un comentario."
End Sub
Private Sub tComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then tUsuario.SetFocus
End Sub
Private Sub tComentario_LostFocus()
    Ayuda ""
End Sub
Private Sub tUsuario_GotFocus()
    tUsuario.SelStart = 0: tUsuario.SelLength = Len(tUsuario.Text)
    Status.SimpleText = " Ingrese su código de Usuario."
End Sub

Private Sub tUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        tUsuario.Tag = 0
        If IsNumeric(tUsuario.Text) Then
            tUsuario.Tag = BuscoUsuarioDigito(CInt(tUsuario.Text), True)
            If Val(tUsuario.Tag) > 0 Then AccionGrabar Else tUsuario.Tag = vbNullString
        Else
            AccionGrabar
        End If
    End If
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionImprimir(Imprimir As Boolean)
Dim Consulta As Boolean
    On Error GoTo errImprimir
    
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    If vsConsulta.Rows > 1 Then

        With vsListado
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        
        vsListado.filename = "Ingreso de Mercadería de Existencia"
        EncabezadoListado vsListado, "Ingreso de Mercadería de Existencia al " & Format(gFechaServidor, FormatoFP), False
        
        If Trim(tComentario.Text) <> "" Then vsListado.Paragraph = "Comentario:= " & Trim(tComentario.Text): vsListado.Paragraph = ""
        If Val(tUsuario.Tag) > 0 Then vsListado.Paragraph = "Usuario:= " & BuscoUsuario(tUsuario.Tag, True): vsListado.Paragraph = ""
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        vsListado.EndDoc
        
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

Private Sub AccionGrabar()
Dim IdBML As Long

    If vsConsulta.Rows = 1 Then MsgBox "No hay artículos ingresados.", vbExclamation, "ATENCIÓN": Exit Sub
    If Val(tUsuario.Tag) = 0 Then MsgBox "Debe ingresar su código de usuario.", vbExclamation, "ATENCIÓN": Exit Sub
    If Not clsGeneral.TextoValido(tComentario.Text) Then MsgBox "Se ingreso un carácter no válido en el comentario.", vbExclamation, "ATENCIÓN": Exit Sub
    If Not clsGeneral.TextoValido(tArea.Text) Then MsgBox "Se ingreso un carácter no válido en el comentario.", vbExclamation, "ATENCIÓN": Exit Sub
    If cLocal.ListIndex = -1 Then MsgBox "Debe seleccionar un local.", vbExclamation, "ATENCIÓN": Foco cLocal: Exit Sub
    If Trim(tArea.Text) = "" Then MsgBox "Debe ingresar un área del local.", vbExclamation, "ATENCIÓN": Foco tArea: Exit Sub
    
    If MsgBox("¿Confirma almacenar los datos ingresados?", vbQuestion + vbYesNo, "GRABAR") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    On Error GoTo ErrBT
    FechaDelServidor
    cBase.BeginTrans
    On Error GoTo ErrRB
    
    IdBML = 0
    Cons = "Select * From BMLocal " _
        & " Where BMLLocal = " & cLocal.ItemData(cLocal.ListIndex) _
        & " And BMLArea = '" & tArea.Text & "' And BMLCodigo = " & CInt(tCodigo.Text)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        IdBML = RsAux!BMLId
        RsAux.Edit
        If Trim(tComentario.Text) = "" Then RsAux!BMLComentario = Null Else RsAux!BMLComentario = Trim(tComentario.Text)
        RsAux!BMLFModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsAux!BMLUsuario = tUsuario.Tag
        RsAux.Update
        RsAux.Close
    Else
        RsAux.AddNew
        RsAux!BMLLocal = cLocal.ItemData(cLocal.ListIndex)
        RsAux!BMLArea = Trim(tArea.Text)
        RsAux!BMLCodigo = Trim(tCodigo.Text)
        If Trim(tComentario.Text) = "" Then RsAux!BMLComentario = Null Else RsAux!BMLComentario = Trim(tComentario.Text)
        RsAux!BMLFModificacion = Format(gFechaServidor, sqlFormatoFH)
        RsAux!BMLUsuario = tUsuario.Tag
        RsAux.Update
        RsAux.Close
        
        Cons = "Select MAX(BMLId) From BMLocal"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        IdBML = RsAux(0)
        RsAux.Close
    End If
    
    
    With vsConsulta
        For I = 1 To .Rows - 1
            
            'Grabo en la tabla del local.
            Cons = "Select * From BMRenglon Where BMRIdBML = " & IdBML _
                & " And BMRArticulo = " & CLng(.Cell(flexcpData, I, 1)) & " And BMREstado = " & CLng(.Cell(flexcpData, I, 2))
            
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            
            If CLng(.Cell(flexcpData, I, 0)) = 0 Then
                'Es baja.
                If RsAux.EOF Then
                    RsAux.Close
                    cBase.RollbackTrans
                    Screen.MousePointer = 0
                    MsgBox "No hay tantos artículos " & Trim(.Cell(flexcpText, I, 2)) & " para dar de baja en los datos del local ingresados.", vbCritical, "ATENCIÓN"
                    Exit Sub
                Else
                    If RsAux!BMRcantidad < CInt(.Cell(flexcpText, I, 1)) Then
                        RsAux.Close
                        cBase.RollbackTrans
                        Screen.MousePointer = 0
                        MsgBox "No hay tantos artículos " & Trim(.Cell(flexcpText, I, 2)) & " para dar de baja para los datos ingresados.", vbCritical, "ATENCIÓN"
                        Exit Sub
                    Else
                        If RsAux!BMRcantidad = CInt(.Cell(flexcpText, I, 1)) Then
                            RsAux.Delete
                        Else
                            RsAux.Edit
                            RsAux!BMRcantidad = RsAux!BMRcantidad - CInt(.Cell(flexcpText, I, 1))
                            RsAux.Update
                        End If
                        RsAux.Close
                    End If
                End If
            Else
                If RsAux.EOF Then
                    RsAux.AddNew
                    RsAux!BMRIdBML = IdBML
                    RsAux!BMRArticulo = CLng(.Cell(flexcpData, I, 1))
                    RsAux!BMREstado = CLng(.Cell(flexcpData, I, 2))
                    RsAux!BMRcantidad = CLng(.Cell(flexcpText, I, 1))
                Else
                    RsAux.Edit
                    RsAux!BMRcantidad = RsAux!BMRcantidad + CLng(.Cell(flexcpText, I, 1))
                End If
                RsAux.Update
                RsAux.Close
            End If
        Next I
    End With
    cBase.CommitTrans
    If chVista.Value = 1 Then
        AccionImprimir False
        vsListado.ZOrder 0
    Else
        chVista.Value = 1
    End If
    Me.Refresh
    If MsgBox("Desea imprimir los datos ingresados?", vbQuestion + vbYesNo, "Imprimir") = vbYes Then
        frmSetup.pControl = vsListado
        frmSetup.Show vbModal, Me
        Me.Refresh
        If frmSetup.pOK Then vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
    End If
    DeshabilitoIngreso
    bNuevo.Enabled = True: bGrabar.Enabled = False: bCancelar.Enabled = False
    Screen.MousePointer = 0
    Exit Sub
ErrBT:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar la transacción.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
ErrRB:
    Resume ErrResumo
ErrResumo:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al intentar grabar los datos, reintente.", Err.Description
    Screen.MousePointer = 0
    Exit Sub
End Sub
Private Sub DeshabilitoIngreso()
    cLocal.ListIndex = -1: cLocal.Enabled = False: cLocal.BackColor = Inactivo
    tArea.Text = "": tArea.Enabled = False: tArea.BackColor = Inactivo
    tCodigo.Text = "": tCodigo.Enabled = False: tCodigo.BackColor = Inactivo
    tArticulo.Text = "": tArticulo.Enabled = False: tArticulo.BackColor = Inactivo
    tCantidad.Text = vbNullString: tCantidad.Enabled = False: tCantidad.BackColor = Inactivo
    cEstado.Enabled = False: cEstado.BackColor = Inactivo: cEstado.ListIndex = -1
    tComentario.BackColor = Inactivo: tComentario.Enabled = False: tComentario.Text = vbNullString
    vsConsulta.Rows = 1
    tUsuario.Enabled = False: tUsuario.BackColor = Inactivo: tUsuario.Text = vbNullString: tUsuario.Tag = vbNullString
    cCaso.Enabled = False: cCaso.BackColor = Inactivo
End Sub
Private Sub HabilitoIngreso()
    cLocal.Enabled = True: cLocal.BackColor = obligatorio
    tArea.Text = "": tArea.Enabled = True: tArea.BackColor = obligatorio
    tArticulo.Enabled = True: tArticulo.BackColor = obligatorio
    tCodigo.Text = "": tCodigo.Enabled = True: tCodigo.BackColor = obligatorio
    tCantidad.Text = vbNullString: tCantidad.Enabled = True: tCantidad.BackColor = obligatorio
    cEstado.Enabled = True: cEstado.BackColor = obligatorio
    tComentario.BackColor = Blanco: tComentario.Enabled = True: tComentario.Text = vbNullString
    tUsuario.Enabled = True: tUsuario.BackColor = obligatorio: tUsuario.Text = vbNullString: tUsuario.Tag = vbNullString
    cCaso.Enabled = True: cCaso.BackColor = obligatorio
End Sub

Private Sub AccionCancelar()
    bNuevo.Enabled = True: bGrabar.Enabled = False: bCancelar.Enabled = False
    DeshabilitoIngreso
End Sub

Private Sub AccionNuevo()
    chVista.Value = 0
    bNuevo.Enabled = False: bGrabar.Enabled = True: bCancelar.Enabled = True
    HabilitoIngreso
    Foco tArticulo  'Chinura para que agarre la primera vez
    cLocal.SetFocus
End Sub
Private Sub InsertoRenglon()
On Error GoTo ErrControl
Dim aValor As Integer

    If Val(tArticulo.Tag) = 0 Then MsgBox "No hay seleccionado un artículo.", vbExclamation, "ATENCIÓN": Foco tArticulo: Exit Sub
    If Not IsNumeric(tCantidad.Text) Then MsgBox "La cantidad ingresada no es correcta.", vbInformation, "ATENCIÓN": Foco tCantidad: Exit Sub
    If CInt(tCantidad.Text) < 1 Then MsgBox "La cantidad ingresada no es correcta.", vbInformation, "ATENCIÓN": Foco tCantidad: Exit Sub
    If cEstado.ListIndex = -1 Then MsgBox "No hay seleccionado un estado.", vbInformation, "ATENCIÓN": Foco cEstado: Exit Sub
    
    'Busco en la grilla si ya se ingreso este artículo para ese estado.
    With vsConsulta
        For I = 1 To .Rows - 1
            If CLng(.Cell(flexcpData, I, 1)) = CLng(tArticulo.Tag) And CLng(.Cell(flexcpData, I, 2)) = cEstado.ItemData(cEstado.ListIndex) Then
                'Veo si es alta o baja lo que ingreso.
                If CInt(.Cell(flexcpData, I, 0)) = 1 Then
                    'Es alta.
                    If cCaso.ItemData(cCaso.ListIndex) = 1 Then
                        If MsgBox("El artículo ya tiene un alta de " & Trim(.Cell(flexcpText, I, 1)) & " artículos." & Chr(13) _
                                & "¿Desea agregarle la cantidad ingresada?", vbQuestion + vbYesNo, "Artículo Ingresado") = vbNo Then tArticulo.Text = "": tCantidad.Text = "": cEstado.Text = "": Foco tArticulo: Exit Sub
                        
                        .Cell(flexcpText, I, 1) = CInt(.Cell(flexcpText, I, 1)) + CInt(tCantidad.Text)
                        tArticulo.Text = "": tCantidad.Text = "": cEstado.Text = "": Foco tArticulo: Exit Sub
                    
                    Else
                        If MsgBox("El artículo ya tiene un alta de " & Trim(.Cell(flexcpText, I, 1)) & " artículos." & Chr(13) _
                                & "¿Desea restarle la cantidad ingresada?", vbQuestion + vbYesNo, "Artículo Ingresado") = vbNo Then tArticulo.Text = "": tCantidad.Text = "": cEstado.Text = "": Foco tArticulo: Exit Sub
                        
                        'Antes de cambiarlo a baja verifico que el final exista en el local
                        aValor = HayStockLocal
                        If aValor < Abs(CInt(.Cell(flexcpText, I, 1)) - CInt(tCantidad.Text)) Then MsgBox "No hay tantos artículos para dar de baja con ese estado en el local seleccionado.", vbInformation, "ATENCIÓN": Exit Sub
                        If CInt(.Cell(flexcpText, I, 1)) - CInt(tCantidad.Text) < 0 Then
                            .Cell(flexcpText, I, 0) = "Baja": .Cell(flexcpData, I, 0) = 0
                            .Cell(flexcpPicture, I, 0) = ImgList.ListImages("Baja").ExtractIcon
                        End If
                        .Cell(flexcpText, I, 1) = Abs(CInt(.Cell(flexcpText, I, 1)) - CInt(tCantidad.Text))
                        If CInt(.Cell(flexcpText, I, 1)) = 0 Then .RemoveItem I
                        tArticulo.Text = "": tCantidad.Text = "": cEstado.Text = "": Foco tArticulo
                        Exit Sub
                    End If
                Else
                    'Es baja
                    If cCaso.ItemData(cCaso.ListIndex) = 0 Then
                        If MsgBox("El artículo ya tiene una Baja de " & Trim(.Cell(flexcpText, I, 1)) & " artículos." & Chr(13) _
                                & "¿Desea restarle la cantidad ingresada?", vbQuestion + vbYesNo, "Artículo Ingresado") = vbNo Then tArticulo.Text = "": tCantidad.Text = "": cEstado.Text = "": Foco tArticulo: Exit Sub
                        
                        If (CInt(.Cell(flexcpText, I, 1)) * -1) + CInt(tCantidad.Text) > 0 Then
                            .Cell(flexcpText, I, 0) = "Alta": .Cell(flexcpData, I, 0) = 1
                            .Cell(flexcpPicture, I, 0) = ImgList.ListImages("Alta").ExtractIcon
                        End If
                        .Cell(flexcpText, I, 1) = Abs((CInt(.Cell(flexcpText, I, 1)) * -1) + CInt(tCantidad.Text))
                        tArticulo.Text = "": tCantidad.Text = "": cEstado.Text = "": Foco tArticulo: Exit Sub
                    Else
                        If MsgBox("El artículo ya tiene una Baja de " & Trim(.Cell(flexcpText, I, 1)) & " artículos." & Chr(13) _
                                & "¿Desea agregarle la cantidad ingresada?", vbQuestion + vbYesNo, "Artículo Ingresado") = vbNo Then tArticulo.Text = "": tCantidad.Text = "": cEstado.Text = "": Foco tArticulo: Exit Sub
                        aValor = HayStockLocal
                        If aValor < CInt(.Cell(flexcpText, I, 1)) + CInt(tCantidad.Text) Then MsgBox "No hay tantos artículos para dar de baja con ese estado en el local seleccionado.", vbInformation, "ATENCIÓN": Exit Sub
                        .Cell(flexcpText, I, 1) = CInt(.Cell(flexcpText, I, 1)) + CInt(tCantidad.Text)
                        tArticulo.Text = "": tCantidad.Text = "": cEstado.Text = "": Foco tArticulo: Exit Sub
                    End If
                End If
            End If
        Next I
    End With

    If cCaso.ItemData(cCaso.ListIndex) = 0 Then
        aValor = HayStockLocal
        If aValor < CInt(tCantidad.Text) Then MsgBox "No hay tantos artículos para dar de baja con ese estado.", vbInformation, "ATENCIÓN": Exit Sub
    End If
    
    Screen.MousePointer = 11
    With vsConsulta
        .AddItem ""
        '1 Es Alta , 0 Es Baja
        If cCaso.ItemData(cCaso.ListIndex) = 1 Then
            .Cell(flexcpText, .Rows - 1, 0) = "Alta"
            .Cell(flexcpData, .Rows - 1, 0) = 1
            .Cell(flexcpPicture, .Rows - 1, 0) = ImgList.ListImages("Alta").ExtractIcon
        Else
            .Cell(flexcpText, .Rows - 1, 0) = "Baja"
            .Cell(flexcpData, .Rows - 1, 0) = 0
            .Cell(flexcpPicture, .Rows - 1, 0) = ImgList.ListImages("Baja").ExtractIcon
        End If
        
        .Cell(flexcpData, .Rows - 1, 1) = tArticulo.Tag
        .Cell(flexcpData, .Rows - 1, 2) = cEstado.ItemData(cEstado.ListIndex)
        
        .Cell(flexcpText, .Rows - 1, 1) = CInt(tCantidad.Text)
        .Cell(flexcpText, .Rows - 1, 2) = Trim(tArticulo.Text)
        .Cell(flexcpText, .Rows - 1, 3) = Trim(cEstado.Text)
        .Cell(flexcpData, .Rows - 1, 3) = 0     'Me digo que lo ingreso ahora.
    End With
    tArticulo.Text = "": tCantidad.Text = "": cEstado.Text = "": cCaso.Text = "": Foco tArticulo
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrControl:
    Screen.MousePointer = vbDefault
    clsGeneral.OcurrioError "Ocurrio un error al intentar insertar el renglón en la lista.", Err.Description
End Sub

Private Function HayStockLocal() As Integer
On Error GoTo ErrBSL
    HayStockLocal = 0
    Screen.MousePointer = vbHourglass
    Cons = "Select * from BMLocal, BMRenglon" _
        & " Where BMLLocal = " & cLocal.ItemData(cLocal.ListIndex) & " And BMLArea = '" & tArea.Text & "'" _
        & " And BMRArticulo = " & tArticulo.Tag _
        & " And BMREstado = " & cEstado.ItemData(cEstado.ListIndex) & " And BMLID = BMRIdBML"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then HayStockLocal = RsAux!BMRcantidad
    RsAux.Close
    Screen.MousePointer = vbDefault
    Exit Function
ErrBSL:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el stock del local.", Err.Description
    Screen.MousePointer = 0
End Function

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
        tArticulo.Text = Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
        tArticulo.Tag = RsAux!ArtID
        RsAux.Close
    End If
    Screen.MousePointer = 0

End Sub

Private Sub BuscoArticuloPorNombre(NomArticulo As String)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
Dim Resultado As Long

    Screen.MousePointer = 11
    
    Cons = "Select ArtId, Código = ArtCodigo, Nombre = ArtNombre from Articulo" _
        & " Where ArtNombre LIKE '" & NomArticulo & "%'" _
        & " Order By ArtNombre"
            
    Dim LiAyuda As New clsListadeAyuda
    LiAyuda.ActivoListaAyuda Cons, False, cBase.Connect
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



