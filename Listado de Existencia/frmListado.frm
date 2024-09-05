VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmListado 
   Caption         =   "Listado de Existencias"
   ClientHeight    =   7530
   ClientLeft      =   1170
   ClientTop       =   2220
   ClientWidth     =   13080
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
   ScaleWidth      =   13080
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   4440
      TabIndex        =   23
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
      TabIndex        =   26
      Top             =   1020
      Width           =   9555
      _Version        =   196608
      _ExtentX        =   16854
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
      EmptyColor      =   -2147483632
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   6075
      TabIndex        =   27
      Top             =   6720
      Width           =   6135
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   19
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   18
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
         TabIndex        =   21
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
         TabIndex        =   22
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
         TabIndex        =   11
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   12
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
      TabIndex        =   25
      Top             =   7275
      Width           =   13080
      _ExtentX        =   23072
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
            Object.Width           =   14870
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
      Height          =   975
      Left            =   60
      TabIndex        =   24
      Top             =   0
      Width           =   10575
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4260
         TabIndex        =   10
         Top             =   600
         Width           =   3375
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cmbOpcion 
         Height          =   315
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox cCDesde 
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox tFHasta 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2340
         TabIndex        =   3
         Text            =   "28/12/2000"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4260
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1020
         TabIndex        =   1
         Text            =   "28/12/2000"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre:"
         Height          =   255
         Left            =   3540
         TabIndex        =   9
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Art.:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Respaldo:"
         Height          =   255
         Left            =   7800
         TabIndex        =   28
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "al"
         Height          =   255
         Left            =   2100
         TabIndex        =   2
         Top             =   300
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   3540
         TabIndex        =   4
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&F/Compra:"
         Height          =   255
         Left            =   120
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

Enum TipoCV
    Compra = 1
    Comercio = 2
    Importacion = 3
End Enum

Private rsAux As rdoResultset, rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean

Private Sub AccionLimpiar()
    tFecha.Text = "": tFHasta.Text = ""
    tArticulo.Text = "": tArticulo.Tag = "0"
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
        vsConsulta.ZOrder 0
    Else
        AccionImprimir
        vsListado.ZOrder 0
    End If

End Sub

Private Sub cmbOpcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cmbTipo
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco txtNombre
End Sub

Private Sub Label1_Click()
    Foco tArticulo
End Sub

Private Sub Label2_Click()
    Foco tFHasta
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = "0"
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrTA

    If KeyCode = vbKeyReturn And Trim(tArticulo.Text) <> "" Then
        tArticulo.Tag = "0"
        Screen.MousePointer = 11
        If Not IsNumeric(tArticulo.Text) Then
            cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtNombre Like '" & tArticulo.Text & "%'"
            Dim LiAyuda  As New clsListadeAyuda
            LiAyuda.ActivoListaAyuda cons, False, miConexion.TextoConexion(logComercio)
            If LiAyuda.ItemSeleccionado <> "" Then tArticulo.Text = LiAyuda.ItemSeleccionado Else tArticulo.Text = "0"
            Set LiAyuda = Nothing
        End If
        If CLng(tArticulo.Text) > 0 Then
            cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtCodigo = " & CLng(tArticulo.Text)
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurReadOnly)
            If rsAux.EOF Then
                rsAux.Close
                MsgBox "No se encontró un artículo con ese código.", vbInformation, "ATENCIÓN"
            Else
                tArticulo.Text = Trim(rsAux!Nombre)
                tArticulo.Tag = rsAux!Artid
                rsAux.Close
                Foco cmbOpcion
            End If
        Else
            tArticulo.Text = "0"
        End If
        Screen.MousePointer = 0
    Else
        If KeyCode = vbKeyReturn Then Foco cmbOpcion
    End If
    Exit Sub
ErrTA:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0


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
        tFecha.Text = Format(tFecha.Text, FormatoFP)
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
    InicializoGrillas
    CargoCosteos
    AccionLimpiar
    bCargarImpresion = True
    
    cmbOpcion.AddItem "No listar Repuestos con Costo 0", 0
    cmbOpcion.AddItem "Listar sólo Repuestos con Costo 0", 1
    cmbOpcion.AddItem "Listar Todos los Artículos", 2
    cmbOpcion.ListIndex = 0
    
    'Cargo Tipos de Articulos
    cons = "SELECT TipCodigo, TipNombre FROM Tipo ORDER BY TipNombre"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not rsAux.EOF
        cmbTipo.AddItem rsAux!TipNombre
        cmbTipo.ItemData(cmbTipo.NewIndex) = rsAux!TipCodigo
        rsAux.MoveNext
    Loop
    rsAux.Close

    
    With vsListado
        .PhysicalPage = True
        .PaperSize = 1
        .Orientation = orPortrait
        .Zoom = 100
        .MarginLeft = 800: .MarginRight = 250
        .MarginBottom = 750: .MarginTop = 750
    End With
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub CargoCosteos()

    cCDesde.AddItem ""
    cCDesde.ItemData(cCDesde.NewIndex) = 0
    
    'Cargo Combos de Costeos        --------------------------------------------------------------------------------
    cons = "Select CabID, CabMesCosteo from CMCabezal Order by CabMesCosteo Desc"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not rsAux.EOF
        cCDesde.AddItem Format(rsAux!CabMesCosteo, "MM/YYYY")
        cCDesde.ItemData(cCDesde.NewIndex) = rsAux!CabID
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    cCDesde.ListIndex = -1
    '------------------------------------------------------------------------------------------------------------------------

End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .OutlineBar = flexOutlineBarNone ' = flexOutlineBarComplete
        .OutlineCol = 0
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<|<Artículo|<Compra|<Fecha|>Q|>Costo|>Total|>TC|>U$S|"
            
        .WordWrap = False
        AnchoEncabezado Pantalla:=True
        .MergeCol(0) = True: .ColAlignment(0) = flexAlignLeftBottom
        .MergeCells = flexMergeSpill
        
    End With
      
End Sub

Private Sub AnchoEncabezado(Optional Pantalla As Boolean = False, Optional Impresora As Boolean = False)

    With vsConsulta
        
        If Pantalla Then
            .ColWidth(0) = 0: .ColWidth(1) = 2700: .ColWidth(2) = 650: .ColWidth(3) = 800: .ColWidth(4) = 1100
            .ColWidth(5) = 1400: .ColWidth(6) = 1400: .ColWidth(7) = 600: .ColWidth(8) = 1300: .ColWidth(9) = 10
        End If
        
        If Impresora Then
            .ColWidth(0) = 0: .ColWidth(1) = 1300: .ColWidth(2) = 460: .ColWidth(3) = 570: .ColWidth(4) = 500
            .ColWidth(5) = 650: .ColWidth(6) = 900: .ColWidth(7) = 300: .ColWidth(8) = 650: .ColWidth(9) = 10
        End If
        
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
    Set miConexion = Nothing
    End
    
End Sub

Private Sub AccionConsultar()

    If Not IsDate(tFecha.Text) And IsDate(tFHasta.Text) Then
        MsgBox "Ingrese la fecha desde.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Sub
    End If
    If IsDate(tFecha.Text) And Not IsDate(tFHasta.Text) Then
        If Trim(tFHasta.Text) = "" Then
            tFHasta.Text = tFecha.Text
        Else
            MsgBox "La fecha hasta no es correcta.", vbExclamation, "ATENCIÓN"
            Foco tFHasta: Exit Sub
        End If
    End If
    If IsDate(tFecha.Text) And IsDate(tFHasta.Text) Then
        If CDate(tFecha.Text) > CDate(tFHasta.Text) Then
            MsgBox "Los rangos de fecha no son correctos.", vbExclamation, "ATENCIÓN"
            Foco tFecha: Exit Sub
        End If
    End If
    On Error GoTo errConsultar
    Screen.MousePointer = 11
    bCargarImpresion = True
    
    cons = "Select ArtID, ArtCodigo, ArtNombre, ArtClase, 'IDCompra' = Compra.ComCodigo, 'Fecha' = CM.ComFecha, ComCantidad, ComCosto, CM.ComTipo, ComTC "
    If cCDesde.ListIndex > 0 Then
        cons = cons & " From Articulo, rspCMCompra CM" _
                        & " Left Outer Join Compra On Compra.ComCodigo = CM.ComCodigo" _
                & " Where ComArticulo = ArtID " _
                & " And ComCosteo = " & cCDesde.ItemData(cCDesde.ListIndex)
    Else
        cons = cons & " From Articulo, CMCompra CM" _
                        & " Left Outer Join Compra On Compra.ComCodigo = CM.ComCodigo" _
                & " Where ComArticulo = ArtID "
    End If
    If IsDate(tFecha.Text) Then
        cons = cons & " And CM.ComFecha >= '" & Format(tFecha.Text & " 00:00:00", sqlFormatoFH) & "'" _
            & " And CM.ComFecha <= '" & Format(tFHasta.Text & " 23:59:59", sqlFormatoFH) & "'"
    End If
    
    If Val(tArticulo.Tag) > 0 Then cons = cons & " And ComArticulo = " & CLng(tArticulo.Tag)
    If cmbTipo.ListIndex <> -1 Then cons = cons & " And ArtTipo = " & cmbTipo.ItemData(cmbTipo.ListIndex)
    
    If txtNombre.Text <> "" Then cons = cons & " And ArtNombre LIKE '" & Replace(txtNombre.Text, " ", "%") & "%'"
    
    cons = cons & " Order by ArtNombre, Fecha DESC"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        rsAux.Close: Screen.MousePointer = 0: InicializoGrillas: Exit Sub
    End If
    
    Dim aIdAnterior As Long: aIdAnterior = 0
    Dim xClase As Long, QCompras As Integer
    Dim pblnAdd As Boolean
    
    With vsConsulta
        .Rows = 1
        Do While Not rsAux.EOF
            
            If aIdAnterior <> rsAux!Artid Then
                If aIdAnterior <> 0 And QCompras = 0 Then
                    .RemoveItem (.Rows - 1)
                End If
                QCompras = 0
                aIdAnterior = rsAux!Artid
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
                If Not IsNull(rsAux!ArtClase) Then xClase = rsAux!ArtClase Else xClase = -1
            End If
            
            '0- No listar Repuestos con Costo 0
            '1- Listar sólo Repuestos con Costo 0
            '2- Listar Todos los Artículos

            Select Case cmbOpcion.ListIndex
                Case 0: pblnAdd = Not (rsAux!ComCosto = 0 And xClase = prmClaseRepuesto)
                Case 1: pblnAdd = (rsAux!ComCosto = 0 And xClase = prmClaseRepuesto)
                Case 2: pblnAdd = True
            End Select
            
            'If Not (rsAux!ComCosto = 0 And (xClase = prmClaseRepuesto) And (cNoCosto0.Value = vbChecked)) Then
            If pblnAdd Then
                QCompras = QCompras + 1
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 1) = " "
                If rsAux!ComTipo = TipoCV.Compra Then .Cell(flexcpText, .Rows - 1, 2) = rsAux!IdCompra
                
                .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!Fecha, "dd/mm/yy")
                .Cell(flexcpText, .Rows - 1, 4) = rsAux!ComCantidad
                .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!ComCosto, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 6) = Format(rsAux!ComCantidad * rsAux!ComCosto, FormatoMonedaP)
                If rsAux!ComTipo = TipoCV.Compra Then
                    .Cell(flexcpText, .Rows - 1, 7) = Format(rsAux!ComTC, FormatoMonedaP)
                    .Cell(flexcpText, .Rows - 1, 8) = Format(CCur(.Cell(flexcpText, .Rows - 1, 6)) / rsAux!ComTC, FormatoMonedaP)
                End If
            End If
            
            rsAux.MoveNext
        Loop
        rsAux.Close
        
        .Select 1, 0, 1, 1
        .Sort = flexSortGenericAscending
        
        .Subtotal flexSTSum, 0, 4, "0", Colores.Gris, Colores.Rojo, False, "%s"
        .Subtotal flexSTSum, 0, 6: .Subtotal flexSTSum, 0, 8
        .AddItem ""
        .Subtotal flexSTSum, -1, 6, , Colores.Gris, Colores.Rojo, False, "Total Existencias"
        .Subtotal flexSTSum, -1, 8
    End With
    
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tFHasta_GotFocus()
    With tFHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tFHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tArticulo
End Sub

Private Sub tFHasta_LostFocus()
    If IsDate(tFHasta.Text) Then tFHasta.Text = Format(tFHasta.Text, FormatoFP)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco bConsultar
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
'            .Columns = 2
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        aTexto = Trim(tFecha.Text)
        If Trim(tFHasta.Text) <> "" Then aTexto = aTexto & " al " & Trim(tFHasta.Text)
        If Trim(aTexto) = "" Then
        
            If cCDesde.ListIndex > 0 Then
                aTexto = "al " & Format(UltimoDia(CDate(cCDesde.Text)), "dd Mmmm yyyy")
            Else
                cons = "select Max(CabMesCosteo) from CMCabezal"
                Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                If Not rsAux.EOF Then
                    If Not IsNull(rsAux(0)) Then aTexto = "al " & Format(UltimoDia(rsAux(0)), "dd Mmmm yyyy")
                End If
                rsAux.Close
            End If
        End If
        aTexto = "Listado de Existencias  " & aTexto
        EncabezadoListado vsListado, aTexto, False
        vsListado.FileName = "Listado de Existencia"
            
        With vsConsulta
            .Redraw = False
'            .FontSize = 6
'            AnchoEncabezado Impresora:=True
            vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
'            AnchoEncabezado Pantalla:=True
'            .FontSize = 8
            .Redraw = True
        End With
        
        vsListado.EndDoc
        'bCargarImpresion = False
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
    Status.Panels(4).Text = strTexto
End Sub
