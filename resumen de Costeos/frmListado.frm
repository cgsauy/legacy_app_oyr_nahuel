VERSION 5.00
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListado 
   Caption         =   "Resumen de Costeos"
   ClientHeight    =   7530
   ClientLeft      =   1950
   ClientTop       =   2070
   ClientWidth     =   10200
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
   ScaleWidth      =   10200
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   8280
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8LCtl.VSFlexGrid vsConsulta 
      Height          =   1095
      Left            =   2520
      TabIndex        =   26
      Top             =   3360
      Width           =   2655
      _cx             =   4683
      _cy             =   1931
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
      Rows            =   50
      Cols            =   10
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
      Height          =   4455
      Left            =   120
      TabIndex        =   10
      Top             =   1080
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
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   11595
      TabIndex        =   11
      Top             =   6720
      Width           =   11655
      Begin VB.CommandButton butExcel 
         Height          =   310
         Left            =   5400
         Picture         =   "frmListado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Exportar a excel"
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0784
         Height          =   310
         Left            =   4500
         Picture         =   "frmListado.frx":0886
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4140
         Picture         =   "frmListado.frx":0DB8
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   3180
         Picture         =   "frmListado.frx":1232
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2820
         Picture         =   "frmListado.frx":131C
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   2220
         Picture         =   "frmListado.frx":1406
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3780
         Picture         =   "frmListado.frx":1640
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4860
         Picture         =   "frmListado.frx":1742
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5940
         Picture         =   "frmListado.frx":1B08
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":1C0A
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmListado.frx":1F0C
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1860
         Picture         =   "frmListado.frx":224E
         Style           =   1  'Graphical
         TabIndex        =   13
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
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6600
         TabIndex        =   24
         Top             =   120
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   476
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lQ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   150
         Width           =   375
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   7275
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
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
            Object.Width           =   9763
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
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   10335
      Begin VB.ComboBox cCHasta 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cCDesde 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   255
         Width           =   4215
      End
      Begin VB.TextBox tGrupo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Al"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   675
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Costeos del "
         Height          =   195
         Left            =   180
         TabIndex        =   0
         Top             =   315
         Width           =   915
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Grupo:"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   645
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RsAux As rdoResultset, rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean

Dim arrCosteo() As String

Private Sub AccionLimpiar()
    
    cCDesde.ListIndex = -1: cCHasta.ListIndex = -1
    tArticulo.Text = ""
    tGrupo.Text = ""
    vsConsulta.Rows = 1
    lQ.Caption = ""
    
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

Private Sub cCDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cCHasta.SetFocus
End Sub

Private Sub cCHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tArticulo
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
    Foco tArticulo
End Sub

Private Sub Label2_Click()
    cCHasta.SetFocus
End Sub

Private Sub Label3_Click()
    tGrupo.SetFocus
End Sub

Private Sub Label5_Click()
    cCDesde.SetFocus
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = 0
End Sub

Private Sub tArticulo_GotFocus()
    With tArticulo: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tArticulo_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrTA
    
    If KeyCode = vbKeyReturn Then
        If Trim(tArticulo.Text) = "" Then Foco tGrupo: Exit Sub
        If Val(tArticulo.Tag) <> 0 Then Foco tGrupo: Exit Sub
        
        Screen.MousePointer = 11
        tArticulo.Text = Replace(Trim(tArticulo.Text), " ", "%")
        
        If Not IsNumeric(tArticulo.Text) Then
            Cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtNombre Like '" & tArticulo.Text & "%'"
        Else
            Cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtCodigo = " & CLng(tArticulo.Text)
        End If
        
        Dim aQ As Integer, aIDSel As Long, aTexto As String
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aQ = 1
            aIDSel = RsAux!ArtID: aTexto = Format(RsAux(1), "(#,000,000)") & " " & Trim(RsAux!Nombre)
            RsAux.MoveNext: If Not RsAux.EOF Then aQ = 2: aIDSel = 0
        End If
        RsAux.Close
    
        Select Case aQ
            Case 0: MsgBox "No hay datos que coincidan con el texto ingersado.", vbExclamation, "No hay datos"
            
            Case 2:
                        Dim miLista As New clsListadeAyuda
                        aIDSel = miLista.ActivarAyuda(cBase, Cons, 4000, 1, "Lista de Datos")
                        Me.Refresh
                        If aIDSel > 0 Then
                            aIDSel = miLista.RetornoDatoSeleccionado(0)
                            aTexto = Format(miLista.RetornoDatoSeleccionado(1), "(#,000,000)")
                            aTexto = aTexto & " " & miLista.RetornoDatoSeleccionado(2)
                        End If
                        Set miLista = Nothing
        End Select
    
        If aIDSel > 0 Then
            tArticulo.Text = Trim(aTexto)
            tArticulo.Tag = aIDSel
        End If
        
        Screen.MousePointer = 0
    End If
    Exit Sub

ErrTA:
    clsGeneral.OcurrioError "Error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()

    On Error GoTo ErrLoad

    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    
    AccionLimpiar
    
    InicializoGrillas
    CargoCombos
    
    bCargarImpresion = True
    With vsListado
        .PhysicalPage = True
        .PaperSize = 1
        .Orientation = orPortrait
        .Zoom = 100
        .MarginLeft = 900: .MarginRight = 350
        .MarginBottom = 750: .MarginTop = 750
    End With
    
    Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
                    
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .OutlineBar = flexOutlineBarNone ' flexOutlineBarComplete
        .OutlineCol = 0
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<Artículo|Costeo|>Cantidad|>$ Compra|>$ Venta|>Ganancia|>% Ganancia|>Promedio %|"
            
        .WordWrap = False
        AnchoEncabezado Pantalla:=True
        .MergeCells = flexMergeSpill
        .MergeCol(0) = True
        
    End With
      
End Sub

Private Sub CargoCombos()

    ReDim arrCosteo(0): arrCosteo(0) = ""
    
    'Cargo Combos de Costeos        --------------------------------------------------------------------------------
    Cons = "Select CabID, CabMesCosteo from CMCabezal Order by CabMesCosteo Desc"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not RsAux.EOF
        cCDesde.AddItem Format(RsAux!CabMesCosteo, "MM/YYYY")
        cCDesde.ItemData(cCDesde.NewIndex) = RsAux!CabID
        
        cCHasta.AddItem Format(RsAux!CabMesCosteo, "MM/YYYY")
        cCHasta.ItemData(cCHasta.NewIndex) = RsAux!CabID
        
        If arrCosteo(0) <> "" Then
            ReDim Preserve arrCosteo(UBound(arrCosteo) + 1)
        End If
        arrCosteo(UBound(arrCosteo)) = RsAux!CabID & "|" & Format(RsAux!CabMesCosteo, "mm/yyyy")
        
            
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If cCDesde.ListCount > 0 Then cCDesde.ListIndex = cCDesde.ListCount - 1
    If cCHasta.ListCount > 0 Then cCHasta.ListIndex = 0
    '------------------------------------------------------------------------------------------------------------------------

End Sub

Private Function getFechaCosteo(mCosteo As Long) As String

Dim I As Integer, sData() As String
    getFechaCosteo = " "
    For I = LBound(arrCosteo) To UBound(arrCosteo)
        sData = Split(arrCosteo(I), "|")
        If sData(0) = mCosteo Then
            getFechaCosteo = sData(1)
            Exit For
        End If
    Next
    
End Function

Private Sub AnchoEncabezado(Optional Pantalla As Boolean = False, Optional Impresora As Boolean = False)

    '<Artículo|Costeo|>Q Costeada|>$ Compra|>$ Venta|>Ganancia|
    With vsConsulta
        
        If Pantalla Then
            .ColWidth(0) = 3200: .ColWidth(1) = 850: .ColWidth(2) = 1000
            .ColWidth(3) = 1600: .ColWidth(4) = 1600: .ColWidth(5) = 1400
        End If
        
        If Impresora Then
            .ColWidth(0) = 0: .ColWidth(1) = 600: .ColWidth(2) = 570 '430
            .ColWidth(3) = 1000: .ColWidth(4) = 1000: .ColWidth(5) = 1000
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
    
    picBotones.Width = vsListado.Width
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 150
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub AccionConsultar()
On Error GoTo errConsultar
    
    If Not ValidoCampos Then Exit Sub
    
Dim aIDCDesde As Long, aIDCHasta As Long
Dim aArticulo As Long, aTxtArticulo As String

Dim aCostoP As Currency, aVentaP As Currency, aCantidadP As Long
Dim aCostoT As Currency, aVentaT As Currency

Dim rsSer As rdoResultset
Dim mValorX As Currency

    Screen.MousePointer = 11
    bCargarImpresion = True
    aArticulo = 0
    lQ.Tag = 0: lQ.Caption = ""
    aIDCDesde = cCDesde.ItemData(cCDesde.ListIndex)
    aIDCHasta = cCHasta.ItemData(cCHasta.ListIndex)
    
    Dim aQ As Long
    aQ = 0
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    Cons = "Select Count(Distinct(CosArticulo)) From CMCosteo, Articulo " & _
               " Where CosArticulo = ArtID" & _
               " And CosID Between " & aIDCDesde & " And " & aIDCHasta
               
    If Val(tArticulo.Tag) <> 0 Then Cons = Cons & " And CosArticulo = " & Val(tArticulo.Tag)
    If Val(tGrupo.Tag) <> 0 Then Cons = Cons & " And CosArticulo In (Select AGrArticulo from ArticuloGrupo Where AGrGrupo =" & Val(tGrupo.Tag) & ")"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then aQ = RsAux(0)
    RsAux.Close
    
    If aQ = 0 Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        Screen.MousePointer = 0:  Exit Sub
    End If
    pbProgreso.Max = aQ
    '-------------------------------------------------------------------------------------------------------------------------------------------
    
    Cons = "Select CosID, ArtCodigo, ArtNombre, Sum(Abs(CosCantidad)) as CosCantidad, " & _
                                    " Sum(CosCosto * CosCantidad) as CosCosto, Sum(CosVenta * CosCantidad) as CosVenta " & _
               " From CMCosteo, Articulo " & _
               " Where CosArticulo = ArtID" & _
               " And CosID Between " & aIDCDesde & " And " & aIDCHasta
    
    If Val(tArticulo.Tag) <> 0 Then Cons = Cons & " And CosArticulo = " & Val(tArticulo.Tag)
    If Val(tGrupo.Tag) <> 0 Then Cons = Cons & " And CosArticulo In (Select AGrArticulo from ArticuloGrupo Where AGrGrupo =" & Val(tGrupo.Tag) & ")"
    
    Cons = Cons & _
               " Group by CosID, ArtCodigo, ArtNombre" & _
               " Order by ArtCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        RsAux.Close: Screen.MousePointer = 0: InicializoGrillas: Exit Sub
    End If
    
    aCostoP = 0: aVentaP = 0: aCantidadP = 0
    aCostoT = 0: aVentaT = 0
    
    With vsConsulta
        .Rows = 1: .Refresh: .Redraw = False
        Do While Not RsAux.EOF
            
            If aArticulo <> RsAux!ArtCodigo Then      '--------------------------------------------------------------
                If aArticulo <> 0 Then
                    pbProgreso.Value = pbProgreso.Value + 1
                    InsertoTotalArticulo aTxtArticulo, aCantidadP, aCostoP, aVentaP
                    
                    aCostoT = aCostoT + aCostoP
                    aVentaT = aVentaT + aVentaP
                    aCostoP = 0: aVentaP = 0: aCantidadP = 0
                End If
                
                aTxtArticulo = "(" & Format(RsAux!ArtCodigo, "#,000,000") & ") " & Trim(RsAux!ArtNombre)
                aArticulo = RsAux!ArtCodigo
                
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Trim(aTxtArticulo)
                '.Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Azul
            End If  '---------------------------------------------------------------------------------------------------------
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Trim(aTxtArticulo)
            .Cell(flexcpText, .Rows - 1, 1) = getFechaCosteo(RsAux!CosID) 'Format(rsAux!CosID, "dd/mm/yy")
            
            .Cell(flexcpText, .Rows - 1, 2) = RsAux!CosCantidad
            .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!CosCosto, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!CosVenta, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 5) = Format((RsAux!CosVenta - RsAux!CosCosto), FormatoMonedaP)
            
            aCantidadP = aCantidadP + .Cell(flexcpValue, .Rows - 1, 2)
            aCostoP = aCostoP + .Cell(flexcpValue, .Rows - 1, 3)
            aVentaP = aVentaP + .Cell(flexcpValue, .Rows - 1, 4)
            
            '% Ganancia (Gan /Venta * 100)
            mValorX = Format((RsAux!CosVenta - RsAux!CosCosto), "0.00")
            If RsAux!CosVenta <> 0 Then mValorX = mValorX / RsAux!CosVenta * 100 Else mValorX = mValorX * 100
            .Cell(flexcpText, .Rows - 1, 6) = Format(mValorX, "##0.00")
            
            'Prom % el porc / Qvendida
            .Cell(flexcpText, .Rows - 1, 7) = Format(mValorX / RsAux!CosCantidad, "##0.00")
                        
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        If aArticulo <> 0 Then
            pbProgreso.Value = pbProgreso.Value + 1
            InsertoTotalArticulo aTxtArticulo, aCantidadP, aCostoP, aVentaP
            aCostoT = aCostoT + aCostoP
            aVentaT = aVentaT + aVentaP
        End If
        
        .Select 1, 0, 1, 0
        .Sort = flexSortGenericAscending
        
        InsertoTotalArticulo "Total General", 0, aCostoT, aVentaT, True
        
        
        'Recorro para sacar los títulos iniciales (estaban x el orden)
        For I = 1 To .Rows - 1
            If .Cell(flexcpText, I, 1) = "" Then .Cell(flexcpText, I, 0) = ""
        Next
        .Redraw = True
    End With
    
    pbProgreso.Value = 0
    aCostoT = 0: aVentaT = 0
    
    'Hay que cargar el Total p c/u de los Costeos   -----------------------------------------------------------------------
    If vsConsulta.Rows > 1 Then vsConsulta.AddItem ""
    
    Cons = "Select CosID, " & _
                        " Sum(CosCosto * CosCantidad) as CosCosto, Sum(CosVenta * CosCantidad) as CosVenta " & _
               " From CMCosteo " & _
               " Where CosID Between " & aIDCDesde & " And " & aIDCHasta & _
               " Group by CosID " & _
               " Order by CosID"
               
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        Do While Not RsAux.EOF
            With vsConsulta
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = getFechaCosteo(RsAux!CosID)
                
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!CosCosto, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!CosVenta, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 5) = Format((RsAux!CosVenta - RsAux!CosCosto), "#,##0.00")
                
                aCostoT = aCostoT + .Cell(flexcpValue, .Rows - 1, 3)
                aVentaT = aVentaT + .Cell(flexcpValue, .Rows - 1, 4)
                
                '% Ganancia (Gan /Venta * 100)
                mValorX = Format((RsAux!CosVenta - RsAux!CosCosto), "0.00")
                If RsAux!CosVenta <> 0 Then mValorX = mValorX / RsAux!CosVenta * 100 Else mValorX = mValorX * 100
                .Cell(flexcpText, .Rows - 1, 6) = Format(mValorX, "##0.00")
            End With
    
            RsAux.MoveNext
        Loop
        
        With vsConsulta
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = "Total"
            
            .Cell(flexcpText, .Rows - 1, 3) = Format(aCostoT, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 4) = Format(aVentaT, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 5) = Format(aVentaT - aCostoT, FormatoMonedaP)
            
            '% Ganancia (Gan /Venta * 100)
            mValorX = Format((aVentaT - aCostoT), "0.00")
            If aVentaT <> 0 Then mValorX = mValorX / aVentaT * 100 Else mValorX = mValorX * 100
            .Cell(flexcpText, .Rows - 1, 6) = Format(mValorX, "##0.00")
                
            .Cell(flexcpBackColor, .Rows - 1, 1, , .Cols - 1) = Colores.Rojo: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
        End With
    End If
    RsAux.Close
    
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    vsConsulta.Redraw = True
    clsGeneral.OcurrioError "Error al realizar la consulta de datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub InsertoTotalArticulo(Articulo As String, Q As Long, Costo As Currency, Venta As Currency, Optional TGeneral As Boolean = False)

    With vsConsulta
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = Trim(Articulo)
        .Cell(flexcpText, .Rows - 1, 1) = "Sub Total"
        
        If Not TGeneral Then .Cell(flexcpText, .Rows - 1, 2) = Q
        .Cell(flexcpText, .Rows - 1, 3) = Format(Costo, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 4) = Format(Venta, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 5) = Format(Venta - Costo, FormatoMonedaP)
        
        If Not TGeneral Then
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Azul
        Else
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
        End If
                
    End With
    
End Sub

Private Sub tGrupo_Change()
    tGrupo.Tag = 0
End Sub

Private Sub tGrupo_GotFocus()
    tGrupo.SelStart = 0: tGrupo.SelLength = Len(tGrupo.Text)
End Sub

Private Sub tGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrTA
    
    If KeyCode = vbKeyReturn Then
        If Trim(tGrupo.Text) = "" Then Foco bConsultar: Exit Sub
        If Val(tGrupo.Tag) <> 0 Then Foco bConsultar: Exit Sub
        
        Screen.MousePointer = 11
        tGrupo.Text = Replace(Trim(tGrupo.Text), " ", "%")
        Cons = "Select GruCodigo, GruNombre as Grupo From Grupo Where GruNombre Like '" & tGrupo.Text & "%' Order by GruNombre"
        
        Dim aQ As Integer, aIDSel As Long, aTexto As String
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aQ = 1
            aIDSel = RsAux!GruCodigo: aTexto = Trim(RsAux!Grupo)
            RsAux.MoveNext: If Not RsAux.EOF Then aQ = 2: aIDSel = 0
        End If
        RsAux.Close
    
        Select Case aQ
            Case 0: MsgBox "No hay datos que coincidan con el texto ingersado.", vbExclamation, "No hay datos"
            
            Case 2:
                        Dim miLista As New clsListadeAyuda
                        aIDSel = miLista.ActivarAyuda(cBase, Cons, 4000, 1, "Lista de Datos")
                        Me.Refresh
                        If aIDSel > 0 Then
                            aIDSel = miLista.RetornoDatoSeleccionado(0)
                            aTexto = miLista.RetornoDatoSeleccionado(1)
                        End If
                        Set miLista = Nothing
        End Select
    
        If aIDSel > 0 Then
            tGrupo.Text = Trim(aTexto)
            tGrupo.Tag = aIDSel
        End If
        
        Screen.MousePointer = 0
    End If
    Exit Sub

ErrTA:
    clsGeneral.OcurrioError "Error al buscar los grupos de artículos.", Err.Description
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
            .Columns = 1
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        aTexto = "Resumen de Costeos"
        
        EncabezadoListado vsListado, aTexto, False
        vsListado.FileName = "Resumen de Costeos"
         
        With vsConsulta
        '    .Redraw = False
        '    .FontSize = 6
        '    AnchoEncabezado Impresora:=True
            .ExtendLastCol = False: vsListado.RenderControl = .hwnd: vsConsulta.ExtendLastCol = True
        '    AnchoEncabezado Pantalla:=True
        '    .FontSize = 8
        '    .Redraw = True
        End With
        
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
    clsGeneral.OcurrioError "Error al realizar la impresión", Err.Description
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

Private Sub vsListado_NewPage()
    lQ.Tag = Val(lQ.Tag) + 1
    lQ.Caption = lQ.Tag: lQ.Refresh
End Sub

Private Function ValidoCampos() As Boolean

    ValidoCampos = False
    
    If cCDesde.ListIndex = -1 Then
        MsgBox "Seleccione el costeo para realizar la consulta de resumen.", vbExclamation, "Falta Costeo Desde"
        cCDesde.SetFocus: Exit Function
    End If
    
    If cCHasta.ListIndex = -1 Then
        MsgBox "Seleccione el costeo para realizar la consulta de resumen.", vbExclamation, "Falta Costeo Hasta"
        cCHasta.SetFocus: Exit Function
    End If
    
    If cCHasta.ItemData(cCHasta.ListIndex) < cCDesde.ItemData(cCDesde.ListIndex) Then
        MsgBox "El rango de costeos seleccionado no es correcto.", vbExclamation, "Rango de Costeos Incorrecto"
        cCDesde.SetFocus: Exit Function
    End If
    
    ValidoCampos = True
    
End Function
