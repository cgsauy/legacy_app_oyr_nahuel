VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmListado 
   Caption         =   "Diferencias Balance (Vista Simple - Corrige LIFO)"
   ClientHeight    =   7530
   ClientLeft      =   1530
   ClientTop       =   2040
   ClientWidth     =   11715
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
   ScaleWidth      =   11715
   Begin VB.CommandButton bAyuda 
      Caption         =   "Ayuda"
      Height          =   315
      Left            =   8940
      TabIndex        =   23
      Top             =   20
      Width           =   915
   End
   Begin VB.CheckBox cVerificado 
      Caption         =   "Eliminar Artículos Verificados."
      Height          =   195
      Left            =   1740
      TabIndex        =   22
      Top             =   300
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CheckBox cQCero 
      Caption         =   "Eliminar Totales en 0"
      Height          =   195
      Left            =   1740
      TabIndex        =   21
      Top             =   60
      Width           =   1815
   End
   Begin VB.ComboBox cEstado 
      Height          =   315
      Left            =   5220
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   20
      Width           =   3555
   End
   Begin VB.CheckBox chRep 
      Caption         =   "Repuestos"
      Height          =   195
      Left            =   60
      TabIndex        =   18
      Top             =   300
      Width           =   1335
   End
   Begin VB.CheckBox chArt 
      Caption         =   "Mercadería"
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   60
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4575
      Left            =   2880
      TabIndex        =   12
      Top             =   1320
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8070
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
      GridLinesFixed  =   4
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
      Height          =   5595
      Left            =   60
      TabIndex        =   14
      Top             =   540
      Width           =   9735
      _Version        =   196608
      _ExtentX        =   17171
      _ExtentY        =   9869
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
      ScaleWidth      =   8055
      TabIndex        =   15
      Top             =   6720
      Width           =   8115
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   8
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   7
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   0
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   5820
         TabIndex        =   16
         Top             =   120
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   476
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   7275
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
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
            AutoSize        =   2
            Key             =   "bdcomercio"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9895
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Filtros:"
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   60
      Width           =   675
   End
   Begin VB.Menu MnuAcciones 
      Caption         =   "MnuAcciones"
      Visible         =   0   'False
      Begin VB.Menu MnuTitulo 
         Caption         =   "Titulo"
      End
      Begin VB.Menu MnuL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPlantilla 
         Caption         =   "Ejecutar Plantilla Balance"
      End
      Begin VB.Menu MnuComentarios 
         Caption         =   "Ingresar Comentarios Inventario"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuL2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerCol 
         Caption         =   "Ocultar/Mostrar [p]"
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

Dim rsAux As rdoResultset, rs1 As rdoResultset
Dim rsRep As rdoResultset

Private aTexto As String
Dim bCargarImpresion As Boolean

Private Enum eCol
    Tipo = 0
    Articulo
    conSano
    conARec
    conRoto
    conARetEnv
    conTotal
    Lifo
    difConLifo
    LifoCosto
    UCFecha         'Ultima compra
    UCCosto
    Notas
End Enum

Private savNotas As Boolean

Private fBalance As Date

Private Sub AccionLimpiar()
    'vsConsulta.Rows = 2
End Sub

Private Sub bAyuda_Click()
    frmAyuda.Show , Me
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

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    
    zfn_CargofBalance
    
    picBotones.BorderStyle = vbBSNone
    InicializoGrillas
    AccionLimpiar
    bCargarImpresion = True
    
    With vsListado
        .PaperSize = 1
        .Orientation = orPortrait
        .PhysicalPage = True
        .Zoom = 100
        .MarginLeft = 500: .MarginRight = 250
        .MarginBottom = 550: .MarginTop = 550
    End With
    
    With cEstado
        .AddItem "(listar todos)", 0
        .AddItem "Última compra con costo cero", 1
        .AddItem "Última compra menor a 6 meses", 2
        .AddItem "Última compra entre 6 y 18 meses", 3
        .AddItem "Última compra mayor a 18 meses o LIFO = 0", 4
        .AddItem "Solamente artículos sin diferencias", 5
        .ListIndex = 0
    End With
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()
On Error Resume Next
    With vsConsulta
        .GridLinesFixed = flexGridFlatHorz
        .GridLines = flexGridFlatHorz
        .OutlineBar = flexOutlineBarNone ' = flexOutlineBarComplete
        .OutlineCol = 0
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        '.FormatString = "<Tipo|<Artículo|^Sistema (Locales)|^Sistema (Locales)|^Sistema (Locales)|^Sistema (Locales)|^Sistema (Locales)" & _
                                                        "|^Se Contó (Balance)|^Se Contó (Balance)|^Se Contó (Balance)|^Se Contó (Balance)|^Se Contó (Balance)" & _
                                                        "|TOTAL|>LIFO|>Ctdo.-Lifo|>Costo Aprox|ucFecha|ucCosto|<NOTAS"
                                                        
        .FormatString = "<Tipo|<Artículo" & _
                                         "|^Se Contó (Balance)|^Se Contó (Balance)|^Se Contó (Balance)|^Se Contó (Balance)|^Se Contó (Balance)" & _
                                         "|>LIFO|>Ctdo.-Lifo|>Costo Aprox|ucFecha|ucCosto|<NOTAS"
                                                        
            
        .WordWrap = False: .MergeCells = flexMergeSpill
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .ColHidden(eCol.UCCosto) = True: .ColHidden(eCol.UCFecha) = True
        .ColHidden(eCol.Tipo) = True
        .ColWidth(eCol.Articulo) = 2700

        .ColWidth(eCol.conSano) = 600: .ColWidth(eCol.conARec) = 500: .ColWidth(eCol.conRoto) = 400: .ColWidth(eCol.conARetEnv) = 900: .ColWidth(eCol.conTotal) = 700
        
        .ColWidth(eCol.difConLifo) = 900
        .ColWidth(eCol.Lifo) = 700: .ColWidth(eCol.Notas) = 1500
        
        .AddItem ""
       
        .Cell(flexcpText, .Rows - 1, eCol.conSano) = "SAN": .Cell(flexcpText, .Rows - 1, eCol.conARec) = "ARec": .Cell(flexcpText, .Rows - 1, eCol.conRoto) = "ROT"
        .Cell(flexcpText, .Rows - 1, eCol.conARetEnv) = "(Ret+Env)": .Cell(flexcpText, .Rows - 1, eCol.conTotal) = "Total"
                
        Dim iX As Integer, xCaption As String
        xCaption = "Ocultar/Mostrar '[p]'"
        With mnuVerCol
            iX = 0 '5: Load .Item(iX)
            .Item(iX).Caption = Replace(xCaption, "[p]", "contado SAN"): .Item(iX).Tag = eCol.conSano
            iX = iX + 1: Load .Item(iX)
            .Item(iX).Caption = Replace(xCaption, "[p]", "contado ARec"): .Item(iX).Tag = eCol.conARec
            iX = iX + 1: Load .Item(iX)
            .Item(iX).Caption = Replace(xCaption, "[p]", "contado ROT"): .Item(iX).Tag = eCol.conRoto
            iX = iX + 1: Load .Item(iX)
            .Item(iX).Caption = Replace(xCaption, "[p]", "contado ARetirarEnviar"): .Item(iX).Tag = eCol.conARetEnv
        End With
        '---------------------------------------------------------------------------------------------------------------------------
        
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, 1) = .BackColorFixed
        
        .Cell(flexcpBackColor, 0, eCol.conSano, .Rows - 1, eCol.conTotal) = Colores.Gris
        .Cell(flexcpBackColor, .Rows - 1, eCol.conSano, .Rows - 1, eCol.conARetEnv) = Colores.Blanco
                
        .Cell(flexcpBackColor, 0, eCol.Lifo, .Rows - 1, eCol.Lifo) = Colores.Blanco
        .Cell(flexcpBackColor, 0, eCol.difConLifo, .Rows - 1, eCol.UCCosto) = Colores.Gris
        .Cell(flexcpBackColor, 0, eCol.Notas, .Rows - 1, .Cols - 1) = Colores.Blanco
        
        .GridLines = flexGridFlat
        .FixedRows = 2
        .Editable = True
        
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
    
    vsListado.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Left = vsListado.Left
    
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
    CierroConexionBDMov
    
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Private Sub AccionConsultar()

    On Error GoTo errConsultar
    Dim aValor As Long, qAuxiliar As Long, sqlVerificados As String
    
    sqlVerificados = " Not IN (Select CInArticulo from ComentarioInventario Where CinVerificado is Not Null)"
    
    If chArt.Value = vbUnchecked And chArt.Value = chRep.Value Then
        MsgBox "Seleccione los artículos a incluir en la consulta", vbExclamation, "Falta selección de Artículos"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    bCargarImpresion = True
    vsConsulta.Rows = 2
    
    Dim aQ As Long, QSano As Long, aID As Long
    Dim aRow As Long
    Dim rsBal As rdoResultset
    aQ = 0
    
    'Cargo Artículos de la Tabla LIFO   ---------------------------------------------------------------------------------------------------------
    cons = "Select Count(Distinct(ComArticulo)) From CMCompra" & _
               " Where ComArticulo Not In (Select ArtId from Articulo where ArtTipo = " & paTipoArticuloServicio & ")"
    If cVerificado.Value = vbChecked Then cons = cons & " And ComArticulo " & sqlVerificados
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux(0)) Then aQ = rsAux(0)
    rsAux.Close
    
    If aQ > 0 Then
        pbProgreso.Min = 0: pbProgreso.Value = 0
        pbProgreso.Max = aQ
        
        cons = "Select  ArtID, ArtCodigo, ArtNombre, Sum(ComCantidad) as Q, Max(ComFecha) as UCompra" & _
                    " From CMCompra, Articulo " & _
                    " Where ComArticulo = ArtID" & _
                    " And ArtTipo <> " & paTipoArticuloServicio
        If cVerificado.Value = vbChecked Then cons = cons & " And ArtID " & sqlVerificados
        cons = cons & " Group by ArtID, ArtCodigo, ArtNombre"
                    
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        Do While Not rsAux.EOF
            pbProgreso.Value = pbProgreso.Value + 1
                
            With vsConsulta
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, eCol.Articulo) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
                If EsRepuesto(rsAux!ArtCodigo) Then .Cell(flexcpText, .Rows - 1, eCol.Tipo) = "Repuestos " Else .Cell(flexcpText, .Rows - 1, eCol.Tipo) = "Mercadería"
                aValor = rsAux!ArtID: .Cell(flexcpData, .Rows - 1, eCol.Articulo) = aValor
                'aValor = rsAux!ArtCodigo: .Cell(flexcpData, .Rows - 1, eCol.locSano) = aValor
                
                .Cell(flexcpText, .Rows - 1, eCol.Lifo) = Format(rsAux!Q, "#,##0")
                If Not IsNull(rsAux!UCompra) Then .Cell(flexcpText, .Rows - 1, eCol.UCFecha) = Format(rsAux!UCompra, "dd/mm/yyyy")
            End With
            rsAux.MoveNext
        Loop
        rsAux.Close
        
        pbProgreso.Min = 0: pbProgreso.Value = 0
        'Cargo el último costo de la ultima compra  ---
         cons = " Select ComArticulo, ComCosto From CMCompra, Articulo " & _
                    " Where ComArticulo = ArtID " & _
                    " And ArtTipo <> " & paTipoArticuloServicio & _
                    " And ComFecha = ( Select Max(ComFecha) from CMCompra as cmVal Where cmVal.ComArticulo = CMCompra.ComArticulo )"
        
        If cVerificado.Value = vbChecked Then cons = cons & " And ArtID " & sqlVerificados
                           
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        Do While Not rsAux.EOF
            If pbProgreso.Value < pbProgreso.Max Then pbProgreso.Value = pbProgreso.Value + 1
            
            aRow = RowEnLista(rsAux!ComArticulo)
            If aRow > 0 Then
                If Not IsNull(rsAux!ComCosto) Then
                    vsConsulta.Cell(flexcpText, aRow, eCol.UCCosto) = Format(rsAux!ComCosto, "0.00")
                    vsConsulta.Cell(flexcpData, aRow, eCol.UCCosto) = Format(rsAux!ComCosto, "0.00")
                End If
            End If
            rsAux.MoveNext
        Loop
        rsAux.Close
    End If
        
    'Cargo Artículos del STOCK CONTEO  ---------------------------------------------------------------------------------------------------------
    cons = "Select Count(Distinct(BMRArticulo)) from BMRenglon, BMLocal" & _
               " Where BMRArticulo Not In (Select ArtId from Articulo Where ArtTipo = " & paTipoArticuloServicio & ")" & _
               " And BMRIDBML = BMLID  And BMLCodigo = 1"
    If cVerificado.Value = vbChecked Then cons = cons & " And BMRArticulo " & sqlVerificados
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux(0)) Then aQ = rsAux(0)
    rsAux.Close
    
    If aQ > 0 Then
        pbProgreso.Min = 0: pbProgreso.Value = 0
        pbProgreso.Max = aQ
        
        cons = "Select ArtID, ArtCodigo, ArtNombre, BMREstado, Sum(BMRCantidad) as Q" & _
                    " From BMRenglon, Articulo, BMLocal" & _
                    " Where BMRArticulo = ArtID " & _
                    " And ArtTipo <> " & paTipoArticuloServicio & _
                    " And BMRIDBML = BMLID  And BMLCodigo = 1"
        If cVerificado.Value = vbChecked Then cons = cons & " And ArtID " & sqlVerificados
        cons = cons & " Group by ArtID, ArtCodigo, ArtNombre, BMREstado " & _
                    " Order by ArtID"
                    
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        aID = 0
        Do While Not rsAux.EOF
            
            If aID <> rsAux!ArtID Then
                aRow = RowEnLista(rsAux!ArtID)
                aID = rsAux!ArtID
                pbProgreso.Value = pbProgreso.Value + 1
            End If
            
            With vsConsulta
                If aRow = 0 Then
                    .AddItem ""
                    aRow = .Rows - 1
                    
                    .Cell(flexcpText, aRow, eCol.Articulo) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
                    If EsRepuesto(rsAux!ArtCodigo) Then .Cell(flexcpText, aRow, eCol.Tipo) = "Repuestos " Else .Cell(flexcpText, aRow, eCol.Tipo) = "Mercadería"
                    aValor = rsAux!ArtID: .Cell(flexcpData, aRow, eCol.Articulo) = aValor
                    'aValor = rsAux!ArtCodigo: .Cell(flexcpData, .Rows - 1, eCol.locSano) = aValor
                End If
                
                Select Case rsAux!BMREstado
                    Case paEstadoArticuloEntrega: .Cell(flexcpText, aRow, eCol.conSano) = Format(rsAux!Q, "#,##0")
                    Case paEstadoArticuloARecuperar: .Cell(flexcpText, aRow, eCol.conARec) = Format(rsAux!Q, "#,##0")
                    Case paEstadoArticuloRoto: .Cell(flexcpText, aRow, eCol.conRoto) = Format(rsAux!Q, "#,##0")
                End Select
            End With
            rsAux.MoveNext
        Loop
        rsAux.Close
    End If
    pbProgreso.Value = 0
    
    
    'Sumo Al CONTEO  lo de los Camiones, Eduardo y Compañia ---------------------------------------------------------------------------------------------------------
    cons = "Select Count(Distinct(StLArticulo)) from StockLocal" & _
               " Where StLArticulo Not In (Select ArtId from Articulo Where ArtTipo = " & paTipoArticuloServicio & ")" & _
               " And ( (StLTipoLocal = 2 And StLLocal in (" & paLocalCompania & ", " & paLocalEduardo & ")) " & _
                        " Or (StlTipoLocal = 1))"
                
    If cVerificado.Value = vbChecked Then cons = cons & " And StLArticulo " & sqlVerificados
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux(0)) Then aQ = rsAux(0)
    rsAux.Close
    
    If aQ > 0 Then
        pbProgreso.Min = 0: pbProgreso.Value = 0
        pbProgreso.Max = aQ
        
        cons = "Select ArtID, ArtCodigo, ArtNombre, StLEstado, Sum(StLCantidad) as Q " & _
                   " From StockLocal, Articulo" & _
                   " Where StLArticulo = ArtID " & _
                   " And ArtTipo <> " & paTipoArticuloServicio & _
                   " And ( (StLTipoLocal = 2 And StLLocal in (" & paLocalCompania & ", " & paLocalEduardo & ")) " & _
                        " Or (StlTipoLocal = 1))"
        If cVerificado.Value = vbChecked Then cons = cons & " And ArtID " & sqlVerificados
        cons = cons & " Group by ArtID, ArtCodigo, ArtNombre, StLEstado" & _
                   " Order by ArtID"
                    
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        aID = 0
        Do While Not rsAux.EOF
            
            If aID <> rsAux!ArtID Then
                aRow = RowEnLista(rsAux!ArtID)
                aID = rsAux!ArtID
                pbProgreso.Value = pbProgreso.Value + 1
            End If
            
            With vsConsulta
                If aRow = 0 Then
                    .AddItem ""
                    aRow = .Rows - 1
                    
                    .Cell(flexcpText, aRow, eCol.Articulo) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
                    If EsRepuesto(rsAux!ArtCodigo) Then .Cell(flexcpText, aRow, eCol.Tipo) = "Repuestos " Else .Cell(flexcpText, aRow, eCol.Tipo) = "Mercadería"
                    aValor = rsAux!ArtID: .Cell(flexcpData, aRow, eCol.Articulo) = aValor
                    'aValor = rsAux!ArtCodigo: .Cell(flexcpData, .Rows - 1, eCol.locSano) = aValor
                End If
                
                Select Case rsAux!StLEstado
                    Case paEstadoArticuloEntrega: .Cell(flexcpText, aRow, eCol.conSano) = .Cell(flexcpValue, aRow, eCol.conSano) + Format(rsAux!Q, "#,##0")
                    Case paEstadoArticuloARecuperar: .Cell(flexcpText, aRow, eCol.conARec) = .Cell(flexcpValue, aRow, eCol.conARec) + Format(rsAux!Q, "#,##0")
                    Case paEstadoArticuloRoto: .Cell(flexcpText, aRow, eCol.conRoto) = .Cell(flexcpValue, aRow, eCol.conRoto) + Format(rsAux!Q, "#,##0")
                End Select
            End With
            rsAux.MoveNext
        Loop
        rsAux.Close
    End If
    pbProgreso.Value = 0
    
    'Veo Filtro Art o REP-----------------------------------------------------------------------------------------------------------
    If chArt.Value <> chRep.Value Then
        Dim sSaco As String
        If chArt.Value = vbUnchecked Then sSaco = "Mercadería" Else sSaco = "Repuestos"
        aRow = 2
        For I = 2 To vsConsulta.Rows - 1
            With vsConsulta
                If Trim(.Cell(flexcpText, aRow, eCol.Tipo)) = sSaco Then .RemoveItem aRow Else aRow = aRow + 1
            End With
        Next
    End If
    '------------------------------------------------------------------------------------------------------------------------------------
    
    'Cargo lo que fue vendido (xtelefono y envios) y está para Retirar ------------------------------------------------------------------
    cons = "Select RVTArticulo as ArtID, Sum(RVTARetirar) as Q" & _
                " From VentaTelefonica, RenglonVtaTelefonica " & _
                " Where VTeTipo = 7 " & _
                " And VTeDocumento is Null And VTeAnulado is null " & _
                " And RVTARetirar > 0 And VTeCodigo = RVTVentaTelefonica " & _
                " Group by RVTArticulo"
                
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        aRow = RowEnLista(rsAux!ArtID)
        aID = rsAux!ArtID
        If aRow > 0 Then
            With vsConsulta
                .Cell(flexcpText, aRow, eCol.conARetEnv) = .Cell(flexcpValue, aRow, eCol.conARetEnv) + Format(rsAux!Q, "#,##0")
            End With
        End If
        rsAux.MoveNext
    Loop
    rsAux.Close
                 
    cons = " Select REvArticulo as ArtID, Sum (REvAEntregar) as Q " & _
            " From Envio, RenglonEnvio " & _
            " Where EnvTipo = 3 " & _
            " And EnvEstado NOT IN (2,4,5) " & _
            " And EnvDocumento Is Not Null " & _
            " And REvAEntregar > 0 And EnvCodigo = REvEnvio " & _
            " Group by REvArticulo"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        aRow = RowEnLista(rsAux!ArtID)
        aID = rsAux!ArtID
        If aRow > 0 Then
            With vsConsulta
                .Cell(flexcpText, aRow, eCol.conARetEnv) = .Cell(flexcpValue, aRow, eCol.conARetEnv) + Format(rsAux!Q, "#,##0")
            End With
        End If
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    'Cargo los comentarios para cada Articulo   ------------------------------------------------------------------
    cons = " Select * from ComentarioInventario"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        aRow = RowEnLista(rsAux!CInArticulo)
        If aRow > 0 Then vsConsulta.Cell(flexcpText, aRow, eCol.Notas) = Trim(rsAux!CInTexto)
        rsAux.MoveNext
    Loop
    rsAux.Close
   
    'Valido Filtro de Q Estados---------------------------------------------------------------------------------------------------------
    If vsConsulta.Rows > 2 Then
        If cEstado.ListIndex > 0 And cEstado.ListIndex <> 5 Then
            Dim bRemove As Boolean, fCompra As Date, QMeses As Integer
            '.AddItem "(listar todos)", 0
            '.AddItem "Última compra con costo cero", 1
            '.AddItem "Última compra menor a 6 meses", 2
            '.AddItem "Última compra entre 6 y 18 meses", 3
            '.AddItem "Última compra mayor a 18 meses", 4
            '.AddItem "Solamente artículos sin diferencias", 5
            
            aRow = 2
            With vsConsulta
                For I = 2 To .Rows - 1
                    bRemove = True
                    
                    Select Case cEstado.ListIndex
                        Case 1  '"Última compra con costo cero", 1
                            If IsNumeric(.Cell(flexcpText, aRow, eCol.UCCosto)) And .Cell(flexcpValue, aRow, eCol.UCCosto) = 0 Then bRemove = False
                                                      
                            
                        Case 2  '"Última compra menor a 6 meses", 2
                                    If IsDate(.Cell(flexcpText, aRow, eCol.UCFecha)) Then
                                        fCompra = CDate(.Cell(flexcpText, aRow, eCol.UCFecha))
                                        QMeses = (DateDiff("m", fCompra, fBalance))
                                        If QMeses < 6 Then bRemove = False
                                    End If
                                    
                                    If Not bRemove Then 'Si la ultima compra fue con costo cero lo borro
                                        If IsNumeric(.Cell(flexcpText, aRow, eCol.UCCosto)) And .Cell(flexcpValue, aRow, eCol.UCCosto) = 0 Then bRemove = True
                                    End If
                                    
                        Case 3  '"Última compra entre 6 y 18 meses", 3
                                    If IsDate(.Cell(flexcpText, aRow, eCol.UCFecha)) Then
                                        fCompra = CDate(.Cell(flexcpText, aRow, eCol.UCFecha))
                                        QMeses = (DateDiff("m", fCompra, fBalance))
                                        If QMeses >= 6 And QMeses <= 18 Then bRemove = False
                                    End If
                                    
                                    If Not bRemove Then 'Si la ultima compra fue con costo cero lo borro
                                        If IsNumeric(.Cell(flexcpText, aRow, eCol.UCCosto)) And .Cell(flexcpValue, aRow, eCol.UCCosto) = 0 Then bRemove = True
                                    End If
                                    
                        Case 4  '"Última compra mayor a 18 meses o LIFO = 0, 4
                                    If .Cell(flexcpValue, aRow, eCol.Lifo) = 0 Then
                                        bRemove = False
                                    Else
                                        If IsDate(.Cell(flexcpText, aRow, eCol.UCFecha)) Then
                                            fCompra = CDate(.Cell(flexcpText, aRow, eCol.UCFecha))
                                            QMeses = (DateDiff("m", fCompra, fBalance))
                                            If QMeses > 18 Then bRemove = False
                                        End If
                                        If Not bRemove Then 'Si la ultima compra fue con costo cero lo borro
                                            If IsNumeric(.Cell(flexcpText, aRow, eCol.UCCosto)) And .Cell(flexcpValue, aRow, eCol.UCCosto) = 0 Then bRemove = True
                                        End If
                                    End If
                        Case 5  '"Artículos sin diferencias"
                        
                    End Select
                    If bRemove Then .RemoveItem aRow Else aRow = aRow + 1
                Next
            End With
        End If
    End If
    '--------------------------------------------------------------------------------------------------------------------------------------
    
    'Proceso Las Diferencias ---------------------------------------------------------------------------------------------------------
    aQ = vsConsulta.Rows - 2
    If aQ > 0 Then
        pbProgreso.Min = 0: pbProgreso.Value = 0
        pbProgreso.Max = aQ
        
        Dim fAuxiliar As Date
        
        For I = 2 To vsConsulta.Rows - 1
            pbProgreso.Value = pbProgreso.Value + 1
            With vsConsulta
            
            .Cell(flexcpText, I, eCol.conTotal) = Format(.Cell(flexcpValue, I, eCol.conSano) + .Cell(flexcpValue, I, eCol.conARec) + .Cell(flexcpValue, I, eCol.conRoto) - .Cell(flexcpValue, I, eCol.conARetEnv), "#,##0")   'T-T
                
            .Cell(flexcpText, I, eCol.difConLifo) = Format(.Cell(flexcpValue, I, eCol.conTotal) - .Cell(flexcpValue, I, eCol.Lifo), "#,##0")
            
            'Costo Merca .LIFO
            If IsNumeric(.Cell(flexcpText, I, eCol.UCCosto)) Then
                .Cell(flexcpText, I, eCol.LifoCosto) = Format(.Cell(flexcpValue, I, eCol.difConLifo) * .Cell(flexcpValue, I, eCol.UCCosto), "#,##0.00")
                If .Cell(flexcpValue, I, eCol.UCCosto) = 0 Then .Cell(flexcpFontItalic, I, eCol.Articulo) = True
            End If
            
            If IsDate(.Cell(flexcpText, I, eCol.UCFecha)) Then
                fAuxiliar = CDate(.Cell(flexcpText, I, eCol.UCFecha))
                Select Case (DateDiff("m", fAuxiliar, fBalance))
                Case Is > 18: .Cell(flexcpBackColor, I, eCol.Articulo) = &HC0C0C0
                Case Is > 6: .Cell(flexcpBackColor, I, eCol.Articulo) = &HE0E0E0
                End Select
            End If
            
            If .Cell(flexcpValue, I, eCol.conTotal) <> 0 Then
                If (.Cell(flexcpValue, I, eCol.conSano) * 100 / .Cell(flexcpValue, I, eCol.conTotal)) <= 10 Then
                    .Cell(flexcpBackColor, I, eCol.conSano, , eCol.conRoto) = &HE0E0E0
                End If
            End If
            End With
        Next
    End If
    pbProgreso.Value = 0
    '------------------------------------------------------------------------------------------------------------------------------------
    
    'Valido Filtro de Q Estados---------------------------------------------------------------------------------------------------------
    If vsConsulta.Rows > 2 Then
        If cEstado.ListIndex = 5 Then   '"Solamente artículos sin diferencias", 5
            aRow = 2
            With vsConsulta
                For I = 2 To .Rows - 1
                    bRemove = True
                    
                    If .Cell(flexcpValue, aRow, eCol.difConLifo) = 0 And .Cell(flexcpValue, aRow, eCol.difConLifo) = 0 Then bRemove = False
                    
                    If Not bRemove Then 'Si la ultima compra fue con costo cero lo borro
                        If IsNumeric(.Cell(flexcpText, aRow, eCol.UCCosto)) And .Cell(flexcpValue, aRow, eCol.UCCosto) = 0 Then bRemove = True
                    End If

                    If bRemove Then .RemoveItem aRow Else aRow = aRow + 1
                Next
            End With
        End If
    End If
    '--------------------------------------------------------------------------------------------------------------------------------------
    
    
    If vsConsulta.Rows > 2 And cQCero.Value = vbChecked Then
        aRow = 2
        For I = 2 To vsConsulta.Rows - 1
            With vsConsulta
                If .Cell(flexcpValue, aRow, eCol.difConLifo) = 0 And .Cell(flexcpValue, aRow, eCol.difConLifo) = 0 Then
                    .RemoveItem aRow
                Else
                    aRow = aRow + 1
                End If
            End With
        Next
    End If
    '--------------------------------------------------------------------------------------------------------------------------------------
    With vsConsulta
        If .Rows > 2 Then
        .Cell(flexcpBackColor, 0, eCol.conTotal, .Rows - 1, eCol.conTotal) = Colores.Gris
        .Cell(flexcpBackColor, 0, eCol.difConLifo, .Rows - 1, eCol.difConLifo) = Colores.Gris
        
        .Select 2, 1
        .Sort = flexSortGenericAscending
        
        .Subtotal flexSTSum, 0, eCol.conTotal, "#,##0", vbInactiveBorder
        .Subtotal flexSTSum, 0, eCol.conSano, "#,##0": .Subtotal flexSTSum, 0, eCol.conARec, "#,##0": .Subtotal flexSTSum, 0, eCol.conRoto, "#,##0": .Subtotal flexSTSum, 0, eCol.conARetEnv, "#,##0"
        .Subtotal flexSTSum, 0, eCol.Lifo, "#,##0": .Subtotal flexSTSum, 0, eCol.difConLifo, "#,##0"
        .Subtotal flexSTSum, 0, eCol.LifoCosto, "#,##0.00"
    
        .Subtotal flexSTSum, -1, eCol.conTotal, "#,##0", vbInactiveBorder, , , "TOTAL"
        .Subtotal flexSTSum, -1, eCol.conSano, "#,##0": .Subtotal flexSTSum, -1, eCol.conARec, "#,##0": .Subtotal flexSTSum, -1, eCol.conRoto, "#,##0": .Subtotal flexSTSum, -1, eCol.conARetEnv, "#,##0"
        .Subtotal flexSTSum, -1, eCol.Lifo, "#,##0": .Subtotal flexSTSum, -1, eCol.difConLifo, "#,##0"
        .Subtotal flexSTSum, -1, eCol.LifoCosto, "#,##0.00"

        .Cell(flexcpAlignment, .FixedRows, eCol.Articulo + 1, .Rows - 1, eCol.Notas - 1) = flexAlignRightCenter
        End If
    End With
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Function RowEnLista(idArticulo As Long) As Long
    RowEnLista = 0
    Dim I As Long
    
    With vsConsulta
        For I = 1 To .Rows - 1
            If idArticulo = .Cell(flexcpData, I, eCol.Articulo) Then
                RowEnLista = I: Exit For
            End If
        Next
    End With
    
End Function
Private Sub BuscoUltimoCosto(idArticulo As Long, Cantidad As Long)

Dim rsCom As rdoResultset
Dim Costo1 As Currency, aQ As Long

    Costo1 = 0: aQ = Abs(Cantidad)
    'Primero Compras con costo 0 ------------------------------------------------------------
    cons = "Select * from CMCompra " _
            & " Where ComFecha <= '06/30/2000'" _
            & " And ComArticulo =  " & idArticulo _
            & " And ComCantidad > 0 " _
            & " And ComCosto = 0" _
            & " Order by ComFecha DESC"
    Set rsCom = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsCom.EOF
        aQ = aQ - rsCom!ComCantidad
        If aQ <= 0 Then Exit Do
        rsCom.MoveNext
    Loop
    rsCom.Close
    '-----------------------------------------------------------------------------------------------
    '2) Con Costo--------------------------------------------------------------------------------
    If aQ > 0 Then
        cons = "Select * from CMCompra " _
                & " Where ComFecha <= '06/30/2000'" _
                & " And ComArticulo =  " & idArticulo _
                & " And ComCantidad > 0 " _
                & " And ComCosto > 0" _
                & " Order by ComFecha DESC"
        Set rsCom = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        Do While Not rsCom.EOF
            If aQ >= rsCom!ComCantidad Then
                Costo1 = Costo1 + (rsCom!ComCantidad * rsCom!ComCosto)
            Else
                Costo1 = Costo1 + (aQ * rsCom!ComCosto)
            End If
            
            aQ = aQ - rsCom!ComCantidad
            If aQ <= 0 Then Exit Do
            rsCom.MoveNext
        Loop
        rsCom.Close
    End If
    '-----------------------------------------------------------------------------------------------
    
    If Cantidad < 0 Then Costo1 = Costo1 * -1
    With vsConsulta
        .Cell(flexcpText, .Rows - 1, 6) = Format(Costo1, "#,##0.00")
    End With
    
End Sub

Private Sub MnuPlantilla_Click()
    Call vsConsulta_DblClick
End Sub

Private Sub mnuVerCol_Click(Index As Integer)
    vsConsulta.ColHidden(Val(mnuVerCol(Index).Tag)) = Not vsConsulta.ColHidden(Val(mnuVerCol(Index).Tag))
End Sub

Private Sub vsConsulta_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Debug.Print savNotas & " EDT=" & vsConsulta.EditText & "  TXT=" & vsConsulta.Cell(flexcpText, Row, Col)
    
    If Col = eCol.difConLifo Then
        Exit Sub
    End If
    
    If Not savNotas Then Exit Sub
    
    On Error GoTo errSAV
    Screen.MousePointer = 11
    
    Dim idArticulo As Long, mTXT As String
    idArticulo = vsConsulta.Cell(flexcpData, Row, eCol.Articulo)
    mTXT = vsConsulta.EditText
    
    cons = " Select * from ComentarioInventario Where CInArticulo = " & idArticulo
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If mTXT <> "" Then
        If rsAux.EOF Then rsAux.AddNew Else rsAux.Edit
        rsAux!CInArticulo = idArticulo
        rsAux!CInTexto = mTXT
        rsAux.Update
    Else
        If Not rsAux.EOF Then rsAux.Delete
    End If
    rsAux.Close
    
    Screen.MousePointer = 0
    Exit Sub
    
errSAV:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Error al actualizar el comentario del artículo", Err.Description
End Sub

Private Sub vsConsulta_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = vsConsulta.Cell(flexcpData, Row, eCol.Articulo) = 0 Or (Col <> eCol.Notas And Col <> eCol.difConLifo)
    savNotas = False
End Sub

Private Sub vsConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDelete Then
        Dim idArticulo As Long, mTXT As String
        idArticulo = vsConsulta.Cell(flexcpData, vsConsulta.Row, eCol.Articulo)
        If idArticulo = 0 Or vsConsulta.Row < vsConsulta.FixedRows Then Exit Sub
        
        If vsConsulta.Cell(flexcpFontBold, vsConsulta.Row, eCol.difConLifo) = True Then
            MsgBox "Este artículo ya fue corregido ...", vbInformation, "Artículo ya corregido"
            Exit Sub
        End If
        
        Dim bOK As Boolean
        bOK = gb_CorrigoLifo(idArticulo, vsConsulta.Row)
        
        If bOK Then
            vsConsulta.Cell(flexcpFontBold, vsConsulta.Row, eCol.difConLifo) = True
            vsConsulta.Cell(flexcpForeColor, vsConsulta.Row, eCol.difConLifo) = vbWhite
            vsConsulta.Cell(flexcpBackColor, vsConsulta.Row, eCol.difConLifo) = vbBlue
        End If
        
    End If
End Sub

Private Sub vsConsulta_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
        Case eCol.Notas
            savNotas = vsConsulta.EditText <> vsConsulta.Cell(flexcpText, Row, Col)
        Case eCol.difConLifo
            If Not IsNumeric(vsConsulta.EditText) Then Cancel = True
    End Select
    
    
End Sub

Private Sub vsConsulta_DblClick()
    On Error GoTo errMnu
    If vsConsulta.Rows <= 2 Then Exit Sub
    Dim aIdReg As String
    
    aIdReg = vsConsulta.Cell(flexcpData, vsConsulta.Row, 2)
    If aIdReg = "" Then Exit Sub
    
    EjecutarApp App.Path & "\appExploreMsg.exe ", paPlBalance & ":" & aIdReg
    Exit Sub

errMnu:
End Sub

Private Sub vsConsulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsConsulta.Rows > 2 And Button = vbRightButton Then
        If Val(vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)) <> 0 Then
            MnuTitulo.Caption = Trim(vsConsulta.Cell(flexcpText, vsConsulta.Row, 1))
            PopupMenu MnuAcciones
        End If
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
        With vsListado
            .StartDoc
            .Columns = 1
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        aTexto = "Diferencias Balance"
        EncabezadoListado vsListado, aTexto, False
        vsListado.FileName = "Diferencias Balance"
        
        vsListado.Paragraph = "Filtros: " & cEstado.Text
        With vsConsulta
            .Redraw = False
            vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
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
    clsGeneral.OcurrioError "Error al realizar la impresión", Err.Description
End Sub

Private Sub AccionConfigurar()
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
End Sub

Public Function EsRepuesto(idArticulo As Long) As Boolean

    If idArticulo > 250000 Then EsRepuesto = True Else EsRepuesto = False
    Exit Function
    
    cons = "Select * from ArticuloGrupo Where AGrArticulo = " & idArticulo & " And AGrGrupo = " & paGrupoRepuesto
    Set rsRep = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsRep.EOF Then EsRepuesto = True Else EsRepuesto = False
    rsRep.Close
    
End Function

Private Function zfn_CargofBalance()

    fBalance = Date
    
    cons = "Select Max(CabMesCosteo) as fBalance from CMCabezal"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux!fBalance) Then fBalance = rsAux!fBalance
    rsAux.Close
    
    fBalance = DateAdd("m", 1, fBalance)
    fBalance = DateAdd("d", Day(fBalance) * -1, fBalance)
    
End Function

Private Function gb_CorrigoLifo(idArticulo As Long, idRow As Long) As Boolean
Dim mQDiff As Long, mTXT_MSG As String, xArticulo As String
Dim mERR As String


    gb_CorrigoLifo = False
    
    xArticulo = vsConsulta.Cell(flexcpText, idRow, eCol.Articulo)
    mQDiff = vsConsulta.Cell(flexcpValue, idRow, eCol.difConLifo)
    'mQDiff = Contado - Lifo --> SI < 0 hay que sacar del lifo , SI > 0 hay que agregar
    
    If mQDiff = 0 Then
        MsgBox "No hay diferencias para actualizar en el lifo de " & xArticulo, vbInformation, "No hay diferencias"
        Exit Function
    End If
    
    On Error GoTo errCorregir
    Dim mQAuxiliar As Long, iBase As Long
    
    Dim rsUPD As rdoResultset, sqlUPD As String

    If mQDiff < 0 Then
        mQAuxiliar = Abs(mQDiff)
        mTXT_MSG = "Dar de BAJA del LIFO " & Abs(mQDiff)
        If MsgBox(mTXT_MSG, vbQuestion + vbYesNo + vbDefaultButton2, xArticulo) = vbNo Then Exit Function
        
        For iBase = 1 To 2
            mERR = "1"
            mQAuxiliar = Abs(mQDiff)
            
            cons = "Select * from CMCompra " & _
                    " Where ComArticulo = " & idArticulo & " And ComCantidad > 0" & _
                    " Order by ComFecha DESC"
            
            If iBase = 1 Then Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If iBase = 2 Then Set rsAux = cBaseMov.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            
            Do While Not rsAux.EOF
                mERR = "2"
                
                sqlUPD = "Select * from CMCompra " & _
                              " Where ComArticulo = " & idArticulo & _
                              " And ComCodigo = " & rsAux("ComCodigo").Value & _
                              " And ComFecha = '" & Format(rsAux("ComFecha").Value, "mm/dd/yyyy") & "'" & _
                              " And ComTipo = " & rsAux("ComTipo").Value
                              
                If iBase = 1 Then Set rsUPD = cBase.OpenResultset(sqlUPD, rdOpenDynamic, rdConcurValues)
                If iBase = 2 Then Set rsUPD = cBaseMov.OpenResultset(sqlUPD, rdOpenDynamic, rdConcurValues)
                
                If rsAux!ComCantidad <= mQAuxiliar Then
                    mERR = "3"
                    'Borro Registro
                    mQAuxiliar = (mQAuxiliar - rsAux!ComCantidad)
                    rsUPD.Delete                    'rsAux.Delete
                    
                Else    'Edito y Modifico
                    mERR = "4"
                                        
                    rsUPD.Edit
                    mERR = "5"
                    rsUPD!ComCantidad = (rsUPD!ComCantidad - mQAuxiliar)
                    rsUPD.Update
                    mERR = "6"
                    mQAuxiliar = 0
                End If
                
                rsUPD.Close
                
                rsAux.MoveNext
                mERR = "7"
                If mQAuxiliar = 0 Then Exit Do
            Loop
            rsAux.Close
            mERR = "8"
        Next
        
    Else
        Dim mMCosto As Currency: mMCosto = 0
        If IsNumeric(vsConsulta.Cell(flexcpData, idRow, eCol.UCCosto)) Then
            mMCosto = vsConsulta.Cell(flexcpData, idRow, eCol.UCCosto)
        End If
        
        mQAuxiliar = mQDiff
        mTXT_MSG = "Agregar al LIFO " & Abs(mQDiff) & " con fecha de balance (" & Format(fBalance, "dd/mm/yyyy") & ") y costo unitario de " & Format(mMCosto, "#,##0.00")
        If MsgBox(mTXT_MSG, vbQuestion + vbYesNo + vbDefaultButton2, xArticulo) = vbNo Then Exit Function
        
        For iBase = 1 To 2
            
            mQAuxiliar = Abs(mQDiff)

            Dim idMinimaCompra As Long: idMinimaCompra = 0      '1) busco id de compra minima   ---------
            cons = "Select Min(ComCodigo) From CMCompra"
            
            If iBase = 1 Then Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If iBase = 2 Then Set rsAux = cBaseMov.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            
            If Not rsAux.EOF Then idMinimaCompra = rsAux(0)
            rsAux.Close
            idMinimaCompra = idMinimaCompra - 1                         '---------------------------------------------

            cons = "Select * from CMCompra Where ComArticulo = 0"
            If iBase = 1 Then Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If iBase = 2 Then Set rsAux = cBaseMov.OpenResultset(cons, rdOpenDynamic, rdConcurValues)

            rsAux.AddNew
            rsAux!ComFecha = fBalance
            rsAux!ComArticulo = idArticulo
            rsAux!ComCantidad = mQAuxiliar
            rsAux!ComCodigo = idMinimaCompra
            rsAux!ComTipo = 2
            rsAux!ComCosto = mMCosto
            rsAux!ComQOriginal = mQAuxiliar
            rsAux.Update
            rsAux.Close
        Next
        
    End If
    
    gb_CorrigoLifo = True
    Exit Function
    
errCorregir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError mERR & ") Error al realizar la corrección del LIFO " & iBase, Err.Description
End Function
