VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmListado 
   Caption         =   "Diferencias Balance"
   ClientHeight    =   7530
   ClientLeft      =   1380
   ClientTop       =   2070
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
   Begin VB.CheckBox cVerificado 
      Caption         =   "Eliminar Artículos Verificados."
      Height          =   195
      Left            =   6900
      TabIndex        =   22
      Top             =   300
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CheckBox cQCero 
      Caption         =   "Eliminar Totales en 0"
      Height          =   195
      Left            =   6900
      TabIndex        =   21
      Top             =   60
      Width           =   1815
   End
   Begin VB.ComboBox cEstado 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   0
      Width           =   4215
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
         NumPanels       =   4
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
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12541
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Filtros:"
      Height          =   255
      Left            =   1800
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
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsAux As rdoResultset, rs1 As rdoResultset
Dim rsRep As rdoResultset

Private aTexto As String
Dim bCargarImpresion As Boolean

Private Sub AccionLimpiar()
    'vsConsulta.Rows = 2
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
    On Error GoTo ErrLoad
    
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
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
        .FormatString = "<Tipo|<Artículo|>Sistema (Locales)|>|>|>|>|>Se Contó (Balance)|>|>|>|>|TOTAL|>LIFO|>Ctdo.-Sist|>Ctdo.-Lifo|"
            
        .WordWrap = False: .MergeCells = flexMergeSpill
        .ColHidden(0) = True
        .ColWidth(1) = 2800
        .ColWidth(2) = 600: .ColWidth(3) = 350: .ColWidth(4) = 700: .ColWidth(5) = 400: .ColWidth(6) = 700
        .ColWidth(7) = 600: .ColWidth(8) = 350: .ColWidth(9) = 700: .ColWidth(10) = 400: .ColWidth(11) = 700
        .ColWidth(12) = 700: .ColWidth(13) = 700
        .ColWidth(14) = 900: .ColWidth(15) = 1150
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 2) = "SAN": .Cell(flexcpText, .Rows - 1, 3) = "AR": .Cell(flexcpText, .Rows - 1, 4) = "ST": .Cell(flexcpText, .Rows - 1, 5) = "ROT": .Cell(flexcpText, .Rows - 1, 6) = "Total"
        .Cell(flexcpText, .Rows - 1, 7) = "SAN": .Cell(flexcpText, .Rows - 1, 8) = "AR": .Cell(flexcpText, .Rows - 1, 9) = "ST": .Cell(flexcpText, .Rows - 1, 10) = "ROT": .Cell(flexcpText, .Rows - 1, 11) = "Total"
        .Cell(flexcpText, .Rows - 1, 12) = "Prop."
        .Cell(flexcpText, .Rows - 1, 14) = "T.Ctdo - SI"
        .Cell(flexcpText, .Rows - 1, 15) = "(C-AR-AE)-Lif"
        
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, 1) = .BackColorFixed
        .Cell(flexcpBackColor, 0, 2, .Rows - 1, 6) = Colores.clNaranja
        .Cell(flexcpBackColor, 0, 7, .Rows - 1, 11) = Colores.Gris
        .Cell(flexcpBackColor, 0, 12, .Rows - 1, 12) = Colores.Blanco
        .Cell(flexcpBackColor, 0, 13, .Rows - 1, 13) = Colores.clCeleste
        .Cell(flexcpBackColor, 0, 14, .Rows - 1, 14) = Colores.Obligatorio
        .Cell(flexcpBackColor, 0, 15, .Rows - 1, 16) = Colores.Blanco
        
                
        .GridLines = flexGridFlat
        .FixedRows = 2
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
        
        cons = "Select  ArtID, ArtCodigo, ArtNombre, Sum(ComCantidad) as Q" & _
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
                .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
                If EsRepuesto(rsAux!ArtCodigo) Then .Cell(flexcpText, .Rows - 1, 0) = "Repuestos " Else .Cell(flexcpText, .Rows - 1, 0) = "Mercadería"
                aValor = rsAux!ArtID: .Cell(flexcpData, .Rows - 1, 1) = aValor
                aValor = rsAux!ArtCodigo: .Cell(flexcpData, .Rows - 1, 2) = aValor
                
                .Cell(flexcpText, .Rows - 1, 13) = Format(rsAux!Q, "#,##0")
            End With
            rsAux.MoveNext
        Loop
        rsAux.Close
    End If
    
    'Cargo Artículos del STOCK ACTUAL   ---------------------------------------------------------------------------------------------------------
    cons = "Select Count(Distinct(StLArticulo)) from StockLocal" & _
               " Where StLArticulo Not In (Select ArtId from Articulo Where ArtTipo = " & paTipoArticuloServicio & ")"
    If cVerificado.Value = vbChecked Then cons = cons & " And StLArticulo" & sqlVerificados
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux(0)) Then aQ = rsAux(0)
    rsAux.Close
    
    If aQ > 0 Then
        pbProgreso.Min = 0: pbProgreso.Value = 0
        pbProgreso.Max = aQ
        
        cons = "Select ArtID, ArtCodigo, ArtNombre, StLEstado, Sum(StLCantidad) as Q " & _
                   " From StockLocal, Articulo" & _
                   " Where StLArticulo = ArtID " & _
                   " And ArtTipo <> " & paTipoArticuloServicio
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
                    
                    .Cell(flexcpText, aRow, 1) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
                    If EsRepuesto(rsAux!ArtCodigo) Then .Cell(flexcpText, aRow, 0) = "Repuestos " Else .Cell(flexcpText, aRow, 0) = "Mercadería"
                    aValor = rsAux!ArtID: .Cell(flexcpData, aRow, 1) = aValor
                    aValor = rsAux!ArtCodigo: .Cell(flexcpData, .Rows - 1, 2) = aValor
                End If
                
                Select Case rsAux!StLEstado
                    Case paEstadoArticuloEntrega: .Cell(flexcpText, aRow, 2) = Format(rsAux!Q, "#,##0")
                    Case paEstadoArticuloARecuperar: .Cell(flexcpText, aRow, 3) = Format(rsAux!Q, "#,##0")
                    
                    Case paEstadoArticuloRoto: .Cell(flexcpText, aRow, 5) = Format(rsAux!Q, "#,##0")
                End Select
                .Cell(flexcpText, aRow, 4) = Format(.Cell(flexcpValue, aRow, 2) + .Cell(flexcpValue, aRow, 3), "#,##0")
                .Cell(flexcpText, aRow, 6) = Format(.Cell(flexcpValue, aRow, 4) + .Cell(flexcpValue, aRow, 5), "#,##0")
                
            End With
            rsAux.MoveNext
        Loop
        rsAux.Close
    End If
    
    pbProgreso.Value = 0
    
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
                    
                    .Cell(flexcpText, aRow, 1) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
                    If EsRepuesto(rsAux!ArtCodigo) Then .Cell(flexcpText, aRow, 0) = "Repuestos " Else .Cell(flexcpText, aRow, 0) = "Mercadería"
                    aValor = rsAux!ArtID: .Cell(flexcpData, aRow, 1) = aValor
                    aValor = rsAux!ArtCodigo: .Cell(flexcpData, .Rows - 1, 2) = aValor
                End If
                
                Select Case rsAux!BMREstado
                    Case paEstadoArticuloEntrega: .Cell(flexcpText, aRow, 7) = Format(rsAux!Q, "#,##0")
                    Case paEstadoArticuloARecuperar: .Cell(flexcpText, aRow, 8) = Format(rsAux!Q, "#,##0")
                    Case paEstadoArticuloRoto: .Cell(flexcpText, aRow, 10) = Format(rsAux!Q, "#,##0")
                End Select
                .Cell(flexcpText, aRow, 9) = Format(.Cell(flexcpValue, aRow, 7) + .Cell(flexcpValue, aRow, 8), "#,##0")
                .Cell(flexcpText, aRow, 11) = Format(.Cell(flexcpValue, aRow, 9) + .Cell(flexcpValue, aRow, 10), "#,##0")
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
                    
                    .Cell(flexcpText, aRow, 1) = Format(rsAux!ArtCodigo, "(#,000,000)") & " " & Trim(rsAux!ArtNombre)
                    If EsRepuesto(rsAux!ArtCodigo) Then .Cell(flexcpText, aRow, 0) = "Repuestos " Else .Cell(flexcpText, aRow, 0) = "Mercadería"
                    aValor = rsAux!ArtID: .Cell(flexcpData, aRow, 1) = aValor
                    aValor = rsAux!ArtCodigo: .Cell(flexcpData, .Rows - 1, 2) = aValor
                End If
                
                Select Case rsAux!StLEstado
                    Case paEstadoArticuloEntrega: .Cell(flexcpText, aRow, 7) = .Cell(flexcpValue, aRow, 7) + Format(rsAux!Q, "#,##0")
                    Case paEstadoArticuloARecuperar: .Cell(flexcpText, aRow, 8) = .Cell(flexcpValue, aRow, 8) + Format(rsAux!Q, "#,##0")
                    Case paEstadoArticuloRoto: .Cell(flexcpText, aRow, 10) = .Cell(flexcpValue, aRow, 10) + Format(rsAux!Q, "#,##0")
                End Select
                .Cell(flexcpText, aRow, 9) = Format(.Cell(flexcpValue, aRow, 7) + .Cell(flexcpValue, aRow, 8), "#,##0")
                .Cell(flexcpText, aRow, 11) = Format(.Cell(flexcpValue, aRow, 9) + .Cell(flexcpValue, aRow, 10), "#,##0")
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
                If Trim(.Cell(flexcpText, aRow, 0)) = sSaco Then
                    .RemoveItem aRow
                Else
                    aRow = aRow + 1
                End If
            End With
        Next
    End If
    '------------------------------------------------------------------------------------------------------------------------------------
    
    'Cargo el TOTAL Propiedad (Total Conteo - ARetirar - AEnviar) ------------------------------------------------------------------
    aQ = vsConsulta.Rows - 2
    If aQ > 0 Then
        pbProgreso.Min = 0: pbProgreso.Value = 0
        pbProgreso.Max = aQ
        
        Dim aKey As Integer
        
        For I = 2 To vsConsulta.Rows - 1
            pbProgreso.Value = pbProgreso.Value + 1
            
            qAuxiliar = vsConsulta.Cell(flexcpValue, I, 11)
            aKey = 0
            cons = " Select StTEstado, StTCantidad From StockTotal " & _
                       " Where StTArticulo = " & vsConsulta.Cell(flexcpData, I, 1) & _
                        " And StTTipoEstado = 2 And StTCantidad <> 0 " & _
                        " Group by  StTEstado, StTCantidad"
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            Do While Not rsAux.EOF
                qAuxiliar = qAuxiliar - rsAux!StTCantidad
                aKey = aKey + rsAux!StTEstado
                rsAux.MoveNext
            Loop
            rsAux.Close
            
            'Le sumo las vtas telefonicas que no fueron facturadas (porque va contra el lifo)-----------------
            cons = "Select Sum(RVTARetirar) as Q" & _
                    " From VentaTelefonica, RenglonVtaTelefonica " & _
                    " Where VTeTipo = 7 " & _
                    " And VTeDocumento is Null And VTeAnulado is null " & _
                    " And RVTArticulo = " & vsConsulta.Cell(flexcpData, I, 1) & " And RVTARetirar > 0 " & _
                    " And VTeCodigo = RVTVentaTelefonica " & _
                            " Union All " & _
                    " Select Sum (REvAEntregar) as Q " & _
                    " From Envio, RenglonEnvio " & _
                    " Where EnvTipo = 3 " & _
                    " And EnvEstado NOT IN (2,4,5) " & _
                    " And EnvDocumento is not Null " & _
                    " And REvArticulo = " & vsConsulta.Cell(flexcpData, I, 1) & " And REvAEntregar > 0 " & _
                    " And EnvCodigo = REvEnvio "
            
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            Do While Not rsAux.EOF
                If Not IsNull(rsAux!Q) Then qAuxiliar = qAuxiliar + rsAux!Q
                rsAux.MoveNext
            Loop
            rsAux.Close
            '---------------------------------------------------------------------------------------------------------
            
            vsConsulta.Cell(flexcpText, I, 12) = Format(qAuxiliar, "#,##0")
            vsConsulta.Cell(flexcpData, I, 12) = aKey
            
        Next
    End If
    pbProgreso.Value = 0
    
    'Valido Filtro de Q Estados---------------------------------------------------------------------------------------------------------
    If vsConsulta.Rows > 2 Then
        If Trim(cEstado.Text) <> "" And cEstado.ListIndex <> -1 Then
            aRow = 2
            For I = 2 To vsConsulta.Rows - 1
                
                With vsConsulta
                    If .Cell(flexcpData, aRow, 12) <> cEstado.ItemData(cEstado.ListIndex) Then
                        .RemoveItem aRow
                    Else
                        aRow = aRow + 1
                    End If
                End With
            Next
        End If
    End If
    '--------------------------------------------------------------------------------------------------------------------------------------
    
    'Proceso Las Diferencias ---------------------------------------------------------------------------------------------------------
    aQ = vsConsulta.Rows - 2
    If aQ > 0 Then
        pbProgreso.Min = 0: pbProgreso.Value = 0
        pbProgreso.Max = aQ
        
        For I = 2 To vsConsulta.Rows - 1
            pbProgreso.Value = pbProgreso.Value + 1
            With vsConsulta
                '.Cell(flexcpText, I, 14) = Format(.Cell(flexcpValue, I, 9) - .Cell(flexcpValue, I, 4), "#,##0") 'ST-ST
                '.Cell(flexcpText, I, 15) = Format(.Cell(flexcpValue, I, 11) - .Cell(flexcpValue, I, 6), "#,##0")    'T-T
                .Cell(flexcpText, I, 14) = Format(.Cell(flexcpValue, I, 11) - .Cell(flexcpValue, I, 6), "#,##0")    'T-T
                
                '.Cell(flexcpText, I, 16) = Format(.Cell(flexcpValue, I, 9) - .Cell(flexcpValue, I, 12), "#,##0")
                '.Cell(flexcpText, I, 17) = Format(.Cell(flexcpValue, I, 11) - .Cell(flexcpValue, I, 12), "#,##0")
                .Cell(flexcpText, I, 15) = Format(.Cell(flexcpValue, I, 12) - .Cell(flexcpValue, I, 13), "#,##0")
            End With
        Next
    End If
    pbProgreso.Value = 0
    '------------------------------------------------------------------------------------------------------------------------------------
    'Valido Filtro de Q Estados---------------------------------------------------------------------------------------------------------
    If vsConsulta.Rows > 2 And cQCero.Value = vbChecked Then
        aRow = 2
        For I = 2 To vsConsulta.Rows - 1
            With vsConsulta
                If .Cell(flexcpValue, aRow, 14) = 0 And .Cell(flexcpValue, aRow, 15) = 0 Then
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
        .Cell(flexcpBackColor, 0, 2, .Rows - 1, 6) = Colores.clNaranja
        .Cell(flexcpBackColor, 0, 7, .Rows - 1, 11) = Colores.Gris
        .Cell(flexcpBackColor, 0, 12, .Rows - 1, 12) = Colores.Blanco
        .Cell(flexcpBackColor, 0, 13, .Rows - 1, 13) = Colores.clCeleste
        .Cell(flexcpBackColor, 0, 14, .Rows - 1, 14) = Colores.Obligatorio
        .Cell(flexcpBackColor, 0, 15, .Rows - 1, 16) = Colores.Blanco
        
        .Select 2, 1   ', 1, 2
        .Sort = flexSortGenericAscending
        
        .Subtotal flexSTSum, 0, 6, "#,##0", vbInactiveBorder
        .Subtotal flexSTSum, 0, 2, "#,##0": .Subtotal flexSTSum, 0, 3, "#,##0": .Subtotal flexSTSum, 0, 4, "#,##0": .Subtotal flexSTSum, 0, 5, "#,##0"
        .Subtotal flexSTSum, 0, 7, "#,##0": .Subtotal flexSTSum, 0, 8, "#,##0": .Subtotal flexSTSum, 0, 9, "#,##0": .Subtotal flexSTSum, 0, 10, "#,##0"
        .Subtotal flexSTSum, 0, 11, "#,##0": .Subtotal flexSTSum, 0, 13, "#,##0": .Subtotal flexSTSum, 0, 14, "#,##0": .Subtotal flexSTSum, 0, 15, "#,##0"
        .Subtotal flexSTSum, 0, 12, "#,##0"
    
        .Subtotal flexSTSum, -1, 6, "#,##0", vbInactiveBorder, , , "TOTAL"
        .Subtotal flexSTSum, -1, 2, "#,##0": .Subtotal flexSTSum, -1, 3, "#,##0": .Subtotal flexSTSum, -1, 4, "#,##0": .Subtotal flexSTSum, -1, 5, "#,##0"
        .Subtotal flexSTSum, -1, 7, "#,##0": .Subtotal flexSTSum, -1, 8, "#,##0": .Subtotal flexSTSum, -1, 9, "#,##0": .Subtotal flexSTSum, -1, 10, "#,##0"
        .Subtotal flexSTSum, -1, 11, "#,##0": .Subtotal flexSTSum, -1, 13, "#,##0": .Subtotal flexSTSum, -1, 14, "#,##0": .Subtotal flexSTSum, -1, 15, "#,##0"
        .Subtotal flexSTSum, -1, 12, "#,##0"
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
            If idArticulo = .Cell(flexcpData, I, 1) Then
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

Private Sub MnuComentarios_Click()
 On Error GoTo errMnu
    If vsConsulta.Rows <= 2 Then Exit Sub
    Dim aIdReg As String
    
    aIdReg = vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)
    If aIdReg = "" Then Exit Sub
    
    EjecutarApp App.Path & "\Comentario_Inventario.exe ", CStr(aIdReg)
    Exit Sub

errMnu:
End Sub

Private Sub MnuPlantilla_Click()
    Call vsConsulta_DblClick
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
