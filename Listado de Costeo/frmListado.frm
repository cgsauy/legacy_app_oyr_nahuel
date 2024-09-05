VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmListado 
   Caption         =   "Listado del Costeo y su Ganancia"
   ClientHeight    =   7530
   ClientLeft      =   2025
   ClientTop       =   1665
   ClientWidth     =   10830
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
   ScaleWidth      =   10830
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4035
      Left            =   4440
      TabIndex        =   4
      Top             =   1800
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   7117
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
      TabIndex        =   7
      Top             =   720
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
      EmptyColor      =   8421504
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   11595
      TabIndex        =   8
      Top             =   6720
      Width           =   11655
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4500
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4140
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
         Left            =   3180
         Picture         =   "frmListado.frx":0EF0
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2820
         Picture         =   "frmListado.frx":0FDA
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   2220
         Picture         =   "frmListado.frx":10C4
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3780
         Picture         =   "frmListado.frx":12FE
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4860
         Picture         =   "frmListado.frx":1400
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5460
         Picture         =   "frmListado.frx":17C6
         Style           =   1  'Graphical
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1860
         Picture         =   "frmListado.frx":1F0C
         Style           =   1  'Graphical
         TabIndex        =   10
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
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   265
         Left            =   6000
         TabIndex        =   21
         Top             =   140
         Width           =   5295
         _ExtentX        =   9340
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
         TabIndex        =   22
         Top             =   150
         Width           =   375
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   7275
      Width           =   10830
      _ExtentX        =   19103
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
            Object.Width           =   10901
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
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   10335
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   4395
      End
      Begin VB.TextBox tMes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Mes del Costeo:"
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   255
         Width           =   1215
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
    Compra = 1              'Compra Comun (a proveedores de mercaderia locales)
    Comercio = 2            'Cualquier documento del comercio (ctdo, cred, etc...)
    Importacion = 3        'Compra (que entra por importaciones)
    Servicio = 4              'Documento ralacionado a Servicios (Ventas por servicios no facturados)
End Enum

Private rsAux As rdoResultset, rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean

Private Sub AccionLimpiar()
    tMes.Text = "": tArticulo.Text = ""
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
        
        If Val(tArticulo.Text) > 0 Then
            cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtCodigo = " & CLng(tArticulo.Text)
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurReadOnly)
            If rsAux.EOF Then
                rsAux.Close
                MsgBox "No se encontró un artículo con ese código.", vbInformation, "ATENCIÓN"
            Else
                tArticulo.Text = Trim(rsAux!Nombre)
                tArticulo.Tag = rsAux!ArtID
                rsAux.Close
                Foco bConsultar
            End If
        Else
            tArticulo.Text = "0"
        End If
        Screen.MousePointer = 0
    Else
        If KeyCode = vbKeyReturn Then Foco bConsultar
    End If
    Exit Sub

ErrTA:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tMes_GotFocus()
    With tMes: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tMes_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tArticulo
End Sub

Private Sub Label5_Click()
    Foco tMes
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
        .PhysicalPage = True
        .PaperSize = 1
        .Orientation = orPortrait
        .Zoom = 100
        .MarginLeft = 900: .MarginRight = 350
        .MarginBottom = 750: .MarginTop = 750
    End With
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .OutlineBar = flexOutlineBarNone ' flexOutlineBarComplete
        .OutlineCol = 0
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<Artículo|Fecha|Tipo|<Número|>Q Venta|>$ Compra (x1)|>$ Venta (x1)|>Ganancia|"
            
        .WordWrap = False
        AnchoEncabezado Pantalla:=True
        .MergeCells = flexMergeSpill
        .MergeCol(0) = True
    End With
      
End Sub

Private Sub AnchoEncabezado(Optional Pantalla As Boolean = False, Optional Impresora As Boolean = False)

    With vsConsulta
        
        If Pantalla Then
            .ColWidth(0) = 0: .ColWidth(1) = 950: .ColWidth(2) = 1000: .ColWidth(3) = 1200: .ColWidth(4) = 1000
            .ColWidth(5) = 1600: .ColWidth(6) = 1600: .ColWidth(7) = 1400
        End If
        
        If Impresora Then
            .ColWidth(0) = 0: .ColWidth(1) = 570: .ColWidth(2) = 430: .ColWidth(3) = 500: .ColWidth(4) = 560
            .ColWidth(5) = 1000: .ColWidth(6) = 1000: .ColWidth(7) = 1000
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

Dim aIDCosteo As Long
Dim aArticulo As Long, aTxtArticulo As String

Dim aCostoP As Currency, aVentaP As Currency, aCantidadP As Long
Dim aCostoT As Currency, aVentaT As Currency

Dim rsSer As rdoResultset

    On Error GoTo errConsultar
    Screen.MousePointer = 11
    bCargarImpresion = True
    aArticulo = 0
    lQ.Tag = 0: lQ.Caption = ""
    'Primero Hay que sacar el ID del Costeo para el mes------------------------------------------------------------------------------
    aIDCosteo = 0
    cons = "Select * from CMCabezal Where CabMesCosteo = '" & Format(tMes.Text, sqlFormatoF) & "'"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then aIDCosteo = rsAux!CabID
    rsAux.Close
    
    If aIDCosteo = 0 Then
        MsgBox "No existe registro del costeo para el mes de " & Trim(tMes.Text), vbExclamation, "ATENCIÓN"
        Foco tMes: Screen.MousePointer = 0: Exit Sub
    End If
    '------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim aQ As Long
    aQ = 0
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    cons = "Select Count(*) from CMCosteo Where CosID = " & aIDCosteo
    If Val(tArticulo.Tag) <> 0 Then cons = cons & " And CosArticulo = " & Val(tArticulo.Tag)
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then aQ = rsAux(0)
    rsAux.Close
    
    If aQ = 0 Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        Screen.MousePointer = 0:  Exit Sub
    End If
    pbProgreso.Max = aQ
    '-------------------------------------------------------------------------------------------------------------------------------------------
    
    cons = "Select * from CMCosteo " _
                    & " Left Outer Join Renglon On CosIDVenta = RenDocumento And CosArticulo = RenArticulo And CosTipoVenta = " & TipoCV.Comercio _
                        & " Left Outer Join Documento On RenDocumento = DocCodigo" _
            & " Where CosID = " & aIDCosteo
    If Val(tArticulo.Tag) <> 0 Then cons = cons & " And CosArticulo = " & Val(tArticulo.Tag)
    cons = cons & " Order by CosArticulo"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        rsAux.Close: Screen.MousePointer = 0: InicializoGrillas: Exit Sub
    End If
    
    aCostoP = 0: aVentaP = 0: aCantidadP = 0
    aCostoT = 0: aVentaT = 0
    
    With vsConsulta
        .Rows = 1: .Refresh: .Redraw = False
        Do While Not rsAux.EOF
            pbProgreso.Value = pbProgreso.Value + 1
            
            If aArticulo <> rsAux!CosArticulo Then      '--------------------------------------------------------------
                If aArticulo <> 0 Then
                    InsertoTotalArticulo aTxtArticulo, aCantidadP, aCostoP, aVentaP
                    
                    aCostoT = aCostoT + aCostoP
                    aVentaT = aVentaT + aVentaP
                    aCostoP = 0: aVentaP = 0: aCantidadP = 0
                End If
                
                cons = "Select * from Articulo Where ArtID = " & rsAux!CosArticulo
                Set rs1 = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                If Not rsAux.EOF Then
                    aTxtArticulo = "(" & Format(rs1!ArtCodigo, "#,000,000") & ") " & Trim(rs1!ArtNombre)
                    aArticulo = rsAux!CosArticulo
                End If
                rs1.Close
                
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Trim(aTxtArticulo)
                .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Azul
            End If  '---------------------------------------------------------------------------------------------------------
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Trim(aTxtArticulo)
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!DocFecha, "dd/mm/yy")
            
            Select Case rsAux!CosTipoVenta
                Case TipoCV.Comercio
                    .Cell(flexcpText, .Rows - 1, 2) = RetornoNombreDocumento(rsAux!DocTipo, Abreviacion:=True)
                    .Cell(flexcpText, .Rows - 1, 3) = Trim(rsAux!DocSerie) & " " & rsAux!DocNumero
                
                Case TipoCV.Compra: .Cell(flexcpText, .Rows - 1, 2) = "Venta Simulada"
                
                Case TipoCV.Servicio:
                    .Cell(flexcpText, .Rows - 1, 2) = "SER"
                    .Cell(flexcpText, .Rows - 1, 3) = rsAux!CosIDVenta
                    'Saco fecha del servicio
                    cons = "Select * from Servicio Where SerCodigo = " & rsAux!CosIDVenta
                    Set rsSer = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                    If Not rsSer.EOF Then
                        If Not IsNull(rsSer!SerFCumplido) Then .Cell(flexcpText, .Rows - 1, 1) = Format(rsSer!SerFCumplido, "dd/mm/yy") Else .Cell(flexcpText, .Rows - 1, 1) = " "
                    End If
                    rsSer.Close
            End Select
            
            .Cell(flexcpText, .Rows - 1, 4) = rsAux!CosCantidad
            .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!CosCosto, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = Format(rsAux!CosVenta, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 7) = Format((rsAux!CosVenta - rsAux!CosCosto) * rsAux!CosCantidad, FormatoMonedaP)
            
            aCantidadP = aCantidadP + .Cell(flexcpValue, .Rows - 1, 4)
            aCostoP = aCostoP + (.Cell(flexcpValue, .Rows - 1, 4) * .Cell(flexcpValue, .Rows - 1, 5))
            aVentaP = aVentaP + (.Cell(flexcpValue, .Rows - 1, 4) * .Cell(flexcpValue, .Rows - 1, 6))
                        
            rsAux.MoveNext
        Loop
        rsAux.Close
        
        If aArticulo <> 0 Then
            InsertoTotalArticulo aTxtArticulo, aCantidadP, aCostoP, aVentaP
            aCostoT = aCostoT + aCostoP
            aVentaT = aVentaT + aVentaP
        End If
        
        .Select 1, 0, 1, 0
        .Sort = flexSortGenericAscending
        
        InsertoTotalArticulo "Total General", 0, aCostoT, aVentaT, True
        
        .Redraw = True
    End With
    
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Sub InsertoTotalArticulo(Articulo As String, Q As Long, Costo As Currency, Venta As Currency, Optional TGeneral As Boolean = False)

    With vsConsulta
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = Trim(Articulo)
        
        If Not TGeneral Then .Cell(flexcpText, .Rows - 1, 4) = Q
        .Cell(flexcpText, .Rows - 1, 5) = Format(Costo, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 6) = Format(Venta, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 7) = Format(Venta - Costo, FormatoMonedaP)
        
        If Not TGeneral Then
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Azul
        Else
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.osGris: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
        End If
    End With
    
End Sub

Private Sub tMes_LostFocus()
    If IsDate(tMes.Text) Then tMes.Text = Format(tMes.Text, "Mmmm yyyy")
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
            .Columns = 2
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        aTexto = "Listado del Costeo - " & Trim(tMes.Text)
        If Val(tArticulo.Tag) <> 0 Then aTexto = aTexto & " (" & Trim(tArticulo.Text) & ")"
        EncabezadoListado vsListado, aTexto, False
        vsListado.FileName = "Listado de Costos"
         
        With vsConsulta
            .Redraw = False
            .FontSize = 6
            AnchoEncabezado Impresora:=True
            vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
            AnchoEncabezado Pantalla:=True
            .FontSize = 8
            .Redraw = True
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
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

Private Sub vsListado_NewPage()
    lQ.Tag = Val(lQ.Tag) + 1
    lQ.Caption = lQ.Tag: lQ.Refresh
End Sub
