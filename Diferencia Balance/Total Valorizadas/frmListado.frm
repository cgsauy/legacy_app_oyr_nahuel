VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmListado 
   Caption         =   "Diferencias Balance"
   ClientHeight    =   7530
   ClientLeft      =   1710
   ClientTop       =   1815
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
      TabIndex        =   14
      Top             =   60
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
      Width           =   11880
      _ExtentX        =   20955
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
            Object.Width           =   12832
            TextSave        =   ""
            Object.Tag             =   ""
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

Enum TipoCV
    Compra = 1
    Comercio = 2
    Importacion = 3
End Enum

Private RsAux As rdoResultset, rs1 As rdoResultset
Dim rsRep As rdoResultset

Private aTexto As String
Dim bCargarImpresion As Boolean

Private Sub AccionLimpiar()
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
        .Zoom = 100
        .MarginLeft = 800: .MarginRight = 250
        .MarginBottom = 750: .MarginTop = 750
    End With
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .OutlineBar = flexOutlineBarNone ' = flexOutlineBarComplete
        .OutlineCol = 0
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<Tipo|<Artículo|>Se Contó|>LIFO|>Diferencia|>Q Rotos|>Valorizada|"
            
        .WordWrap = False: .MergeCells = flexMergeSpill
        '.ColWidth(0) = 3800: .ColWidth(1) = 800: .ColWidth(2) = 800: .ColWidth(5) = 1500: .ColWidth(6) = 1200
        .ColHidden(0) = True
        .ColWidth(1) = 3800: .ColWidth(2) = 800: .ColWidth(3) = 800: .ColWidth(6) = 1500:
        
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
Dim aDif As Long

    On Error GoTo errConsultar
    Screen.MousePointer = 11
    bCargarImpresion = True
    
    Dim aQ As Long, QRoto As Long, QSano As Long
    Dim rsBal As rdoResultset
    aQ = 0
    
    Cons = "Select Count(Distinct(ComArticulo)) From CMCompra"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then aQ = aQ + RsAux(0)
    RsAux.Close
    
    Cons = "Select Count(Distinct(CCBArticulo)) from ContadoCalcBalance" & _
                " Where CCBArticulo Not In (Select ComArticulo from CMCompra)"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then aQ = aQ + RsAux(0)
    RsAux.Close
    
    If aQ = 0 Then
        MsgBox "No hay datos a desplegar.", vbInformation, "ATENCION"
        Screen.MousePointer = 0: vsConsulta.Rows = 1: Exit Sub
    End If
    
    pbProgreso.Min = 0
    pbProgreso.Max = aQ
    
    With vsConsulta
        .Rows = 1
        Me.Refresh: .Redraw = False
        pbProgreso.Value = 0
        'Articulo De la existencia--------------------------------------------------------------------------------------------------------
        Cons = "Select  ArtID, ArtCodigo, ArtNombre, Sum(ComCantidad) as Q" & _
                    " From CMCompra, Articulo " & _
                    " Where ComArticulo = ArtID" & _
                    " Group by ArtID, ArtCodigo, ArtNombre"
                    
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            pbProgreso.Value = pbProgreso.Value + 1
            
            QRoto = 0: QSano = 0
            Cons = "Select CCBEstado, Sum(CCBCantidad) as Q From ContadoCalcBalance Where CCBArticulo = " & RsAux!ArtID & " Group by CCBEstado"
            Set rsBal = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            Do While Not rsBal.EOF
                Select Case rsBal!CCBEstado
                    Case paEstadoArticuloEntrega, paEstadoArticuloARecuperar: QSano = QSano + rsBal!Q
                    Case Else: QRoto = QRoto + rsBal!Q
                End Select
                rsBal.MoveNext
            Loop
            rsBal.Close
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
            If EsRepuesto(RsAux!ArtCodigo) Then .Cell(flexcpText, .Rows - 1, 0) = "Repuestos " Else .Cell(flexcpText, .Rows - 1, 0) = "Mercadería"
            
            .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!Q, "#,##0")
            
            .Cell(flexcpText, .Rows - 1, 2) = Format(QSano, "#,##0")
            If QRoto <> 0 Then .Cell(flexcpText, .Rows - 1, 5) = Format(QRoto, "#,##0")
            
            aDif = .Cell(flexcpValue, .Rows - 1, 2) - .Cell(flexcpValue, .Rows - 1, 3)
            .Cell(flexcpText, .Rows - 1, 4) = Format(aDif, "#,##0")
            
            If aDif < 0 Then BuscoUltimoCosto RsAux!ArtID, aDif
            
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        'Articulos del Conteo Q' no estan en la existencia-------------------------------------------------------------------------------------
        Cons = "Select ArtID, ArtCodigo, ArtNombre, Sum(CCBCantidad) as Q" & _
                   " From ContadoCalcBalance, Articulo" & _
                   " Where CCBArticulo Not In (Select ComArticulo from CMCompra)" & _
                   " And CCBArticulo = ArtID " & _
                   " Group by ArtID, ArtCodigo, ArtNombre"
                   
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            pbProgreso.Value = pbProgreso.Value + 1
            
            QSano = RsAux!Q
            
            QRoto = 0
            Cons = "Select Sum(CCBCantidad) as Q From ContadoCalcBalance Where CCBArticulo = " & RsAux!ArtID & " And CCBEstado Not In (" & paEstadoArticuloEntrega & ", " & paEstadoArticuloARecuperar & ")"
            Set rsBal = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rsBal.EOF Then If Not IsNull(rsBal!Q) Then QRoto = rsBal!Q
            rsBal.Close
            
            QSano = QSano - QRoto
                       
            .AddItem ""
            If EsRepuesto(RsAux!ArtCodigo) Then .Cell(flexcpText, .Rows - 1, 0) = "Repuestos " Else .Cell(flexcpText, .Rows - 1, 0) = "Mercadería"
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
            
            .Cell(flexcpText, .Rows - 1, 2) = Format(QSano, "#,##0")
            
            .Cell(flexcpText, .Rows - 1, 3) = 0
            .Cell(flexcpText, .Rows - 1, 4) = Format(.Cell(flexcpValue, .Rows - 1, 2), "#,##0")
            
            .Cell(flexcpText, .Rows - 1, 5) = Format(QRoto, "#,##0")
            
            .Cell(flexcpText, .Rows - 1, 6) = "0.00" 'Format(QRoto, "#,##0.00")
                
            RsAux.MoveNext
        Loop
        RsAux.Close
        
        .Select 1, 0, 1, 2
        .Sort = flexSortGenericAscending
        
        .Subtotal flexSTSum, 0, 6, , vbInactiveBorder, , True
        .Subtotal flexSTSum, 0, 2, "#,##0": .Subtotal flexSTSum, 0, 3, "#,##0": .Subtotal flexSTSum, 0, 4, "#,##0": .Subtotal flexSTSum, 0, 5, "#,##0"
        
        .Subtotal flexSTSum, -1, 6, , vbInactiveBorder, , True, "TOTAL"
        .Subtotal flexSTSum, -1, 2, "#,##0": .Subtotal flexSTSum, -1, 3, "#,##0": .Subtotal flexSTSum, -1, 4, "#,##0": .Subtotal flexSTSum, -1, 5, "#,##0"
        pbProgreso.Value = 0
        .Redraw = True
    End With
    
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub BuscoUltimoCosto(idArticulo As Long, Cantidad As Long)

Dim rsCom As rdoResultset
Dim Costo1 As Currency, aQ As Long

    Costo1 = 0: aQ = Abs(Cantidad)
    'Primero Compras con costo 0 ------------------------------------------------------------
    Cons = "Select * from CMCompra " _
            & " Where ComFecha <= '06/30/2000'" _
            & " And ComArticulo =  " & idArticulo _
            & " And ComCantidad > 0 " _
            & " And ComCosto = 0" _
            & " Order by ComFecha DESC"
    Set rsCom = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsCom.EOF
        aQ = aQ - rsCom!ComCantidad
        If aQ <= 0 Then Exit Do
        rsCom.MoveNext
    Loop
    rsCom.Close
    '-----------------------------------------------------------------------------------------------
    '2) Con Costo--------------------------------------------------------------------------------
    If aQ > 0 Then
        Cons = "Select * from CMCompra " _
                & " Where ComFecha <= '06/30/2000'" _
                & " And ComArticulo =  " & idArticulo _
                & " And ComCantidad > 0 " _
                & " And ComCosto > 0" _
                & " Order by ComFecha DESC"
        Set rsCom = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
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
            .Columns = 1
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        aTexto = "Diferencias Balance"
        EncabezadoListado vsListado, aTexto, False
        vsListado.filename = "Diferencias Balance"
            
        With vsConsulta
            .Redraw = False
            vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hWnd: vsConsulta.ExtendLastCol = True
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

Public Function EsRepuesto(idArticulo As Long) As Boolean

    If idArticulo > 250000 Then EsRepuesto = True Else EsRepuesto = False
    Exit Function
    
    Cons = "Select * from ArticuloGrupo Where AGrArticulo = " & idArticulo & " And AGrGrupo = " & paGrupoRepuesto
    Set rsRep = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not rsRep.EOF Then EsRepuesto = True Else EsRepuesto = False
    rsRep.Close
    
End Function
