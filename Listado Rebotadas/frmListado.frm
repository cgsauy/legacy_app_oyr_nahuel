VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Begin VB.Form frmListado 
   Caption         =   "Listado de Ventas Rebotadas"
   ClientHeight    =   7530
   ClientLeft      =   1305
   ClientTop       =   2010
   ClientWidth     =   9330
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
   ScaleWidth      =   9330
   Begin VSFlex8LCtl.VSFlexGrid vsConsulta 
      Height          =   855
      Left            =   2040
      TabIndex        =   23
      Top             =   5400
      Width           =   2775
      _cx             =   4895
      _cy             =   1508
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
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
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   600
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
      ForeColor       =   &H00000080&
      Height          =   795
      Left            =   120
      TabIndex        =   21
      Top             =   60
      Width           =   10995
      Begin VB.CheckBox oServicios 
         Caption         =   "Servicios No facturados"
         Height          =   255
         Left            =   1740
         TabIndex        =   3
         Top             =   480
         Value           =   1  'Checked
         Width           =   2235
      End
      Begin VB.TextBox tArticulo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   240
         Width           =   4215
      End
      Begin VB.CheckBox oCompra 
         Caption         =   "Devoluciones de Compras"
         Height          =   255
         Left            =   1740
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox oContado 
         Caption         =   "Contados"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox oCredito 
         Caption         =   "Créditos"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   280
         Width           =   735
      End
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4455
      Left            =   120
      TabIndex        =   19
      Top             =   900
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
      ScaleWidth      =   6675
      TabIndex        =   20
      Top             =   6720
      Width           =   6735
      Begin VB.CommandButton butExcel 
         Height          =   310
         Left            =   3720
         Picture         =   "frmListado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Exportar a excel"
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0784
         Height          =   310
         Left            =   4800
         Picture         =   "frmListado.frx":0886
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0DB8
         Style           =   1  'Graphical
         TabIndex        =   14
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":1640
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   5160
         Picture         =   "frmListado.frx":1742
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5760
         Picture         =   "frmListado.frx":1B08
         Style           =   1  'Graphical
         TabIndex        =   17
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
         TabIndex        =   6
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   7
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
      TabIndex        =   18
      Top             =   7275
      Width           =   9330
      _ExtentX        =   16457
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
            Object.Width           =   8229
            TextSave        =   ""
            Key             =   ""
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
    Compra = 1              'Compra Comun (a proveedores de mercaderia locales)
    Comercio = 2            'Cualquier documento del comercio (ctdo, cred, etc...)
    Importacion = 3        'Compra (que entra por importaciones)
    Servicio = 4              'Documento ralacionado a Servicios (Ventas por servicios no facturados)
End Enum

Private RsAux As rdoResultset, rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean

Dim aTotalGeneral As Currency

Private Sub AccionLimpiar()
    vsConsulta.Rows = 1
    tArticulo.Text = ""
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

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        .Cols = 1: .Rows = 1:
        .FormatString = "Fecha|Factura|<Código|<Artículo|>Q|>Costo (x1)|>Costo Venta|"
        .WordWrap = False
        AnchoEncabezado Pantalla:=True
        .MergeCells = flexMergeSpill
        .MergeCol(0) = True
    End With
      
End Sub

Private Sub AnchoEncabezado(Optional Pantalla As Boolean = False, Optional Impresora As Boolean = False)

    With vsConsulta
        
        If Pantalla Then
            .ColWidth(0) = 950: .ColWidth(1) = 950: .ColWidth(2) = 950: .ColWidth(3) = 4000: .ColWidth(4) = 1000: .ColWidth(5) = 1400: .ColWidth(6) = 1400
        End If
        
        If Impresora Then
            .ColWidth(0) = 700: .ColWidth(1) = 650: .ColWidth(2) = 950: .ColWidth(3) = 1750: .ColWidth(4) = 400: .ColWidth(5) = 800: .ColWidth(6) = 1000
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
    
    
    vsListado.Width = Me.ScaleWidth - (vsListado.Left * 2)
    fFiltros.Width = vsListado.Width
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
    End
    
End Sub

Private Sub AccionConsultar()
    
    On Error GoTo errConsultar
    Screen.MousePointer = 11
    aTotalGeneral = 0
    bCargarImpresion = True
       
    vsConsulta.Rows = 1: vsConsulta.Refresh
    
    If oContado.Value = vbChecked Then              'Ventas Contado--------------------------------------------------------------------------------------
        Cons = "Select * from CMVenta, Documento, Articulo " _
                & " Where VenCodigo = DocCodigo " _
                & " And VenArticulo = ArtID" _
                & " And VenTipo = " & TipoCV.Comercio _
                & " And DocTipo IN (" & TipoDocumento.Contado & ", " & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial & ")"
        
        If Val(tArticulo.Tag) <> 0 Then Cons = Cons & " And VenArticulo = " & Val(tArticulo.Tag)
        
        Cons = Cons & " Order by VenFecha, DocNumero, DocSerie"
     
        CargoGrilla Cons, "Ventas Contado"
    End If
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If oCredito.Value = vbChecked Then              'Ventas Credito--------------------------------------------------------------------------------------
        Cons = "Select * from CMVenta, Documento, Articulo " _
                & " Where VenCodigo = DocCodigo " _
                & " And VenArticulo = ArtID" _
                & " And VenTipo = " & TipoCV.Comercio _
                & " And DocTipo IN (" & TipoDocumento.Credito & ", " & TipoDocumento.NotaCredito & ")"
        
        If Val(tArticulo.Tag) <> 0 Then Cons = Cons & " And VenArticulo = " & Val(tArticulo.Tag)
        
        Cons = Cons & " Order by VenFecha, DocNumero, DocSerie"
        
        CargoGrilla Cons, "Ventas Crédito"
    End If
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If oCompra.Value = vbChecked Then                'Devoluciones--------------------------------------------------------------------------------------
        Cons = "Select *, DocSerie= '', DocNumero= '' from CMVenta, Articulo " _
                & " Where VenArticulo = ArtID" _
                & " And VenTipo = " & TipoCV.Compra
        
        If Val(tArticulo.Tag) <> 0 Then Cons = Cons & " And VenArticulo = " & Val(tArticulo.Tag)
        Cons = Cons & " Order by VenFecha"
        
        CargoGrilla Cons, "Devoluciones de Compras"
    End If
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If oServicios.Value = vbChecked Then                'Servicios--------------------------------------------------------------------------------------
        Cons = "Select *, DocSerie= '', VenCodigo as DocNumero  from CMVenta, Articulo " _
                & " Where VenArticulo = ArtID" _
                & " And VenTipo = " & TipoCV.Servicio
        
        If Val(tArticulo.Tag) <> 0 Then Cons = Cons & " And VenArticulo = " & Val(tArticulo.Tag)
        Cons = Cons & " Order by VenFecha"
        
        CargoGrilla Cons, "Servicios No Facturados"
    End If
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If vsConsulta.Rows = 1 Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
    Else
        With vsConsulta
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = "Total de Ventas Rebotadas"
            .Cell(flexcpText, .Rows - 1, 5) = Format(aTotalGeneral, FormatoMonedaP)
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
            .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
        End With
        CargoResumen
    End If
        
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    clsGeneral.OcurrioError "Error al realizar la consulta de datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CargoResumen()

    
    With vsConsulta
        .AddItem "": .AddItem "": .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = "RESUMEN DE VENTAS REBOTADAS"
        '.Cell(flexcpText, .Rows - 1, 3) = "Q+/Q-": .Cell(flexcpText, .Rows - 1, 4) = "Costo Q+": .Cell(flexcpText, .Rows - 1, 5) = "Costo Q-"
        .Cell(flexcpText, .Rows - 1, 4) = "Q+/Q-": .Cell(flexcpText, .Rows - 1, 5) = "Costo Q+": .Cell(flexcpText, .Rows - 1, 6) = "Costo Q-"
        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
    End With
    
    
    Cons = "Select ArtCodigo, ArtNombre, Q = Sum(VenCantidad), P = Sum(VenCantidad * VenPrecio) " _
           & " From CMVenta, Articulo" _
           & " Where VenArticulo = ArtID " _
           & " And VenCantidad > 0 "
    If Val(tArticulo.Tag) <> 0 Then Cons = Cons & " And VenArticulo = " & Val(tArticulo.Tag)
    
    Cons = Cons & " Group by ArtCodigo, ArtNombre" _
                       & " Union All" _
           & " Select ArtCodigo, ArtNombre, Q = Sum(VenCantidad), P = Sum(VenCantidad * VenPrecio) " _
           & " From CMVenta, Articulo" _
           & " Where VenArticulo = ArtID " _
           & " And VenCantidad < 0 "

    If Val(tArticulo.Tag) <> 0 Then Cons = Cons & " And VenArticulo = " & Val(tArticulo.Tag)
    
    Cons = Cons & " Group by ArtCodigo, ArtNombre" _
                       & " Order by ArtCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    Dim aArticulo As String: aArticulo = ""
    
    Do While Not RsAux.EOF
        With vsConsulta
            
            If aArticulo <> Trim(RsAux!ArtNombre) Then
                aArticulo = Trim(RsAux!ArtNombre)
                .AddItem ""
                '.Cell(flexcpText, .Rows - 1, 1) = "(" & Format(RsAux!ArtCodigo, "#,000,000") & ") " & Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!ArtCodigo, "#,000,000")
                .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 4) = RsAux!Q
            Else
                If Trim(.Cell(flexcpText, .Rows - 1, 4)) = "" Then
                    .Cell(flexcpText, .Rows - 1, 4) = RsAux!Q
                Else
                    .Cell(flexcpText, .Rows - 1, 4) = .Cell(flexcpText, .Rows - 1, 4) & "/" & RsAux!Q
                End If
            End If
            
            If RsAux!P > 0 Then
                .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!P, FormatoMonedaP)
            Else
                .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!P, FormatoMonedaP)
            End If
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close


End Sub


Private Sub CargoGrilla(Consulta As String, Titulo As String)

Dim aTxtArticulo As String
Dim aTotalVentas As Currency
Dim aMes As Date, aMesRow As Date
    
    aTotalVentas = 0
    With vsConsulta
        aMes = CDate("1/1/1900")
        Set RsAux = cBase.OpenResultset(Consulta, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Trim(Titulo)
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
    
            Do While Not RsAux.EOF
                If Format(aMes, "mm/yyyy") = "01/1900" Then aMes = Format(RsAux!VenFecha, "mm/yyyy")
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!VenFecha, "dd/mm/yyyy")
                If Not IsNull(RsAux!DocSerie) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!DocSerie) & " " & RsAux!DocNumero
                
'                aTxtArticulo = "(" & Format(RsAux!ArtCodigo, "#,000,000") & ") " & Trim(RsAux!ArtNombre)
'                .Cell(flexcpText, .Rows - 1, 2) = Trim(aTxtArticulo)
                
                .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ArtCodigo, "#,000,000")
                .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!ArtNombre)
                
                .Cell(flexcpText, .Rows - 1, 4) = RsAux!VenCantidad
                .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!VenPrecio, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!VenPrecio * RsAux!VenCantidad, FormatoMonedaP)
                aTotalVentas = aTotalVentas + .Cell(flexcpValue, .Rows - 1, 6)
                aTotalGeneral = aTotalGeneral + .Cell(flexcpValue, .Rows - 1, 6)
                RsAux.MoveNext
                
                If Not RsAux.EOF Then
                    If Format(RsAux!VenFecha, "mm/yyyy") <> Format(aMes, "mm/yyyy") Then 'And Format(aMes, "mm/yyyy") <> "01/1900" Then
                        .AddItem ""
                        .Cell(flexcpText, .Rows - 1, 5) = Format(aMes, "Mmm yyyy")
                        .Cell(flexcpText, .Rows - 1, 6) = Format(aTotalVentas, FormatoMonedaP)
                        .Cell(flexcpBackColor, .Rows - 1, 3, , .Cols - 1) = Colores.Gris
                        
                        aMes = Format(RsAux!VenFecha, "mm/yyyy")
                        aTotalVentas = 0
                    Else
                        aMes = Format(RsAux!VenFecha, "mm/yyyy")
                    End If
                Else
                    .AddItem ""
                    .Cell(flexcpText, .Rows - 1, 5) = Format(aMes, "Mmm yyyy")
                    .Cell(flexcpText, .Rows - 1, 6) = Format(aTotalVentas, FormatoMonedaP)
                    .Cell(flexcpBackColor, .Rows - 1, 3, , .Cols - 1) = Colores.Gris
                    aTotalVentas = 0
                End If
                
            Loop
        End If
        RsAux.Close
        
    End With

End Sub

Private Sub oCompra_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then oServicios.SetFocus
End Sub

Private Sub oContado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then oCredito.SetFocus
End Sub

Private Sub oCredito_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then oCompra.SetFocus
End Sub

Private Sub oServicios_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tArticulo
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
            Cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtNombre Like '" & tArticulo.Text & "%'"
            Dim LiAyuda  As New clsListadeAyuda
            LiAyuda.ActivoListaAyuda Cons, False, miConexion.TextoConexion(logComercio)
            If LiAyuda.ItemSeleccionado <> "" Then tArticulo.Text = LiAyuda.ItemSeleccionado Else tArticulo.Text = "0"
            Set LiAyuda = Nothing
        End If
        
        If Val(tArticulo.Text) <> 0 Then
            Cons = "Select ArtID, 'Código' = ArtCodigo, 'Nombre' = ArtNombre From Articulo Where ArtCodigo = " & CLng(tArticulo.Text)
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
            If RsAux.EOF Then
                RsAux.Close
                MsgBox "No se encontró un artículo con ese código.", vbInformation, "ATENCIÓN"
            Else
                tArticulo.Text = Trim(RsAux!Nombre)
                tArticulo.Tag = RsAux!ArtID
                RsAux.Close
                Foco bConsultar
            End If
        Else
            tArticulo.Text = ""
        End If
        Screen.MousePointer = 0
        
    Else
        If KeyCode = vbKeyReturn Then Foco bConsultar
    End If
    
    Exit Sub
ErrTA:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo.", Err.Description
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
            .StartDoc
'            .Columns = 2
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        Cons = "select Max(CabMesCosteo) from CMCabezal"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If Not IsNull(RsAux(0)) Then aTexto = "al " & Format(UltimoDia(RsAux(0)), "dd Mmmm yyyy")
        End If
        RsAux.Close
        
        EncabezadoListado vsListado, "Listado de Ventas Rebotadas " & aTexto, False
        vsListado.FileName = "Listado de Ventas Rebotadas"
            
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

