VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmCompra 
   Caption         =   "Compra de Mercadería"
   ClientHeight    =   6690
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10845
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCompra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   10845
   Begin VB.PictureBox picBotones 
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   120
      ScaleHeight     =   405
      ScaleWidth      =   2175
      TabIndex        =   18
      Top             =   6000
      Width           =   2175
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmCompra.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   1800
         Picture         =   "frmCompra.frx":0744
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   1080
         Picture         =   "frmCompra.frx":0846
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   50
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   720
         Picture         =   "frmCompra.frx":0C0C
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   50
         Width           =   310
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   4320
      TabIndex        =   14
      Top             =   1560
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   6435
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "sucursal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   8361
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
      Height          =   1020
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   9735
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   5
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox tHasta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin AACombo99.AACombo cTipo 
         Height          =   315
         Left            =   840
         TabIndex        =   7
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
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
      Begin VB.TextBox tDesde 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   5160
         TabIndex        =   9
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
      End
      Begin VB.Label Label5 
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   615
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Moneda:"
         Height          =   255
         Left            =   4200
         TabIndex        =   8
         Top             =   615
         Width           =   1095
      End
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4455
      Left            =   240
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   9615
      _Version        =   196608
      _ExtentX        =   16960
      _ExtentY        =   7858
      _StockProps     =   229
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
End
Attribute VB_Name = "frmCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsAux As rdoResultset
Private aTexto As String

Private Sub AccionLimpiar()
    cTipo.Text = ""
    cMoneda.Text = ""
    tDesde.Text = "": tHasta.Text = ""
    tProveedor.Text = ""
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub
Private Sub bImprimir_Click()
    AccionImprimir
End Sub
Private Sub bNoFiltros_Click()
    AccionLimpiar
End Sub

Private Sub cTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco cMoneda
End Sub

Private Sub cMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco bConsultar
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()

    On Error GoTo ErrLoad

    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    InicializoGrillas
    AccionLimpiar
    
    'Cargo los datos-------------------------------------------------------------------------------------------------
    tDesde.Text = Format(PrimerDia(Now), "dd/mm/yyyy")
    tHasta.Text = Format(UltimoDia(Now), "dd/mm/yyyy")
    
    cTipo.AddItem "Contado"
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraContado
    cTipo.AddItem "Crédito"
    cTipo.ItemData(cTipo.NewIndex) = TipoDocumento.CompraCredito

    cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo cons, cMoneda
    '--------------------------------------------------------------------------------------------------------------------
    
    PropiedadesImpresion
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = "Fecha|Proveedor|<Factura|Artículo|>Q|>Unitario|>Costo|>Cofis|>I.V.A.|>Total|>Total U$S|"
            
        .WordWrap = False
        .ColWidth(0) = 900: .ColWidth(1) = 1300: .ColWidth(2) = 900: .ColWidth(3) = 2700: .ColWidth(4) = 500: .ColWidth(5) = 1100: .ColWidth(6) = 1200
        .ColWidth(7) = 950: .ColWidth(8) = 1200: .ColWidth(9) = 1300
        .ColWidth(10) = 1300
        
        .ColDataType(5) = flexDTCurrency: .ColDataType(6) = flexDTCurrency: .ColDataType(7) = flexDTCurrency: .ColDataType(8) = flexDTCurrency: .ColDataType(9) = flexDTCurrency
    End With
      
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyE: AccionConsultar
            Case vbKeyQ: AccionLimpiar
            Case vbKeyI: AccionImprimir
            
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11
    picBotones.BorderStyle = vbFlat
    fFiltros.Width = Me.Width - (fFiltros.Left * 2.5)
    
    picBotones.Top = Me.ScaleHeight - (Status.Height + picBotones.Height + 90)
    
    vsConsulta.Left = fFiltros.Left
    vsConsulta.Top = fFiltros.Top + fFiltros.Height + 50
    vsConsulta.Height = Me.ScaleHeight - (vsConsulta.Top + Status.Height + picBotones.Height + 120)
    vsConsulta.Width = fFiltros.Width
    
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub AccionConsultar()

Dim aUnitario As Currency, aTotal As Currency, aIva As Currency, aCofis As Currency
Dim aCompra As Long
Dim aDolares As Currency

Dim aTCostoContado As Currency, aTCostoCredito As Currency, aTCostos As Currency
Dim aTIvaContado As Currency, aTIvaCredito As Currency
Dim aTCofisContado As Currency, aTCofisCredito As Currency
Dim aTDolares As Currency
Dim aOtrosCostos As Currency

    On Error GoTo errConsultar
    If Not ValidoFiltros Then Exit Sub
    
    Screen.MousePointer = 11
    'Armo la consulta de datos------------------------------------------------------------------------------------------------------------------------------
    cons = " Select * from Compra, CompraRenglon, ProveedorCliente, Articulo" _
        & " Where ComFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59") & "'"
        
        '" Where ComFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59:59") & "'"
            
    If cTipo.ListIndex <> -1 Then
        Select Case cTipo.ItemData(cTipo.ListIndex)
            Case TipoDocumento.CompraContado: cons = cons & " And ComTipoDocumento In (" & TipoDocumento.CompraContado & ", " & TipoDocumento.CompraNotaDevolucion & ")"
            Case TipoDocumento.CompraCredito: cons = cons & " And ComTipoDocumento In (" & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraNotaCredito & ")"
        End Select
    Else
        cons = cons & " And ComTipoDocumento In (" & TipoDocumento.CompraContado & ", " & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraNotaDevolucion & ")"
    End If
    
    If cMoneda.ListIndex <> -1 Then cons = cons & " And ComMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
            
    If Val(tProveedor.Tag) <> 0 Then cons = cons & " And ComProveedor = " & Val(tProveedor.Tag)
    
    cons = cons & " And ComCodigo = CReCompra " _
                       & " And ComProveedor = PClCodigo " _
                       & " And CReArticulo = ArtId " _
                       & " Order by ComFecha, ComCodigo"
    '------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        vsConsulta.Rows = 1
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        rsAux.Close: Screen.MousePointer = 0: Exit Sub
    End If
    
    aTCostoContado = 0: aTCostoCredito = 0
    aTIvaContado = 0: aTIvaCredito = 0
    aTCofisContado = 0: aTCofisCredito = 0
    aDolares = 0: aTDolares = 0
    aTCostos = 0: aOtrosCostos = 0
    
    With vsConsulta
        .Rows = 1: aCompra = 0
        aTotal = 0: aIva = 0: aCofis = 0
        Do While Not rsAux.EOF
            '"Fecha|Proveedor|<Factura|Artículo|>Q|>Unitario|>Costo|>I.V.A.|>Total"
            
            If aCompra <> rsAux!ComCodigo Then
                If aCompra <> 0 Then    'Inserto Total Factura
                    .AddItem ""
                    .Cell(flexcpText, .Rows - 1, 6) = Format(aTotal, FormatoMonedaP)
                    .Cell(flexcpText, .Rows - 1, 7) = Format(aCofis, FormatoMonedaP)
                    .Cell(flexcpText, .Rows - 1, 8) = Format(aIva, FormatoMonedaP)
                    .Cell(flexcpText, .Rows - 1, 9) = Format(aTotal + aIva + aCofis, FormatoMonedaP)
                    If aDolares <> 0 Then .Cell(flexcpText, .Rows - 1, 10) = Format(aDolares, FormatoMonedaP)
                    
                    aTotal = 0: aIva = 0: aCofis = 0
                    aDolares = 0
                    .AddItem ""
                End If
                aCompra = rsAux!ComCodigo
                
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!ComFecha, "dd/mm/yy")
                .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!PClNombre)
                
                If Not IsNull(rsAux!ComSerie) Then aTexto = Trim(rsAux!ComSerie) & " " Else aTexto = ""
                If Not IsNull(rsAux!ComNumero) Then aTexto = aTexto & Trim(rsAux!ComNumero)
                .Cell(flexcpText, .Rows - 1, 2) = Trim(aTexto)
                
                If rsAux!ComMoneda <> paMonedaPesos Then
                    aDolares = rsAux!ComImporte
                    If Not IsNull(rsAux!ComIva) Then aDolares = aDolares + rsAux!ComIva
                    If Not IsNull(rsAux!ComCofis) Then aDolares = aDolares + rsAux!ComCofis
                                        
                    aTotal = Format(rsAux!ComImporte * rsAux!ComTC, FormatoMonedaP)
                    If Not IsNull(rsAux!ComIva) Then aIva = rsAux!ComIva * rsAux!ComTC
                    If Not IsNull(rsAux!ComCofis) Then aCofis = rsAux!ComCofis * rsAux!ComTC
                    
                Else
                    aTotal = rsAux!ComImporte
                    If Not IsNull(rsAux!ComIva) Then aIva = rsAux!ComIva
                    If Not IsNull(rsAux!ComCofis) Then aCofis = rsAux!ComCofis
                End If
                
                'Agregue esta línea por el cambio a Zureo: antes las notas se grababan en negativo y ahora en positivo
                If rsAux!ComTipoDocumento = TipoDocumento.CompraNotaCredito Or rsAux!ComTipoDocumento = TipoDocumento.CompraNotaDevolucion Then
                    aDolares = Abs(aDolares) * -1
                    aTotal = Abs(aTotal) * -1
                    aIva = Abs(aIva) * -1
                    aCofis = Abs(aCofis) * -1
                End If
                
                aTDolares = aTDolares + aDolares
                
                If rsAux!ComTipoDocumento = TipoDocumento.CompraContado Or rsAux!ComTipoDocumento = TipoDocumento.CompraNotaDevolucion Then
                    .Cell(flexcpText, .Rows - 1, 0) = .Cell(flexcpText, .Rows - 1, 0) & " *"
                    aTCostoContado = aTCostoContado + aTotal
                    aTIvaContado = aTIvaContado + aIva
                    aTCofisContado = aTCofisContado + aCofis
                Else
                    aTCostoCredito = aTCostoCredito + aTotal
                    aTIvaCredito = aTIvaCredito + aIva
                    aTCofisCredito = aTCofisCredito + aCofis
                End If
                
            Else
                .AddItem ""
            End If
            
            If rsAux!ComMoneda <> paMonedaPesos Then aUnitario = rsAux!CRePrecioU * rsAux!ComTC Else aUnitario = rsAux!CRePrecioU
            
            .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!ArtCodigo, "#,000,000") & " " & Trim(rsAux!ArtNombre)
            .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!CReCantidad, "#,##0")
            
            .Cell(flexcpText, .Rows - 1, 5) = Format(aUnitario, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpText, .Rows - 1, 5) * rsAux!CReCantidad, FormatoMonedaP)
            
            If rsAux!ComTipoDocumento = TipoDocumento.CompraNotaCredito Or rsAux!ComTipoDocumento = TipoDocumento.CompraNotaDevolucion Then
                .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpValue, .Rows - 1, 6) * -1, FormatoMonedaP)
                .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.RojoClaro
            End If
            aTCostos = aTCostos + .Cell(flexcpValue, .Rows - 1, 6)
            
            If Not rsAux!ArtAMercaderia Then
                .Cell(flexcpFontItalic, .Rows - 1, 3, , 6) = True
                aOtrosCostos = aOtrosCostos + .Cell(flexcpValue, .Rows - 1, 6)
            End If
            rsAux.MoveNext
        Loop
        
        .AddItem ""     'Agrego el total de la ultima factura
        .Cell(flexcpText, .Rows - 1, 6) = Format(aTotal, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 7) = Format(aCofis, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 8) = Format(aIva, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 9) = Format(aTotal + aIva + aCofis, FormatoMonedaP)
        If aDolares <> 0 Then .Cell(flexcpText, .Rows - 1, 10) = Format(aDolares, FormatoMonedaP)
        .AddItem ""
        
        rsAux.Close
        
        'Lineas finales de Totals------------------------------------------------------------------------------------------------------------------------
        .AddItem "": .Cell(flexcpText, .Rows - 1, 3) = "TOTAL CONTADO"
        .Cell(flexcpText, .Rows - 1, 6) = Format(aTCostoContado, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 7) = Format(aTCofisContado, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 8) = Format(aTIvaContado, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 9) = Format(aTCostoContado + aTIvaContado + aTCofisContado, FormatoMonedaP)
        
        .AddItem "": .Cell(flexcpText, .Rows - 1, 3) = "TOTAL CREDITO"
        .Cell(flexcpText, .Rows - 1, 6) = Format(aTCostoCredito, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 7) = Format(aTCofisCredito, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 8) = Format(aTIvaCredito, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 9) = Format(aTCostoCredito + aTIvaCredito + aTCofisCredito, FormatoMonedaP)
        
        .AddItem "": .Cell(flexcpText, .Rows - 1, 3) = "TOTAL COMPRAS"
        .Cell(flexcpText, .Rows - 1, 6) = Format(aTCostoCredito + aTCostoContado, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 7) = Format(aTCofisCredito + aTCofisContado, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 8) = Format(aTIvaCredito + aTIvaContado, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 9) = Format(aTCostoCredito + aTCostoContado + aTIvaCredito + aTIvaContado + aTCofisCredito + aTCofisContado, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 10) = Format(aTDolares, FormatoMonedaP)
        
        .AddItem "": .Cell(flexcpText, .Rows - 1, 3) = "TOTAL COSTOS"
        .Cell(flexcpText, .Rows - 1, 6) = Format(aTCostos, FormatoMonedaP)
        
        .Cell(flexcpBackColor, .Rows - 4, 3, .Rows - 1, .Cols - 1) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 4, 3, .Rows - 1, .Cols - 1) = True
        
        If aOtrosCostos > 0 Then
            .AddItem "": .Cell(flexcpText, .Rows - 1, 3) = "OTROS COSTOS"
            .Cell(flexcpText, .Rows - 1, 6) = Format(aOtrosCostos, FormatoMonedaP)
            .Cell(flexcpBackColor, .Rows - 1, 3, , .Cols - 1) = Colores.osGris: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco: .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
        End If
        '----------------------------------------------------------------------------------------------------------------------------------------------------
     
        .AddItem ""
        aTCostoContado = 0: aTCostoCredito = 0
        aTIvaContado = 0: aTIvaCredito = 0
        aTCofisContado = 0: aTCofisCredito = 0
        aDolares = 0: aTDolares = 0
        aTCostos = 0: aOtrosCostos = 0

        
        '2) Cargo Todas las Facturas Asignadas al Subrubro Compra de Mercaderia q no soon "compras" -----------------------------------------------------
        cons = " Select * from Compra, ProveedorCliente, GastoSubRubro" _
                & " Where ComFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59") & "'"
                ' " Where ComFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59:59") & "'"
                
                If cTipo.ListIndex <> -1 Then
                    Select Case cTipo.ItemData(cTipo.ListIndex)
                        Case TipoDocumento.CompraContado: cons = cons & " And ComTipoDocumento In (" & TipoDocumento.CompraContado & ", " & TipoDocumento.CompraNotaDevolucion & ")"
                        Case TipoDocumento.CompraCredito: cons = cons & " And ComTipoDocumento In (" & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraNotaCredito & ")"
                    End Select
                Else
                    cons = cons & " And ComTipoDocumento In (" & TipoDocumento.CompraContado & ", " & TipoDocumento.CompraCredito & ", " & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraNotaDevolucion & ")"
                End If
        
                If cMoneda.ListIndex <> -1 Then cons = cons & " And ComMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
                If Val(tProveedor.Tag) <> 0 Then cons = cons & " And ComProveedor = " & Val(tProveedor.Tag)
        
        cons = cons & " And ComCodigo Not IN (Select CReCompra From CompraRenglon)" & _
                           " And ComCodigo = GSrIDCompra And GSRIDSubRubro = " & paSubrubroCompraMercaderia & _
                           " And ComProveedor = PClCodigo " & _
                           " Order by ComFecha, ComCodigo"
        '------------------------------------------------------------------------------------------------------------------------------------------------------------
    
        Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        Do While Not rsAux.EOF
        With vsConsulta
            aTotal = 0: aIva = 0: aCofis = 0
            '"Fecha|Proveedor|<Factura|Artículo|>Q|>Unitario|>Costo|>I.V.A.|>Total"
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!ComFecha, "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!PClNombre)

            If Not IsNull(rsAux!ComSerie) Then aTexto = Trim(rsAux!ComSerie) & " " Else aTexto = ""
            If Not IsNull(rsAux!ComNumero) Then aTexto = aTexto & Trim(rsAux!ComNumero)
            .Cell(flexcpText, .Rows - 1, 2) = Trim(aTexto)
            
            If rsAux!ComMoneda <> paMonedaPesos Then
                aDolares = rsAux!ComImporte
                If Not IsNull(rsAux!ComIva) Then aDolares = aDolares + rsAux!ComIva
                If Not IsNull(rsAux!ComCofis) Then aDolares = aDolares + rsAux!ComCofis
                
                
                aTotal = Format(rsAux!ComImporte * rsAux!ComTC, FormatoMonedaP)
                If Not IsNull(rsAux!ComIva) Then aIva = rsAux!ComIva * rsAux!ComTC
                If Not IsNull(rsAux!ComCofis) Then aCofis = rsAux!ComCofis * rsAux!ComTC
                
            Else
                aDolares = 0
                aTotal = rsAux!ComImporte
                If Not IsNull(rsAux!ComIva) Then aIva = rsAux!ComIva
                If Not IsNull(rsAux!ComCofis) Then aCofis = rsAux!ComCofis
            End If
            
            aTDolares = aTDolares + aDolares
            
            'Agregue esta línea por el cambio a Zureo: antes las notas se grababan en negativo y ahora en positivo
            If rsAux!ComTipoDocumento = TipoDocumento.CompraNotaCredito Or rsAux!ComTipoDocumento = TipoDocumento.CompraNotaDevolucion Then
                aDolares = Abs(aDolares) * -1
                aTotal = Abs(aTotal) * -1
                aIva = Abs(aIva) * -1
                aCofis = Abs(aCofis) * -1
            End If
            
            
            .Cell(flexcpText, .Rows - 1, 6) = Format(aTotal, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 7) = Format(aCofis, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 8) = Format(aIva, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 9) = Format(aTotal + aIva + aCofis, FormatoMonedaP)
            If aDolares <> 0 Then .Cell(flexcpText, .Rows - 1, 10) = Format(aDolares, FormatoMonedaP)
            
            If rsAux!ComTipoDocumento = TipoDocumento.CompraNotaCredito Or rsAux!ComTipoDocumento = TipoDocumento.CompraNotaDevolucion Then
                .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.RojoClaro
            End If
            If rsAux!ComTipoDocumento = TipoDocumento.CompraContado Or rsAux!ComTipoDocumento = TipoDocumento.CompraNotaDevolucion Then
                .Cell(flexcpText, .Rows - 1, 0) = .Cell(flexcpText, .Rows - 1, 0) & " *"
                aTCostoContado = aTCostoContado + aTotal
                aTIvaContado = aTIvaContado + aIva
                aTCofisContado = aTCofisContado + aCofis
            Else
                aTCostoCredito = aTCostoCredito + aTotal
                aTIvaCredito = aTIvaCredito + aIva
                aTCofisCredito = aTCofisCredito + aCofis
            End If
            rsAux.MoveNext
        End With
        Loop
        rsAux.Close
        '--------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        'Lineas finales de Totals------------------------------------------------------------------------------------------------------------------------
        If aTCostoContado + aTIvaContado + aTCofisContado <> 0 Then
            .AddItem "": .Cell(flexcpText, .Rows - 1, 3) = "TOTAL CONTADO"
            .Cell(flexcpText, .Rows - 1, 6) = Format(aTCostoContado, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 7) = Format(aTCofisContado, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 8) = Format(aTIvaContado, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 9) = Format(aTCostoContado + aTIvaContado + aTCofisContado, FormatoMonedaP)
            .Cell(flexcpBackColor, .Rows - 1, 3, .Rows - 1, .Cols - 1) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, 3, .Rows - 1, .Cols - 1) = True
        End If
        If aTCostoCredito + aTIvaCredito + aTCofisCredito <> 0 Then
            .AddItem "": .Cell(flexcpText, .Rows - 1, 3) = "TOTAL CREDITO"
            .Cell(flexcpText, .Rows - 1, 6) = Format(aTCostoCredito, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 7) = Format(aTCofisCredito, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 8) = Format(aTIvaCredito, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 9) = Format(aTCostoCredito + aTIvaCredito + aTCofisCredito, FormatoMonedaP)
            .Cell(flexcpBackColor, .Rows - 1, 3, .Rows - 1, .Cols - 1) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, 3, .Rows - 1, .Cols - 1) = True
        End If
        If aTCostoCredito + aTCostoContado + aTIvaCredito + aTIvaContado + aTCofisCredito + aTCofisContado <> 0 Then
            .AddItem "": .Cell(flexcpText, .Rows - 1, 3) = "TOTAL DTOS COMPRAS"
            .Cell(flexcpText, .Rows - 1, 6) = Format(aTCostoCredito + aTCostoContado, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 7) = Format(aTCofisCredito + aTCofisContado, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 8) = Format(aTIvaCredito + aTIvaContado, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 9) = Format(aTCostoCredito + aTCostoContado + aTIvaCredito + aTIvaContado + aTCofisCredito + aTCofisContado, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 10) = Format(aTDolares, FormatoMonedaP)
            .Cell(flexcpBackColor, .Rows - 1, 3, .Rows - 1, .Cols - 1) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, 3, .Rows - 1, .Cols - 1) = True
        End If
        '----------------------------------------------------------------------------------------------------------------------------------------------------
                
        Screen.MousePointer = 0
        
    End With
    Exit Sub

errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Function ValidoFiltros() As Boolean

    ValidoFiltros = False
   
    If Not IsDate(tDesde.Text) Then
        MsgBox "Debe ingresar una fecha válida para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    If Not IsDate(tHasta.Text) Then
        MsgBox "Debe ingresar una fecha válida para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco tHasta: Exit Function
    End If
    
    If CDate(tDesde.Text) > CDate(tHasta.Text) Then
        MsgBox "El rango de fechas para realizar la consulta no es correcto..", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    ValidoFiltros = True
    
End Function

Private Sub AccionImprimir()
Dim J As Integer

    If vsConsulta.Rows = 1 Then
        MsgBox "No hay datos en la lista para realizar la impresión.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    On Error GoTo errPrint
    Screen.MousePointer = 11
    
    With vsListado
        .Orientation = orLandscape
        If Not .PrintDialog(pdPrinterSetup) Then Screen.MousePointer = 0: Exit Sub
        
        .Preview = True
        .StartDoc
                
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN"
            Screen.MousePointer = vbDefault: Exit Sub
        End If
    
        EncabezadoListado vsListado, "Listado de Compras desde " & Trim(tDesde.Text) & " hasta " & Trim(tHasta.Text), False
        .FileName = "Compra de Mercadería"
        
        vsConsulta.ExtendLastCol = False: .RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        
        .EndDoc
        .PrintDoc
        
    End With
    Screen.MousePointer = 0
    Exit Sub

errPrint:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión. ", Err.Description
End Sub

Private Sub Label1_Click()
    Foco tDesde
End Sub

Private Sub Label4_Click()
    Foco tHasta
End Sub

Private Sub tDesde_GotFocus()
    With tDesde: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(tDesde.Text) Then tHasta.Text = Format(UltimoDia(tDesde.Text), "dd/mm/yyyy")
        Foco tHasta
    End If
End Sub

Private Sub tDesde_LostFocus()
    If IsDate(tDesde.Text) Then tDesde.Text = Format(tDesde.Text, "dd/mm/yyyy")
End Sub

Private Sub tHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tProveedor
End Sub

Private Sub tHasta_LostFocus()
    If IsDate(tHasta.Text) Then tHasta.Text = Format(tHasta.Text, "dd/mm/yyyy")
End Sub

Private Sub tProveedor_Change()
    tProveedor.Tag = 0
End Sub

Private Sub tProveedor_GotFocus()
    tProveedor.SelStart = 0: tProveedor.SelLength = Len(tProveedor.Text)
End Sub

Private Sub tProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF1 And Val(tProveedor.Tag) <> 0 Then AccionListaDeAyuda
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
    On Error GoTo errBuscar
    
    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) <> 0 Or Trim(tProveedor.Text) = "" Then Foco cTipo: Exit Sub
        
        Screen.MousePointer = 11
        cons = "Select PClCodigo, PClNombre, PClFantasia from ProveedorCliente " _
                & " Where PClNombre like '" & Trim(tProveedor.Text) & "%' Or PClFantasia like '" & Trim(tProveedor.Text) & "%'"
        
        Dim aLista As New clsListadeAyuda
        aLista.ActivoListaAyuda cons, False, miConexion.TextoConexion(logComercio), 5500
        If aLista.ValorSeleccionado <> 0 Then
            tProveedor.Text = Trim(aLista.ItemSeleccionado)
            tProveedor.Tag = aLista.ValorSeleccionado
            
            Foco cTipo
        Else
            tProveedor.Text = ""
        End If
        Set aLista = Nothing
        Screen.MousePointer = 0
    End If
    Exit Sub
    Screen.MousePointer = 0

errBuscar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al procesar la lista de ayuda.", Err.Description
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub PropiedadesImpresion()

    With vsListado
        .PhysicalPage = True
        .PaperSize = vbPRPSLetter
        .Orientation = orLandscape
        .PreviewMode = pmPrinter
        .PreviewPage = 1
        .Zoom = 100
        .MarginLeft = 450: .MarginRight = 350
        .MarginBottom = 750: .MarginTop = 750
    End With
    
End Sub
