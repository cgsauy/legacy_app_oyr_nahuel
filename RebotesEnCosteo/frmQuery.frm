VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{191D08B9-4E92-4372-BF17-417911F14390}#1.5#0"; "orGridPreview.ocx"
Begin VB.Form frmQuery 
   Caption         =   "Previa de rebotes al costeo"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pbProgreso 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   14
      Top             =   5385
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrid 
      Height          =   2055
      Left            =   60
      TabIndex        =   11
      Top             =   3180
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   3625
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
      BackColorFixed  =   8421376
      ForeColorFixed  =   10551295
      BackColorSel    =   11829830
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   16384
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   4
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
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
   Begin VB.PictureBox pFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   8955
      TabIndex        =   10
      Top             =   360
      Width           =   8955
      Begin VB.ComboBox cboTipoArticulo 
         Height          =   315
         Left            =   5700
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   1200
         Width           =   2895
      End
      Begin VB.ComboBox cboTipoFiltro 
         Height          =   315
         Left            =   1380
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox tItem 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1380
         TabIndex        =   3
         Text            =   "11111111111111111fdgsdsfgdsf fdsgsfdgsdfgdsf"
         Top             =   840
         Width           =   5175
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   315
         Left            =   1380
         TabIndex        =   1
         Top             =   420
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMMM yyyy"
         Format          =   43188227
         CurrentDate     =   37582
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo artículo:"
         Height          =   195
         Left            =   4560
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Filrar:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Simular costeo en busca de rebotes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   3480
      End
      Begin VB.Label lblUltimoCosteo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Último costeo:"
         ForeColor       =   &H00666666&
         Height          =   195
         Left            =   3480
         TabIndex        =   12
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Mes a costear:"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
   End
   Begin MSComctlLib.StatusBar sbHelpLine 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Top             =   5580
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9075
            Key             =   "help"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "grid"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3775
            MinWidth        =   3775
            Key             =   "progress"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilMenu 
      Left            =   6660
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuery.frx":0442
            Key             =   "print"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuery.frx":0554
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuery.frx":086E
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuery.frx":0CC0
            Key             =   "play"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuery.frx":0DD2
            Key             =   "cleanfilter"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tooMenu 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exit"
            Object.ToolTipText     =   "Salir. [Ctrl+X]"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "play"
            Object.ToolTipText     =   "Consultar. [Ctrl+E]"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "stop"
            Object.ToolTipText     =   "Cancelar carga."
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cleanfilter"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Imprimir. [Ctrl+P]"
         EndProperty
      EndProperty
   End
   Begin orGridPreview.GridPreview grPrint 
      Left            =   60
      Top             =   0
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmQuery"
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

Private bQueryCancel As Boolean
Private colArticulos As Collection
Private colVentas As Collection
Private colCompras As Collection

Private Sub cboTipoArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then ActionPlay
End Sub

Private Sub cboTipoFiltro_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cboTipoArticulo.SetFocus
End Sub

Private Sub dtpDesde_Change()
    InitGrid
End Sub

Private Sub dtpDesde_GotFocus()
    Helpline "Mes a simular."
End Sub

Private Sub dtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        tItem.SetFocus
    End If
End Sub

Private Sub dtpDesde_LostFocus()
    Helpline ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> vbCtrlMask Then Exit Sub
    Select Case KeyCode
        Case vbKeyP: ActionPrint
        Case vbKeyE: ActionPlay
        Case vbKeyX: Unload Me
    End Select
End Sub

Private Sub Form_Load()
Dim fHeader As New StdFont
Dim fFooter As New StdFont
    ObtengoSeteoForm Me, WidthIni:=Me.Width, HeightIni:=Me.Height
    With tooMenu
        .ImageList = ilMenu
        .Buttons("exit").Image = "exit"
        .Buttons("play").Image = "play"
        .Buttons("stop").Image = "stop"
        .Buttons("cleanfilter").Image = "cleanfilter"
        .Buttons("print").Image = "print"
    End With
    
    With fHeader
        .Bold = True
        .Name = "Arial"
        .Size = 11
    End With
    With fFooter
        .Bold = True
        .Name = "Tahoma"
        .Size = 10
    End With
    
    With grPrint
        .Caption = "Simular costeo en busca de rebotes"
        .FileName = "Simular costeo en busca de rebotes"
        .Font = Font
        Set .HeaderFont = fHeader
        .Header = "Simular costeo en busca de rebotes"
        .Orientation = opPortrait
        .PaperSize = 1
        .PageBorder = opTopBottom
    End With
    
    InitGrid
    
    With vsGrid
        .BackColorFixed = &H4000&    '&H787800
        '.BackColorAlternate = &HE0E0C4
    End With
    SetButton True
    
    ActionCleanFilter
    dtpDesde.Value = Date
    
    Set fHeader = Nothing
    Set fFooter = Nothing
    
    CargoArticulos
    CargarUltimoCosteo
    
    
    CargoCombo "SELECT TipCodigo, TipNombre FROM Tipo Order by TipNombre", cboTipoArticulo
    
    cboTipoFiltro.Clear
    cboTipoFiltro.AddItem "De plaza"
    cboTipoFiltro.AddItem "Importados"
    cboTipoFiltro.AddItem "Repuestos"
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    pFilter.Move ScaleLeft, tooMenu.Height + 30, ScaleWidth
    vsGrid.Move ScaleLeft, pFilter.Top + pFilter.Height, ScaleWidth, ScaleHeight - (pFilter.Top + pFilter.Height + sbHelpLine.Height + pbProgreso.Height)
    'pbProgreso.Left =
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    CierroConexion
    Set clsGeneral = Nothing
    GuardoSeteoForm Me
End Sub

Private Sub Label1_Click()
    tItem.SetFocus
End Sub

Private Sub tItem_Change()
    tItem.Tag = ""
    InitGrid
End Sub

Private Sub tItem_GotFocus()
    With tItem
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Helpline "Ingrese un artículo."
End Sub

Private Sub tItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Val(tItem.Tag) > 0 Or Trim(tItem.Text) = "" Then
            cboTipoFiltro.SetFocus
        Else
            bd_BuscoArticulo
        End If
    End If
End Sub

Private Sub tItem_LostFocus()
    Helpline ""
End Sub

Private Sub tooMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "exit": Unload Me
        Case "play": ActionPlay
        Case "cleanfilter": ActionCleanFilter
        Case "print": ActionPrint
        Case "stop": ActionStop
    End Select
End Sub

Private Sub ActionStop()
    bQueryCancel = True
End Sub

Private Sub ActionPlay()
On Error GoTo errAP
Dim lAux As Long
    
    If PrimerDia(dtpDesde.Value) < UltimoDia(lblUltimoCosteo.Tag) Then
        MsgBox "Debe ingresar un mes superior al último costeo.", vbExclamation, "ATENCIÓN"
        dtpDesde.SetFocus
        Exit Sub
    End If
    
    'Screen.MousePointer = 11
    InitGrid
    bQueryCancel = False
    SetButton False
    DoEvents
    Helpline "Consultando ..."
    HelplineGrid ""
    bd_CargoTablaCMCompra
    DoEvents
    If bQueryCancel Then GoTo evSalir
    bd_CargoTablaCMVenta
    DoEvents
    If bQueryCancel Then GoTo evSalir
    CargoTablaCMCosteo
    DoEvents
    
evSalir:
    Helpline ""
    pbProgreso.Value = 0
    SetButton True
    Screen.MousePointer = 0
    Exit Sub
errAP:
    vsGrid.Redraw = True
    clsGeneral.OcurrioError "Ocurrió el siguiente error al consultar.", Err.Description
    Screen.MousePointer = 0
    SetButton True
    HelplineGrid ""
    pbProgreso.Value = 0
End Sub

Private Sub ActionCleanFilter()
On Error Resume Next
    tItem.Text = ""
    tItem.SetFocus
End Sub

Private Sub ActionPrint()
On Error GoTo errPrint
Dim sFilter As String
    
    sFilter = ""
    If Val(tItem.Tag) > 0 Then sFilter = "Artículo: " & Trim(tItem.Text)

    vsGrid.ExtendLastCol = False
    With grPrint
        If sFilter <> "" Then
            .LineBeforeGrid "Filtros", ifontsize:=9, bbold:=True, bitalic:=True
            .LineBeforeGrid sFilter
            .LineBeforeGrid ""
        End If
        .AddGrid vsGrid.hwnd
        .ShowPreview
    End With
    vsGrid.ExtendLastCol = True
    Exit Sub
errPrint:
    clsGeneral.OcurrioError "Ocurrió un error al intentar imprimir.", Err.Description
End Sub

Private Sub SetButton(ByVal bPlay As Boolean)
    
    With tooMenu
        .Buttons("play").Enabled = bPlay
        .Buttons("stop").Enabled = Not bPlay
    End With
    pFilter.Enabled = bPlay
    Me.Refresh
End Sub

Private Sub Helpline(ByVal sText As String)
    sbHelpLine.Panels("help").Text = sText
    sbHelpLine.Refresh
End Sub

Private Sub HelplineGrid(ByVal sText As String)
    sbHelpLine.Panels("grid").Text = sText
    sbHelpLine.Refresh
End Sub

Private Sub InitGrid()
    With vsGrid
        .Cols = 1
        .Rows = 1
        .ExtendLastCol = True
        .FormatString = "Artículo|Mes|Rebote|"
        .ColWidth(1) = 1400: .ColWidth(0) = 3200
        .ColWidth(2) = 1200
        .MergeCol(0) = True
    End With
End Sub

Private Sub vsGrid_GotFocus()
    Helpline "[Botón derecho] Opciones"
End Sub

Private Sub vsGrid_LostFocus()
    Helpline ""
End Sub

Private Sub vsGrid_RowColChange()
    HelplineGrid "Lín:" & vsGrid.Row & " Col:" & vsGrid.Col + 1
End Sub

Private Sub bd_CargoTablaCMVenta()

    Dim fBetween As String
    fBetween = " Between '" & Format(CDate(lblUltimoCosteo.Tag) + 1, sqlFormatoFH) & "' And '" & Format(UltimoDia(dtpDesde.Value) & " 23:59", sqlFormatoFH) & "'"
    Dim strDocumentos As String
    strDocumentos = TipoDocumento.Contado & ", " & TipoDocumento.Credito _
        & ", " & TipoDocumento.NotaCredito & ", " & TipoDocumento.NotaDevolucion & ", " & TipoDocumento.NotaEspecial
    
    
    Dim sTipoArticulo As String
    If cboTipoArticulo.ListIndex > -1 Then
        sTipoArticulo = " AND ArtTipo IN (SELECT TipID FROM dbo.InTipos(" & cboTipoArticulo.ItemData(cboTipoArticulo.ListIndex) & "))"
    End If
    
    Dim sTipoFiltro As String
    If cboTipoFiltro.ListIndex > -1 Then
        Select Case cboTipoFiltro.ListIndex
            Case 1
                sTipoFiltro = " AND ArtSeImporta = 1 AND ArtID NOT IN(SELECT AGrArticulo FROM cgsa.dbo.ArticuloGrupo where AGrGrupo = 14)"
            Case 2
                sTipoFiltro = " AND ArtID IN(SELECT AGrArticulo FROM cgsa.dbo.ArticuloGrupo where AGrGrupo = 14)"
            Case 0
                sTipoFiltro = " AND ArtSeImporta = 0 AND ArtID NOT IN(SELECT AGrArticulo FROM cgsa.dbo.ArticuloGrupo where AGrGrupo = 14)"
        End Select
    End If
    
    Dim sQyArticulo As String
    If Val(tItem.Tag) > 0 Then
        sQyArticulo = " AND ArtID = " & Val(tItem.Tag)
    End If
    
    Helpline "Cargando VENTAS ..."
    Dim totalReg As Long
    pbProgreso.Value = 0
    Cons = "SELECT COUNT(*) FROM CMVenta INNER JOIN Articulo ON ArtID = VenArticulo " & sQyArticulo & sTipoFiltro & sTipoArticulo
    
    Cons = Cons & " UNION ALL Select COUNT(*) FROM Compra, CompraRenglon " _
          & " INNER JOIN Articulo ON ArtID = CReArticulo " & sQyArticulo & sTipoFiltro & sTipoArticulo _
          & " WHERE ComCodigo = CReCompra" _
          & " AND ComFecha " & fBetween _
          & " AND ComTipoDocumento In (" & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraNotaDevolucion & ")" '_
          '& " AND (CReArticulo = " & Val(tItem.Tag) & " OR 0 = " & Val(tItem.Tag) & ")"
    
    Cons = Cons & " UNION ALL Select COUNT(*) FROM Servicio, ServicioRenglon" & _
               " INNER JOIN Articulo ON ArtID = SReMotivo " & sQyArticulo & sTipoFiltro & sTipoArticulo & _
               " Where SerCodigo = SReServicio " & _
               " And SerEstadoServicio = 5 " & _
               " And SerDocumento Is Null " & _
               " And SerFCumplido " & fBetween & _
               " And SReTipoRenglon = 2 " '& _
               '" AND (SReMotivo = " & Val(tItem.Tag) & " OR 0 = " & Val(tItem.Tag) & ")"
    
    Cons = Cons & " UNION ALL Select COUNT(*) " _
        & " From Documento, Renglon" _
        & " INNER JOIN Articulo ON ArtID = RenArticulo " & sQyArticulo & sTipoFiltro & sTipoArticulo _
        & " Where DocTipo IN (" & strDocumentos & ")" _
        & " And DocFecha " & fBetween _
        & " And DocAnulado = 0 And DocCodigo = RenDocumento " '_
        '& " AND (RenArticulo = " & Val(tItem.Tag) & " OR 0 = " & Val(tItem.Tag) & ")"
               
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsAux.EOF
        totalReg = totalReg + RsAux(0)
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Set colVentas = New Collection
    If totalReg = 0 Then Exit Sub
    pbProgreso.Max = totalReg
    pbProgreso.Value = 0
    
    DoEvents
    If bQueryCancel Then Exit Sub
    
    Dim oVenta As clsCMVenta
    Cons = "SELECT * FROM CMVenta INNER JOIN Articulo ON ArtID = VenArticulo " & sQyArticulo & sTipoFiltro & sTipoArticulo
    Cons = Cons & " ORDER BY VenFecha, VenArticulo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsAux.EOF
        Set oVenta = New clsCMVenta
        colVentas.Add oVenta
        With oVenta
            .VenArticulo = RsAux("VenArticulo")
            .VenCantidad = RsAux("VenCantidad")
            .VenCodigo = RsAux("VenCodigo")
            .VenFecha = RsAux("VenFecha")
            .VenPrecio = RsAux("VenPrecio")
            .VenTipo = RsAux("VenTipo")
        End With
        RsAux.MoveNext
        pbProgreso.Value = pbProgreso.Value + 1
    Loop
    RsAux.Close
    
    DoEvents
    If bQueryCancel Then Exit Sub
    
    Dim oCMAux As clsCMVenta
    Dim index As Long
    
    'Segundo paso Cargo Notas de Compras.
    '& " AND (CReArticulo = " & Val(tItem.Tag) & " OR 0 = " & Val(tItem.Tag) & ")"
    Cons = "Select CReArticulo, CReCantidad, CReCompra, ComFecha FROM Compra, CompraRenglon" _
          & " INNER JOIN Articulo ON ArtID = CReArticulo " & sQyArticulo & sTipoFiltro & sTipoArticulo _
          & " Where ComCodigo = CReCompra" _
          & " And ComFecha " & fBetween _
          & " And ComTipoDocumento In (" & TipoDocumento.CompraNotaCredito & ", " & TipoDocumento.CompraNotaDevolucion & ")" _
          & " ORDER BY ComFecha"
          
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)

    Do While Not RsAux.EOF
    
        index = 0
        For I = colVentas.Count To 1 Step -1
            Set oCMAux = colVentas.Item(I)
            If oCMAux.VenFecha > RsAux!ComFecha Then
                index = I
            Else
                Exit For
            End If
        Next
        
        Set oVenta = New clsCMVenta
        If colVentas.Count = 0 Or index = 0 Then
            colVentas.Add oVenta
        Else
            colVentas.Add oVenta, , index
        End If
        With oVenta
            .VenArticulo = RsAux("CReArticulo")
            .VenCantidad = RsAux!CReCantidad
            .VenCodigo = RsAux!CReCompra
            .VenFecha = RsAux!ComFecha
            '.VenPrecio = RsAux("VenPrecio")
            .VenTipo = TipoCV.Compra
        End With
        RsAux.MoveNext
        pbProgreso.Value = pbProgreso.Value + 1
    Loop
    RsAux.Close
    
    DoEvents
    If bQueryCancel Then Exit Sub
    
    '21/5/2001 - Cargo los servicios con costo que no fueron facturados y estan cumplidos--------------------------------------------------------
    '               Como no fueron facturados, los servicios van a entrar con costo de venta 0
    '" AND (SReMotivo = " & Val(tItem.Tag) & " OR 0 = " & Val(tItem.Tag) & ")"
    Cons = "Select SReMotivo, SReCantidad, SerCodigo, SerFCumplido From Servicio, ServicioRenglon" & _
               " INNER JOIN Articulo ON ArtID = SReMotivo " & sQyArticulo & sTipoFiltro & sTipoArticulo & _
               " Where SerCodigo = SReServicio " & _
               " And SerEstadoServicio = 5 " & _
               " And SerDocumento Is Null " & _
               " And SerFCumplido " & fBetween & _
               " And SReTipoRenglon = 2 " & _
               " ORDER BY SerFCumplido"
                   
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        
        index = 0
        For I = colVentas.Count To 1 Step -1
            Set oCMAux = colVentas.Item(I)
            If oCMAux.VenFecha > RsAux!SerFCumplido Then
                index = I
            Else
                Exit For
            End If
        Next
        
        Set oVenta = New clsCMVenta
        If colVentas.Count = 0 Or index = 0 Then
            colVentas.Add oVenta
        Else
            colVentas.Add oVenta, , index
        End If
        With oVenta
            .VenArticulo = RsAux("SReMotivo")
            .VenCantidad = RsAux!SReCantidad
            .VenCodigo = RsAux!SerCodigo
            .VenFecha = RsAux!SerFCumplido
            .VenTipo = TipoCV.Servicio
        End With
        RsAux.MoveNext
        pbProgreso.Value = pbProgreso.Value + 1
    Loop
    RsAux.Close
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
       
    
    'Primer Paso Copio las Ventas---------------------------------------
    'Traigo los documentos Ctdo y Cred, Nota Esp, Nota de Cred. y  Nota de Dev. que no estén anulados
    Cons = "Select DocFecha, DocMoneda, DocTipo, RenArticulo, RenCantidad, RenDocumento " _
        & " From Documento, Renglon" _
        & " INNER JOIN Articulo ON ArtID = RenArticulo " & sQyArticulo & sTipoFiltro & sTipoArticulo _
        & " Where DocTipo IN (" & strDocumentos & ")" _
        & " And DocFecha " & fBetween _
        & " And DocAnulado = 0 And DocCodigo = RenDocumento " _
        & " AND (RenArticulo = " & Val(tItem.Tag) & " OR 0 = " & Val(tItem.Tag) & ")" _
        & " ORDER BY DocFecha"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        
        'Posiciono ordenado por fecha.
        index = 0
        For I = colVentas.Count To 1 Step -1
            Set oCMAux = colVentas.Item(I)
            If oCMAux.VenFecha > RsAux!DocFecha Then
                index = I
            Else
                DoEvents
                If bQueryCancel Then Exit Sub
                Exit For
            End If
        Next
        
        Set oVenta = New clsCMVenta
        If colVentas.Count = 0 Or index = 0 Then
            colVentas.Add oVenta
        Else
            colVentas.Add oVenta, , index
        End If
        With oVenta
            .VenArticulo = RsAux("RenArticulo")
            Select Case RsAux!DocTipo
                Case TipoDocumento.NotaCredito, TipoDocumento.NotaDevolucion, TipoDocumento.NotaEspecial
                    .VenCantidad = RsAux!RenCantidad * -1
                                
                Case Else: .VenCantidad = RsAux!RenCantidad
            End Select
            .VenCodigo = RsAux!RenDocumento
            .VenFecha = RsAux!DocFecha
            '.VenPrecio = RsAux("VenPrecio")
            .VenTipo = TipoCV.Comercio
        End With
        '-------------------------------------------------------------------------------------------------
        RsAux.MoveNext
        pbProgreso.Value = pbProgreso.Value + 1
        Me.Refresh
    Loop
    RsAux.Close
    pbProgreso.Value = 0
    
End Sub

Private Sub bd_CargoTablaCMCompra()

    '1) Cargar las compras del mes (Credito y Contado)
    '2) Cargar las importaciones del mes (con fecha de arribo del costeo en el mes)
    
    'ATENCIÓN: solo van los articulos que tengan el campo ArtAMercaderia = True !!!!!!!
    
    pbProgreso.Value = 0
    Dim totalReg As Long
    
Dim oCompra As clsCMCompra
    Helpline "Cargando COMPRAS ..."
    Dim fBetween As String
    fBetween = " Between '" & Format(CDate(lblUltimoCosteo.Tag) + 1, sqlFormatoFH) & "' And '" & Format(UltimoDia(dtpDesde.Value) & " 23:59", sqlFormatoFH) & "'"
    
    Dim sQyArticulo As String
    If Val(tItem.Tag) > 0 Then
        sQyArticulo = " AND ArtID = " & Val(tItem.Tag)
    End If
    
    Dim sTipoArticulo As String
    If cboTipoArticulo.ListIndex > -1 Then
        sTipoArticulo = " AND ArtTipo IN (SELECT TipID FROM dbo.InTipos(" & cboTipoArticulo.ItemData(cboTipoArticulo.ListIndex) & "))"
    End If
    
    Dim sTipoFiltro As String
    If cboTipoFiltro.ListIndex > -1 Then
        Select Case cboTipoFiltro.ListIndex
            Case 1
                sTipoFiltro = " AND ArtSeImporta = 1 AND ArtID NOT IN(SELECT AGrArticulo FROM cgsa.dbo.ArticuloGrupo where AGrGrupo = 14)"
            Case 2
                sTipoFiltro = " AND ArtID IN(SELECT AGrArticulo FROM cgsa.dbo.ArticuloGrupo where AGrGrupo = 14)"
            Case 0
                sTipoFiltro = " AND ArtSeImporta = 0 AND ArtID NOT IN(SELECT AGrArticulo FROM cgsa.dbo.ArticuloGrupo where AGrGrupo = 14)"
        End Select
    End If
    
    
    Cons = "SELECT Count(*) FROM CMCompra INNER JOIN Articulo ON ComArticulo = ArtID " & sQyArticulo & sTipoFiltro & sTipoArticulo
'    If Val(tItem.Tag) > 0 Then
'         Cons = Cons & " WHERE ComArticulo = " & Val(tItem.Tag)
'    End If
    
    Cons = Cons & " UNION ALL SELECT Count(*) " _
          & " FROM Compra, CompraRenglon INNER JOIN Articulo ON CReArticulo = ArtID And ArtAMercaderia = 1 " & sQyArticulo & sTipoFiltro & sTipoArticulo _
          & " Where ComCodigo = CReCompra" _
          & " And ComFecha " & fBetween _
          & " And ComTipoDocumento In (" & TipoDocumento.CompraContado & ", " & TipoDocumento.CompraCredito & ")" _
        & " UNION ALL " _
        & "Select COUNT(*) from CosteoCarpeta, CosteoArticulo INNER JOIN Articulo ON CArIdArticulo = ArtId And ArtAMercaderia = 1 " & sQyArticulo & sTipoFiltro _
        & " Where CCaFArribo " & fBetween & " And CCaID = CArIDCosteo" _
        & " UNION ALL " _
        & " Select Count(*) from CosteoCarpeta " _
           & " Where CCaFArribo " & fBetween
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsAux.EOF
        totalReg = totalReg + RsAux(0)
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    Set colCompras = New Collection
    If totalReg = 0 Then Exit Sub
    
    pbProgreso.Max = totalReg
    pbProgreso.Value = 0
    
    '0 cargo la tabla de compras
    Cons = "SELECT * FROM CMCompra INNER JOIN Articulo ON ComArticulo = ArtID " & sQyArticulo & sTipoFiltro & sTipoArticulo
    Cons = Cons & " ORDER BY ComFecha"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        Set oCompra = New clsCMCompra
        colCompras.Add oCompra
        With oCompra
            .ComArticulo = RsAux!ComArticulo
            .ComCantidad = RsAux!ComCantidad
            .ComCodigo = RsAux!ComCodigo
            .ComFecha = RsAux!ComFecha
            .ComTipo = RsAux!ComTipo
        End With
        RsAux.MoveNext
        pbProgreso.Value = pbProgreso.Value + 1
    Loop
    RsAux.Close
    
    DoEvents
    If bQueryCancel Then Exit Sub
    
    Dim index As Integer
    Dim oCMAux As clsCMCompra
    
    '1) Compras del Mes (contado y credito)------------------------------------------------------------------------------------------------------------------------
    Cons = "SELECT CReArticulo Articulo, CReCantidad Cantidad, ComCodigo ID, ComFecha Fecha, 1 Tipo " _
          & " FROM Compra, CompraRenglon INNER JOIN Articulo ON CReArticulo = ArtID And ArtAMercaderia = 1 " & sQyArticulo & sTipoFiltro & sTipoArticulo _
          & " Where ComCodigo = CReCompra" _
          & " And ComFecha " & fBetween _
          & " And ComTipoDocumento In (" & TipoDocumento.CompraContado & ", " & TipoDocumento.CompraCredito & ")" _
        & " UNION ALL " _
        & "Select CArIdArticulo Articulo, CArCantidad Cantidad, CCaID ID, CCaFArribo Fecha, 3 Tipo FROM CosteoCarpeta, CosteoArticulo INNER JOIN Articulo ON CArIdArticulo = ArtId " & sQyArticulo & sTipoFiltro & sTipoArticulo _
        & " Where CCaFArribo " & fBetween & " And CCaID = CArIDCosteo" _
        & " ORDER BY Fecha"
          
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        index = 0
        For I = colCompras.Count To 1 Step -1
            Set oCMAux = colCompras.Item(I)
            If oCMAux.ComFecha > RsAux!Fecha Then
                index = I
            Else
                Exit For
            End If
        Next
        
        Set oCompra = New clsCMCompra
        If colCompras.Count = 0 Or index = 0 Then
            colCompras.Add oCompra
        Else
            colCompras.Add oCompra, , index
        End If
        With oCompra
            .ComArticulo = RsAux!Articulo
            .ComCantidad = RsAux!Cantidad
            .ComCodigo = RsAux!ID
            .ComFecha = RsAux!Fecha
            .ComTipo = RsAux!Tipo
        End With
        RsAux.MoveNext
        pbProgreso.Value = pbProgreso.Value + 1
    Loop
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    DoEvents
    If bQueryCancel Then Exit Sub
    
    Dim rsCom As rdoResultset
    'Busca los componetes: Articulos que estan an el remito y no en el embarque----------------------------------------------
    Cons = "Select * from CosteoCarpeta " _
           & " Where CCaFArribo " & fBetween
           
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        
        Cons = "Select * from RemitoCompra, " & _
                    " RemitoCompraRenglon INNER JOIN Articulo ON RCRArticulo = ArtID And ArtAMercaderia = 1 " & sQyArticulo & sTipoFiltro & sTipoArticulo & _
                    " Where RCoCodigo = RCRRemito" & _
                    " And RCoTipoFolder = " & RsAux!CCaNivelFolder & _
                    " And RCoIDFolder = " & RsAux!CCaFolder & _
                    " And RCRArticulo Not in (Select AFoArticulo from ArticuloFolder Where AFoTipo = RCoTipoFolder And AFoCodigo = RCoIdFolder)" & _
                    " ORDER BY RCoFecha"
        Set rsCom = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
        
        Do While Not rsCom.EOF
        
            index = 0
            For I = colCompras.Count To 1 Step -1
                Set oCMAux = colCompras.Item(I)
                If oCMAux.ComFecha > RsAux!CCaFArribo Then
                    index = I
                Else
                    Exit For
                End If
            Next
        
            Set oCompra = New clsCMCompra
            If colCompras.Count = 0 Or index = 0 Then
                colCompras.Add oCompra
            Else
                colCompras.Add oCompra, , index
            End If
            With oCompra
                .ComArticulo = RsAux!RCRArticulo
                .ComCantidad = RsAux!RCRCantidad
                .ComCodigo = RsAux!CCaID
                .ComFecha = RsAux!RCoFecha
                .ComTipo = TipoCV.Importacion
            End With
            rsCom.MoveNext
        Loop
        rsCom.Close
        '----------------------------------------------------------------------------------------------------------------------------------------
        RsAux.MoveNext
        pbProgreso.Value = pbProgreso.Value + 1
    Loop
    RsAux.Close
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
End Sub

Private Sub CargarUltimoCosteo()
On Error GoTo errCUC
    Screen.MousePointer = 11
    Cons = "SELECT MAX(CabMesCosteo) FROM CMCabezal"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    lblUltimoCosteo.Caption = "Último costeo: " & Format(RsAux(0), "MMMM yyyy")
    lblUltimoCosteo.Tag = UltimoDia(RsAux(0))
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
errCUC:
    clsGeneral.OcurrioError "Error al buscar la fecha del último costeo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function BuscarCompraArticuloMenorAFecha(ByVal idArticulo As Long, ByVal Fecha As Date) As clsCMCompra
    Dim oCompra As clsCMCompra
    For I = colCompras.Count To 1 Step -1
        Set oCompra = colCompras(I)
        If (oCompra.ComArticulo = idArticulo And oCompra.ComCantidad > 0 And oCompra.ComFecha <= Fecha) Then
            Set BuscarCompraArticuloMenorAFecha = oCompra
            Exit Function
        End If
    Next
    Set BuscarCompraArticuloMenorAFecha = Nothing
End Function

Private Function BuscarCompraArticuloMayorAFecha(ByVal idArticulo As Long, ByVal Fecha As Date) As clsCMCompra
    Dim oCompra As clsCMCompra
    For Each oCompra In colCompras
        If (oCompra.ComArticulo = idArticulo And oCompra.ComCantidad <> 0 And oCompra.ComFecha >= Fecha) Then
            Set BuscarCompraArticuloMayorAFecha = oCompra
            Exit Function
        End If
    Next
    Set BuscarCompraArticuloMayorAFecha = Nothing
End Function

Private Function BuscarCompraArticuloCantidadMes(ByVal idArticulo As Long, ByVal Mes As Date, ByVal borrar As Boolean) As Integer
    BuscarCompraArticuloCantidadMes = 0
    I = 1
    Dim oCompra As clsCMCompra
    Do While I <= colCompras.Count
        Set oCompra = colCompras(I)
        If oCompra.ComFecha > UltimoDia(Mes) Then Exit Function
        
        If (oCompra.ComArticulo = idArticulo And oCompra.ComCantidad <> 0 And oCompra.ComFecha <= UltimoDia(Mes) And oCompra.ComFecha >= PrimerDia(Mes)) Then
            BuscarCompraArticuloCantidadMes = oCompra.ComCantidad + BuscarCompraArticuloCantidadMes
            If borrar Then
                colCompras.Remove I
                I = I - 1
            End If
        End If
        I = I + 1
    Loop
End Function

Private Function BuscarVentaArticuloCantidadMes(ByVal idArticulo As Long, ByVal Fecha As Date, ByVal borrar As Boolean) As Integer
    BuscarVentaArticuloCantidadMes = 0
    I = 1
    Dim oVta As clsCMVenta
    Do While I <= colVentas.Count
        Set oVta = colVentas(I)
        'Por seguridad
        If oVta.VenFecha > UltimoDia(Fecha) Then Exit Function
        
        If (oVta.VenArticulo = idArticulo And oVta.VenCantidad <> 0 And oVta.VenFecha <= UltimoDia(Fecha) And oVta.VenFecha >= PrimerDia(Fecha)) Then
            BuscarVentaArticuloCantidadMes = BuscarVentaArticuloCantidadMes + oVta.VenCantidad
            If borrar Then
                colVentas.Remove I
                I = I - 1
            End If
        End If
        
        I = I + 1
    Loop
End Function

Private Function BuscarVentaArticuloTotalCantidadMenorAFecha(ByVal idArticulo As Long, ByVal Fecha As Date) As Integer
    BuscarVentaArticuloTotalCantidadMenorAFecha = 0
    Dim oVta As clsCMVenta
    For Each oVta In colVentas
        'Por seguridad
        If oVta.VenFecha > UltimoDia(Fecha) Then Exit Function
        If (oVta.VenArticulo = idArticulo And oVta.VenCantidad <> 0 And oVta.VenFecha <= UltimoDia(Fecha)) Then
            BuscarVentaArticuloTotalCantidadMenorAFecha = BuscarVentaArticuloTotalCantidadMenorAFecha + oVta.VenCantidad
        End If
    Next
End Function

Private Function BuscarCompraArticuloTotalCantidadMenorAFecha(ByVal idArticulo As Long, ByVal Fecha As Date) As Integer
    BuscarCompraArticuloTotalCantidadMenorAFecha = 0
    Dim oCompra As clsCMCompra
    For Each oCompra In colCompras
        If (oCompra.ComArticulo = idArticulo And oCompra.ComCantidad <> 0 And oCompra.ComFecha <= Fecha) Then
            BuscarCompraArticuloTotalCantidadMenorAFecha = oCompra.ComCantidad + BuscarCompraArticuloTotalCantidadMenorAFecha
        End If
    Next
End Function

Private Sub CargoTablaCMCosteo()
Dim QyCos As rdoQuery
Dim RsVen As rdoResultset, rsCom As rdoResultset
Dim aFVenta As Date, aArticulo As Long
Dim aQVenta As Long, aQCompra As Long, aQCosteo As Long
Dim aQVentaOriginal As Long
Dim bBorroVenta As Boolean
    
    Dim colMes As New Collection
    Dim oMes As New clsMesCantidad
    
    pbProgreso.Max = colVentas.Count + colCompras.Count
    pbProgreso.Value = 0
    
    Helpline "SIMULANDO COSTEO ..."
    
    Dim oVta As clsCMVenta
    Dim oArt As clsArticulo
    Dim oCompra As clsCMCompra
    
    For Each oVta In colVentas
        aArticulo = oVta.VenFecha
        aQVenta = oVta.VenCantidad
        aQVentaOriginal = aQVenta
        
        Do While aQVenta <> 0
        
            Set oArt = BuscarArticuloCollection(oVta.VenArticulo)
            If (oArt Is Nothing) Then
                MsgBox "NO hay artículo."
            End If

            'Si el artículo es del tipo Servicio lo costeo contra costo 0
            If oArt.Tipo = 151 Then
                aQCosteo = aQVenta
                aQVenta = 0
            Else
        
                'Voy a la maxima fecha de Compra <= a la fecha de venta ------------------------------------
                Set oCompra = BuscarCompraArticuloMenorAFecha(oVta.VenArticulo, oVta.VenFecha)
                If Not oCompra Is Nothing Then               'Hay una FC <= FV
                    
                    If aQVenta > 0 Then                 'VENTA DE MERCADERIA---------------------------------------------------
                        aQCompra = oCompra.ComCantidad
                        If aQVenta > aQCompra Then
                            aQVenta = aQVenta - aQCompra
                            aQCosteo = aQCompra
                        Else
                            aQCosteo = aQVenta
                            aQVenta = 0
                        End If
                        oCompra.ComCantidad = oCompra.ComCantidad - aQCosteo
                    Else        'DEVOLUCION DE MERCADERIA---------------------------------------------------
                                  'La cantidad debe ser siempre menor a la original, sino voy al inmediato anterior (x q voy a sumar 1 sino me paso)
                                  'IRMA: la sumamos igual, no importa si nos pasamos
                        aQCompra = oCompra.ComCantidad
                        aQCosteo = aQVenta      'QVenta es negativa --> devolucion
                        aQVenta = 0
                        oCompra.ComCantidad = oCompra.ComCantidad - aQCosteo
                    End If
                Else                                        'NO Hay una FC <= FV
                    
                    Set oCompra = BuscarCompraArticuloMayorAFecha(oVta.VenArticulo, oVta.VenFecha)
                    If Not oCompra Is Nothing Then  'Hay una FC >= FV
                    
                        If aQVenta > 0 Then                 'VENTA DE MERCADERIA---------------------------------------------------
                            aQCompra = oCompra.ComCantidad
                            If aQVenta > aQCompra Then
                                aQVenta = aQVenta - aQCompra
                                aQCosteo = aQCompra
                            Else
                                aQCosteo = aQVenta
                                aQVenta = 0
                            End If
                            
                            oCompra.ComCantidad = oCompra.ComCantidad - aQCosteo
                        
                        Else        'DEVOLUCION DE MERCADERIA---------------------------------------------------
                                  'La cantidad debe ser siempre menor a la original, sino voy al inmediato siguiente
                                  'Cambiamos, siempre le sumamos  no importa si me paso en la QdeCompra !!!! 22/5/00
                            aQCompra = oCompra.ComCantidad
                            aQCosteo = aQVenta
                            aQVenta = 0
                            oCompra.ComCantidad = oCompra.ComCantidad - aQCosteo
                            
                        End If
                        
                    Else
                        'Si no hay datos queda remanente, Primero updateo con lo que queda remanente en la venta
                        '11 de Mayo de 2000 - 1) Si es una devolucion y queda remanete la costeo contra costo 0 (aQVenta < 0)
                                          ' 2) Registro un suceso en la grilla y borro la Venta para que no quede remanete (aQVenta = 0 And bBorroVenta)
'                        If aQVenta < 0 Then
'                            aQVenta = 0: bBorroVenta = True
'                        Else
                            oVta.VenCantidad = aQVenta
'                        End If
                        Exit Do
                    End If
                End If
            End If
        Loop
        'Si la venta quedó en cero elimino el registro de la venta
        If aQVenta = 0 Or bBorroVenta Then
            oVta.VenCantidad = aQVenta
        End If
        pbProgreso.Value = pbProgreso.Value + 1
    Next
    
    Helpline "Eliminando registros costeados ..."
    pbProgreso.Value = 0
    If (colVentas.Count > 0) Then pbProgreso.Max = colVentas.Count
    
    I = 1
    Do While I <= colVentas.Count
        Set oVta = colVentas(I)
        If (oVta.VenCantidad = 0) Then
            colVentas.Remove I
            I = I - 1
        End If
        I = I + 1
        pbProgreso.Value = pbProgreso.Value + 1
    Loop
    
    'LAS COMPRAS NO LAS DESPLIEGO.
'    I = 1
'    Do While I <= colCompras.Count
'        Set oCompra = colCompras(I)
'        If (oCompra.ComCantidad = 0) Then
'            colCompras.Remove I
'            I = I - 1
'        End If
'        I = I + 1
'        pbProgreso.Value = pbProgreso.Value + 1
'    Loop
        
    Dim bYaInicie As Boolean
    Dim total As Integer
    
    pbProgreso.Value = 0
    If (colVentas.Count > 0) Then pbProgreso.Max = colVentas.Count
    Helpline ""
    'Genero las diferencias.
    Dim aFCpa As Date
    Dim oMesF As clsMesCantidad
    Dim colorFila As Long
    colorFila = vbWhite
    'Do While (colCompras.Count > 0 Or colVentas.Count > 0)
    If (colVentas.Count > 0) Then pbProgreso.Max = colVentas.Count
    Do While colVentas.Count > 0
        pbProgreso.Value = 0
        'Tomo el primer artículo y lo proceso.
        Set oVta = colVentas(1)
        Set oArt = BuscarArticuloCollection(oVta.VenArticulo)
        Set colMes = New Collection
        I = 1
        Do While (I <= colVentas.Count)
            
            Set oVta = colVentas(I)
            If (oVta.VenArticulo = oArt.ID) Then
                aFVenta = Format(CDate(PrimerDia(oVta.VenFecha)), "dd/mm/yyyy")
                
                Set oMes = New clsMesCantidad
                oMes.Mes = "1/1/1901"
                
                For Each oMesF In colMes
                    If (oMesF.Mes = aFVenta) Then
                        Set oMes = oMesF
                        Exit For
                    End If
                Next
                
                If oMes.Mes <> aFVenta Then
                    Set oMes = New clsMesCantidad
                    oMes.Mes = aFVenta
                    colMes.Add oMes
                End If
                oMes.Cantidad = oMes.Cantidad + oVta.VenCantidad
                
                colVentas.Remove I
                I = I - 1
                
            End If
            I = I + 1
            pbProgreso.Value = pbProgreso.Value + 1
            
        Loop
        
        For Each oMesF In colMes
            With vsGrid
                .AddItem "(" & Format(oArt.Codigo, "#,#00,000") & ") " & oArt.Nombre
                .Cell(flexcpText, .Rows - 1, 1) = Format(oMesF.Mes, "MMMM yyyy")
                .Cell(flexcpText, .Rows - 1, 2) = oMesF.Cantidad
                .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = colorFila
            End With
        Next
        
        If colorFila <> vbWhite Then colorFila = vbWhite Else colorFila = &HC0C0C0
    Loop
    
    
End Sub

'Private Sub CargoTablaCMCosteoOLD()
'
'Dim QyCos As rdoQuery
'Dim RsVen As rdoResultset, rsCom As rdoResultset
'
'Dim aFVenta As Date, aArticulo As Long
'Dim aQVenta As Long, aQCompra As Long, aQCosteo As Long
'Dim aQVentaOriginal As Long
'Dim bBorroVenta As Boolean
'
'    Dim colMes As New Collection
'    Dim oMes As New clsMesCantidad
'    colMes.Add oMes
'    oMes.Mes = PrimerDia(lblUltimoCosteo.Tag)
'    oMes.Compras = BuscarCompraArticuloTotalCantidadMenorAFecha(Val(tItem.Tag), UltimoDia(oMes.Mes))
'    oMes.Ventas = BuscarVentaArticuloTotalCantidadMenorAFecha(Val(tItem.Tag), UltimoDia(oMes.Mes))
'    oMes.Cantidad = oMes.Compras - oMes.Ventas
'
''    With vsGrid
''        .AddItem Format(lblUltimoCosteo.Tag, "MMMM yyyy")
''        .Cell(flexcpText, .Rows - 1, 2) = BuscarCompraArticuloTotalCantidadMenorAFecha(Val(tItem.Tag), UltimoDia(lblUltimoCosteo.Tag))
''        .Cell(flexcpText, .Rows - 1, 3) = BuscarVentaArticuloTotalCantidadMenorAFecha(Val(tItem.Tag), UltimoDia(lblUltimoCosteo.Tag))
''        .Cell(flexcpText, .Rows - 1, 4) = Val(.Cell(flexcpText, .Rows - 1, 2)) - Val(.Cell(flexcpText, .Rows - 1, 3))
''    End With
'
'    aFVenta = CDate("1/2/1900")
'
'    If colVentas.Count > 0 Then
'        aFVenta = colVentas(colVentas.Count).VenFecha
'    End If
'    If colCompras.Count > 0 Then
'        If CDate(aFVenta) < colCompras(colCompras.Count).ComFecha Then
'            aFVenta = colCompras(colCompras.Count).ComFecha
'        End If
'    End If
'
'    If aFVenta <> CDate("1/2/1900") Then
'
'        Dim dIr As Date
'        dIr = CDate(lblUltimoCosteo.Tag) + 1
'        Do While dIr < aFVenta
'
'            Set oMes = New clsMesCantidad
'            colMes.Add oMes
'            oMes.Mes = PrimerDia(dIr)
'            oMes.Compras = BuscarCompraArticuloCantidadMes(Val(tItem.Tag), UltimoDia(oMes.Mes), False)
'            oMes.Ventas = BuscarVentaArticuloCantidadMes(Val(tItem.Tag), UltimoDia(oMes.Mes), False)
'            dIr = DateAdd("M", 1, dIr)
'        Loop
'    End If
'
'    Dim oVta As clsCMVenta
'    Dim oArt As clsArticulo
'    Dim oCompra As clsCMCompra
'    For Each oVta In colVentas
'        aArticulo = oVta.VenFecha
'        aQVenta = oVta.VenCantidad
'        aQVentaOriginal = aQVenta
'
'        Do While aQVenta <> 0
'
'            Set oArt = BuscarArticuloCollection(oVta.VenArticulo)
'
'            'Si el artículo es del tipo Servicio lo costeo contra costo 0
'            If oArt.Tipo = 151 Then
'                aQCosteo = aQVenta
'                aQVenta = 0
'            Else
'
'                'Voy a la maxima fecha de Compra <= a la fecha de venta ------------------------------------
'                Set oCompra = BuscarCompraArticuloMenorAFecha(oVta.VenArticulo, oVta.VenFecha)
'                If Not oCompra Is Nothing Then               'Hay una FC <= FV
'
'                    If aQVenta > 0 Then                 'VENTA DE MERCADERIA---------------------------------------------------
'                        aQCompra = oCompra.ComCantidad
'                        If aQVenta > aQCompra Then
'                            aQVenta = aQVenta - aQCompra
'                            aQCosteo = aQCompra
'                        Else
'                            aQCosteo = aQVenta
'                            aQVenta = 0
'                        End If
'                        oCompra.ComCantidad = oCompra.ComCantidad - aQCosteo
'                    Else        'DEVOLUCION DE MERCADERIA---------------------------------------------------
'                                  'La cantidad debe ser siempre menor a la original, sino voy al inmediato anterior (x q voy a sumar 1 sino me paso)
'                                  'IRMA: la sumamos igual, no importa si nos pasamos
'                        aQCompra = oCompra.ComCantidad
'                        aQCosteo = aQVenta      'QVenta es negativa --> devolucion
'                        aQVenta = 0
'                        oCompra.ComCantidad = oCompra.ComCantidad - aQCosteo
'                    End If
'                Else                                        'NO Hay una FC <= FV
'
'                    Set oCompra = BuscarCompraArticuloMayorAFecha(oVta.VenArticulo, oVta.VenFecha)
'                    If Not oCompra Is Nothing Then  'Hay una FC >= FV
'
'                        If aQVenta > 0 Then                 'VENTA DE MERCADERIA---------------------------------------------------
'                            aQCompra = oCompra.ComCantidad
'                            If aQVenta > aQCompra Then
'                                aQVenta = aQVenta - aQCompra
'                                aQCosteo = aQCompra
'                            Else
'                                aQCosteo = aQVenta
'                                aQVenta = 0
'                            End If
'
'                            oCompra.ComCantidad = oCompra.ComCantidad - aQCosteo
'
'                        Else        'DEVOLUCION DE MERCADERIA---------------------------------------------------
'                                  'La cantidad debe ser siempre menor a la original, sino voy al inmediato siguiente
'                                  'Cambiamos, siempre le sumamos  no importa si me paso en la QdeCompra !!!! 22/5/00
'                            aQCompra = rsCom!ComCantidad
'                            aQCosteo = aQVenta
'                            aQVenta = 0
'                            oCompra.ComCantidad = oCompra.ComCantidad - aQCosteo
'
'                        End If
'
'                    Else
'                        'Si no hay datos queda remanente, Primero updateo con lo que queda remanente en la venta
'                        '11 de Mayo de 2000 - 1) Si es una devolucion y queda remanete la costeo contra costo 0 (aQVenta < 0)
'                                          ' 2) Registro un suceso en la grilla y borro la Venta para que no quede remanete (aQVenta = 0 And bBorroVenta)
''                        If aQVenta < 0 Then
''                            aQVenta = 0: bBorroVenta = True
''                        Else
'                            oVta.VenCantidad = aQVenta
''                        End If
'                        Exit Do
'                    End If
'                End If
'            End If
'        Loop
'        'Si la venta quedó en cero elimino el registro de la venta
'        If aQVenta = 0 Or bBorroVenta Then
'            oVta.VenCantidad = aQVenta
'        End If
'    Next
'
'    I = 1
'    Do While I <= colVentas.Count
'        Set oVta = colVentas(I)
'        If (oVta.VenCantidad = 0) Then
'            colVentas.Remove I
'            I = I - 1
'        End If
'        I = I + 1
'    Loop
'
'    I = 1
'    Do While I <= colCompras.Count
'        Set oCompra = colCompras(I)
'        If (oCompra.ComCantidad = 0) Then
'            colCompras.Remove I
'            I = I - 1
'        End If
'        I = I + 1
'    Loop
'
'    Dim bYaInicie As Boolean
'
'    Dim total As Integer
'
'    'Genero las diferencias.
'    Dim aFCpa As Date
'    'Dim colMes As New Collection
'    'Dim oMes As clsMesCantidad
'    Do While (colCompras.Count > 0 Or colVentas.Count > 0)
'
'        'Agrupo por mes.
'        If (colCompras.Count > 0) Then
'            Set oCompra = colCompras.Item(1)
'            aFCpa = oCompra.ComFecha
'        Else
'            aFCpa = "1/1/2100"
'        End If
'
'        If colVentas.Count > 0 Then
'            Set oVta = colVentas.Item(1)
'            aFVenta = oVta.VenFecha
'        Else
'            aFVenta = "1/1/2100"
'        End If
'
'        If (aFCpa < aFVenta) Then
'            aFVenta = aFCpa
'        End If
'
'        If aFVenta > CDate(lblUltimoCosteo.Tag) Then
'            For Each oMes In colMes
'                If (Format(oMes.Mes, "mm/yyyy") = Format(aFVenta, "mm/yyyy")) Then
'                    oMes.Cantidad = BuscarCompraArticuloCantidadMes(tItem.Tag, oMes.Mes, True) - BuscarVentaArticuloCantidadMes(tItem.Tag, oMes.Mes, True)
'                    total = total + oMes.Cantidad
'                    Exit For
'                End If
'            Next
'        Else
'            total = total + (BuscarCompraArticuloCantidadMes(tItem.Tag, aFVenta, True) - BuscarVentaArticuloCantidadMes(tItem.Tag, aFVenta, True))
'        End If
''        Set oMes = New clsMesCantidad
''        colMes.Add oMes
''        oMes.mes = aFVenta
''        oMes.Cantidad = BuscarCompraArticuloCantidadMes(tItem.Tag, oMes.mes) - BuscarVentaArticuloCantidadMes(tItem.Tag, oMes.mes)
'    Loop
'    For Each oMes In colMes
'        With vsGrid
'            .AddItem Format(oMes.Mes, "MMMM yyyy")
'            .Cell(flexcpText, .Rows - 1, 2) = oMes.Compras
'            .Cell(flexcpText, .Rows - 1, 3) = oMes.Ventas
'            .Cell(flexcpText, .Rows - 1, 4) = oMes.Cantidad
'        End With
'    Next
'
'    With vsGrid
'        .AddItem "Resultado"
'        If (total < 0) Then
'            .Cell(flexcpText, .Rows - 1, 3) = total
'        Else
'            .Cell(flexcpText, .Rows - 1, 2) = total
'        End If
'        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = &HC0FFFF
'        .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
'    End With
'
'
'
'End Sub


Private Sub CargoArticulos()
On Error GoTo errCUC
    Screen.MousePointer = 11
    Dim oArt As clsArticulo
    Set colArticulos = New Collection
    Cons = "SELECT ArtID, ArtCodigo, ArtNombre, ArtTipo FROM Articulo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    Do While Not RsAux.EOF
        Set oArt = New clsArticulo
        colArticulos.Add oArt
        With oArt
            .Codigo = RsAux("ArtCodigo")
            .ID = RsAux("ArtID")
            .Nombre = Trim(RsAux("ArtNombre"))
            .Tipo = RsAux("ArtTipo")
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    Screen.MousePointer = 0
    Exit Sub
errCUC:
    clsGeneral.OcurrioError "Error al buscar la fecha del último costeo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function BuscarArticuloCollection(ByVal idArt As Long) As clsArticulo
    Dim oArt As clsArticulo
    For Each oArt In colArticulos
        If idArt = oArt.ID Then
            Set BuscarArticuloCollection = oArt
            Exit Function
        End If
    Next
    Set BuscarArticuloCollection = Nothing
End Function

Private Sub bd_BuscoArticulo()
'Atención el mapeo de error lo hago antes de entrar al procedimiento
Dim Resultado As Long

    Screen.MousePointer = 11
    If IsNumeric(tItem.Text) Then
        bd_CargoArticuloPorCodigo tItem.Text
    Else
        Cons = "Select ArtId, Código = ArtCodigo, Nombre = ArtNombre from Articulo" _
            & " Where ArtNombre LIKE '" & Replace(tItem.Text, " ", "%") & "%'" _
            & " Order By ArtNombre"
                
        Dim LiAyuda As New clsListadeAyuda
        If LiAyuda.ActivarAyuda(cBase, Cons, , 1, "Buscar artículo") > 0 Then
            Resultado = LiAyuda.RetornoDatoSeleccionado(1)
        Else
            Resultado = 0
        End If
        If Resultado > 0 Then bd_CargoArticuloPorCodigo Resultado
        Set LiAyuda = Nothing       'Destruyo la clase.
    End If
    Screen.MousePointer = 0
    
End Sub

Private Sub bd_CargoArticuloPorCodigo(ByVal CodArticulo As Long)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
    
    Screen.MousePointer = 11
    Cons = "Select * From Articulo Where ArtCodigo = " & CodArticulo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurReadOnly)
    
    If RsAux.EOF Then
        RsAux.Close
        tItem.Tag = ""
        MsgBox "No existe un artículo que posea ese código.", vbExclamation, "ATENCIÓN"
    Else
        tItem.Text = Format(RsAux!ArtCodigo, "#,000,000") & " " & Trim(RsAux!ArtNombre)
        tItem.Tag = RsAux!ArtID
        RsAux.Close
    End If
    Screen.MousePointer = 0
End Sub

