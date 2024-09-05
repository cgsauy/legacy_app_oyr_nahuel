VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmPagos 
   Caption         =   "Pagos a Proveedores"
   ClientHeight    =   7620
   ClientLeft      =   1200
   ClientTop       =   1800
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPagos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   10800
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   2055
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   4575
      _Version        =   196608
      _ExtentX        =   8070
      _ExtentY        =   3625
      _StockProps     =   229
      BorderStyle     =   1
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
      Zoom            =   60
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
      ForeColor       =   &H00000080&
      Height          =   660
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   10155
      Begin VB.TextBox tProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1020
         MaxLength       =   40
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox tHasta 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7020
         MaxLength       =   12
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox tDesde 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5280
         MaxLength       =   12
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   6480
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Proveedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   4680
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   7275
      TabIndex        =   18
      Top             =   6240
      Width           =   7335
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmPagos.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmPagos.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmPagos.frx":0ABE
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
         Picture         =   "frmPagos.frx":0F38
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
         Picture         =   "frmPagos.frx":1022
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
         Picture         =   "frmPagos.frx":110C
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "frmPagos.frx":1346
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Picture         =   "frmPagos.frx":1448
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5640
         Picture         =   "frmPagos.frx":180E
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
         Picture         =   "frmPagos.frx":1910
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
         Picture         =   "frmPagos.frx":1C12
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
         Picture         =   "frmPagos.frx":1F54
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
         Picture         =   "frmPagos.frx":2256
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6240
         TabIndex        =   23
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
      TabIndex        =   19
      Top             =   7365
      Width           =   10800
      _ExtentX        =   19050
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
            Object.Width           =   10848
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   3255
      Left            =   1200
      TabIndex        =   21
      Top             =   1920
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5741
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
End
Attribute VB_Name = "frmPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsAux As rdoResultset
Private aTexto As String
Dim aDisponibilidades As String

Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConfigurar_Click()
    AccionConfigurar
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

Private Sub Form_Load()
    
    On Error Resume Next
    ObtengoSeteoForm Me, Me.Left, Me.Top, Me.Width, Me.Height
    frmPagos.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    frmPagos.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    frmPagos.Status.Panels("bd") = "BD: " & PropiedadesConnect(txtConexion, Database:=True) & " "
    
    pbProgreso.Value = 0
    InicializoGrilla
    vsConsulta.ZOrder 0
    
    picBotones.BorderStyle = vbBSNone
    PropiedadesImpresion
    
    FechaDelServidor
    tDesde.Text = Format(PrimerDia(gFechaServidor), "dd/mm/yyyy")
    tHasta.Text = Format(UltimoDia(gFechaServidor), "dd/mm/yyyy")
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = fFiltros.Width
    
    picBotones.Width = vsListado.Width
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 50
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = vsListado.Height
    vsConsulta.Left = vsListado.Left
    
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
    Foco tDesde
End Sub

Private Sub Label2_Click()
    Foco tProveedor
End Sub

Private Sub Label3_Click()
    Foco tHasta
End Sub

Private Sub tDesde_GotFocus()
    With tDesde: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    With vsListado
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    
    EncabezadoListado vsListado, "Pagos a Proveedores -  Desde " & Trim(tDesde.Text) & " al " & Trim(tHasta.Text), False
    vsListado.filename = "Pagos a Proveedores"
    
    vsConsulta.ExtendLastCol = False
    vsListado.RenderControl = vsConsulta.hWnd
    vsConsulta.ExtendLastCol = True
    
    vsListado.EndDoc
    
    If Imprimir Then
        frmSetup.pControl = vsListado
        frmSetup.Show vbModal, Me
        Me.Refresh
        If frmSetup.pOK Then vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errImprimir:
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub PropiedadesImpresion()
  
  With vsListado
        .PaperSize = vbPRPSLetter
        .PhysicalPage = True
        .Orientation = orPortrait
        
        .PreviewMode = pmPrinter
        
        .PreviewPage = 1
        .FontName = "Tahoma": .FontSize = 10: .FontBold = False: .FontItalic = False
        .TextAlign = 0: .PageBorder = 3
        .Columns = 1
        .TableBorder = tbBoxRows
        .Zoom = 100
    End With

End Sub


Private Sub tDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tHasta
End Sub

Private Sub tDesde_LostFocus()
    If IsDate(tDesde.Text) Then tDesde.Text = Format(tDesde.Text, "dd/mm/yyyy")
End Sub

Private Sub tHasta_GotFocus()
    With tHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub tHasta_LostFocus()
    If IsDate(tHasta.Text) Then tHasta.Text = Format(tHasta.Text, "dd/mm/yyyy")
End Sub

Private Sub tProveedor_Change()
    tProveedor.Tag = 0
End Sub

Private Sub tProveedor_GotFocus()
    With tProveedor: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tProveedor_KeyPress(KeyAscii As Integer)
On Error GoTo errBuscar
    
    If KeyAscii = vbKeyReturn Then
        If Val(tProveedor.Tag) <> 0 Or Trim(tProveedor.Text) = "" Then Foco tDesde: Exit Sub
        Screen.MousePointer = 11
        Cons = "Select PClCodigo, PClNombre, PClFantasia from ProveedorCliente " _
                & " Where PClNombre like '" & Trim(tProveedor.Text) & "%' Or PClFantasia like '" & Trim(tProveedor.Text) & "%'" _
                & " Order by PClNombre"
        
        Dim aLista As New clsListadeAyuda
        aLista.ActivoListaAyuda Cons, False, txtConexion, 5500
        If aLista.ValorSeleccionado <> 0 Then
            tProveedor.Text = Trim(aLista.ItemSeleccionado)
            tProveedor.Tag = aLista.ValorSeleccionado
            Foco tDesde
        Else
            tProveedor.Text = ""
        End If
        Set aLista = Nothing
        Screen.MousePointer = 0
    End If
    Exit Sub
    Screen.MousePointer = 0

errBuscar:
    clsGeneral.OcurrioError "Ocurrió un error al procesar la lista de ayuda.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub vsConsulta_DblClick()
    With vsConsulta
        If .Rows = 1 Then Exit Sub
        If .Cell(flexcpText, .Row, 8) = "" Then Exit Sub
        For I = 1 To .Rows - 1
            If .Cell(flexcpValue, I, 0) = .Cell(flexcpValue, .Row, 8) Then
                .Select I, 0, , .Cols - 1
                I = .CellTop
                Exit For
            End If
        Next
    End With
    
End Sub


Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionConsultar()
 
Dim Rs1 As rdoResultset
Dim aPagosP As Currency, aPagosD As Currency

    On Error GoTo ErrCDML
    If Not ValidoDatos Then Exit Sub
    
    Screen.MousePointer = 11
    chVista.Value = vbUnchecked
    aPagosP = 0: aPagosD = 0
    vsConsulta.Rows = 1: vsConsulta.Refresh
    
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    pbProgreso.Value = 0
    Cons = "Select Count(*) from Compra " _
           & " Where ComFecha Between '" & Format(tDesde.Text, sqlFormatoF) & "' And '" & Format(tHasta.Text, sqlFormatoF) & "'" _
           & " And ComTipoDocumento = " & TipoDocumento.CompraReciboDePago
    If Val(tProveedor.Tag) <> 0 Then Cons = Cons & " And ComProveedor = " & Val(tProveedor.Tag)
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux(0) = 0 Then
            MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
            RsAux.Close: Screen.MousePointer = 0: Exit Sub
        End If
        pbProgreso.Max = RsAux(0)
    End If
    RsAux.Close
    '-----------------------------------------------------------------------------------------------------------------

    Cons = "Select * from Compra, ProveedorCliente" _
           & " Where ComFecha Between '" & Format(tDesde.Text, sqlFormatoF) & "' And '" & Format(tHasta.Text, sqlFormatoF) & "'" _
           & " And ComTipoDocumento = " & TipoDocumento.CompraReciboDePago _
           & " And ComProveedor = PClCodigo"
    If Val(tProveedor.Tag) <> 0 Then Cons = Cons & " And ComProveedor = " & Val(tProveedor.Tag)
    
    'Cons = Cons & " Order by ComFecha, ComProveedor"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCIÓN"
        Screen.MousePointer = 0: RsAux.Close: Exit Sub
    End If
    
    With vsConsulta
    Do While Not RsAux.EOF
        pbProgreso.Value = pbProgreso.Value + 1
            
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!ComFecha, "Mmmm yyyy")
        If Not IsNull(RsAux!PClNombre) Then aTexto = Trim(RsAux!PClNombre) Else aTexto = Trim(RsAux!PClFantasia)
        .Cell(flexcpText, .Rows - 1, 1) = aTexto
        
        .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!ComCodigo, "#,##0")
        .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!ComFecha, "dd/mm/yy")
                
        aTexto = RetornoNombreDocumento(RsAux!ComTipoDocumento, Abreviacion:=True) & " "
        If Not IsNull(RsAux!ComSerie) Then aTexto = aTexto & Trim(RsAux!ComSerie) & " "
        If Not IsNull(RsAux!ComNumero) Then aTexto = aTexto & RsAux!ComNumero
        .Cell(flexcpText, .Rows - 1, 4) = aTexto
                
        If RsAux!ComMoneda = paMonedaPesos Then
            .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!ComImporte, FormatoMonedaP)
            If Not IsNull(RsAux!ComIva) Then .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!ComIva, FormatoMonedaP) Else .Cell(flexcpText, .Rows - 1, 6) = "0.00"
            
            .Cell(flexcpText, .Rows - 1, 7) = Format(.Cell(flexcpValue, .Rows - 1, 5) + .Cell(flexcpValue, .Rows - 1, 6), FormatoMonedaP)
                        
        Else
            .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!ComImporte * RsAux!ComTC, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 8) = Format(RsAux!ComImporte, FormatoMonedaP)
            If Not IsNull(RsAux!ComIva) Then
                .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!ComIva * RsAux!ComTC, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 8) = Format(RsAux!ComImporte + RsAux!ComIva, FormatoMonedaP)
            Else
                .Cell(flexcpText, .Rows - 1, 6) = "0.00"
            End If
            .Cell(flexcpText, .Rows - 1, 7) = Format(.Cell(flexcpValue, .Rows - 1, 6) + .Cell(flexcpValue, .Rows - 1, 5), FormatoMonedaP)
                        
        End If

        RsAux.MoveNext
    Loop
    RsAux.Close
   
   .Select 1, 0, 1, 3
   .Sort = flexSortGenericAscending
        
   .SubtotalPosition = flexSTBelow
   .Subtotal flexSTSum, 1, 5, , Colores.Gris, , False, "Total %s"
   .Subtotal flexSTSum, 1, 6: .Subtotal flexSTSum, 1, 7: .Subtotal flexSTSum, 1, 8
   
   .Subtotal flexSTSum, 0, 5, , Colores.Rojo, Colores.Blanco, True, "Total %s"
   .Subtotal flexSTSum, 0, 6: .Subtotal flexSTSum, 0, 7: .Subtotal flexSTSum, 0, 8
    
    End With
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    Exit Sub
    
ErrCDML:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub AccionLimpiar()
    tDesde.Text = ""
    tProveedor.Text = ""
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

Private Function ValidoDatos() As Boolean
    On Error Resume Next
    ValidoDatos = False
            
    If Not IsDate(tDesde.Text) Then
        MsgBox "La fecha desde ingresada para consultar no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    If Not IsDate(tHasta.Text) Then
        MsgBox "La fecha hasta ingresada para consultar no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tHasta: Exit Function
    End If
    If CDate(tDesde.Text) > CDate(tHasta.Text) Then
        MsgBox "El rango de fechas ingresado para consultar no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    ValidoDatos = True
End Function

Private Sub InicializoGrilla()

    On Error Resume Next
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = "Mes|Proveedor|<Fecha|<ID Gasto|<Documento|>Importe $|>I.V.A. $|>Total $|>Total U$S|"
        .ColWidth(0) = 0: .ColWidth(1) = 0
        .ColWidth(2) = 900: .ColWidth(3) = 900: .ColWidth(4) = 1600
        .ColWidth(5) = 1400: .ColWidth(6) = 1100: .ColWidth(7) = 1500: .ColWidth(8) = 1400
        
        .MergeCells = flexMergeSpill
        .WordWrap = False
        .OutlineBar = flexOutlineBarNone
        .ColDataType(0) = flexDTDate
    End With
      
End Sub


