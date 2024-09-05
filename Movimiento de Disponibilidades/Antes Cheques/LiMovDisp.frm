VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.0#0"; "AACOMBO.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmListado 
   Caption         =   "Movimiento de Disponibilidades"
   ClientHeight    =   5820
   ClientLeft      =   1995
   ClientTop       =   2115
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LiMovDisp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   9345
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrilla 
      Height          =   1095
      Left            =   2160
      TabIndex        =   22
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1931
      _ConvInfo       =   1
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
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   4
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
      ExtendLastCol   =   0   'False
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
      Height          =   1095
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   1935
      _Version        =   196608
      _ExtentX        =   3413
      _ExtentY        =   1931
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
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   8655
      Begin VB.TextBox tHasta 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7200
         MaxLength       =   12
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin AACombo99.AACombo cDisponibilidad 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
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
         Text            =   ""
      End
      Begin VB.TextBox tDesde 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5280
         MaxLength       =   12
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   6600
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   4680
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Disponibilidad:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   7995
      TabIndex        =   19
      Top             =   5040
      Width           =   8055
      Begin VB.CheckBox chVista 
         DownPicture     =   "LiMovDisp.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "LiMovDisp.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "LiMovDisp.frx":0ABE
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
         Picture         =   "LiMovDisp.frx":0F38
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
         Picture         =   "LiMovDisp.frx":1022
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
         Picture         =   "LiMovDisp.frx":110C
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
         Picture         =   "LiMovDisp.frx":1346
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
         Picture         =   "LiMovDisp.frx":1448
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Limpiar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5400
         Picture         =   "LiMovDisp.frx":180E
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
         Picture         =   "LiMovDisp.frx":1910
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
         Picture         =   "LiMovDisp.frx":1C12
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
         Picture         =   "LiMovDisp.frx":1F54
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
         Picture         =   "LiMovDisp.frx":2256
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   5940
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
      TabIndex        =   20
      Top             =   5565
      Width           =   9345
      _ExtentX        =   16484
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
            Object.Width           =   8281
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

Private strEncabezado As String, strFormato As String
Private aTexto As String
Private cSaldoInicial As Currency, strFechaSaldo As String, strHoraSaldo As String
Dim gMonedaD As Long

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
    AccionImprimir
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

Private Sub cDisponibilidad_GotFocus()
    With cDisponibilidad
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub
Private Sub cDisponibilidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDesde
End Sub

Private Sub cDisponibilidad_LostFocus()
    cDisponibilidad.SelStart = 0
End Sub

Private Sub chVista_Click()
    
    If chVista.Value = 0 Then
        vsGrilla.ZOrder 0
    Else
        With vsListado          'Selecciono Listado
            Screen.MousePointer = 11
            If vsGrilla.ColHidden(8) Then
                vsGrilla.ColWidth(9) = 3000: .Orientation = orPortrait
            Else
                vsGrilla.ColWidth(9) = 4500: .Orientation = orLandscape
            End If
            vsGrilla.ExtendLastCol = False
            
            .StartDoc
            EncabezadoListado vsListado, "Movimiento de Disponibilidades desde " & tDesde.Text & " hasta " & tHasta.Text, True
            .RenderControl = vsGrilla.hWnd
            .EndDoc
            Screen.MousePointer = 0
        End With
        
        vsListado.ZOrder 0: vsGrilla.ExtendLastCol = True: vsGrilla.ColWidth(9) = 3000
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
            Case vbKeyI: AccionImprimir
            Case vbKeyC: AccionConfigurar
            
            Case vbKeyL: If chVista.Value = 0 Then chVista.Value = 1 Else chVista.Value = 0
            Case vbKeyX: Unload Me
        End Select
    End If
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad

    ObtengoSeteoForm Me, 1000, 500, 3840, 4230
    picBotones.BorderStyle = vbBSNone
    
    PropiedadesImpresion
    LimpioGrilla
    
    'Cargo disponibilidades.-------------------------------
    Cons = "Select DisID, DisNombre From NivelPermiso, Disponibilidad " _
        & " Where NPeNivel IN (Select UNiNivel From UsuarioNivel Where UNiUsuario = " & paCodigoDeUsuario & ")" _
        & " And NPeAplicacion = DisAplicacion" _
        & " Group by DisID, DisNombre" _
        & " Order by DisNombre"
    CargoCombo Cons, cDisponibilidad
    BuscoCodigoEnCombo cDisponibilidad, paDisponibilidad
    '--------------------------------------------------------------
    tDesde.Text = Format(Now, FormatoFP)
    tHasta.Text = Format(Now, FormatoFP)
    
    vsGrilla.ColHidden(7) = True: vsGrilla.ColHidden(8) = True
    pbProgreso.Value = 0
    Exit Sub
    
ErrLoad:
    clGeneral.OcurrioError "Ocurrió un error inesperado al cargar el formulario.", Err.Description
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = fFiltros.Width
    
    picBotones.Width = vsListado.Width
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 250
    
    vsGrilla.Width = vsListado.Width
    vsGrilla.Height = vsListado.Height
    vsGrilla.Left = vsListado.Left
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next

    GuardoSeteoForm Me
    
    CierroConexion
    Set clGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Private Sub Label1_Click()
    Foco tDesde
End Sub

Private Sub AccionImprimir()
    
    On Error GoTo errImprimir
    Me.Refresh
    
    vsGrilla.ExtendLastCol = False
    With vsListado
        .StartDoc
        .filename = "MovDisponibilidades"
        EncabezadoListado vsListado, "Movimiento de Disponibilidades desde " & tDesde.Text & " hasta " & tHasta.Text, True
        .RenderControl = vsGrilla.hWnd
        .EndDoc
    End With
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    Me.Refresh
    If Not frmSetup.pOK Then Exit Sub
    
    vsListado.PrintDoc , frmSetup.pPaginaD, frmSetup.pPaginaH
    vsGrilla.ExtendLastCol = True
    
    Exit Sub
errImprimir:
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al realizar la impresión", Err.Description
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
        .Zoom = 60
    End With

End Sub


Private Sub ArmoEncabezadoTabla()

    strFormato = "+<1500|+<1250|+<3100|+>1300|+>1300|+>1300|+<4600"
    strEncabezado = "Fecha|Cheque|Proveedor|Debe|Haber|Saldo|Concepto"
    
    With vsListado
        .FontSize = 10: .FontBold = True
        .TableBorder = tbBoxRows
        .AddTable strFormato, strEncabezado, "", Inactivo
        .FontSize = 8: .FontBold = False
    End With
    
End Sub
Private Sub tdesde_GotFocus()
    With tDesde: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tHasta
End Sub

Private Sub tDesde_LostFocus()
    tDesde.SelStart = 0
    If IsDate(tDesde.Text) Then tDesde.Text = Format(tDesde.Text, FormatoFP) Else tDesde.Text = ""
End Sub
Private Sub tHasta_GotFocus()
    With tHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub
Private Sub tHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco bConsultar
End Sub
Private Sub tHasta_LostFocus()
    tHasta.SelStart = 0
    If IsDate(tHasta.Text) Then tHasta.Text = Format(tHasta.Text, FormatoFP) Else tHasta.Text = ""
End Sub

Private Sub vsGrilla_Click()

    With vsGrilla
        If .MouseRow = 0 Then
            .ColSel = .MouseCol
            If .ColSort(.MouseCol) = flexSortGenericAscending Then
                .ColSort(.MouseCol) = flexSortGenericDescending
            Else
                .ColSort(.MouseCol) = flexSortGenericAscending
            End If
            .Sort = flexSortUseColSort
        End If
    End With
    
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionConsultar()

    On Error GoTo ErrCDML
    
    If Not ValidoDatos Then Exit Sub
    Screen.MousePointer = 11
    
    vsGrilla.ZOrder 0: vsGrilla.Rows = 1: vsGrilla.Refresh
    chVista.Value = 0
    
    'Busco el si hay un saldo inicial para esa disponibilidad.----------------------------------------------------------------------------------
    cSaldoInicial = 0: strFechaSaldo = "": strHoraSaldo = ""
    
    Cons = "Select * From SaldoDisponibilidad " _
        & " Where SDiDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
        & " And SDiFecha = (Select MAX(SDiFecha) From SaldoDisponibilidad " _
            & " Where SDiDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
            & " And SDiFecha <= '" & Format(tDesde.Text & " " & "23:59:59", sqlFormatoFH) & "')"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then
        cSaldoInicial = RsAux!SDiSaldo
        strFechaSaldo = RsAux!SDiFecha
        strHoraSaldo = RsAux!SDiHora
    Else
        strFechaSaldo = tDesde.Text
        strHoraSaldo = "00:00:00"
    End If
    RsAux.Close
    '----------------------------------------------------------------------------------------------------------------------------------------------
    
    'Busco la moneda de la disponibilidad (para habilitar columnas de pesos)-----------------------------------------------------------
    Cons = "Select * From Disponibilidad Where DisId = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    gMonedaD = RsAux!DisMoneda
    RsAux.Close
    If gMonedaD = paMonedaPesos Then
        vsGrilla.ColHidden(7) = True: vsGrilla.ColHidden(8) = True
    Else
        vsGrilla.ColHidden(7) = False: vsGrilla.ColHidden(8) = False
    End If
    '----------------------------------------------------------------------------------------------------------------------------------------------
    
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    pbProgreso.Value = 0
    Cons = "Select Count(*) from  MovimientoDisponibilidad " _
                & ", MovimientoDisponibilidadRenglon" _
                & " Where MDRIDDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
                & " And ((MDiFecha = '" & Format(strFechaSaldo, sqlFormatoF) & "' And MDiHora >= '" & strHoraSaldo & "')" _
                & " Or  MDiFecha > '" & Format(strFechaSaldo, sqlFormatoF) & "') " _
                & " And MDiFecha <= '" & Format(tHasta.Text, sqlFormatoF) & "'" _
                & " And MDIId = MDRIDMovimiento "
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
    
    Cons = "Select * From  MovimientoDisponibilidad " _
                        & " Left Outer Join Compra On MDiIdCompra = ComCodigo" _
                        & " Left Outer Join ProveedorCliente On ComProveedor = PClCodigo " _
                & ", MovimientoDisponibilidadRenglon left Outer Join Cheque On  MDRIdCheque = CheId " _
                & " Where MDRIDDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
                & " And ((MDiFecha = '" & Format(strFechaSaldo, sqlFormatoF) & "' And MDiHora >= '" & strHoraSaldo & "')" _
                        & " Or  MDiFecha > '" & Format(strFechaSaldo, sqlFormatoF) & "') " _
                & " And MDiFecha <= '" & Format(tHasta.Text, sqlFormatoF) & "'" _
                & " And MDIId = MDRIDMovimiento " _
                & " Order by MDIFecha, MDiHora"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsAux.EOF Then
        vsGrilla.Redraw = False
        CargoDatos
        RsAux.Close
        vsGrilla.Redraw = True
    Else
        RsAux.Close
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbExclamation, "ATENCIÓN"
    End If
    
    Screen.MousePointer = 0
    Exit Sub
ErrCDML:
    clGeneral.OcurrioError "Ocurrió un error al cargar los datos.", Err.Description
    vsGrilla.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatos()
Dim cSaldo As Currency
Dim aTC As Currency, aSaldoPesos As Currency        'Columnas en pesos
Dim Rs1 As rdoResultset

    cSaldo = 0
    
    'Hay que sacar la TC a Pesos si la moneda es distinta
    If cSaldoInicial <> 0 And gMonedaD <> paMonedaPesos Then
        aTC = TasadeCambio(CInt(gMonedaD), paMonedaPesos, CDate(strFechaSaldo))
        aSaldoPesos = cSaldoInicial * aTC
    End If
    
    If cSaldoInicial > 0 Then
        With vsGrilla
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(strFechaSaldo, "dd/mm/yy hh:mm")
            .Cell(flexcpText, .Rows - 1, 4) = Format(cSaldoInicial, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = Format(cSaldoInicial, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 8) = Format(aSaldoPesos, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 9) = "Saldo Inicial"
        End With
    ElseIf cSaldoInicial < 0 Then
        With vsGrilla
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(strFechaSaldo, "dd/mm/yy hh:mm")
            .Cell(flexcpText, .Rows - 1, 5) = Format(cSaldoInicial, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = Format(cSaldoInicial, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 8) = Format(aSaldoPesos, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 9) = "Saldo Inicial"
        End With
    End If
    cSaldo = cSaldoInicial
    With vsGrilla
    
    Do While Not RsAux.EOF
            pbProgreso.Value = pbProgreso.Value + 1
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!MDiFecha, "dd/mm/yy") & " " & Format(RsAux!MDiHora, "hh:mm")
            If Not IsNull(RsAux!CheSerie) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!CheSerie) & " " & Trim(RsAux!CheNumero)
            If Not IsNull(RsAux!CheVencimiento) Then .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!CheVencimiento, "dd/mm/yy")
            
            If Not IsNull(RsAux!PClFantasia) Then
                .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!PClFantasia)
            Else
                'Veo si es transferencia---------------------------------------------------------------------------------------
                Cons = "Select * from MovimientoDisponibilidad, MovimientoDisponibilidadRenglon, TipoMovDisponibilidad, Disponibilidad " _
                                & " Where MDiID = MDRIDMovimiento " _
                                & " And MDRidMovimiento = " & RsAux!MDiID _
                                & " And MDRIdDisponibilidad <> " & RsAux!MDRIdDisponibilidad _
                                & " And MDRIdDisponibilidad = DisID And MDITipo = TMDCodigo"
                    
                Set Rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not Rs1.EOF Then
                    If Not IsNull(Rs1!TMDTransferencia) Then
                        If Rs1!TMDTransferencia = 1 Then
                            If Not Rs1.EOF Then .Cell(flexcpText, .Rows - 1, 3) = Trim(Rs1!DisNombre)
                        End If
                    End If
                End If
                Rs1.Close
                '-------------------------------------------------------------------------------------------------------------------
            End If
            
            If Not IsNull(RsAux!MDRDebe) Then
                .Cell(flexcpText, .Rows - 1, 4) = Format(Abs(RsAux!MDRDebe), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 7) = Format(Abs(RsAux!MDRImportePesos), FormatoMonedaP)
                cSaldo = cSaldo + .Cell(flexcpValue, .Rows - 1, 4)
            End If
            
            If Not IsNull(RsAux!MDRHaber) Then
                .Cell(flexcpText, .Rows - 1, 5) = Format(Abs(RsAux!MDRHaber), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 7) = Format(Abs(RsAux!MDRImportePesos) * -1, FormatoMonedaP)
                cSaldo = cSaldo - .Cell(flexcpValue, .Rows - 1, 5)
            End If
            
            .Cell(flexcpText, .Rows - 1, 6) = Format(cSaldo, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 8) = Format(.Cell(flexcpValue, .Rows - 2, 8) + .Cell(flexcpValue, .Rows - 1, 7), FormatoMonedaP)
            
            If Not IsNull(RsAux!MDiComentario) Then
                .Cell(flexcpText, .Rows - 1, 9) = Trim(RsAux!MDiComentario)
            Else
                If Not IsNull(RsAux!ComComentario) Then .Cell(flexcpText, .Rows - 1, 9) = Trim(RsAux!ComComentario)
            End If
        
        RsAux.MoveNext
    Loop
    
    .AddItem ""
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 4, , Colores.Rojo, Colores.Blanco, True, "Total"
    .Subtotal flexSTSum, -1, 5, , Colores.Rojo, Colores.Blanco, True, "Total"
    .Cell(flexcpText, .Rows - 1, 6) = .Cell(flexcpText, .Rows - 3, 6)
    .Cell(flexcpText, .Rows - 1, 8) = .Cell(flexcpText, .Rows - 3, 8)
    
    End With
    pbProgreso.Value = 0
    
End Sub

Private Sub AccionLimpiar()
'On Error Resume Next
    cDisponibilidad.Text = ""
    LimpioGrilla
    chVista.Value = 0
    With vsListado
        .StartDoc: .EndDoc
    End With
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

Private Sub LimpioGrilla()
    With vsGrilla
        .Redraw = False
        .ExtendLastCol = True
        .Rows = 1: .Cols = 1
        
        .FormatString = "<Fecha|<Cheque|Vence|<Proveedor|>Debe|>Haber|>Saldo|>Debe/Haber $|>Saldo $|<Concepto"
        .ColWidth(0) = 1250: .ColWidth(1) = 1000: .ColWidth(2) = 750: .ColWidth(3) = 1900: .ColWidth(4) = 1200: .ColWidth(5) = 1200:: .ColWidth(6) = 1200
        .ColWidth(7) = 1200: .ColWidth(8) = 1300: .ColWidth(9) = 3500
        .ColDataType(0) = flexDTDate
        .AllowUserResizing = flexResizeColumns
        .Redraw = True
    End With
End Sub


Private Function ValidoDatos() As Boolean

    ValidoDatos = False
    
    If cDisponibilidad.ListIndex = -1 Then
        MsgBox "Seleccione una disponbilidad para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco cDisponibilidad: Exit Function
    End If
    
    If Not IsDate(tDesde.Text) Or Not IsDate(tHasta.Text) Then
        MsgBox "Las fechas ingresadas no son válidas.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    If CDate(tDesde.Text) > CDate(tHasta.Text) Then
        MsgBox "Las fechas ingresadas no son válidas.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    ValidoDatos = True
    
End Function
