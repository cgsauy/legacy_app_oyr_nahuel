VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMayor 
   Caption         =   "Mayor Analítico de Disponibilidades"
   ClientHeight    =   8490
   ClientLeft      =   1200
   ClientTop       =   1665
   ClientWidth     =   11685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMayor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11685
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
      Zoom            =   60
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
      ForeColor       =   &H00000080&
      Height          =   660
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   10335
      Begin AACombo99.AACombo cDisponibilidad 
         Height          =   315
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSComCtl2.DTPicker tDesde 
         Height          =   315
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   23789569
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker tHasta 
         Height          =   315
         Left            =   2580
         TabIndex        =   3
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   23789569
         CurrentDate     =   37543
      End
      Begin VB.Label Label3 
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Disponibilidad:"
         Height          =   255
         Left            =   4320
         TabIndex        =   4
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   7875
      TabIndex        =   18
      Top             =   6240
      Width           =   7935
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmMayor.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmMayor.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmMayor.frx":0ABE
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
         Picture         =   "frmMayor.frx":0F38
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
         Picture         =   "frmMayor.frx":1022
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
         Picture         =   "frmMayor.frx":110C
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
         Picture         =   "frmMayor.frx":1346
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
         Picture         =   "frmMayor.frx":1448
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5400
         Picture         =   "frmMayor.frx":180E
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
         Picture         =   "frmMayor.frx":1910
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
         Picture         =   "frmMayor.frx":1C12
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
         Picture         =   "frmMayor.frx":1F54
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
         Picture         =   "frmMayor.frx":2256
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6000
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
      Top             =   8235
      Width           =   11685
      _ExtentX        =   20611
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
            Object.Width           =   12409
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   3255
      Left            =   120
      TabIndex        =   21
      Top             =   2820
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
Attribute VB_Name = "frmMayor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAux As rdoResultset
Private aTexto As String
Dim aDisponibilidades As String

'">#Ref|Fecha|<Subrubro|<Documento|>Importe (G) $|>Cofis (G) $|>I.V.A. (G) $|>Debe $|>Haber $|>Importe M/E|<Proveedor|"
Private Enum eCols
    ID
    Fecha
    Subrubro
    Documento
    Importe
    Cofis
    IVA
    Debe
    Haber
    ImporteME
    Proveedor
End Enum

Private Type typDato
    Rubro As String
    ImporteD As Currency
    ImporteH As Currency
End Type
    
Dim arrRubros() As typDato

Dim OBJ_Gastos As typDato       'Asientos de Gastos

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

Private Sub cDisponibilidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then bConsultar.SetFocus
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
    ObtengoSeteoForm Me, 1000, 500, 3840, 4230
    
    InicializoGrilla
    pbProgreso.Value = 0
    picBotones.BorderStyle = vbBSNone
    PropiedadesImpresion
    
    
    tDesde.Value = CDate("01/" & Format(Date, "mm/yyyy"))
    tHasta.Value = DateAdd("d", -1, DateAdd("m", 1, tDesde.Value))
    
    cons = "Select DisId, DisNombre from Disponibilidad Order by DisNombre"
    
    'Cargo disponibilidades.-------------------------------
    cons = "Select DisID, DisNombre From NivelPermiso, Disponibilidad " _
        & " Where NPeNivel IN (Select UNiNivel From UsuarioNivel Where UNiUsuario = " & paCodigoDeUsuario & ")" _
        & " And NPeAplicacion = DisAplicacion" _
        & " Group by DisID, DisNombre " _
        & " Order by DisNombre"
    
    CargoCombo cons, cDisponibilidad
    
    aDisponibilidades = ""
    For I = 0 To cDisponibilidad.ListCount - 1
        aDisponibilidades = aDisponibilidades & cDisponibilidad.ItemData(I) & ","
    Next I
    aDisponibilidades = mID(aDisponibilidades, 1, Len(aDisponibilidades) - 1)
    '--------------------------------------------------------------
    
    CargoConstantesSubrubros ComoRubros:=True
    
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
    
    vsConsulta.ZOrder 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next

    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Private Sub Label2_Click()
    Foco cDisponibilidad
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
    
    zfn_EncabezadoListado vsListado, "Mayor Analítico de Disponibilidades - " & Trim(cDisponibilidad.Text) & " - " & tDesde.Value & " al " & tHasta.Value, True
    vsListado.FileName = "Mayor Analítico"
    
    vsConsulta.ExtendLastCol = False
    vsListado.RenderControl = vsConsulta.hwnd
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
    clsGeneral.OcurrioError "Error al realizar la impresión", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub PropiedadesImpresion()
  
    With vsListado
        .PhysicalPage = True
        .PaperSize = vbPRPSLetter
        .Orientation = orLandscape
        .PreviewMode = pmScreen
        .PreviewPage = 1
        .Zoom = 100
        .MarginLeft = 500: .MarginRight = 350
        .MarginBottom = 750: .MarginTop = 750
    End With
    
End Sub

Private Sub tDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        tHasta.Value = DateAdd("d", -1, DateAdd("m", 1, tDesde.Value))
        tHasta.SetFocus
    End If
End Sub

Private Sub tDesde_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        tHasta.Value = DateAdd("d", -1, DateAdd("m", 1, tDesde.Value))
        tHasta.SetFocus
    End If
End Sub

Private Sub tHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then cDisponibilidad.SetFocus

End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionConsultar()
 
Dim aDisponibilidad As Long, aMovimiento As Long
Dim rs1 As rdoResultset

Dim aSubTotal As Currency, aSubTotalME As Currency
Dim aTotalGeneral As Currency, aTotalBancarias As Currency

Dim mTDebe As Currency, mTHaber As Currency, mTIVA As Currency, mTCofis As Currency
Dim bAlDebe As Boolean

Dim mTotalG_ME As Currency

    On Error GoTo ErrCDML
    If Not ValidoDatos Then Exit Sub
    
    Dim fDesde As Date, fHasta As Date
    ReDim arrRubros(0)
    
    Screen.MousePointer = 11
    chVista.Value = vbUnchecked
    
    fDesde = tDesde.Value: fHasta = tHasta.Value
    
    aDisponibilidad = 0: aMovimiento = 0: aTotalGeneral = 0: aTotalBancarias = 0
    vsConsulta.Rows = 1: vsConsulta.Refresh
    
    OBJ_Gastos.Rubro = "Por Gastos"
    OBJ_Gastos.ImporteD = 0: OBJ_Gastos.ImporteH = 0
    
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    pbProgreso.Value = 0
    cons = "Select Count(*) from  MovimientoDisponibilidad " _
                                    & " left Outer Join Compra On MDiIdCompra = ComCodigo" _
                                        & " left Outer Join GastoSubrubro On ComCodigo = GSrIDCompra, " _
                                & " MovimientoDisponibilidadRenglon " _
           & " Where MDIId = MDRIDMovimiento " _
           & " And MDiFecha Between '" & Format(fDesde, sqlFormatoF) & "' AND '" & Format(fHasta, sqlFormatoF) & " 23:59" & "'"
           
    cons = cons & " And MDRIdDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If rsAux(0) = 0 Then
            MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
            rsAux.Close: Screen.MousePointer = 0: Exit Sub
        End If
        pbProgreso.Max = rsAux(0)
    End If
    rsAux.Close
    '-----------------------------------------------------------------------------------------------------------------
     
    cons = "Select * from  MovimientoDisponibilidad " _
                                    & " left Outer Join Compra On MDiIdCompra = ComCodigo" _
                                        & " left Outer Join GastoSubrubro On ComCodigo = GSrIDCompra" _
                                            & " left Outer Join Subrubro On GSrIDSubrubro = SRuID " _
                                            & " left Outer Join Rubro On SRuRubro = RubID " _
                                            & " left Outer Join ProveedorCliente On ComProveedor = PClCodigo, " _
                                & " MovimientoDisponibilidadRenglon, " _
                                & " TipoMovDisponibilidad " _
           & " Where MDIId = MDRIDMovimiento " _
           & " And MDiFecha Between '" & Format(fDesde, sqlFormatoF) & "' AND '" & Format(fHasta, sqlFormatoF) & " 23:59" & "'" _
           & " And MDiTipo = TMDCodigo "
           
    cons = cons & " And MDRIdDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    cons = cons & " Order by MDRIdDisponibilidad, MDiFecha, MDRIdMovimiento"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCIÓN"
        Screen.MousePointer = 0: rsAux.Close: Exit Sub
    End If
    
    aMovimiento = 0: aSubTotal = 0: aSubTotalME = 0
    mTDebe = 0: mTHaber = 0: mTIVA = 0: mTCofis = 0
    
    vsConsulta.Redraw = False
    With vsConsulta
    
        If Not rsAux.EOF Then
            'Agrego el Saldo Inicial de la disponibilidad   ------------------------------
            bAlDebe = True
            cons = "Select Top 1 * FROM SaldoDisponibilidad, Disponibilidad " & _
                    " Where SDiFecha <= '" & Format(fDesde, "mm/dd/yyyy") & "'" & _
                    " And SDiHora = '00:00:00'" & _
                    " And DisID = SDiDisponibilidad " & _
                    " And SDiDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) & " Order by SDIFecha desc"
                    
            Set rs1 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not rs1.EOF Then
                bAlDebe = (rs1!SDiSaldo > 0)
                If rs1!SDiSaldo > 0 Then mTDebe = Abs(rs1!SDiSaldo) Else mTHaber = Abs(rs1!SDiSaldo)
                aTexto = Trim(rs1!DisNombre)
            End If
            rs1.Close
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, eCols.Fecha) = Format(fDesde, "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, eCols.Subrubro) = "Saldo Inicial " & aTexto
            .Cell(flexcpText, .Rows - 1, IIf(bAlDebe, eCols.Debe, eCols.Haber)) = Format(IIf(bAlDebe, mTDebe, mTHaber), FormatoMonedaP)
        End If
    
    Do While Not rsAux.EOF
        pbProgreso.Value = pbProgreso.Value + 1
        
        .AddItem ""
        bAlDebe = Not IsNull(rsAux!MDRDebe)
        If aMovimiento <> rsAux!MDiID Then
            
            If Not IsNull(rsAux!MDiIdCompra) Then
                .Cell(flexcpText, .Rows - 1, eCols.ID) = Format(rsAux!MDiIdCompra, "#,##0")
                If Not IsNull(rsAux!PClNombre) Then .Cell(flexcpText, .Rows - 1, eCols.Proveedor) = Trim(rsAux!PClNombre)
            Else
                .Cell(flexcpText, .Rows - 1, eCols.ID) = Format(rsAux!MDiID, "#,##0")
                If Not IsNull(rsAux!MDiComentario) Then .Cell(flexcpText, .Rows - 1, eCols.Proveedor) = Trim(rsAux!MDiComentario)
                
                If rsAux!TMDListado = 0 Then .Cell(flexcpData, .Rows - 1, eCols.Subrubro) = CStr(rsAux!MDiTipo)
            End If
            
            .Cell(flexcpText, .Rows - 1, eCols.Fecha) = Format(rsAux!MDiFecha, "dd/mm/yy")
            
            If Not IsNull(rsAux!ComCodigo) Then 'Documento
                If Not IsNull(rsAux!ComSerie) Then aTexto = Trim(rsAux!ComSerie) & " " Else aTexto = ""
                If Not IsNull(rsAux!ComNumero) Then aTexto = aTexto & rsAux!ComNumero
                .Cell(flexcpText, .Rows - 1, eCols.Documento) = aTexto
            End If
            
            
            If Not IsNull(rsAux!ComMoneda) Then
                If rsAux!ComMoneda = paMonedaPesos Then
                    If Not IsNull(rsAux!ComImporte) Then .Cell(flexcpText, .Rows - 1, eCols.Importe) = Format(rsAux!ComImporte, FormatoMonedaP)
                    If Not IsNull(rsAux!ComCofis) Then .Cell(flexcpText, .Rows - 1, eCols.Cofis) = Format(rsAux!ComCofis, FormatoMonedaP)
                    If Not IsNull(rsAux!ComIVA) Then .Cell(flexcpText, .Rows - 1, eCols.IVA) = Format(rsAux!ComIVA, FormatoMonedaP)
                Else
                    If Not IsNull(rsAux!ComImporte) Then .Cell(flexcpText, .Rows - 1, eCols.Importe) = Format(rsAux!ComImporte * rsAux!ComTC, FormatoMonedaP)
                    If Not IsNull(rsAux!ComCofis) Then .Cell(flexcpText, .Rows - 1, eCols.Cofis) = Format(rsAux!ComCofis * rsAux!ComTC, FormatoMonedaP)
                    If Not IsNull(rsAux!ComIVA) Then .Cell(flexcpText, .Rows - 1, eCols.IVA) = Format(rsAux!ComIVA * rsAux!ComTC, FormatoMonedaP)
                End If
            End If
            
            .Cell(flexcpText, .Rows - 1, IIf(bAlDebe, eCols.Debe, eCols.Haber)) = Format(rsAux!MDRImportePesos, FormatoMonedaP)
            'If bAlDebe Then
            '    mTDebe = mTDebe + .Cell(flexcpValue, .Rows - 1, eCols.Debe)
            'Else
            '    mTHaber = mTHaber + .Cell(flexcpValue, .Rows - 1, eCols.Haber)
            'End If
            
            mTIVA = mTIVA + (Abs(.Cell(flexcpValue, .Rows - 1, eCols.IVA)) * IIf(bAlDebe, 1, -1))
            mTCofis = mTCofis + (Abs(.Cell(flexcpValue, .Rows - 1, eCols.Cofis)) * IIf(bAlDebe, 1, -1))
                       
            mTotalG_ME = 0
            If Not IsNull(rsAux!ComCodigo) Then
                If rsAux!ComMoneda <> paMonedaPesos Then
                    mTotalG_ME = rsAux!ComImporte + IIf(IsNull(rsAux!ComIVA), 0, rsAux!ComIVA) + IIf(IsNull(rsAux!ComCofis), 0, rsAux!ComCofis)
                End If
            Else
                If bAlDebe Then
                    If rsAux!MDRDebe <> rsAux!MDRImportePesos Then mTotalG_ME = rsAux!MDRDebe
                Else
                    If rsAux!MDRHaber <> rsAux!MDRImportePesos Then mTotalG_ME = rsAux!MDRHaber
                End If
            End If
            If mTotalG_ME <> 0 Then .Cell(flexcpText, .Rows - 1, eCols.ImporteME) = Format(mTotalG_ME, "#,##0.00")
                            
            aSubTotalME = aSubTotalME + mTotalG_ME
            '-------------------------------------------------------------------------------------------------------------------------------------------------
        End If
            
        If Not IsNull(rsAux!RubNombre) Then     'Compras
            .Cell(flexcpText, .Rows - 1, eCols.Subrubro) = Trim(rsAux!RubNombre)
        Else
            'No es una compra es un Movimiento
            aTexto = ""
            If Not IsNull(rsAux!MDiTipo) Then
                Dim aSR As Long: aSR = 0
                Select Case rsAux!MDiTipo
                    Case paMCNotaCredito, paMCAnulacion: aSR = paSubrubroDeudoresPorVenta
                    Case paMCChequeDiferido: aSR = paSubrubroCDAlCobro
                    Case paMCVtaTelefonica: aSR = paSubrubroVtasTelACobrar
                    Case paMCLiquidacionCamionero: aSR = paSubrubroCobranzaVtasTel
                    Case paMCSenias: aSR = paSubrubroSeniasRecibidas
                    Case paMCIngresosOperativos: aSR = SRIngresosOperativos(rsAux!MDiComentario)
                        
                    Case Else     'Transferecnias entre cuentas, Hay que buscar la otra punta
                        If Not IsNull(rsAux!TMDTransferencia) Then
                            If rsAux!TMDTransferencia = 1 Then
                                cons = "Select RubCodigo, RubNombre from MovimientoDisponibilidadRenglon, Disponibilidad, SubRubro, Rubro " _
                                        & " Where MDRidMovimiento = " & rsAux!MDiID _
                                        & " And MDRIdDisponibilidad <> " & rsAux!MDRIdDisponibilidad _
                                        & " And MDRIdDisponibilidad = DisID And DisIDSubrubro = SRuID And SRuRubro = RubID"
                            
                                Set rs1 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                                If Not rs1.EOF Then aTexto = Format(rs1!RubCodigo, "000000000") & " " & Trim(rs1!RubNombre)
                                rs1.Close
                            End If
                        End If
                        aSR = 0     'Para que no entre el el IF
                End Select
                    
                If aSR <> 0 Then aTexto = RetornoConstanteSubrubro(aSR)
                If aTexto <> "" Then aTexto = mID(aTexto, InStr(aTexto, " ") + 1)
            End If
            
            .Cell(flexcpText, .Rows - 1, eCols.Subrubro) = aTexto
        End If
            
        If Not IsNull(rsAux!ComMoneda) And Not IsNull(rsAux!GSrImporte) Then
            Dim mImporteGPesos As Currency
            Dim mPorc As Currency
            Dim mTotalG As Currency
            
            If rsAux!ComMoneda = paMonedaPesos Then mImporteGPesos = rsAux!GSrImporte Else mImporteGPesos = rsAux!GSrImporte * rsAux!ComTC
            mImporteGPesos = Abs(mImporteGPesos)
            mTotalG = rsAux!ComImporte
            If Not IsNull(rsAux!ComIVA) Then mTotalG = mTotalG + rsAux!ComIVA
            If Not IsNull(rsAux!ComCofis) Then mTotalG = mTotalG + rsAux!ComCofis
            
            mPorc = rsAux!MDRImportePesos * 100 / mTotalG
            .Cell(flexcpText, .Rows - 1, IIf(bAlDebe, eCols.Debe, eCols.Haber)) = Format((mImporteGPesos * mPorc / 100), FormatoMonedaP)
            
            'If rsAux!ComMoneda = paMonedaPesos Then
            '    .Cell(flexcpText, .Rows - 1, IIf(bAlDebe, eCols.Debe, eCols.Haber)) = Format(Abs(rsAux!GSrImporte), FormatoMonedaP)
            'Else
            '    .Cell(flexcpText, .Rows - 1, IIf(bAlDebe, eCols.Debe, eCols.Haber)) = Format(Abs(rsAux!GSrImporte * rsAux!ComTC), FormatoMonedaP)
            'End If
        End If
        
        If bAlDebe Then
            mTDebe = mTDebe + .Cell(flexcpValue, .Rows - 1, eCols.Debe)
        Else
            mTHaber = mTHaber + .Cell(flexcpValue, .Rows - 1, eCols.Haber)
        End If
        
        aMovimiento = rsAux!MDiID
        
        If bAlDebe Then
            fnc_AddTotal txtRubro:=(.Cell(flexcpText, .Rows - 1, eCols.Subrubro)), valImporte:=.Cell(flexcpValue, .Rows - 1, eCols.Debe)
        Else
            fnc_AddTotal txtRubro:=(.Cell(flexcpText, .Rows - 1, eCols.Subrubro)), valImporte:=.Cell(flexcpValue, .Rows - 1, eCols.Haber) * -1
        End If
        
        If Not IsNull(rsAux!MDiIdCompra) Then
            If bAlDebe Then
                OBJ_Gastos.ImporteD = OBJ_Gastos.ImporteD + .Cell(flexcpValue, .Rows - 1, eCols.Debe)
            Else
                OBJ_Gastos.ImporteH = OBJ_Gastos.ImporteH + .Cell(flexcpValue, .Rows - 1, eCols.Haber)
            End If
        End If
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    'Agrego el SubTotal-------------------------------------------------------------------
    .AddItem ""
    .Cell(flexcpText, .Rows - 1, eCols.Subrubro) = "IVA"
    .Cell(flexcpText, .Rows - 1, IIf(mTIVA >= 0, eCols.Debe, eCols.Haber)) = Format(Abs(mTIVA), "#,##0.00")
    mTDebe = mTDebe + .Cell(flexcpValue, .Rows - 1, eCols.Debe): mTHaber = mTHaber + .Cell(flexcpValue, .Rows - 1, eCols.Haber)
    
    .AddItem ""
    .Cell(flexcpText, .Rows - 1, eCols.Subrubro) = "COFIS"
    .Cell(flexcpText, .Rows - 1, IIf(mTCofis >= 0, eCols.Debe, eCols.Haber)) = Format(Abs(mTCofis), "#,##0.00")
    mTDebe = mTDebe + .Cell(flexcpValue, .Rows - 1, eCols.Debe): mTHaber = mTHaber + .Cell(flexcpValue, .Rows - 1, eCols.Haber)
    
    .AddItem ""
    .Cell(flexcpText, .Rows - 1, eCols.Debe) = Format(mTDebe, "#,##0.00")
    .Cell(flexcpText, .Rows - 1, eCols.Haber) = Format(mTHaber, "#,##0.00")
    If aSubTotalME <> 0 Then .Cell(flexcpText, .Rows - 1, eCols.ImporteME) = Format(aSubTotalME, "#,##0.00")
    .Cell(flexcpBackColor, .Rows - 1, eCols.Debe, , eCols.ImporteME) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, eCols.Debe, , eCols.ImporteME) = True
    
    .AddItem ""
    .Cell(flexcpText, .Rows - 1, IIf(mTDebe > mTHaber, eCols.Debe, eCols.Haber)) = Format(IIf(mTDebe > mTHaber, (mTDebe - mTHaber), (mTHaber - mTDebe)), "#,##0.00")
    .Cell(flexcpBackColor, .Rows - 1, eCols.Debe, , eCols.ImporteME) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, eCols.Debe, , eCols.ImporteME) = True
    End With
    
    fnc_ResumenFinal mTIVA, mTCofis
    
    fnc_UnificoMovimientos
    
    Screen.MousePointer = 0
    vsConsulta.Redraw = True
    pbProgreso.Value = 0
    Exit Sub
    
ErrCDML:
    clsGeneral.OcurrioError "Error al cargar los datos.", Err.Description
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Function fnc_UnificoMovimientos()
Dim iX As Integer, iJJ As Integer
Dim fAux As Date, mID As Long
Dim bEXIT As Boolean, bEXIT2 As Boolean
Dim mImporte As Currency

    With vsConsulta
        Do While Not bEXIT
            
            If Val(.Cell(flexcpData, iX, eCols.Subrubro)) <> 0 And IsDate(.Cell(flexcpText, iX, eCols.Fecha)) Then
                
                mID = Val(.Cell(flexcpData, iX, eCols.Subrubro))
                fAux = CDate(.Cell(flexcpText, iX, eCols.Fecha))
                
                iJJ = iX + 1
                
                bEXIT2 = False
                Do While Not bEXIT2
                    If Not IsDate(.Cell(flexcpText, iJJ, eCols.Fecha)) Then Exit Do
                    If CDate(.Cell(flexcpText, iJJ, eCols.Fecha)) <> fAux Then Exit Do
                    
                    If Val(.Cell(flexcpData, iJJ, eCols.Subrubro)) = mID Then
                        mImporte = .Cell(flexcpValue, iX, eCols.Debe) - .Cell(flexcpValue, iX, eCols.Haber)
                        mImporte = mImporte + .Cell(flexcpValue, iJJ, eCols.Debe) - .Cell(flexcpValue, iJJ, eCols.Haber)
                        
                        .Cell(flexcpText, iX, IIf(mImporte >= 0, eCols.Debe, eCols.Haber)) = Format(Abs(mImporte), "#,##0.00")
                        .Cell(flexcpText, iX, eCols.Proveedor) = ""
                        .RemoveItem iJJ
                    Else
                        iJJ = iJJ + 1
                    End If
                Loop
                
            End If
            iX = iX + 1
            If iX >= (.Rows - 1) Then bEXIT = True
        Loop
        
    End With

End Function
   
Private Sub fnc_ResumenFinal(mTIVA As Currency, mTCofis As Currency)
Dim iX As Integer
Dim TOT_Debe As Currency, TOT_Haber As Currency

    With vsConsulta
        .AddItem ""
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, eCols.Subrubro) = "Resumen por Rubros"
        .Cell(flexcpText, .Rows - 1, eCols.Debe) = "Debe $": .Cell(flexcpText, .Rows - 1, eCols.Haber) = "Haber $"
        .Cell(flexcpBackColor, .Rows - 1, eCols.Subrubro, , eCols.Haber) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, eCols.Subrubro, , eCols.Haber) = True
        
        For iX = LBound(arrRubros) To UBound(arrRubros)
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, eCols.Subrubro) = arrRubros(iX).Rubro
            
            .Cell(flexcpText, .Rows - 1, eCols.Debe) = Format(arrRubros(iX).ImporteD, "#,##0.00")
            TOT_Debe = TOT_Debe + arrRubros(iX).ImporteD
    
            .Cell(flexcpText, .Rows - 1, eCols.Haber) = Format(arrRubros(iX).ImporteH, "#,##0.00")
            TOT_Haber = TOT_Haber + arrRubros(iX).ImporteH
        Next
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, eCols.Subrubro) = "IVA"
        .Cell(flexcpText, .Rows - 1, IIf(mTIVA >= 0, eCols.Debe, eCols.Haber)) = Format(Abs(mTIVA), "#,##0.00")
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, eCols.Subrubro) = "COFIS"
        .Cell(flexcpText, .Rows - 1, IIf(mTCofis >= 0, eCols.Debe, eCols.Haber)) = Format(Abs(mTCofis), "#,##0.00")
        
        If mTIVA >= 0 Then TOT_Debe = TOT_Debe + Abs(mTIVA) Else TOT_Haber = TOT_Haber + Abs(mTIVA)
        If mTCofis >= 0 Then TOT_Debe = TOT_Debe + Abs(mTCofis) Else TOT_Haber = TOT_Haber + Abs(mTCofis)
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, eCols.Debe) = Format(TOT_Debe, "#,##0.00")
        .Cell(flexcpText, .Rows - 1, eCols.Haber) = Format(TOT_Haber, "#,##0.00")
        .Cell(flexcpBackColor, .Rows - 1, eCols.Subrubro, , eCols.Haber) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, eCols.Subrubro, , eCols.Haber) = True
        
        .AddItem ""
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, eCols.Subrubro) = OBJ_Gastos.Rubro
        .Cell(flexcpText, .Rows - 1, eCols.Debe) = Format(OBJ_Gastos.ImporteD, "#,##0.00")
        .Cell(flexcpText, .Rows - 1, eCols.Haber) = Format(OBJ_Gastos.ImporteH, "#,##0.00")
    End With

End Sub

Private Function fnc_AddTotal(txtRubro As String, valImporte As Currency)
    
    If Trim(arrRubros(0).Rubro) = "" Then
        arrRubros(0).Rubro = txtRubro
        If valImporte >= 0 Then
            arrRubros(0).ImporteD = Abs(valImporte)
        Else
            arrRubros(0).ImporteH = Abs(valImporte)
        End If
        Exit Function
    End If
    
    Dim iX As Integer, bOk As Boolean
    
    bOk = False
    
    For iX = LBound(arrRubros) To UBound(arrRubros)
        With arrRubros(iX)
            If Trim(.Rubro) = Trim(txtRubro) Then
            
                If valImporte >= 0 Then .ImporteD = .ImporteD + Abs(valImporte) Else .ImporteH = .ImporteH + Abs(valImporte)
                bOk = True: Exit For
            End If
        End With
    Next
    If Not bOk Then
        iX = UBound(arrRubros) + 1
        ReDim Preserve arrRubros(iX)
        With arrRubros(iX)
            .Rubro = Trim(txtRubro)
            If valImporte >= 0 Then .ImporteD = Abs(valImporte) Else .ImporteH = Abs(valImporte)
        End With
    End If
    
End Function

Private Sub AccionLimpiar()
    cDisponibilidad.Text = ""
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

Private Function ValidoDatos() As Boolean
    On Error Resume Next
    ValidoDatos = False
    
    If Not IsDate(tDesde.Value) Or Not IsDate(tHasta.Value) Then
        MsgBox "Las fechas ingresadas para consultar no son correctas.", vbExclamation, "Faltan Datos"
        tDesde.SetFocus: Exit Function
    End If
    
    If tDesde.Value > tHasta.Value Then
        MsgBox "Las fechas ingresadas para consultar no son correctas.", vbExclamation, "Faltan Datos"
        tDesde.SetFocus: Exit Function
    End If
    
    If cDisponibilidad.ListIndex = -1 Then
        MsgBox "Seleccione la disponibilidad para consultar el Mayor Analítico.", vbExclamation, "Faltan Datos"
        cDisponibilidad.SetFocus: Exit Function
    End If
    
    ValidoDatos = True
    
End Function

Private Sub InicializoGrilla()

    On Error Resume Next
    With vsConsulta
        .Cols = 1: .Rows = 1:
        '.FormatString = ">#Ref|<Subrubro|<Documento|<Nº Cheque|>Importe (G) $|>Cofis (G) $|>I.V.A. (G) $|>Importe $|>Importe M/E|<Concepto|"
        
        .FormatString = ">#Ref|Fecha|<Rubro|<Documento|>Importe (G) $|>Cofis (G) $|>I.V.A. (G) $|>Debe $|>Haber $|>Importe M/E|<Proveedor|"
        
        .ColWidth(eCols.ID) = 850: .ColWidth(eCols.Fecha) = 750: .ColWidth(eCols.Subrubro) = 2500: .ColWidth(eCols.Documento) = 1100
        .ColWidth(eCols.Importe) = 1100: .ColWidth(eCols.Cofis) = 950: .ColWidth(eCols.Debe) = 1400
        .ColWidth(eCols.Haber) = 1400
        .ColWidth(eCols.ImporteME) = 1100: .ColWidth(eCols.Proveedor) = 4400
        
        .WordWrap = False
        .ColHidden(eCols.ID) = True
        
'        ">#Ref|Fecha|<Subrubro|<Documento|>Importe (G) $|>Cofis (G) $|>I.V.A. (G) $|>Debe $|>Haber $|>Importe M/E|<Proveedor|"
    End With
      
End Sub

Private Function zfn_EncabezadoListado(vsPrint As Control, strTitulo As String, sNombreEmpresa As Boolean)
    
    With vsPrint
        .HdrFont = "Arial"
        .HdrFontSize = 10
        .HdrFontBold = False
    End With
    
    If sNombreEmpresa Then
        vsPrint.Header = strTitulo + "||Carlos Gutiérrez S.A."
    Else
        vsPrint.Header = strTitulo
    End If
    'vsPrint.HdrFontBold = False: vsPrint.FontBold = False
    'vsPrint.HdrFontSize = 10: vsPrint.Footer = Format(Now, "dd/mm/yy hh:mm")
    
End Function

