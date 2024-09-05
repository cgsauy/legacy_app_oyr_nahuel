VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.0#0"; "AACOMBO.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmListado 
   Caption         =   "Diario de Movimientos"
   ClientHeight    =   7065
   ClientLeft      =   165
   ClientTop       =   450
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
   ScaleHeight     =   7065
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   2055
      Left            =   120
      TabIndex        =   18
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
      TabIndex        =   21
      Top             =   120
      Width           =   11055
      Begin VB.TextBox tHasta 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2760
         MaxLength       =   12
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin AACombo99.AACombo cTipo 
         Height          =   315
         Left            =   6000
         TabIndex        =   5
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
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
         Left            =   840
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo de Movimientos:"
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   10995
      TabIndex        =   19
      Top             =   6240
      Width           =   11055
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   15
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   14
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
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5640
         Picture         =   "frmListado.frx":17C6
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
         Picture         =   "frmListado.frx":18C8
         Style           =   1  'Graphical
         TabIndex        =   7
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
         TabIndex        =   9
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
         TabIndex        =   8
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
      TabIndex        =   20
      Top             =   6810
      Width           =   11880
      _ExtentX        =   20955
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
            Object.Width           =   12753
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   3255
      Left            =   1200
      TabIndex        =   6
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
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsAux As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean

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

Private Sub cTipo_KeyDown(KeyCode As Integer, Shift As Integer)
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
    pbProgreso.Value = 0
    InicializoGrilla
    vsConsulta.ZOrder 0
    
    picBotones.BorderStyle = vbBSNone
    PropiedadesImpresion
    
    tDesde.Text = Format(PrimerDia(Now), "dd/mm/yyyy")
    tHasta.Text = Format(UltimoDia(Now), "dd/mm/yyyy")
    
    
    cTipo.AddItem "Entradas"
    cTipo.ItemData(cTipo.NewIndex) = 1
    cTipo.AddItem "Salidas"
    cTipo.ItemData(cTipo.NewIndex) = 2
    '--------------------------------------------------------------
    
    CargoConstantesSubrubros
    bCargarImpresion = True
    
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
    Set clGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Private Sub Label1_Click()
    Foco tDesde
End Sub

Private Sub Label2_Click()
    Foco cTipo
End Sub

Private Sub tDesde_GotFocus()
    With tDesde: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    If bCargarImpresion Then
        With vsListado
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        aTexto = "Diario de Movimientos (" & UCase(Trim(cTipo.Text)) & ") - Del " & Trim(tDesde.Text) & " al " & Trim(tHasta.Text)
        EncabezadoListado vsListado, aTexto, True
        vsListado.filename = "Diario de Movimientos"
        
        vsConsulta.ExtendLastCol = False
        vsListado.RenderControl = vsConsulta.hWnd
        vsConsulta.ExtendLastCol = True
        
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


Private Sub tDesde_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If IsDate(tDesde.Text) Then tHasta.Text = Format(UltimoDia(CDate(tDesde.Text)), "dd/mm/yyyy")
        Foco tHasta
    End If
    
End Sub

Private Sub tDesde_LostFocus()
    If IsDate(tDesde.Text) Then tDesde.Text = Format(tDesde.Text, "dd/mm/yyyy")
End Sub


Private Sub tHasta_GotFocus()
    With tHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cTipo
End Sub

Private Sub tHasta_LostFocus()
    If IsDate(tHasta.Text) Then tHasta.Text = Format(tHasta.Text, "dd/mm/yyyy")
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionConsultar()
 
Dim aDisponibilidad As Long, aMovimiento As Long

Dim Rs1 As rdoResultset
Dim aSubTotal As Currency, aTotalGeneral As Currency, aTotalBancarias As Currency
Dim bBancaria As Boolean

Dim aIDSR As Long, bIgualSR As Boolean
Dim aIDMov As Long, aFecha As Date

    On Error GoTo ErrCDML
    bCargarImpresion = True
    
    If Not ValidoDatos Then Exit Sub
    
    Screen.MousePointer = 11
    chVista.Value = vbUnchecked
    
    aDisponibilidad = 0: aMovimiento = 0: aTotalGeneral = 0: aTotalBancarias = 0
    vsConsulta.Rows = 1: vsConsulta.Refresh
    
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    pbProgreso.Value = 0
    Cons = "Select Count(*) from  MovimientoDisponibilidad " _
                                    & " left Outer Join Compra On MDiIdCompra = ComCodigo" _
                                        & " left Outer Join GastoSubrubro On ComCodigo = GSrIDCompra, " _
                                & " MovimientoDisponibilidadRenglon " _
           & " Where MDIId = MDRIDMovimiento " _
           & " And MDiFecha Between '" & Format(tDesde.Text, sqlFormatoF) & "' And '" & Format(tHasta.Text, sqlFormatoF) & "'" _
           
    Select Case cTipo.ItemData(cTipo.ListIndex)
        Case 1: Cons = Cons & " And MDRDebe <> NULL"
        Case 2: Cons = Cons & " And MDRHaber <> NULL"
    End Select
    
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
    
    Cons = "Select * from  MovimientoDisponibilidad " _
                                    & " left Outer Join Compra On MDiIdCompra = ComCodigo" _
                                        & " left Outer Join GastoSubrubro On ComCodigo = GSrIDCompra" _
                                            & " left Outer Join Subrubro On GSrIDSubrubro = SRuID, " _
                                & " MovimientoDisponibilidadRenglon left Outer Join Cheque On  MDRIdCheque = CheId, " _
                                & " TipoMovDisponibilidad " _
           & " Where MDIId = MDRIDMovimiento " _
           & " And MDiFecha Between '" & Format(tDesde.Text, sqlFormatoF) & "' And '" & Format(tHasta.Text, sqlFormatoF) & "'" _
           & " And MDiTipo = TMDCodigo"
    
    'Combo Tipo de Movimiento 1= Entrada     2= Salida
    Select Case cTipo.ItemData(cTipo.ListIndex)
        Case 1: Cons = Cons & " And MDRDebe <> NULL"
        Case 2: Cons = Cons & " And MDRHaber <> NULL"
    End Select
    
    Cons = Cons & " Order by MDiFecha, MDiIDCompra, GSRIdSubrubro"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCIÓN"
        Screen.MousePointer = 0: RsAux.Close: Exit Sub
    End If
        
    aIDMov = 0: aFecha = "1/1/1900"
    With vsConsulta
    Do While Not RsAux.EOF
        pbProgreso.Value = pbProgreso.Value + 1
        
        If aFecha <> RsAux!MDiFecha Then
            If aFecha > "1/1/1900" Then .AddItem "": .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!MDiFecha, "dd/mm/yyyy"): .Cell(flexcpText, .Rows - 1, 1) = " "
            
            aFecha = RsAux!MDiFecha
            
            .AddItem "": .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!MDiFecha, "dd/mm/yyyy")
            .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True: .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio
        End If
        
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!MDiFecha, "dd/mm/yyyy")
        
        If Not IsNull(RsAux!SRuCodigo) Then
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!SRuCodigo, "000000000") & " " & Trim(RsAux!SRuNombre)
        Else
            .Cell(flexcpText, .Rows - 1, 1) = " "
            
            'Cargo el Proveedor del la compra
            aTexto = ""
            If Not IsNull(RsAux!ComProveedor) Then
                Cons = "Select * from ProveedorCliente Where PClCodigo = " & RsAux!ComProveedor
                Set Rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not Rs1.EOF Then aTexto = aTexto & Trim(Rs1!PClNombre)
                Rs1.Close
                
            Else
                If Not IsNull(RsAux!MDiTipo) Then
                    Dim aSR As Long: aSR = 0
                    Select Case RsAux!MDiTipo
                        Case paMCNotaCredito, paMCAnulacion: aSR = paSubrubroDeudoresPorVenta
                        Case paMCChequeDiferido: aSR = paSubrubroCDAlCobro
                        Case paMCVtaTelefonica: aSR = paSubrubroVtasTelACobrar
                        Case paMCLiquidacionCamionero: aSR = paSubrubroCobranzaVtasTel
                        Case paMCSenias: aSR = paSubrubroSeniasRecibidas
                        
                        Case Else     'Transferecnias entre cuentas, Hay que buscar la otra punta
                                If Not IsNull(RsAux!TMDTransferencia) Then
                                    If RsAux!TMDTransferencia = 1 Then
                                        Cons = "Select * from MovimientoDisponibilidadRenglon, Disponibilidad, SubRubro " _
                                                & " Where MDRidMovimiento = " & RsAux!MDiID _
                                                & " And MDRIdDisponibilidad <> " & RsAux!MDRIdDisponibilidad _
                                                & " And MDRIdDisponibilidad = DisID And DisIDSubrubro = SRuID"
                                    
                                        Set Rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                                        If Not Rs1.EOF Then aTexto = Format(Rs1!SRuCodigo, "000000000") & " " & Trim(Rs1!SRuNombre)
                                        Rs1.Close
                                    End If
                                End If
                                aSR = 0     'Para que no entre el el IF
                    End Select
                    
                    If aSR <> 0 Then aTexto = RetornoConstanteSubrubro(aSR)
                End If
            End If
            
            If aTexto <> "" Then .Cell(flexcpText, .Rows - 1, 1) = aTexto
        End If
                            
        'Cargo el Concepto -------> Siempre que hay compra cargo el Proveedor               -----------------------------------------------
        '                           --------> Si el proveedor es N/D cargo el Rubro
        aTexto = ""
        If Not IsNull(RsAux!ComProveedor) Then
            Cons = "Select * from ProveedorCliente Where PClCodigo = " & RsAux!ComProveedor
            Set Rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not Rs1.EOF Then aTexto = Trim(Rs1!PClNombre)
            Rs1.Close
            If Trim(UCase(aTexto)) = "N/D" And Not IsNull(RsAux!SRuCodigo) Then
                Cons = "Select * from Rubro Where RubID = " & RsAux!SRuRubro
                Set Rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not Rs1.EOF Then aTexto = Trim(Rs1!RubNombre)
                Rs1.Close
            End If
        End If
        If Not IsNull(RsAux!MDiComentario) Then
            If Trim(RsAux!MDiComentario) <> Trim(aTexto) And UCase(Trim(RsAux!MDiComentario)) <> "N/D" Then
                If aTexto <> "" Then aTexto = aTexto & " // "
                aTexto = aTexto & Trim(RsAux!MDiComentario)
            End If
        Else
            If Trim(RsAux!ComComentario) <> Trim(aTexto) Then
                If aTexto <> "" Then aTexto = aTexto & " // "
                aTexto = aTexto & Trim(RsAux!ComComentario)
            Else
                If aTexto = "" Then aTexto = Trim(RsAux!TMDNombre)
            End If
        End If
        .Cell(flexcpText, .Rows - 1, 2) = aTexto
        '-------------------------------------------------------------------------------------------------------------------------------------------------
            
        
        If aIDMov <> RsAux!MDiID Then
            
            If Not IsNull(RsAux!CheSerie) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!CheSerie) & " " & RsAux!CheNumero
            If Not IsNull(RsAux!MDiIdCompra) Then .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!MDiIdCompra, "#,##0")
            
            If Not IsNull(RsAux!GSRIdSubrubro) Then aIDSR = RsAux!GSRIdSubrubro Else aIDSR = 0
            bIgualSR = False
            
            'OJO CON EL IVA DE LA COMPRA !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            If Not IsNull(RsAux!ComMoneda) Then
                If RsAux!ComMoneda = paMonedaPesos Then
                    If Not IsNull(RsAux!ComImporte) Then .Cell(flexcpText, .Rows - 1, 5) = Format(Abs(RsAux!ComImporte), FormatoMonedaP)
                    If Not IsNull(RsAux!ComIva) Then .Cell(flexcpText, .Rows - 1, 6) = Format(Abs(RsAux!ComIva), FormatoMonedaP)
                Else
                    If Not IsNull(RsAux!ComImporte) Then .Cell(flexcpText, .Rows - 1, 5) = Format(Abs(RsAux!ComImporte) * RsAux!ComTC, FormatoMonedaP)
                    If Not IsNull(RsAux!ComIva) Then .Cell(flexcpText, .Rows - 1, 6) = Format(Abs(RsAux!ComIva) * RsAux!ComTC, FormatoMonedaP)
                End If
            
'                .Cell(flexcpText, .Rows - 1, 7) = Format(.Cell(flexcpText, .Rows - 1, 5) + .Cell(flexcpText, .Rows - 1, 6), FormatoMonedaP)
                
            Else
                .Cell(flexcpText, .Rows - 1, 5) = Format(Abs(RsAux!MDRImportePesos), FormatoMonedaP)
            End If
            
        Else
'            If Not IsNull(RsAux!GSRIdSubrubro) Then
                If aIDSR = RsAux!GSRIdSubrubro Then bIgualSR = True
'            End If
        End If
        
        If Not IsNull(RsAux!ComMoneda) And Not IsNull(RsAux!GSrImporte) Then
            If Not bIgualSR And RsAux!GSRIdSubrubro <> aIDSR Then
                aIDSR = RsAux!GSRIdSubrubro
                If RsAux!ComMoneda = paMonedaPesos Then
                    If Not IsNull(RsAux!GSrImporte) Then .Cell(flexcpText, .Rows - 1, 5) = Format(Abs(RsAux!GSrImporte), FormatoMonedaP)
                Else
                    If Not IsNull(RsAux!GSrImporte) Then .Cell(flexcpText, .Rows - 1, 5) = Format(Abs(RsAux!GSrImporte) * RsAux!ComTC, FormatoMonedaP)
                End If
            End If
        End If
                        
        .Cell(flexcpText, .Rows - 1, 7) = Format(Abs(RsAux!MDRImportePesos), FormatoMonedaP)
            
        aIDMov = RsAux!MDiID

        RsAux.MoveNext
    Loop
    RsAux.Close
    
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, 0, 5, , Colores.Obligatorio, , True, "Subtotal"
    .Subtotal flexSTSum, 0, 6, , Colores.Rojo, Colores.Blanco, True
    .Subtotal flexSTSum, 0, 7, , Colores.Rojo, Colores.Blanco, True
    
    .Subtotal flexSTSum, -1, 5, , Colores.Rojo, Colores.Blanco, True, "Total General"
    .Subtotal flexSTSum, -1, 6, , Colores.Rojo, Colores.Blanco, True
    .Subtotal flexSTSum, -1, 7, , Colores.Rojo, Colores.Blanco, True
    
    End With
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    Exit Sub
    
ErrCDML:
    Screen.MousePointer = 0
    clGeneral.OcurrioError "Ocurrió un error al cargar los datos.", Err.Description
End Sub

Private Function CargoResumenMovimientos(idDisponibilidad) As Currency
Dim RsRsm As rdoResultset
Dim aRetorno As Currency
Dim aSR As Long: aSR = 0
    
    aRetorno = 0
    Cons = "Select MDiTipo, TMDNombre, Importe = Sum(MDRImportePesos) " _
           & " From  MovimientoDisponibilidad, MovimientoDisponibilidadRenglon, TipoMovDisponibilidad" _
           & " Where MDIId = MDRIDMovimiento" _
           & " And MDRDebe <> NULL " _
           & " And MDiFecha = '" & Format(tDesde.Text, sqlFormatoF) & "'" _
           & " And MDRIdDisponibilidad = " & idDisponibilidad _
           & " And MDiTipo = TMDCodigo " _
           & " And TMDListado = 0 " _
           & " Group by MDiTipo, TMDNombre"
    Set RsRsm = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    With vsConsulta
    Do While Not RsRsm.EOF
        .AddItem ""
        aRetorno = aRetorno + RsRsm!Importe
        aTexto = ""
        If Not IsNull(RsRsm!MDiTipo) Then
            Select Case RsRsm!MDiTipo
                Case paMCNotaCredito, paMCAnulacion: aSR = paSubrubroDeudoresPorVenta
                Case paMCChequeDiferido: aSR = paSubrubroCDAlCobro
                Case paMCVtaTelefonica: aSR = paSubrubroVtasTelACobrar
                Case paMCLiquidacionCamionero: aSR = paSubrubroCobranzaVtasTel
            End Select
            If aSR <> 0 Then aTexto = RetornoConstanteSubrubro(aSR)
        End If
        .Cell(flexcpText, .Rows - 1, 1) = aTexto
        
        
        .Cell(flexcpText, .Rows - 1, 6) = Format(RsRsm!Importe, FormatoMonedaP)
    
        .Cell(flexcpText, .Rows - 1, 7) = Trim(RsRsm!TMDNombre)
    
        RsRsm.MoveNext
    Loop
    RsRsm.Close
    
    CargoResumenMovimientos = aRetorno
    End With
    
End Function

Private Sub AccionLimpiar()
    tDesde.Text = ""
    cTipo.Text = ""
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

Private Function ValidoDatos() As Boolean
    On Error Resume Next
    ValidoDatos = False
    
    If Not IsDate(tDesde.Text) Then
        MsgBox "La fecha ingresada para consultar no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    If Not IsDate(tHasta.Text) Then
        MsgBox "La fecha ingresada para consultar no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tHasta: Exit Function
    End If
    
    If CDate(tDesde.Text) > CDate(tHasta.Text) Then
        MsgBox "El rango de fechas ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tHasta: Exit Function
    End If

    If cTipo.ListIndex = -1 Then
        MsgBox "Debe seleccionar el tipo de movimiento a consultar (entradas o salidas).", vbExclamation, "ATENCIÓN"
        Foco cTipo: Exit Function
    End If
    
    ValidoDatos = True
    
End Function

Private Sub InicializoGrilla()

    On Error Resume Next
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = "Día|Subrubro|<Descripción|<Nº Cheque|<ID_Compra|>Importe $|>I.V.A. $|>Total $|"
        .ColWidth(0) = 0
        .ColWidth(1) = 3250: .ColWidth(2) = 2250: .ColWidth(3) = 1250
        .ColWidth(5) = 1200: .ColWidth(6) = 1200: .ColWidth(7) = 1400
        
        .WordWrap = False
        .MergeCells = flexMergeSpill
    End With
      
End Sub


