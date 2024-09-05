VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.0#0"; "AACOMBO.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmDiarioEgresos 
   Caption         =   "Diario de Egresos de Caja"
   ClientHeight    =   7065
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDiarioEgresos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   2055
      Left            =   120
      TabIndex        =   15
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
      TabIndex        =   18
      Top             =   120
      Width           =   7695
      Begin AACombo99.AACombo cDisponibilidad 
         Height          =   315
         Left            =   3840
         TabIndex        =   3
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
      Begin VB.TextBox tFecha 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Disponibilidad:"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Fecha:"
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
      ScaleWidth      =   6075
      TabIndex        =   16
      Top             =   6240
      Width           =   6135
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmDiarioEgresos.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmDiarioEgresos.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmDiarioEgresos.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmDiarioEgresos.frx":0F38
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmDiarioEgresos.frx":1022
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmDiarioEgresos.frx":110C
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "frmDiarioEgresos.frx":1346
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Picture         =   "frmDiarioEgresos.frx":1448
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5400
         Picture         =   "frmDiarioEgresos.frx":180E
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmDiarioEgresos.frx":1910
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmDiarioEgresos.frx":1C12
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmDiarioEgresos.frx":1F54
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmDiarioEgresos.frx":2256
         Style           =   1  'Graphical
         TabIndex        =   5
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
      TabIndex        =   17
      Top             =   6810
      Width           =   7860
      _ExtentX        =   13864
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
            Object.Width           =   5662
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   3255
      Left            =   120
      TabIndex        =   19
      Top             =   3000
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
Attribute VB_Name = "frmDiarioEgresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsAux As rdoResultset
Private aTexto As String

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
    picBotones.BorderStyle = vbBSNone
    PropiedadesImpresion
    
    tFecha.Text = Format(Now, "dd/mm/yyyy")
    
    Cons = "Select DisId, DisNombre from Disponibilidad Order by DisNombre"
    CargoCombo Cons, cDisponibilidad
    
    CargoConstantesSubrubros
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = fFiltros.Width
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = vsListado.Height
    
    vsConsulta.ZOrder 0
    
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
    Foco tFecha
End Sub


Private Sub Label2_Click()
    Foco cDisponibilidad
End Sub

Private Sub tFecha_GotFocus()
    With tFecha: .SelStart = 0: .SelLength = Len(.Text): End With
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
    
    EncabezadoListado vsListado, "Diario de Egresos de Caja al " & Trim(tFecha.Text), True
    vsListado.filename = "Diario de Egresos de Caja"
    
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
    Screen.MousePointer = 0
    msgError.MuestroError "Ocurrió un error al realizar la impresión", Err.Description
End Sub

Private Sub PropiedadesImpresion()
  
  With vsListado
        .PaperSize = vbPRPSLetter
        .PhysicalPage = True
        .Orientation = orLandscape
        
        .PreviewMode = pmPrinter
        
        .PreviewPage = 1
        .FontName = "Tahoma": .FontSize = 10: .FontBold = False: .FontItalic = False
        .TextAlign = 0: .PageBorder = 3
        .Columns = 1
        .TableBorder = tbBoxRows
        .Zoom = 60
    End With

End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cDisponibilidad
End Sub

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionConsultar()
 
Dim aDisponibilidad As Long, aMovimiento As Long
Dim Rs1 As rdoResultset
Dim aSubTotal As Currency, aTotalGeneral As Currency, aTotalBancarias As Currency
Dim bBancaria As Boolean

    On Error GoTo ErrCDML
    If Not ValidoDatos Then Exit Sub
    
    Screen.MousePointer = 11
    chVista.Value = vbUnchecked
    
    aDisponibilidad = 0: aMovimiento = 0: aTotalGeneral = 0: aTotalBancarias = 0
    vsConsulta.Rows = 1
    
    Cons = "Select * from  MovimientoDisponibilidad " _
                                    & " left Outer Join Compra On MDiIdCompra = ComCodigo" _
                                        & " left Outer Join GastoSubrubro On ComCodigo = GSrIDCompra" _
                                            & " left Outer Join Subrubro On GSrIDSubrubro = SRuID, " _
                                & " MovimientoDisponibilidadRenglon left Outer Join Cheque On  MDRIdCheque = CheId, " _
                                & " TipoMovDisponibilidad " _
           & " Where MDIId = MDRIDMovimiento " _
           & " And MDRHaber <> NULL" _
           & " And MDiFecha = '" & Format(tFecha.Text, sqlFormatoF) & "'" _
           & " And MDiTipo = TMDCodigo and TMDListado = 1 "
           
    If cDisponibilidad.ListIndex <> -1 Then Cons = Cons & " And MDRIdDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    
    Cons = Cons & " Order by MDRIdDisponibilidad, MDRIdMovimiento"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCIÓN"
        Screen.MousePointer = 0: RsAux.Close: Exit Sub
    End If
    
    With vsConsulta
    Do While Not RsAux.EOF
    
        If aDisponibilidad <> RsAux!MDRIdDisponibilidad Then
            If aDisponibilidad <> 0 Then    '-------------------------------------------------------
                .AddItem ""     'Agrego el SubTotal
                .Cell(flexcpText, .Rows - 1, 6) = Format(aSubTotal, FormatoMonedaP)
                .Cell(flexcpBackColor, .Rows - 1, 6) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, 6) = True
                
                If bBancaria Then aTotalBancarias = aTotalBancarias + aSubTotal
                aTotalGeneral = aTotalGeneral + aSubTotal
            End If
            '-------------------------------------------------------------------------------------------
            aMovimiento = 0: aSubTotal = 0
            aDisponibilidad = RsAux!MDRIdDisponibilidad
            
            Cons = "Select * from Disponibilidad Where DisID = " & aDisponibilidad
            Set Rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not Rs1.EOF Then
                aTexto = Rs1!DisNombre
                If Not IsNull(Rs1!DisSucursal) Then bBancaria = True Else bBancaria = False
            End If
            Rs1.Close
            
            .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = aTexto
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio
        End If
        
        .AddItem ""
        
        If aMovimiento <> RsAux!MDiId Then
            If Not IsNull(RsAux!MDiIdCompra) Then .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!MDiIdCompra, "#,##0")
        End If
        
        If Not IsNull(RsAux!SRuCodigo) Then aTexto = Format(RsAux!SRuCodigo, "000000000") & " "
        
        If Not IsNull(RsAux!SRuNombre) Then
            aTexto = aTexto & Trim(RsAux!SRuNombre)
            .Cell(flexcpText, .Rows - 1, 1) = aTexto
        Else
            'Cargo el Proveedor del la compra
            If Not IsNull(RsAux!ComProveedor) Then
                Cons = "Select * from ProveedorCliente Where PClCodigo = " & RsAux!ComProveedor
                Set Rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not Rs1.EOF Then aTexto = Trim(Rs1!PClNombre)
                Rs1.Close
                .Cell(flexcpText, .Rows - 1, 1) = aTexto
            Else
                aTexto = ""
                If Not IsNull(RsAux!MDiTipo) Then
                    Dim aSR As Long: aSR = 0
                    Select Case RsAux!MDiTipo
                        Case paMCNotaCredito, paMCAnulacion: aSR = paSubrubroDeudoresPorVenta
                        Case paMCChequeDiferido: aSR = paSubrubroCDAlCobro
                        Case paMCVtaTelefonica: aSR = paSubrubroVtasTelACobrar
                        Case paMCLiquidacionCamionero: aSR = paSubrubroCobranzaVtasTel
                    End Select
                    If aSR <> 0 Then aTexto = RetornoConstanteSubrubro(aSR)
                End If
            
                .Cell(flexcpText, .Rows - 1, 1) = aTexto
            End If
        End If
        
        If aMovimiento <> RsAux!MDiId Then
            'Documento
            If Not IsNull(RsAux!ComCodigo) Then
                aTexto = RetornoNombreDocumento(RsAux!ComTipoDocumento, True) & " "
                If Not IsNull(RsAux!ComSerie) Then aTexto = aTexto & Trim(RsAux!ComSerie) & " "
                If Not IsNull(RsAux!ComNumero) Then aTexto = aTexto & RsAux!ComNumero
                .Cell(flexcpText, .Rows - 1, 2) = aTexto
            End If
            
            If Not IsNull(RsAux!CheSerie) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!CheSerie) & " " & RsAux!CheNumero
            
            If Not IsNull(RsAux!ComMoneda) Then
                If RsAux!ComMoneda = paMonedaPesos Then
                    If Not IsNull(RsAux!ComImporte) Then .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!ComImporte, FormatoMonedaP)
                    If Not IsNull(RsAux!ComIva) Then .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!ComIva, FormatoMonedaP)
                Else
                    If Not IsNull(RsAux!ComImporte) Then .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!ComImporte * RsAux!ComTC, FormatoMonedaP)
                    If Not IsNull(RsAux!ComIva) Then .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!ComIva * RsAux!ComTC, FormatoMonedaP)
                End If
            End If
            
            .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!MDRImportePesos, FormatoMonedaP)
            aSubTotal = aSubTotal + RsAux!MDRImportePesos
        
            If RsAux!SRuID = paSubrubroCompraMercaderia Then
                Cons = "Select * from ProveedorCliente Where PClCodigo = " & RsAux!ComProveedor
                Set Rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                If Not Rs1.EOF Then aTexto = Trim(Rs1!PClNombre)
                Rs1.Close
                .Cell(flexcpText, .Rows - 1, 7) = aTexto
            Else
                If Not IsNull(RsAux!MDiComentario) Then
                    .Cell(flexcpText, .Rows - 1, 7) = Trim(RsAux!MDiComentario)
                Else
                    If Not IsNull(RsAux!ComComentario) Then
                        .Cell(flexcpText, .Rows - 1, 7) = Trim(RsAux!ComComentario)
                    Else
                        .Cell(flexcpText, .Rows - 1, 7) = Trim(RsAux!TMDNombre)
                    End If
                End If
            End If
        End If
        
        If Not IsNull(RsAux!ComMoneda) And Not IsNull(RsAux!GSrImporte) Then
            If RsAux!ComMoneda = paMonedaPesos Then
                .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!GSrImporte, FormatoMonedaP)
            Else
                .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!GSrImporte * RsAux!ComTC, FormatoMonedaP)
            End If
        End If
        
        
        aMovimiento = RsAux!MDiId
        RsAux.MoveNext
    Loop
    RsAux.Close
       
    'Agrego el SubTotal-------------------------------------------------------------------
    .AddItem ""     'Agrego el SubTotal
    .Cell(flexcpText, .Rows - 1, 6) = Format(aSubTotal, FormatoMonedaP)
    .Cell(flexcpBackColor, .Rows - 1, 6) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, 6) = True
    
    If bBancaria Then aTotalBancarias = aTotalBancarias + aSubTotal
    aTotalGeneral = aTotalGeneral + aSubTotal
    '-------------------------------------------------------------------------------------------
    
    'Agrego el Resumen-------------------------------------------------------------------
    .AddItem "": .AddItem ""     'Agrego el SubTotales
    .Cell(flexcpText, .Rows - 1, 1) = "Total General"
    .Cell(flexcpText, .Rows - 1, 2) = Format(aTotalGeneral, FormatoMonedaP): .Cell(flexcpAlignment, .Rows - 1, 2) = 6
    .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
    
    If aTotalBancarias <> 0 Then
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = "Total Bancarias"
        .Cell(flexcpText, .Rows - 1, 2) = Format(aTotalBancarias, FormatoMonedaP): .Cell(flexcpAlignment, .Rows - 1, 2) = 6
        .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
    End If
    End With
    
    CargoDatosCheques
    
    Screen.MousePointer = 0
    Exit Sub
    
ErrCDML:
    Screen.MousePointer = 0
    clGeneral.OcurrioError "Ocurrió un error al cargar los datos.", Err.Description
End Sub

Private Sub CargoDatosCheques()

Dim aDisponibilidad As Long, aSubTotal As Currency
Dim Rs1 As rdoResultset

    aDisponibilidad = 0
    Cons = "Select CheID, CheIdDisponibilidad, CheSerie, CheNumero, CheLibrado, CheImporte from  MovimientoDisponibilidad, MovimientoDisponibilidadRenglon, Cheque " _
           & " Where MDIId = MDRIDMovimiento " _
           & " And MDRHaber <> NULL" _
           & " And MDRIdCheque = CheId " _
           & " And MDiFecha = '" & Format(tFecha.Text, sqlFormatoF) & "'"
    
    If cDisponibilidad.ListIndex <> -1 Then Cons = Cons & " And MDRIdDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
           
    Cons = Cons & " Group by CheID, CheIDDisponibilidad, CheSerie, CheNumero, CheLibrado, CheImporte  " _
                       & " Order by CheIDDisponibilidad, CheSerie, CheNumero"
           
    'Hay que agregar la fecha
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If RsAux.EOF Then RsAux.Close: Exit Sub
    With vsConsulta
    
    .AddItem "": .AddItem ""
    .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
    .Cell(flexcpText, .Rows - 1, 1) = "Disponibilidad": .Cell(flexcpText, .Rows - 1, 2) = "Nº Cheque": .Cell(flexcpText, .Rows - 1, 3) = "Librado": .Cell(flexcpText, .Rows - 1, 4) = "Importe"
    
    Do While Not RsAux.EOF
        If aDisponibilidad <> RsAux!CheIdDisponibilidad Then
            If aDisponibilidad <> 0 Then    '-------------------------------------------------------
                .AddItem "" 'Agrego el SubTotal
                .Cell(flexcpText, .Rows - 1, 4) = Format(aSubTotal, FormatoMonedaP)
                .Cell(flexcpBackColor, .Rows - 1, 4) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, 4) = True
                .AddItem ""
            End If
            '-------------------------------------------------------------------------------------------
            aSubTotal = 0
            aDisponibilidad = RsAux!CheIdDisponibilidad
            
            Cons = "Select * from Disponibilidad Where DisID = " & aDisponibilidad
            Set Rs1 = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
            If Not Rs1.EOF Then aTexto = Rs1!DisNombre
            Rs1.Close
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = aTexto
        Else
            .AddItem ""
        End If
        
        .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!CheSerie) & " " & RsAux!CheNumero
        .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!CheLibrado, "dd/mm/yyyy")
        .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!CheImporte, FormatoMonedaP)
        
        aSubTotal = aSubTotal + RsAux!CheImporte
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    .AddItem "" 'Agrego el SubTotal
    .Cell(flexcpText, .Rows - 1, 4) = Format(aSubTotal, FormatoMonedaP)
    .Cell(flexcpBackColor, .Rows - 1, 4) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, 4) = True
    
    End With
    
End Sub

Private Sub AccionLimpiar()
    tFecha.Text = ""
    cDisponibilidad.Text = ""
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

Private Function ValidoDatos() As Boolean
    On Error Resume Next
    ValidoDatos = False
    
    If Not IsDate(tFecha.Text) Then
        MsgBox "La fecha ingresada para consultar no es correcta.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    
    ValidoDatos = True
End Function

Private Sub InicializoGrilla()

    On Error Resume Next
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = ">ID Compra|<Subrubro|<Documento|<Nº Cheque|>Importe (C) $|>I.V.A. (C) $|>Importe $|<Concepto|"
        .ColWidth(0) = 900: .ColWidth(1) = 2800: .ColWidth(2) = 1200: .ColWidth(3) = 1250: .ColWidth(4) = 1200: .ColWidth(5) = 1200: .ColWidth(6) = 1400:: .ColWidth(7) = 4800
        
        .WordWrap = False
    End With
      
End Sub

