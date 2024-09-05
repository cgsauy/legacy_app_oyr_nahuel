VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmDiarioIngresos 
   Caption         =   "Diario de Ingresos de Caja"
   ClientHeight    =   7065
   ClientLeft      =   2100
   ClientTop       =   1515
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
   Icon            =   "frmDiarioIngresos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   7860
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
      ScaleWidth      =   7275
      TabIndex        =   16
      Top             =   6240
      Width           =   7335
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmDiarioIngresos.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmDiarioIngresos.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmDiarioIngresos.frx":0ABE
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
         Picture         =   "frmDiarioIngresos.frx":0F38
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
         Picture         =   "frmDiarioIngresos.frx":1022
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
         Picture         =   "frmDiarioIngresos.frx":110C
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
         Picture         =   "frmDiarioIngresos.frx":1346
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
         Picture         =   "frmDiarioIngresos.frx":1448
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5640
         Picture         =   "frmDiarioIngresos.frx":180E
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
         Picture         =   "frmDiarioIngresos.frx":1910
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
         Picture         =   "frmDiarioIngresos.frx":1C12
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
         Picture         =   "frmDiarioIngresos.frx":1F54
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
         Picture         =   "frmDiarioIngresos.frx":2256
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6240
         TabIndex        =   21
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
      Left            =   1200
      TabIndex        =   19
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
Attribute VB_Name = "frmDiarioIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsAux As rdoResultset
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
    pbProgreso.Value = 0
    InicializoGrilla
    vsConsulta.ZOrder 0
    
    picBotones.BorderStyle = vbBSNone
    PropiedadesImpresion
    
    tFecha.Text = Format(Now, "dd/mm/yyyy")
    
    'Cargo disponibilidades.-------------------------------
    cons = "Select DisID, DisNombre From NivelPermiso, Disponibilidad " _
        & " Where NPeNivel IN (Select UNiNivel From UsuarioNivel Where UNiUsuario = " & paCodigoDeUsuario & ")" _
        & " And NPeAplicacion = DisAplicacion"
    
    CargoCombo cons, cDisponibilidad
    
    aDisponibilidades = ""
    For I = 0 To cDisponibilidad.ListCount - 1
        aDisponibilidades = aDisponibilidades & cDisponibilidad.ItemData(I) & ","
    Next I
    aDisponibilidades = Mid(aDisponibilidades, 1, Len(aDisponibilidades) - 1)
    '--------------------------------------------------------------
        
    CargoConstantesSubrubros
    
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
    
    EncabezadoListado vsListado, "Diario de Ingresos de Caja al " & Trim(tFecha.Text), True
    vsListado.FileName = "Diario de Ingresos de Caja"
    
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
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
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
Dim rs1 As rdoResultset

Dim aSubTotal As Currency, aSubTotalME As Currency
Dim aTotalGeneral As Currency, aTotalBancarias As Currency
Dim bBancaria As Boolean

Dim mTotalG_ME As Currency

    On Error GoTo ErrCDML
    If Not ValidoDatos Then Exit Sub
    
    Screen.MousePointer = 11
    chVista.Value = vbUnchecked
    
    aDisponibilidad = 0: aMovimiento = 0: aTotalGeneral = 0: aTotalBancarias = 0
    vsConsulta.Rows = 1: vsConsulta.Refresh
    
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    pbProgreso.Value = 0
    cons = "Select Count(*) from  MovimientoDisponibilidad " _
                                    & " left Outer Join Compra On MDiIdCompra = ComCodigo" _
                                        & " left Outer Join GastoSubrubro On ComCodigo = GSrIDCompra, " _
                                & " MovimientoDisponibilidadRenglon " _
           & " Where MDIId = MDRIDMovimiento " _
           & " And MDRDebe <> NULL" _
           & " And MDiFecha = '" & Format(tFecha.Text, sqlFormatoF) & "'" _
           
    If cDisponibilidad.ListIndex <> -1 Then
        cons = cons & " And MDRIdDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    Else
        cons = cons & " And MDRIdDisponibilidad In (" & aDisponibilidades & ")"
    End If
    
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
                                            & " left Outer Join Subrubro On GSrIDSubrubro = SRuID, " _
                                & " MovimientoDisponibilidadRenglon left Outer Join Cheque On  MDRIdCheque = CheId, " _
                                & " TipoMovDisponibilidad " _
           & " Where MDIId = MDRIDMovimiento " _
           & " And MDRDebe <> NULL" _
           & " And MDiFecha = '" & Format(tFecha.Text, sqlFormatoF) & "'" _
           & " And MDiTipo = TMDCodigo"
           
    If cDisponibilidad.ListIndex <> -1 Then
        cons = cons & " And MDRIdDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    Else
        cons = cons & " And MDRIdDisponibilidad IN (" & aDisponibilidades & ")"
    End If
    
    cons = cons & " Order by MDRIdDisponibilidad, MDRIdMovimiento"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If rsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCIÓN"
        Screen.MousePointer = 0: rsAux.Close: Exit Sub
    End If
    
    With vsConsulta
    Do While Not rsAux.EOF
        pbProgreso.Value = pbProgreso.Value + 1
        If aDisponibilidad <> rsAux!MDRIdDisponibilidad Then
            
            If aDisponibilidad <> 0 Then    '-------------------------------------------------------
                aSubTotal = aSubTotal + CargoResumenMovimientos(aDisponibilidad)
                .AddItem ""     'Agrego el SubTotal
                .Cell(flexcpText, .Rows - 1, 7) = Format(aSubTotal, "#,##0.00")
                .Cell(flexcpBackColor, .Rows - 1, 7) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, 7) = True
                If aSubTotalME <> 0 Then
                    .Cell(flexcpText, .Rows - 1, 8) = Format(aSubTotalME, "#,##0.00")
                    .Cell(flexcpBackColor, .Rows - 1, 8) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, 8) = True
                End If
                
                If bBancaria Then aTotalBancarias = aTotalBancarias + aSubTotal
                aTotalGeneral = aTotalGeneral + aSubTotal
            End If
            '-------------------------------------------------------------------------------------------
           
            aMovimiento = 0: aSubTotal = 0: aSubTotalME = 0
            aDisponibilidad = rsAux!MDRIdDisponibilidad
            
            cons = "Select * from Disponibilidad Where DisID = " & aDisponibilidad
            Set rs1 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            If Not rs1.EOF Then
                aTexto = rs1!DisNombre
                If Not IsNull(rs1!DisSucursal) Then bBancaria = True Else bBancaria = False
            End If
            rs1.Close
            
            .AddItem "": .Cell(flexcpText, .Rows - 1, 1) = aTexto
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio
            
        End If
        
        If rsAux!TMDListado Then
            .AddItem ""
            
            If aMovimiento <> rsAux!MDIId Then
                If Not IsNull(rsAux!MDiIdCompra) Then
                    .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!MDiIdCompra, "#,##0")
                Else
                    .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!MDIId, "#,##0")
                End If
                
                'Cargo el Concepto -------> Siempre que hay compra cargo el Proveedor               -----------------------------------------------
                '                           --------> Si el proveedor es N/D cargo el Rubro
                aTexto = ""
                If Not IsNull(rsAux!ComProveedor) Then
                    cons = "Select * from ProveedorCliente Where PClCodigo = " & rsAux!ComProveedor
                    Set rs1 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                    If Not rs1.EOF Then aTexto = Trim(rs1!PClFantasia)
                    rs1.Close
                    If Trim(UCase(aTexto)) = "N/D" And Not IsNull(rsAux!SRuCodigo) Then
                        cons = "Select * from Rubro Where RubID = " & rsAux!SRuRubro
                        Set rs1 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                        If Not rs1.EOF Then aTexto = Trim(rs1!RubNombre)
                        rs1.Close
                    End If
                End If
                If Not IsNull(rsAux!MDiComentario) Then
                    If Trim(rsAux!MDiComentario) <> Trim(aTexto) And UCase(Trim(rsAux!MDiComentario)) <> "N/D" Then
                        If aTexto <> "" Then aTexto = aTexto & " // "
                        aTexto = aTexto & Trim(rsAux!MDiComentario)
                    End If
                Else
                    If Trim(rsAux!ComComentario) <> Trim(aTexto) And UCase(Trim(rsAux!ComComentario)) <> "N/D" Then
                        If aTexto <> "" Then aTexto = aTexto & " // "
                        aTexto = aTexto & Trim(rsAux!ComComentario)
                    Else
                        If aTexto = "" Then aTexto = Trim(rsAux!TMDNombre)
                    End If
                End If
                
                mTotalG_ME = 0
                If Not IsNull(rsAux!ComCodigo) Then
                    If rsAux!ComMoneda <> paMonedaPesos Then
                        mTotalG_ME = rsAux!ComImporte + IIf(IsNull(rsAux!ComIVa), 0, rsAux!ComIVa) + IIf(IsNull(rsAux!ComCofis), 0, rsAux!ComCofis)
                        aTexto = "TC " & Format(rsAux!ComTC, "0.0##") & "  " & aTexto
                    End If
                Else
                    If rsAux!MDRDebe <> rsAux!MDRImportePesos Then
                        aTexto = "TC " & Format(rsAux!MDRImportePesos / rsAux!MDRDebe, "0.0##") & "  " & aTexto
                        mTotalG_ME = rsAux!MDRDebe
                    End If
                End If
                
                If mTotalG_ME <> 0 Then .Cell(flexcpText, .Rows - 1, 8) = Format(mTotalG_ME, "#,##0.00")
                aSubTotalME = aSubTotalME + mTotalG_ME
                .Cell(flexcpText, .Rows - 1, 9) = aTexto
                '-------------------------------------------------------------------------------------------------------------------------------------------------
            End If
            
            If Not IsNull(rsAux!SRuCodigo) Then aTexto = Format(rsAux!SRuCodigo, "000000000") & " "
            
            If Not IsNull(rsAux!SRuNombre) Then
                aTexto = aTexto & Trim(rsAux!SRuNombre)
                .Cell(flexcpText, .Rows - 1, 1) = aTexto
            Else
                'Cargo el Proveedor del la compra
                aTexto = ""
                If Not IsNull(rsAux!ComProveedor) Then
                    cons = "Select * from ProveedorCliente Where PClCodigo = " & rsAux!ComProveedor
                    Set rs1 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                    If Not rs1.EOF Then aTexto = aTexto & Trim(rs1!PClNombre)
                    rs1.Close
                Else
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
                                        cons = "Select * from MovimientoDisponibilidadRenglon, Disponibilidad, SubRubro " _
                                                & " Where MDRidMovimiento = " & rsAux!MDIId _
                                                & " And MDRIdDisponibilidad <> " & rsAux!MDRIdDisponibilidad _
                                                & " And MDRIdDisponibilidad = DisID And DisIDSubrubro = SRuID"
                                    
                                        Set rs1 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                                        If Not rs1.EOF Then
                                            aTexto = Format(rs1!SRuCodigo, "000000000") & " " & Trim(rs1!SRuNombre)
                                            'If rs1!MDRHaber <> rs1!MDRImportePesos Then
                                            '    .Cell(flexcpText, .Rows - 1, 9) = "TC " & Format(rs1!MDRImportePesos / rs1!MDRHaber, "0.0##") & "  " & .Cell(flexcpText, .Rows - 1, 9)
                                            '    mTotalG_ME = rs1!MDRHaber
                                            'End If
                                        End If
                                        rs1.Close
                                        
                                        'If mTotalG_ME <> 0 Then .Cell(flexcpText, .Rows - 1, 8) = Format(mTotalG_ME, "#,##0.00")
                                        'aSubTotalME = aSubTotalME + mTotalG_ME
                                    End If
                                End If
                                aSR = 0     'Para que no entre el el IF
                                    
                        End Select
                        If aSR <> 0 Then aTexto = RetornoConstanteSubrubro(aSR)
                    End If
                End If
                
                .Cell(flexcpText, .Rows - 1, 1) = aTexto
            End If
            
            If aMovimiento <> rsAux!MDIId Then
                
                If Not IsNull(rsAux!ComCodigo) Then 'Documento
                    aTexto = RetornoNombreDocumento(rsAux!ComTipoDocumento, True) & " "
                    If Not IsNull(rsAux!ComSerie) Then aTexto = aTexto & Trim(rsAux!ComSerie) & " "
                    If Not IsNull(rsAux!ComNumero) Then aTexto = aTexto & rsAux!ComNumero
                    .Cell(flexcpText, .Rows - 1, 2) = aTexto
                End If
                
                If Not IsNull(rsAux!CheSerie) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(rsAux!CheSerie) & " " & rsAux!CheNumero
                
                If Not IsNull(rsAux!ComMoneda) Then
                    If rsAux!ComMoneda = paMonedaPesos Then
                        If Not IsNull(rsAux!ComImporte) Then .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!ComImporte, FormatoMonedaP)
                        If Not IsNull(rsAux!ComCofis) Then .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!ComCofis, FormatoMonedaP)
                        If Not IsNull(rsAux!ComIVa) Then .Cell(flexcpText, .Rows - 1, 6) = Format(rsAux!ComIVa, FormatoMonedaP)
                    Else
                        If Not IsNull(rsAux!ComImporte) Then .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!ComImporte * rsAux!ComTC, FormatoMonedaP)
                        If Not IsNull(rsAux!ComCofis) Then .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!ComCofis * rsAux!ComTC, FormatoMonedaP)
                        If Not IsNull(rsAux!ComIVa) Then .Cell(flexcpText, .Rows - 1, 6) = Format(rsAux!ComIVa * rsAux!ComTC, FormatoMonedaP)
                    End If
                End If
                
                .Cell(flexcpText, .Rows - 1, 7) = Format(rsAux!MDRImportePesos, FormatoMonedaP)
                aSubTotal = aSubTotal + rsAux!MDRImportePesos
            End If
            
            If Not IsNull(rsAux!ComMoneda) And Not IsNull(rsAux!GSrImporte) Then
                If rsAux!ComMoneda = paMonedaPesos Then
                    If Not IsNull(rsAux!GSrImporte) Then .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!GSrImporte, FormatoMonedaP)
                Else
                    If Not IsNull(rsAux!GSrImporte) Then .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!GSrImporte * rsAux!ComTC, FormatoMonedaP)
                End If
            End If
                            
            aMovimiento = rsAux!MDIId
        End If

        rsAux.MoveNext
    Loop
    rsAux.Close
    
    If .Rows > 1 Then
        'Agrego el SubTotal-------------------------------------------------------------------
        aSubTotal = aSubTotal + CargoResumenMovimientos(aDisponibilidad)
        .AddItem ""     'Agrego el SubTotal
        .Cell(flexcpText, .Rows - 1, 7) = Format(aSubTotal, "#,##0.00")
        .Cell(flexcpBackColor, .Rows - 1, 7) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, 7) = True
        If aSubTotalME <> 0 Then
            .Cell(flexcpText, .Rows - 1, 8) = Format(aSubTotalME, "#,##0.00")
            .Cell(flexcpBackColor, .Rows - 1, 8) = Colores.Obligatorio: .Cell(flexcpFontBold, .Rows - 1, 8) = True
        End If
                
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
    Else
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCIÓN"
    End If
    
    End With
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    Exit Sub
    
ErrCDML:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function CargoResumenMovimientos(idDisponibilidad) As Currency
Dim RsRsm As rdoResultset
Dim aRetorno As Currency
Dim aSR As Long: aSR = 0
    
Dim bAdd As Boolean

    aRetorno = 0
    cons = "Select MDiTipo, TMDNombre, TMDSubRubro, TMDTransferencia, Importe = Sum(MDRImportePesos) " _
           & " From  MovimientoDisponibilidad, MovimientoDisponibilidadRenglon, TipoMovDisponibilidad" _
           & " Where MDIId = MDRIDMovimiento" _
           & " And MDRDebe is Not NULL " _
           & " And MDiFecha = '" & Format(tFecha.Text, sqlFormatoF) & "'" _
           & " And MDRIdDisponibilidad = " & idDisponibilidad _
           & " And MDiTipo = TMDCodigo " _
           & " And TMDListado = 0 " _
           & " Group by MDiTipo, TMDNombre, TMDSubRubro, TMDTransferencia"
    Set RsRsm = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    
    With vsConsulta
    Do While Not RsRsm.EOF
        
        aRetorno = aRetorno + RsRsm!Importe
        aTexto = ""
        'If Not IsNull(RsRsm!MDiTipo) Then
        '    Select Case RsRsm!MDiTipo
        '        Case paMCNotaCredito, paMCAnulacion: aSR = paSubrubroDeudoresPorVenta
        '        Case paMCChequeDiferido: aSR = paSubrubroCDAlCobro
        '        Case paMCVtaTelefonica: aSR = paSubrubroVtasTelACobrar
        '        Case paMCLiquidacionCamionero: aSR = paSubrubroCobranzaVtasTel
        '    End Select
        '    If aSR <> 0 Then aTexto = RetornoConstanteSubrubro(aSR)
        'End If
        bAdd = True
        If Not IsNull(RsRsm!TMDTransferencia) Then
            If RsRsm!TMDTransferencia = 1 Then
                Dim rs1 As rdoResultset
                
                cons = "Select DisIDSubrubro, Importe = Sum(MDRImportePesos) " _
                        & " From MovimientoDisponibilidadRenglon, Disponibilidad " _
                        & " Where MDRDebe is  NULL " _
                        & " And MDRIdDisponibilidad <> " & idDisponibilidad _
                        & " And MDRIdDisponibilidad = DisID" _
                        & " And MDRIdMovimiento IN ( " _
                                    & " Select MDiID From MovimientoDisponibilidad, MovimientoDisponibilidadRenglon" _
                                    & " Where MDIId = MDRIDMovimiento" _
                                    & " And MDRDebe is Not NULL " _
                                    & " And MDiFecha = '" & Format(tFecha.Text, sqlFormatoF) & "'" _
                                    & " And MDRIdDisponibilidad = " & idDisponibilidad _
                                    & " And MDiTipo = " & RsRsm!MDiTipo & " )" _
                        & " Group by DisIDSubrubro"
            
                Set rs1 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                Do While Not rs1.EOF
                    bAdd = False
                    .AddItem ""
                    aSR = rs1!DisIDSubrubro
                    aTexto = RetornoSubrubroFormateado(aSR)
                    .Cell(flexcpText, .Rows - 1, 1) = aTexto
                    
                    .Cell(flexcpText, .Rows - 1, 7) = Format(rs1!Importe, FormatoMonedaP)
                    .Cell(flexcpText, .Rows - 1, 9) = Trim(RsRsm!TMDNombre)
                    
                    rs1.MoveNext
                Loop
                rs1.Close
            End If
        End If
        If bAdd Then
            .AddItem ""
            If Not IsNull(RsRsm!TMDSubRubro) Then aSR = RsRsm!TMDSubRubro Else aSR = 0
            aTexto = RetornoSubrubroFormateado(aSR)
            .Cell(flexcpText, .Rows - 1, 1) = aTexto
            
            .Cell(flexcpText, .Rows - 1, 7) = Format(RsRsm!Importe, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 9) = Trim(RsRsm!TMDNombre)
        End If
        RsRsm.MoveNext
    Loop
    RsRsm.Close
    
    CargoResumenMovimientos = aRetorno
    End With
    
End Function

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
        .FormatString = ">#Ref|<Subrubro|<Documento|<Nº Cheque|>Importe (C) $|>Cofis (C) $|>I.V.A. (C) $|>Importe $|>Importe M/E|<Concepto|"
        .ColWidth(0) = 900: .ColWidth(1) = 2800: .ColWidth(2) = 1200: .ColWidth(3) = 1250
        .ColWidth(4) = 1200: .ColWidth(5) = 950: .ColWidth(6) = 1200: .ColWidth(7) = 1400
         .ColWidth(8) = 1100: .ColWidth(9) = 4400
        
        .WordWrap = False
    End With
      
End Sub


