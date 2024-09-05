VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmListado 
   Caption         =   "Corrijo Stock Virtual"
   ClientHeight    =   6810
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10365
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
   ScaleHeight     =   6810
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   3555
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6271
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
      SubtotalPosition=   0
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
      Height          =   2475
      Left            =   60
      TabIndex        =   18
      Top             =   720
      Width           =   11415
      _Version        =   196608
      _ExtentX        =   20135
      _ExtentY        =   4366
      _StockProps     =   229
      BackColor       =   -2147483634
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
      BackColor       =   -2147483634
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   9975
      TabIndex        =   19
      Top             =   6000
      Width           =   10035
      Begin VB.CommandButton bArreglo 
         Height          =   310
         Left            =   480
         Picture         =   "frmListado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Arreglar"
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pBar 
         Height          =   255
         Left            =   4920
         TabIndex        =   20
         Top             =   120
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":07C4
         Height          =   310
         Left            =   3900
         Picture         =   "frmListado.frx":08C6
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Vista. [Ctrl+L]"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   3540
         Picture         =   "frmListado.frx":0DF8
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2760
         Picture         =   "frmListado.frx":1272
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2400
         Picture         =   "frmListado.frx":135C
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1980
         Picture         =   "frmListado.frx":1446
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3180
         Picture         =   "frmListado.frx":1680
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":1782
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":1884
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1260
         Picture         =   "frmListado.frx":1B86
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1620
         Picture         =   "frmListado.frx":1EC8
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   900
         Picture         =   "frmListado.frx":21CA
         Style           =   1  'Graphical
         TabIndex        =   6
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
      Top             =   6555
      Width           =   10365
      _ExtentX        =   18283
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
            AutoSize        =   1
            Object.Width           =   10081
            Key             =   "msg"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Key             =   "fecha"
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
      Height          =   675
      Left            =   60
      TabIndex        =   16
      Top             =   0
      Width           =   11175
      Begin AACombo99.AACombo cGrupo 
         Height          =   315
         Left            =   6240
         TabIndex        =   3
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
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
      Begin VB.TextBox tArticulo 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "&Grupo:"
         Height          =   255
         Left            =   5640
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private aTexto As String
Private RsSt As rdoResultset

Private Sub AccionLimpiar()
    pBar.Value = 0
    tArticulo.Text = "": cGrupo.Text = ""
    vsConsulta.Rows = 1
End Sub

Private Sub bArreglo_Click()
Dim resp As Integer
    If paCodigoDeUsuario = 0 Then MsgBox "No tengo usuario logueado.", vbExclamation, "ATENCIÓN": Exit Sub
    If vsConsulta.Rows > 1 Then
        If MsgBox("¿Confirma arreglar el stock?", vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then
            resp = MsgBox("¿Para los artículos encontrados desea que la diferencia existente se traslade al ESTADO SANO?" & Chr(13) & "Si presiona <SI> se hará en todos  los artículos." _
                & Chr(13) & "Si presiona <No> no se hará contra el estado sano." & Chr(13) & " Si presiona <Cancel> se consultara para cada artículo de la lista esta condición.", vbQuestion + vbYesNoCancel + vbMsgBoxRight, "ATENCIÓN")
            ArregloStock resp
        End If
    End If
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

Private Sub cGrupo_Click()
    tArticulo.Text = "": vsConsulta.Rows = 1
End Sub

Private Sub cGrupo_Change()
    tArticulo.Text = "": vsConsulta.Rows = 1
End Sub

Private Sub cGrupo_GotFocus()
    With cGrupo
        .SelStart = 0: .SelLength = Len(.Text)
    End With
End Sub

Private Sub cGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
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
    Cons = "Select GruCodigo, GruNombre from Grupo order by GruNombre"
    CargoCombo Cons, cGrupo
    InicializoGrillas
    AccionLimpiar
    vsListado.Orientation = orPortrait: vsListado.BackColor = vbWhite
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub
Private Sub InicializoGrillas()
Dim aValor As Long
  On Error GoTo ErrArmoGrilla
    With vsConsulta
        .Redraw = False
        .WordWrap = False
        .Cols = 1: .Rows = 1
        
        .FormatString = "Artículo|>Tabla A Retirar|>A Retirar |>Fact. Retira |>Tabla A Enviar |>A Envíar |>Fact. Envía |"
        .ColWidth(0) = 2800
        .MergeCol(0) = True
        .Redraw = True
    End With
    Screen.MousePointer = 0
    Exit Sub
ErrArmoGrilla:
    clsGeneral.OcurrioError "Ocurrio un error al armar la grilla.", Err.Description
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

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11

    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
    
    fFiltros.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = fFiltros.Width
    vsListado.Left = fFiltros.Left
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = vsListado.Height
    vsConsulta.Left = vsListado.Left
    
    picBotones.Width = vsListado.Width
    pBar.Width = picBotones.Width - (pBar.Left + 50)
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    GuardoSeteoForm Me
    CierroConexion
    Set clsGeneral = Nothing
    Set miconexion = Nothing
    End
End Sub

Private Sub AccionConsultar()
On Error GoTo errConsultar

    cBase.QueryTimeout = 30
    Screen.MousePointer = 11
    InicializoGrillas
    If vsConsulta.Cols = 2 Then
        MsgBox "No hay datos para los filtros ingresados, verifique.", vbExclamation, "ATENCIÓN"
        Screen.MousePointer = 0: Exit Sub
    End If
    CargoStock
    cBase.QueryTimeout = 15
    Screen.MousePointer = 0
    Exit Sub
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    cBase.QueryTimeout = 15
End Sub

Private Sub Label1_Click()
    Foco cGrupo
End Sub

Private Sub Label4_Click()
    Foco tArticulo
End Sub

Private Sub tArticulo_Change()
    tArticulo.Tag = "0": vsConsulta.Rows = 1: cGrupo.Text = ""
End Sub
Private Sub tArticulo_GotFocus()
    With tArticulo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrAP
    Screen.MousePointer = 11
    If KeyAscii = vbKeyReturn Then
        If Trim(tArticulo.Text) <> "" Then
            If IsNumeric(tArticulo.Text) Then
                BuscoArticuloPorCodigo tArticulo.Text
            Else
                BuscoArticuloPorNombre tArticulo.Text
            End If
            If Val(tArticulo.Tag) > 0 Then vsConsulta.Rows = 1: bConsultar.SetFocus
        Else
            cGrupo.SetFocus
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrAP:
    clsGeneral.OcurrioError "Ocurrio un error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Sub tArticulo_LostFocus()
    Ayuda ""
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    
    If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub
    With vsListado
        .StartDoc
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    
    EncabezadoListado vsListado, "Arreglo de stock virtual", False
    vsListado.FileName = "Arreglo de stock virtual"
    vsListado.Paragraph = ""
    
    vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
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

Private Sub AccionConfigurar()
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
End Sub

Private Sub Ayuda(strTexto As String)
    Status.Panels("msg").Text = strTexto
End Sub
Private Sub ArregloStock(Condicion As Integer)
Dim ContraSano As Boolean
    
    If Condicion = vbYes Then ContraSano = True Else ContraSano = False
    
    Screen.MousePointer = 11
    FechaDelServidor
    
    With vsConsulta
        For I = 1 To .Rows - 1
            
            If .RowHidden(I) = False Then
                
                If Condicion = vbCancel Then
                    If MsgBox("¿Para el artículo " & Trim(.Cell(flexcpText, I, 0)) & " desea que la diferencia existente se traslade al ESTADO SANO?" & Chr(13) & "Si presiona <SI> se hará en todos  los artículos." _
                        , vbQuestion + vbYesNo, "ATENCIÓN") = vbYes Then ContraSano = True Else ContraSano = False
                End If
                
                On Error GoTo errbt
                cBase.BeginTrans
                On Error GoTo errr
                
                'Para el artículo vuelvo a cargar los valores. Por si cambiaron.
                'ARETIRAR
                .Cell(flexcpText, I, 1) = ConsultoStockTotal(Val(.Cell(flexcpData, I, 0)), TipoMovimientoEstado.ARetirar)
                .Cell(flexcpText, I, 2) = StockARetirar(Val(.Cell(flexcpData, I, 0)))
                .Cell(flexcpText, I, 3) = StockAFacturarRetira(Val(.Cell(flexcpData, I, 0)))
                If Val(.Cell(flexcpText, I, 1)) <> Val(.Cell(flexcpText, I, 2)) + Val(.Cell(flexcpText, I, 3)) Then
                    'Hay diferencia entonces lo arreglo.
                    'Si la diferencia es mayor entonces quito del stocktotal.
                    
                    Cons = "Select * from StockTotal Where StTArticulo = " & Val(.Cell(flexcpData, I, 0)) _
                        & " And StTTipoEstado = " & TipoEstadoMercaderia.Virtual _
                        & " And StTEstado = " & TipoMovimientoEstado.ARetirar
                    Set RsSt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If RsSt.EOF Then
                        If Val(.Cell(flexcpText, I, 2)) + Val(.Cell(flexcpText, I, 3)) <> 0 Then
                            RsSt.AddNew
                            RsSt!StTArticulo = Val(.Cell(flexcpData, I, 0))
                            RsSt!StTTipoEstado = TipoEstadoMercaderia.Virtual
                            RsSt!StTEstado = TipoMovimientoEstado.ARetirar
                            RsSt!StTCantidad = Val(.Cell(flexcpText, I, 2)) + Val(.Cell(flexcpText, I, 3))
                            RsSt.Update
                            MarcoMovimientoStockEstado paCodigoDeUsuario, Val(.Cell(flexcpData, I, 0)), Val(.Cell(flexcpText, I, 2)) + Val(.Cell(flexcpText, I, 3)), TipoMovimientoEstado.ARetirar, 1, TipoDocumento.ArregloStock, 1
                        End If
                    Else
                        If RsSt!StTCantidad + ((Val(.Cell(flexcpText, I, 1)) - (Val(.Cell(flexcpText, I, 2)) + Val(.Cell(flexcpText, I, 3)))) * -1) = 0 Then
                            RsSt.Delete
                        Else
                            RsSt.Edit
                            RsSt!StTCantidad = RsSt!StTCantidad + ((Val(.Cell(flexcpText, I, 1)) - (Val(.Cell(flexcpText, I, 2)) + Val(.Cell(flexcpText, I, 3)))) * -1)
                            RsSt.Update
                        End If
                        MarcoMovimientoStockEstado paCodigoDeUsuario, Val(.Cell(flexcpData, I, 0)), ((Val(.Cell(flexcpText, I, 1)) - (Val(.Cell(flexcpText, I, 3)) + Val(.Cell(flexcpText, I, 3)))) * -1), TipoMovimientoEstado.ARetirar, 1, TipoDocumento.ArregloStock, 1
                    End If
                    RsSt.Close
                    If ContraSano Then
                        Cons = "Select * from StockTotal Where StTArticulo = " & Val(.Cell(flexcpData, I, 0)) _
                            & " And StTTipoEstado = " & TipoEstadoMercaderia.Fisico _
                            & " And StTEstado = " & paEstadoArticuloEntrega
                        Set RsSt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        If RsSt.EOF Then
                            If Val(.Cell(flexcpText, I, 2)) + Val(.Cell(flexcpText, I, 3)) <> 0 Then
                                RsSt.AddNew
                                RsSt!StTArticulo = Val(.Cell(flexcpData, I, 0))
                                RsSt!StTTipoEstado = TipoEstadoMercaderia.Fisico
                                RsSt!StTEstado = paEstadoArticuloEntrega
                                RsSt!StTCantidad = (Val(.Cell(flexcpText, I, 2)) + Val(.Cell(flexcpText, I, 3))) * -1
                                RsSt.Update
                            End If
                        Else
                            RsSt.Edit
                            RsSt!StTCantidad = RsSt!StTCantidad + (Val(.Cell(flexcpText, I, 1)) - (Val(.Cell(flexcpText, I, 2)) + Val(.Cell(flexcpText, I, 3))))
                            RsSt.Update
                        End If
                        RsSt.Close
                    End If
                End If
                
                'AENVIAR
                .Cell(flexcpText, I, 4) = ConsultoStockTotal(Val(.Cell(flexcpData, I, 0)), TipoMovimientoEstado.AEntregar)
                .Cell(flexcpText, I, 5) = StockEnEnvio(Val(.Cell(flexcpData, I, 0)))
                .Cell(flexcpText, I, 6) = StockAFacturarEnvia(Val(.Cell(flexcpData, I, 0)))
                If Val(.Cell(flexcpText, I, 4)) <> Val(.Cell(flexcpText, I, 5)) + Val(.Cell(flexcpText, I, 6)) Then
                    'Hay diferencia entonces lo arreglo.
                    'Si la diferencia es mayor entonces quito del stocktotal.
                    
                    Cons = "Select * from StockTotal Where StTArticulo = " & Val(.Cell(flexcpData, I, 0)) _
                        & " And StTTipoEstado = " & TipoEstadoMercaderia.Virtual _
                        & " And StTEstado = " & TipoMovimientoEstado.AEntregar
                    Set RsSt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                    If RsSt.EOF Then
                        If Val(.Cell(flexcpText, I, 5)) + Val(.Cell(flexcpText, I, 6)) <> 0 Then
                            RsSt.AddNew
                            RsSt!StTArticulo = Val(.Cell(flexcpData, I, 0))
                            RsSt!StTTipoEstado = TipoEstadoMercaderia.Virtual
                            RsSt!StTEstado = TipoMovimientoEstado.AEntregar
                            RsSt!StTCantidad = Val(.Cell(flexcpText, I, 5)) + Val(.Cell(flexcpText, I, 6))
                            RsSt.Update
                            MarcoMovimientoStockEstado paCodigoDeUsuario, Val(.Cell(flexcpData, I, 0)), Val(.Cell(flexcpText, I, 5)) + Val(.Cell(flexcpText, I, 6)), TipoMovimientoEstado.AEntregar, 1, TipoDocumento.ArregloStock, 1
                        End If
                    Else
                        If RsSt!StTCantidad + ((Val(.Cell(flexcpText, I, 4)) - (Val(.Cell(flexcpText, I, 5)) + Val(.Cell(flexcpText, I, 6)))) * -1) = 0 Then
                            RsSt.Delete
                        Else
                            RsSt.Edit
                            RsSt!StTCantidad = RsSt!StTCantidad + ((Val(.Cell(flexcpText, I, 4)) - (Val(.Cell(flexcpText, I, 5)) + Val(.Cell(flexcpText, I, 6)))) * -1)
                            RsSt.Update
                        End If
                        MarcoMovimientoStockEstado paCodigoDeUsuario, Val(.Cell(flexcpData, I, 0)), ((Val(.Cell(flexcpText, I, 4)) - (Val(.Cell(flexcpText, I, 5)) + Val(.Cell(flexcpText, I, 6)))) * -1), TipoMovimientoEstado.AEntregar, 1, TipoDocumento.ArregloStock, 1
                    End If
                    RsSt.Close
                    If ContraSano Then
                        Cons = "Select * from StockTotal Where StTArticulo = " & Val(.Cell(flexcpData, I, 0)) _
                            & " And StTTipoEstado = " & TipoEstadoMercaderia.Fisico _
                            & " And StTEstado = " & paEstadoArticuloEntrega
                        Set RsSt = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                        If RsSt.EOF Then
                            If Val(.Cell(flexcpText, I, 5)) + Val(.Cell(flexcpText, I, 6)) <> 0 Then
                                RsSt.AddNew
                                RsSt!StTArticulo = Val(.Cell(flexcpData, I, 0))
                                RsSt!StTTipoEstado = TipoEstadoMercaderia.Fisico
                                RsSt!StTEstado = paEstadoArticuloEntrega
                                RsSt!StTCantidad = (Val(.Cell(flexcpText, I, 5)) + Val(.Cell(flexcpText, I, 6))) * -1
                                RsSt.Update
                            End If
                        Else
                            RsSt.Edit
                            RsSt!StTCantidad = RsSt!StTCantidad + (Val(.Cell(flexcpText, I, 4)) - (Val(.Cell(flexcpText, I, 5)) + Val(.Cell(flexcpText, I, 6))))
                            RsSt.Update
                        End If
                        RsSt.Close
                    End If
                End If
                cBase.CommitTrans
            End If
        Next I
        .Rows = 1
    End With
    
    Screen.MousePointer = 0
    Exit Sub
errbt:
    clsGeneral.OcurrioError "Ocurrio un error al iniciar la transacción", Err.Description
    Screen.MousePointer = 0
    Exit Sub
errr:
    Resume rolb
rolb:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Ocurrio un error al grabar", Err.Description
    Screen.MousePointer = 0
    
End Sub

Private Sub CargoStock()
On Error GoTo ErrCS
Dim idArticulo As Long, aValor As Long
    
    Screen.MousePointer = 11
    'Recorro la fila 0 y saco los distintos conteos que tengo.
    'Para cada uno veo si ya inserte el artículo y estado, entonces pongo su cantidad
    'al final recorro nuevamente y marco los que poseen diferencias.
    Cons = "Select count(*) from Articulo "
                        
    If Val(tArticulo.Tag) > 0 Then Cons = Cons & " Where ArtID = " & Val(tArticulo.Tag)
    If cGrupo.ListIndex > -1 Then Cons = Cons & " Where ArtID IN (Select AGrArticulo From ArticuloGrupo Where AGrGrupo = " & cGrupo.ItemData(cGrupo.ListIndex) & ")"

    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    pBar.Max = RsAux(0)
    RsAux.Close
    
    If pBar.Max = 0 Then MsgBox "No hay datos a desplegar.", vbExclamation, "ATENCIÓN"
    
    idArticulo = 0
    With vsConsulta
        
        Cons = "Select * from Articulo "
                        
        If Val(tArticulo.Tag) > 0 Then Cons = Cons & " Where ArtID = " & Val(tArticulo.Tag)
        If cGrupo.ListIndex > -1 Then Cons = Cons & " Where ArtID IN (Select AGrArticulo From ArticuloGrupo Where AGrGrupo = " & cGrupo.ItemData(cGrupo.ListIndex) & ")"

        Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
        
        Do While Not RsAux.EOF
            
            .AddItem Format(RsAux!ArtCodigo, "#,000,000") & " " & Trim(RsAux!ArtNombre)
            idArticulo = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = idArticulo
            idArticulo = 0: .Cell(flexcpData, .Rows - 1, 1) = idArticulo
            
            .Cell(flexcpText, .Rows - 1, 1) = ConsultoStockTotal(RsAux!ArtID, TipoMovimientoEstado.ARetirar)
            .Cell(flexcpText, .Rows - 1, 2) = StockARetirar(RsAux!ArtID)
            .Cell(flexcpText, .Rows - 1, 3) = StockAFacturarRetira(RsAux!ArtID)
            If Val(.Cell(flexcpText, .Rows - 1, 1)) <> Val(.Cell(flexcpText, .Rows - 1, 2)) + Val(.Cell(flexcpText, .Rows - 1, 3)) Then
                .Cell(flexcpBackColor, .Rows - 1, 1) = Colores.obligatorio
                idArticulo = 1: .Cell(flexcpData, .Rows - 1, 1) = idArticulo
            End If
            
            .Cell(flexcpText, .Rows - 1, 4) = ConsultoStockTotal(RsAux!ArtID, TipoMovimientoEstado.AEntregar)
            .Cell(flexcpText, .Rows - 1, 5) = StockEnEnvio(RsAux!ArtID)
            .Cell(flexcpText, .Rows - 1, 6) = StockAFacturarEnvia(RsAux!ArtID)
            If Val(.Cell(flexcpText, .Rows - 1, 4)) <> Val(.Cell(flexcpText, .Rows - 1, 5)) + Val(.Cell(flexcpText, .Rows - 1, 6)) Then
                .Cell(flexcpBackColor, .Rows - 1, 4) = Colores.obligatorio
                idArticulo = 1: .Cell(flexcpData, .Rows - 1, 1) = idArticulo
            End If
            RsAux.MoveNext
            pBar.Value = pBar.Value + 1
            If .Cell(flexcpData, .Rows - 1, 1) = 0 Then .RowHidden(.Rows - 1) = True
        Loop
        RsAux.Close
    End With
    pBar.Value = 0
    Screen.MousePointer = 0
    Exit Sub
ErrCS:
    clsGeneral.OcurrioError "Ocurrio un error al cargar el stock.", Err.Description
    Screen.MousePointer = 0
End Sub
Private Function EstaInsertado(idArticulo As Long, idEstado As Long)
Dim Cont As Integer
    EstaInsertado = 0
    With vsConsulta
        For Cont = 1 To .Rows - 1
            If Val(.Cell(flexcpData, Cont, 0)) = idArticulo And Val(.Cell(flexcpData, Cont, 1)) = idEstado Then EstaInsertado = Cont: Exit For
        Next Cont
    End With
End Function

Private Sub BuscoArticuloPorCodigo(CodArticulo As Long)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
    
    Screen.MousePointer = 11
    Cons = "Select * From Articulo Where ArtCodigo = " & CodArticulo
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    If RsAux.EOF Then
        RsAux.Close
        tArticulo.Tag = "0"
        MsgBox "No existe un artículo que posea ese código.", vbExclamation, "ATENCIÓN"
    Else
        tArticulo.Text = Format(RsAux!ArtCodigo, "#,000,000") & " " & Trim(RsAux!ArtNombre)
        tArticulo.Tag = RsAux!ArtID
        RsAux.Close
    End If
    Screen.MousePointer = 0

End Sub

Private Sub BuscoArticuloPorNombre(NomArticulo As String)
'Atención el mapeo de error lo hago antes de entrar al procedimiento
Dim Resultado As Long

    Screen.MousePointer = 11
    
    Cons = "Select ArtId, Código = ArtCodigo, Nombre = ArtNombre from Articulo" _
        & " Where ArtNombre LIKE '" & NomArticulo & "%'" _
        & " Order By ArtNombre"
            
    Dim LiAyuda As New clsListadeAyuda
    LiAyuda.ActivoListaAyuda Cons, False, cBase.Connect
    Screen.MousePointer = 11
    If LiAyuda.ItemSeleccionado <> "" Then
        Resultado = LiAyuda.ItemSeleccionado
    Else
        Resultado = 0
    End If
    If Resultado > 0 Then BuscoArticuloPorCodigo Resultado
    Set LiAyuda = Nothing       'Destruyo la clase.
    Screen.MousePointer = 0
    
End Sub

Private Function StockAFacturarEnvia(Articulo As Long) As Long
    Cons = "Select Sum(REvAEntregar) From Envio, RenglonEnvio" _
        & " Where EnvTipo = " & TipoEnvio.Cobranza _
        & " And EnvEstado NOT IN (2,4,5)" _
        & " And EnvDocumento <> Null " _
        & " And REvArticulo = " & Articulo _
        & " And REvAEntregar > 0" _
        & "And EnvCodigo = REvEnvio"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If IsNull(RsSt(0)) Then StockAFacturarEnvia = 0 Else StockAFacturarEnvia = RsSt(0)
    RsSt.Close
End Function

Private Function StockAFacturarRetira(Articulo As Long) As Long
    Cons = "Select Sum(RVTARetirar) From VentaTelefonica, RenglonVtaTelefonica " _
            & " Where VTeTipo = " & TipoDocumento.ContadoDomicilio _
            & " And VTeDocumento = Null And VTeAnulado = Null " _
            & " And RVTArticulo = " & Articulo _
            & "And VTeCodigo = RVTVentaTelefonica"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If IsNull(RsSt(0)) Then StockAFacturarRetira = 0 Else StockAFacturarRetira = RsSt(0)
    RsSt.Close
End Function

Private Function StockEnEnvio(Articulo As Long) As Long
    Cons = "Select Sum(REvAEntregar) From Envio, RenglonEnvio" _
        & " Where EnvTipo = " & TipoEnvio.Entrega _
        & " And EnvEstado NOT IN (2,4,5)" _
        & " And EnvDocumento <> Null And REvArticulo = " & Articulo _
        & " And REvAEntregar <> 0" _
        & "And EnvCodigo = REvEnvio"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If IsNull(RsSt(0)) Then StockEnEnvio = 0 Else StockEnEnvio = RsSt(0)
    RsSt.Close
End Function

Private Function StockARetirar(Articulo As Long) As Long
    StockARetirar = 0
    Cons = "Select Sum(RenARetirar) From Documento, Renglon " _
        & " Where DocTipo IN (" & TipoDocumento.Contado & ", " & TipoDocumento.Credito & ") " _
        & " And RenArticulo = " & Articulo _
        & " And RenDocumento = DocCodigo And DocAnulado = 0"
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not IsNull(RsSt(0)) Then StockARetirar = RsSt(0)
    RsSt.Close
    'Tambien hay en Remitos.
    Cons = "Select Sum(RReAEntregar) From Remito, RenglonRemito " _
            & " Where RReArticulo = " & Articulo _
            & " And RReRemito = RemCodigo "
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not IsNull(RsSt(0)) Then StockARetirar = StockARetirar + RsSt(0)
    RsSt.Close
    Exit Function
End Function

Private Function ConsultoStockTotal(Articulo As Long, Estado As Integer)
    ConsultoStockTotal = 0
    Cons = "Select * from StockTotal Where StTArticulo = " & Articulo _
        & " And StTTipoEstado = " & TipoEstadoMercaderia.Virtual _
        & " And StTEstado = " & Estado
    Set RsSt = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not RsSt.EOF Then ConsultoStockTotal = RsSt!StTCantidad
    RsSt.Close
End Function
