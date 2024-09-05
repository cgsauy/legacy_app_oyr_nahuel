VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmListado 
   Caption         =   "Movimiento de Disponibilidades"
   ClientHeight    =   5820
   ClientLeft      =   1995
   ClientTop       =   2520
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
      Begin VB.CommandButton bInteres 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Cálculo de Intereses."
         Top             =   120
         Width           =   310
      End
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
         Left            =   3120
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
         Left            =   2760
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
         Left            =   2100
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
         Left            =   1380
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
         Left            =   1740
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
         Left            =   1020
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

Dim aTexto As String
Dim cSaldoInicial As Currency, strFechaSaldo As String, strHoraSaldo As String
Dim gMonedaD As Long
Dim bBancaria As Boolean

Dim prmQSaldos As Long, prmSumaSaldos As Currency

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

Private Sub bInteres_Click()
    AccionFrmIntereses
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
        AccionImprimir
        vsListado.ZOrder 0
        Exit Sub
        With vsListado          'Selecciono Listado
            Screen.MousePointer = 11
            If vsGrilla.ColHidden(8) Then
                vsGrilla.ColWidth(9) = 3000: .Orientation = orPortrait
            Else
                vsGrilla.ColWidth(9) = 3600: .Orientation = orLandscape
            End If
            vsGrilla.ExtendLastCol = False
            
            .StartDoc
            EncabezadoListado vsListado, "Movimiento de Disponibilidades desde " & tDesde.Text & " hasta " & tHasta.Text, True
            .RenderControl = vsGrilla.hwnd
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
            Case vbKeyI: AccionImprimir True
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
    cons = "Select DisID, DisNombre From NivelPermiso, Disponibilidad " _
        & " Where NPeNivel IN (Select UNiNivel From UsuarioNivel Where UNiUsuario = " & paCodigoDeUsuario & ")" _
        & " And NPeAplicacion = DisAplicacion" _
        & " Group by DisID, DisNombre" _
        & " Order by DisNombre"
    CargoCombo cons, cDisponibilidad
    BuscoCodigoEnCombo cDisponibilidad, paDisponibilidad
    '--------------------------------------------------------------
    tDesde.Text = Format(Now, FormatoFP)
    tHasta.Text = Format(Now, FormatoFP)
    
    vsGrilla.ColHidden(7) = True: vsGrilla.ColHidden(8) = True
    pbProgreso.Value = 0
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error inesperado al cargar el formulario.", Err.Description
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
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Private Sub Label1_Click()
    Foco tDesde
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    Screen.MousePointer = 11
    Me.Refresh
    
    vsGrilla.ExtendLastCol = False
    With vsListado
         .Orientation = orLandscape
        If vsGrilla.ColHidden(8) Then vsGrilla.ColWidth(9) = 4900 Else vsGrilla.ColWidth(9) = 2800
        
        .StartDoc
        .FileName = "Movimiento de Disponibilidades"
        aTexto = "Movimiento de Disponibilidades": If cDisponibilidad.Text <> "" Then aTexto = aTexto & " (" & Trim(cDisponibilidad.Text) & ") "
        aTexto = aTexto & " desde " & tDesde.Text & " hasta " & tHasta.Text
        
        EncabezadoListado vsListado, aTexto, True
        .RenderControl = vsGrilla.hwnd
        .EndDoc
    
    End With
    vsGrilla.ExtendLastCol = True
    
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
        .PreviewMode = pmPrinter
        .PreviewPage = 1
        .FontName = "Tahoma": .FontSize = 10: .FontBold = False: .FontItalic = False
        .TextAlign = 0: .PageBorder = 3
        .Columns = 1
        .TableBorder = tbBoxRows
        .Zoom = 100
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

Private Sub vsGrilla_DblClick()

    If Not bBancaria Then Exit Sub
    With vsGrilla
        If .Row <= 1 Then Exit Sub
        If .Cell(flexcpText, .Row, 0) = "" Then Exit Sub
        If Not IsDate(.Cell(flexcpText, .Row, 0)) Then Exit Sub
        
        If .Cell(flexcpForeColor, .Row, 0) = vbBlack Then
            .Cell(flexcpForeColor, .Row, 0, , .Cols - 1) = Colores.RojoClaro: .Cell(flexcpBackColor, .Row, 0, , .Cols - 1) = Colores.Gris
            
            If .Cell(flexcpText, .Row, 4) <> "" Then
                .Cell(flexcpText, .Rows - 1, 4) = Format(.Cell(flexcpText, .Rows - 1, 4) - .Cell(flexcpValue, .Row, 4), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpText, .Rows - 1, 6) - .Cell(flexcpValue, .Row, 4), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 8) = Format(.Cell(flexcpText, .Rows - 1, 8) - Abs(.Cell(flexcpValue, .Row, 7)), FormatoMonedaP)
            Else
                .Cell(flexcpText, .Rows - 1, 5) = Format(.Cell(flexcpText, .Rows - 1, 5) - .Cell(flexcpValue, .Row, 5), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpText, .Rows - 1, 6) + .Cell(flexcpValue, .Row, 5), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 8) = Format(.Cell(flexcpText, .Rows - 1, 8) + Abs(.Cell(flexcpValue, .Row, 7)), FormatoMonedaP)
            End If
    
        Else
            .Cell(flexcpForeColor, .Row, 0, , .Cols - 1) = vbBlack: .Cell(flexcpBackColor, .Row, 0, , .Cols - 1) = Colores.Blanco
            
            If .Cell(flexcpText, .Row, 4) <> "" Then
                .Cell(flexcpText, .Rows - 1, 4) = Format(.Cell(flexcpText, .Rows - 1, 4) + .Cell(flexcpValue, .Row, 4), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpText, .Rows - 1, 6) + .Cell(flexcpValue, .Row, 4), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 8) = Format(.Cell(flexcpText, .Rows - 1, 8) + Abs(.Cell(flexcpValue, .Row, 7)), FormatoMonedaP)
            Else
                .Cell(flexcpText, .Rows - 1, 5) = Format(.Cell(flexcpText, .Rows - 1, 5) + .Cell(flexcpValue, .Row, 5), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpText, .Rows - 1, 6) - .Cell(flexcpValue, .Row, 5), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 8) = Format(.Cell(flexcpText, .Rows - 1, 8) - Abs(.Cell(flexcpValue, .Row, 7)), FormatoMonedaP)
            End If

        End If
        .Cell(flexcpText, .Rows - 1, 9) = Format(.Cell(flexcpValue, .Rows - 1, 6) * .Cell(flexcpData, .Rows - 1, 0), FormatoMonedaP)
        
    End With
    
End Sub

Private Sub vsGrilla_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then Call vsGrilla_DblClick
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionConsultar()

    On Error GoTo ErrCDML
    
    If Not ValidoDatos Then Exit Sub
    Screen.MousePointer = 11
    
    prmQSaldos = 0: prmSumaSaldos = 0
    
    vsGrilla.ZOrder 0: vsGrilla.Rows = 1: vsGrilla.Refresh
    chVista.Value = 0
    
    'Busco el si hay un saldo inicial para esa disponibilidad.----------------------------------------------------------------------------------
    cSaldoInicial = 0: strFechaSaldo = "": strHoraSaldo = ""
    
    cons = "Select * From SaldoDisponibilidad " _
        & " Where SDiDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
        & " And SDiFecha = (Select MAX(SDiFecha) From SaldoDisponibilidad " _
            & " Where SDiDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
            & " And SDiFecha <= '" & Format(tDesde.Text & " " & "23:59:59", sqlFormatoFH) & "')"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurReadOnly)
    If Not rsAux.EOF Then
        cSaldoInicial = rsAux!SDiSaldo
        strFechaSaldo = rsAux!SDiFecha
        strHoraSaldo = rsAux!SDiHora
    Else
        strFechaSaldo = tDesde.Text
        strHoraSaldo = "00:00:00"
    End If
    rsAux.Close
    '----------------------------------------------------------------------------------------------------------------------------------------------
    
    'Busco la moneda de la disponibilidad (para habilitar columnas de pesos)-----------------------------------------------------------
    cons = "Select * From Disponibilidad Where DisId = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurReadOnly)
    gMonedaD = rsAux!DisMoneda
    If Not IsNull(rsAux!DisSucursal) Then bBancaria = True Else bBancaria = False
    rsAux.Close
    If gMonedaD = paMonedaPesos Then
        vsGrilla.ColHidden(7) = True: vsGrilla.ColHidden(8) = True
    Else
        vsGrilla.ColHidden(7) = False: vsGrilla.ColHidden(8) = False
    End If
    vsGrilla.Refresh
    '----------------------------------------------------------------------------------------------------------------------------------------------
    
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    pbProgreso.Value = 0
    cons = "Select Count(*) from  MovimientoDisponibilidad, MovimientoDisponibilidadRenglon" _
                & " Where MDRIDDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
                & " And ((MDiFecha = '" & Format(strFechaSaldo, sqlFormatoF) & "' And MDiHora >= '" & strHoraSaldo & "')" _
                & " Or  MDiFecha > '" & Format(strFechaSaldo, sqlFormatoF) & "') " _
                & " And MDiFecha <= '" & Format(tHasta.Text, sqlFormatoF) & "'" _
                & " And MDIId = MDRIDMovimiento "
                
    If bBancaria Then
        'Saco los Movimientos con Cheques Diferidos entre las fechas
        cons = cons & " UNION ALL " _
            & "Select Count(*) From  MovimientoDisponibilidad, MovimientoDisponibilidadRenglon, Cheque " _
                & " Where MDRIDDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
                & " And ((MDiFecha = '" & Format(strFechaSaldo, sqlFormatoF) & "' And MDiHora < '" & strHoraSaldo & "')" _
                        & " Or  MDiFecha < '" & Format(strFechaSaldo, sqlFormatoF) & "') " _
                & " And CheVencimiento Between '" & Format(strFechaSaldo, sqlFormatoF) & "' AND '" & Format(tHasta.Text, sqlFormatoF) & "'" _
                & " And MDIId = MDRIDMovimiento " _
                & " And MDRIdCheque = CheId "
    End If
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        Dim aCantidad As Long: aCantidad = 0
        Do While Not rsAux.EOF
            If rsAux(0) <> 0 Then aCantidad = aCantidad + rsAux(0)
            rsAux.MoveNext
        Loop
    End If
    rsAux.Close
    If aCantidad = 0 Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        Screen.MousePointer = 0: Exit Sub
    End If
    pbProgreso.Max = aCantidad
    '-----------------------------------------------------------------------------------------------------------------
       
    CargoDatos
    
    If bBancaria Then zJuntoDepositos
    
    Screen.MousePointer = 0
    Exit Sub
ErrCDML:
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos.", Err.Description
    vsGrilla.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub CargoDatos()

Dim aTC As Currency, aSaldoPesos As Currency        'Columnas en pesos
Dim rs1 As rdoResultset
Dim aIDCheque As Long
Dim aFInicio As String

    cons = "Select * From  MovimientoDisponibilidad " _
                        & " Left Outer Join Compra On MDiIdCompra = ComCodigo" _
                        & " Left Outer Join ProveedorCliente On ComProveedor = PClCodigo " _
                & ", MovimientoDisponibilidadRenglon left Outer Join Cheque On  MDRIdCheque = CheId " _
                & " Where MDRIDDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
                & " And ((MDiFecha = '" & Format(strFechaSaldo, sqlFormatoF) & "' And MDiHora >= '" & strHoraSaldo & "')" _
                        & " Or  MDiFecha > '" & Format(strFechaSaldo, sqlFormatoF) & "') " _
                & " And MDiFecha <= '" & Format(tHasta.Text, sqlFormatoF) & "'" _
                & " And MDIId = MDRIDMovimiento "
    
    If bBancaria Then
        'Saco los Movimientos con Cheques Diferidos entre las fechas
        cons = cons & " UNION ALL " _
                & "Select * From  MovimientoDisponibilidad " _
                        & " Left Outer Join Compra On MDiIdCompra = ComCodigo" _
                        & " Left Outer Join ProveedorCliente On ComProveedor = PClCodigo " _
                & ", MovimientoDisponibilidadRenglon, Cheque " _
                & " Where MDRIDDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
                & " And ((MDiFecha = '" & Format(strFechaSaldo, sqlFormatoF) & "' And MDiHora < '" & strHoraSaldo & "')" _
                        & " Or  MDiFecha < '" & Format(strFechaSaldo, sqlFormatoF) & "') " _
                & " And CheVencimiento Between '" & Format(strFechaSaldo, sqlFormatoF) & "' AND '" & Format(tHasta.Text, sqlFormatoF) & "'" _
                & " And MDIId = MDRIDMovimiento " _
                & " And MDRIdCheque = CheId "
    
        cons = cons & " Order by MDRIDCheque"
    Else
        cons = cons & " Order by MDiFecha, MDiHora"
    End If
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurReadOnly)
    If rsAux.EOF Then rsAux.Close: Exit Sub
    
    vsGrilla.Redraw = False

    aIDCheque = 0
    'Hay que sacar la TC a Pesos si la moneda es distinta
    If cSaldoInicial <> 0 And gMonedaD <> paMonedaPesos Then
        aTC = TasadeCambio(CInt(gMonedaD), paMonedaPesos, CDate(strFechaSaldo))
        aSaldoPesos = cSaldoInicial * aTC
    End If
    
    If cSaldoInicial >= 0 And strFechaSaldo <> "" Then
        With vsGrilla
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(strFechaSaldo, "dd/mm/yy hh:mm")
            .Cell(flexcpText, .Rows - 1, 6) = Format(cSaldoInicial, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 8) = Format(aSaldoPesos, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 9) = "Saldo Inicial"
        End With
        prmQSaldos = prmQSaldos + 1
        prmSumaSaldos = prmSumaSaldos + cSaldoInicial
        
    ElseIf cSaldoInicial < 0 And strFechaSaldo <> "" Then
        With vsGrilla
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(strFechaSaldo, "dd/mm/yy hh:mm")
            .Cell(flexcpText, .Rows - 1, 6) = Format(cSaldoInicial, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 8) = Format(aSaldoPesos, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 9) = "Saldo Inicial"
        End With
        prmQSaldos = prmQSaldos + 1
        prmSumaSaldos = prmSumaSaldos + cSaldoInicial
    End If
    
    If strFechaSaldo <> "" Then aFInicio = Format(strFechaSaldo, "yyyy/mm/dd") Else aFInicio = Format(tDesde.Text, "yyyy/mm/dd")
    With vsGrilla
    
    Do While Not rsAux.EOF
        pbProgreso.Value = pbProgreso.Value + 1
        
        If aIDCheque <> rsAux!MDRIdCheque Or rsAux!MDRIdCheque = 0 Then
            
            If Not IsNull(rsAux!CheVencimiento) Then            'Cheuqe Diferido
                If Format(rsAux!CheVencimiento, "yyyy/mm/dd") < CDate(aFInicio) Or _
                   Format(rsAux!CheVencimiento, "yyyy/mm/dd") > Format(tHasta.Text, "yyyy/mm/dd") Then GoTo Siguiente
            End If
            
            aIDCheque = rsAux!MDRIdCheque
            .AddItem ""
            Select Case rsAux!MDiFecha
                Case Is < CDate(aFInicio): .Cell(flexcpText, .Rows - 1, 0) = Format(strFechaSaldo, "dd/mm/yy hh:mm")
                Case Is > CDate(aFInicio): .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!MDiFecha, "dd/mm/yy") & " " & Format(rsAux!MDiHora, "hh:mm")
                Case Is = CDate(aFInicio): If rsAux!MDiHora < strHoraSaldo Then .Cell(flexcpText, .Rows - 1, 0) = Format(strFechaSaldo, "dd/mm/yy hh:mm") Else .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!MDiFecha, "dd/mm/yy") & " " & Format(rsAux!MDiHora, "hh:mm")
            End Select
            
            If Not IsNull(rsAux!CheSerie) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!CheSerie) & " " & Trim(rsAux!CheNumero)
            If Not IsNull(rsAux!CheVencimiento) Then .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!CheVencimiento, "dd/mm/yy")
            
            If Not IsNull(rsAux!PClFantasia) Then
                .Cell(flexcpText, .Rows - 1, 3) = Trim(rsAux!PClFantasia)
            Else
                'Veo si es transferencia---------------------------------------------------------------------------------------
                cons = " Select * from MovimientoDisponibilidad, MovimientoDisponibilidadRenglon, TipoMovDisponibilidad, Disponibilidad " _
                        & " Where MDiID = MDRIDMovimiento " _
                        & " And MDRidMovimiento = " & rsAux!MDiID _
                        & " And MDRIdDisponibilidad <> " & rsAux!MDRIdDisponibilidad _
                        & " And MDRIdDisponibilidad = DisID And MDITipo = TMDCodigo"
                    
                Set rs1 = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                If Not rs1.EOF Then
                    If Not IsNull(rs1!TMDTransferencia) Then
                        If rs1!TMDTransferencia = 1 Then
                            If Not rs1.EOF Then .Cell(flexcpText, .Rows - 1, 3) = Trim(rs1!DisNombre)
                        End If
                    End If
                End If
                rs1.Close
                '-------------------------------------------------------------------------------------------------------------------
            End If
            
            If Not IsNull(rsAux!MDRDebe) Then
                .Cell(flexcpText, .Rows - 1, 4) = Format(Abs(rsAux!MDRDebe), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 7) = Format(Abs(rsAux!MDRImportePesos), FormatoMonedaP)
            End If
            
            If Not IsNull(rsAux!MDRHaber) Then
                .Cell(flexcpText, .Rows - 1, 5) = Format(Abs(rsAux!MDRHaber), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 7) = Format(Abs(rsAux!MDRImportePesos) * -1, FormatoMonedaP)
            End If
            
            If Not IsNull(rsAux!ComComentario) Then     '07/05/2003 Cambie el orden x los movs de cheques diferidos
                .Cell(flexcpText, .Rows - 1, 9) = Trim(rsAux!ComComentario)
            Else
                If Not IsNull(rsAux!MDiComentario) Then .Cell(flexcpText, .Rows - 1, 9) = Trim(rsAux!MDiComentario)
            End If
            If Not IsNull(rsAux!MDiIdCompra) Then
                .Cell(flexcpText, .Rows - 1, 9) = "(G:" & Format(rsAux!MDiIdCompra, "0") & ") " & .Cell(flexcpText, .Rows - 1, 9)
            Else
                If Not IsNull(rsAux!MDiID) Then .Cell(flexcpText, .Rows - 1, 9) = Format(rsAux!MDiID, "(0) ") & .Cell(flexcpText, .Rows - 1, 9)
            End If
        
        Else    'El cheque ya está ingresado
            .Cell(flexcpText, .Rows - 1, 3) = ""
            If Not IsNull(rsAux!MDRDebe) Then
                    .Cell(flexcpText, .Rows - 1, 4) = Format(.Cell(flexcpValue, .Rows - 1, 4) + Abs(rsAux!MDRDebe), FormatoMonedaP)
                    .Cell(flexcpText, .Rows - 1, 7) = Format(.Cell(flexcpValue, .Rows - 1, 7) + Abs(rsAux!MDRImportePesos), FormatoMonedaP)
            End If
        
            If Not IsNull(rsAux!MDRHaber) Then
                .Cell(flexcpText, .Rows - 1, 5) = Format(.Cell(flexcpValue, .Rows - 1, 5) + Abs(rsAux!MDRHaber), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 7) = Format(Abs(Abs(.Cell(flexcpValue, .Rows - 1, 7)) + Abs(rsAux!MDRImportePesos)) * -1, FormatoMonedaP)
            End If
        End If

Siguiente:
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    'Ordeno por fecha y Cargo los campos de saldos     -------------------------------------------------------------------------------------
    If bBancaria Then
        .Select 1, 0, 1, 2
        .Sort = flexSortGenericAscending
    End If
    
    For I = 2 To .Rows - 1
        If .Cell(flexcpText, I, 4) <> "" Then   'DEBE
            .Cell(flexcpText, I, 6) = Format(.Cell(flexcpValue, I - 1, 6) + .Cell(flexcpValue, I, 4), FormatoMonedaP)
        Else    'HABER
            .Cell(flexcpText, I, 6) = Format(.Cell(flexcpValue, I - 1, 6) - .Cell(flexcpValue, I, 5), FormatoMonedaP)
        End If
        prmSumaSaldos = prmSumaSaldos + .Cell(flexcpValue, I, 6)
        prmQSaldos = prmQSaldos + 1
        .Cell(flexcpText, I, 8) = Format(.Cell(flexcpValue, I - 1, 8) + .Cell(flexcpValue, I, 7), FormatoMonedaP)
    Next
    '------------------------------------------------------------------------------------------------------------------------------------------------
    
    .AddItem ""
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 4, , RGB(128, 128, 128), Colores.Blanco, True, "Total"
    .Subtotal flexSTSum, -1, 5, , RGB(128, 128, 128), Colores.Blanco, True, "Total"
    .Cell(flexcpText, .Rows - 1, 6) = .Cell(flexcpText, .Rows - 3, 6): .Cell(flexcpText, .Rows - 1, 8) = .Cell(flexcpText, .Rows - 3, 8)
    .Cell(flexcpText, .Rows - 1, 9) = " A TC del " & tHasta.Text
    If bBancaria Then
        .AddItem "Conciliación"
        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = RGB(105, 105, 105): .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco: .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
        .Cell(flexcpText, .Rows - 1, 4) = .Cell(flexcpText, .Rows - 2, 4): .Cell(flexcpText, .Rows - 1, 5) = .Cell(flexcpText, .Rows - 2, 5)
        .Cell(flexcpText, .Rows - 1, 6) = .Cell(flexcpText, .Rows - 2, 6): .Cell(flexcpText, .Rows - 1, 8) = .Cell(flexcpText, .Rows - 2, 8)
    End If
    
    'Pesos corregidos a la fecha hasta
    .Cell(flexcpData, .Rows - 1, 0) = TasadeCambio(CInt(gMonedaD), paMonedaPesos, CDate(tHasta.Text))
    .Cell(flexcpText, .Rows - 1, 9) = Format(.Cell(flexcpValue, .Rows - 1, 6) * .Cell(flexcpData, .Rows - 1, 0), FormatoMonedaP)
    
    End With
    pbProgreso.Value = 0: vsGrilla.Redraw = True
    
End Sub

Private Sub AccionLimpiar()

    On Error Resume Next
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
        .ColWidth(0) = 1250: .ColWidth(1) = 1000: .ColWidth(2) = 750: .ColWidth(3) = 1900: .ColWidth(4) = 1400: .ColWidth(5) = 1400:: .ColWidth(6) = 1500
        .ColWidth(7) = 1400: .ColWidth(8) = 1500
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

Private Sub AccionFrmIntereses()

    If vsGrilla.Rows = 1 Then Exit Sub
    
    On Error GoTo errInteres
    Screen.MousePointer = 11
    Dim prmInteresesG As Currency
    prmInteresesG = 0
    
    'Proceso Intereses Ganados  --------------------------------------------------------------------------------------------------------
     cons = "Select * From  MovimientoDisponibilidad, MovimientoDisponibilidadRenglon, GastoSubRubro " _
                & " Where MDRIDDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
                & " And ((MDiFecha = '" & Format(strFechaSaldo, sqlFormatoF) & "' And MDiHora >= '" & strHoraSaldo & "')" _
                        & " Or  MDiFecha > '" & Format(strFechaSaldo, sqlFormatoF) & "') " _
                & " And MDiFecha <= '" & Format(tHasta.Text, sqlFormatoF) & "'" _
                & " And MDiId = MDRIDMovimiento " _
                & " And MDiIDCompra = GSrIdCompra And GSrIDSubRubro = " & paSubrubroIntBanGan

    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        If Not IsNull(rsAux!MDRDebe) Then prmInteresesG = prmInteresesG + Format(Abs(rsAux!MDRDebe), FormatoMonedaP)
            
        If Not IsNull(rsAux!MDRHaber) Then prmInteresesG = prmInteresesG + Format(Abs(rsAux!MDRHaber) * -1, FormatoMonedaP)
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    '------------------------------------------------------------------------------------------------------------------------------------------

    Screen.MousePointer = 0
    
    frmInteres.prmQSaldos = prmQSaldos
    frmInteres.prmSumaSaldos = prmSumaSaldos
    frmInteres.prmIGanados = prmInteresesG
    frmInteres.Show vbModal, Me
    
    Exit Sub

errInteres:
    clsGeneral.OcurrioError "Error al calcular los intereses ganados.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function zJuntoDepositos()

On Error GoTo errFncZ

Dim mPatron As String, mLetter As String, mDay As String
Dim mTotal As Currency, mRow1 As Integer

    mPatron = "*Depósito de Ch.*"
    mDay = "00/00/0000"
    mLetter = ""
    
    With vsGrilla
        For I = 1 To .Rows - 1
            If .Cell(flexcpText, I, 9) Like mPatron Then
                If mDay = Format(.Cell(flexcpText, I, 0), "dd/mm/yyyy") Then
                    .Cell(flexcpText, I, 1) = "(" & mLetter & ") "
                    mTotal = mTotal + .Cell(flexcpValue, I, 4)
                Else
                    If mDay <> "00/00/0000" Then
                        .Cell(flexcpText, mRow1, 1) = "(" & mLetter & ") " & mTotal
                    End If
                    
                    If mLetter <> "" Then
                        mLetter = Chr(Asc(mLetter) + 1)
                    Else
                        mLetter = "A"
                    End If
                    
                    mDay = Format(.Cell(flexcpText, I, 0), "dd/mm/yyyy")
                    mTotal = .Cell(flexcpValue, I, 4)
                    mRow1 = I
                    
                End If
            End If
        
        Next I
        
        If mTotal <> 0 And mRow1 Then
            .Cell(flexcpText, mRow1, 1) = "(" & mLetter & ") " & mTotal
        End If
        
    End With

errFncZ:
End Function

