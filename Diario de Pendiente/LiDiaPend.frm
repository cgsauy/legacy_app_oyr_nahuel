VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{191D08B9-4E92-4372-BF17-417911F14390}#1.5#0"; "orGridPreview.ocx"
Begin VB.Form frmListado 
   Caption         =   "Diario de Pendiente Contado"
   ClientHeight    =   3870
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LiDiaPend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin orGridPreview.GridPreview gpPrint 
      Left            =   5760
      Top             =   1440
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsDetalle 
      Height          =   1335
      Left            =   5880
      TabIndex        =   16
      Top             =   1080
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2355
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "Camion       |<       Fecha|>      $ Pendiente "
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrilla 
      Height          =   1095
      Left            =   1440
      TabIndex        =   9
      Top             =   840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1931
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
      BackColorFixed  =   -2147483636
      ForeColorFixed  =   -2147483634
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   14737632
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
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
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros"
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   11655
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   7200
         TabIndex        =   15
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
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
      Begin VB.TextBox tHasta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox tDesde 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin AACombo99.AACombo cSucursal 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
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
      Begin AACombo99.AACombo cboCamion 
         Height          =   315
         Left            =   8880
         TabIndex        =   18
         Top             =   240
         Width           =   2355
         _ExtentX        =   4154
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
         Caption         =   "&Camión:"
         Height          =   255
         Left            =   8160
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "&Moneda:"
         Height          =   255
         Left            =   6480
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "&Pendientes al:"
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   4200
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Sucursal:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   6075
      TabIndex        =   6
      Top             =   3000
      Width           =   6135
      Begin VB.CommandButton bPreview 
         Height          =   310
         Left            =   960
         Picture         =   "LiDiaPend.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Preview."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   600
         Picture         =   "LiDiaPend.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   1320
         Picture         =   "LiDiaPend.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Limpiar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5400
         Picture         =   "LiDiaPend.frx":0E3C
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "LiDiaPend.frx":0F3E
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3615
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "terminal"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12065
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cambios
    ' 13/4 le agregue a la consulta que las ventas teléfonicas tengan envreclamocobro.
    ' 14/1/05   Agregue lista de días
    ' 20/01/05 Agregue fecha prometida.
Option Explicit

Private strInsertados As String
Private strEncabezado As String, strFormato As String
Private aTexto As String
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

Private Sub bPreview_Click()
    AccionImprimir True
End Sub


Private Sub cboCamion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco bConsultar
End Sub

Private Sub cMoneda_GotFocus()
    With cMoneda: .SelStart = 0: .SelLength = Len(.Text): End With
    Ayuda "Seleccione la moneda con que se facturó."
End Sub
Private Sub cMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco cboCamion
End Sub
Private Sub cMoneda_LostFocus()
    With cMoneda: .SelStart = 0: End With
    Ayuda ""
End Sub

Private Sub cSucursal_GotFocus()
    With cSucursal
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda " Seleccione una sucursal."
End Sub
Private Sub cSucursal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tHasta  'Foco tDesde
End Sub
Private Sub cSucursal_LostFocus()
    cSucursal.SelStart = 0: Ayuda ""
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
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

Private Sub Form_Load()
On Error GoTo ErrLoad


    ObtengoSeteoForm Me, 1000, 500, 3840, 4230
    picBotones.BorderStyle = vbBSNone
        
    LimpioGrilla
    
    'Cargo Sucursales.-------------------------------
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal " _
        & " Order By SucAbreviacion "
    CargoCombo Cons, cSucursal
    cSucursal.AddItem "Todos"
    cSucursal.ItemData(cSucursal.NewIndex) = 0
    'Por defecto pongo todos.
    For I = 0 To cSucursal.ListCount - 1
        If cSucursal.ItemData(I) = 0 Then cSucursal.ListIndex = I: Exit For
    Next
    Cons = "Select MonCodigo, MonSigno From Moneda Order by MonSigno"
    CargoCombo Cons, cMoneda
    BuscoCodigoEnCombo cMoneda, CLng(paMonedaPesos)
    CargoCombo "SELECT CamCodigo, CamNombre FROM Camion WHERE CamHabilitado = 1 ORDER BY CamNombre", cboCamion
    '--------------------------------------------------------------
    tDesde.Text = Format(Date, FormatoFP)
    tHasta.Text = Format(Date, FormatoFP)
    
    
    'Inicializo control de impresión.
    Dim fHeader As New StdFont
    With fHeader
        .Bold = True
        .Name = "Arial"
        .Size = 11
    End With
    With gpPrint
        .Caption = "Diario de pendientes"
        .FileName = "DiarioPendiente"
        .Font = Font
        Set .HeaderFont = fHeader
        .Orientation = opPortrait
        .PaperSize = 1
        .PageBorder = opTopBottom
        .MarginTop = 1000
        .MarginLeft = 500
        .MarginRight = 400
    End With
    '-----------------------------------------------
    
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error inesperado al cargar el formulario.", Err.Description
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    'vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    
    
    fFiltros.Width = Me.ScaleWidth - 240
    
    
    vsGrilla.Move 120, vsGrilla.Top, fFiltros.Width - vsDetalle.Width, Me.ScaleHeight - (vsGrilla.Top + Status.Height + picBotones.Height + 70)
    vsDetalle.Move vsGrilla.Width + 120, vsGrilla.Top, vsDetalle.Width, vsGrilla.Height
    
    picBotones.Top = vsGrilla.Height + vsGrilla.Top + 70
    
    
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
    Foco cSucursal
End Sub

Private Sub AccionImprimir(Optional bPreview As Boolean = False)
    On Error GoTo ErrImprimir
    
    vsGrilla.ExtendLastCol = False
    With gpPrint
        '.Header = "Diario de Pendiente Contados desde " & tDesde.Text & " hasta " & tHasta.Text
        .Header = "Diario de Pendiente Contados al " & tHasta.Text
        .AddGrid vsGrilla.hwnd
        .AddGrid vsDetalle.hwnd
        If bPreview Then
            .ShowPreview
        Else
            .GoPrint
        End If
    End With
    vsGrilla.ExtendLastCol = True
    Exit Sub
    
ErrImprimir:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub

Private Sub Label2_Click()
    Foco tDesde
End Sub

Private Sub Label3_Click()
    Foco tHasta
End Sub

Private Sub tDesde_GotFocus()
    With tDesde
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese desde que fecha desea consultar."
End Sub
Private Sub tDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(tDesde.Text) Then
            Foco tHasta
        Else
            MsgBox "La fecha desde no es correcta.", vbExclamation, "ATENCIÓN"
        End If
    End If
End Sub
Private Sub tDesde_LostFocus()
    tDesde.SelStart = 0: Ayuda ""
    If IsDate(tDesde.Text) Then tDesde.Text = Format(tDesde.Text, FormatoFP)
End Sub
Private Sub tHasta_GotFocus()
    With tHasta
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Ayuda "Ingrese hasta que fecha desea consultar."
End Sub
Private Sub tHasta_KeyPress(KeyAscii As Integer)
    If vbKeyReturn = KeyAscii Then Foco cMoneda
End Sub
Private Sub tHasta_LostFocus()
    Ayuda ""
    If IsDate(tHasta.Text) Then tHasta.Text = Format(tHasta.Text, FormatoFP)
End Sub

Private Sub AccionConsultar2()

    On Error GoTo ErrCDML
    
    If Not ValidoDatos Then Exit Sub
        
    Screen.MousePointer = 11
    vsGrilla.ZOrder 0
    vsDetalle.ZOrder 0
    LimpioGrilla
    
    vsGrilla.Redraw = True
    vsDetalle.Redraw = True
    
    
    Cons = "SELECT Documento.*, Renglon.*, Articulo.*, CodTipo = EnvCodigo, DpeTipo Tipo, SucAbreviacion, IsNull(CamNombre, '') as CamNombre, CASE WHEN DPeTipo = 1 THEN EnvFechaEntregado ELSE VisFModificacion END as FEnt , IsNull(EnvCamion, 0) as CamCod, CASE WHEN DPeTipo = 1 THEN EnvFechaPrometida ELSE VisFecha END as FProm " _
        & "FROM DocumentoPendiente " _
        & "INNER JOIN Documento ON DPeDocumento = DocCodigo And DocAnulado = 0 INNER JOIN Renglon ON DocCodigo = RenDocumento " _
        & "INNER JOIN Articulo ON RenArticulo = ArtID INNER JOIN Sucursal ON SucCodigo = DocSucursal " _
        & "LEFT OUTER JOIN Envio ON EnvCodigo = DPeIDTipo " _
        & "LEFT OUTER JOIN ServicioVisita ON VisServicio = DPeIDTipo AND VisDocumento = DPeDocumento " _
        & "LEFT OUTER JOIN Camion ON CASE WHEN DPeTipo = 1 THEN EnvCamion ELSE VisCamion END = CamCodigo " _
        & "WHERE (DPeIDLiquidacion Is Null OR DPeFLiquidacion > '" & Format(tHasta.Text, "yyyymmdd 23:59:59") & "')" _
        & "AND DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) & " AND DPeDisponibilidad IS NOT Null " _
        & "AND DocFecha <= '" & Format(tHasta.Text, "yyyymmdd 23:59:59") & "'"
        
        '& "AND DocFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59:59") & "'"

    If cboCamion.ListIndex > -1 Then Cons = Cons & " AND CASE WHEN DPeTipo = 1 THEN EnvCamion ELSE VisCamion END = " & cboCamion.ItemData(cboCamion.ListIndex)

    Cons = Cons & " Order by DocSucursal, DocCodigo, CamNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    CargoDatos True
    RsAux.Close
    vsDetalle.Redraw = True
    vsGrilla.Redraw = True
    
    Screen.MousePointer = 0
    Exit Sub
ErrCDML:
    vsGrilla.Redraw = True
    vsDetalle.Redraw = True
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos.", Err.Description
End Sub


Private Sub AccionConsultar()

    On Error GoTo ErrCDML
    AccionConsultar2
    Exit Sub
    
    If Not ValidoDatos Then Exit Sub
        
    Screen.MousePointer = 11
    vsGrilla.ZOrder 0
    vsDetalle.ZOrder 0
    LimpioGrilla
        
    'Saco ventas telefonicas que esten a confirmar y no se le haya hecho nota o anulado el documento.
    Cons = "Select Documento.*, Renglon.*, Articulo.*, CodTipo = EnvCodigo, Tipo = 1, SucAbreviacion, '' as CamNombre, EnvFechaEntregado as FEnt, EnvCamion as CamCod, EnvFechaPrometida as FProm " _
        & " From VentaTelefonica, Documento, Renglon, Articulo, Sucursal, Envio " _
        & " Where DocTipo = " & TipoDocumento.Contado _
        & " And DocFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59:59") & "'" _
        & " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
        & " And DocCodigo Not IN (Select NotFactura From Nota) " _
        & " AND DocCodigo Not IN (SELECT DPeDocumento FROM DocumentoPendiente WHERE DPeTipo = 1)"
    
    If cSucursal.ItemData(cSucursal.ListIndex) > 0 Then Cons = Cons & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
        
        
    '& " And EnvEstado IN ( " & EstadoEnvio.AConfirmar & ", " & EstadoEnvio.AImprimir & ") And EnvReclamoCobro <> 0"
    Cons = Cons _
        & " And DocAnulado = 0" _
        & " And EnvReclamoCobro <> 0" _
        & " And DocSucursal = SucCodigo And EnvLiquidacion Is Null And DocCodigo = VTeDocumento And DocCodigo = EnvDocumento " _
        & " And DocCodigo = RenDocumento And RenArticulo = ArtId"
    
    'Saco las ventas telefónicas.
'    Cons = Cons & " UNION ALL Select Documento.*, Renglon.*, Articulo.*, CodTipo = EnvCodigo, Tipo = 1, SucAbreviacion, CamNombre, EnvFechaEntregado as FEnt, EnvCamion as CamCod, EnvFechaPrometida as FProm " _
'        & " From VentaTelefonica, Documento, Renglon, Articulo, Sucursal, Envio, Camion " _
'        & " Where DocTipo = " & TipoDocumento.Contado _
'        & " And DocFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59:59") & "'" _
'        & " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
'
'    If cSucursal.ItemData(cSucursal.ListIndex) > 0 Then Cons = Cons & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
'
'    '& " And EnvEstado IN ( " & EstadoEnvio.Impreso & ", " & EstadoEnvio.Entregado & ") And EnvReclamoCobro <> 0"
'    Cons = Cons _
'        & " And DocAnulado = 0" _
'        & " And EnvReclamoCobro <> 0" _
'        & " And DocSucursal = SucCodigo And EnvLiquidacion Is Null And DocCodigo = VTeDocumento And DocCodigo = EnvDocumento " _
'        & " And DocCodigo = RenDocumento And RenArticulo = ArtId AND EnvCamion = CamCodigo"

    'Le uno todas las facturas que se hicieron para cobrar el Flete.
    Cons = Cons & " UNION ALL " _
        & "Select Documento.*, Renglon.*, Articulo.*, CodTipo = EnvCodigo, Tipo = 1, SucAbreviacion, IsNull(CamNombre, '') as CamNombre, EnvFechaEntregado as FEnt, EnvCamion as CamCod, EnvFechaPrometida as FProm " _
            & " From Documento, Renglon, Envio" _
            & " Left Outer Join Camion On EnvCamion = CamCodigo" _
            & ", Articulo, Sucursal" _
        & " Where DocTipo = " & TipoDocumento.Contado _
        & " And DocFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59:59") & "'" _
        & " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
      '  & " AND DocCodigo Not IN (SELECT DPeDocumento FROM DocumentoPendiente WHERE DPeTipo = 1)"
    
    If cSucursal.ItemData(cSucursal.ListIndex) > 0 Then Cons = Cons & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
    
    Cons = Cons _
        & " And DocAnulado = 0" _
        & " And EnvEstado Not In ( " & EstadoEnvio.Anulado & ", " & EstadoEnvio.Rebotado & ")" _
        & " And EnvLiquidacion Is Null  And EnvFormaPago = " & TipoPagoEnvio.PagaDomicilio _
        & " And DocSucursal = SucCodigo And DocCodigo = EnvDocumentoFactura And EnvDocumento <> EnvDocumentoFactura " _
        & " And DocCodigo = RenDocumento And RenArticulo = ArtId"
        
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    CargoDatos False
    RsAux.Close
    
    
    'Le uno las diferencias de Envios.
    'Cons = Cons & " UNION ALL "
    Cons = " Select Documento.*, Renglon.*, Articulo.*, CodTipo = EnvCodigo, Tipo = 1, SucAbreviacion, IsNull(CamNombre, '') as CamNombre, EnvFechaEntregado as FEnt , EnvCamion as CamCod, EnvFechaPrometida as FProm " _
            & " From Documento, Renglon, Articulo, DiferenciaEnvio, Sucursal, Envio " _
            & " Left Outer Join Camion On EnvCamion = CamCodigo" _
        & " Where DocTipo = " & TipoDocumento.Contado _
        & " And DocFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59:59") & "'" _
        & " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
        & " AND DocCodigo Not IN (SELECT DPeDocumento FROM DocumentoPendiente WHERE DPeTipo = 1)"
    
    If cSucursal.ItemData(cSucursal.ListIndex) > 0 Then Cons = Cons & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
    
    Cons = Cons _
        & " And DocAnulado = 0" _
        & " And EnvEstado Not IN ( " & EstadoEnvio.Anulado & ", " & EstadoEnvio.Rebotado & ")" _
        & " And EnvLiquidacion Is Null " _
        & " And DEvFormaPago = " & TipoPagoEnvio.PagaDomicilio _
        & " And DocSucursal = SucCodigo And DocCodigo = DEvDocumento And EnvCodigo = DEvEnvio " _
        & " And DocCodigo = RenDocumento And RenArticulo = ArtId" _
    & " UNION ALL " _
        & "Select Documento.*, Renglon.*, Articulo.*, CodTipo = SerCodigo, Tipo = 2,  SucAbreviacion, CamNombre, Null as FEnt, CamCodigo as CamCod, VisFecha as FProm " _
        & " From DocumentoPendiente, Documento, Renglon, Articulo, Sucursal, Servicio, ServicioVisita, Camion " _
        & " Where DPeTipo = " & DocPendiente.Servicio _
        & " And DPeDisponibilidad = 1  AND DPeIDLiquidacion IS NULL And DocFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59:59") & "'" _
        & " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) & " And DocAnulado = 0"
    
    If cSucursal.ItemData(cSucursal.ListIndex) > 0 Then Cons = Cons & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
    
    Cons = Cons & " And DPeDocumento = DocCodigo And DocCodigo = RenDocumento And RenArticulo = ArtID " _
        & " And DocSucursal = SucCodigo And DPeIDTipo = SerCodigo And SerCodigo = VisServicio And VisCamion = CamCodigo And VisLiquidada IS Null"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    CargoDatos False
    RsAux.Close
        
        
    'Cons = Cons & " UNION ALL "
     Cons = "SELECT Documento.*, Renglon.*, Articulo.*, CodTipo = EnvCodigo, Tipo = 1, SucAbreviacion, IsNull(CamNombre, '') as CamNombre, EnvFechaEntregado as FEnt , IsNull(EnvCamion, 0) as CamCod, EnvFechaPrometida as FProm " _
        & "FROM DocumentoPendiente " _
        & "INNER JOIN Documento ON DPeDocumento = DocCodigo INNER JOIN Renglon ON DocCodigo = RenDocumento " _
        & "INNER JOIN Articulo ON RenArticulo = ArtID INNER JOIN Sucursal ON SucCodigo = DocSucursal INNER JOIN Envio ON EnvCodigo = DPeIDTipo AND EnvLiquidacion IS NULL " _
        & "LEFT OUTER JOIN Camion ON EnvCamion = CamCodigo " _
        & "WHERE DPeTipo = 1 And DPeDisponibilidad = 1 " _
        & "AND DocFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59:59") & "'" _
        & "AND DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) & " And DocAnulado = 0"
        
    Cons = Cons & " Order by DocSucursal, DocCodigo, CamNombre"
    
    cBase.QueryTimeout = 120
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    
    vsDetalle.Redraw = False
    vsGrilla.Redraw = False
        
    
    If Not RsAux.EOF Then
'        vsDetalle.Redraw = False
 '       vsGrilla.Redraw = False
        CargoDatos True
        
  '      vsGrilla.Redraw = True
   '     vsDetalle.Redraw = True
    '    RsAux.Close
'    Else
     '   RsAux.Close
      '  MsgBox "No hay datos a desplegar.", vbExclamation, "ATENCIÓN"
    End If
    RsAux.Close
    
    vsGrilla.Redraw = True
    vsDetalle.Redraw = True
    
    Screen.MousePointer = 0
    Exit Sub
    
    Cons = "SELECT Documento.*, Renglon.*, Articulo.*, CodTipo = EnvCodigo, Tipo = 1, SucAbreviacion, IsNull(CamNombre, '') as CamNombre, EnvFechaEntregado as FEnt , IsNull(EnvCamion, 0) as CamCod, EnvFechaPrometida as FProm " _
        & "FROM DocumentoPendiente " _
        & "INNER JOIN Documento ON DPeDocumento = DocCodigo INNER JOIN Renglon ON DocCodigo = RenDocumento " _
        & "INNER JOIN Articulo ON RenArticulo = ArtID INNER JOIN Sucursal ON SucCodigo = DocSucursal INNER JOIN Envio ON EnvCodigo = DPeIDTipo " _
        & "LEFT OUTER JOIN Camion ON EnvCamion = CamCodigo " _
        & "WHERE DPeTipo = 1 And DPeIDLiquidacion Is Null " _
        & "AND DocFecha Between '" & Format(tDesde.Text, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Text, "mm/dd/yyyy 23:59:59") & "'" _
        & "AND DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) & " And DocAnulado = 0"
        
    Cons = Cons & " Order by DocSucursal, DocCodigo, CamNombre"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurReadOnly)
    CargoDatos True
    RsAux.Close
    vsDetalle.Redraw = True
    vsGrilla.Redraw = True
    
    Screen.MousePointer = 0
    Exit Sub
ErrCDML:
    vsGrilla.Redraw = True
    vsDetalle.Redraw = True
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al cargar los datos.", Err.Description
End Sub
Private Function ExisteEnGrilla(CodDoc As Long, IDArt As Long)
    ExisteEnGrilla = False
    If InStr(strInsertados & ",", "," & CodDoc & ":" & IDArt & ",") = 0 Then
        strInsertados = strInsertados & "," & CodDoc & ":" & IDArt
    Else
        ExisteEnGrilla = True
    End If
    Exit Function
    With vsGrilla
        For I = 0 To .Rows - 1
            If .Cell(flexcpData, I, 0) = CodDoc And .Cell(flexcpData, I, 1) = IDArt Then
                ExisteEnGrilla = True: Exit Function
            End If
        Next I
    End With
End Function
Private Sub CargoDatos(ByVal sumodatos As Boolean)
Dim aDoc As Long, aArt As Long
    
    Do While Not RsAux.EOF
        aDoc = RsAux!DocCodigo
        aArt = RsAux!ArtID
        'Esto es una chinura era la forma más rápida para no cambiar las consultas.
        If Not ExisteEnGrilla(aDoc, aArt) Then
            With vsGrilla
                .AddItem ""
                'Esto es una chinura era la forma más rápida para no cambiar las consultas.
                .Cell(flexcpData, .Rows - 1, 0) = aDoc
                .Cell(flexcpData, .Rows - 1, 1) = aArt
                .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!SucAbreviacion)
                .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!DocFecha, "dd/mm/yy hh:mm")
                .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!DocSerie) & " " & Trim(RsAux!DocNumero)
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!ArtCodigo, "#,000,000")
                .Cell(flexcpText, .Rows - 1, 4) = Trim(RsAux!ArtNombre)
                .Cell(flexcpText, .Rows - 1, 5) = RsAux!RenCantidad
                .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!RenCantidad * (RsAux!RenPrecio), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 7) = Trim(RsAux!CamNombre)
                If RsAux!Tipo = 1 Then
                    .Cell(flexcpText, .Rows - 1, 8) = "E. " & Trim(RsAux!CodTipo)
                Else
                    .Cell(flexcpText, .Rows - 1, 8) = "S. " & Trim(RsAux!CodTipo)
                End If
                If Not IsNull(RsAux!FEnt) Then .Cell(flexcpText, .Rows - 1, 9) = Format(RsAux!FEnt, "dd/mm/yy")
                If Not IsNull(RsAux!FProm) Then .Cell(flexcpText, .Rows - 1, 10) = Format(RsAux!FProm, "dd/mm/yy")
                If Not IsNull(RsAux!CamCod) Then loc_InsertIntoArrayDetalle .Cell(flexcpText, .Rows - 1, 7), .Cell(flexcpValue, .Rows - 1, 6), .Cell(flexcpText, .Rows - 1, 9)
                
            End With
        End If
        RsAux.MoveNext
    Loop
    
    If Not sumodatos Then Exit Sub
    
    With vsGrilla
        .Subtotal flexSTClear
        .Subtotal flexSTSum, 0, 6, , Obligatorio, , True
        .Subtotal flexSTSum, -1, 6, , Obligatorio, Rojo, True, "Total"
    End With
    
    If vsDetalle.Rows > vsDetalle.FixedRows Then
'        Cons = "SELECT CamNombre, LiqID, LiqTotal-SUM(LCoCobrado) Importe " & _
'            "FROM Liquidacion LEFT OUTER JOIN LiquidacionCobro ON LiqID = LCoLiquidacion INNER JOIN Camion ON LiqEnte = CamCodigo " & _
'            "GROUP BY liqid, CamNombre, LiqTotal HAVING LiqTotal <> Sum(LCoCobrado)"

        '20120428 si el total a reclamar va a dif. de cambio entonces la suma me retornaba null y no presentaba dicho importe.
        Cons = "SELECT CamNombre, LiqID, LiqTotal-IsNull(SUM(LCoCobrado), 0) Importe " & _
            "FROM Liquidacion LEFT OUTER JOIN LiquidacionCobro ON LiqID = LCoLiquidacion INNER JOIN Camion ON LiqEnte = CamCodigo "
            
        If cboCamion.ListIndex > -1 Then Cons = Cons & " AND CamCodigo = " & cboCamion.ItemData(cboCamion.ListIndex)
        
        Cons = Cons & "GROUP BY liqid, CamNombre, LiqTotal HAVING LiqTotal <> IsNull(Sum(LCoCobrado), 0) Order By CamNombre"


        Dim rsDif As rdoResultset
        Set rsDif = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not rsDif.EOF
            
            With vsDetalle
                .AddItem Trim(rsDif("CamNombre"))
                .Cell(flexcpText, .Rows - 1, 1) = "Dif:" & rsDif("LiqID")
                .Cell(flexcpText, .Rows - 1, 2) = Format(rsDif("Importe"), FormatoMonedaP)
            End With
            
            rsDif.MoveNext
        Loop
        rsDif.Close
        
        With vsDetalle
'            If .Rows > .FixedRows Then
                .Select .FixedRows, 0, .Rows - 1, 1
                .Sort = flexSortGenericDescending
                .MergeCells = flexMergeRestrictColumns And flexMergeSpill
                .MergeCol(0) = True
                .Subtotal flexSTClear
                .Subtotal flexSTSum, 0, 2, , &HE0FFFF, , True, "Total"
                .Select 1, 0
'            End If
        End With
    End If
    If vsGrilla.Rows > 1 Then vsGrilla.SetFocus
End Sub

Private Sub loc_InsertIntoArrayDetalle(ByVal sCamion As String, ByVal Importe As Currency, ByVal sFecha As String)
Dim iQ As Integer
    
    'Busco si ya inserte el camión.
    If sCamion = "" Then sCamion = "... S/C ..."
    If IsDate(sFecha) Then sFecha = Format(CDate(sFecha), "dd/mm/yy")
    
    With vsDetalle
    
        For iQ = .FixedRows To .Rows - 1
            If Trim(.Cell(flexcpText, iQ, 0)) = Trim(sCamion) And sFecha = .Cell(flexcpText, iQ, 1) Then
                'Incremento el importe
                .Cell(flexcpText, iQ, 2) = Format(CCur(.Cell(flexcpValue, iQ, 2)) + Importe, FormatoMonedaP)
                Exit Sub
            End If
        Next iQ
        
        .AddItem Trim(sCamion)
        .Cell(flexcpText, .Rows - 1, 1) = sFecha
        .Cell(flexcpText, .Rows - 1, 2) = Format(Importe, FormatoMonedaP)
        
    End With
    
End Sub

Private Sub AccionLimpiar()
On Error Resume Next
    LimpioGrilla
End Sub

Private Sub LimpioGrilla()
    
    vsDetalle.Rows = 1
    strInsertados = ""
    With vsGrilla
        .Redraw = False
        .ExtendLastCol = True
        .Clear
        .Rows = 1
        .Cols = 1
        .FormatString = "Sucursal|<Fecha|<Factura|>Código|<Nombre|>  Q|>Contado|<Camión|<Tipo|Entregado|Prometida|"
        .ColWidth(0) = 110: .ColWidth(1) = 1250: .ColWidth(2) = 800: .ColWidth(3) = 820: .ColWidth(4) = 2200: .ColWidth(6) = 1050: .ColWidth(7) = 900: .ColWidth(8) = 900: .ColWidth(9) = 1000: .ColWidth(10) = 1000
        .ColWidth(5) = 350
        .MergeCells = flexMergeSpill
        '.MergeCol(0) = True
        .OutlineBar = flexOutlineBarSimple
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        .AllowUserResizing = flexResizeColumns
        .Redraw = True
    End With
End Sub
Private Sub Ayuda(strTexto As String)
    Status.Panels(3).Text = strTexto
End Sub
Private Function ValidoDatos() As Boolean
    ValidoDatos = False
    If cSucursal.ListIndex = -1 Then
        MsgBox "No selecciono una sucursal válida.", vbExclamation, "ATENCIÓN"
        cSucursal.SetFocus: Exit Function
    End If
'    If Not IsDate(tDesde.Text) Then
'        MsgBox "No ingreso una fecha válida, verifique.", vbExclamation, "ATENCIÓN"
'        tDesde.SetFocus: Exit Function
'    End If
    If Not IsDate(tHasta.Text) Then
        MsgBox "No ingreso una fecha válida, verifique.", vbExclamation, "ATENCIÓN"
        tHasta.SetFocus: Exit Function
    End If
'    If CDate(tDesde.Text) > CDate(tHasta.Text) Then
'        MsgBox "No se ingreso un rango de fechas válido, verifique.", vbExclamation, "ATENCIÓN"
'        Foco tDesde: Exit Function
'    End If
    If cMoneda.ListIndex = -1 Then
        MsgBox "Seleccione una moneda.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Function
    End If
    ValidoDatos = True
End Function

