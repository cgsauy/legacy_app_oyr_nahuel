VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmListado 
   Caption         =   "Asientos de Gastos a Crédito"
   ClientHeight    =   7590
   ClientLeft      =   1815
   ClientTop       =   2055
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
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   4680
      TabIndex        =   20
      Top             =   960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7646
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
      HighLight       =   1
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
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4455
      Left            =   120
      TabIndex        =   23
      Top             =   720
      Width           =   11415
      _Version        =   196608
      _ExtentX        =   20135
      _ExtentY        =   7858
      _StockProps     =   229
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
      Zoom            =   100
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   9555
      TabIndex        =   24
      Top             =   6720
      Width           =   9615
      Begin VB.CommandButton bExportar 
         Height          =   310
         Left            =   5280
         Picture         =   "frmListado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Exportar"
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":074C
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":084E
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0D80
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
         Picture         =   "frmListado.frx":11FA
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
         Picture         =   "frmListado.frx":12E4
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
         Picture         =   "frmListado.frx":13CE
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
         Picture         =   "frmListado.frx":1608
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
         Picture         =   "frmListado.frx":170A
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5640
         Picture         =   "frmListado.frx":1AD0
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":1BD2
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
         Picture         =   "frmListado.frx":1ED4
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
         Picture         =   "frmListado.frx":2216
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
         Picture         =   "frmListado.frx":2518
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6000
         TabIndex        =   25
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
      TabIndex        =   22
      Top             =   7335
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
            Key             =   ""
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
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton bExpandir 
         Caption         =   "Expandir"
         Height          =   315
         Left            =   8400
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chComentarios 
         Caption         =   "&Ver Comentarios"
         Height          =   195
         Left            =   6720
         TabIndex        =   6
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox tFHasta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Text            =   "28/12/2000"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Text            =   "28/12/2000"
         Top             =   240
         Width           =   1095
      End
      Begin AACombo99.AACombo cTipoListado 
         Height          =   315
         Left            =   4740
         TabIndex        =   5
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
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
      Begin VB.Label Label1 
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   4260
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   255
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog ctlDlg 
      Left            =   -60
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************* CAMBIOS ********************************************
' (1) 12/12/00  - Proc: CargoContraQueVan
'      Se agegó el parametro paSubrubroDifCostoImp por las importaciones que ya fueron costeadas y luego pagas (se lo habíamos quitado ??)
                                

Option Explicit

Private rsAux As rdoResultset, rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean
Dim txtAcreedoresVarios As String

Private Sub AccionLimpiar()
    tFecha.Text = "": tFHasta.Text = ""
    vsConsulta.Rows = 1
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub bExpandir_Click()
With vsConsulta
        For I = .FixedRows To .Rows - 1
            If .IsSubtotal(I) Then
                If Val(bExpandir.Tag) = 0 Then
                     .IsCollapsed(I) = flexOutlineExpanded
                Else
                     If .RowOutlineLevel(I) > 0 Then .IsCollapsed(I) = flexOutlineCollapsed
                End If
            End If
        Next
    End With
    bExpandir.Tag = IIf(Val(bExpandir.Tag) = 0, 1, 0)
End Sub

Private Sub bExportar_Click()
On Error GoTo errCancel
    
    With ctlDlg
        .CancelError = True
        .FileName = Me.Caption 'AsientosVarios
        .Filter = "Libro de Microsoft Exel|*.xls|" & _
                     "Texto (delimitado por tabulaciones)|*.txt|" & "Texto (delimitado por comas)|*.txt"
        
        .ShowSave
        'Confirma exportar el contenido de la lista al archivo:
        If MsgBox("Confirma exportar el contenido de la lista al archivo: " & .FileName, vbQuestion + vbYesNo) = vbYes Then
        
            With vsConsulta
                Dim iR As Integer, iX As Integer
                iX = .FixedRows
                
                For iR = .FixedRows To .Rows - 1
                    If Not (.IsSubtotal(iX) And (.RowOutlineLevel(iX) = 0)) Then
                        If Trim(.Cell(flexcpText, iX, 0)) <> "" Then 'And .Cell(flexcpForeColor, iX, 0) <> vbRed Then
                            .RemoveItem iX
                            iX = iX - 1
                        End If
                    End If
                    iX = iX + 1
                Next
                
            End With
            
            On Error GoTo errSaving
            Screen.MousePointer = 11
            Me.Refresh
            DoEvents
            
            Dim mSSetting As SaveLoadSettings
            
            Select Case .FilterIndex
                Case 1, 2: mSSetting = flexFileTabText
                Case 3: mSSetting = flexFileCommaText
            End Select
            
            vsConsulta.SaveGrid .FileName, mSSetting, True
                
            Screen.MousePointer = 0
        End If
        
    End With

errCancel:
    Screen.MousePointer = 0: Exit Sub
errSaving:
     clsGeneral.OcurrioError "Error al exportar el contenido de la lista.", Err.Description
    Screen.MousePointer = 0
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

Private Sub bConfigurar_Click()
    AccionConfigurar
End Sub

Private Sub bAnterior_Click()
    IrAPagina vsListado, vsListado.PreviewPage - 1
End Sub

Private Sub chComentarios_Click()
    
    'If chComentarios.Value = vbChecked Then
    '    vsConsulta.ColWidth(7) = 1600
    '    vsConsulta.ColWidth(9) = 1400
    '    vsConsulta.Cell(flexcpText, 0, 9) = "Comentarios"
    'Else
    '    vsConsulta.ColWidth(7) = 0
    '    vsConsulta.Cell(flexcpText, 0, 8) = ""
    '    vsConsulta.Cell(flexcpText, 0, 9) = ""
    'End If
    
    If chComentarios.Value = vbChecked Then
        vsConsulta.ColWidth(7) = 600: vsConsulta.Cell(flexcpText, 0, 7) = "T/C"
        vsConsulta.ColWidth(8) = 1100: vsConsulta.Cell(flexcpText, 0, 8) = "Proveedor"
        vsConsulta.ColWidth(9) = 1400: vsConsulta.Cell(flexcpText, 0, 9) = "Comentarios"
    Else
        vsConsulta.ColWidth(7) = 0: vsConsulta.Cell(flexcpText, 0, 7) = ""
        vsConsulta.ColWidth(8) = 0: vsConsulta.Cell(flexcpText, 0, 8) = ""
        vsConsulta.ColWidth(9) = 0: vsConsulta.Cell(flexcpText, 0, 9) = ""
    End If
    
End Sub

Private Sub cTipoListado_GotFocus()
    With cTipoListado: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub cTipoListado_KeyPress(KeyAscii As Integer)
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

Private Sub Label2_Click()
    Foco tFHasta
End Sub

Private Sub tFecha_GotFocus()
    With tFecha: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tFHasta
End Sub
Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then
        tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
        If Not IsDate(tFHasta.Text) Then tFHasta.Text = tFecha.Text
    End If
End Sub

Private Sub Label5_Click()
    Foco tFecha
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrLoad

    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    InicializoGrillas
    AccionLimpiar
    bCargarImpresion = True
    
    CargoConstantesSubrubros
    cons = "Select * from Subrubro Where SRuID = " & paSubrubroAcreedoresVarios
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then txtAcreedoresVarios = Format(rsAux!SRuCodigo, "000000000") & " " & Trim(rsAux!SRuNombre)
    rsAux.Close
    
    cTipoListado.AddItem "Ingresos": cTipoListado.ItemData(cTipoListado.NewIndex) = 0
    cTipoListado.AddItem "Egresos": cTipoListado.ItemData(cTipoListado.NewIndex) = 1
    
    'vsListado.Orientation = orPortrait
    'vsListado.MarginBottom = 750: vsListado.MarginTop = 750
    
    With vsListado
        .PhysicalPage = True
        .PaperSize = vbPRPSLetter
        .Orientation = orPortrait
        .PreviewMode = pmScreen
        .PreviewPage = 1
        .Zoom = 100
        .MarginLeft = 450: .MarginRight = 350
        .MarginBottom = 750: .MarginTop = 750
    End With
    
    pbProgreso.Value = 0
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    vsListado.MarginRight = 300
    With vsConsulta
        .OutlineBar = flexOutlineBarSimple
        '.OutlineBar = flexOutlineBarComplete 'flexOutlineBarNone
        .OutlineCol = 0
        .MultiTotals = True
        .ColSel = 0
        .ColSort(0) = flexSortStringAscending
        .Sort = flexSortUseColSort
        
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<Rubro|<SubRubro|>Importe $|>Cofis $|>I.V.A. $|>Total $|>Total M/E|<|<|<|"
            
        .WordWrap = False
        .ColWidth(0) = 2300: .ColWidth(1) = 2150: .ColWidth(2) = 1550
        .ColWidth(3) = 900: .ColWidth(4) = 1400: .ColWidth(5) = 1550: .ColWidth(6) = 1400
        .ColWidth(7) = 0: .ColWidth(8) = 0: .ColWidth(9) = 0
        .MergeCells = flexMergeSpill
        .MergeCol(0) = True
    End With
      
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
    
    picBotones.Width = vsListado.Width
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 150
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = vsListado.Height
    vsConsulta.Left = vsListado.Left
    
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    End
    
End Sub

Private Sub AccionConsultar()
Dim IdCompra As Long
Dim IvaCompra As Currency, IvaGastos As Currency
Dim CofisCompra As Currency, CofisGastos As Currency
Dim RsRub As rdoResultset

    If Not ValidoCampos Then Exit Sub
    
    On Error GoTo errConsultar
    Screen.MousePointer = 11
    pbProgreso.Value = 0
    bCargarImpresion = True
    
    vsConsulta.Rows = 1: vsConsulta.Refresh
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    cons = " Select Count(*) from Compra, GastoSubRubro " _
            & " Where ComFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And ComCodigo = GSrIDCompra"
    
    Select Case cTipoListado.ItemData(cTipoListado.ListIndex)
        Case 0: cons = cons & " And ComTipoDocumento = " & TipoDocumento.CompraNotaCredito        '0- Ingresos
        Case 1: cons = cons & " And ComTipoDocumento = " & TipoDocumento.CompraCredito              '1- Egresos
    End Select
    
    cons = cons & " And ComDCDe is null"        'Cambio 13/11/00        Irma
    
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

    cons = "Select * From Compra, GastoSubRubro, SubRubro, Rubro, ProveedorCliente" _
            & " Where ComFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And ComCodigo = GSrIDCompra" _
            & " And GSrIDSubRubro = SRuID " _
            & " And SRuRubro = RubID" _
            & " And ComProveedor = PClCodigo"
    
    Select Case cTipoListado.ItemData(cTipoListado.ListIndex)
        Case 0: cons = cons & " And ComTipoDocumento = " & TipoDocumento.CompraNotaCredito        '0- Ingresos
        Case 1: cons = cons & " And ComTipoDocumento = " & TipoDocumento.CompraCredito              '1- Egresos
    End Select
    
    cons = cons & " And ComDCDe is null"        'Cambio 13/11/00        Irma
    
    cons = cons & " Order by ComCodigo"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        rsAux.Close: Screen.MousePointer = 0: InicializoGrillas: Exit Sub
    End If
    
    IdCompra = 0
    IvaCompra = 0: IvaGastos = 0: CofisCompra = 0: CofisGastos = 0
    
    vsConsulta.Rows = 1: vsConsulta.Redraw = False
    Do While Not rsAux.EOF
        pbProgreso.Value = pbProgreso.Value + 1
        With vsConsulta
            .AddItem ""
            'En el data del item 0 pongo si el rubro expande o no
            If rsAux!RubExpandir Then .Cell(flexcpData, .Rows - 1, 0) = 1 Else .Cell(flexcpData, .Rows - 1, 0) = 0
            
            .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!RubCodigo, "000000000") & " " & Trim(rsAux!RubNombre)
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!SRuCodigo, "000000000") & " " & Trim(rsAux!SRuNombre)
            
            .Cell(flexcpText, .Rows - 1, 8) = " "
            .Cell(flexcpText, .Rows - 1, 9) = " "
            .Cell(flexcpText, .Rows - 1, 10) = " "
            If Not IsNull(rsAux!PClNombre) Then .Cell(flexcpText, .Rows - 1, 8) = Trim(rsAux!PClNombre)
            If Not IsNull(rsAux!ComComentario) Then .Cell(flexcpText, .Rows - 1, 9) = Trim(rsAux!ComComentario)
            
            
            If IdCompra <> rsAux!ComCOdigo Then
                If rsAux!ComMoneda = paMonedaPesos Then
                    .Cell(flexcpText, .Rows - 1, 2) = Format(Abs(rsAux!ComImporte), FormatoMonedaP)
                Else
                    .Cell(flexcpText, .Rows - 1, 2) = Format(Abs(rsAux!ComImporte) * rsAux!ComTC, FormatoMonedaP)
                End If
                
                If Not IsNull(rsAux!ComIVA) Then
                    If rsAux!ComMoneda = paMonedaPesos Then
                        .Cell(flexcpText, .Rows - 1, 4) = Format(Abs(rsAux!ComIVA), FormatoMonedaP)
                    Else
                        .Cell(flexcpText, .Rows - 1, 4) = Format(Abs(rsAux!ComIVA) * rsAux!ComTC, FormatoMonedaP)
                        .Cell(flexcpText, .Rows - 1, 6) = Format(Abs(rsAux!ComIVA), FormatoMonedaP)
                    End If
                    If paSubrubroCompraMercaderia = rsAux!SRuID Then
                        IvaCompra = IvaCompra + .Cell(flexcpValue, .Rows - 1, 4)
                    Else
                        IvaGastos = IvaGastos + .Cell(flexcpValue, .Rows - 1, 4)
                    End If
                End If
                
                If Not IsNull(rsAux!ComCofis) Then
                    If rsAux!ComMoneda = paMonedaPesos Then
                        .Cell(flexcpText, .Rows - 1, 3) = Format(Abs(rsAux!ComCofis), FormatoMonedaP)
                    Else
                        .Cell(flexcpText, .Rows - 1, 3) = Format(Abs(rsAux!ComCofis) * rsAux!ComTC, FormatoMonedaP)
                        .Cell(flexcpText, .Rows - 1, 6) = Format(Abs(rsAux!ComCofis) + .Cell(flexcpValue, .Rows - 1, 6), FormatoMonedaP)
                    End If
                    If paSubrubroCompraMercaderia = rsAux!SRuID Then
                        CofisCompra = CofisCompra + .Cell(flexcpValue, .Rows - 1, 3)
                    Else
                        CofisGastos = CofisGastos + .Cell(flexcpValue, .Rows - 1, 3)
                    End If
                End If
                
            End If
            If Not IsNull(rsAux!GSrIDSubrubro) Then
                If rsAux!ComMoneda = paMonedaPesos Then
                    .Cell(flexcpText, .Rows - 1, 2) = Format(Abs(rsAux!GSrImporte), FormatoMonedaP)
                Else
                    .Cell(flexcpText, .Rows - 1, 2) = Format(Abs(rsAux!GSrImporte) * rsAux!ComTC, FormatoMonedaP)
                    .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpValue, .Rows - 1, 6) + Abs(rsAux!GSrImporte), FormatoMonedaP)
                End If
            End If
            .Cell(flexcpText, .Rows - 1, 5) = Format(.Cell(flexcpValue, .Rows - 1, 2) + .Cell(flexcpValue, .Rows - 1, 3) + .Cell(flexcpValue, .Rows - 1, 4), FormatoMonedaP)
            
            If Not IsNull(rsAux!ComTC) Then .Cell(flexcpText, .Rows - 1, 7) = Format(rsAux!ComTC, "#.000")
                'If Trim(.Cell(flexcpText, .Rows - 1, 7)) = "" Then .Cell(flexcpText, .Rows - 1, 7) = Format(rsAux!ComTC, "#.000") Else .Cell(flexcpText, .Rows - 1, 7) = Format(rsAux!ComTC, "#.000") & " / " & .Cell(flexcpText, .Rows - 1, 7)
            'End If
            
            IdCompra = rsAux!ComCOdigo
            
            rsAux.MoveNext
        End With
    Loop
    rsAux.Close
    
    Dim aTotal As Currency
    With vsConsulta
        .Select 1, 0, 1, 1
        .Sort = flexSortGenericAscending
        
        'Totales para la Columna 1
        .Subtotal flexSTSum, 1, 2, , , , False, "%s"
        .Subtotal flexSTSum, 1, 3: .Subtotal flexSTSum, 1, 4: .Subtotal flexSTSum, 1, 5: .Subtotal flexSTSum, 1, 6
        .Cell(flexcpForeColor, 1, 0, .Rows - 1, .Cols - 1) = Colores.Azul
        
        'Totales para la Columna 0
        .Subtotal flexSTSum, 0, 2, , , , , "%s"
        .Subtotal flexSTSum, 0, 3: .Subtotal flexSTSum, 0, 4: .Subtotal flexSTSum, 0, 5: .Subtotal flexSTSum, 0, 6

        'Total de todos los Renglones
        .Subtotal flexSTSum, -1, 2, , Colores.Gris, , True, "Total"
        .Subtotal flexSTSum, -1, 3: .Subtotal flexSTSum, -1, 4: .Subtotal flexSTSum, -1, 5: .Subtotal flexSTSum, -1, 6
        aTotal = .Cell(flexcpValue, .Rows - 1, 5)
        
        If IvaCompra <> 0 Then
            .AddItem "I.V.A. Compras  21401"
            .Cell(flexcpText, .Rows - 1, 4) = Format(IvaCompra, FormatoMonedaP)
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
        End If
        If IvaGastos <> 0 Then
            .AddItem "I.V.A. Gastos     21403"
            .Cell(flexcpText, .Rows - 1, 4) = Format(IvaGastos, FormatoMonedaP)
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
        End If
        If CofisCompra <> 0 Then
            .AddItem "Cofis Compras    21411"
            .Cell(flexcpText, .Rows - 1, 3) = Format(CofisCompra, FormatoMonedaP)
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
        End If
        If CofisGastos <> 0 Then
            .AddItem "Cofis Gastos       21413"
            .Cell(flexcpText, .Rows - 1, 3) = Format(CofisGastos, FormatoMonedaP)
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
        End If
    End With
    
    AgrupoCamposEnGrilla
        
    'Cargo contra que concepto Van---------------------------------------------------------------------------
    CargoContraQueVan aTotal
    
    vsConsulta.Redraw = True
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True: pbProgreso.Value = 0: Screen.MousePointer = 0
End Sub

Private Sub CargoContraQueVan(aImporteAcreedores As Currency)
On Error GoTo errCQV
Dim rsPro As rdoResultset

    'Cargo contra que concepto Van---------------------------------------------------------------------------
    With vsConsulta
        .AddItem "": .AddItem ""
        Select Case cTipoListado.ItemData(cTipoListado.ListIndex)
            Case 0: .Cell(flexcpText, .Rows - 1, 1) = "Concepto al DEBE"        '0- Ingresos
            Case 1: .Cell(flexcpText, .Rows - 1, 1) = "Concepto al HABER"      '1- Egresos
        End Select
        .Cell(flexcpBackColor, .Rows - 1, 1, , .Cols - 1) = RGB(0, 0, 1) 'Colores.Inactivo
        .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco: .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
    End With
    '------------------------------------------------------------------------------------------------------------------------------------------------------

    'Cargo los pagos de divisas para poner los rubros de los bancos
    cons = "Select * From Compra, GastoSubRubro " & _
            " Where ComFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" & _
            " And ComCodigo = GSrIDCompra"
            
    Select Case cTipoListado.ItemData(cTipoListado.ListIndex)
        Case 0: cons = cons & " And ComTipoDocumento = " & TipoDocumento.CompraNotaCredito        '0- Ingresos
        Case 1: cons = cons & " And ComTipoDocumento = " & TipoDocumento.CompraCredito              '1- Egresos
    End Select
    
    'ref. (1)
    cons = cons & " And GSrIDSubrubro In (" & paSubrubroDivisa & ", " & paSubrubroDifCambioG & ", " & paSubrubroDifCambio & ", " & paSubrubroDifCostoImp & ")"
    'Cons = Cons & " And GSrIDSubrubro In (" & paSubrubroDivisa & ", " & paSubrubroDifCambioG & ", " & paSubrubroDifCambio & ")"
    
    cons = cons & " And ComDCDe is null"        'Cambio 13/11/00        Irma
    
    cons = cons & " Order by ComProveedor"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Dim aIdEmpresa As Long, aSubTotal As Currency, aSubTotalP As Currency, aSubTotalD As Currency, txtSubRubro As String
    aIdEmpresa = 0
    Dim mIdSR As Long
    
    If Not rsAux.EOF Then
        Do While Not rsAux.EOF
            If aIdEmpresa <> rsAux!ComProveedor Then
                
                If aIdEmpresa <> 0 And aSubTotal <> 0 Then
                    With vsConsulta
                        .AddItem ""
                        .Cell(flexcpText, .Rows - 1, 1) = txtSubRubro
                        .Cell(flexcpText, .Rows - 1, 5) = Format(aSubTotal, FormatoMonedaP)
                        If aSubTotalD > 0 Then
                            .Cell(flexcpText, .Rows - 1, 6) = Format(aSubTotalD, FormatoMonedaP)
                            .Cell(flexcpText, .Rows - 1, 7) = Format(aSubTotalP / aSubTotalD, "#.000") & " (" & Format(aSubTotalP, FormatoMonedaP) & " / " & Format(aSubTotalD, FormatoMonedaP) & ")"
                        End If
                        aImporteAcreedores = aImporteAcreedores - aSubTotal
                    
                        .Cell(flexcpData, .Rows - 1, 1) = CStr(aIdEmpresa) & "|"
                    End With
                End If
                
                aIdEmpresa = rsAux!ComProveedor
                
                'Saco datos del Proveedor del Gasto
                mIdSR = 0
                cons = " Select * from EmpresaDato Left Outer Join SubRubro On EDaSRubroContable = SRuID" & _
                            " Where EDaCodigo = " & aIdEmpresa & _
                            " And EDaTipoEmpresa = " & TipoEmpresa.Cliente
                Set rsPro = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                If Not rsPro.EOF Then
                    If Not IsNull(rsPro!SRuNombre) Then
                        txtSubRubro = Format(rsPro!SRuCodigo, "000000000") & " " & Trim(rsPro!SRuNombre)
                    Else
                        txtSubRubro = "Sin Datos"
                        
                        'Saco nombre de Empresa ----------------------------------------------------------------------------------------
                        Dim rsDE As rdoResultset
                        cons = "Select * from CEmpresa Where CEmCliente = " & rsPro!EDaCodigo
                        Set rsDE = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                        If Not rsDE.EOF Then txtSubRubro = txtSubRubro & " (" & Trim(rsDE!CEmNombre) & ")"
                        rsDE.Close
                        '----------------------------------------------------------------------------------------------------------------------
                    End If
                End If
                rsPro.Close
                               
                                
                aSubTotal = 0: aSubTotalD = 0: aSubTotalP = 0
            End If
            
            If Not (txtSubRubro Like "Sin Datos*") Or (txtSubRubro Like "Sin Datos*" And rsAux!GSrIDSubrubro <> paSubrubroDifCostoImp) Then
                If rsAux!ComMoneda = paMonedaPesos Then
                    aSubTotal = aSubTotal + Format(Abs(rsAux!GSrImporte), FormatoMonedaP)
                Else
                    aSubTotal = aSubTotal + Format(Abs(rsAux!GSrImporte) * rsAux!ComTC, FormatoMonedaP)
                    aSubTotalD = aSubTotalD + Abs(rsAux!GSrImporte)
                    aSubTotalP = aSubTotalP + Format(Abs(rsAux!GSrImporte) * rsAux!ComTC, FormatoMonedaP)
                End If
            End If
            
            rsAux.MoveNext
        Loop
        
        If aSubTotal <> 0 Then          'Add Ultima Empresa !!! -------------------------------------------------------------------------------
            With vsConsulta
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 1) = txtSubRubro
                .Cell(flexcpText, .Rows - 1, 5) = Format(aSubTotal, FormatoMonedaP)
                If aSubTotalD > 0 Then
                    .Cell(flexcpText, .Rows - 1, 6) = Format(aSubTotalD, FormatoMonedaP)
                    .Cell(flexcpText, .Rows - 1, 7) = Format(aSubTotalP / aSubTotalD, "#.000") & " (" & Format(aSubTotalP, FormatoMonedaP) & " / " & Format(aSubTotalD, FormatoMonedaP) & ")"
                End If
                aImporteAcreedores = aImporteAcreedores - aSubTotal
                
                .Cell(flexcpData, .Rows - 1, 1) = CStr(aIdEmpresa) & "|"
            End With
        End If
        
    End If
    rsAux.Close
    '---------------------------------------------------------------------------------------------------------------
    
    'Cargo contra que concepto Van---------------------------------------------------------------------------
    
    With vsConsulta
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = txtAcreedoresVarios
        .Cell(flexcpData, .Rows - 1, 1) = "0|0"
        .Cell(flexcpText, .Rows - 1, 5) = Format(aImporteAcreedores, FormatoMonedaP)
    End With
    
    '------------------------------------------------------------------------------------------------------------------------------------------------------
    Exit Sub
    
errCQV:
    clsGeneral.OcurrioError "Error al cargar Contra Que Van.", Err.Number & "-" & Err.Description
End Sub

Private Sub tFHasta_GotFocus()
    With tFHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tFHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cTipoListado
End Sub

Private Sub tFHasta_LostFocus()
    If IsDate(tFHasta.Text) Then tFHasta.Text = Format(tFHasta.Text, "dd/mm/yyyy")
End Sub

Private Sub vsConsulta_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)

    vsConsulta.Row = vsConsulta.MouseRow
    
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    If bCargarImpresion Then
        If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub
        
        vsListado.PaperSize = 1
        If chComentarios.Value = vbUnchecked Then
            vsListado.Orientation = orPortrait
        Else
            vsListado.Orientation = orLandscape
            vsConsulta.ColWidth(7) = 600
            vsConsulta.ColWidth(8) = 1600
            vsConsulta.ColWidth(9) = 2300
        End If
        
        With vsListado
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        EncabezadoListado vsListado, "Asientos de Gastos a Crédito (" & Trim(cTipoListado.Text) & ") - Del " & Trim(tFecha.Text) & " al " & Trim(tFHasta.Text), False
        vsListado.FileName = "Listado Asientos de Gastos a Credito"
            
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        
        vsListado.EndDoc
        'bCargarImpresion = False
        If chComentarios.Value = vbChecked Then vsConsulta.ColWidth(7) = 1400
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
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionConfigurar()
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
End Sub

Private Sub AgrupoCamposEnGrilla()

On Error GoTo ErrACEG
Dim aData As Boolean

    With vsConsulta
        For I = 1 To .Rows - 1
            If .IsSubtotal(I) Then
                If aData Then   'Hay que expandir la rama de los subrubros
                    Select Case .RowOutlineLevel(I)
                        Case 0: .IsCollapsed(I) = flexOutlineExpanded
                        Case 1: .IsCollapsed(I) = flexOutlineCollapsed
                    End Select
                Else
                    If .RowOutlineLevel(I) <> -1 Then .IsCollapsed(I) = flexOutlineCollapsed
                End If
            Else
                If .Cell(flexcpData, I, 0) = 1 Then aData = True Else aData = False
            End If
            
        Next I
    End With
    
    Exit Sub
ErrACEG:
End Sub

Private Sub vsConsulta_Collapsed()
    
    If vsConsulta.RowOutlineLevel(vsConsulta.Row) <> 1 Then Exit Sub
    
    With vsConsulta
        If .IsCollapsed(.Row) Then
            .Cell(flexcpForeColor, .Row, 0, , .Cols - 1) = Colores.Azul
        Else
            .Cell(flexcpForeColor, .Row, 0, , .Cols - 1) = vbBlack
        End If
    End With
    
End Sub

Private Sub vsConsulta_DblClick()
    
    If vsConsulta.Rows = vsConsulta.FixedRows Then Exit Sub
    If vsConsulta.Rows = -1 Then Exit Sub
    
    Dim txtDATA As String
    txtDATA = vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)
    
    'txtDATA = Proveedor|SubRubro
    
    txtDATA = vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)
    If InStr(txtDATA, "|") = 0 Then Exit Sub
    Dim arrVAL() As String
    
    Dim mFrm As New frmDetalleD
    With mFrm
        arrVAL = Split(txtDATA, "|")
        .prmIDAcreedor = arrVAL(0)
        .prmAcreedorTXT = vsConsulta.Cell(flexcpText, vsConsulta.Row, 1)
        .prmFecha1 = tFecha.Text
        .prmFecha2 = tFHasta.Text
        
        Select Case cTipoListado.ItemData(cTipoListado.ListIndex)
            Case 0: .prmIN_TDocs = TipoDocumento.CompraNotaCredito        '0- Ingresos
            Case 1: .prmIN_TDocs = TipoDocumento.CompraCredito             '1- Egresos
        End Select
        
        .Show , Me
    End With
    
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Function ValidoCampos() As Boolean

    ValidoCampos = False
    
    If Not IsDate(tFecha.Text) And IsDate(tFHasta.Text) Then
        MsgBox "Ingrese la fecha desde.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
    End If
    If IsDate(tFecha.Text) And Not IsDate(tFHasta.Text) Then
        If Trim(tFHasta.Text) = "" Then
            tFHasta.Text = tFecha.Text
        Else
            MsgBox "La fecha hasta no es correcta.", vbExclamation, "ATENCIÓN"
            Foco tFHasta: Exit Function
        End If
    End If
    If IsDate(tFecha.Text) And IsDate(tFHasta.Text) Then
        If CDate(tFecha.Text) > CDate(tFHasta.Text) Then
            MsgBox "Los rangos de fecha no son correctos.", vbExclamation, "ATENCIÓN"
            Foco tFecha: Exit Function
        End If
    End If
        
    If cTipoListado.ListIndex = -1 Then
        MsgBox "Debe seleccioar el tipo de movimientos a listr (Ingresos/Egresos).", vbExclamation, "ATENCIÓN"
        Foco cTipoListado: Exit Function
    End If
    ValidoCampos = True
    
End Function
