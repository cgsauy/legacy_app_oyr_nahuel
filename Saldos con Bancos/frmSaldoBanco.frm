VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Begin VB.Form frmSaldoBanco 
   Caption         =   "Saldos con Bancos"
   ClientHeight    =   6420
   ClientLeft      =   285
   ClientTop       =   1605
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSaldoBanco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   12105
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   5160
      TabIndex        =   2
      Top             =   3120
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
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   6165
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
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
            Object.Width           =   13150
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
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   10095
      Begin VB.CheckBox chVista 
         Caption         =   "A la Vista"
         Height          =   255
         Left            =   6240
         TabIndex        =   15
         Top             =   720
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chPlazo 
         Caption         =   "Plazo BL"
         Height          =   255
         Left            =   6240
         TabIndex        =   14
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chCobranza 
         Caption         =   "Cobranza"
         Height          =   255
         Left            =   4920
         TabIndex        =   13
         Top             =   720
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chAnticipado 
         Caption         =   "Anticipado"
         Height          =   255
         Left            =   4920
         TabIndex        =   12
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chPaga 
         Alignment       =   1  'Right Justify
         Caption         =   "Divisa Paga"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Value           =   2  'Grayed
         Width           =   1140
      End
      Begin VB.PictureBox picBotones 
         BorderStyle     =   0  'None
         Height          =   425
         Left            =   7800
         ScaleHeight     =   420
         ScaleWidth      =   2175
         TabIndex        =   6
         Top             =   600
         Width           =   2175
         Begin VB.CommandButton bImprimir 
            Height          =   310
            Left            =   720
            Picture         =   "frmSaldoBanco.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Imprimir."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bNoFiltros 
            Height          =   310
            Left            =   1080
            Picture         =   "frmSaldoBanco.frx":0544
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Quitar filtros."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bCancelar 
            Height          =   310
            Left            =   1800
            Picture         =   "frmSaldoBanco.frx":090A
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Salir."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bConsultar 
            Height          =   310
            Left            =   120
            Picture         =   "frmSaldoBanco.frx":0A0C
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Ejecutar."
            Top             =   50
            Width           =   310
         End
      End
      Begin AACombo99.AACombo cBanco 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
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
         BackStyle       =   0  'Transparent
         Caption         =   "Formas de Pago de Divisas"
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4815
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   11415
      _Version        =   196608
      _ExtentX        =   20135
      _ExtentY        =   8493
      _StockProps     =   229
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
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSaldoBanco.frx":0D0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSaldoBanco.frx":1028
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSaldoBanco.frx":1342
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSaldoBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsAux As rdoResultset
Private aTexto As String

Private Sub AccionLimpiar()
    cBanco.Text = ""
    chPaga.Value = vbUnchecked
    chAnticipado.Value = vbChecked
    chCobranza.Value = vbChecked
    chVista.Value = vbChecked
    chPlazo.Value = vbChecked
End Sub

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

Private Sub cBanco_GotFocus()
    With cBanco
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chPaga.SetFocus
End Sub

Private Sub cBanco_LostFocus()
    cBanco.SelStart = 0
End Sub

Private Sub chAnticipado_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chCobranza.SetFocus
End Sub

Private Sub chCobranza_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chPlazo.SetFocus
End Sub

Private Sub chPaga_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chAnticipado.SetFocus
End Sub

Private Sub chPaga_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then EstadoCheck chPaga
End Sub

Private Sub chPlazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chVista.SetFocus
End Sub

Private Sub chVista_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub Label5_Click()
    Foco cBanco
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()

    On Error GoTo ErrLoad

    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    InicializoGrillas
    AccionLimpiar

    cons = "Select BLoCodigo, BLoNombre From BancoLocal Order by BLoNombre"
    CargoCombo cons, cBanco
    
    cBanco.AddItem "(Sin Banco Emisor)"
    cBanco.ItemData(cBanco.NewIndex) = 0
    
    FechaDelServidor
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = ">Embarca|Prometido|<Forma de Pago|<Plazo|Vence|<LC|<Carpeta|Paga|Moneda|>Importe Divisa U$S|>T/C|"
            
        .WordWrap = True
        .ColWidth(0) = 1150: .ColWidth(1) = 950:
        .ColWidth(3) = 800: .ColWidth(4) = 950: .ColWidth(5) = 1000: .ColWidth(6) = 800: .ColWidth(9) = 1600: .ColWidth(10) = 950
        .ColDataType(0) = flexDTDate: .ColDataType(1) = flexDTDate: .ColDataType(4) = flexDTDate: .ColDataType(9) = flexDTCurrency
    End With
      
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

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Screen.MousePointer = 11
    picBotones.BorderStyle = vbFlat
    fFiltros.Width = Me.Width - (fFiltros.Left * 2.5)
    
    vsConsulta.Left = fFiltros.Left
    vsConsulta.Top = fFiltros.Top + fFiltros.Height + 50
    vsConsulta.Height = Me.ScaleHeight - (vsConsulta.Top + Status.Height + 90)
    vsConsulta.Width = fFiltros.Width
    
    Screen.MousePointer = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub AccionConsultar()

Dim Vence As Date
Dim aValor As Long, aIndice As Integer

    If cBanco.ListIndex = -1 Then MsgBox "Debe seleccionar el banco para realizar la consulta.", vbExclamation, "ATENCIÓN": Exit Sub
    
    If chAnticipado.Value = vbUnchecked And chCobranza.Value = vbUnchecked And chPlazo.Value = vbUnchecked And chVista.Value = vbUnchecked Then
        MsgBox "Debe seleccionar la forma de pago de la divisa para realizar la consulta.", vbExclamation, "ATENCIÓN": Exit Sub
    End If
    
    On Error GoTo errConsultar
    Screen.MousePointer = 11
    
    Dim idBanco As Long
    idBanco = cBanco.ItemData(cBanco.ListIndex)
    
    cons = "Select * from Carpeta, Moneda, Embarque"
    If idBanco > 0 Then
        cons = cons & " Where CarBCoEmisor = " & cBanco.ItemData(cBanco.ListIndex)
    Else
        cons = cons & " Where CarBCoEmisor Is Null"
    End If
    
    cons = cons & " And EmbCarpeta = CarID" _
                        & " And EmbMoneda = MonCodigo" _
                        & " And EmbDivisa > 0 " _
                        & " And CarFAnulada is Null"
            
    Select Case chPaga.Value
        Case vbChecked: cons = cons & " And EmbDivisaPaga = 1"
        Case vbUnchecked: cons = cons & " And EmbDivisaPaga = 0"
    End Select
    
    'Formas de pago
    Dim aEn As String: aEn = ""
    
    If chAnticipado.Value = vbChecked Then aEn = aEn & FormaPago.cFPAnticipado & ","
    If chCobranza.Value = vbChecked Then aEn = aEn & FormaPago.cFPCobranza & ","
    If chVista.Value = vbChecked Then aEn = aEn & FormaPago.cFPVista & ","
    If chPlazo.Value = vbChecked Then aEn = aEn & FormaPago.cFPPlazoBL & ","
    
    If Trim(aEn) <> "" Then
        aEn = Mid(aEn, 1, Len(aEn) - 1)
        cons = cons & " And CarFormaPago In ( " & aEn & ")"
    End If
            
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    With vsConsulta
    .Rows = 1
    Do While Not rsAux.EOF
        
        .AddItem ""
        If Not IsNull(rsAux!EmbFEmbarque) Then .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!EmbFEmbarque, "dd/mm/yyyy")
        If Not IsNull(rsAux!EmbFEPrometido) Then .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!EmbFEPrometido, "dd/mm/yyyy")
        
        If Not IsNull(rsAux!CarFormaPago) Then .Cell(flexcpText, .Rows - 1, 2) = RetornoFormaPago(rsAux!CarFormaPago)
        If Not IsNull(rsAux!CarPlazo) Then If rsAux!CarPlazo <> 0 Then .Cell(flexcpText, .Rows - 1, 3) = rsAux!CarPlazo
        
        If Not IsNull(rsAux!EmbFEmbarque) Then  'Trabajo con fecha de embarque real
            If Not IsNull(rsAux!CarPlazo) Then
                Vence = SumoDias(Format(rsAux!EmbFEmbarque, "dd/mm/yyyy"), CLng(rsAux!CarPlazo))
            Else
                Vence = rsAux!EmbFEmbarque
            End If
            .Cell(flexcpText, .Rows - 1, 4) = Format(Vence, "dd/mm/yyyy")
            
        Else        'Trabajo con fecha de embarque prometido
            If Not IsNull(rsAux!EmbFEPrometido) Then
                If Not IsNull(rsAux!CarPlazo) Then
                    Vence = SumoDias(Format(rsAux!EmbFEPrometido, "dd/mm/yyyy"), CLng(rsAux!CarPlazo))
                Else
                    Vence = rsAux!EmbFEPrometido
                End If
                .Cell(flexcpText, .Rows - 1, 4) = Format(Vence, "dd/mm/yyyy")
            End If
        End If
        
        aValor = rsAux!EmbID: .Cell(flexcpData, .Rows - 1, 0) = aValor
        
        'Cargo el icono asociado segun la fecha ------------------------------------------------------------
        If rsAux!EmbDivisaPaga = 0 Then
            If Vence < gFechaServidor Then
                aIndice = 1
            Else
                If Weekday(gFechaServidor) + paVencimientoLC > 6 Then      'Hay un S y D el el medio
                    If Vence <= gFechaServidor + paVencimientoLC + ((Weekday(gFechaServidor) + paVencimientoLC) / 7) * 2 Then
                        aIndice = 2
                    Else
                        aIndice = 3
                    End If
                Else
                    If Vence <= gFechaServidor + paVencimientoLC Then aIndice = 2 Else aIndice = 3
                End If
            End If
            .Cell(flexcpPicture, .Rows - 1, 0) = ImageList1.ListImages(aIndice).ExtractIcon
        End If
        '-----------------------------------------------------------------------------------------------------------
        
        If Not IsNull(rsAux!CarCartaCredito) Then .Cell(flexcpText, .Rows - 1, 5) = Trim(rsAux!CarCartaCredito)
        .Cell(flexcpText, .Rows - 1, 6) = Trim(rsAux!CarCodigo) & "." & Trim(rsAux!EmbCodigo)
                
        If rsAux!EmbDivisaPaga = 0 Then .Cell(flexcpText, .Rows - 1, 7) = "No" Else .Cell(flexcpText, .Rows - 1, 7) = "Si"
        
        .Cell(flexcpText, .Rows - 1, 8) = Trim(rsAux!MonSigno)
        If Not IsNull(rsAux!EmbDivisa) Then
            If Not IsNull(rsAux!EmbArbitraje) Then
                .Cell(flexcpText, .Rows - 1, 9) = Format(rsAux!EmbDivisa / rsAux!EmbArbitraje, "##,##0.00")
            Else
                .Cell(flexcpText, .Rows - 1, 9) = Format(rsAux!EmbDivisa, "##,##0.00")
            End If
        Else
            .Cell(flexcpText, .Rows - 1, 9) = "0.00"
        End If
        
        If Not IsNull(rsAux!EmbArbitraje) Then .Cell(flexcpText, .Rows - 1, 10) = Format(rsAux!EmbArbitraje, "#,##0.000") Else .Cell(flexcpText, .Rows - 1, 10) = "0.000"

        rsAux.MoveNext
    Loop
    rsAux.Close
    
    .Subtotal flexSTSum, -1, 9, , &H80&, vbWhite, True, "Total General"
    Screen.MousePointer = 0
    End With
    
    Exit Sub
errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub


Private Sub AccionImprimir()
Dim J As Integer

    If vsConsulta.Rows = 1 Then
        MsgBox "No hay datos en la lista para realizar la impresión.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    On Error GoTo errPrint
    Screen.MousePointer = 11
    
    With vsListado
    
        If Not .PrintDialog(pdPrinterSetup) Then Screen.MousePointer = 0: Exit Sub
        
        .Preview = True
        .StartDoc
                
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN"
            Screen.MousePointer = vbDefault: Exit Sub
        End If
    
        EncabezadoListado vsListado, "Importaciones - Saldos con Bancos.", False
        
        .FileName = "Saldos con Bancos"
        
        .FontSize = 9: .FontBold = True
        If cBanco.ListIndex <> -1 Then aTexto = cBanco.Text Else aTexto = "(Todos)"
        .Paragraph = "": .Paragraph = "Banco: " & aTexto: .Paragraph = ""
        .FontSize = 8: .FontBold = False
        vsConsulta.ExtendLastCol = False: .RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        
        .EndDoc
        .PrintDoc
        
    End With
    
    Screen.MousePointer = 0
    Exit Sub

errPrint:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión. ", Err.Description
End Sub

Private Sub vsConsulta_Click()
    
    On Error Resume Next
    With vsConsulta
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

Private Function EstadoCheck(chBox As CheckBox)

    Select Case chBox.Value
        Case vbGrayed:  chBox.Value = vbChecked
        Case vbChecked:  chBox.Value = vbUnchecked
        Case vbUnchecked:  chBox.Value = vbGrayed
    End Select
    
End Function

