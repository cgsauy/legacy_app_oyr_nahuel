VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.0#0"; "AACOMBO.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmChequesLibrados 
   Caption         =   "Cheques Girados"
   ClientHeight    =   6420
   ClientLeft      =   165
   ClientTop       =   450
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
   Icon            =   "frmChequesLibrados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   2040
      TabIndex        =   6
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
      TabIndex        =   8
      Top             =   6165
      Width           =   12105
      _ExtentX        =   21352
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
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   11175
      Begin VB.TextBox tHasta 
         Height          =   285
         Left            =   6960
         TabIndex        =   5
         Text            =   "10/12/2000"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tDesde 
         Height          =   285
         Left            =   5160
         TabIndex        =   3
         Text            =   "10/12/2000"
         Top             =   240
         Width           =   975
      End
      Begin VB.PictureBox picBotones 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   8160
         ScaleHeight     =   465
         ScaleWidth      =   2175
         TabIndex        =   10
         Top             =   200
         Width           =   2175
         Begin VB.CommandButton bImprimir 
            Height          =   310
            Left            =   720
            Picture         =   "frmChequesLibrados.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Imprimir."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bNoFiltros 
            Height          =   310
            Left            =   1080
            Picture         =   "frmChequesLibrados.frx":0544
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Quitar filtros."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bCancelar 
            Height          =   310
            Left            =   1800
            Picture         =   "frmChequesLibrados.frx":090A
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Salir."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bConsultar 
            Height          =   310
            Left            =   120
            Picture         =   "frmChequesLibrados.frx":0A0C
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Ejecutar."
            Top             =   50
            Width           =   310
         End
      End
      Begin AACombo99.AACombo cDisponibilidad 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   6360
         TabIndex        =   4
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "D&esde:"
         Height          =   255
         Left            =   4560
         TabIndex        =   2
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "&Disponibilidad:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4455
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   11415
      _Version        =   196608
      _ExtentX        =   20135
      _ExtentY        =   7858
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
End
Attribute VB_Name = "frmChequesLibrados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RsAux As rdoResultset
Private aTexto As String

Private Sub AccionLimpiar()
    cDisponibilidad.Text = ""
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

Private Sub cDisponibilidad_GotFocus()
    With cDisponibilidad
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cDisponibilidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDesde
End Sub

Private Sub cDisponibilidad_LostFocus()
    cDisponibilidad.SelStart = 0
End Sub

Private Sub Label5_Click()
    Foco cDisponibilidad
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

    Cons = "Select DisID, DisNombre From Disponibilidad Where DisSucursal <> NULL Order by DisNombre"
    CargoCombo Cons, cDisponibilidad
    
    tDesde.Text = Format(PrimerDia(PrimerDia(Now) - 1), "dd/mm/yyyy")
    tHasta.Text = Format(PrimerDia(Now) - 1, "dd/mm/yyyy")
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .OutlineBar = flexOutlineBarSimple
                
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "|Nº Cheque|Librado|Vence|>Importe|ID_Gasto|Proveedor|>Importe|>Ch.- Gastos|"
            
        .WordWrap = True
        .ColWidth(0) = 165: .ColWidth(1) = 1100: .ColWidth(4) = 1200: .ColWidth(3) = 950: .ColWidth(2) = 950
        .ColWidth(6) = 2300: .ColWidth(7) = 1200: .ColWidth(8) = 1000
        .ColDataType(2) = flexDTCurrency
                
        '.MergeCol(0) = True: .MergeCol(1) = True: .MergeCol(2) = True: .MergeCol(3) = True: .MergeCol(4) = True
        .ColAlignment(0) = flexAlignLeftTop: .ColAlignment(1) = flexAlignLeftTop: .ColAlignment(4) = flexAlignRightTop: .ColAlignment(3) = flexAlignLeftTop
        .MergeCells = flexMergeSpill

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

Private Sub AccionConsultar(Optional OrdenadoPorCheque As Boolean = False)

Dim aImporte As Currency

    On Error GoTo errConsultar
    If Not ValidoFiltros Then Exit Sub
    
    Screen.MousePointer = 11
    
    Cons = "Select * from Cheque, Disponibilidad , ChequePago, Compra, ProveedorCliente " _
           & " Where CheIDDisponibilidad = DisID " _
           & " And CheID = CPaIDCheque " _
           & " And CPaIDCompra = ComCodigo And ComProveedor = PClCodigo"
     If cDisponibilidad.ListIndex <> -1 Then Cons = Cons & " And CheIDDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    
    Cons = Cons & " And CheLibrado Between '" & Format(tDesde.Text, sqlFormatoF) & "' And '" & Format(tHasta.Text, sqlFormatoF) & "'"
    
    If Not OrdenadoPorCheque Then
        Cons = Cons & " ORDER BY CheIDDisponibilidad, CheLibrado, CheID"
    Else
        Cons = Cons & " ORDER BY CheIDDisponibilidad, CheSerie, CheNumero"
    End If
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        RsAux.Close: Screen.MousePointer = 0: Exit Sub
    End If
    
    Dim aAnterior As Long: aAnterior = 0
    Dim aSuma As Currency: aSuma = 0
    aImporte = 0
    
    With vsConsulta
        .Rows = 1
        Do While Not RsAux.EOF
            
            If aAnterior <> RsAux!CPaIDCheque And aSuma <> 0 Then
                If (aImporte - aSuma) <> 0 Then .Cell(flexcpText, .Rows - 1, 8) = Format(aImporte - aSuma, FormatoMonedaP)
            End If
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Trim(RsAux!DisNombre)
            .Cell(flexcpText, .Rows - 1, 1) = " "
            If aAnterior <> RsAux!CPaIDCheque Then
                aSuma = 0
                aAnterior = RsAux!CPaIDCheque
                aImporte = RsAux!CheImporte
            
                .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!CheSerie) & " " & RsAux!CheNumero
                If Not IsNull(RsAux!CheVencimiento) Then .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!CheVencimiento, "dd/mm/yyyy")
                .Cell(flexcpText, .Rows - 1, 2) = Format(RsAux!CheLibrado, "dd/mm/yyyy")
                .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!CheImporte, FormatoMonedaP)
            End If
            
            .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!CPaIDCompra, "#,##0")
            .Cell(flexcpText, .Rows - 1, 6) = Trim(RsAux!PClNombre)
            .Cell(flexcpText, .Rows - 1, 7) = Format(RsAux!CPaImporte, FormatoMonedaP)
            
            aSuma = aSuma + RsAux!CPaImporte
            
            RsAux.MoveNext
            
        Loop
        RsAux.Close
        
        Screen.MousePointer = 0
        
        .Subtotal flexSTSum, 0, 4, , Colores.Obligatorio, &H80&, True, "%s"
        .Subtotal flexSTSum, 0, 7, , Colores.Obligatorio, &H80&, True
        
    End With
    Exit Sub

errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Function ValidoFiltros() As Boolean

    ValidoFiltros = False
    If Not IsDate(tDesde.Text) Then
        MsgBox "La fecha desde ingresada no es correcta. Verifique.", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    If Not IsDate(tHasta.Text) Then
        MsgBox "La fecha hasta ingresada no es correcta. Verifique.", vbExclamation, "ATENCIÓN"
        Foco tHasta: Exit Function
    End If
    
    If CDate(tDesde.Text) > CDate(tHasta.Text) Then
        MsgBox "El período de fechas ingresado no es correcto (desde debe ser menor al hasta).", vbExclamation, "ATENCIÓN"
        Foco tDesde: Exit Function
    End If
    
    ValidoFiltros = True
    
End Function

Private Sub AccionImprimir()
Dim J As Integer

    If vsConsulta.Rows = 1 Then
        MsgBox "No hay datos en la lista para realizar la impresión.", vbExclamation, "ATENCIÓN"
        Exit Sub
    End If
    
    On Error GoTo errPrint
    Screen.MousePointer = 11
    
    With vsListado
    
        If Not .PrintDialog(pdPrinterSetup) Then Screen.MousePointer = 11: Exit Sub
        
        .Preview = True
        .StartDoc
                
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN"
            Screen.MousePointer = vbDefault: Exit Sub
        End If
    
        EncabezadoListado vsListado, "Cheques Librados.", False
        .filename = "Cheques Librados."
        
        .FontSize = 9: .FontBold = True
        If cDisponibilidad.ListIndex <> -1 Then aTexto = cDisponibilidad.Text Else aTexto = "(Todos)"
        aTexto = "Disponibilidad: " & aTexto & "         del " & Trim(tDesde.Text) & " al " & Trim(tHasta.Text)
        .Paragraph = "": .Paragraph = aTexto: .Paragraph = ""
        .FontSize = 8: .FontBold = False
        vsConsulta.ExtendLastCol = False: .RenderControl = vsConsulta.hWnd: vsConsulta.ExtendLastCol = True
        
        .EndDoc
        .PrintDoc
        
    End With
    Screen.MousePointer = 0
    Exit Sub

errPrint:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión. ", Err.Description
End Sub

Private Sub tDesde_GotFocus()
    With tDesde: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tHasta
End Sub

Private Sub tDesde_LostFocus()
    If IsDate(tDesde.Text) Then tDesde.Text = Format(tDesde.Text, "dd/mm/yyyy")
End Sub

Private Sub tHasta_GotFocus()
    With tHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tHasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If IsDate(tHasta.Text) Then tHasta.Text = Format(tHasta.Text, "dd/mm/yyyy")
        bConsultar.SetFocus
    End If
End Sub

Private Sub tHasta_LostFocus()
    If IsDate(tHasta.Text) Then tHasta.Text = Format(tHasta.Text, "dd/mm/yyyy")
End Sub


Private Sub vsConsulta_Click()

    With vsConsulta
         If .MouseCol = 1 And .MouseRow = 0 And .Rows > 0 Then
            If MsgBox("Dese ordenar la consulta por número de cheque.", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            AccionConsultar True
         End If
    End With
    
End Sub

Private Sub vsConsulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 And vsConsulta.Rows > 1 Then PopupMenu MnuAccesos, X:=X + vsConsulta.Left, Y:=Y + vsConsulta.Top
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

