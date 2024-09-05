VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCarpetasAbiertas 
   Caption         =   "Carpetas Abiertas"
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
   Icon            =   "frmCarpetasAbiertas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin VSFlex8LCtl.VSFlexGrid vsConsulta 
      Height          =   1455
      Left            =   4320
      TabIndex        =   11
      Top             =   960
      Width           =   4095
      _cx             =   7223
      _cy             =   2566
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
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
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
      AutoSearchDelay =   2
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
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
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
            Object.Width           =   13123
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
      Height          =   660
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton butExcel 
         Height          =   310
         Left            =   6240
         Picture         =   "frmCarpetasAbiertas.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Exportar a excel"
         Top             =   240
         Width           =   310
      End
      Begin VB.CheckBox chArribadas 
         Caption         =   "&Arribadas"
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox tNumero 
         Height          =   285
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.PictureBox picBotones 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   5400
         ScaleHeight     =   405
         ScaleWidth      =   3015
         TabIndex        =   5
         Top             =   200
         Width           =   3015
         Begin VB.CommandButton bImprimir 
            Height          =   310
            Left            =   1200
            Picture         =   "frmCarpetasAbiertas.frx":0784
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Imprimir."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bNoFiltros 
            Height          =   310
            Left            =   1560
            Picture         =   "frmCarpetasAbiertas.frx":0886
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Quitar filtros."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bCancelar 
            Height          =   310
            Left            =   2280
            Picture         =   "frmCarpetasAbiertas.frx":0C4C
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Salir."
            Top             =   50
            Width           =   310
         End
         Begin VB.CommandButton bConsultar 
            Height          =   310
            Left            =   120
            Picture         =   "frmCarpetasAbiertas.frx":0D4E
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Ejecutar."
            Top             =   50
            Width           =   310
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "D&esde Nº de carpeta:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   255
         Width           =   1575
      End
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4455
      Left            =   240
      TabIndex        =   4
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
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCarpetasAbiertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RsAux As rdoResultset
Private aTexto As String

Private Sub AccionLimpiar()
    chArribadas.Value = vbGrayed
    tNumero.Text = ""
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

Private Sub butExcel_Click()
On Error GoTo errBE
    Dim sFile As String
    sFile = fnc_Browse(1, Replace(Me.Caption, "/", "-") & ".xls", "Exportar a excel")
    If sFile = "" Then Exit Sub
    vsConsulta.SaveGrid sFile, flexFileExcel, SaveExcelSettings.flexXLSaveFixedRows Or SaveExcelSettings.flexXLSaveRaw
errBE:
End Sub

Private Function fnc_Browse(ByVal xToFile As Byte, ByVal sFileN As String, ByVal sDialogT As String, Optional bShowSave As Boolean = True) As String
On Error GoTo errCancel
fnc_Browse = ""
 
    'Inicializo INITDIR
'    fnc_ValDirectory
            
    With cdFile
        .CancelError = True
        .DialogTitle = sDialogT
    'Var global
        '.InitDir = mExportDir
        If bShowSave Then .FileName = sFileN
        .Flags = cdlOFNOverwritePrompt Or cdlOFNCreatePrompt Or cdlOFNPathMustExist
        Select Case xToFile '1-Excel;   2-csv;  3=html
            Case 1: .Filter = "Libro de Microsoft Excel|*.xls"
            Case 2: .Filter = "Archivo de texto (csv)|*.csv"
            Case 3: .Filter = "Archivo html (*.htm)|*.htm"""
        End Select
        If bShowSave Then
            .ShowSave
        Else
            .ShowOpen
        End If
    End With
    fnc_Browse = cdFile.FileName
errCancel:
End Function



Private Sub chArribadas_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub chArribadas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    With chArribadas
    If Button = vbRightButton Then
        Select Case .Value
            Case vbGrayed:  .Value = vbChecked
            Case vbChecked:  .Value = vbUnchecked
            Case vbUnchecked:  .Value = vbGrayed
        End Select
    End If
    End With
    
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
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = "Carpeta|LC|Banco|^Apertura|>Divisa|Transporte|Embarque|^Arribó|>Q|Artículo"
            
        .WordWrap = False
        .ColWidth(0) = 800: .ColWidth(1) = 800: .ColWidth(2) = 1300: .ColWidth(3) = 900: .ColWidth(4) = 1600: .ColWidth(5) = 1400
        .ColWidth(7) = 900: .ColWidth(8) = 550 ': .ColWidth(9) = 2700
        .ColDataType(2) = flexDTCurrency
        
        .MergeCells = flexMergeRestrictAll
        .MergeCol(0) = True: .MergeCol(1) = True: .MergeCol(2) = True: .MergeCol(3) = True: .MergeCol(4) = True: .MergeCol(4) = True
        .ColAlignment(0) = flexAlignLeftTop: .ColAlignment(1) = flexAlignLeftTop: .ColAlignment(2) = flexAlignLeftTop: .ColAlignment(3) = flexAlignRightTop: .ColAlignment(4) = flexAlignLeftTop: .ColAlignment(4) = flexAlignLeftTop
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
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub AccionConsultar()
Dim rsMon As rdoResultset
Dim aImporte As Currency
Dim aMonedaID As Long, aMonedaTXT As String: aMonedaID = 0

    On Error GoTo errConsultar
    If Not ValidoFiltros Then Exit Sub
    
    Screen.MousePointer = 11
    Cons = " Select * From ArticuloFolder, " _
                                 & " Embarque left outer join Transporte on EmbTransporte = TraCodigo, " _
                                 & " Carpeta left outer join BancoLocal on CarBcoEmisor = BLoCodigo, " _
                                 & " Articulo" _
            & " Where AFoTipo = " & Folder.cFEmbarque _
            & " And AFoCodigo = EmbID" _
            & " And AFoArticulo = ArtID " _
            & " And EmbCarpeta = CarID"
        
    If Val(tNumero.Text) <> 0 Then Cons = Cons & " And CarCodigo > " & tNumero.Text
        
    Select Case chArribadas.Value
        Case vbChecked: Cons = Cons & " And EmbFArribo <> null"
        Case vbUnchecked: Cons = Cons & " And EmbFArribo = null"
    End Select
    Cons = Cons & " Order by AFoTipo, AFoCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If RsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        RsAux.Close: Screen.MousePointer = 0: Exit Sub
    End If

    With vsConsulta
        .Rows = 1
        Do While Not RsAux.EOF
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = RsAux!CarCodigo & "." & Trim(RsAux!EmbCodigo)
            .Cell(flexcpText, .Rows - 1, 8) = RsAux!AFoCantidad
            .Cell(flexcpText, .Rows - 1, 9) = Trim(RsAux!ArtNombre)
            If Not IsNull(RsAux!CarCartaCredito) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!CarCartaCredito)
            
            If Not IsNull(RsAux!BLoNombre) Then .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!BLoNombre)
            
            If Not IsNull(RsAux!EmbMoneda) Then
                If aMonedaID <> RsAux!EmbMoneda Then
                    Cons = "Select * from Moneda Where MonCodigo = " & RsAux!EmbMoneda
                    Set rsMon = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    aMonedaTXT = ""
                    aMonedaID = RsAux!EmbMoneda
                    If Not rsMon.EOF Then If Not IsNull(rsMon!MonSigno) Then aMonedaTXT = Trim(rsMon!MonSigno)
                    rsMon.Close
                End If
                .Cell(flexcpText, .Rows - 1, 4) = aMonedaTXT & " " & Format(RsAux!EmbDivisa, "##,##0.00")
            End If
            
            If Not IsNull(RsAux!EmbFEmbarque) Then
                .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!EmbFEmbarque, "dd/mm/yy")
            Else
                If Not IsNull(RsAux!EmbFEPrometido) Then
                    .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!EmbFEPrometido, "dd/mm/yy")
                    .Cell(flexcpForeColor, .Rows - 1, 6) = Colores.Rojo: .Cell(flexcpFontItalic, .Rows - 1, 6) = True
                End If
            End If
            
            
            If Not IsNull(RsAux!CarFApertura) Then .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!CarFApertura, "dd/mm/yy")
            If Not IsNull(RsAux!TraNombre) Then .Cell(flexcpText, .Rows - 1, 5) = Trim(RsAux!TraNombre)
            If Not IsNull(RsAux!EmbFArribo) Then .Cell(flexcpText, .Rows - 1, 7) = Format(RsAux!EmbFArribo, "dd/mm/yy")
            
            RsAux.MoveNext
            
        Loop
        RsAux.Close
        
        Screen.MousePointer = 0
        
    End With
    Exit Sub

errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Function ValidoFiltros() As Boolean

    ValidoFiltros = False
    If Trim(tNumero.Text) <> "" Then
        If Not IsNumeric(tNumero.Text) Then
            MsgBox "El número de carpeta ingresado no es correcto. Verifique.", vbExclamation, "ATENCIÓN"
            Foco tNumero: Exit Function
        End If
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
        .Orientation = orLandscape
        If Not .PrintDialog(pdPrinterSetup) Then Screen.MousePointer = 0: Exit Sub
        
        .Preview = True
        .StartDoc
                
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "ATENCIÓN"
            Screen.MousePointer = vbDefault: Exit Sub
        End If
    
        EncabezadoListado vsListado, "Carpetas Abiertas.", False
        .FileName = "Carpetas Abiertas."
        
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

Private Sub Label1_Click()
    Foco tNumero
End Sub

Private Sub tNumero_GotFocus()
    With tNumero: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then chArribadas.SetFocus
End Sub

Private Sub tNumero_LostFocus()
    If IsDate(tNumero.Text) Then tNumero.Text = Format(tNumero.Text, "dd/mm/yyyy")
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

