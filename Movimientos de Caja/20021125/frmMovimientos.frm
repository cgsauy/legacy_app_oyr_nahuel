VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Begin VB.Form frmMovimientos 
   Caption         =   "Movimientos de Caja"
   ClientHeight    =   6420
   ClientLeft      =   540
   ClientTop       =   1860
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
   Icon            =   "frmMovimientos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   12105
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   4320
      TabIndex        =   6
      Top             =   1560
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
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   "sucursal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   "terminal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   "bd"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10663
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
      Height          =   1020
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   9735
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
      End
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         MaxLength       =   12
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.PictureBox picBotones 
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   7440
         ScaleHeight     =   405
         ScaleWidth      =   2175
         TabIndex        =   10
         Top             =   200
         Width           =   2175
         Begin VB.CommandButton bImprimir 
            Height          =   310
            Left            =   720
            Picture         =   "frmMovimientos.frx":0442
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
            Picture         =   "frmMovimientos.frx":0544
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
            Picture         =   "frmMovimientos.frx":090A
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
            Picture         =   "frmMovimientos.frx":0A0C
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   "Ejecutar."
            Top             =   50
            Width           =   310
         End
      End
      Begin AACombo99.AACombo cMovimiento 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
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
      Begin VB.Label Label3 
         Caption         =   "&Fecha:"
         Height          =   255
         Left            =   5280
         TabIndex        =   4
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "&Disponibilidad:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   260
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Movimiento:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   615
         Width           =   1095
      End
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4455
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   9615
      _Version        =   196608
      _ExtentX        =   16960
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
Attribute VB_Name = "frmMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsAux As rdoResultset
Private aTexto As String

Private Sub AccionLimpiar()
    cDisponibilidad.Text = ""
    cMovimiento.Text = ""
    tFecha.Text = ""
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

Private Sub cDisponibilidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco cMovimiento
End Sub

Private Sub cMovimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco tFecha
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
    
    cons = "Select DisID, DisNombre from Disponibilidad Order by DisNombre"
    CargoCombo cons, cDisponibilidad
    
    
    cons = "Select TMDCodigo, TMDNombre from TipoMovDisponibilidad Order by TMDNombre"
    CargoCombo cons, cMovimiento
    
    If Trim(Command()) <> "" Then       'Viene Disponibilidad:Movimiento:dd/mm/yyyy
        Dim aStr As String
        aStr = Trim(Command())
        BuscoCodigoEnCombo cDisponibilidad, Val(Mid(aStr, 1, InStr(aStr, ":") - 1))
        aStr = Mid(aStr, InStr(aStr, ":") + 1, Len(aStr))
        BuscoCodigoEnCombo cMovimiento, Val(Mid(aStr, 1, InStr(aStr, ":") - 1))
        aStr = Mid(aStr, InStr(aStr, ":") + 1, Len(aStr))
        tFecha.Text = Format(aStr, "dd/mm/yyyy")
        AccionConsultar
    Else
        BuscoCodigoEnCombo cDisponibilidad, paDisponibilidad
        tFecha.Text = Format(Now, "dd/mm/yyyy")
    End If
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = "ID|Tipo de Movimiento|Fecha|Hora|Comentario|>Debe|>Haber|"
            
        .WordWrap = True
        .ColWidth(0) = 700: .ColWidth(1) = 2400: .ColWidth(2) = 800: .ColWidth(3) = 600: .ColWidth(4) = 3500: .ColWidth(5) = 1200: .ColWidth(6) = 1200
        .ColDataType(2) = flexDTCurrency
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
        
    cons = "Select * from MovimientoDisponibilidad, MovimientoDisponibilidadRenglon, TipoMovDisponibilidad" _
           & " Where MDiID = MDRIdMovimiento" _
           & " And MDiTipo = TMDCodigo" _
           & " And MDRIdDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) _
           & " And MDiFecha = '" & Format(tFecha.Text, "mm/dd/yyyy") & "'"
           
    If cMovimiento.ListIndex <> -1 Then cons = cons & " And MDiTipo = " & cMovimiento.ItemData(cMovimiento.ListIndex)
    
    cons = cons & " Order by MDiFecha, MDiHora"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        vsConsulta.Rows = 1
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        rsAux.Close: Screen.MousePointer = 0: Exit Sub
    End If

    With vsConsulta
        .Rows = 1
        Do While Not rsAux.EOF
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!MDiID, "#,##0")
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!TMDNombre)
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!MDiFecha, "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!MDiHora, "hh:mm")
            If Not IsNull(rsAux!MDiComentario) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(rsAux!MDiComentario)
            
            If Not IsNull(rsAux!MDRDebe) Then .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!MDRDebe, FormatoMonedaP)
            If Not IsNull(rsAux!MDRHaber) Then .Cell(flexcpText, .Rows - 1, 6) = Format(rsAux!MDRHaber, FormatoMonedaP)

                        
            rsAux.MoveNext
            
        Loop
        rsAux.Close
        
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 5, , , , , "Subtotal"
        .Subtotal flexSTSum, -1, 6, , , , , "Subtotal"
        
        Dim aTotal As Currency
        aTotal = .Cell(flexcpValue, .Rows - 1, 5) - .Cell(flexcpValue, .Rows - 1, 6)
        
        If aTotal <> 0 Then
            .AddItem "Total"
            If aTotal < 0 Then .Cell(flexcpText, .Rows - 1, 6) = Format(aTotal, FormatoMonedaP) Else .Cell(flexcpText, .Rows - 1, 5) = Format(aTotal, FormatoMonedaP)
        End If
        
        .Cell(flexcpBackColor, 1, 5, .Rows - 1, 7) = Colores.Obligatorio
        .Cell(flexcpForeColor, 1, 6, .Rows - 1) = Colores.Rojo
        .Cell(flexcpBackColor, .Rows - 2, 0, .Rows - 1, 7) = Colores.Obligatorio
            
        Screen.MousePointer = 0
        
    End With
    Exit Sub

errConsultar:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Function ValidoFiltros() As Boolean

    ValidoFiltros = False
    If cDisponibilidad.ListIndex = -1 Then
        MsgBox "Debe seleccionar la disponibilidad para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco cDisponibilidad: Exit Function
    End If
    
    If Not IsDate(tFecha.Text) Then
        MsgBox "Debe ingresar una fecha para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Function
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
    
        EncabezadoListado vsListado, "Movimientos de Caja.", False
        .FileName = "Movimientos de Caja."
        
        .Paragraph = "Disponibilidad: " & Trim(cDisponibilidad.Text) & Chr(vbKeyTab) & "Fecha: " & tFecha.Text
        
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
    Foco tFecha
End Sub

Private Sub tFecha_GotFocus()
    With tFecha: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

