VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmListado 
   Caption         =   "Saldo Mercadería Importación A Recibir"
   ClientHeight    =   7530
   ClientLeft      =   1305
   ClientTop       =   2010
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
   ScaleHeight     =   7530
   ScaleWidth      =   11880
   Begin VB.Frame frmFiltro 
      Caption         =   "Filtros"
      ForeColor       =   &H00000080&
      Height          =   555
      Left            =   60
      TabIndex        =   19
      Top             =   60
      Width           =   7755
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "&Procesar gastos anteriores al:"
         Height          =   255
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   2235
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   4320
      TabIndex        =   14
      Top             =   1260
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
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   4455
      Left            =   60
      TabIndex        =   16
      Top             =   660
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
      PreviewMode     =   1
      Zoom            =   70
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   11655
      TabIndex        =   17
      Top             =   6720
      Width           =   11715
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmListado.frx":0EF0
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmListado.frx":0FDA
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmListado.frx":10C4
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "frmListado.frx":12FE
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Picture         =   "frmListado.frx":1400
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5400
         Picture         =   "frmListado.frx":17C6
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":18C8
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmListado.frx":1BCA
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmListado.frx":1F0C
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmListado.frx":220E
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   5940
         TabIndex        =   18
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
      TabIndex        =   15
      Top             =   7275
      Width           =   11880
      _ExtentX        =   20955
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
            Object.Width           =   12753
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

Private rsAux As rdoResultset
Dim rsFol As rdoResultset

Private aTexto As String
Dim bCargarImpresion As Boolean
Dim bHayDivisa As Boolean

Private Sub AccionLimpiar()
    vsConsulta.Rows = 1
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
    FechaDelServidor
    tFecha.Text = Format(gFechaServidor, "dd/mm/yyyy")
    
    frmListado.Status.Panels("terminal") = "Terminal: " & miConexion.NombreTerminal
    frmListado.Status.Panels("usuario") = "Usuario: " & miConexion.UsuarioLogueado(Nombre:=True)
    frmListado.Status.Panels("bd") = "BD: " & PropiedadesConnect(cBase.Connect, Database:=True) & " "
        
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    InicializoGrillas
    bCargarImpresion = True
    
    With vsListado
        .PhysicalPage = True
        .PaperSize = 1
        .Orientation = orPortrait
        .Zoom = 100
        .MarginLeft = 1000: .MarginRight = 250
        .MarginBottom = 750: .MarginTop = 750
    End With
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        '.FormatString = "<Nº de Carpeta|>Saldo en Pesos|"
        .FormatString = "<Carpeta|<Id_Gasto|<Fecha|<Documento|>Pesos (C)|>Dólares (C)|>Saldo en Pesos|"
        .ColWidth(0) = 1100: .ColWidth(1) = 900: .ColWidth(2) = 1000: .ColWidth(3) = 1000: .ColWidth(4) = 1500: .ColWidth(5) = 1300: .ColWidth(6) = 1700
        
        .WordWrap = False
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

    frmFiltro.Left = 60
    vsListado.Top = frmFiltro.Top + frmFiltro.Height + 100: vsListado.Left = frmFiltro.Left
    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
    frmFiltro.Width = Me.ScaleWidth - (vsListado.Left * 2)
    vsListado.Width = Me.ScaleWidth - (vsListado.Left * 2)
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = vsListado.Height
    vsConsulta.Left = vsListado.Left
    
     picBotones.Width = vsConsulta.Width
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 150
    
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

Dim aCantidad As Long
Dim aTexto As String
Dim aGastos As Currency

    On Error GoTo errConsultar
    If Not IsDate(tFecha.Text) Then
        MsgBox "Debe ingresar una fecha válida para realizar la consulta.", vbExclamation, " ATENCIÓN"
        Foco tFecha: Exit Sub
    End If
    
    If MsgBox("Confirma realizar la consulta al " & Format(tFecha.Text, "Long Date"), vbQuestion + vbYesNo, "Consultar") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    bCargarImpresion = True
       
    vsConsulta.Rows = 1: vsConsulta.Refresh
    aCantidad = 0
    
    cons = "Select Count(*) from Carpeta " & _
            " Where CarCosteada = 0" & _
            " And CarFAnulada Is Null" & _
            " And CarCodigo > 3772"
               
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux(0)) Then aCantidad = rsAux(0)
    rsAux.Close
    
    If aCantidad = 0 Then
        MsgBox "No hay datos a desplegar de deaudores al " & Format(gFechaServidor, "Long Date") & ".", vbInformation, "ATENCION"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    pbProgreso.Value = 0
    pbProgreso.Max = aCantidad
    vsConsulta.Redraw = False
    
    cons = "Select * from Carpeta " & _
        " Where CarCosteada = 0" & _
        " And CarFAnulada Is Null" & _
        " And CarCodigo > 3772"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        pbProgreso.Value = pbProgreso.Value + 1
        With vsConsulta
                        
            aGastos = 0: bHayDivisa = False
            
            ProcesoGastos rsAux!CarId, rsAux!CarCodigo
            
            'If Not bHayDivisa Then .Cell(flexcpBackColor, .Row, 0, , .Cols - 1) = Colores.Inactivo Else .Cell(flexcpBackColor, .Row, 0, , .Cols - 1) = Colores.Blanco
                        
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    With vsConsulta
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, 0, 6, , Colores.Inactivo, , False, "Total %s"
        .Subtotal flexSTSum, -1, 6, , Colores.Rojo, Colores.Blanco, True, "TOTAL"
        
        For I = 1 To .Rows - 2
            If .IsSubtotal(I) Then .IsCollapsed(I) = flexOutlineCollapsed
        Next
        
    End With
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
    pbProgreso.Value = 0
    
    Exit Sub
errConsultar:
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True
    Screen.MousePointer = 0
    pbProgreso.Value = 0
End Sub

Private Function ProcesoGastos(aFolder As Long, aTxtFolder As String) As Currency

Dim aRetorno As Currency
    
    On Error GoTo errProceso
    aRetorno = 0
    '.FormatString = "<Carpeta|<Id_Gasto|<Fecha|>Documento|>Pesos|>Dólares|>Saldo en Pesos|"
    
    'Saco lo que queda por costear de los gastos del Nivel Carpeta-------------------------------------------------
    cons = "Select * from GastoImportacion, Compra" & _
                " Where GImFolder = " & aFolder & _
                " And GImNivelFolder = " & Folder.cFCarpeta & _
                " And GImIDCompra = ComCodigo" & _
                " And ComFecha < '" & Format(tFecha.Text, sqlFormatoF) & " 23:59'"
    Set rsFol = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsFol.EOF
        With vsConsulta
            .AddItem Trim(aTxtFolder)
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsFol!ComCodigo, "#,###")
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsFol!ComFecha, "dd/mm/yyyy")
            If Not IsNull(rsFol!ComSerie) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(rsFol!ComSerie) & " "
            If Not IsNull(rsFol!ComNumero) Then .Cell(flexcpText, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 3) & rsFol!ComNumero
            
            If rsFol!ComMoneda <> paMonedaPesos Then
                .Cell(flexcpText, .Rows - 1, 4) = Format(rsFol!ComImporte * rsFol!ComTC, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 5) = Format(rsFol!ComImporte, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 6) = Format(rsFol!GImCostear * rsFol!ComTC, FormatoMonedaP)
                
            Else
                .Cell(flexcpText, .Rows - 1, 4) = Format(rsFol!ComImporte, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 6) = Format(rsFol!GImCostear, FormatoMonedaP)
            End If
        End With
        
        If rsFol!GImIdSubRubro = paSubrubroDivisa And Not bHayDivisa Then bHayDivisa = True
        rsFol.MoveNext
    Loop
    rsFol.Close
    '-------------------------------------------------------------------------------------------------------------------------
    
    'Saco lo que queda por costear de los gastos del Nivel Embarque-------------------------------------------------
    cons = "Select * from GastoImportacion, Compra" & _
                " Where GImFolder In (Select EmbID from Embarque Where EmbCarpeta = " & aFolder & " And EmbCosteado = 0)" & _
                " And GImNivelFolder = " & Folder.cFEmbarque & _
                " And GImIDCompra = ComCodigo" & _
                " And ComFecha < '" & Format(tFecha.Text, sqlFormatoF) & " 23:59'"
    Set rsFol = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsFol.EOF
        With vsConsulta
            .AddItem Trim(aTxtFolder)
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsFol!ComCodigo, "#,###")
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsFol!ComFecha, "dd/mm/yyyy")
            If Not IsNull(rsFol!ComSerie) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(rsFol!ComSerie) & " "
            If Not IsNull(rsFol!ComNumero) Then .Cell(flexcpText, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 3) & rsFol!ComNumero

            If rsFol!ComMoneda <> paMonedaPesos Then
                'aRetorno = aRetorno + Format((rsFol!GImCostear * rsFol!ComTC), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 4) = Format(rsFol!ComImporte * rsFol!ComTC, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 5) = Format(rsFol!ComImporte, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 6) = Format(rsFol!GImCostear * rsFol!ComTC, FormatoMonedaP)
            Else
                'aRetorno = aRetorno + Format(rsFol!GImCostear, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 4) = Format(rsFol!ComImporte, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 6) = Format(rsFol!GImCostear, FormatoMonedaP)
            End If
        End With
        
        If rsFol!GImIdSubRubro = paSubrubroDivisa And Not bHayDivisa Then bHayDivisa = True
        rsFol.MoveNext
    Loop
    rsFol.Close
    '-------------------------------------------------------------------------------------------------------------------------
    
    'Saco lo que queda por costear de los gastos del Nivel Subcarpeta-------------------------------------------------
    cons = "Select * from GastoImportacion, Compra" & _
                " Where GImFolder In (Select SubID From Embarque, SubCarpeta Where EmbCarpeta = " & aFolder & " And SubEmbarque = EmbId And SubCosteada = 0)" & _
                " And GImNivelFolder = " & Folder.cFSubCarpeta & _
                " And GImIDCompra = ComCodigo" & _
                " And ComFecha < '" & Format(tFecha.Text, sqlFormatoF) & " 23:59'"
    Set rsFol = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsFol.EOF
        With vsConsulta
            .AddItem Trim(aTxtFolder)
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsFol!ComCodigo, "#,###")
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsFol!ComFecha, "dd/mm/yyyy")
            If Not IsNull(rsFol!ComSerie) Then .Cell(flexcpText, .Rows - 1, 3) = Trim(rsFol!ComSerie) & " "
            If Not IsNull(rsFol!ComNumero) Then .Cell(flexcpText, .Rows - 1, 3) = .Cell(flexcpText, .Rows - 1, 3) & rsFol!ComNumero

            If rsFol!ComMoneda <> paMonedaPesos Then
                'aRetorno = aRetorno + Format((rsFol!GImCostear * rsFol!ComTC), FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 4) = Format(rsFol!ComImporte * rsFol!ComTC, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 5) = Format(rsFol!ComImporte, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 6) = Format(rsFol!GImCostear * rsFol!ComTC, FormatoMonedaP)
            Else
                'aRetorno = aRetorno + Format(rsFol!GImCostear, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 4) = Format(rsFol!ComImporte, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 6) = Format(rsFol!GImCostear, FormatoMonedaP)
            End If
        End With
        
        If rsFol!GImIdSubRubro = paSubrubroDivisa And Not bHayDivisa Then bHayDivisa = True
        rsFol.MoveNext
    Loop
    rsFol.Close
    '-------------------------------------------------------------------------------------------------------------------------
    
    ProcesoGastos = aRetorno
    Exit Function

errProceso:
    clsGeneral.OcurrioError "Ocurrió un error al procesar los gastos de la carpeta.", Err.Description
End Function

Private Sub Label1_Click()
    Foco tFecha
End Sub

Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0: tFecha.SelLength = Len(tFecha.Text)
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

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    If bCargarImpresion Then
        If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub
        With vsListado
            .StartDoc
            .Columns = 1
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        EncabezadoListado vsListado, "Saldo Mercadería Importación a Recibir al " & Format(tFecha.Text, "Long Date"), False
        vsListado.FileName = "Saldo Mercaderia Importacion"
            
        With vsConsulta
            .Redraw = False
            '.FontSize = 6
            'AnchoEncabezado Impresora:=True
            vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
            'AnchoEncabezado Pantalla:=True
            '.FontSize = 8
            .Redraw = True
        End With
        
        vsListado.EndDoc
        'bCargarImpresion = False
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
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la impresión", Err.Description
End Sub

Private Sub AccionConfigurar()
    
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
    
End Sub

