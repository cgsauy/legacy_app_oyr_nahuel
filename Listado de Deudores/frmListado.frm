VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmListado 
   Caption         =   "Listado de Deudores"
   ClientHeight    =   7530
   ClientLeft      =   1695
   ClientTop       =   3285
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
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   4320
      TabIndex        =   12
      Top             =   720
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
      TabIndex        =   14
      Top             =   60
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
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   11655
      TabIndex        =   15
      Top             =   6720
      Width           =   11715
      Begin VB.CommandButton bCBase 
         Caption         =   "CB"
         Height          =   310
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   8
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   7
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   0
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6180
         TabIndex        =   16
         Top             =   120
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   476
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin ComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
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

Enum TipoCV
    Compra = 1
    Comercio = 2
    Importacion = 3
End Enum

Private rsAux As rdoResultset, rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean

Dim aTotalGeneral As Currency

Private Sub AccionLimpiar()
    vsConsulta.Rows = 1
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bCBase_Click()


Dim newB As String
    On Error GoTo errCh
    If MsgBox("Ud. sabe lo que está haciendo !!?", vbQuestion + vbYesNo + vbDefaultButton2, "Realmente desea cambiar la base") = vbNo Then Exit Sub
    
    newB = InputBox("Ingrese el texto del login para la nueva conexión" & vbCrLf & _
                "Id de aplicación en archivo de conexiones.", "Cambio de Base de Datos")
    
    If Trim(newB) = "" Then Exit Sub
    If MsgBox("Está seguro de cambiar la base de datos al login " & newB, vbQuestion + vbYesNo + vbDefaultButton2, "Cambiar Base") = vbNo Then Exit Sub
    
    newB = miConexion.TextoConexion(newB)
    If Trim(newB) = "" Then Exit Sub
    
    Screen.MousePointer = 11
    On Error Resume Next
    cBase.Close
    On Error GoTo errCh
    Set cBase = Nothing
    Set cBase = eBase.OpenConnection("", rdDriverNoPrompt, , newB)
    
    Screen.MousePointer = 0
    
    MsgBox "Ahora está trabajanbo en la nueva base de datos.", vbExclamation, "Base Cambiada OK"
    Exit Sub
    
errCh:
    clsGeneral.OcurrioError "Error de Conexión", Err.Description
    Screen.MousePointer = 0

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
        .MarginLeft = 800: .MarginRight = 250
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
        .FormatString = "CI/RUC|Nombre|Factura|>Saldo|"
            
        .WordWrap = False
        AnchoEncabezado Pantalla:=True
        .MergeCells = flexMergeSpill
        .MergeCol(0) = True
    End With
      
End Sub

Private Sub AnchoEncabezado(Optional Pantalla As Boolean = False, Optional Impresora As Boolean = False)

    With vsConsulta
        
        If Pantalla Then
            .ColWidth(0) = 1260: .ColWidth(1) = 3750: .ColWidth(2) = 915: .ColWidth(3) = 1600
        End If
        
        If Impresora Then
            .ColWidth(0) = 1000: .ColWidth(1) = 2300: .ColWidth(2) = 750: .ColWidth(3) = 1400
        End If
        
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

    vsListado.Top = 100: vsListado.Left = 60
    vsListado.Height = Me.ScaleHeight - (vsListado.Top + Status.Height + picBotones.Height + 70)
    picBotones.Top = vsListado.Height + vsListado.Top + 70
    
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
    Set clsGeneral = Nothing
    End
    
End Sub

Private Sub AccionConsultar()

Dim aCantidad As Long
Dim aTexto As String

    On Error GoTo errConsultar
    If MsgBox("Confirma realizar la consulta de deudores al " & Format(gFechaServidor, "Long Date"), vbQuestion + vbYesNo, "Consultar") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    bCargarImpresion = True
       
    vsConsulta.Rows = 1: vsConsulta.Refresh
    aCantidad = 0
    
    cons = "Select Count(*) from Credito, Documento" _
           & " Where CreTipo IN (" & TipoCredito.Normal & ", " & TipoCredito.Gestor & ")" _
           & " And CreSaldoFactura > 0 " _
           & " And CreFactura = DocCodigo" _
           & " And DocMoneda = " & paMonedaPesos _
           & " And DocAnulado = 0"
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
    
    cons = "Select * from Credito, Documento, Cliente " _
                    & "Left Outer Join CPersona On CPeCliente = CliCodigo " _
                    & "Left Outer Join CEmpresa On CEmCliente = CliCodigo " _
           & " Where CreTipo IN (" & TipoCredito.Normal & ", " & TipoCredito.Gestor & ")" _
           & " And CreSaldoFactura > 0 " _
           & " And CreFactura = DocCodigo" _
           & " And DocMoneda = " & paMonedaPesos _
           & " And DocAnulado = 0" _
           & " And DocCliente = CliCodigo"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    Do While Not rsAux.EOF
        pbProgreso.Value = pbProgreso.Value + 1
        With vsConsulta
            .AddItem ""
            
            Select Case rsAux!CliTipo
                Case TipoCliente.Cliente
                    If Not IsNull(rsAux!CliCiRuc) Then .Cell(flexcpText, .Rows - 1, 0) = clsGeneral.RetornoFormatoCedula(rsAux!CliCiRuc)
                    .Cell(flexcpText, .Rows - 1, 1) = ArmoNombre(Format(rsAux!CPeApellido1, "#"), Format(rsAux!CPeApellido2, "#"), Format(rsAux!CPeNombre1, "#"), Format(rsAux!CPeNombre2, "#"))
                
                Case TipoCliente.Empresa
                    If Not IsNull(rsAux!CliCiRuc) Then .Cell(flexcpText, .Rows - 1, 0) = clsGeneral.RetornoFormatoRuc(rsAux!CliCiRuc)
                    If Not IsNull(rsAux!CEmNombre) Then .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!CEmNombre) Else .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!CEmFantasia)
            End Select
            
            .Cell(flexcpText, .Rows - 1, 2) = Trim(rsAux!DocSerie) & " " & Format(rsAux!DocNumero, "000000")
            .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!CreSaldoFactura, FormatoMonedaP)
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    With vsConsulta
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 3, , Colores.Rojo, Colores.Blanco, False, "TOTAL"
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
            .Columns = 2
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        EncabezadoListado vsListado, "Listado de Deudores en pesos al " & Format(gFechaServidor, "Long Date"), False
        vsListado.FileName = "Listado de Deudores"
            
        With vsConsulta
            .Redraw = False
            .FontSize = 6
            AnchoEncabezado Impresora:=True
            vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
            AnchoEncabezado Pantalla:=True
            .FontSize = 8
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

Function ArmoNombre(Ape1 As String, Ape2 As String, Nom1 As String, Nom2 As String) As String
    ArmoNombre = Trim(Ape1) & " " & Trim(Ape2)
    ArmoNombre = Trim(ArmoNombre) & ", " & Trim(Nom1) & " " & Trim(Nom2)
End Function

