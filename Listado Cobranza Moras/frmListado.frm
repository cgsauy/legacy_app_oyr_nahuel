VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListado 
   Caption         =   "Cobranza de Moras"
   ClientHeight    =   7530
   ClientLeft      =   1230
   ClientTop       =   1785
   ClientWidth     =   10830
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
   ScaleWidth      =   10830
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   4440
      TabIndex        =   4
      Top             =   1800
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
      Left            =   120
      TabIndex        =   7
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
      Zoom            =   70
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   11595
      TabIndex        =   8
      Top             =   6720
      Width           =   11655
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":0442
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   265
         Left            =   6000
         TabIndex        =   21
         Top             =   140
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
      TabIndex        =   6
      Top             =   7275
      Width           =   10830
      _ExtentX        =   19103
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
            Object.Width           =   10874
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Listado Cobranza de Moras"
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
      TabIndex        =   5
      Top             =   0
      Width           =   10335
      Begin AACombo99.AACombo cSucursal 
         Height          =   315
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
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
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   7920
         TabIndex        =   3
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
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
      Begin MSComCtl2.DTPicker tDesde 
         Height          =   315
         Left            =   3840
         TabIndex        =   22
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   9306113
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker tHasta 
         Height          =   315
         Left            =   5700
         TabIndex        =   23
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   9306113
         CurrentDate     =   37543
      End
      Begin VB.Label Label4 
         Caption         =   "&Hasta:"
         Height          =   255
         Left            =   5160
         TabIndex        =   25
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "&Desde:"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Moneda:"
         Height          =   255
         Left            =   7200
         TabIndex        =   2
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Sucursal:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   285
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RsAux As rdoResultset, rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean

Private Sub AccionLimpiar()
    cSucursal.Text = "": cMoneda.Text = ""
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
    Me.Refresh

End Sub


Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub cSucursal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tDesde
End Sub

Private Sub Label1_Click()
    Foco cSucursal
End Sub

Private Sub Label2_Click()
    Foco cMoneda
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
    
    'Cargo las sucursales y las monedas en loc combos-------------------------------------------------------------------------
    Cons = "Select SucCodigo, SucAbreviacion From Sucursal Where SucDcontado <> null Or SucDCredito <> Null"
    CargoCombo Cons, cSucursal
    
    Cons = "Select MonCodigo, MonSigno From Moneda"
    CargoCombo Cons, cMoneda
    BuscoCodigoEnCombo cMoneda, paMonedaFacturacion
    
    FechaDelServidor
    tDesde.Value = Format(gFechaServidor, "dd/mm/yyyy")
    tHasta.Value = Format(gFechaServidor, "dd/mm/yyyy")
    '----------------------------------------------------------------------------------------------------------------------------------
    
    bCargarImpresion = True
    vsListado.PaperSize = 1
    vsListado.MarginRight = 350
    vsListado.MarginLeft = 350
    vsListado.Orientation = orPortrait
    vsListado.Zoom = 100
    Exit Sub
    
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    vsListado.MarginTop = 500
    vsListado.MarginBottom = 600
    
    With vsConsulta
        .OutlineBar = flexOutlineBarNone ' flexOutlineBarComplete
        .OutlineCol = 0
        .MultiTotals = True
        .SubtotalPosition = flexSTBelow
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<Sucursal|<Fecha|<N.Débito|>Mora Neta|>IVA Mora|>Total Mora|<Recibo"
            
        .WordWrap = False
        .ColWidth(0) = 0: .ColWidth(1) = 550: .ColWidth(2) = 800
        .ColWidth(3) = 900: .ColWidth(4) = 750: .ColWidth(5) = 1100: .ColWidth(6) = 700: .ColWidth(6) = 1100
         
        .MergeCells = flexMergeSpill
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
    
    vsConsulta.Top = vsListado.Top
    vsConsulta.Width = vsListado.Width
    vsConsulta.Height = vsListado.Height
    vsConsulta.Left = vsListado.Left
    
    picBotones.Width = vsListado.Width
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
Dim aIDSucursal As Long, aTxtSucursal As String
Dim rs1 As rdoResultset

Dim m_NETO As Currency, m_IVA  As Currency
Dim mT_NETO As Currency, mT_IVA  As Currency

    On Error GoTo errConsultar
    If Not ValidoCampos Then Exit Sub
    
    Screen.MousePointer = 11
    vsConsulta.ZOrder 0
    Me.Refresh
    bCargarImpresion = True
    
    aIDSucursal = 0
    
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    'cons = "Select Count(*) From Documento " & _
            " Where DocTipo = " & TipoDocumento.ReciboDePago & _
            " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) & _
            " And DocIVA Is NOT Null And DocIVA <> 0 And DocAnulado = 0 " & _
            " And DocFecha Between '" & Format(tDesde.Value, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Value, "mm/dd/yyyy 23:59:59") & "'"

    Cons = "Select Count(*) From Documento " & _
            " Where DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) & _
            " And  (( DocTipo = " & TipoDocumento.NotaDebito & " )" & _
               " OR (  DocTipo = " & TipoDocumento.ReciboDePago & " And DocIVA < 0 ) )" & _
            " And DocFecha Between '" & Format(tDesde.Value, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Value, "mm/dd/yyyy 23:59:59") & "'"

    If cSucursal.ListIndex <> -1 Then Cons = Cons & " And DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then
        If RsAux(0) = 0 Then
            MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
            RsAux.Close: Screen.MousePointer = 0: vsConsulta.Rows = 1: Exit Sub
        End If
        pbProgreso.Max = RsAux(0)
    End If
    RsAux.Close
    '-------------------------------------------------------------------------------------------------------------------------------------------
    
'    cons = "Select DocSucursal, DocAnulado, DocFecha, DocSerie, DocNumero, (Sum(DPaMora) - DocIVA) as DPAMoraNeta, DocIVA, Sum(DPaMora) as DPaMoraTotal" & _
            " From Documento, DocumentoPago " & _
            " Where DocCodigo = DPaDocQSalda" & _
            " And  DocTipo = " & TipoDocumento.ReciboDePago & _
            " And DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) & _
            " And DocIVA Is NOT Null And DocIVA <> 0 And DocAnulado = 0" & _
            " And DocFecha Between '" & Format(tDesde.Value, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Value, "mm/dd/yyyy 23:59:59") & "'"
    
    For I = 1 To 2
        
        If I = 2 Then Cons = Cons & " UNION ALL " Else Cons = ""
        
        Cons = Cons & "Select Moras.DocCodigo, Moras.DocTotal *" & IIf(I = 1, 1, -1) & " as DocTotal, Moras.DocIVA *" & IIf(I = 1, 1, -1) & " as DocIVA, " & _
                        " Moras.DocSucursal as DocSucursal, Moras.DocAnulado as DocAnulado, Moras.DocFecha as DocFecha, Moras.DocSerie as DocSerie, Moras.DocNumero as DocNumero, Recibo.DocSerie as RecSerie, Recibo.DocNumero as RecNumero " & _
            " From Documento Moras, DocumentoPago, Documento as Recibo " & _
            " Where Moras.DocCodigo = DPaDocASaldar And DPaDocQSalda = Recibo.DocCodigo" & _
            " And  Moras.DocTipo = " & TipoDocumento.NotaDebito & _
            " And Moras.DocMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
            
            If cSucursal.ListIndex <> -1 Then Cons = Cons & " And Moras.DocSucursal = " & cSucursal.ItemData(cSucursal.ListIndex)
            
            If I = 1 Then
                Cons = Cons & " And Recibo.DocIVA = 0 And Moras.DocFecha Between '" & Format(tDesde.Value, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Value, "mm/dd/yyyy 23:59:59") & "'"
            Else
                Cons = Cons & " And Recibo.DocIVA < 0 And Recibo.DocFecha Between '" & Format(tDesde.Value, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Value, "mm/dd/yyyy 23:59:59") & "'"
            End If
                
    Next
    
    
    Cons = Cons & " UNION ALL " & _
        "SELECT Moras.DocCodigo, Moras.DocTotal *-1 as DocTotal, Moras.DocIVA *-1 as DocIVA,  Moras.DocSucursal as DocSucursal, " & _
        "Moras.DocAnulado as DocAnulado, Moras.DocFecha as DocFecha, Moras.DocSerie as DocSerie, Moras.DocNumero as DocNumero, " & _
        "Recibo.DocSerie as RecSerie, Recibo.DocNumero as RecNumero " & _
        "FROM Documento Moras INNER JOIN DocumentoPago ON Moras.DocCodigo = DPaDocASaldar " & _
        "INNER JOIN Documento Recibo ON Recibo.DocCodigo = DPaDocQSalda " & _
        "WHERE Moras.DocFecha Between '" & Format(tDesde.Value, "mm/dd/yyyy 00:00:00") & "' And '" & Format(tHasta.Value, "mm/dd/yyyy 23:59:59") & "' AND Moras.DocTipo = " & TipoDocumento.NotaCreditoMora
    
    Cons = Cons & " Order by DocSucursal, DocSerie, DocNumero "
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    If RsAux.EOF Then
        RsAux.Close
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        Screen.MousePointer = 0: vsConsulta.Rows = 1: Exit Sub
    End If
    
    Dim xIDAnterior As Long: xIDAnterior = 0
    vsConsulta.Rows = 1: vsConsulta.Refresh
    Do While Not RsAux.EOF
        If pbProgreso.Value + 1 <= pbProgreso.Max Then pbProgreso.Value = pbProgreso.Value + 1
        
        If aIDSucursal <> RsAux!DocSucursal Then
            If aIDSucursal <> 0 Then
                With vsConsulta
                    .AddItem ""
                    .Cell(flexcpText, .Rows - 1, 3) = Format(m_NETO, FormatoMonedaP)
                    .Cell(flexcpText, .Rows - 1, 4) = Format(m_IVA, FormatoMonedaP)
                    .Cell(flexcpText, .Rows - 1, 5) = Format(m_NETO + m_IVA, FormatoMonedaP)
                    .Cell(flexcpBackColor, .Rows - 1, 3, , 5) = Colores.Obligatorio
                    .AddItem ""
                End With
                mT_NETO = mT_NETO + m_NETO: mT_IVA = mT_IVA + m_IVA
                m_NETO = 0: m_IVA = 0
            End If
        
            aIDSucursal = RsAux!DocSucursal
            Cons = "Select * from Sucursal Where SucCodigo = " & RsAux!DocSucursal
            Set rs1 = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not rs1.EOF Then aTxtSucursal = Trim(rs1!SucAbreviacion) Else aTxtSucursal = ""
            rs1.Close
            
            vsConsulta.AddItem aTxtSucursal
            vsConsulta.Cell(flexcpBackColor, vsConsulta.Rows - 1, 0, , 6) = Colores.osGris
            vsConsulta.Cell(flexcpForeColor, vsConsulta.Rows - 1, 0, , 5) = vbWhite
            vsConsulta.Cell(flexcpFontBold, vsConsulta.Rows - 1, 0, , 5) = True

        End If
        
        With vsConsulta
            .AddItem aTxtSucursal
            .Cell(flexcpText, .Rows - 1, 6) = Trim(RsAux!RecSerie) & "" & Trim(RsAux!RecNumero)
            .Cell(flexcpText, .Rows - 1, 1) = " "
            If RsAux!DocCodigo <> xIDAnterior Then
                .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!DocFecha, "dd/mm")
                .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!DocSerie) & "" & Trim(RsAux!DocNumero)
                
            
                .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!DocTotal - RsAux!DocIVA, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!DocIVA, FormatoMonedaP)
                .Cell(flexcpText, .Rows - 1, 5) = Format(RsAux!DocTotal, FormatoMonedaP)
            End If
            
            If RsAux!DocAnulado Then
                .Cell(flexcpBackColor, .Rows - 1, 1, , 5) = vbButtonFace 'Colores.Gris
            Else
                If RsAux!DocCodigo <> xIDAnterior Then
                    m_NETO = m_NETO + .Cell(flexcpValue, .Rows - 1, 3)
                    m_IVA = m_IVA + .Cell(flexcpValue, .Rows - 1, 4)
                End If
            End If
            
            xIDAnterior = RsAux!DocCodigo
            
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    With vsConsulta
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 3) = Format(m_NETO, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 4) = Format(m_IVA, FormatoMonedaP)
        .Cell(flexcpText, .Rows - 1, 5) = Format(m_NETO + m_IVA, FormatoMonedaP)
        .Cell(flexcpBackColor, .Rows - 1, 3, , 5) = Colores.osGris
        .Cell(flexcpForeColor, .Rows - 1, 3, , 5) = vbWhite

        mT_NETO = mT_NETO + m_NETO: mT_IVA = mT_IVA + m_IVA
        If cSucursal.ListIndex = -1 Then
            .AddItem ""
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 3) = Format(mT_NETO, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 4) = Format(mT_IVA, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 5) = Format(mT_NETO + mT_IVA, FormatoMonedaP)
            .Cell(flexcpBackColor, .Rows - 1, 3, , 5) = Colores.osGris
            .Cell(flexcpForeColor, .Rows - 1, 3, , 5) = vbWhite
        End If
    End With
    
    pbProgreso.Value = 0: Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al realizar la consulta de datos.", Err.Description
End Sub

Private Sub tDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco tHasta
End Sub

Private Sub tHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco cMoneda
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
            vsListado.Columns = 2
            
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        EncabezadoListado vsListado, "Cobranza de Moras en " & Trim(cMoneda.Text) & " - " & tDesde.Value & " al " & tHasta.Value, False
        vsListado.FileName = "Cobranza de Moras"
        
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
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

Private Function ValidoCampos() As Boolean
    
    ValidoCampos = False
    
    If Not IsDate(tDesde.Value) Or Not IsDate(tHasta.Value) Then
        MsgBox "Debe ingresar un fecha para realizar la consulta de datos.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    If tDesde.Value > tHasta.Value Then
        MsgBox "El rango de fechas ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Debe seleccionar la moneda para realizar la consulta de datos.", vbExclamation, "ATENCIÓN"
        Exit Function
    End If
    
    ValidoCampos = True
    
End Function
