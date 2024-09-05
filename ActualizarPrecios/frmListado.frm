VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Begin VB.Form frmListado 
   Caption         =   "Actualización de Precios"
   ClientHeight    =   7335
   ClientLeft      =   1725
   ClientTop       =   2610
   ClientWidth     =   11175
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
   ScaleHeight     =   7335
   ScaleWidth      =   11175
   Begin VB.Frame fValores 
      Caption         =   "Valores Globales"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   4560
      TabIndex        =   36
      Top             =   120
      Width           =   6495
      Begin VB.OptionButton oTC 
         Caption         =   "Por &TC M/E"
         Height          =   195
         Left            =   2640
         TabIndex        =   11
         Top             =   600
         Width           =   1155
      End
      Begin VB.OptionButton oPorcentaje 
         Caption         =   "% Au&mento"
         Height          =   195
         Left            =   2640
         TabIndex        =   10
         Top             =   300
         Width           =   1155
      End
      Begin VB.TextBox tTCTasa 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4560
         TabIndex        =   14
         Top             =   540
         Width           =   675
      End
      Begin VB.CommandButton bGrabar 
         Caption         =   "&Grabar"
         Height          =   300
         Left            =   5460
         TabIndex        =   37
         Top             =   520
         Width           =   795
      End
      Begin VB.CommandButton bCalcular 
         Caption         =   "Aplica&r"
         Height          =   300
         Left            =   5460
         TabIndex        =   15
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox tGVigencia 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   900
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox tGAumento 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   240
         Width           =   675
      End
      Begin AACombo99.AACombo cGPlan 
         Height          =   315
         Left            =   900
         TabIndex        =   9
         Top             =   555
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
      Begin AACombo99.AACombo cTCMoneda 
         Height          =   315
         Left            =   3840
         TabIndex        =   13
         Top             =   540
         Width           =   735
         _ExtentX        =   1296
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
      Begin VB.Label Label6 
         Caption         =   "&Vigencia:"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   255
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "&Plan:"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   615
         Width           =   435
      End
   End
   Begin ComctlLib.TabStrip tabPrecios 
      Height          =   1635
      Left            =   120
      TabIndex        =   34
      Top             =   1620
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2884
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Precios &Contado"
            Key             =   "contado"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   " Cuo&tas  "
            Key             =   "cuotas"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fFiltros 
      Caption         =   "Filtros de Consulta"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   60
      TabIndex        =   33
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtValorDolar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox tGrupo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   600
         Width           =   3015
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   1020
         TabIndex        =   1
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
      Begin VB.Label Label2 
         Caption         =   "Valor Dólar:"
         Height          =   255
         Left            =   2460
         TabIndex        =   2
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Pre&cios en:"
         Height          =   195
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   795
      End
      Begin VB.Label lIdFiltro 
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   795
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsContado 
      Height          =   4575
      Left            =   2880
      TabIndex        =   28
      Top             =   1320
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8070
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
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
      TabIndex        =   30
      Top             =   1920
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
      ScaleWidth      =   8055
      TabIndex        =   31
      Top             =   6720
      Width           =   8115
      Begin VB.CheckBox chVista 
         DownPicture     =   "frmListado.frx":030A
         Height          =   310
         Left            =   4440
         Picture         =   "frmListado.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Picture         =   "frmListado.frx":093E
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmListado.frx":0DB8
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmListado.frx":0EA2
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Picture         =   "frmListado.frx":0F8C
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Picture         =   "frmListado.frx":11C6
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Picture         =   "frmListado.frx":12C8
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5400
         Picture         =   "frmListado.frx":168E
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmListado.frx":1790
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Picture         =   "frmListado.frx":1A92
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Picture         =   "frmListado.frx":1DD4
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Picture         =   "frmListado.frx":20D6
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin ComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   5820
         TabIndex        =   32
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
      TabIndex        =   29
      Top             =   7080
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Key             =   "rebaja"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   15637
            Key             =   "help"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsCuotas 
      Height          =   4575
      Left            =   8640
      TabIndex        =   35
      Top             =   1140
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   8070
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
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
End
Attribute VB_Name = "frmListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typCoef
    Plan As Integer
    TCuota As Integer
    Coeficiente As Currency
    QCtas As Integer
End Type

Dim arrCoef() As typCoef
Dim arrRND() As String

Private aTexto As String
Dim aValor As Long, aVCuota As Currency

Dim I As Long
Dim mFormato As String

Private Sub AccionLimpiar()
    tGrupo.Text = ""
    
    
    tGVigencia.Text = Format(Date, "dd/mm/yyyy 19:30")
    tGAumento.Text = "" ': cGPlan.Text = ""
    
    bGrabar.Enabled = False
    vsContado.Rows = 1: vsCuotas.Rows = 1
    cMoneda.Enabled = True
    
End Sub

Private Sub bCalcular_Click()
    
    If vsContado.Rows = 1 Then
        MsgBox "Debe cargar los datos de los artículos a procesar.", vbExclamation, "Faltan Artículos"
        Exit Sub
    End If
    
    If Not IsNumeric(tGAumento.Text) Then tGAumento.Text = 0
    If oTC.Value Then
        If cTCMoneda.ListIndex = -1 Then
            MsgBox "Para hacer los cálculos por tasa de cambio en moneda extranjera debe seleccionar la moneda.", vbExclamation, "Falta Moneda TC"
            Foco cTCMoneda: Exit Sub
        End If
        If Not IsNumeric(tTCTasa.Text) Then
            MsgBox "Para hacer los cálculos por tasa de cambio en moneda extranjera debe ingresar la tasa.", vbExclamation, "Falta Tasa de Cambio"
            Foco tTCTasa: Exit Sub
        End If
    End If
    vsContado.Editable = False: vsCuotas.Editable = False
    
    Dim bAll As Boolean
    bAll = True
    If MsgBox("Los cálculos se van a realizar para los artículos que no tienen precio." & vbCrLf & _
                   "Para recalcular todos los precios presione NO.", vbQuestion + vbYesNo, "Sólo a Artículo sin Precios Ingresados") = vbYes Then bAll = False
    
    AccionCalcular bAll
    
    vsContado.Editable = True: vsCuotas.Editable = True
    
End Sub

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()
    AccionConsultar
End Sub

Private Sub bGrabar_Click()
    
    If Not ValidoDatos Then Exit Sub
    If MsgBox("Confirma actualizar los precios ingresados.", vbQuestion + vbYesNo, "Grabar Precios") = vbNo Then Exit Sub
    
    AccionGrabar
    
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


Private Sub cGPlan_GotFocus()
    With cGPlan: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub cGPlan_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If oPorcentaje.Value Then Foco tGAumento
        If oTC.Value Then Foco cTCMoneda
    End If
End Sub

Private Sub chVista_Click()
    
    If chVista.Value = 0 Then
        'vsContado.ZOrder 0
        vsListado.Visible = False
    Else
        AccionImprimir
        vsListado.Visible = True
        vsListado.ZOrder 0
    End If

End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tGrupo
End Sub

Private Sub cTCMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyReturn And cTCMoneda.ListIndex <> -1 Then
        If Not IsNumeric(tTCTasa.Text) Then
            Dim mMoneda1 As Integer, mMoneda2 As Integer
        
            mMoneda1 = cTCMoneda.ItemData(cTCMoneda.ListIndex)
            mMoneda2 = cMoneda.ItemData(cMoneda.ListIndex)
                        
            tTCTasa.Text = TasadeCambio(mMoneda1, mMoneda2, Now, TipoTC:=paTipoTC)
        End If
        Foco tTCTasa
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()
    On Error GoTo ErrLoad
    lIdFiltro.Tag = 1
    'InicializoArrayCoef
    
    InicializoArrayRND
    dis_CargoArrayMonedas
    
    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    picBotones.BorderStyle = vbBSNone
    InicializoGrillas
    
    txtValorDolar.Text = 1
    Cons = "select dbo.Dolar(2, getdate())"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then txtValorDolar.Text = RsAux(0)
    RsAux.Close
    
    
    Cons = "Select PlaCodigo, PlaNombre from TipoPlan Order by PlaNombre"
    CargoCombo Cons, cGPlan
    
    Cons = "Select MonCodigo, MonSigno from Moneda Where MonFactura = 1 Order by MonSigno"
    CargoCombo Cons, cMoneda
    Cons = "Select MonCodigo, MonSigno from Moneda Order by MonSigno"
    CargoCombo Cons, cTCMoneda
    BuscoCodigoEnCombo cMoneda, CLng(paMonedaPesos)
    
    AccionLimpiar
    oPorcentaje.Value = True
    
    With vsListado
        .PaperSize = 1
        .Orientation = orPortrait
        .PhysicalPage = True
        .Zoom = 100
        .MarginLeft = 500: .MarginRight = 250
        .MarginBottom = 750: .MarginTop = 750
    End With
    
    tabPrecios.SelectedItem = tabPrecios.Tabs("contado")
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoArrayRND()
    ReDim arrRND(0)
'    For I = 0 To 9
'        ReDim Preserve arrRND(I)
'        arrRND(I) = 0
'    Next
ReDim arrRND(9)
arrRND(0) = -1
arrRND(1) = -1
arrRND(2) = -2
arrRND(3) = 2
arrRND(4) = 1
arrRND(5) = 0
arrRND(6) = -1
arrRND(7) = 1
arrRND(8) = 2
arrRND(9) = 1

'0 -1
'1 -1
'2 -2
'3 +2
'4 +1
'6 -1
'7 +1
'8 +2
'9 +1

End Sub
Private Sub InicializoGrillas()

    On Error Resume Next
    With vsContado
        .Cols = 1: .Rows = 1:
        .FormatString = "<Artículo|>Contado (V)|<Plan (V)|Fecha Vigencia|>Contado (M/E)|<Plan (N)|>Contado Nuevo|>% Aum.|^Costo|>Margen|>Stock| "
            
        .WordWrap = False: .MergeCells = flexMergeSpill
        .ColWidth(0) = 3500: .ColWidth(1) = 1000: .ColWidth(2) = 750: .ColWidth(3) = 1290: .ColWidth(5) = 800
        .ColWidth(6) = 1300: .ColWidth(7) = 700
        .ColWidth(8) = 1300
        .ColWidth(10) = 800
        
        .GridLinesFixed = flexGridFlat
        .GridLines = flexGridFlat
        
        If Not miConexion.AccesoAlMenu("Estadísticas") Then
            .ColHidden(8) = True
            .ColHidden(9) = True
            .Cell(flexcpText, 0, 8) = ""
            .Cell(flexcpText, 0, 9) = ""
        End If
    End With
    
    With vsCuotas
        .Cols = 1: .Rows = 1
        .FormatString = "<Artículo|>Contado (N)"
            
        .WordWrap = False: .MergeCells = flexMergeSpill
        .ColWidth(0) = 3500: .ColWidth(1) = 1000
        
        .Cell(flexcpData, 0, .Cols - 1) = paTipoCuotaContado
        'Agrego las columnas p/cada tipocuota-----------------------------------------------------
        Dim aData As String
        Cons = "Select * from TipoCuota " & _
                    " Where TCuVencimientoE is null And TCuDeshabilitado is null" & _
                    " And TCuCodigo <> " & paTipoCuotaContado & _
                    " And TCuCodigo In (Select Distinct(CoeTipoCuota) from Coeficiente Where CoeCoeficiente <> 1) " & _
                    " Order by TCuOrden"
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        Do While Not RsAux.EOF
            aData = RsAux!TCuCodigo
            .Cols = .Cols + 1
            .Cell(flexcpText, 0, .Cols - 1) = Trim(RsAux!TCuAbreviacion)
            .Cell(flexcpData, 0, .Cols - 1) = aData
            If .ColWidth(.Cols - 1) < 500 Then .ColWidth(.Cols - 1) = 500
            .ColWidth(.Cols - 1) = 700
            .ColAlignment(.Cols - 1) = flexAlignRightCenter
            RsAux.MoveNext
        Loop
        RsAux.Close
        '--------------------------------------------------------------------------------------------------
        .FixedCols = 1
        .GridLinesFixed = flexGridFlat
        .GridLines = flexGridFlat
        .ExtendLastCol = False
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
    fFiltros.Left = 60: fFiltros.Top = 60
    
    fValores.Top = fFiltros.Top: fValores.Left = fFiltros.Left + fFiltros.Width + 60
    fValores.Width = Me.ScaleWidth - (fValores.Left + 60)
    
    With vsListado
        .Top = fFiltros.Top + fFiltros.Height + 60
        .Height = Me.ScaleHeight - (.Top + Status.Height + picBotones.Height + 70)
        picBotones.Top = .Height + vsListado.Top + 70
        
        .Width = Me.ScaleWidth - (.Left * 2)
        .Left = fFiltros.Left
    End With
    
    tabPrecios.Top = vsListado.Top: tabPrecios.Left = vsListado.Left
    tabPrecios.Width = vsListado.Width: tabPrecios.Height = vsListado.Height
    
    With tabPrecios
        vsContado.Top = .ClientTop:  vsContado.Left = .ClientLeft
        vsContado.Width = .ClientWidth: vsContado.Height = .ClientHeight
    End With
    With vsContado
        vsCuotas.Top = .Top: vsCuotas.Left = .Left
        vsCuotas.Width = .Width: vsCuotas.Height = .Height
    End With
    
    picBotones.Width = vsListado.Width
    pbProgreso.Width = picBotones.Width - pbProgreso.Left - 150
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

Private Sub AccionCalcular(Optional bTodos As Boolean = False)

Dim miAumento As Currency, miValor As Currency
Dim miPlan As Long

    On Error GoTo errCalcular
    Screen.MousePointer = 11
    
    miPlan = 0
    If cGPlan.ListIndex <> -1 Then miPlan = cGPlan.ItemData(cGPlan.ListIndex)
    miAumento = 1 + (CCur(tGAumento.Text) / 100)
    
    If oTC.Value Then CargoPreciosME cTCMoneda.ItemData(cTCMoneda.ListIndex), bTodos
    
    For I = 1 To vsContado.Rows - 1
        With vsContado
            '<Artículo|>Contado (V)|<Plan (V)|Fecha Vigencia|>Contado ME|<Pan Nuevo|>Contado Nuevo|>% Aum.|
            If bTodos Or (Not bTodos And Trim(.Cell(flexcpText, I, 6)) = "") Then
                If miPlan = 0 And .Cell(flexcpText, I, 2) <> "" Then
                    .Cell(flexcpText, I, 5) = .Cell(flexcpText, I, 2)
                    aValor = .Cell(flexcpData, I, 2)
                Else
                    .Cell(flexcpText, I, 5) = Trim(cGPlan.Text)
                    aValor = miPlan
                End If
                 .Cell(flexcpData, I, 5) = aValor
                 
                If oPorcentaje.Value Then           'Ajustes por % de aumento   ---------------------
                    If Trim(.Cell(flexcpText, I, 1)) <> "" Then
                        miValor = .Cell(flexcpValue, I, 1) * miAumento
                    
                        .Cell(flexcpText, I, 6) = FormatoImporte(miValor)
                        .Cell(flexcpForeColor, I, 6) = ColorAjuste(.Cell(flexcpValue, I, 6))
                    
                        .Cell(flexcpText, I, 7) = (miAumento - 1) * 100
                        .Cell(flexcpBackColor, I, 7) = .BackColor
                    Else
                        .Cell(flexcpBackColor, I, 7) = vbButtonFace
                    End If
                Else                        'Ajustes por TC ME   ----------------------------------------------
                    If Trim(.Cell(flexcpText, I, 4)) <> "" Then     'Precio ME
                        miValor = .Cell(flexcpValue, I, 4) * CCur(tTCTasa.Text)
                        .Cell(flexcpText, I, 6) = FormatoImporte(miValor)
                        .Cell(flexcpForeColor, I, 6) = ColorAjuste(.Cell(flexcpValue, I, 6))
                
                        If Trim(.Cell(flexcpText, I, 1)) <> "" Then
                            .Cell(flexcpText, I, 7) = Format((.Cell(flexcpValue, I, 6) * 100) / .Cell(flexcpValue, I, 1) - 100, "0.00")
                            .Cell(flexcpBackColor, I, 7) = .BackColor
                        Else
                            .Cell(flexcpBackColor, I, 7) = vbButtonFace
                        End If
                        
                    Else
                        .Cell(flexcpBackColor, I, 7) = vbButtonFace
                    End If
                End If
                              
                
                CalculoValoresCuotas I, .Cell(flexcpData, I, 5), .Cell(flexcpValue, I, 6)
            End If
        End With
    Next
    
    Screen.MousePointer = 0
    bGrabar.Enabled = True
    Exit Sub
    
errCalcular:
    clsGeneral.OcurrioError "Error al realizar los cálculos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub CalculoValoresCuotas(idRow As Long, idPlan As Long, miContado As Currency)

Dim idTCuota As Long, j As Integer
Dim miCoef As Currency, miQCtas As Integer

    With vsContado
        If Not .ColHidden(9) And .Cell(flexcpValue, idRow, 8) <> 0 Then
            .Cell(flexcpText, idRow, 9) = Format((.Cell(flexcpValue, idRow, 6) * 100 / .Cell(flexcpValue, idRow, 8)) - 100, "0.00")
        End If
    End With

    With vsCuotas
        .Cell(flexcpText, idRow, 1) = FormatoImporte(miContado) 'Format(miContado, "#,##0")
        aValor = 1: .Cell(flexcpData, idRow, 1) = aValor
        .Cell(flexcpForeColor, idRow, 1) = ColorAjuste(.Cell(flexcpValue, idRow, 1))
        .Cell(flexcpBackColor, idRow, 1) = .BackColor
        
        'Veo en que columna esta la cuota para ingresar el valor
        For j = 2 To .Cols - 1
            idTCuota = Val(.Cell(flexcpData, 0, j))
            If arrCoef_Coeficiente(idPlan, idTCuota, miCoef, miQCtas) Then
                If miCoef <> 1 Then
                    'aVCuota = Format((miContado * miCoef) / miQCtas, "#,##0")
                    aVCuota = (miContado * miCoef) / miQCtas
                
                    .Cell(flexcpText, idRow, j) = FormatoImporte(aVCuota) ' Format(aVCuota, "#,##0")
                    aValor = miQCtas: .Cell(flexcpData, idRow, j) = aValor
                    .Cell(flexcpForeColor, idRow, j) = ColorAjuste(.Cell(flexcpValue, idRow, j))
                    .Cell(flexcpBackColor, idRow, j) = .BackColor
                End If
            Else
                'Saco el precio con el plan Anterior y lo pongo deshabilitado
                Dim oldPlan As Long
                oldPlan = vsContado.Cell(flexcpData, idRow, 2)
                
                If arrCoef_Coeficiente(oldPlan, idTCuota, miCoef, miQCtas) Then
                    'aVCuota = Format((miContado * miCoef) / miQCtas, "#,##0")
                    aVCuota = (miContado * miCoef) / miQCtas
                
                    .Cell(flexcpText, idRow, j) = FormatoImporte(aVCuota) ' Format(aVCuota, "#,##0")
                    aValor = miQCtas: .Cell(flexcpData, idRow, j) = aValor
                    .Cell(flexcpForeColor, idRow, j) = ColorAjuste(.Cell(flexcpValue, idRow, j))
                    .Cell(flexcpBackColor, idRow, j) = Colores.clCeleste
                Else
                    .Cell(flexcpText, idRow, j) = ""
                    aValor = 0: .Cell(flexcpData, idRow, j) = aValor
                    .Cell(flexcpBackColor, idRow, j) = .BackColorFixed
                End If
            End If
            
        Next
    End With
    
End Sub

Private Sub AccionConsultar(Optional porArticulo As Long = 0, Optional porFiltros As String = "", Optional mPlan As String = "")

    On Error GoTo errConsultar
    If cMoneda.ListIndex = -1 Then
        MsgBox "Seleccione la moneda para consultar los precios.", vbExclamation, "Falta Moneda"
        Foco cMoneda: Exit Sub
    End If
    
    Screen.MousePointer = 11
    
    If vsContado.Rows = 1 Then
        Dim mMoneda As Long
        mMoneda = Val(cMoneda.ItemData(cMoneda.ListIndex))
        InicializoArrayCoef mMoneda
        mFormato = dis_arrMonedaProp(mMoneda, enuMoneda.pRedondeo)
    End If
    
    bGrabar.Enabled = False
    
    Dim aQ As Long
    
    Cons = "Select * " & _
          " From Articulo Left Outer Join PrecioVigente On ArtID = PViArticulo " & _
                                                                " And PViTipoCuota = " & paTipoCuotaContado & _
                                                                " And PViMoneda = " & cMoneda.ItemData(cMoneda.ListIndex)
    
    Cons = Cons & " Left Outer Join TipoPlan On PViPlan = PlaCodigo " & _
                  " Left Outer Join SumaStockTotal On ArtID = SSTArticulo "

    If Trim(mPlan) <> "" Then
        If Mid(mPlan, 1, 1) = "1" Then
            Cons = Cons & " And PViPlan = " & Mid(mPlan, 2)
        Else
            Cons = Cons & " And PViPlan <> " & Mid(mPlan, 2)
        End If
    End If
    
    Cons = Cons & " Where ArtEnUso = 1 "
               
    If porArticulo > 0 Then
        Cons = Cons & " And ArtCodigo = " & porArticulo
    Else
        Cons = Cons & porFiltros
    End If
    
    aQ = DCount(Cons)
    If aQ = 0 Then
        MsgBox "No hay datos a procesar para los filtros ingresados.", vbInformation, "No hay datos"
        Screen.MousePointer = 0: Exit Sub
    End If
    
    Cons = clsGeneral.Replace(Cons, "*", "ArtCodigo, ArtNombre, ArtID, PViPrecio, PViPlan, PViVigencia, PlaNombre, PlaCodigo, SSTCantidad, dbo.CostoArticulo(ArtId, " & txtValorDolar.Text & ") as Costo ")
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    pbProgreso.Value = 0: pbProgreso.Max = aQ
    
    Do While Not RsAux.EOF
        With vsContado
            If Not EnLista(RsAux!ArtID) Then
                'Si el PlaNombre es Nulo y el PViPlan NO, no lo cargo porque el precio vigente no es del mismo plan consultado
                If Not (IsNull(RsAux!PlaCodigo) And Not IsNull(RsAux!PViPlan)) Then
                    
                    '<Artículo|>Contado (V)|<Plan (V)|Fecha Vigencia|<Pan Nuevo|>Contado Nuevo|>% Aum.|
                    .AddItem ""
                    .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
                    aValor = RsAux!ArtID: .Cell(flexcpData, .Rows - 1, 0) = aValor
                    
                    vsCuotas.AddItem ""
                    vsCuotas.Cell(flexcpText, vsCuotas.Rows - 1, 0) = Format(RsAux!ArtCodigo, "(#,000,000)") & " " & Trim(RsAux!ArtNombre)
                    vsCuotas.Cell(flexcpData, vsCuotas.Rows - 1, 0) = aValor
                    
                    If Not IsNull(RsAux!PViPrecio) Then
                        .Cell(flexcpText, .Rows - 1, 1) = FormatoImporte(RsAux!PViPrecio)
                    Else
                        .Cell(flexcpText, .Rows - 1, 1) = " "
                        .Cell(flexcpBackColor, .Rows - 1, 1) = vbButtonFace
                    End If
                    
                    If Not IsNull(RsAux!PlaNombre) Then
                        .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!PlaNombre)
                        aValor = RsAux!PlaCodigo: .Cell(flexcpData, .Rows - 1, 2) = aValor
                    
                        .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!PViVigencia, "dd/mm/yy hh:mm")
                    Else
                        .Cell(flexcpBackColor, .Rows - 1, 2) = vbButtonFace
                        .Cell(flexcpBackColor, .Rows - 1, 3) = vbButtonFace
                    End If
                    
                    'Costo
                    If Not IsNull(RsAux!Costo) And Not .ColHidden(8) Then
                        .Cell(flexcpText, .Rows - 1, 8) = FormatoImporte(RsAux!Costo) 'Format(rsAux!Costo, "#,##0.00")
                        .Cell(flexcpAlignment, .Rows - 1, 8) = flexAlignRightCenter
                    End If
                    
                    .Cell(flexcpText, .Rows - 1, 10) = Format(RsAux!SSTCantidad, "#,##0")
                End If
            End If
        End With
        RsAux.MoveNext
        pbProgreso.Value = pbProgreso.Value + 1
    Loop
    RsAux.Close
    
    With vsContado
        .Redraw = False
        .Select 1, 0, .Rows - 1
        .Sort = flexSortGenericAscending
        .Select 1, 0, 1, 0
        .Redraw = True
    End With
    With vsCuotas
        .Redraw = False
        .Select 1, 0, .Rows - 1
        .Sort = flexSortGenericAscending
        .Select 1, 0, 1, 0
        .Redraw = True
    End With
    
    If vsContado.Rows > 1 Then cMoneda.Enabled = False
    
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    'If Trim(tGAumento.Text) = "" Then Foco tGAumento
    Exit Sub
    
errConsultar:
    clsGeneral.OcurrioError "Error al realizar la consulta de datos.", Err.Description
    vsContado.Redraw = True
    vsCuotas.Redraw = True
    Screen.MousePointer = 0
End Sub

Function EnLista(idArticulo As Long, Optional nroRow As Long) As Boolean
Dim j As Integer
    
    EnLista = False
    With vsContado
        For j = 1 To .Rows - 1
            If .Cell(flexcpData, j, 0) = idArticulo Then
                nroRow = j
                EnLista = True: Exit For
            End If
        Next
    End With
    
End Function


Private Sub oPorcentaje_Click()
    
    tGAumento.Enabled = True: tGAumento.BackColor = vbWindowBackground
    cTCMoneda.Enabled = False: cTCMoneda.BackColor = vbButtonFace
    tTCTasa.Enabled = False: tTCTasa.BackColor = vbButtonFace
End Sub

Private Sub oPorcentaje_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco tGAumento
End Sub

Private Sub oTC_Click()
    tGAumento.Enabled = False: tGAumento.BackColor = vbButtonFace
    cTCMoneda.Enabled = True: cTCMoneda.BackColor = vbWindowBackground
    tTCTasa.Enabled = True: tTCTasa.BackColor = vbWindowBackground
End Sub

Private Sub oTC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Foco cTCMoneda
End Sub

Private Sub tabPrecios_Click()
    
    Select Case tabPrecios.SelectedItem.Key
        Case "contado": vsContado.ZOrder 0
        Case "cuotas": vsCuotas.ZOrder 0
    End Select
    
End Sub

Private Sub tGAumento_GotFocus()
    With tGAumento: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tGAumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bCalcular.SetFocus
End Sub

Private Sub tGrupo_Change()
    tGrupo.Tag = 0
End Sub

Private Sub tGrupo_GotFocus()
    With tGrupo: .SelStart = 0: .SelLength = Len(.Text): End With
    Status.Panels("help").Text = "[F1]- Selección de filtros."
End Sub

Private Sub tGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF1 Then
        frmFiltros.Show vbModal, Me
        Me.Refresh
        DoEvents
        With frmFiltros
            If Not .prmOK Then Exit Sub
            
            Dim mySQL As String, myPlan As String
            Dim myVigencia As String
            
            If Trim(.prmGrupo) <> "" Then           'Filtro de Grupos de Artículos      ---------------------------------------------
                If Mid(.prmGrupo, 1, 1) = "1" Then
                    mySQL = mySQL & " And ArtID IN "
                Else
                    mySQL = mySQL & " And ArtID Not IN "
                End If
                mySQL = mySQL & "(Select AGrArticulo from ArticuloGrupo Where AGrGrupo = " & Mid(.prmGrupo, 2) & ")"
            End If
            
            If Trim(.prmTipo) <> "" Then           'Filtro de Tipo de Artículos        ---------------------------------------------
                If Mid(.prmTipo, 1, 1) = "1" Then
                    'mySQL = mySQL & " And ArtTipo = " & Mid(.prmTipo, 2)
                    mySQL = mySQL & " And ArtTipo IN (Select TArId from dbo.TipoArticulo(" & Mid(.prmTipo, 2) & ") ) "
                Else
                    mySQL = mySQL & " And ArtTipo NOT IN (Select TArId from dbo.TipoArticulo(" & Mid(.prmTipo, 2) & ") ) "
                End If
            End If
        
            If Trim(.prmMarca) <> "" Then           'Filtro por Marca de Artículo        ---------------------------------------------
                If Mid(.prmMarca, 1, 1) = "1" Then
                    mySQL = mySQL & " And ArtMarca = " & Mid(.prmMarca, 2)
                Else
                    mySQL = mySQL & " And ArtMarca <> " & Mid(.prmMarca, 2)
                End If
            End If
        
            If Trim(.prmProveedor) <> "" Then          'Filtro por Proveedor de Artículo        ---------------------------------------------
                If Mid(.prmProveedor, 1, 1) = "1" Then
                    mySQL = mySQL & " And ArtProveedor = " & Mid(.prmProveedor, 2)
                Else
                    mySQL = mySQL & " And ArtProveedor <> " & Mid(.prmProveedor, 2)
                End If
            End If
            
            If Trim(.prmLista) <> "" Then          'Filtro por Listas de Artículo        ---------------------------------------------
                If Mid(.prmLista, 1, 1) = "1" Then
                    mySQL = mySQL & " And ArtID In (Select AFaArticulo From ArticuloFacturacion Where AFaLista = " & Mid(.prmLista, 2) & ")"
                Else
                    mySQL = mySQL & " And ArtID In (Select AFaArticulo From ArticuloFacturacion Where AFaLista <> " & Mid(.prmLista, 2) & ")"
                End If
            End If
            
            If Trim(.prmExclusivo) <> "" Then             'Filtro por Artículo Exclusivo        ---------------------------------------------
                mySQL = mySQL & " And ArtID In (Select AFaArticulo From ArticuloFacturacion Where AFaExclusivo = " & Trim(.prmExclusivo) & ")"
            End If
            
            If Trim(.prmHabilitado) <> "" Then             'Filtro por Artículo Exclusivo        ---------------------------------------------
                If .prmHabilitado = 1 Then
                    mySQL = mySQL & " And ArtHabilitado = 'S' "
                Else
                    mySQL = mySQL & " And ArtHabilitado <> 'S' "
                End If
            End If
            
            If Trim(.prmPrecio) <> "" Then
                If (Mid(.prmPrecio, 1, 1) = "<" Or Mid(.prmPrecio, 1, 1) = ">") And IsNumeric(Mid(.prmPrecio, 2)) Then
                    mySQL = mySQL & " And PViPrecio " & Mid(.prmPrecio, 1, 1) & " " & CCur(Mid(.prmPrecio, 2))
                End If
            End If
            
            myPlan = .prmPlan
            
                        
            If Trim(.prmVigencia) <> "" Then      'Filtro Vigencia del Precio
                    mySQL = mySQL & " And PViVigencia " & .prmVigencia
            End If

        End With
        
        If Trim(mySQL) <> "" Then AccionConsultar porFiltros:=mySQL, mPlan:=myPlan
    End If
End Sub

Private Sub tGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(tGrupo.Text) = "" Then Foco tGVigencia: Exit Sub
        If Val(tGrupo.Tag) <> 0 Then Foco tGVigencia: Exit Sub
        
        If IsNumeric(tGrupo.Text) Then
            AccionConsultar porArticulo:=CLng(tGrupo.Text)
            tGrupo.Text = "": tGrupo.SetFocus
            Exit Sub
        End If
        
        On Error GoTo errBuscaG
        
        Dim aQ As Integer, aIdGrupo As Long
        aQ = 0: aIdGrupo = 0
        
        Cons = "Select ArtCodigo as Codigo, ArtNombre  as Nombre from Articulo " & _
                   "Where ArtNombre like '" & Trim(tGrupo.Text) & "%'" & _
                   " Order by ArtNombre"
        
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            aQ = 1
            aIdGrupo = RsAux!Codigo: aTexto = Trim(RsAux!Nombre)
            RsAux.MoveNext: If Not RsAux.EOF Then aQ = 2
        End If
        RsAux.Close
        
        Select Case aQ
            Case 0: MsgBox "No hay datos que coincidan con el texto ingersado.", vbExclamation, "No hay datos"
            
            Case 2:
                        Dim miLista As New clsListadeAyuda
                        aIdGrupo = miLista.ActivarAyuda(cBase, Cons, 4000, 1, "Lista de Datos")
                        Me.Refresh
                        If aIdGrupo > 0 Then
                            aIdGrupo = miLista.RetornoDatoSeleccionado(0)
                            aTexto = miLista.RetornoDatoSeleccionado(1)
                        End If
                        Set miLista = Nothing
        End Select
        
        If aIdGrupo > 0 Then
            AccionConsultar porArticulo:=aIdGrupo
            tGrupo.Text = ""
        End If
        Screen.MousePointer = 0
    End If
   
    Exit Sub
errBuscaG:
    clsGeneral.OcurrioError "Error al buscar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tGrupo_LostFocus()
Status.Panels("help").Text = ""
End Sub

Private Sub tGVigencia_GotFocus()
    With tGVigencia: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub tGVigencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cGPlan
End Sub

Private Sub tGVigencia_LostFocus()

    If IsDate(tGVigencia.Text) Then
        tGVigencia.Text = Format(tGVigencia.Text, "dd/mm/yyyy hh:mm")
    Else
        tGVigencia.Text = ""
    End If
    
End Sub

Private Sub tTCTasa_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bCalcular.SetFocus
End Sub

Private Sub txtValorDolar_GotFocus()
    txtValorDolar.SelStart = 0: txtValorDolar.SelLength = Len(txtValorDolar.Text)
End Sub

Private Sub vsContado_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim miAumento As Currency, miValor As Currency

    '<Artículo|>Contado (V)|<Plan (V)|Fecha Vigencia|Contado (ME)|<Pan (N)|>Contado Nuevo|>% Aum.|
    Select Case Col
        
        Case 5      'Cambia Plan
                With vsContado
                    CalculoValoresCuotas Row, .Cell(flexcpData, Row, 5), .Cell(flexcpValue, Row, 6)
                End With
                
        Case 6      'Precio Contado
            With vsContado
                'Ajusto el % de aumento
                .Cell(flexcpForeColor, Row, 6) = ColorAjuste(.Cell(flexcpValue, Row, 6))
                
                If .Cell(flexcpValue, Row, 1) <> 0 Then
                    miAumento = ((.Cell(flexcpValue, Row, 6) * 100) / .Cell(flexcpValue, Row, 1)) - 100
                    If InStr(CStr(miAumento), ".") <> 0 Then miAumento = Format(miAumento, "0.00")
                    .Cell(flexcpText, Row, 7) = miAumento
                End If
                
                
                CalculoValoresCuotas Row, .Cell(flexcpData, Row, 5), .Cell(flexcpValue, Row, 6)
            End With
            
        Case 7      '% Aumento
            With vsContado
                'Ajusto el Valor contado -> Cambia Procentaje
                miAumento = 1 + (.Cell(flexcpValue, Row, Col) / 100)
             
                miValor = .Cell(flexcpValue, Row, 1) * miAumento
                .Cell(flexcpText, Row, 6) = FormatoImporte(miValor) 'Format(miValor, "#,##0")
                .Cell(flexcpForeColor, Row, 6) = ColorAjuste(.Cell(flexcpValue, Row, 6))
            
                CalculoValoresCuotas Row, .Cell(flexcpData, Row, 5), .Cell(flexcpValue, Row, 6)
            End With
        
        Case 9      'Cambio Margen Ganancia
            
            With vsContado
                .Cell(flexcpText, Row, 9) = Format(.Cell(flexcpValue, Row, 9), "0.00")
                
                'Precio Contado
                miValor = .Cell(flexcpValue, Row, 8) * (1 + (.Cell(flexcpValue, Row, 9) / 100))
                .Cell(flexcpText, Row, 6) = FormatoImporte(miValor)
                .Cell(flexcpForeColor, Row, 6) = ColorAjuste(.Cell(flexcpValue, Row, 6))
                
                'Ajusto el % de aumento
                If .Cell(flexcpValue, Row, 1) <> 0 Then
                    miAumento = ((.Cell(flexcpValue, Row, 6) * 100) / .Cell(flexcpValue, Row, 1)) - 100
                    If InStr(CStr(miAumento), ".") <> 0 Then miAumento = Format(miAumento, "0.00")
                    .Cell(flexcpText, Row, 7) = miAumento
                End If
                                
                CalculoValoresCuotas Row, .Cell(flexcpData, Row, 5), .Cell(flexcpValue, Row, 6)
            End With
            
    End Select

End Sub

Private Sub vsContado_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    '<Artículo|>Contado (V)|<Plan (V)|Fecha Vigencia|<Pan (N)|>Contado Nuevo|>% Aum.|
    If Col <> 5 And Col <> 6 And Col <> 7 And Col <> 9 Then Cancel = True
    If Col = 7 Then If Trim(vsContado.Cell(flexcpText, Row, Col)) = "" Then Cancel = True
    
End Sub

Private Sub vsContado_GotFocus()
    Status.Panels("help").Text = "[Del]- Elimina artículo de la lista."
End Sub

Private Sub vsContado_KeyDown(KeyCode As Integer, Shift As Integer)

    If vsContado.Rows = 1 Then Exit Sub
    On Error Resume Next
    If KeyCode = vbKeyDelete And vsContado.Col = 0 Then
        vsCuotas.RemoveItem vsContado.Row
        vsContado.RemoveItem vsContado.Row
        
        If vsContado.Rows = 1 Then cMoneda.Enabled = True
    End If
    
End Sub

Private Sub vsContado_RowColChange()
    On Error Resume Next
    Static inhere%
    inhere = False
    If inhere Then Exit Sub
    inhere = True
    
    vsCuotas.Select vsContado.Row, vsCuotas.Col
    
    inhere = False
    
End Sub

Private Sub vsContado_Scroll()
    On Error Resume Next
    Static inhere%
    inhere = False
    If inhere Then Exit Sub
    inhere = True
    
    vsCuotas.TopRow = vsContado.TopRow
    
    inhere = False
    
End Sub

Private Sub vsContado_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '<Artículo|>Contado (V)|<Plan (V)|Fecha Vigencia|>Contado (ME)|<Pan (N)|>Contado Nuevo|>% Aum.|
    Select Case Col
        Case 5
            With vsContado
                If Trim(.EditText) = "" Then Cancel = True: Exit Sub
                Dim miPlan As Long: miPlan = 0
                
                Cons = "Select * from TipoPlan Where PlaNombre = '" & Trim(.EditText) & "'"
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                If Not RsAux.EOF Then miPlan = RsAux!PlaCodigo
                RsAux.Close
                
                If miPlan = 0 Then Cancel = True: Exit Sub
                .EditText = UCase(.EditText)
                .Cell(flexcpData, Row, Col) = miPlan
                
            End With
        
        Case 6          'Precio Contado
                With vsContado
                    If Not IsNumeric(.EditText) Then Cancel = True: Exit Sub
                    .EditText = FormatoImporte(.EditText) 'Format(.EditText, "#,##0")
                End With
                
        Case 7, 9         '% Aumento
            With vsContado
                If Not IsNumeric(.EditText) Then Cancel = True: Exit Sub
            End With
            
    End Select
    
End Sub

Private Sub vsCuotas_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '"<Artículo|>Contado (N)"
    If Col < 2 Then Cancel = True
    If vsCuotas.Cell(flexcpBackColor, Row, Col) = vsCuotas.BackColorFixed Then Cancel = True: Exit Sub

End Sub

Private Sub vsCuotas_DblClick()
On Error GoTo err2C
Dim pblnCancel As Boolean

    If vsCuotas.Col < 2 Then pblnCancel = True
    If vsCuotas.Cell(flexcpBackColor, vsCuotas.Row, vsCuotas.Col) = vsCuotas.BackColorFixed Then pblnCancel = True

    If Not pblnCancel Then
        Dim pcurValor As Currency, idx As Integer, pbytValor As Byte
        pcurValor = vsCuotas.Cell(flexcpValue, vsCuotas.Row, vsCuotas.Col)
        
        pbytValor = Right(pcurValor, 1)
        pcurValor = pcurValor + Val(arrRND(pbytValor))
        
        vsCuotas.Cell(flexcpText, vsCuotas.Row, vsCuotas.Col) = Format(pcurValor, IIf(InStr(mFormato, ".") <> 0, "#,##0.00", "#,##0"))
    End If
'1 -1
'2 -2
'3 +2
'4 +1
'6 -1
'7 +1
'8 +2
'9 +1
'0 -1
'5 +0

err2C:
End Sub

Private Sub vsCuotas_GotFocus()
    
    Status.Panels("help").Text = "[Del]- Deshabilita/habilita Cuota p/Artículo           [Shift+Del]- Deshabilita/habilita Todas"
    
End Sub

Private Sub vsCuotas_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsCuotas.Rows = 1 Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
        If vsCuotas.Col < 2 Then Exit Sub
        
        With vsCuotas
            If Shift = vbShiftMask Then
                For I = 1 To .Rows - 1
                    If .Cell(flexcpBackColor, I, .Col) = .BackColor Then
                        .Cell(flexcpBackColor, I, .Col) = Colores.clCeleste
                    Else
                        .Cell(flexcpBackColor, I, .Col) = .BackColor
                    End If
                Next
            
            Else
                If .Cell(flexcpBackColor, .Row, .Col) = .BackColor Then
                    .Cell(flexcpBackColor, .Row, .Col) = Colores.clCeleste
                Else
                    .Cell(flexcpBackColor, .Row, .Col) = .BackColor
                End If
            End If
        End With
    End If
    
End Sub

Private Sub vsCuotas_RowColChange()
On Error Resume Next
    Static inhere%
    inhere = False
    If inhere Then Exit Sub
    inhere = True
    
    vsContado.Select vsCuotas.Row, vsContado.Col
    
    inhere = False
End Sub

Private Sub vsCuotas_Scroll()
    On Error Resume Next
    Static inhere%
    inhere = False
    If inhere Then Exit Sub
    inhere = True
    
    vsContado.TopRow = vsCuotas.TopRow
    
    inhere = False
End Sub

Private Sub vsCuotas_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '"<Artículo|>Contado (N)"
    With vsCuotas
        If Not IsNumeric(.EditText) Then Cancel = True: Exit Sub
        .EditText = Format(.EditText, "#,##0")
                
        If Trim(.EditText) <> Trim(.Cell(flexcpText, Row, Col)) Then
            .Cell(flexcpForeColor, Row, Col) = vbBlue 'ColorAjuste(CCur(.EditText))
        End If
        
        If Abs(((CCur(.EditText) * 100) / .Cell(flexcpValue, Row, Col)) - 100) > 1 Then
            If MsgBox("El cambio de precio supera el 1 % con respecto al valor original" & vbCrLf & _
                            "Valor original " & .Cell(flexcpText, Row, Col) & "      Ud. ingresó " & .EditText & vbCrLf & vbCrLf & _
                            "Continúa con el ingreso.", vbYesNo + vbDefaultButton2 + vbExclamation, "Variación Mayor al 1%") = vbNo Then
                    Cancel = True
                    Exit Sub
            End If
        End If
        
        '(Cta Vieja - CtaNueva) * qCtas
        Dim aRebaja As Currency, aPorc As Currency
        
        aRebaja = Format((.Cell(flexcpValue, Row, Col) - CCur(.EditText)) * .Cell(flexcpData, Row, Col), "#,##0")
        aPorc = aRebaja / (.Cell(flexcpValue, Row, Col) * .Cell(flexcpData, Row, Col))
        aPorc = Format(aPorc * 100, "0.00")
        Status.Panels("rebaja").Text = "Rebaja   $ " & aRebaja & "  (" & aPorc & " %)"
    End With
    
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    With vsListado
        .StartDoc
        .Columns = 1
        If .Error Then
            MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & Err.Number & "- " & Err.Description, vbInformation, "ATENCIÓN": Screen.MousePointer = vbDefault
            Screen.MousePointer = 0: Exit Sub
        End If
    End With        '----------------------------------------------------------------------------------------------------------------------------------------------
    
    aTexto = "Actualización de Precios"
    EncabezadoListado vsListado, aTexto, False
    vsListado.FileName = "Precios"
    
    With vsContado
        .Redraw = False
        .ExtendLastCol = False: vsListado.RenderControl = .hwnd: .ExtendLastCol = True
        .Redraw = True
    End With
    vsListado.Paragraph = ""
    
    With vsCuotas
        .Redraw = False
        '.ExtendLastCol = False:
        vsListado.RenderControl = .hwnd
        ': .ExtendLastCol = True
        .Redraw = True
    End With
    
    vsListado.EndDoc
    
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
    clsGeneral.OcurrioError "Error al realizar la impresión", Err.Description
End Sub

Private Sub AccionConfigurar()
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
End Sub

Public Function ColorAjuste(cImporte As Currency) As Long
    On Error GoTo errCA
    
    ColorAjuste = vsContado.ForeColor
    'If (Right(Importe, 2) / Importe) < 0.008 Then ColorAjuste = vbRed
    
    Dim Importe As String
    Importe = CStr(cImporte)
    If InStr(Importe, ".") <> 0 Then Importe = Mid(Importe, 1, InStr(Importe, ".") - 1)
    
    If Len(Importe) < 3 Then Exit Function
    If (Right(Importe, 2) / Right(Importe, 3)) < 0.01 Then ColorAjuste = vbRed: Exit Function
    
    If Len(Importe) < 4 Then Exit Function
    If (Right(Importe, 3) / Right(Importe, 4)) < 0.01 Then ColorAjuste = vbRed: Exit Function
    
    If Len(Importe) < 5 Then Exit Function
    If (Right(Importe, 4) / Right(Importe, 5)) < 0.01 Then ColorAjuste = vbRed: Exit Function

errCA:
    
End Function

Private Sub InicializoArrayCoef(idMoneda As Long)
    On Error GoTo errIniArray
    ReDim arrCoef(0)
    Dim aIdx As Integer: aIdx = 0
    
    Cons = "Select * from Coeficiente, TipoCuota" _
            & " Where CoeTipoCuota = TCuCodigo" _
            & " And CoeMoneda = " & idMoneda _
            & " And TCuVencimientoE Is Null" _
            & " And CoeTipoCuota <> " & paTipoCuotaContado
            
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        ReDim Preserve arrCoef(aIdx)
        With arrCoef(aIdx)
            .Coeficiente = RsAux!CoeCoeficiente
            .Plan = RsAux!CoePlan
            .TCuota = RsAux!CoeTipoCuota
            .QCtas = RsAux!TCuCantidad
        End With
        
        aIdx = aIdx + 1
        RsAux.MoveNext
    Loop
    RsAux.Close
    Exit Sub

errIniArray:
    clsGeneral.OcurrioError "Error al inicializar array de coeficientes.", Err.Description
End Sub

Private Function arrCoef_Coeficiente(xPlan As Long, xTCuota As Long, xCoef As Currency, xQCtas As Integer) As Boolean
Dim X As Long
    
    arrCoef_Coeficiente = False
    xCoef = 1: xQCtas = 1
    For X = LBound(arrCoef) To UBound(arrCoef)
        With arrCoef(X)
            If .Plan = xPlan And .TCuota = xTCuota Then
                xCoef = .Coeficiente
                xQCtas = .QCtas
                arrCoef_Coeficiente = True
                Exit For
            End If
        End With
    Next
End Function

Private Function DCount(ByVal miSql As String) As Long
    On Error GoTo errDCount
    DCount = 0
    
    miSql = clsGeneral.Replace(miSql, "*", "Count(*)")
    Set RsAux = cBase.OpenResultset(miSql, rdOpenDynamic, rdConcurValues)
    If Not RsAux.EOF Then If Not IsNull(RsAux(0)) Then DCount = RsAux(0)
    RsAux.Close
    Exit Function
    
errDCount:
    clsGeneral.OcurrioError "Error al consultar la cantidad de registros.", Err.Description
End Function

Private Sub AccionGrabar()
    
    On Error GoTo errGrabar
    Dim aIdArticulo As Long
    Dim aIdPlan As Long, aIdTCuota As Long
    Dim aVigencia As Date
    Dim bIguales As Boolean, bHayIguales As Boolean
    Dim aCol As Integer
    
    Dim mMoneda As Long
    mMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    
    Dim rsUpd As rdoResultset
    Screen.MousePointer = 11
    bHayIguales = False
    FechaDelServidor
    
    If IsDate(tGVigencia.Text) Then
        If CDate(tGVigencia.Text) < gFechaServidor Then
            aVigencia = gFechaServidor
        Else
            aVigencia = CDate(tGVigencia.Text)
        End If
    End If
    
    'Valido Precios Iguales     ------------------------------------------------------------------------------------
    bIguales = False
    For I = 1 To vsContado.Rows - 1
        If vsContado.Cell(flexcpValue, I, 1) = vsContado.Cell(flexcpValue, I, 6) And _
            vsContado.Cell(flexcpData, I, 2) = vsContado.Cell(flexcpData, I, 5) Then
                bIguales = True
                Exit For
        End If
    Next
    '------------------------------------------------------------------------------------
    Dim bActualizarIguales As Boolean
    bActualizarIguales = False
    
    If bIguales Then
        If MsgBox("Hay artículos que no cambiaron de precios." & vbCrLf & _
                    "Ud. quiere actualizarlos ?", vbQuestion + vbYesNo + vbDefaultButton2, "Artículos con Mismo Precio") = vbYes Then
                bActualizarIguales = True
        End If
    End If
    Cons = "Select * from HistoriaPrecio Where HPrArticulo = 0"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    
    pbProgreso.Value = 0: pbProgreso.Max = vsContado.Rows - 1
    
    For I = 1 To vsContado.Rows - 1
        bIguales = False
        
        '<Artículo|>Contado (V)|<Plan (V)|Fecha Vigencia|Contado ME|<Pan (N)|>Contado Nuevo|>% Aum.|"
        If vsContado.Cell(flexcpValue, I, 1) = vsContado.Cell(flexcpValue, I, 6) And _
            vsContado.Cell(flexcpData, I, 2) = vsContado.Cell(flexcpData, I, 5) Then bIguales = True
        
        If Not bActualizarIguales Then
            If bIguales Then bHayIguales = True
        Else
            bIguales = False
        End If
        
        If Not bIguales Then
            With vsContado
                aIdArticulo = .Cell(flexcpData, I, 0)
                aIdPlan = .Cell(flexcpData, I, 5)
            End With
            
            With vsCuotas
                For aCol = 1 To .Cols - 1
                    If Trim(.Cell(flexcpText, I, aCol)) <> "" Then
                        Select Case .Cell(flexcpBackColor, I, aCol)
                            Case .BackColor
                                    RsAux.AddNew
                                    RsAux!HPrArticulo = aIdArticulo
                                    RsAux!HPrTipoCuota = .Cell(flexcpData, 0, aCol)
                                    RsAux!HPrMoneda = mMoneda
                                    RsAux!HPrVigencia = Format(aVigencia, "mm/dd/yyyy hh:mm:ss")
                                    
                                    RsAux!HPrPlan = aIdPlan
                                    RsAux!HPrPrecio = .Cell(flexcpValue, I, aCol) * .Cell(flexcpData, I, aCol)
                                    RsAux!HPrHabilitado = True
                                    RsAux.Update
                            
                            Case Colores.clCeleste          'Deshabilito el Vigente
                                    Cons = "Update PrecioVigente " & _
                                               " Set PViHabilitado = 0 " & _
                                               " Where PViArticulo = " & aIdArticulo & _
                                               " And PViTipoCuota = " & .Cell(flexcpData, 0, aCol) & _
                                               " And PViMoneda = " & mMoneda
                                    cBase.Execute Cons
                        End Select
                    End If
                Next
            End With
            
        End If
        pbProgreso.Value = pbProgreso.Value + 1
    Next
    RsAux.Close

    pbProgreso.Value = 0
    Screen.MousePointer = 0
    
    If bHayIguales Then
        MsgBox "Hubieron artículos que no se tomaron en cuenta para la actualización." & vbCrLf & _
                    "Causa: Mismo plan y precio contado.", vbExclamation, "Precios Actualizados OK"
    End If
    bGrabar.Enabled = False
    Exit Sub

errGrabar:
    clsGeneral.OcurrioError "Error al grabar los precios.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function FormatoImporte(mImporte As Currency) As String
    On Error Resume Next
    Dim aRet As Currency, mFmt As String
    
    aRet = Redondeo(mImporte, mFormato)
    mFmt = "#,##0"
    
    If InStr(mFormato, ".") <> 0 Then mFmt = "#,##0.00"
    
    FormatoImporte = Format(aRet, mFmt)
    
End Function

Private Function ValidoDatos() As Boolean
On Error GoTo errValidar
    
    ValidoDatos = False
    With vsCuotas
        For I = 1 To .Rows - 1
            If .Cell(flexcpValue, I, 1) = 0 Then
                MsgBox "Hay artículos con precio contado igual a cero." & vbCrLf & _
                            "Si no se van a actualizar, elimine los artículos de la lista.", vbExclamation, "Artículos sin Precios Ingresados"
                Exit Function
            End If
            
            If vsContado.Cell(flexcpData, I, 5) = 0 Then
                MsgBox "Hay artículos que no tienen ingresado el plan de financiación." & vbCrLf & _
                            "Si no se van a actualizar, elimine los artículos de la lista.", vbExclamation, "Artículos sin Plan de Financiación"
                Exit Function
            End If

        Next
    
    End With
    
    ValidoDatos = True
    Exit Function

errValidar:
End Function

Private Function CargoPreciosME(idMoneda As Long, Todos As Boolean)

    'Saco todos los códigos de artículos
    If vsContado.Rows = 1 Then Exit Function
    
    Dim strIds As String
    For I = 1 To vsContado.Rows - 1
        If Todos Or (Not Todos And Trim(vsContado.Cell(flexcpText, I, 6)) = "") Then
            strIds = strIds & vsContado.Cell(flexcpData, I, 0) & ","
            vsContado.Cell(flexcpText, I, 4) = ""
        End If
    Next
    If Trim(strIds) = "" Then Exit Function
    
    Dim mMERound  As String
    mMERound = 1
    Cons = "Select * from Moneda Where MonCodigo = " & idMoneda
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        vsContado.Cell(flexcpText, 0, 4) = "Contado " & Trim(RsAux!MonSigno)
        If Not IsNull(RsAux!MonRedondeo) Then mMERound = Trim(RsAux!MonRedondeo)
    End If
    RsAux.Close
    
    strIds = Mid(strIds, 1, Len(strIds) - 1)
    
    Dim mRow As Long
    
    Cons = "Select * From PrecioVigente " & _
               " Where PViArticulo In (" & strIds & ")" & _
               " And PViTipoCuota = " & paTipoCuotaContado & _
               " And PViMoneda = " & idMoneda
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    Do While Not RsAux.EOF
        If EnLista(RsAux!PViArticulo, mRow) Then
            vsContado.Cell(flexcpText, mRow, 4) = Format(Redondeo(RsAux!PViPrecio, mMERound), "#,##0.00")
        End If
        RsAux.MoveNext
    Loop
    RsAux.Close
    
End Function

