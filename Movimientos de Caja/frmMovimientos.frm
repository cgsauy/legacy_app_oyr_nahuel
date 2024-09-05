VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMovimientos 
   Caption         =   "Movimientos de Caja"
   ClientHeight    =   6420
   ClientLeft      =   1860
   ClientTop       =   2670
   ClientWidth     =   9330
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
   ScaleWidth      =   9330
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   60
      ScaleHeight     =   435
      ScaleWidth      =   7695
      TabIndex        =   12
      Top             =   5640
      Width           =   7755
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Picture         =   "frmMovimientos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   1740
         Picture         =   "frmMovimientos.frx":0744
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   1080
         Picture         =   "frmMovimientos.frx":0846
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   720
         Picture         =   "frmMovimientos.frx":0C0C
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin MSComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   2400
         TabIndex        =   17
         Top             =   120
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   2835
      Left            =   4320
      TabIndex        =   8
      Top             =   1560
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5001
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
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   6165
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "sucursal"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "terminal"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "usuario"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "bd"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5689
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
      TabIndex        =   9
      Top             =   0
      Width           =   9735
      Begin AACombo99.AACombo cDisponibilidad 
         Height          =   315
         Left            =   1200
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
      Begin AACombo99.AACombo cMovimiento 
         Height          =   315
         Left            =   5640
         TabIndex        =   3
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
      End
      Begin MSComCtl2.DTPicker dDesde 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   23592961
         CurrentDate     =   37543
      End
      Begin MSComCtl2.DTPicker dHasta 
         Height          =   315
         Left            =   3120
         TabIndex        =   7
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Format          =   23592961
         CurrentDate     =   37543
      End
      Begin VB.Label Label4 
         Caption         =   "Al:"
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   675
         Width           =   435
      End
      Begin VB.Label Label3 
         Caption         =   "Desde el:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   675
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "&Disponibilidad:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   315
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Movimiento:"
         Height          =   255
         Left            =   4680
         TabIndex        =   2
         Top             =   315
         Width           =   1095
      End
   End
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   3375
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   9615
      _Version        =   196608
      _ExtentX        =   16960
      _ExtentY        =   5953
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
      Zoom            =   70
      EmptyColor      =   0
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   10260
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimientos.frx":0D0E
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimientos.frx":1028
            Key             =   "help"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimientos.frx":113A
            Key             =   "consultar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimientos.frx":1294
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimientos.frx":13EE
            Key             =   "print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimientos.frx":1548
            Key             =   "limpiar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimientos.frx":165A
            Key             =   "vista2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimientos.frx":17B4
            Key             =   "vista1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimientos.frx":190E
            Key             =   "move2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimientos.frx":1A68
            Key             =   "move3"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimientos.frx":1BC2
            Key             =   "move4"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimientos.frx":1D1C
            Key             =   "move1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMovimientos.frx":1E76
            Key             =   "configprint"
         EndProperty
      EndProperty
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
    dDesde.Value = Now: dHasta.Value = Now
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
    If KeyCode = vbKeyReturn Then dDesde.SetFocus
End Sub

Private Sub dDesde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dHasta.SetFocus
End Sub

Private Sub dHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()

    On Error GoTo ErrLoad

    ObtengoSeteoForm Me, 100, 100, 9500, 7000
    InicializoControles
    InicializoGrillas
    AccionLimpiar
    
    cons = "Select DisID, DisNombre from Disponibilidad Order by DisNombre"
    CargoCombo cons, cDisponibilidad
    
    
    cons = "Select TMDCodigo, TMDNombre from TipoMovDisponibilidad Order by TMDNombre"
    CargoCombo cons, cMovimiento
    
    If Trim(Command()) <> "" Then       'Viene Disponibilidad:Movimiento:dd/mm/yyyy
        Dim arrPrm() As String
        arrPrm = Split(Trim(Command()), ":")
        For I = LBound(arrPrm) To UBound(arrPrm)
            Select Case I
                Case 0: BuscoCodigoEnCombo cDisponibilidad, Val(arrPrm(I))
                Case 1: BuscoCodigoEnCombo cMovimiento, Val(arrPrm(I))
                Case 2
                        If IsDate(arrPrm(I)) Then dDesde.Value = Format(arrPrm(I), "dd/mm/yyyy")
                        dHasta.Value = dDesde.Value
                
                Case 3
                        If IsDate(arrPrm(I)) Then dHasta.Value = Format(arrPrm(I), "dd/mm/yyyy")
            End Select
        Next
        
        AccionConsultar
    Else
        BuscoCodigoEnCombo cDisponibilidad, paDisponibilidad
        dDesde.Value = Format(Now, "dd/mm/yyyy")
        dHasta.Value = dDesde.Value
    End If
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Error al inicializar el formulario.", Trim(Err.Description)
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = "ID|Tipo de Movimiento|Fecha|Hora|Comentario|>Debe|>Haber|>ID Gasto|>TC|>Debe $U|>Haber $U|"
            
        .WordWrap = False
        .ColWidth(0) = 700: .ColWidth(1) = 1800: .ColWidth(2) = 800: .ColWidth(3) = 600: .ColWidth(4) = 2500
        .ColWidth(5) = 1100: .ColWidth(6) = 1100
        .ColWidth(8) = 600: .ColWidth(9) = 1100: .ColWidth(10) = 1100
        '.ColDataType(2) = flexDTCurrency

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
    fFiltros.Left = 60
    fFiltros.Width = Me.ScaleWidth - (fFiltros.Left * 2)
    
    With picBotones
        .Left = Me.ScaleLeft
        .Width = Me.ScaleWidth
        .Top = Me.ScaleHeight - (.Height + Status.Height)
    End With
    
    With pbProgreso
        .Width = picBotones.Width - .Left - 100
    End With
    
    With vsConsulta
        .Left = fFiltros.Left
        .Top = fFiltros.Top + fFiltros.Height + 50
        .Height = picBotones.Top - .Top - 60
        .Width = fFiltros.Width
    End With
    
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
Dim mQ As Long
Dim mPesos As Currency, mDH As Currency

Dim rsQ As rdoResultset

    On Error GoTo errConsultar
    If Not ValidoFiltros Then Exit Sub
    
    Screen.MousePointer = 11

    Dim mIDMoneda As Long
    mIDMoneda = MonedaDisponibilidad(cDisponibilidad.ItemData(cDisponibilidad.ListIndex))
    
    vsConsulta.ColHidden(9) = (mIDMoneda = paMonedaPesos)
    vsConsulta.ColHidden(10) = vsConsulta.ColHidden(9)
    
    vsConsulta.Rows = 1
    
    'Inicializo progress bar --------------------------------------------------------------------------------------
    cons = "Select Count(*) from MovimientoDisponibilidad, MovimientoDisponibilidadRenglon" & _
            " Where MDiID = MDRIdMovimiento" & _
            " And MDRIdDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) & _
            " And MDiFecha Between " & Format(dDesde.Value, "'mm/dd/yyyy'") & _
                                        " And " & Format(dHasta.Value, "'mm/dd/yyyy'")

    If cMovimiento.ListIndex <> -1 Then cons = cons & " And MDiTipo = " & cMovimiento.ItemData(cMovimiento.ListIndex)
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then
        If Not IsNull(rsAux(0)) Then mQ = rsAux(0)
    End If
    rsAux.Close
    '-----------------------------------------------------------------------------------------------------------------
    If mQ = 0 Then
        Screen.MousePointer = 0
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "No Hay Datos"
        Exit Sub
    End If
    
    pbProgreso.Max = mQ
    pbProgreso.Value = 0
    
    cons = "Select * from MovimientoDisponibilidad, MovimientoDisponibilidadRenglon, TipoMovDisponibilidad" & _
            " Where MDiID = MDRIdMovimiento" & _
            " And MDiTipo = TMDCodigo" & _
            " And MDRIdDisponibilidad = " & cDisponibilidad.ItemData(cDisponibilidad.ListIndex) & _
            " And MDiFecha Between " & Format(dDesde.Value, "'mm/dd/yyyy'") & _
                                        " And " & Format(dHasta.Value, "'mm/dd/yyyy'")
           
    If cMovimiento.ListIndex <> -1 Then cons = cons & " And MDiTipo = " & cMovimiento.ItemData(cMovimiento.ListIndex)
    
    cons = cons & " Order by MDiFecha, MDiHora"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        rsAux.Close: Screen.MousePointer = 0: Exit Sub
    End If

    With vsConsulta
        Do While Not rsAux.EOF
            
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(rsAux!MDiId, "#,##0")
            .Cell(flexcpText, .Rows - 1, 1) = Trim(rsAux!TMDNombre)
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsAux!MDiFecha, "dd/mm/yy")
            .Cell(flexcpText, .Rows - 1, 3) = Format(rsAux!MDiHora, "hh:mm")
            If Not IsNull(rsAux!MDiComentario) Then .Cell(flexcpText, .Rows - 1, 4) = Trim(rsAux!MDiComentario)
            
            If Not IsNull(rsAux!MDRDebe) Then
                .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!MDRDebe, FormatoMonedaP)
                mDH = .Cell(flexcpValue, .Rows - 1, 5)
            End If
            If Not IsNull(rsAux!MDRHaber) Then
                .Cell(flexcpText, .Rows - 1, 6) = Format(rsAux!MDRHaber, FormatoMonedaP)
                mDH = .Cell(flexcpValue, .Rows - 1, 6)
            End If
            
            If Not IsNull(rsAux!MDiIdCompra) Then .Cell(flexcpText, .Rows - 1, 7) = Format(rsAux!MDiIdCompra, "#,##0")
            
            mPesos = Format(rsAux!MDRImportePesos, "#,##0.00")
            If mDH <> mPesos Then
                .Cell(flexcpText, .Rows - 1, 8) = Format(mPesos / mDH, "#,##0.000")
                If Not IsNull(rsAux!MDRDebe) Then .Cell(flexcpText, .Rows - 1, 9) = Format(mPesos, "#,##0.00")
                If Not IsNull(rsAux!MDRHaber) Then .Cell(flexcpText, .Rows - 1, 10) = Format(mPesos, "#,##0.00")
            
            Else
                    'IRMA: 8/04/2003
                    'Si el Tipo de Mov es Compra M/E    busco el otro renglón para poner tasas de cambio    ----------
                    If rsAux!MDiTipo = prmTipoMCompraME Then
                        
                        cons = "Select * from MovimientoDisponibilidadRenglon" & _
                                " Where MDRIDMovimiento = " & rsAux!MDiId & _
                                " And MDRIDDisponibilidad <> " & rsAux!MDRIdDisponibilidad
                        
                        Set rsQ = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                        If Not rsQ.EOF Then
                            .Cell(flexcpText, .Rows - 1, 8) = Format(mPesos / rsQ!MDRImporteCompra, "#,##0.000")
                            .Cell(flexcpText, .Rows - 1, 11) = Format(rsQ!MDRImporteCompra, "#,##0.00")
                            .Cell(flexcpAlignment, .Rows - 1, 11) = flexAlignLeftCenter
                            .ColWidth(11) = 1000
                            'If Not IsNull(rsAux!MDRDebe) Then .Cell(flexcpText, .Rows - 1, 9) = Format(mPesos, "#,##0.00")
                            'If Not IsNull(rsAux!MDRHaber) Then .Cell(flexcpText, .Rows - 1, 10) = Format(mPesos, "#,##0.00")
                        End If
                        rsQ.Close
                        
                    End If
                    '----------------------------------------------------------------------------------------------------------------
                    
            End If
            
            rsAux.MoveNext
            
            pbProgreso.Value = pbProgreso.Value + 1
        Loop
        rsAux.Close
        
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 5, , , , , "Subtotal"
        .Subtotal flexSTSum, -1, 6
        .Subtotal flexSTSum, -1, 9
        .Subtotal flexSTSum, -1, 10
        
        Dim aTotal As Currency, aTotalPU As Currency
        aTotal = .Cell(flexcpValue, .Rows - 1, 5) - .Cell(flexcpValue, .Rows - 1, 6)
        aTotalPU = .Cell(flexcpValue, .Rows - 1, 9) - .Cell(flexcpValue, .Rows - 1, 10)
        
        'If aTotal <> 0 Then
            .AddItem "Total"
            If aTotal < 0 Then .Cell(flexcpText, .Rows - 1, 6) = Format(aTotal, "#,##0.00") Else .Cell(flexcpText, .Rows - 1, 5) = Format(aTotal, "#,##0.00")
            If aTotalPU < 0 Then .Cell(flexcpText, .Rows - 1, 10) = Format(aTotalPU, "#,##0.00") Else .Cell(flexcpText, .Rows - 1, 9) = Format(aTotalPU, "#,##0.00")
        'End If
        
        .Cell(flexcpBackColor, 1, 5, .Rows - 1, 6) = Colores.Obligatorio
        .Cell(flexcpForeColor, 1, 6, .Rows - 1) = Colores.Rojo
        .Cell(flexcpBackColor, 1, 9, .Rows - 1, 10) = Colores.Obligatorio
        .Cell(flexcpForeColor, 1, 10, .Rows - 1) = Colores.Rojo
        .Cell(flexcpBackColor, .Rows - 2, 0, .Rows - 1, .Cols - 1) = Colores.Gris
            
        Screen.MousePointer = 0
        
    End With
    pbProgreso.Value = 0
    Exit Sub

errConsultar:
    clsGeneral.OcurrioError "Error al realizar la consulta de datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ValidoFiltros() As Boolean

    ValidoFiltros = False
    If cDisponibilidad.ListIndex = -1 Then
        MsgBox "Debe seleccionar la disponibilidad para realizar la consulta.", vbExclamation, "Faltan Datos"
        Foco cDisponibilidad: Exit Function
    End If
    
    If Not IsDate(dDesde.Value) Or Not IsDate(dHasta.Value) Then
        MsgBox "Se deben ingresar las fechas para realizar la consulta.", vbExclamation, "Error en Fechas"
        dDesde.SetFocus: Exit Function
    End If
    
    If dDesde.Value > dHasta.Value Then
        MsgBox "Error en el rango de fechas inrgesado.", vbExclamation, "Error en Fechas"
        dDesde.SetFocus: Exit Function
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
        
        .Paragraph = "Disponibilidad: " & Trim(cDisponibilidad.Text) & Chr(vbKeyTab) & _
                            "Del  " & dDesde.Value & " al " & dHasta.Value
        
        vsConsulta.ExtendLastCol = False: .RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        
        .EndDoc
        .PrintDoc
        '.Visible = True: .ZOrder 0
        
    End With
    Screen.MousePointer = 0
    Exit Sub

errPrint:
    clsGeneral.OcurrioError "Error al realizar la impresión. ", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Label1_Click()
    dDesde.SetFocus
End Sub

Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

Private Sub InicializoControles()

    With img1
        bConsultar.Picture = .ListImages("consultar").ExtractIcon
        bImprimir.Picture = .ListImages("print").ExtractIcon
        bNoFiltros.Picture = .ListImages("limpiar").ExtractIcon
        bCancelar.Picture = .ListImages("salir").ExtractIcon
    End With
    
End Sub

Private Function MonedaDisponibilidad(idDisp As Long) As Long

    cons = "Select * from Disponibilidad Where DisID = " & idDisp
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    MonedaDisponibilidad = rsAux!DisMoneda
    rsAux.Close
    
End Function
