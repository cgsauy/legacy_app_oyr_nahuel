VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "AACombo.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VSVIEW3.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmListado 
   Caption         =   "Asientos de Gastos"
   ClientHeight    =   7995
   ClientLeft      =   1260
   ClientTop       =   1785
   ClientWidth     =   11760
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
   ScaleHeight     =   7995
   ScaleWidth      =   11760
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   1140
      TabIndex        =   19
      Top             =   1380
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
      TabIndex        =   22
      Top             =   720
      Width           =   11175
      _Version        =   196608
      _ExtentX        =   19711
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
      AbortWindowPos  =   0
      AbortWindowPos  =   0
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   9555
      TabIndex        =   23
      Top             =   6960
      Width           =   9615
      Begin VB.CommandButton bExportar 
         Height          =   310
         Left            =   5280
         Picture         =   "frmListado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
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
         TabIndex        =   15
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2880
         Picture         =   "frmListado.frx":0D80
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2520
         Picture         =   "frmListado.frx":0E6A
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin MSComctlLib.ProgressBar pbProgreso 
         Height          =   270
         Left            =   6000
         TabIndex        =   24
         Top             =   120
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   7740
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "terminal"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "bd"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12541
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
      TabIndex        =   20
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton bExpandir 
         Caption         =   "Expandir"
         Height          =   315
         Left            =   8580
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chMas 
         Caption         =   "&Mas columnas"
         Height          =   195
         Left            =   7020
         TabIndex        =   25
         Top             =   300
         Width           =   1575
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
         Left            =   4920
         TabIndex        =   5
         Top             =   240
         Width           =   1875
         _ExtentX        =   3307
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
         Left            =   4440
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
   Begin MSComctlLib.ImageList img1 
      Left            =   11160
      Top             =   120
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
            Picture         =   "frmListado.frx":0F54
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":126E
            Key             =   "help"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1380
            Key             =   "consultar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":14DA
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1634
            Key             =   "print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":178E
            Key             =   "limpiar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":18A0
            Key             =   "vista2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":19FA
            Key             =   "vista1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1B54
            Key             =   "move2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1CAE
            Key             =   "move3"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1E08
            Key             =   "move4"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":1F62
            Key             =   "move1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListado.frx":20BC
            Key             =   "configprint"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog ctlDlg 
      Left            =   11160
      Top             =   780
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
Option Explicit

Private rsAux As rdoResultset, rs1 As rdoResultset
Private aTexto As String
Dim bCargarImpresion As Boolean

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
        
            On Error GoTo errSaving
            
            With vsConsulta
                Dim iR As Integer, iX As Integer
                iX = .FixedRows
                
                For iR = .FixedRows To .Rows - 1
                    If Not (.IsSubtotal(iX) And (.RowOutlineLevel(iX) = 0)) Or .Cell(flexcpBackColor, iX, 5) = Colores.Obligatorio Then
                        'If (Trim(.Cell(flexcpText, iX, 2)) <> "" And Trim(.Cell(flexcpText, iX, 1)) = "") Or .Cell(flexcpBackColor, iX, 5) = Colores.Obligatorio Then
                        If Trim(.Cell(flexcpText, iX, 2)) <> "" Then
                            .RemoveItem iX: iX = iX - 1
                        Else
                            If .Cell(flexcpBackColor, iX, 5) = Colores.Obligatorio Then .RemoveItem iX: iX = iX - 1
                        End If
                    Else
                        If .Cell(flexcpBackColor, iX, 5) = Colores.Obligatorio Then .RemoveItem iX: iX = iX - 1
                    End If
                    iX = iX + 1
                Next
            End With

            
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


Private Sub chMas_Click()
    If chMas.Value = vbChecked Then
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
    Ayuda "Ingrese una fecha de compra."
End Sub
Private Sub tFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco tFHasta
End Sub
Private Sub tFecha_LostFocus()
    If IsDate(tFecha.Text) Then
        tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
        If Not IsDate(tFHasta.Text) Then tFHasta.Text = tFecha.Text
    End If
    Ayuda ""
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
    
    cTipoListado.AddItem "Ingresos": cTipoListado.ItemData(cTipoListado.NewIndex) = 0
    cTipoListado.AddItem "Egresos": cTipoListado.ItemData(cTipoListado.NewIndex) = 1
    
    'vsListado.Orientation = orPortrait
    'vsListado.MarginBottom = 750: vsListado.MarginTop = 750
    'vsListado.MarginLeft = 450: vsListado.MarginRight = 450
    
    With vsListado
        .PhysicalPage = True
        .PaperSize = vbPRPSLetter
        .Orientation = orLandscape
        .PreviewMode = pmScreen
        .PreviewPage = 1
        .Zoom = 100
        .MarginLeft = 450: .MarginRight = 350
        .MarginBottom = 750: .MarginTop = 750
    End With
    
    With img1
        bConsultar.Picture = .ListImages("consultar").ExtractIcon
        bPrimero.Picture = .ListImages("move1").ExtractIcon
        bAnterior.Picture = .ListImages("move2").ExtractIcon
        bSiguiente.Picture = .ListImages("move3").ExtractIcon
        bUltima.Picture = .ListImages("move4").ExtractIcon
        
        bImprimir.Picture = .ListImages("print").ExtractIcon
        bConfigurar.Picture = .ListImages("configprint").ExtractIcon
        
        bNoFiltros.Picture = .ListImages("limpiar").ExtractIcon
        bCancelar.Picture = .ListImages("salir").ExtractIcon
        chVista.Picture = .ListImages("vista1").ExtractIcon
        chVista.DownPicture = .ListImages("vista2").ExtractIcon
        
    End With
    pbProgreso.Value = 0
    
    Exit Sub
ErrLoad:
    clsGeneral.OcurrioError "Ocurrió un error al inicializar el formulario.", Trim(Err.Description)
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
Dim fDebeHaber As String

    If Not ValidoCampos Then Exit Sub
    
    On Error GoTo errConsultar
    Screen.MousePointer = 11
    
    Select Case cTipoListado.ItemData(cTipoListado.ListIndex)
        Case 0: fDebeHaber = " MDRHaber = Null "   '0- Ingresos
        Case 1: fDebeHaber = " MDRDebe = Null "  '1- Egresos
    End Select

    bCargarImpresion = True
    
    vsConsulta.Rows = 1: vsConsulta.Refresh
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    cons = " Select Count(*) from MovimientoDisponibilidad, Compra, GastoSubRubro " _
            & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDiIDCompra = ComCodigo " _
            & " And ComCodigo = GSrIDCompra" _
            & " And MDiId In (Select MDRIdMovimiento from MovimientoDisponibilidadRenglon Where " & fDebeHaber & ")"
    
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
    
    cons = "Select * From MovimientoDisponibilidad, Compra, GastoSubRubro, SubRubro, Rubro, ProveedorCliente" _
            & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDiIDCompra = ComCodigo " _
            & " And ComCodigo = GSrIDCompra" _
            & " And GSrIDSubRubro = SRuID " _
            & " And SRuRubro = RubID" _
            & " And MDiId In (Select MDRIdMovimiento from MovimientoDisponibilidadRenglon Where " & fDebeHaber & ")" _
            & " And ComProveedor = PClCodigo"
    
    cons = cons & " Order by MDiID"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If rsAux.EOF Then
        MsgBox "No hay datos a desplegar para los filtros ingresados.", vbInformation, "ATENCION"
        rsAux.Close: Screen.MousePointer = 0: InicializoGrillas: Exit Sub
    End If
    
    ReDim arrTasas(0)           '-----------------------------------------------------------------------------------
    Dim mTC_OK As Currency
    Dim bCargarTasas As Boolean 'Las cargo si se consulta 1 mes sólo
    
    bCargarTasas = (Month(CDate(tFecha.Text)) = Month(CDate(tFHasta.Text)) And Year(CDate(tFecha.Text)) = Year(CDate(tFHasta.Text)))
    If bCargarTasas Then
         mTC_OK = TasadeCambio(paMonedaDolar, paMonedaPesos, DateAdd("d", -1, CDate("01/" & Mid(Format(tFecha.Text, "dd/mm/yyyy"), 4))))
    End If
    '--------------------------------------------------------------------------------------------------------------------
    
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
            
            If rsAux!ComMoneda <> paMonedaPesos Then If Not IsNull(rsAux!ComTC) Then .Cell(flexcpText, .Rows - 1, 7) = Format(rsAux!ComTC, "#0.000")
            If Not IsNull(rsAux!PClNombre) Then .Cell(flexcpText, .Rows - 1, 8) = Trim(rsAux!PClNombre)
            If Not IsNull(rsAux!ComComentario) Then .Cell(flexcpText, .Rows - 1, 9) = Trim(rsAux!ComComentario)
            .Cell(flexcpText, .Rows - 1, 10) = " "
            
            If IdCompra <> rsAux!MDiIdCompra Then
                If rsAux!ComMoneda = paMonedaPesos Then
                    .Cell(flexcpText, .Rows - 1, 2) = Format(Abs(rsAux!ComImporte), "#,##0.00")
                Else
                    .Cell(flexcpText, .Rows - 1, 2) = Format(Abs(rsAux!ComImporte) * rsAux!ComTC, "#,##0.00")
                End If
                
                If Not IsNull(rsAux!ComIVA) Then
                    If rsAux!ComMoneda = paMonedaPesos Then
                        .Cell(flexcpText, .Rows - 1, 4) = Format(Abs(rsAux!ComIVA), "#,##0.00")
                    Else
                        .Cell(flexcpText, .Rows - 1, 4) = Format(Abs(rsAux!ComIVA) * rsAux!ComTC, "#,##0.00")
                        .Cell(flexcpText, .Rows - 1, 6) = Format(Abs(rsAux!ComIVA), "#,##0.00")
                    End If
                    
                    If paSubrubroCompraMercaderia = rsAux!SRuID Then
                        IvaCompra = IvaCompra + .Cell(flexcpValue, .Rows - 1, 4)
                    Else
                        IvaGastos = IvaGastos + .Cell(flexcpValue, .Rows - 1, 4)
                    End If
                End If
                
                If Not IsNull(rsAux!ComCofis) Then
                    If rsAux!ComMoneda = paMonedaPesos Then
                        .Cell(flexcpText, .Rows - 1, 3) = Format(Abs(rsAux!ComCofis), "#,##0.00")
                    Else
                        .Cell(flexcpText, .Rows - 1, 3) = Format(Abs(rsAux!ComCofis) * rsAux!ComTC, "#,##0.00")
                        .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpValue, .Rows - 1, 6) + Abs(rsAux!ComCofis), "#,##0.00")
                    End If
                    
                    If paSubrubroCompraMercaderia = rsAux!SRuID Then
                        CofisCompra = CofisCompra + .Cell(flexcpValue, .Rows - 1, 3)
                    Else
                        CofisGastos = CofisGastos + .Cell(flexcpValue, .Rows - 1, 3)
                    End If
                End If
            End If
            
            If rsAux!ComMoneda = paMonedaPesos Then
                .Cell(flexcpText, .Rows - 1, 2) = Format(Abs(rsAux!GSrImporte), "#,##0.00")
            Else
                .Cell(flexcpText, .Rows - 1, 2) = Format(Abs(rsAux!GSrImporte) * rsAux!ComTC, "#,##0.00")
                .Cell(flexcpText, .Rows - 1, 6) = Format(.Cell(flexcpValue, .Rows - 1, 6) + Abs(rsAux!GSrImporte), "#,##0.00")
                
                'Cargo el array de movimientos en moedas extranjeras
                If bCargarTasas And mTC_OK <> rsAux!ComTC Then
                    arr_Agregar rsAux!ComMoneda, rsAux!ComTC, .Cell(flexcpText, .Rows - 1, 1), _
                                .Cell(flexcpValue, .Rows - 1, 2), .Cell(flexcpValue, .Rows - 1, 3), .Cell(flexcpValue, .Rows - 1, 4), rsAux!MDiIdCompra
                End If
            End If
            .Cell(flexcpText, .Rows - 1, 5) = Format(.Cell(flexcpValue, .Rows - 1, 2) + .Cell(flexcpValue, .Rows - 1, 3) + .Cell(flexcpValue, .Rows - 1, 4), "#,##0.00")
                        
            IdCompra = rsAux!MDiIdCompra
            
            rsAux.MoveNext
        End With
    Loop
    rsAux.Close
    
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
        .Subtotal flexSTSum, -1, 2, , Colores.Obligatorio, &H80&, True, "Total"
        .Subtotal flexSTSum, -1, 3: .Subtotal flexSTSum, -1, 4: .Subtotal flexSTSum, -1, 5: .Subtotal flexSTSum, -1, 6
        
        If IvaCompra <> 0 Then
            .AddItem "I.V.A. Compras  21401"
            .Cell(flexcpText, .Rows - 1, 4) = Format(IvaCompra, "#,##0.00")
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
        End If
        If IvaGastos <> 0 Then
            .AddItem "I.V.A. Gastos     21403"
            .Cell(flexcpText, .Rows - 1, 4) = Format(IvaGastos, "#,##0.00")
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
        End If
        If CofisCompra <> 0 Then
            .AddItem "Cofis Compras    21411"
            .Cell(flexcpText, .Rows - 1, 3) = Format(CofisCompra, "#,##0.00")
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
        End If
        If CofisGastos <> 0 Then
            .AddItem "Cofis Gastos       21413"
            .Cell(flexcpText, .Rows - 1, 3) = Format(CofisGastos, "#,##0.00")
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Obligatorio: .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
        End If
    End With
    
    AgrupoCamposEnGrilla
    
    CargoDetalleAcreedores
    CargoDisponibilidades
    CargoDetalleTasasDeCambio
    
    vsConsulta.Redraw = True
    pbProgreso.Value = 0
    Screen.MousePointer = 0
    Exit Sub
    
errConsultar:
    clsGeneral.OcurrioError "Error al realizar la consulta de datos.", Err.Description
    vsConsulta.Redraw = True: Screen.MousePointer = 0
End Sub

'Contra que van
Private Sub CargoDisponibilidades()

Dim mCellData As Long

    cons = "Select DisID, DisNombre, DisMoneda, DisIDSRCheque, SRuID, SRuCodigo, SRuNombre, Importe = Sum(MDRImportePesos), Debe = Sum(MDRDebe), Haber = Sum(MDRHaber) " _
            & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, Disponibilidad, SubRubro" _
            & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDiID = MDRIDMovimiento" _
            & " And MDiIDCompra Is Not Null and MDRIdDisponibilidad = DisID" _
            & " And DisIDSubrubro = SRuID "
    
    Select Case cTipoListado.ItemData(cTipoListado.ListIndex)
        Case 0: cons = cons & " And MDRHaber = Null "   '0- Ingresos
        Case 1: cons = cons & " And MDRDebe = Null "    '1- Egresos
    End Select
    
    cons = cons & " Group by DisID, DisNombre, DisMoneda, DisIDSRCheque, SRuID, SRuCodigo, SRuNombre"
    cons = cons & " Order by SRuNombre"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If rsAux.EOF Then rsAux.Close: Exit Sub
    
    With vsConsulta
        .AddItem "": .AddItem ""
        
        Select Case cTipoListado.ItemData(cTipoListado.ListIndex)
            Case 0: .Cell(flexcpText, .Rows - 1, 1) = "Conceptos al DEBE" '0- Ingresos
            Case 1: .Cell(flexcpText, .Rows - 1, 1) = "Conceptos al HABER" '1- Egresos
        End Select
        
        .Cell(flexcpData, .Rows - 1, 7) = "DT"
        
        .Cell(flexcpBackColor, .Rows - 1, 1, , .Cols - 1) = Colores.Azul
        .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Blanco
        .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
    End With
    
    Dim aTotal As Currency: aTotal = 0
    Dim aTotalME As Currency: aTotalME = 0
    Dim aImporte As Currency, aImporteDH As Currency
    Dim RsCh As rdoResultset
    
    Do While Not rsAux.EOF
        
        With vsConsulta
            aImporte = rsAux!Importe
            If Not IsNull(rsAux!Debe) Then aImporteDH = rsAux!Debe
            If Not IsNull(rsAux!Haber) Then aImporteDH = rsAux!Haber
            
            'Hay que ver (si la disponibilidad es bancaria, si los monimientos son con cheques diferidos)
            If Not IsNull(rsAux!DisIDSRCheque) Then
                
                cons = "Select SRuCodigo, SRuNombre, Importe = Sum(MDRImportePesos), Debe = Sum(MDRDebe), Haber = Sum(MDRHaber) " _
                        & " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, Disponibilidad, SubRubro, Cheque" _
                        & " Where MDiFecha Between '" & Format(tFecha.Text, "mm/dd/yyyy") & "' AND '" & Format(tFHasta.Text, "mm/dd/yyyy") & "'" _
                        & " And MDiID = MDRIDMovimiento" _
                        & " And MDRIdDisponibilidad = " & rsAux!DisID _
                        & " And MDiIDCompra Is Not Null And MDRIdDisponibilidad = DisID" _
                        & " And DisIDSRCheque = SRuID " _
                        & " And MDRIDCheque = CheID And CheVencimiento Is Not Null" _
                        & " And CheLibrado Between '" & Format(tFecha.Text, "mm/dd/yyyy") & "' AND '" & Format(tFHasta.Text, "mm/dd/yyyy") & "'"
                
                '& " And MDRIdDisponibilidad " IN (Select DisID from Disponibilidad Where DisIDSubrubro =  " & rsAux!SRuID & ")"
                
                Select Case cTipoListado.ItemData(cTipoListado.ListIndex)
                    Case 0: cons = cons & " And MDRHaber = Null "   '0- Ingresos
                    Case 1: cons = cons & " And MDRDebe = Null "    '1- Egresos
                End Select
                cons = cons & " Group by SRuCodigo, SRuNombre"
                
                Set RsCh = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
                If Not RsCh.EOF Then
                    .AddItem ""
                    .Cell(flexcpText, .Rows - 1, 0) = rsAux!DisNombre
                    mCellData = rsAux!DisID
                    .Cell(flexcpData, .Rows - 1, 1) = mCellData: .Cell(flexcpData, .Rows - 1, 7) = "DC" 'Disponib c/Cheque
                    
                    .Cell(flexcpText, .Rows - 1, 1) = Format(RsCh!SRuCodigo, "000000000") & " " & Trim(RsCh!SRuNombre) & " (" & Trim(rsAux!SRuNombre) & ")"
                    .Cell(flexcpText, .Rows - 1, 5) = Format(RsCh!Importe, FormatoMonedaP)
                    If rsAux!DisMoneda <> paMonedaPesos Then
                        
                        If Not IsNull(RsCh!Debe) Then .Cell(flexcpText, .Rows - 1, 6) = Format(RsCh!Debe, FormatoMonedaP)
                        If Not IsNull(RsCh!Haber) Then .Cell(flexcpText, .Rows - 1, 6) = Format(RsCh!Haber, FormatoMonedaP)
                        .Cell(flexcpText, .Rows - 1, 7) = "pTC " & Format(aImporte / .Cell(flexcpValue, .Rows - 1, 6), "#,##0.000")
                        
                        aTotalME = aTotalME + .Cell(flexcpValue, .Rows - 1, 6)
                    End If
                    aTotal = aTotal + .Cell(flexcpText, .Rows - 1, 5)
                    
                    aImporte = aImporte - .Cell(flexcpValue, .Rows - 1, 5)
                    aImporteDH = aImporteDH - .Cell(flexcpValue, .Rows - 1, 6)
                End If
                RsCh.Close
            End If
            If aImporte <> 0 Then
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = rsAux!DisNombre
                mCellData = rsAux!DisID
                .Cell(flexcpData, .Rows - 1, 1) = mCellData: .Cell(flexcpData, .Rows - 1, 7) = "DN" 'Disponib Normal
                
                .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!SRuCodigo, "000000000") & " " & Trim(rsAux!SRuNombre)
                .Cell(flexcpText, .Rows - 1, 5) = Format(aImporte, FormatoMonedaP)
                If rsAux!DisMoneda <> paMonedaPesos Then
                    .Cell(flexcpText, .Rows - 1, 6) = Format(aImporteDH, FormatoMonedaP)
                    .Cell(flexcpText, .Rows - 1, 7) = "pTC " & Format(aImporte / aImporteDH, "#,##0.000")
                    aTotalME = aTotalME + .Cell(flexcpValue, .Rows - 1, 6)
                End If
                
                aTotal = aTotal + .Cell(flexcpText, .Rows - 1, 5)
            End If
        End With
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    With vsConsulta
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = "Total"
        .Cell(flexcpText, .Rows - 1, 5) = Format(aTotal, FormatoMonedaP)
        If aTotalME <> 0 Then .Cell(flexcpText, .Rows - 1, 6) = Format(aTotalME, FormatoMonedaP)
        .Cell(flexcpBackColor, .Rows - 1, 1, , .Cols - 1) = Colores.Obligatorio
        .Cell(flexcpForeColor, .Rows - 1, 0, , .Cols - 1) = Colores.Rojo
        .Cell(flexcpFontBold, .Rows - 1, 0, , .Cols - 1) = True
    End With
    
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
        If chMas.Value = vbUnchecked Then
            vsListado.Orientation = orPortrait
            vsConsulta.ColWidth(7) = 1000
            'vsConsulta.ColWidth(8) = 1200
            'vsConsulta.ColWidth(9) = 1000

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
        
        EncabezadoListado vsListado, "Asientos de Gastos (" & Trim(cTipoListado.Text) & ") - Del " & Trim(tFecha.Text) & " al " & Trim(tFHasta.Text), False
        vsListado.FileName = "Listado Asientos de Gastos"
            
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        
        vsListado.EndDoc
        'bCargarImpresion = False
        If chMas.Value = vbChecked Then vsConsulta.ColWidth(7) = 600: vsConsulta.ColWidth(8) = 1400
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

Private Sub Ayuda(strTexto As String)
    Status.Panels(4).Text = strTexto
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
On Error GoTo errFnc
    
    Select Case vsConsulta.Cell(flexcpData, vsConsulta.Row, 7)
        
        Case "G": EjecutarApp prmPathApp & "Ingreso de Facturas.exe", vsConsulta.Cell(flexcpValue, vsConsulta.Row, 7)
        
        Case "DN", "DC":
                    If Not ValidoCampos Then Exit Sub
                    
                    Dim mFrm As New frmDetalleD
                    With mFrm
                        .prmDisponibilidadID = vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)
                        .prmFecha1 = tFecha.Text
                        .prmFecha2 = tFHasta.Text
                        .prmListarDiferidos = ("DC" = vsConsulta.Cell(flexcpData, vsConsulta.Row, 7))
                        
                         Select Case cTipoListado.ItemData(cTipoListado.ListIndex)
                            Case 0: .prmAlHaber = False '0- Ingresos
                            Case 1: .prmAlHaber = True '1- Egresos
                        End Select
                        
                        .Show , Me
                        .fnc_AccionConsultar
                    End With
    
        Case "DT":
                    If Not ValidoCampos Then Exit Sub
                    
                    Dim xROW As Integer, bEXIT As Boolean
                    
                    Dim mFrmG As New frmDetalleD
                    With mFrm
                        
                        .prmFecha1 = tFecha.Text
                        .prmFecha2 = tFHasta.Text
                        .prmAlHaber = (cTipoListado.ItemData(cTipoListado.ListIndex) = 1)
                        
                        .Show , Me
                        
                        xROW = vsConsulta.Row + 1
                        bEXIT = Not (("DC" = vsConsulta.Cell(flexcpData, xROW, 7)) Or ("DN" = vsConsulta.Cell(flexcpData, xROW, 7)))
                        Do While Not bEXIT
                            
                            .prmListarDiferidos = ("DC" = vsConsulta.Cell(flexcpData, xROW, 7))
                            .prmDisponibilidadID = vsConsulta.Cell(flexcpData, xROW, 1)
                            xROW = xROW + 1
                            
                            bEXIT = Not (("DC" = vsConsulta.Cell(flexcpData, xROW, 7)) Or ("DN" = vsConsulta.Cell(flexcpData, xROW, 7)))
                            
                            .fnc_AccionConsultar
                        Loop
                    End With
    
    End Select
    
errFnc:
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
        MsgBox "Debe seleccionar el tipo de movimientos a listr (Ingresos/Egresos).", vbExclamation, "ATENCIÓN"
        Foco cTipoListado: Exit Function
    End If
    ValidoCampos = True
    
End Function

Private Sub CargoDetalleAcreedores()

Dim aQAcreedores As Long, aIdEmpresa As Long
Dim aAVarios As Currency, aAProveedor As Currency
Dim aAVariosME As Currency, aAProveedorME As Currency
Dim txtSubrubro As String
Dim rsPro As rdoResultset
Dim fDebeHaber As String
Dim bAgregeAcreedor As Boolean

    On Error GoTo errConsultar
    aQAcreedores = 0: aIdEmpresa = 0
    aAVarios = 0: aAProveedor = 0
    aAVariosME = 0: aAProveedorME = 0
    bAgregeAcreedor = False
    
    Select Case cTipoListado.ItemData(cTipoListado.ListIndex)
        Case 0: fDebeHaber = " MDRHaber = Null "   '0- Ingresos
        Case 1: fDebeHaber = " MDRDebe = Null "  '1- Egresos
    End Select
    
    'Inicializo Progress Bar-----------------------------------------------------------------------------------------------------------------
    cons = " Select Count(*) from MovimientoDisponibilidad, Compra, GastoSubRubro " _
            & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDiIDCompra = ComCodigo " _
            & " And ComCodigo = GSrIDCompra" _
            & " And GSrIDSubRubro = " & paSubrubroAcreedoresVarios _
            & " And MDiId In (Select MDRIdMovimiento from MovimientoDisponibilidadRenglon Where " & fDebeHaber & ")"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    If Not rsAux.EOF Then If Not IsNull(rsAux(0)) Then aQAcreedores = rsAux(0)
    rsAux.Close
    If aQAcreedores = 0 Then Exit Sub
    
    pbProgreso.Max = aQAcreedores
    '-----------------------------------------------------------------------------------------------------------------
    
    cons = "Select * From MovimientoDisponibilidad, Compra, GastoSubRubro" _
            & " Where MDiFecha Between '" & Format(tFecha.Text, sqlFormatoF) & "' AND '" & Format(tFHasta.Text, sqlFormatoF) & "'" _
            & " And MDiIDCompra = ComCodigo " _
            & " And ComCodigo = GSrIDCompra" _
            & " And GSrIDSubRubro = " & paSubrubroAcreedoresVarios _
            & " And MDiId In (Select MDRIdMovimiento from MovimientoDisponibilidadRenglon Where " & fDebeHaber & ")"
    cons = cons & " Order by ComProveedor"
    
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    aIdEmpresa = 0
    pbProgreso.Value = 0
    Do While Not rsAux.EOF
        pbProgreso.Value = pbProgreso.Value + 1
        
        If aIdEmpresa <> rsAux!ComProveedor Then
            If aIdEmpresa <> 0 Then
                If aAProveedor > 0 Then         'x si no va directo a AVarios
                    With vsConsulta
                        If Not bAgregeAcreedor Then .AddItem "Detalle Acreedores" Else .AddItem ""
                        bAgregeAcreedor = True
                        .Cell(flexcpText, .Rows - 1, 1) = txtSubrubro
                        .Cell(flexcpText, .Rows - 1, 5) = Format(aAProveedor, FormatoMonedaP)
                        .Cell(flexcpText, .Rows - 1, 6) = Format(aAProveedorME, FormatoMonedaP)
                        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
                    End With
                End If
            End If
            
            aIdEmpresa = rsAux!ComProveedor
            'Saco datos del Proveedor del Gasto
            cons = " Select * from EmpresaDato Left Outer Join SubRubro On EDaSRubroContable = SRuID" & _
                        " Where EDaCodigo = " & aIdEmpresa & _
                        " And EDaTipoEmpresa = " & TipoEmpresa.Cliente
            Set rsPro = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rsPro.EOF Then
                If Not IsNull(rsPro!SRuNombre) Then txtSubrubro = Format(rsPro!SRuCodigo, "000000000") & " " & Trim(rsPro!SRuNombre) Else txtSubrubro = "Sin Datos"
            End If
            rsPro.Close
            aAProveedor = 0: aAProveedorME = 0
        End If
        
        'Tengo que consultar a que estaba asignada la compra original
        cons = "Select * From CompraPago, Compra, GastoSubRubro" _
                & " Where CPaDocQSalda = " & rsAux!Comcodigo _
                & " And CPaDocASaldar = ComCodigo " _
                & " And ComCodigo = GSrIDCompra" _
                & " And GSrIdSubRubro = " & paSubrubroDivisa
        Set rsPro = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsPro.EOF Then
            If rsPro!ComMoneda = paMonedaPesos Then
                aAProveedor = aAProveedor + Abs(rsAux!GSrImporte)
            Else
                aAProveedor = aAProveedor + Format(Abs(rsAux!GSrImporte) * rsAux!ComTC, FormatoMonedaP)
                aAProveedorME = aAProveedorME + Abs(rsAux!GSrImporte)
            End If
        Else
            If rsAux!ComMoneda = paMonedaPesos Then
                aAVarios = aAVarios + Abs(rsAux!GSrImporte)
            Else
                aAVarios = aAVarios + Format(Abs(rsAux!GSrImporte) * rsAux!ComTC, FormatoMonedaP)
                aAVariosME = aAVariosME + Abs(rsAux!GSrImporte)
            End If
        End If
        rsPro.Close
        
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    If aAProveedor > 0 Then         'x si no va directo a AVarios
        With vsConsulta
            If Not bAgregeAcreedor Then .AddItem "Detalle Acreedores" Else .AddItem ""
            bAgregeAcreedor = True
            .Cell(flexcpText, .Rows - 1, 1) = txtSubrubro
            .Cell(flexcpText, .Rows - 1, 5) = Format(aAProveedor, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = Format(aAProveedorME, FormatoMonedaP)
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
        End With
    End If
    
    If aAVarios > 0 And bAgregeAcreedor Then
        'Saco datos del Subrubro General
        cons = " Select * from SubRubro Where SRuID = " & paSubrubroAcreedoresVarios
        Set rsPro = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
        If Not rsPro.EOF Then txtSubrubro = Format(rsPro!SRuCodigo, "000000000") & " " & Trim(rsPro!SRuNombre)
        rsPro.Close
        
        With vsConsulta
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 1) = txtSubrubro
            .Cell(flexcpText, .Rows - 1, 5) = Format(aAVarios, FormatoMonedaP)
            .Cell(flexcpText, .Rows - 1, 6) = Format(aAVariosME, FormatoMonedaP)
            .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Inactivo
        End With
    End If
    '------------------------------------------------------------------------------------------------------------------------------------------------------
    pbProgreso.Value = 0
    Exit Sub
    
errConsultar:
    clsGeneral.OcurrioError "Ocurrió un error al cargar el detalle de acreedores.", Err.Description
End Sub

Private Function CargoDetalleTasasDeCambio()

    If arrTasas(0).Moneda = 0 Then Exit Function
    Dim mQTs As Integer: mQTs = 0
    
    arr_Sort
    
    'Por ahora voy a ignorar la moneda ya que los gastos se ingresan en dolares o pesos 21/03/2003
    Dim mTC As Currency, mNeto As Currency, mCofis As Currency, mIVA As Currency
    
    mTC = 0
    
    With vsConsulta
        .AddItem ""
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = "Movimientos en Moneda Extranjera "
        .Cell(flexcpBackColor, .Rows - 1, 0, , .Cols - 1) = Colores.Gris
    End With
    
    For I = LBound(arrTasas) To UBound(arrTasas)
        With vsConsulta
            
            .AddItem ""
            If mTC <> arrTasas(I).Tasa Then
                
                If mTC <> 0 Then    'TOTAL DE LA TC
                    If mQTs <> 1 Then
                        .Cell(flexcpText, .Rows - 1, 2) = Format(mNeto, "#,##0.00")
                        .Cell(flexcpText, .Rows - 1, 3) = Format(mCofis, "#,##0.00")
                        .Cell(flexcpText, .Rows - 1, 4) = Format(mIVA, "#,##0.00")
                        .Cell(flexcpText, .Rows - 1, 5) = Format(mNeto + mCofis + mIVA, "#,##0.00")
                        .Cell(flexcpText, .Rows - 1, 6) = Format((mNeto + mCofis + mIVA) / mTC, "#,##0.00")
                        .Cell(flexcpBackColor, .Rows - 1, 2, , 6) = Colores.Gris
                        
                        .AddItem ""
                    Else
                        .Cell(flexcpBackColor, .Rows - 2, 2, , 6) = Colores.Gris
                    End If
                End If
                
                .Cell(flexcpText, .Rows - 1, 0) = "TC " & Format(arrTasas(I).Tasa, "#.000")
                .Cell(flexcpAlignment, .Rows - 1, 0) = flexAlignRightCenter
                
                
                mTC = arrTasas(I).Tasa
                mNeto = 0: mCofis = 0: mIVA = 0
                mQTs = 0
            End If
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(arrTasas(I).Rubro)
            .Cell(flexcpText, .Rows - 1, 2) = Format(arrTasas(I).Neto, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 3) = Format(arrTasas(I).Cofis, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 4) = Format(arrTasas(I).IVA, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 5) = Format(arrTasas(I).Neto + arrTasas(I).Cofis + arrTasas(I).IVA, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 6) = Format((arrTasas(I).Neto + arrTasas(I).Cofis + arrTasas(I).IVA) / mTC, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 7) = Format(arrTasas(I).IDGasto, "###,###"): .Cell(flexcpData, .Rows - 1, 7) = "G"
            
            mNeto = mNeto + arrTasas(I).Neto
            mCofis = mCofis + arrTasas(I).Cofis
            mIVA = mIVA + arrTasas(I).IVA
            mQTs = mQTs + 1
            
            If I = UBound(arrTasas) Then
                If mQTs > 1 Then   'TOTAL DE LA TC
                    .AddItem ""
                    .Cell(flexcpText, .Rows - 1, 2) = Format(mNeto, "#,##0.00")
                    .Cell(flexcpText, .Rows - 1, 3) = Format(mCofis, "#,##0.00")
                    .Cell(flexcpText, .Rows - 1, 4) = Format(mIVA, "#,##0.00")
                    .Cell(flexcpText, .Rows - 1, 5) = Format(mNeto + mCofis + mIVA, "#,##0.00")
                    .Cell(flexcpText, .Rows - 1, 6) = Format((mNeto + mCofis + mIVA) / mTC, "#,##0.00")
                End If
                .Cell(flexcpBackColor, .Rows - 1, 2, , 6) = Colores.Gris
            End If
        End With
    
    Next
    
    Dim mME As Currency
    mNeto = 0: mCofis = 0: mIVA = 0: mME = 0
    With vsConsulta
        For I = LBound(arrTasas) To UBound(arrTasas)
            mNeto = mNeto + arrTasas(I).Neto
            mCofis = mCofis + arrTasas(I).Cofis
            mIVA = mIVA + arrTasas(I).IVA
            mME = mME + ((arrTasas(I).Neto + arrTasas(I).Cofis + arrTasas(I).IVA) / arrTasas(I).Tasa)
        Next
        .AddItem "": .AddItem ""
        .Cell(flexcpText, .Rows - 1, 1) = "TOTAL"
        .Cell(flexcpText, .Rows - 1, 2) = Format(mNeto, "#,##0.00")
        .Cell(flexcpText, .Rows - 1, 3) = Format(mCofis, "#,##0.00")
        .Cell(flexcpText, .Rows - 1, 4) = Format(mIVA, "#,##0.00")
        .Cell(flexcpText, .Rows - 1, 5) = Format(mNeto + mCofis + mIVA, "#,##0.00")
        .Cell(flexcpText, .Rows - 1, 6) = Format(mME, "#,##0.00")
        .Cell(flexcpBackColor, .Rows - 1, 2, , 6) = Colores.Gris
    End With
End Function

'/* Movimientos para una Disponibilidad */
'Select RubCodigo, RubNombre, SRuCodigo, SRuNombre, GSrImporte, MDRImportePesos, MDRImporteCompra, MDRDebe, MDRHaber
'From MovimientoDisponibilidad, MovimientoDisponibilidadRenglon, Compra, GastoSubRubro, SubRubro, Rubro
'Where MDiFecha Between '03/15/2003' AND '04/15/2003'
'And MDiIDCompra = ComCodigo  And ComCodigo = GSrIDCompra
'And GSrIDSubRubro = SRuID  And SRuRubro = RubID
'And MDRDebe is Null
'And MDiID = MDRIDMovimiento
'And MDRIDDisponibilidad = 3
'Order by MDiID

