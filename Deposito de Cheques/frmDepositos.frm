VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDepositos 
   BackColor       =   &H8000000B&
   Caption         =   "Depósito de Cheques"
   ClientHeight    =   6180
   ClientLeft      =   420
   ClientTop       =   2280
   ClientWidth     =   11730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDepositos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   11730
   Begin vsViewLib.vsPrinter vsListado 
      Height          =   3795
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   9615
      _Version        =   196608
      _ExtentX        =   16960
      _ExtentY        =   6694
      _StockProps     =   229
      BackColor       =   -2147483634
      BorderStyle     =   1
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
      BackColor       =   -2147483634
   End
   Begin VB.PictureBox picBotones 
      Height          =   495
      Left            =   60
      ScaleHeight     =   435
      ScaleWidth      =   11595
      TabIndex        =   12
      Top             =   5640
      Width           =   11655
      Begin VB.CommandButton bGrabar 
         Enabled         =   0   'False
         Height          =   310
         Left            =   5400
         Picture         =   "frmDepositos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bPrimero 
         Height          =   310
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la primer página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bSiguiente 
         Height          =   310
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la siguiente página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bAnterior 
         Height          =   310
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la página anterior."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConsultar 
         Height          =   310
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Ejecutar."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   4860
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bNoFiltros 
         Height          =   310
         Left            =   4380
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Quitar filtros."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   3300
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bUltima 
         Height          =   310
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Ir a la última página."
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMas 
         Height          =   310
         Left            =   2340
         Picture         =   "frmDepositos.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Zoom Out"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bZMenos 
         Height          =   310
         Left            =   2700
         Picture         =   "frmDepositos.frx":062E
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Zoom In"
         Top             =   120
         Width           =   310
      End
      Begin VB.CommandButton bConfigurar 
         Height          =   310
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Configurar impresora."
         Top             =   120
         Width           =   310
      End
      Begin VB.CheckBox chVista 
         Height          =   310
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   310
      End
      Begin VB.Label lQch 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   9600
         TabIndex        =   28
         Top             =   135
         Width           =   435
      End
      Begin VB.Label lTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7560
         TabIndex        =   26
         Top             =   135
         Width           =   1815
      End
      Begin VB.Label lTTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total a Depositar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   5760
         TabIndex        =   27
         Top             =   120
         Width           =   4335
      End
   End
   Begin VB.Frame frmFiltro 
      Caption         =   "Filtros"
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   11835
      Begin VB.ComboBox cboDe 
         Height          =   315
         Left            =   10320
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cTipo 
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin AACombo99.AACombo cMoneda 
         Height          =   315
         Left            =   4680
         TabIndex        =   3
         Top             =   705
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
      Begin VB.TextBox tFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   705
         Width           =   2055
      End
      Begin AACombo99.AACombo cCondicion 
         Height          =   315
         Left            =   6480
         TabIndex        =   5
         Top             =   705
         Width           =   1515
         _ExtentX        =   2672
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
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&De:"
         Height          =   255
         Left            =   9720
         TabIndex        =   30
         Top             =   750
         Width           =   555
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Tipo:"
         Height          =   255
         Left            =   7920
         TabIndex        =   6
         Top             =   750
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de &Vencimiento:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   730
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Moneda:"
         Height          =   255
         Left            =   4020
         TabIndex        =   2
         Top             =   730
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmDepositos.frx":0718
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   180
         Width           =   8655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "&Condición:"
         Height          =   255
         Left            =   5640
         TabIndex        =   4
         Top             =   730
         Width           =   915
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   4335
      Left            =   3060
      TabIndex        =   10
      ToolTipText     =   "Fondo gris es de otra sucursal, crema es sin documento asociado"
      Top             =   780
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
      FocusRect       =   2
      HighLight       =   0
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
   Begin MSComctlLib.ImageList img1 
      Left            =   1200
      Top             =   1200
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
            Picture         =   "frmDepositos.frx":0800
            Key             =   "salir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDepositos.frx":0B1A
            Key             =   "help"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDepositos.frx":0C2C
            Key             =   "consultar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDepositos.frx":0D86
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDepositos.frx":0EE0
            Key             =   "print"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDepositos.frx":103A
            Key             =   "limpiar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDepositos.frx":114C
            Key             =   "vista2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDepositos.frx":12A6
            Key             =   "vista1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDepositos.frx":1400
            Key             =   "move2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDepositos.frx":155A
            Key             =   "move3"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDepositos.frx":16B4
            Key             =   "move4"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDepositos.frx":180E
            Key             =   "move1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDepositos.frx":1968
            Key             =   "configprint"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList Image1 
      Left            =   120
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDepositos.frx":1A7A
            Key             =   "check"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDepositos.frx":1D94
            Key             =   "nocheck"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuBDerecho 
      Caption         =   "BotonDerecho"
      Visible         =   0   'False
      Begin VB.Menu MnuTitulo 
         Caption         =   "Menú Cheques"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnuL1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMarcar 
         Caption         =   "&Marcar por selección"
      End
      Begin VB.Menu MnuVerSucursal 
         Caption         =   "&Editar Sucursal a Depositar"
      End
      Begin VB.Menu MnuSeguimiento 
         Caption         =   "&Seguimiento de Cheques"
      End
      Begin VB.Menu MnuVerDeuda 
         Caption         =   "&Deuda en Cheques"
      End
      Begin VB.Menu MnuVerL2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVerCancelar 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "frmDepositos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim aTexto As String
Dim bCargarImpresion As Boolean

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bConsultar_Click()

    'Valido el ingreso de los campos para realizar la consulta.------------------------------------------
    aTexto = ValidoPeriodoFechas(tFecha.Text) ', True)
    If aTexto = "" Then
        MsgBox "El período de fechas ingresado no es correcto.", vbExclamation, "ATENCIÓN"
        Foco tFecha: Exit Sub
    End If
    If cMoneda.ListIndex = -1 Then
        MsgBox "Debe seleccionar una moneda para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco cMoneda: Exit Sub
    End If
    
    If cCondicion.ListIndex = -1 Then
        MsgBox "Debe seleccionar la condición de búsqueda para realizar la consulta.", vbExclamation, "ATENCIÓN"
        Foco cCondicion: Exit Sub
    End If
    '------------------------------------------------------------------------------------------------------------
    
    CargoCheques Trim(tFecha.Text)
    If vsConsulta.Rows > 1 Then vsConsulta.SetFocus
    
End Sub

Private Sub bGrabar_Click()
    
    If CCur(lTotal.Caption) = 0 Then
        MsgBox "Debe seleccionar los cheques a depositar.", vbExclamation, "No hay datos"
        Exit Sub
    End If
    
    If Not fnc_ValidoCambios Then Exit Sub
    
    If Not CargoDataGastos Then Exit Sub
    
    frmWizGasto.Show vbModal, Me
    If Not frmWizGasto.prmOK Then Exit Sub
    
    If MsgBox("Confirma grabar el depósito de los cheques seleccionados." & vbCrLf & vbCrLf & _
                    "Además se van a grabar los Gastos o Transferencias relacionados a los depósitos.", vbQuestion + vbYesNo, "Grabar Depósitos") = vbNo Then Exit Sub
    
    AccionGrabar
    
End Sub

Private Function fnc_ValidoCambios() As Boolean

    On Error GoTo errValido
    fnc_ValidoCambios = True
    
    If cMoneda.ListIndex = -1 Then
        MsgBox "Seleccione la moneda para realizar los movimientos por el concepto Depósitos", vbExclamation, "Falta Moneda"
        fnc_ValidoCambios = False
        Exit Function
    End If

Dim misCheques As String
Dim idx As Integer

    Screen.MousePointer = 11
    
    With vsConsulta
        
        For idx = .FixedRows To .Rows - 1
            If .Cell(flexcpBackColor, idx, 0, , .Cols - 1) <> Colores.Inactivo Then
                If misCheques <> "" Then misCheques = misCheques & ","
                misCheques = misCheques & .Cell(flexcpData, idx, 0)
            End If
        Next
        
    End With
    
    Cons = "Select * from ChequeDiferido Where CDiCodigo IN (" & misCheques & ")"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not RsAux.EOF
        With vsConsulta
            'Busco el cheque para comparar el importe
            For idx = .FixedRows To .Rows - 1
                If RsAux!CDiCodigo = .Cell(flexcpData, idx, 0) Then
                    If Format(RsAux!CDiImporte, FormatoMonedaP) <> .Cell(flexcpText, idx, 3) Or RsAux!CDiMoneda <> cMoneda.ItemData(cMoneda.ListIndex) Then
                        fnc_ValidoCambios = False
                        Exit Do
                    End If
                    Exit For
                End If
            Next
        End With
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If Not fnc_ValidoCambios Then
        MsgBox "Los importes de los cheques seleccionados para depositar se modificaron." & vbCrLf & _
                    "Vuelva a cargar los datos para grabar el depósito.", vbExclamation, "Datos Modificados"
    End If
    
    Screen.MousePointer = 0
    Exit Function
errValido:
    clsGeneral.OcurrioError "Error al validar el importe de los cheques.", Err.Description
    Screen.MousePointer = 0
End Function

Private Function CargoDataGastos() As Boolean

Dim I As Integer, bOK As Boolean
    
    CargoDataGastos = False
    Dim mMoneda As Long
    mMoneda = cMoneda.ItemData(cMoneda.ListIndex)
    
    With vsConsulta
    ReDim dGastos(0)
    
    For I = 1 To .Rows - 1
        If .Cell(flexcpBackColor, I, 0, , .Cols - 1) <> Colores.Inactivo Then
            
            bOK = arrG_AddItem(.Cell(flexcpData, I, 5), .Cell(flexcpText, I, 5), .Cell(flexcpValue, I, 3), (.Cell(flexcpForeColor, I, 0) = Colores.Azul), mMoneda, .Cell(flexcpData, I, 2))
            
            If Not bOK Then
                clsGeneral.OcurrioError "Errores al procesar la lista de cheques.", Err.Description
                Exit Function
            End If
        End If
    Next
    
    End With
    CargoDataGastos = True
    
End Function

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

Private Sub cboDe_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then bConsultar.SetFocus
End Sub

Private Sub chVista_Click()
    
    If chVista.Value = 0 Then
        vsListado.Visible = False
    Else
        AccionImprimir
        vsListado.Visible = True: vsListado.ZOrder 0
    End If
    Me.Refresh

End Sub

Private Sub cCondicion_Click()

    If cCondicion.ListIndex = -1 Then Exit Sub
    Select Case cCondicion.ItemData(cCondicion.ListIndex)
        Case 1: lTTotal.Caption = " Total a Depositar:"
        Case 2: lTTotal.Caption = " Total Depositado:"
    End Select
        
End Sub

Private Sub cCondicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And cCondicion.ListIndex <> -1 Then bConsultar.SetFocus
End Sub

Private Sub cMoneda_Change()
    If vsConsulta.Rows > 1 Then vsConsulta.Rows = 1
End Sub

Private Sub cMoneda_Click()
    If vsConsulta.Rows > 1 Then vsConsulta.Rows = 1
End Sub

Private Sub cMoneda_GotFocus()
    cMoneda.SelStart = 0: cMoneda.SelLength = Len(cMoneda.Text)
End Sub

Private Sub cMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Foco cCondicion
End Sub

Private Sub cTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cboDe.SetFocus
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    Screen.MousePointer = 11
    picBotones.BorderStyle = 0
    ObtengoSeteoForm Me
    InicializoGrilla
    AccionLimpiar
    
    tFecha.Text = Format(Now, "d/mm/yyyy")
   
    'Cargo las monedas en el combo-------------------------
    Cons = "Select MonCodigo, MonSigno from Moneda where MonFactura = 1 Order by MonSigno"
    CargoCombo Cons, cMoneda, ""
    '--------------------------------------------------------------
    
    cCondicion.AddItem "A Depositar": cCondicion.ItemData(cCondicion.NewIndex) = 1
    cCondicion.AddItem "Depositados": cCondicion.ItemData(cCondicion.NewIndex) = 2
    
    cTipo.AddItem "(Todos)", 0
    cTipo.AddItem "Al Día", 1
    cTipo.AddItem "Diferidos", 2
    cTipo.ListIndex = 0
    
'    cboDe.AddItem "(Todos)", 0
    cboDe.AddItem "Comercio", 0
    cboDe.AddItem "Fleteros", 1
    cboDe.ListIndex = 0
        
    BuscoCodigoEnCombo cMoneda, CLng(paMonedaPesos)
    BuscoCodigoEnCombo cCondicion, 1
    
    vsListado.Zoom = 100
    vsListado.PaperSize = 1
    vsListado.MarginLeft = 750
    vsListado.Orientation = orPortrait
    
End Sub

Private Sub CargoCheques(Fecha As String)

Dim RsSuc As rdoResultset
Dim aSucCodigo As Long, aSucNombre As String        'Para Depositar
Dim aTotal As Currency
Dim aValor As Long
Dim bADepositar As Boolean

    On Error GoTo errPago
    Screen.MousePointer = 11
    vsConsulta.Rows = 1
    bGrabar.Enabled = False
    bCargarImpresion = True
    aSucCodigo = 0: aTotal = 0
    
    'Armo la Consulta de Cheques-------------------------------------------------------------------------
    Cons = "SELECT ChequeDiferido.*, SucursalDeBanco.*, BancoSSFF.*, IsNull(DocSucursal, 0) DocSucursal FROM ChequeDiferido INNER JOIN SucursalDeBanco ON CDiSucursal = SBaCodigo" _
            & " INNER JOIN BancoSSFF ON CDiBanco = BanCodigo " _
            & " LEFT OUTER JOIN Documento ON CDiDocumento = DocCodigo" _
            & " Where CDiMoneda = " & cMoneda.ItemData(cMoneda.ListIndex) _
            & " And CDiEliminado Is Null"
    
    '& ConsultaDeFecha("And", "CDiVencimiento", Fecha)
     Select Case cTipo.ListIndex
        Case 0: 'Todos
                Cons = Cons & " And ( (" & ConsultaDeFecha("", "CDiVencimiento", Fecha) & ") OR " & _
                                               "( CDiVencimiento Is Null " & ConsultaDeFecha("And", "CDiLibrado", Fecha) & ") )"
        Case 1: 'Al Día
                Cons = Cons & " And ( CDiVencimiento Is Null " & ConsultaDeFecha("And", "CDiLibrado", Fecha) & ")"
                
        Case 2: 'Diferidos
                Cons = Cons & ConsultaDeFecha("And", "CDiVencimiento", Fecha)
    End Select
    
    
    Select Case cCondicion.ItemData(cCondicion.ListIndex)
        Case 1:
                    Cons = Cons & " And CDiCobrado Is NULL"
                    bConsultar.Tag = "AD": bADepositar = True
                    vsConsulta.ColHidden(6) = True
        Case 2:
                    Cons = Cons & " And CDiCobrado Is Not NULL"
                    bConsultar.Tag = "DE":  bADepositar = False
                    vsConsulta.ColHidden(6) = False
    End Select
    
    Select Case cboDe.ListIndex
        Case 0: 'Comercio
                Cons = Cons & " And CDiTag Is Null "
                
        Case 1: 'Fleteros
                Cons = Cons & " And CDiTag = 1 "
                
    End Select
    
    Cons = Cons & " And CDiRebotado is Null " & _
                        " Order by CDiVencimiento" ', CDiSerie, CDiNumero"
    '-----------------------------------------------------------------------------------------------------------
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    
    Do While Not RsAux.EOF
        With vsConsulta
            .AddItem ""
            
            'ItemData (0) = Id_Cheque,  (1) = Id_Cliente, (5) = Id_BCoDepositar
            If Not IsNull(RsAux!CDiVencimiento) Then
                .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!CDiVencimiento, "dd/mm/yy")
            Else
                .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!CDiLibrado, "dd/mm/yy")
                .Cell(flexcpForeColor, .Rows - 1, 0) = Colores.Azul
            End If
            aValor = RsAux!CDiCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            aValor = RsAux!CDiCliente: .Cell(flexcpData, .Rows - 1, 1) = aValor
            
            aValor = 0
            If Not IsNull(RsAux("CDiTag")) Then aValor = RsAux("CDiTag")
            .Cell(flexcpData, .Rows - 1, 2) = aValor
        
            .Cell(flexcpText, .Rows - 1, 1) = Trim(RsAux!BanNombre) & " (" & Trim(RsAux!SBaNombre) & ")"
            .Cell(flexcpText, .Rows - 1, 2) = Trim(RsAux!CDiSerie) & " " & Trim(RsAux!CDiNumero)
        
        
            .Cell(flexcpText, .Rows - 1, 3) = Format(RsAux!CDiImporte, FormatoMonedaP)
            aTotal = aTotal + RsAux!CDiImporte
        
            .Cell(flexcpText, .Rows - 1, 4) = Format(RsAux!CDiLibrado, "dd/mm/yy")
        
            If bADepositar Then
                '.Cell(flexcpPicture, .Rows - 1, 0) = Image1.ListImages("check").ExtractIcon
                'Sucursal a Depositar SBaDeposito-----------------------------------------------------------
                If aSucCodigo <> RsAux!SBaDeposito Then
                    Cons = "Select * from  SucursalDeBanco, BancoSSFF" _
                            & " Where SBaCodigo = " & RsAux!SBaDeposito _
                            & " And SBaBanco = BanCodigo"
                    Set RsSuc = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    aSucCodigo = RsAux!SBaDeposito
                    aSucNombre = Trim(RsSuc!BanNombre) & " (" & Trim(RsSuc!SBaNombre) & ")"
                    RsSuc.Close
                End If
                '-------------------------------------------------------------------------------------------------
            Else
                'Sucursal Depositado CDiDepositado-----------------------------------------------------------
                If aSucCodigo <> RsAux!CDiDepositado Then
                    Cons = "Select * from  SucursalDeBanco, BancoSSFF" _
                            & " Where SBaCodigo = " & RsAux!CDiDepositado _
                            & " And SBaBanco = BanCodigo"
                    Set RsSuc = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
                    aSucCodigo = RsAux!CDiDepositado
                    aSucNombre = Trim(RsSuc!BanNombre) & " (" & Trim(RsSuc!SBaNombre) & ")"
                    RsSuc.Close
                End If
                '-------------------------------------------------------------------------------------------------
            End If
        
            .Cell(flexcpText, .Rows - 1, 5) = aSucNombre
            aValor = aSucCodigo: .Cell(flexcpData, .Rows - 1, 5) = aValor
        
            If Not bADepositar Then .Cell(flexcpText, .Rows - 1, 6) = Format(RsAux!CDiCobrado, "dd/mm/yy")
                
'            If RsAux("DocSucursal") <> paCodigoDeSucursal Then
'                .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = IIf(RsAux("DocSucursal") = 0, &HEEEEEE, &HF0FFFF)
'            End If
            
        
            RsAux.MoveNext
        End With
    Loop
    RsAux.Close
    
    lTotal.Caption = Format(aTotal, FormatoMonedaP)
    lQch.Tag = vsConsulta.Rows - vsConsulta.FixedRows
    lQch.Caption = Format(lQch.Tag, "(0)")
    
    If vsConsulta.Rows > 0 And bADepositar Then bGrabar.Enabled = True
    Screen.MousePointer = 0
    If vsConsulta.Rows = 1 Then MsgBox "No hay datos a desplegar para los filtros ingresados.", vbExclamation, "No hay datos"
    Exit Sub
    
errPago:
    clsGeneral.OcurrioError "Error al cargar los cheques diferidos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    If Me.Width >= 9120 Then
        
        picBotones.Left = 60
        picBotones.Top = Me.ScaleHeight - picBotones.Height - 50
        
        frmFiltro.Left = 60
        frmFiltro.Width = Me.ScaleWidth - (frmFiltro.Left * 2)
        vsConsulta.Left = frmFiltro.Left: vsConsulta.Top = frmFiltro.Top + frmFiltro.Height + 100
        vsConsulta.Width = frmFiltro.Width
        vsConsulta.Height = Me.Height - vsConsulta.Top - picBotones.Height - 450
        
        vsListado.Top = vsConsulta.Top: vsListado.Height = vsConsulta.Height
        vsListado.Left = vsConsulta.Left: vsListado.Width = vsConsulta.Width
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set clsGeneral = Nothing
    Set miConexion = Nothing
    GuardoSeteoForm Me
    End
End Sub

Private Sub Label1_Click()
    Foco tFecha
End Sub

Private Sub Label2_Click()
    Foco cMoneda
End Sub

Private Sub Label4_Click()
    Foco cCondicion
End Sub

Private Sub MnuMarcar_Click()
Dim mRow As Integer, mCol As Integer, mTXT As String

    With vsConsulta
        mCol = .Col: mRow = .Row
        mTXT = .Cell(flexcpText, mRow, mCol)
        
        Dim iJ As Integer
        For iJ = .FixedRows To .Rows - 1
            If Not (.Cell(flexcpText, iJ, mCol) = mTXT) Then
                If .Cell(flexcpBackColor, iJ, 0, , .Cols - 1) <> Colores.Inactivo Then
                    CambioIcono xRow:=CLng(iJ)
                End If
            Else
                If .Cell(flexcpBackColor, iJ, 0, , .Cols - 1) = Colores.Inactivo Then
                    CambioIcono xRow:=CLng(iJ)
                End If
            End If
        Next
        
    End With
        
End Sub

Private Sub MnuSeguimiento_Click()
    'ItemData (0) = Id_Cheque,  (1) = Id_Cliente, (5) = Id_BCoDepositar
    EjecutarApp prmPathApp & "SeguimientoCheques.exe", vsConsulta.Cell(flexcpData, vsConsulta.Row, 0)

End Sub

Private Sub MnuVerDeuda_Click()
    'ItemData (0) = Id_Cheque,  (1) = Id_Cliente, (5) = Id_BCoDepositar
    EjecutarApp prmPathApp & "Deuda en cheques", vsConsulta.Cell(flexcpData, vsConsulta.Row, 1)
End Sub

Private Sub MnuVerSucursal_Click()
    CambioBancoADepositar
End Sub

Private Sub tFecha_GotFocus()
    tFecha.SelStart = 0
    tFecha.SelLength = Len(tFecha.Text)
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        aTexto = ValidoPeriodoFechas(tFecha.Text)
        If aTexto = "" Then MsgBox "El período de fechas ingresado no es correcto.", vbExclamation, "ATENCIÓN": Exit Sub
        Foco cMoneda
        If IsDate(tFecha.Text) Then tFecha.Text = Format(tFecha.Text, "d/mm/yyyy")
    End If
    
End Sub

Private Sub CambioIcono(Optional xRow As Long = -1)
    
    On Error Resume Next
    If xRow = -1 Then xRow = vsConsulta.Row
    
    If vsConsulta.Rows = 1 Then Exit Sub
    If bConsultar.Tag <> "AD" Then Exit Sub
    With vsConsulta
        If .Cell(flexcpBackColor, xRow, 0, , .Cols - 1) <> Colores.Inactivo Then
            .Cell(flexcpBackColor, xRow, 0, , .Cols - 1) = Colores.Inactivo
            lTotal.Caption = Format(CCur(lTotal.Caption) - .Cell(flexcpValue, xRow, 3), FormatoMonedaP)
            
            lQch.Tag = Val(lQch.Tag) - 1
        Else
            lTotal.Caption = Format(CCur(lTotal.Caption) + .Cell(flexcpValue, xRow, 3), FormatoMonedaP)
            .Cell(flexcpBackColor, xRow, 0, , .Cols - 1) = .BackColor
            lQch.Tag = Val(lQch.Tag) + 1
        End If
        
        lQch.Caption = Format(lQch.Tag, "(0)")
    End With
    
End Sub

Private Sub CambioBancoADepositar()
    
    On Error Resume Next
    If vsConsulta.Rows = 1 Then Exit Sub
    
    Dim aCodigo As String
    Dim aCodSucursal As Long
    
    aCodigo = InputBox("Ingrese el codigo de banco y sucursal en donde se va a depositar el cheque (Formato 00-000).", "Sucursal a Depositar")
    
    If Trim(aCodigo) = "" Then Exit Sub
    If InStr(aCodigo, "-") = 0 Then MsgBox "El código ingresado no es correcto.", vbExclamation, "ATENCIÓN": Exit Sub
    
    'ItemData (0) = Id_Cheque,  (1) = Id_Cliente, (5) = Id_BCoDepositar
    On Error GoTo errCargar
    Dim aSucursal As Integer, aBanco As Integer
    aBanco = CInt(Mid(aCodigo, 1, InStr(aCodigo, "-") - 1))
    aSucursal = CInt(Mid(aCodigo, InStr(aCodigo, "-") + 1, Len(aCodigo)))
    
    Screen.MousePointer = 11
    Cons = "Select * from  SucursalDeBanco, BancoSSFF" _
            & " Where SBaCodigoS = " & aSucursal _
            & " And BanCodigoB = " & aBanco _
            & " And SBaBanco = BanCodigo"
    Set RsAux = cBase.OpenResultset(Cons, rdOpenForwardOnly, rdConcurValues)
    If Not RsAux.EOF Then
        aCodigo = Trim(RsAux!BanNombre) & " (" & Trim(RsAux!SBaNombre) & ")"
        aCodSucursal = RsAux!SBaCodigo
    Else
        aCodigo = ""
        Screen.MousePointer = 0
        MsgBox "No existe registro para el código ingresado.", vbExclamation, "ATENCIÓN"
    End If
    RsAux.Close
    
    If aCodigo <> "" Then
        vsConsulta.Cell(flexcpText, vsConsulta.Row, 5) = aCodigo
        vsConsulta.Cell(flexcpData, vsConsulta.Row, 5) = aCodSucursal
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Ocurrió un error al buscar el banco.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionGrabar()

    On Error GoTo errorBT
    FechaDelServidor
    Screen.MousePointer = 11
    
    Dim RsMov As rdoResultset
    Dim mFechaHora As String
    Dim mIDMov As Long, mCompra As Long, I As Integer
    Dim mTCDolar As Currency, mTCPesos As Currency
    
    mTCPesos = 1
    mTCDolar = TasadeCambio(paMonedaDolar, cMoneda.ItemData(cMoneda.ListIndex), UltimoDia(DateAdd("m", -1, gFechaServidor)))
    If cMoneda.ItemData(cMoneda.ListIndex) <> paMonedaPesos Then
        mTCPesos = TasadeCambio(cMoneda.ItemData(cMoneda.ListIndex), paMonedaPesos, UltimoDia(DateAdd("m", -1, gFechaServidor)))
    End If
    mFechaHora = Format(gFechaServidor, "dd/mm/yyyy") & " " & Format(gFechaServidor, "hh:mm:ss")
    
    
    Dim objGeneric As New clsDBFncs
    Dim rdoCZureo As rdoConnection
    If Not objGeneric.get_Connection(rdoCZureo, "ORG01", 10) Then
        MsgBox "Error al conectarse a la base de datos de Zureo.", vbExclamation, "Conexión Zureo"
        Screen.MousePointer = 0
        Exit Sub
    End If
        
    'cBase.BeginTrans    'COMIENZO TRANSACCION ------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    For I = 1 To vsConsulta.Rows - 1
        With vsConsulta
        If .Cell(flexcpBackColor, I, 0, , .Cols - 1) <> Colores.Inactivo Then
            
            'ItemData (0) = Id_Cheque,  (1) = Id_Cliente, (5) = Id_BCoDepositar
            Cons = "Update ChequeDiferido " _
                    & " Set CDiCobrado = '" & Format(gFechaServidor, sqlFormatoFH) & "', " _
                    & " CDiDepositado = " & .Cell(flexcpData, I, 5) _
                    & " Where CDiCodigo = " & .Cell(flexcpData, I, 0)
            cBase.Execute Cons
        End If
        End With
    Next
    
    Dim m_ReturnID As Long
    Dim objComp As New clsComprobantes
    
    Dim OBJ_COM As clsDComprobante, OBJ_CTA As clsDCuenta
    Dim colCuentas As New Collection
    
    Dim xCuentaS As Long, xCuentaE As Long
    Dim xCuentaS_M As Integer, xCuentaE_M As Integer

    'Grabo los Gastos Y Transferencias  --------------------------------------------------------------------------
    For I = LBound(dGastos) To UBound(dGastos)
        mIDMov = 0: mCompra = 0
        
        If dGastos(I).ImporteAlDia <> 0 Then      'Transferencia entre disponibilidades
                    
            Cons = "Select DisID, DisIDSubRubro, DisMoneda  from Disponibilidad " & _
                   " Where DisID IN (" & dGastos(I).IdDisponibilidadEntrada & "," & dGastos(I).IdDisponibilidadSalida & ")"
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            Do While Not RsAux.EOF
                If RsAux!DisID = dGastos(I).IdDisponibilidadEntrada Then
                    xCuentaE = RsAux!DisIDSubrubro
                    xCuentaE_M = RsAux!DisMoneda
                Else
                    xCuentaS = RsAux!DisIDSubrubro
                    xCuentaS_M = RsAux!DisMoneda
                End If
                RsAux.MoveNext
            Loop
            RsAux.Close

            Set OBJ_COM = New clsDComprobante
            With OBJ_COM
                .doAccion = 1 'IIf(xCta1 <> 0, 1, 0)
                .Ente = 0 'dGastos(I).IdProveedorGasto
                .Empresa = 1
                .Fecha = CDate(Format(mFechaHora, "dd/mm/yyyy"))
                .Tipo = 50  'Transferencias Zureo 'TipoDocumento.CompraEntradaCaja
                .Moneda = dGastos(I).idMoneda
                .ImporteTotal = dGastos(I).ImporteAlDia
                .TC = IIf(dGastos(I).idMoneda <> paMonedaPesos, mTCDolar, 1)
                .Memo = "Depósito de Ch.D. " & dGastos(I).SucursalNombre
                .UsuarioAlta = paCodigoDeUsuario
                .UsuarioAutoriza = paCodigoDeUsuario
            End With
        
            Set OBJ_CTA = New clsDCuenta
            With OBJ_CTA
                .VaAlDebe = 0
                .Cuenta = xCuentaS 'dGastos(I).IdSubrubroSalida
                .ImporteComp = dGastos(I).ImporteAlDia
                .ImporteCta = dGastos(I).ImporteAlDia * mTCPesos
                .MonedaCta = xCuentaS_M 'prmMonedaContabilidad
            End With
            colCuentas.Add OBJ_CTA
            Set OBJ_CTA = Nothing
            
            Set OBJ_CTA = New clsDCuenta
            With OBJ_CTA
                .VaAlDebe = 1
                .Cuenta = xCuentaE
                .ImporteComp = dGastos(I).ImporteAlDia
                .ImporteCta = dGastos(I).ImporteAlDia * mTCPesos
                .MonedaCta = xCuentaE_M 'prmMonedaContabilidad
            End With
            colCuentas.Add OBJ_CTA
            Set OBJ_CTA = Nothing
        
            Set OBJ_COM.Cuentas = colCuentas
            If objComp.fnc_PasarComprobante(rdoCZureo, OBJ_COM) Then m_ReturnID = objComp.prm_Comprobante
            Set OBJ_COM = Nothing
        
            'Inserto en la Tabla Movimiento-Disponibilidad--------------------------------------------------------
            'cons = "Select * from MovimientoDisponibilidad Where MDiID = " & mIDMov
            'Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            
            'RsMov.AddNew
            'RsMov!MDiFecha = Format(mFechaHora, "mm/dd/yyyy")
            'RsMov!MDiHora = Format(mFechaHora, "hh:mm:ss")
            'RsMov!MDiTipo = dGastos(I).IdTipoTransferencia
            'RsMov!MDiIdCompra = Null
            'RsMov!MDiComentario = "Depósito de Ch. " & dGastos(I).SucursalNombre
            'RsMov.Update: RsMov.Close
            '------------------------------------------------------------------------------------------------------------

            'Saco el Id de movimiento-------------------------------------------------------------------------------
            'cons = "Select Max(MDiID) from MovimientoDisponibilidad" & _
                      " Where MDiFecha = " & Format(mFechaHora, "'mm/dd/yyyy'") & _
                      " And MDiHora = " & Format(mFechaHora, "'hh:mm:ss'") & _
                      " And MDiTipo = " & dGastos(I).IdTipoTransferencia & _
                      " And MDiIdCompra is Null"
        
            'Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            'mIDMov = RsMov(0)
            'RsMov.Close
            '------------------------------------------------------------------------------------------------------------

            'Grabo en Tabla Movimiento-Disponibilidad-Renglon--------------------------------------------------
            'cons = "Select * from MovimientoDisponibilidadRenglon Where MDRIdMovimiento = " & mIDMov
            'Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            
            'RsMov.AddNew        'Salida
            'RsMov!MDRIdMovimiento = mIDMov
            'RsMov!MDRIdDisponibilidad = dGastos(I).IdDisponibilidadSalida
            'RsMov!MDRIdCheque = 0
            'RsMov!MDRImporteCompra = dGastos(I).ImporteAlDia
            'RsMov!MDRImportePesos = dGastos(I).ImporteAlDia * mTCPesos
            'RsMov!MDRHaber = dGastos(I).ImporteAlDia
            'RsMov.Update
            
            'RsMov.AddNew        'Entrada
            'RsMov!MDRIdMovimiento = mIDMov
            'RsMov!MDRIdDisponibilidad = dGastos(I).IdDisponibilidadEntrada
            'RsMov!MDRIdCheque = 0
            'RsMov!MDRImporteCompra = dGastos(I).ImporteAlDia
            'RsMov!MDRImportePesos = dGastos(I).ImporteAlDia * mTCPesos
            'RsMov!MDRDebe = dGastos(I).ImporteAlDia
            'RsMov.Update
            
            'RsMov.Close
        End If
                
        mIDMov = 0
        If dGastos(I).ImporteDiferido <> 0 Then     'Registro Gasto por Cheques Diferidos
        
            Cons = "Select DisID, DisIDSubRubro, DisMoneda  from Disponibilidad " & _
                   " Where DisID = " & dGastos(I).IdDisponibilidadEntrada
            Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
            If Not RsAux.EOF Then
                xCuentaE = RsAux!DisIDSubrubro
                xCuentaE_M = RsAux!DisMoneda
                RsAux.MoveNext
            End If
            RsAux.Close
                               
            Set OBJ_COM = New clsDComprobante
            With OBJ_COM
                .doAccion = 1 'IIf(xCta1 <> 0, 1, 0)
                .Ente = 0 'dGastos(I).IdProveedorGasto
                .Empresa = 1
                .Fecha = CDate(Format(mFechaHora, "dd/mm/yyyy"))
                .Tipo = TipoDocumento.CompraEntradaCaja
                .Moneda = dGastos(I).idMoneda
                .ImporteTotal = dGastos(I).ImporteDiferido
                .TC = IIf(dGastos(I).idMoneda <> paMonedaPesos, mTCDolar, 1)
                .Memo = "Depósito de Ch.D. " & dGastos(I).SucursalNombre
                .UsuarioAlta = paCodigoDeUsuario
                .UsuarioAutoriza = paCodigoDeUsuario
            End With

        
            Set OBJ_CTA = New clsDCuenta
            With OBJ_CTA
                .VaAlDebe = 0
                .Cuenta = dGastos(I).IdSubrubroSalida
                .ImporteComp = dGastos(I).ImporteDiferido
                .ImporteCta = dGastos(I).ImporteDiferido * mTCPesos
                .MonedaCta = paMonedaPesos 'prmMonedaContabilidad
            End With
            colCuentas.Add OBJ_CTA
            Set OBJ_CTA = Nothing

            Set OBJ_CTA = New clsDCuenta
            With OBJ_CTA
                .VaAlDebe = 1
                .Cuenta = xCuentaE 'dGastos(I).IdDisponibilidadEntrada
                .ImporteComp = dGastos(I).ImporteDiferido
                .ImporteCta = dGastos(I).ImporteDiferido * mTCPesos
                .MonedaCta = xCuentaE_M 'prmMonedaContabilidad
            End With
            colCuentas.Add OBJ_CTA
            Set OBJ_CTA = Nothing

            Set OBJ_CTA = Nothing
            Set OBJ_COM.Cuentas = colCuentas
            
            If objComp.fnc_PasarComprobante(rdoCZureo, OBJ_COM) Then m_ReturnID = objComp.prm_Comprobante
            
            Set OBJ_COM = Nothing
       
            'cons = "Select * from Compra Where ComCodigo = " & mCompra
            'Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            'RsMov.AddNew
            
            'RsMov!ComFecha = Format(mFechaHora, "mm/dd/yyyy")
            'RsMov!ComProveedor = dGastos(I).IdProveedorGasto
            'RsMov!ComTipoDocumento = TipoDocumento.CompraEntradaCaja
            
            'RsMov!ComMoneda = dGastos(I).idMoneda
            'RsMov!ComTC = mTCDolar
            'RsMov!ComImporte = dGastos(I).ImporteDiferido * -1
            'RsMov!ComSaldo = 0
            
            'RsMov!ComComentario = "Depósito de Ch.D. " & dGastos(I).SucursalNombre
            'RsMov!ComFModificacion = Format(gFechaServidor, "mm/dd/yyyy hh:mm:ss")
            'RsMov!ComUsuario = paCodigoDeUsuario
            'RsMov.Update: RsMov.Close
            '--------------------------------------------------------------------------------------------------------------------------------
    
            'cons = "Select Max(ComCodigo) from Compra" & _
                    " Where ComFecha = " & Format(mFechaHora, "'mm/dd/yyyy'") & _
                    " And ComTipoDocumento = " & TipoDocumento.CompraEntradaCaja & _
                    " And ComProveedor = " & dGastos(I).IdProveedorGasto & _
                    " And ComMoneda = " & dGastos(I).idMoneda
            'Set RsMov = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
            'mCompra = RsMov(0)
            'RsMov.Close
    
            'Tabla Gasto Subrubros  ----------------------------------------------------------------------
            'cons = "Select * from GastoSubrubro Where GSrIDCompra = " & mCompra
            'Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            'RsMov.AddNew
            'RsMov!GSrIDCompra = mCompra
            'RsMov!GSrIDSubrubro = dGastos(I).IdSubrubroSalida
            'RsMov!GSrImporte = dGastos(I).ImporteDiferido * -1
            'RsMov.Update: RsMov.Close

            'Inserto en la Tabla Movimiento-Disponibilidad--------------------------------------------------------
            'cons = "Select * from MovimientoDisponibilidad Where MDiID = " & mIDMov
            'Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            
            'RsMov.AddNew
            'RsMov!MDiFecha = Format(mFechaHora, "mm/dd/yyyy")
            'RsMov!MDiHora = Format(mFechaHora, "hh:mm:ss")
            'RsMov!MDiTipo = paMDPagoDeCompra
            'RsMov!MDiIdCompra = mCompra
            'RsMov!MDiComentario = "Depósito de Ch.D. " & dGastos(I).SucursalNombre
            'RsMov.Update: RsMov.Close
            '------------------------------------------------------------------------------------------------------------
            
            'Saco el Id de movimiento-------------------------------------------------------------------------------
            'cons = "Select Max(MDiID) from MovimientoDisponibilidad" & _
                      " Where MDiFecha = " & Format(mFechaHora, "'mm/dd/yyyy'") & _
                      " And MDiHora = " & Format(mFechaHora, "'hh:mm:ss'") & _
                      " And MDiTipo = " & paMDPagoDeCompra & _
                      " And MDiIdCompra = " & mCompra
            'Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            'mIDMov = RsMov(0)
            'RsMov.Close
            '------------------------------------------------------------------------------------------------------------
            
            'Grabo en Tabla Movimiento-Disponibilidad-Renglon--------------------------------------------------
            'cons = "Select * from MovimientoDisponibilidadRenglon Where MDRIdMovimiento = " & mIDMov
            'Set RsMov = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            
            'If RsMov.EOF Then RsMov.AddNew Else RsMov.Edit
           
            'RsMov!MDRIdMovimiento = mIDMov
            'RsMov!MDRIdDisponibilidad = dGastos(I).IdDisponibilidadEntrada
            'RsMov!MDRIdCheque = 0
            
            'RsMov!MDRImporteCompra = dGastos(I).ImporteDiferido
            'RsMov!MDRImportePesos = dGastos(I).ImporteDiferido * mTCPesos
            'RsMov!MDRDebe = dGastos(I).ImporteDiferido
            'RsMov.Update: RsMov.Close
        End If
    Next
    
    
    Set objComp = Nothing
    rdoCZureo.Close
    
    'cBase.CommitTrans    'Fin de la TRANSACCION------------------------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    
    
    bGrabar.Enabled = False
    vsConsulta.Rows = 1
    Screen.MousePointer = 0
    Exit Sub

errorBT:
    clsGeneral.OcurrioError "No se ha podido inicializar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0: Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "No se ha podido realizar la transacción. Reintente la operación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub InicializoGrilla()

    On Error Resume Next
    With vsConsulta
        .Cols = 1: .Rows = 1:
        .FormatString = "Vencimiento|Banco Emisor|<Nº Cheque|>Importe|Librado|Banco a Depositar|Depositado"
            
        .WordWrap = False
        .ColWidth(0) = 1100: .ColWidth(1) = 2260: .ColWidth(3) = 1200: .ColWidth(4) = 750: .ColWidth(5) = 2200
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

End Sub

Private Sub vsConsulta_DblClick()

    If vsConsulta.Rows = 1 Then Exit Sub
    CambioIcono
    
End Sub

Private Sub vsConsulta_KeyDown(KeyCode As Integer, Shift As Integer)

    If vsConsulta.Rows = 1 Then Exit Sub
    On Error GoTo errLista
    
    If KeyCode = vbKeyDelete And bConsultar.Tag = "DE" Then
        With vsConsulta
            'Elimino depósito del cheque---------------------------------------------------
            If MsgBox("Confirma eliminar el depósito del cheque: " & .Cell(flexcpText, .Row, 2) & " del " & .Cell(flexcpText, .Row, 1) & vbCrLf & vbCrLf & _
                            "Si presiona SI va a acceder al Seguimiento de Cheques y allí debe presionar el botón para 'Des-Depositarlo'.", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Depósito") = vbNo Then Exit Sub
            
            Call MnuSeguimiento_Click
            
        End With
        Screen.MousePointer = 0
    End If      '--------------------------------------------------------------------------------------------
    Exit Sub
    
errLista:
    clsGeneral.OcurrioError "Error al ejecutar la operación.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub vsConsulta_KeyPress(KeyAscii As Integer)

    If vsConsulta.Rows = 1 Then Exit Sub
    On Error GoTo errLista
    Select Case KeyAscii
        Case vbKeySpace: CambioIcono
        Case vbKeyReturn: If bGrabar.Enabled Then bGrabar.SetFocus
    End Select
    
    Exit Sub
errLista:
    Screen.MousePointer = 0
    clsGeneral.OcurrioError "Ocurrió un error al procesar la información.", Err.Description
End Sub

Private Sub vsConsulta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    On Error Resume Next
   
    If vsConsulta.Rows = 1 Then Exit Sub
    If bConsultar.Tag = "DE" Then MnuVerSucursal.Enabled = False Else: MnuVerSucursal.Enabled = True
    If Button = vbRightButton Then
        vsConsulta.Select vsConsulta.MouseRow, vsConsulta.MouseCol
        PopupMenu MnuBDerecho
    End If

End Sub

Private Sub AccionImprimir(Optional Imprimir As Boolean = False)
    
    On Error GoTo errImprimir
    'Inicializo Objeto de impresión.---------------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    
    If bCargarImpresion Then
        If vsConsulta.Rows = 1 Then Screen.MousePointer = 0: Exit Sub
        With vsListado
            .StartDoc
            If .Error Then
                MsgBox "No se pudo iniciar el documento de impresión." & Chr(vbKeyReturn) & _
                            Err.Number & "- " & Err.Description & _
                            "VsPrinterError: " & .Error, vbInformation, "ATENCIÓN"
                Screen.MousePointer = 0: Exit Sub
            End If
        End With        '----------------------------------------------------------------------------------------------------------------------------------------------
        
        aTexto = Trim(tFecha.Text) & ", " & Trim(cMoneda.Text) & ", " & Trim(cCondicion.Text)
        EncabezadoListado vsListado, "Depósito de Cheques - " & aTexto, False
        vsListado.FileName = "Deposito de cheques"
        
        vsConsulta.ExtendLastCol = False: vsListado.RenderControl = vsConsulta.hwnd: vsConsulta.ExtendLastCol = True
        vsListado.Paragraph = " ": vsListado.Paragraph = Trim(lTTotal.Caption) & " " & Trim(lTotal.Caption)
        
        vsListado.EndDoc
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

Private Sub AccionLimpiar()
    
    tFecha.Text = ""
    cMoneda.Text = ""
    cCondicion.Text = ""
    
    lTotal.Caption = "0.00"
    lQch.Caption = "": lQch.Tag = 0
    
End Sub
    
Private Sub AccionConfigurar()
    frmSetup.pControl = vsListado
    frmSetup.Show vbModal, Me
End Sub


Private Sub vsListado_EndDoc()
    EnumeroPiedePagina vsListado
End Sub

