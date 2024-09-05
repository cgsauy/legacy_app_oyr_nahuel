VERSION 5.00
Object = "{B443E3A5-0B4D-4B43-B11D-47B68DC130D7}#1.5#0"; "orArticulo.ocx"
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Begin VB.Form frmStockLocal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Stock en Local"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStockLocal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9840
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4155
      Width           =   9840
      Begin VB.CommandButton cbSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7440
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGrilla 
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3413
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   0
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
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
      SubtotalPosition=   0
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
   Begin VB.PictureBox picFiltro 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   9840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   9840
      Begin prjFindArticulo.orArticulo tArticulo 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         FindNombreEnUso =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese el código, o parte del nombre del artículo a buscar y presione enter"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   6495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Artículo:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmStockLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oFnc As New clsFunciones
Dim oCGSA As New clsorCGSA

Private Sub ps_LoadStock()
On Error GoTo errLS
Dim rsA As rdoResultset
Dim sQy As String
    Screen.MousePointer = 11
    sQy = "SELECT StLCantidad, rTRIM(EsMAbreviacion) as EsMAbreviacion" _
        & " FROM StockLocal INNER JOIN EstadoMercaderia ON StlEstado = EsMCodigo" _
        & " WHERE StlArticulo = " & tArticulo.prm_ArtID _
        & " AND StlLocal = " & paCodigoDeSucursal _
        & " And StlCantidad <> 0 And StlEstado <> 0 ORDER BY EsMAbreviacion"
    Set rsA = rdBase.OpenResultset(sQy, rdOpenDynamic, rdConcurValues)
    Do While Not rsA.EOF
        With vsGrilla
            .AddItem rsA("EsMAbreviacion")
            .Cell(flexcpText, .Rows - 1, 1) = rsA("StLCantidad")
        End With
        rsA.MoveNext
    Loop
    rsA.Close
    If vsGrilla.Rows > 1 Then
        vsGrilla.Subtotal flexSTSum, -1, 1, "#,##0", &HC0C0C0, , True, "Total"
        vsGrilla.SetFocus
    End If
    Screen.MousePointer = 0
    Exit Sub
errLS:
    Screen.MousePointer = 0
    oCGSA.OcurrioError "Error al consultar.", Err.Description, "Consulta"
End Sub

Private Sub ps_InitGrid()
    
    With vsGrilla
        .FixedRows = 1
        .Rows = .FixedRows
        
        .FixedCols = 0
        .Cols = 1
        .GridLinesFixed = flexGridFlatHorz
        .GridLines = flexGridFlatHorz
        
        .FormatString = "<Estado|>Cantidad"
                
        .ColWidth(0) = 2500
        .ColWidth(1) = 1500
        
        .ExtendLastCol = False
        .WordWrap = False
               
        .BackColorBkg = .BackColor
        '.GridLinesFixed = flexGridInsetHorz
        .GridLinesFixed = flexGridFlat
        
        .BackColorFixed = vbApplicationWorkspace
        .ForeColorFixed = vbWindowBackground
        .GridColorFixed = &HFAF0EB
        
        .BackColorAlternate = &HF0F0F0
        
        .HighLight = flexHighlightAlways ' flexHighlightWithFocus
        .FocusRect = flexFocusNone
        .BackColorSel = &H800000       'vbInfoBackground
        .ForeColorSel = vbWhite '1 'vbHighlight
        .SheetBorder = .BackColor
        
        '.GridColorFixed = vbButtonShadow
        '.ForeColorFixed = vbHighlightText
'        .BorderStyle = flexBorderNone
        .MergeCells = flexMergeSpill
        
        .SelectionMode = 1
        
        .RowHeight(0) = 320
        .RowHeightMin = 250
        
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
    End With

End Sub

Private Sub cbSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo errGP
    oFnc.GetPositionForm Me
    ps_InitGrid
    With tArticulo
        Set .Connect = rdBase
        .KeyQuerySP = "stktotal"
        .DisplayCodigoArticulo = True
    End With
    Screen.MousePointer = 0
    Exit Sub
errGP:
    oCGSA.OcurrioError "Error al iniciar el formulario.", Err.Description, "Stock en local"
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        vsGrilla.Move 120, picFiltro.Top + picFiltro.Height, Me.ScaleWidth - 240, Me.ScaleHeight - picFiltro.Top + picFiltro.Height
        cbSalir.Left = picBottom.ScaleWidth - cbSalir.Width - 120
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oFnc.SetPositionForm Me
    Set oFnc = Nothing
    Set oCGSA = Nothing
    rdBase.Close
End Sub

Private Sub tArticulo_Change()
    If tArticulo.prm_ArtID > 0 Then
        tArticulo.Tag = ""
        vsGrilla.Rows = 1
    End If
End Sub

Private Sub tArticulo_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrAP
    
    If KeyAscii = vbKeyReturn Then
        If tArticulo.Tag <> "" Then Exit Sub
        vsGrilla.Rows = 1
        
        If tArticulo.prm_ArtID > 0 Then ps_LoadStock
        tArticulo.SelectAll
        
    End If
    Exit Sub
ErrAP:
    oCGSA.OcurrioError "Error al buscar el artículo.", Err.Description
    Screen.MousePointer = 0
End Sub

