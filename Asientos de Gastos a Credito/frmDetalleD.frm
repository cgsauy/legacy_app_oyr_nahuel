VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{191D08B9-4E92-4372-BF17-417911F14390}#1.5#0"; "orGridPreview.ocx"
Begin VB.Form frmDetalleD 
   Caption         =   "Detalle de Gastos"
   ClientHeight    =   5220
   ClientLeft      =   1125
   ClientTop       =   3480
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
   Icon            =   "frmDetalleD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   12105
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picMarco 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   3015
      TabIndex        =   2
      Top             =   4800
      Width           =   3075
      Begin VB.CommandButton bPreview 
         Height          =   310
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Vista Previa"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir."
         Top             =   0
         Width           =   310
      End
   End
   Begin orGridPreview.GridPreview cPrint 
      Left            =   6120
      Top             =   60
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGastos 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4683
      _ConvInfo       =   1
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
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
   Begin VB.Label lCaption 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   525
   End
End
Attribute VB_Name = "frmDetalleD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmIDAcreedor As Long
Public prmFecha1 As String
Public prmFecha2 As String
Public prmIN_TDocs As String
Public prmAcreedorTXT As String

Dim rsSQL As rdoResultset

Private Sub bCancelar_Click()
    Unload Me
End Sub

Private Sub bImprimir_Click()
    
    Screen.MousePointer = 11
    With cPrint
        .AddGrid vsGastos.hwnd
        .GoPrint
    End With
    Screen.MousePointer = 0
    
End Sub

Private Sub bPreview_Click()
    
    Screen.MousePointer = 11
    With cPrint
        vsGastos.ExtendLastCol = False
        .AddGrid vsGastos.hwnd
        .ShowPreview
        vsGastos.ExtendLastCol = True
    End With
    Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
    
    Screen.MousePointer = 11
    
    InicializoForm
    InicializoGrillas
    AccionConsultar
    
    Screen.MousePointer = 0
    
End Sub

Private Sub AccionConsultar()
On Error GoTo errSQL
Dim iX As Byte, mSQL As String

    mSQL = ""
    For iX = 1 To 2
    
        If iX = 1 Then
            mSQL = mSQL & " Select RubNombre, Sum(Abs(GSrImporte) * ComTC), Sum(Abs(GSrImporte) ) "
            
        Else
            mSQL = mSQL & " Union All "
            mSQL = mSQL & "Select RubNombre, Sum(Abs(GSrImporte)), Sum(Abs(GSrImporte)) "
        End If
                
        mSQL = mSQL & _
        " From Compra, GastoSubRubro, SubRubro, Rubro " & _
        " Where ComFecha Between '" & Format(prmFecha1, "mm/dd/yyyy") & "' AND '" & Format(prmFecha2, "mm/dd/yyyy 23:59") & "'" & _
        " And ComCodigo = GSrIDCompra And GSrIDSubRubro = SRuID" & _
        " And SRuRubro = RubID " & _
        " And ComTipoDocumento In (" & prmIN_TDocs & ")" & _
        " And ComDCDe IS Null " & _
        " And GSrIDSubrubro " & IIf(prmIDAcreedor = 0, "NOT", "") & " IN (" & paSubrubroDivisa & ", " & paSubrubroDifCambioG & ", " & paSubrubroDifCambio & ", " & paSubrubroDifCostoImp & ")" & _
        " And ComMoneda " & IIf(iX = 1, "<>", "=") & paMonedaPesos

        If prmIDAcreedor <> 0 Then mSQL = mSQL & " And ComProveedor = " & prmIDAcreedor
        mSQL = mSQL & " Group by RubNombre"
    Next
    mSQL = mSQL & " Order by RubNombre"
    
    Set rsSQL = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
        
    Do While Not rsSQL.EOF
        With vsGastos               '<Rubro|>Total |>Total ME|
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Trim(rsSQL("RubNombre").Value)
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsSQL(1), "#,###.00")
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsSQL(2), "#,###.00")
            
            If .Cell(flexcpText, .Rows - 1, 1) = .Cell(flexcpText, .Rows - 1, 2) Then .Cell(flexcpText, .Rows - 1, 2) = ""
        End With
          
        rsSQL.MoveNext
    Loop
    rsSQL.Close
    
    'Si el Acreeedores es 0 le sumo el IVA de las compras   --------------------------------------------------------------------------------------
    If prmIDAcreedor = 0 Then
        Dim mTIVA As Currency, mTCOFIS As Currency
        mTIVA = 0: mTCOFIS = 0
        mSQL = ""
        For iX = 1 To 2
            If iX = 1 Then
                mSQL = mSQL & " Select Sum(Abs(ComIVA)* ComTC) , Sum(Abs(ComCofis) * ComTC)"
                
            Else
                mSQL = mSQL & " Union All "
                mSQL = mSQL & "Select Sum(Abs(ComIVA)), Sum(Abs(ComCofis)) "
            End If
                    
            mSQL = mSQL & _
                " From Compra " & _
                " Where ComFecha Between '" & Format(prmFecha1, "mm/dd/yyyy") & "' AND '" & Format(prmFecha2, "mm/dd/yyyy 23:59") & "'" & _
                " And ComTipoDocumento In (" & prmIN_TDocs & ")" & _
                " And ComMoneda " & IIf(iX = 1, "<>", "=") & paMonedaPesos & _
                " And ComCodigo NOT IN (" & _
                     " Select GSrIDCompra From GastoSubRubro " & _
                     " Where GSrIDSubrubro IN (" & paSubrubroDivisa & ", " & paSubrubroDifCambioG & ", " & paSubrubroDifCambio & ", " & paSubrubroDifCostoImp & ") )"
        Next
        
        Set rsSQL = cBase.OpenResultset(mSQL, rdOpenDynamic, rdConcurValues)
        Do While Not rsSQL.EOF
            If Not IsNull(rsSQL(0)) Then mTIVA = mTIVA + rsSQL(0)
            If Not IsNull(rsSQL(1)) Then mTCOFIS = mTCOFIS + rsSQL(1)
            rsSQL.MoveNext
        Loop
        rsSQL.Close
        
        With vsGastos
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Trim("IVA")
            .Cell(flexcpText, .Rows - 1, 1) = Format(mTIVA, "#,###.00")
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Trim("Cofis")
            .Cell(flexcpText, .Rows - 1, 1) = Format(mTCOFIS, "#,###.00")
        End With
    End If
    
    With vsGastos
        If .Rows > 1 Then
            '.Cell(flexcpBackColor, 1, 7, .Rows - 1, 8) = RGB(230, 230, 250)
            .Subtotal flexSTSum, -1, 1, , RGB(123, 104, 238), Colores.Blanco, , "Total Detalle"
            .Subtotal flexSTSum, -1, 2
        End If
    End With
    
    Exit Sub

errSQL:
    clsGeneral.OcurrioError "Error al realizar la consulta de datos.", Err.Description
End Sub

Private Sub InicializoGrillas()

    On Error Resume Next
    With vsGastos
        
        .MultiTotals = True
        
        .FixedCols = 0
        .GridLines = flexGridFlatHorz
        .SubtotalPosition = flexSTBelow
        .ExtendLastCol = True
        .BorderStyle = flexBorderNone
        .Appearance = flexFlat
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<Rubro|>Total $|>Total M/E|"
            
        .WordWrap = False
        .ColWidth(0) = 2000: .ColWidth(1) = 1400: .ColWidth(2) = 1400
        
        .MergeCells = flexMergeSpill
    End With
      
      
    'With frmListado.Image
    '    bImprimir.Picture = .ListImages("print").ExtractIcon
    '    bCancelar.Picture = .ListImages("salir").ExtractIcon
    '    bPreview.Picture = .ListImages("vista1").ExtractIcon
        bImprimir.Picture = frmListado.bImprimir.Picture
        bCancelar.Picture = frmListado.bCancelar.Picture
        bPreview.Picture = frmListado.chVista.Picture
    'End With
      
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With picMarco
        .Top = Me.ScaleHeight - .Height '- 50
        .Left = Me.ScaleLeft
        .BorderStyle = 0
    End With
    
    With vsGastos
        .Top = lCaption.Top + lCaption.Height + 100
        .Left = Me.ScaleLeft + 30
        .Width = Me.ScaleWidth - 60
        .Height = picMarco.Top - .Top - 80
    End With
    
End Sub

Private Sub InicializoForm()
On Error GoTo errIni

    lCaption.Caption = "Detalle de Acreedores: " & prmAcreedorTXT
    
    With cPrint
        .Orientation = opPortrait
        .Caption = Me.Caption
        .Header = lCaption.Caption
        .PageBorder = opTopBottom
    End With
    
    Exit Sub

errIni:
    clsGeneral.OcurrioError "Error cargar datos de la disponibilidad.", Err.Description
End Sub
