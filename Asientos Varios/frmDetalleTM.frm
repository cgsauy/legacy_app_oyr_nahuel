VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Object = "{191D08B9-4E92-4372-BF17-417911F14390}#1.5#0"; "orGridPreview.ocx"
Begin VB.Form frmDetalleTM 
   Caption         =   "Asientos Varios - Detalle de Movimientos"
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
   Icon            =   "frmDetalleTM.frx":0000
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
Attribute VB_Name = "frmDetalleTM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmTipoMovimiento As Long
Public prmIDRubro As Long

Public prmFecha1 As String
Public prmFecha2 As String

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
        .AddGrid vsGastos.hwnd
        .ShowPreview
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

    'Detalle de Movimientos para un MDiTipo = XX y MDiFecha = xxx
    cons = "Select MDiID, MDIFecha, DisNombre, SRuCodigo, SRuNombre, DH = 'Haber', Importe = MDrImportePesos, IOriginal = MDRHaber, MDiComentario" & _
                " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, Disponibilidad, SubRubro" & _
                " Where MDiFecha Between '" & Format(prmFecha1, "mm/dd/yyyy") & "' AND '" & Format(prmFecha2, "mm/dd/yyyy") & "'" & _
                " And MDiTipo = " & prmTipoMovimiento & _
                " And MDiIDCompra Is Null" & _
                " And MDRIDDisponibilidad = DisID  And DisIDSubrubro = SRuID" & _
                " And MDRHaber is Not Null And MDiID = MDRIDMovimiento"
    If prmIDRubro <> 0 Then cons = cons & " And SRuID = " & prmIDRubro

    cons = cons & " Union All"
    
    cons = cons & _
            " Select MDiID, MDIFecha, DisNombre, SRuCodigo, SRuNombre, DH = 'Debe', Importe = MDrImportePesos, IOriginal = MDRDebe, MDiComentario" & _
            " From MovimientoDisponibilidadRenglon, MovimientoDisponibilidad, Disponibilidad, SubRubro" & _
            " Where MDiFecha Between '" & Format(prmFecha1, "mm/dd/yyyy") & "' AND '" & Format(prmFecha2, "mm/dd/yyyy") & "'" & _
            " And MDiID = MDRIDMovimiento" & _
            " And MDiTipo = " & prmTipoMovimiento & _
            " And MDRIDDisponibilidad = DisID  And DisIDSubrubro = SRuID" & _
            " And MDRDebe is Not Null And MDiIDCompra Is Null"
    If prmIDRubro <> 0 Then cons = cons & " And SRuID = " & prmIDRubro
    
    cons = cons & " Order by MDiID"
    
    Set rsSQL = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    Dim mImportePesos As Currency, mImporteD As Currency, mTasaCambio As Currency
        
    Do While Not rsSQL.EOF
        With vsGastos
            'Id_Movs|<Fecha|<Disponibilidad|<SubRubro|>Importe (D)|>TC|>Debe $|>Haber $|<Memo"
            .AddItem Format(rsSQL!MDiID, "###,###")
            
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsSQL!MDiFecha, "dd/mm/yy") '& " " & Format(rsSQL!MDiHora, "hh:mm")
            
            .Cell(flexcpText, .Rows - 1, 2) = Trim(rsSQL!DisNombre)
            .Cell(flexcpText, .Rows - 1, 3) = Format(rsSQL!SRuCodigo, "000000000") & " " & Trim(rsSQL!SRuNombre)
            
            mImporteD = Format(Abs(rsSQL!IOriginal), "#,##0.00")
            mImportePesos = Format(Abs(rsSQL!Importe), "#,##0.00")
            
            .Cell(flexcpText, .Rows - 1, 4) = Format(mImporteD, "#,##0.00")
            
            mTasaCambio = 1
            If mImporteD <> mImportePesos Then
                mTasaCambio = mImportePesos / mImporteD
                'mTasaCambio = Redondeo(mTasaCambio, "0.05")
            End If
            .Cell(flexcpText, .Rows - 1, 5) = Format(mTasaCambio, "#,##0.000")
            
            
            If LCase(Trim(rsSQL!DH)) = "debe" Then
                .Cell(flexcpText, .Rows - 1, 6) = Format(mImportePesos, "#,##0.00")
            Else
                .Cell(flexcpText, .Rows - 1, 7) = Format(mImportePesos, "#,##0.00")
            End If
            If Not IsNull(rsSQL!MDiComentario) Then .Cell(flexcpText, .Rows - 1, 8) = Trim(rsSQL!MDiComentario)
                        
        End With
        rsSQL.MoveNext
    Loop
    rsSQL.Close
    
    
    With vsGastos
        If .Rows > 1 Then
            .Cell(flexcpBackColor, 1, 6, .Rows - 1, 7) = RGB(230, 230, 250)
            '.Subtotal flexSTSum, -1, 6, ,  RGB(123, 104, 238), Colores.Blanco, , "Total Detalle"
            '.Subtotal flexSTSum, -1, 7
            
             .ColSort(2) = flexSortGenericAscending
            .ColSort(3) = flexSortGenericAscending
            .ColSort(4) = flexSortNone
            .ColSort(5) = flexSortGenericAscending
        
            .Select 1, 5
            .Sort = flexSortUseColSort
        
            .OutlineBar = flexOutlineBarSimple
            .OutlineCol = 0
            .MultiTotals = True
            
            .Subtotal flexSTSum, 5, 6, , RGB(123, 104, 238), Colores.Blanco, , , 5, True
            .Subtotal flexSTSum, 5, 4, , , , , , 5, True
            .Subtotal flexSTSum, 5, 7, , , , , , 5, True
            
            
            CollapseAll
            
            .Subtotal flexSTSum, -1, 6, , RGB(123, 104, 238), Colores.Blanco, , "Total Detalle"
            .Subtotal flexSTSum, -1, 4, , , , , , 5, True
            .Subtotal flexSTSum, -1, 7
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
        
        .Cols = 1: .Rows = 1
        .FormatString = ">Id_Movs|<Fecha|<Disponibilidad|<SubRubro|>Importe (D)|>TC|>Debe $|>Haber $|<Memo"
            
        .WordWrap = False
        .ColWidth(0) = 800: .ColWidth(1) = 800
        .ColWidth(2) = 1900: .ColWidth(3) = 1900
        
        .ColWidth(4) = 1200: .ColWidth(5) = 700: .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        .ColWidth(8) = 4000
        
        .MergeCells = flexMergeSpill
    
    End With
      
      
    With frmListado.img1
        bImprimir.Picture = .ListImages("print").ExtractIcon
        bCancelar.Picture = .ListImages("salir").ExtractIcon
        bPreview.Picture = .ListImages("vista1").ExtractIcon
    End With
      
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

Private Sub vsGastos_DblClick()
On Error GoTo errFnc
    
    If vsGastos.Rows = 1 Then Exit Sub
    If vsGastos.Row = vsGastos.Rows - 1 Then Exit Sub
    
    If vsGastos.Cell(flexcpValue, vsGastos.Row, 0) <> 0 Then
        EjecutarApp prmPathApp & "Transferencias de Disponibilidades.exe", vsGastos.Cell(flexcpValue, vsGastos.Row, 0)
    End If
    
    
errFnc:
End Sub

Private Sub InicializoForm()
On Error GoTo errIni

Dim mTexto As String

    mTexto = ""
    cons = "Select * From TipoMovDisponibilidad Where TMDCodigo = " & prmTipoMovimiento
    Set rsSQL = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsSQL.EOF Then
        mTexto = Trim(rsSQL!TMDNombre)
    End If
    rsSQL.Close
    
    lCaption.Caption = "Detalle de Movimientos de Disponibilidades del Tipo: " & mTexto
    
    With cPrint
        .Orientation = opLandscape
        .Caption = Me.Caption
        .Header = lCaption.Caption
        .PageBorder = opTopBottom
    End With
    
    Exit Sub

errIni:
    clsGeneral.OcurrioError "Error cargar datos de la disponibilidad.", Err.Description
End Sub

Private Function CollapseAll()
'On Error Resume Next

    With vsGastos
        For I = .FixedRows To .Rows - .FixedRows
            If .IsSubtotal(I) Then
                .IsCollapsed(I) = flexOutlineCollapsed
            End If
        Next
    End With
    
End Function
