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
      TabIndex        =   1
      Top             =   4800
      Width           =   3075
      Begin VB.CommandButton bPreview 
         Height          =   310
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Vista Previa"
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bCancelar 
         Height          =   310
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Salir."
         Top             =   0
         Width           =   310
      End
      Begin VB.CommandButton bImprimir 
         Height          =   310
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   2
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
      Left            =   60
      TabIndex        =   0
      Top             =   60
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
End
Attribute VB_Name = "frmDetalleD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public prmDisponibilidadID As Long

Public prmFecha1 As String
Public prmFecha2 As String
Public prmAlHaber As Boolean

Public prmListarDiferidos As Boolean

Dim prmDisponibilidadN As String
Dim prmDisponibilidadM As Long
Dim prmDisponibilidadSRCh As Long

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
    
    zfn_InicializoControles
    
    Screen.MousePointer = 0
    
End Sub

Public Function fnc_AccionConsultar() As Boolean
On Error GoTo errSQL

    fnc_TituloDisponibilidad
    
    'SQL de Gatos para un Disponibilidad X
    cons = "Select MDiIdCompra, MDiFecha, MDiHora, RubCodigo, RubNombre, SRuCodigo, SRuNombre, " & _
                            " GSrImporte, MDRImportePesos, MDRImporteCompra, MDRDebe, MDRHaber, " & _
                            " ComCofis, ComIVA " & _
                " From MovimientoDisponibilidad, MovimientoDisponibilidadRenglon, Compra, GastoSubRubro, SubRubro, Rubro " & _
                " Where MDiID = MDRIDMovimiento" & _
                " And MDiIDCompra = ComCodigo And ComCodigo = GSrIDCompra" & _
                " And GSrIDSubRubro = SRuID  And SRuRubro = RubID" & _
                " And MDiFecha Between '" & Format(prmFecha1, "mm/dd/yyyy") & "' AND '" & Format(prmFecha2, "mm/dd/yyyy") & "'" & _
                " And MDRIdDisponibilidad = " & prmDisponibilidadID & _
                " And MDiIDCompra Is Not Null"
    
    If prmAlHaber Then
        cons = cons & " And MDRDebe Is Null "    '1- Egresos
    Else
        cons = cons & " And MDRHaber Is Null "   '0- Ingresos
    End If
    
    If prmDisponibilidadSRCh <> 0 Then
        If prmListarDiferidos Then
            cons = cons & " And MDRIDCheque IN "
        Else
            cons = cons & " And MDRIDCheque NOT IN "
        End If
        
        cons = cons & " ( Select CheID from Cheque " & _
                                " Where CheVencimiento Is Not Null" & _
                                " And CheLibrado Between '" & Format(prmFecha1, "mm/dd/yyyy") & "' AND '" & Format(prmFecha2, "mm/dd/yyyy") & "' )"
    End If
    
    cons = cons & " Order by MDiFecha, MDiHora, MDiID"
   
    
    Set rsSQL = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    Dim mTC_APesos As Currency, mTC_ADisp As Currency
    Dim mIDCompra As Long, mImporteD As Currency
    
    Dim TOT_Importe As Currency, TOT_Cofis As Currency, TOT_IVA As Currency, TOT_DH As Currency, TOT_ME As Currency
        
    Do While Not rsSQL.EOF
        With vsGastos
            '<Id Gasto|<Fecha|<Rubro|<SubRubro|>Importe $|>Cofis|>IVA|>Debe/Haber|>TC|
            .AddItem ""
            
            If mIDCompra <> rsSQL!MDiIdCompra Then
                .Cell(flexcpText, .Rows - 1, 0) = Format(rsSQL!MDiIdCompra, "###,###")
                .Cell(flexcpText, .Rows - 1, 1) = Format(rsSQL!MDiFecha, "dd/mm/yy") '& " " & Format(rsSQL!MDiHora, "hh:mm")
            End If
            
            .Cell(flexcpText, .Rows - 1, 2) = Format(rsSQL!RubCodigo, "000000000") & " " & Trim(rsSQL!RubNombre)
            .Cell(flexcpText, .Rows - 1, 3) = Format(rsSQL!SRuCodigo, "000000000") & " " & Trim(rsSQL!SRuNombre)
            
            If Not IsNull(rsSQL!MDRDebe) Then
                mImporteD = rsSQL!MDRDebe
            Else
                mImporteD = rsSQL!MDRHaber
            End If
            mTC_ADisp = 1
            mTC_APesos = rsSQL!MDRImportePesos / rsSQL!MDRImporteCompra
            mTC_ADisp = rsSQL!MDRImportePesos / mImporteD
            
            .Cell(flexcpText, .Rows - 1, 4) = Format(rsSQL!GSrImporte * mTC_APesos, "#,##0.00")
            
            If mIDCompra <> rsSQL!MDiIdCompra Then
                If Not IsNull(rsSQL!ComCofis) Then .Cell(flexcpText, .Rows - 1, 5) = Format(rsSQL!ComCofis * mTC_APesos, "#,##0.00")
                If Not IsNull(rsSQL!ComIVA) Then .Cell(flexcpText, .Rows - 1, 6) = Format(rsSQL!ComIVA * mTC_APesos, "#,##0.00")
                
                .Cell(flexcpText, .Rows - 1, 8) = Format(mImporteD, "#,##0.00")
                If mTC_ADisp <> 1 Then .Cell(flexcpText, .Rows - 1, 9) = Format(mTC_ADisp, "#,##0.000")
            
                .Cell(flexcpText, .Rows - 1, 7) = Format(rsSQL!MDRImportePesos, "#,##0.00")
            End If
            
            'mImporteD = .Cell(flexcpValue, .Rows - 1, 4) + .Cell(flexcpValue, .Rows - 1, 5) + .Cell(flexcpValue, .Rows - 1, 6)
            '.Cell(flexcpText, .Rows - 1, 7) = Format(mImporteD, "#,##0.00")
            
            TOT_Importe = TOT_Importe + .Cell(flexcpValue, .Rows - 1, 4)
            TOT_Cofis = TOT_Cofis + .Cell(flexcpValue, .Rows - 1, 5)
            TOT_IVA = TOT_IVA + .Cell(flexcpValue, .Rows - 1, 6)
            TOT_DH = TOT_DH + .Cell(flexcpValue, .Rows - 1, 7)
            TOT_ME = TOT_ME + .Cell(flexcpValue, .Rows - 1, 8)
            mIDCompra = rsSQL!MDiIdCompra
            
        End With
        rsSQL.MoveNext
    Loop
    rsSQL.Close
    
    
    With vsGastos
        If .Rows > 1 Then
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 4) = Format(TOT_Importe, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 5) = Format(TOT_Cofis, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 6) = Format(TOT_IVA, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 7) = Format(TOT_DH, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 8) = Format(TOT_ME, "#,##0.00")
            '.Cell(flexcpBackColor, 1, 7, .Rows - 1, 8) = RGB(230, 230, 250)
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 2) = RGB(128, 128, 128)
            .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 2) = Colores.Blanco
            
            '.Cell(flexcpBackColor, 1, 7, .Rows - 1, 8) = RGB(230, 230, 250)
            '.Subtotal flexSTSum, -1, 4, , RGB(123, 104, 238), Colores.Blanco, , "Total Detalle"
            '.Subtotal flexSTSum, -1, 5: .Subtotal flexSTSum, -1, 6: .Subtotal flexSTSum, -1, 7
            '.Subtotal flexSTSum, -1, 8
        End If
        .AddItem ""
    End With
    Exit Function

errSQL:
    clsGeneral.OcurrioError "Error al realizar la consulta de datos.", Err.Description
End Function

Private Sub InicializoGrillas()

  
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
        .Top = 30
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
        EjecutarApp prmPathApp & "Ingreso de Facturas.exe", vsGastos.Cell(flexcpValue, vsGastos.Row, 0)
    End If
    
    
errFnc:
End Sub

Private Sub zfn_InicializoControles()
On Error GoTo errIni
        
    With cPrint
        .Orientation = opLandscape
        .Caption = Me.Caption
        .Header = Me.Caption
        .PageBorder = opTopBottom
    End With
    

    With vsGastos
        .MultiTotals = True
        .FixedCols = 0
        .GridLines = flexGridFlatHorz
        .SubtotalPosition = flexSTBelow
        .ExtendLastCol = True
        .BorderStyle = flexBorderNone
        .Appearance = flexFlat
        
        .Cols = 1: .Rows = 1:
        .FormatString = "<Id Gasto|<Fecha|<Rubro|<SubRubro|>Importe $(G)|>Cofis $(G)|>IVA $(G)|>Debe/Haber $|>Debe/Haber|>TC|"
            
        .WordWrap = False
        .ColWidth(0) = 800: .ColWidth(1) = 800
        .ColWidth(2) = 1900: .ColWidth(3) = 2150
        
        .ColWidth(4) = 1200: .ColWidth(5) = 800: .ColWidth(6) = 1000
        .ColWidth(7) = 1300
        .ColWidth(9) = 650
        
        .MergeCells = flexMergeSpill
    End With
      
      
    With frmListado.img1
        bImprimir.Picture = .ListImages("print").ExtractIcon
        bCancelar.Picture = .ListImages("salir").ExtractIcon
        bPreview.Picture = .ListImages("vista1").ExtractIcon
    End With
    
    Exit Sub
errIni:
    clsGeneral.OcurrioError "Error cargar datos de la disponibilidad.", Err.Description
End Sub

Private Function fnc_TituloDisponibilidad()

    prmDisponibilidadSRCh = 0
    
    cons = "Select * from Disponibilidad Where DisID = " & prmDisponibilidadID
    Set rsSQL = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    If Not rsSQL.EOF Then   'DisNombre, DisMoneda, DisIDSRCheque
        prmDisponibilidadN = Trim(rsSQL!DisNombre)
        prmDisponibilidadM = Trim(rsSQL!DisMoneda)
        If Not IsNull(rsSQL!DisIDSRCheque) Then prmDisponibilidadSRCh = rsSQL!DisIDSRCheque
    End If
    rsSQL.Close
    
    'lCaption.Caption = "Detalle de Gastos Pagos/Cobrados con la Disponibilidad " & prmDisponibilidadN
    With vsGastos
        .AddItem ""
        .Cell(flexcpText, .Rows - 1, 0) = prmDisponibilidadN
        .Cell(flexcpFontBold, .Rows - 1, 0) = True
    End With
    
End Function
