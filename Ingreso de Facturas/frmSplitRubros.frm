VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmSplitRubros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Varios Rubros"
   ClientHeight    =   2550
   ClientLeft      =   3390
   ClientTop       =   4005
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplitRubros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5925
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bExit 
      Caption         =   "&Cancelar"
      Height          =   325
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   915
   End
   Begin VB.CommandButton bOK 
      Caption         =   "&Aceptar"
      Height          =   325
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   915
   End
   Begin VB.TextBox tImporte 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      MaxLength       =   13
      TabIndex        =   2
      Top             =   420
      Width           =   915
   End
   Begin VB.TextBox tSubRubro 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   420
      Width           =   3315
   End
   Begin VB.TextBox tRubro 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3315
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsGasto 
      Height          =   1305
      Left            =   60
      TabIndex        =   3
      Top             =   780
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2302
      _ConvInfo       =   1
      Appearance      =   0
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
      FocusRect       =   2
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   5
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "S&ubrubro:"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "I&mporte:"
      Height          =   255
      Left            =   4260
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Rubro:"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmSplitRubros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim idX As Integer, aValor As Long
Dim I As Integer

Public prmTotal As Currency
Public prmOK As Boolean
Public prmEdit As Boolean

Private Sub CargoGrillaDesdeArray()
On Error GoTo errCargar

Dim mSuma As Currency

    For idX = LBound(arrRubros) To UBound(arrRubros)
        If arrRubros(idX).IdRubro <> 0 Then
            With vsGasto
                .AddItem ""
                .Cell(flexcpText, .Rows - 1, 0) = Trim(arrRubros(idX).TextoRubro)
                aValor = arrRubros(idX).IdRubro: .Cell(flexcpData, .Rows - 1, 0) = aValor
                
                .Cell(flexcpText, .Rows - 1, 1) = Trim(arrRubros(idX).TextoSRubro)
                aValor = arrRubros(idX).IdSRubro: .Cell(flexcpData, .Rows - 1, 1) = aValor
                
                .Cell(flexcpText, .Rows - 1, 2) = Format(arrRubros(idX).Importe, "#,##0.00")
                
                mSuma = mSuma + arrRubros(idX).Importe
            End With
        End If
    Next
    
    If mSuma = prmTotal Then vsGasto.TabIndex = 0
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los rubros asignados.", Err.Description
End Sub

Private Sub InicializoControles()

    On Error Resume Next
    Me.BackColor = RGB(255, 250, 205)
    bOK.BackColor = Me.BackColor
    bExit.BackColor = Me.BackColor
    
    tRubro.BackColor = RGB(250, 250, 210)
    tSubRubro.BackColor = tRubro.BackColor
    tImporte.BackColor = tRubro.BackColor
    
    With vsGasto
        .Rows = 1: .Cols = 1
        .Editable = False
        .FormatString = "<Rubro|<Subrubro|>Importe|"
        .ExtendLastCol = True
        .WordWrap = True
        .ColWidth(0) = 2000: .ColWidth(1) = 2200: .ColWidth(2) = 1200
        .ColDataType(2) = flexDTCurrency
        
        .BackColor = tRubro.BackColor
        .BackColorBkg = tRubro.BackColor
        .BackColorFixed = RGB(250, 245, 205)
        .ForeColorFixed = &H800000
        .BorderStyle = flexBorderNone ' flexBorderFlat
        .GridLinesFixed = flexGridInsetHorz
        
        .BackColorSel = RGB(240, 230, 140)
        .ForeColorSel = Colores.RojoClaro
        
        .HighLight = flexHighlightWithFocus
        .FocusRect = flexFocusNone
    End With
    
    If Not prmEdit Then
        tRubro.Enabled = False
        tSubRubro.Enabled = False
        tImporte.Enabled = False
        bOK.Enabled = False
    End If
    
End Sub

Private Sub bExit_Click()
    Unload Me
End Sub

Private Sub bOK_Click()
    
    If Not ValidoCampos Then Exit Sub
    prmOK = True
    Unload Me
    
End Sub

Private Sub Form_Load()
    On Error Resume Next
    prmOK = False
    InicializoControles
    
    CargoGrillaDesdeArray
    
End Sub

Private Sub tImporte_GotFocus()
    tImporte.SelStart = 0
    tImporte.SelLength = Len(tImporte.Text)
End Sub

Private Sub tImporte_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Val(tSubRubro.Tag) = 0 And Trim(tImporte.Text) = "" Then bOK.SetFocus: Exit Sub
        If Val(tSubRubro.Tag) = 0 Or Not IsNumeric(tImporte.Text) Then Exit Sub
        
        'Verifico si el gasto está en la lista----------------------------------------
        For I = 1 To vsGasto.Rows - 1
            If vsGasto.Cell(flexcpData, I, 1) = Val(tSubRubro.Tag) Then
                MsgBox "El subrubro seleccionado ya está ingresado. Verifique la lista de distribución.", vbInformation, "Subrubro Ingresado"
                Exit Sub
            End If
        Next
        
        'Agrego el Gasto A la lista de gastos----------------------------------------
        On Error GoTo errAgregar
        Screen.MousePointer = 11
        Dim aValor As Integer
        With vsGasto
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Trim(tRubro.Text)
            aValor = Val(tRubro.Tag): .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            .Cell(flexcpText, .Rows - 1, 1) = Trim(tSubRubro.Text)
            aValor = Val(tSubRubro.Tag): .Cell(flexcpData, .Rows - 1, 1) = aValor
            
            .Cell(flexcpText, .Rows - 1, 2) = Format(tImporte.Text, FormatoMonedaP)
        End With
        
        tImporte.Text = "": tRubro.Text = ""
        Dim aSuma As Currency: aSuma = 0
        For I = 1 To vsGasto.Rows - 1
            aSuma = aSuma + vsGasto.Cell(flexcpValue, I, 2)
        Next
        If prmTotal - aSuma <> 0 Then Foco tRubro Else bOK.SetFocus
        
        Screen.MousePointer = 0     '-------------------------------------------------------------------
    End If
    Exit Sub
    
errAgregar:
    clsGeneral.OcurrioError "Error al agregar el gasto a la lista.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tRubro_Change()
    
    If Val(tRubro.Tag) <> 0 Then
        tRubro.Tag = 0
        If Val(tSubRubro.Tag) <> 0 Then tSubRubro.Text = ""
    End If
    
End Sub

Private Sub tRubro_GotFocus()
    tRubro.SelStart = 0: tRubro.SelLength = Len(tRubro.Text)
End Sub

Private Sub tRubro_KeyPress(KeyAscii As Integer)
On Error GoTo errBS
    
    If KeyAscii = vbKeyReturn Then
        
        If Val(tRubro.Tag) <> 0 Then Foco tSubRubro: Exit Sub
        If Trim(tRubro.Text) = "" Then Foco tSubRubro: Exit Sub
        
        ing_BuscoRubro tRubro
        
    End If
    Exit Sub

errBS:
    clsGeneral.OcurrioError "Error al buscar el rubro.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub tSubRubro_Change()
    tSubRubro.Tag = 0
End Sub

Private Sub tSubRubro_GotFocus()
    tSubRubro.SelStart = 0: tSubRubro.SelLength = Len(tSubRubro.Text)
End Sub

Private Sub tSubRubro_KeyPress(KeyAscii As Integer)
On Error GoTo errBS
    
    If KeyAscii = vbKeyReturn Then
                
        If Trim(tSubRubro.Text) = "" And vsGasto.Rows > 1 Then
            bOK.SetFocus: Exit Sub
        End If
        
        If Val(tSubRubro.Tag) <> 0 Then
                    
                Dim aSuma As Currency: aSuma = 0
                For I = 1 To vsGasto.Rows - 1
                    aSuma = aSuma + vsGasto.Cell(flexcpValue, I, 2)
                Next
                tImporte.Text = Format(prmTotal - aSuma, "#,##0.00")
                
                Foco tImporte
                Exit Sub
        End If
        
        If Trim(tSubRubro.Text) <> "" Then ing_BuscoSubrubro tRubro, tSubRubro
        
    End If
    Exit Sub

errBS:
    clsGeneral.OcurrioError "Error al buscar el subrubro.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Function ValidoCampos() As Boolean
On Error GoTo errValidar

    ValidoCampos = False
    
    Dim mSuma As Currency: mSuma = 0
    For I = 1 To vsGasto.Rows - 1
        mSuma = mSuma + vsGasto.Cell(flexcpValue, I, 2)
    Next
    
    If prmTotal - mSuma <> 0 Then
        MsgBox "El importe del gasto a distribuir es de " & Format(prmTotal, "#,##0.00") & vbCrLf & _
                    "Se han asignado rubros por un total del " & Format(mSuma, "#,##0.00") & vbCrLf & vbCrLf & _
                    "Verifique los datos ingresados.", vbExclamation, "Diferencia en Asignación"
        Foco tSubRubro: Exit Function
    End If
    
    'Cargo array con los datos ------------------------------------------------------------
    ReDim arrRubros(0)
    idX = 0
    For I = 1 To vsGasto.Rows - 1
    
        ReDim Preserve arrRubros(idX)
        
        With arrRubros(idX)
            .IdRubro = vsGasto.Cell(flexcpData, I, 0)
            .TextoRubro = Trim(vsGasto.Cell(flexcpText, I, 0))
            .IdSRubro = vsGasto.Cell(flexcpData, I, 1)
            .TextoSRubro = Trim(vsGasto.Cell(flexcpText, I, 1))
            .Importe = CCur(vsGasto.Cell(flexcpText, I, 2))
        End With
        
        idX = idX + 1
    Next
    '----------------------------------------------------------------------------------------------
                
    ValidoCampos = True
    Exit Function

errValidar:
    clsGeneral.OcurrioError "Error al validar los datos.", Err.Description
End Function

Private Sub vsGasto_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Not prmEdit Then Exit Sub
    If vsGasto.Rows = 1 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyDelete
            vsGasto.RemoveItem vsGasto.Row
            
        Case vbKeyReturn, vbKeySpace
            With vsGasto
                tRubro.Text = Trim(.Cell(flexcpText, .Row, 0))
                tRubro.Tag = vsGasto.Cell(flexcpData, .Row, 0)
                
                tSubRubro.Text = Trim(vsGasto.Cell(flexcpText, .Row, 1))
                tSubRubro.Tag = vsGasto.Cell(flexcpData, .Row, 1)
                
                tImporte.Text = vsGasto.Cell(flexcpText, .Row, 2)
                
                .RemoveItem .Row
            End With
            Foco tRubro
    End Select

    
End Sub

