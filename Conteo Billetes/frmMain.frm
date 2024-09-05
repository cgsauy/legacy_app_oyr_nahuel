VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "Vsflex6d.ocx"
Object = "{6CF96EEB-5F9F-11D3-B46E-827621868276}#2.1#0"; "aacombo.ocx"
Object = "{D2FFAA40-074A-11D1-BAA2-444553540000}#3.0#0"; "VsVIEW3.ocx"
Object = "{923DD7D8-A030-4239-BCD4-51FDB459E0FE}#4.0#0"; "orgComboCalculator.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H8000000B&
   Caption         =   "Conteo de Billetes"
   ClientHeight    =   5145
   ClientLeft      =   2760
   ClientTop       =   3330
   ClientWidth     =   7665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7665
   Begin orgCalculatorFlat.orgCalculator tQBillete 
      Height          =   285
      Left            =   2520
      TabIndex        =   18
      Top             =   1140
      Width           =   1095
      _ExtentX        =   1931
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
      Formato         =   "#,##0"
   End
   Begin VB.PictureBox picPie 
      Height          =   690
      Left            =   180
      ScaleHeight     =   630
      ScaleWidth      =   5715
      TabIndex        =   11
      Top             =   4200
      Width           =   5775
      Begin VB.CommandButton bPrint 
         Caption         =   "&Imprimir"
         Height          =   315
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   300
         Width           =   795
      End
      Begin VB.CommandButton bGrabar 
         Caption         =   "&Grabar"
         Height          =   315
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   300
         Width           =   795
      End
      Begin VB.CommandButton bExit 
         Caption         =   "&Salir"
         Height          =   315
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   300
         Width           =   795
      End
      Begin VB.Label lTotalU 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   4335
      End
      Begin VB.Label lTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox tFecha 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   360
      Width           =   1035
   End
   Begin VB.ComboBox cDisponibilidad 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin AACombo99.AACombo cUbicacion 
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   1140
      Width           =   2355
      _ExtentX        =   4154
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
   Begin VB.TextBox tIBillete 
      Appearance      =   0  'Flat
      Height          =   290
      Left            =   3660
      TabIndex        =   8
      Top             =   1140
      Width           =   795
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsLista 
      Height          =   1935
      Left            =   60
      TabIndex        =   10
      Top             =   1620
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   3413
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
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
   Begin vsViewLib.vsPrinter vsPrinter 
      Height          =   3795
      Left            =   4320
      TabIndex        =   15
      Top             =   780
      Visible         =   0   'False
      Width           =   3075
      _Version        =   196608
      _ExtentX        =   5424
      _ExtentY        =   6694
      _StockProps     =   229
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   End
   Begin VB.Label Label5 
      Caption         =   "&Lista"
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   195
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&De $:"
      Height          =   195
      Left            =   3660
      TabIndex        =   7
      Top             =   900
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Q Billetes:"
      Height          =   195
      Left            =   2520
      TabIndex        =   6
      Top             =   900
      Width           =   795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Ubicación:"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha:"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dis&ponibilidad:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu MnuMover 
      Caption         =   "Mover a:"
      Visible         =   0   'False
      Begin VB.Menu MnuTitulo 
         Caption         =   "Mover a:"
      End
      Begin VB.Menu MnuML1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMovTo 
         Caption         =   "Mover"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim prmGrabar As Boolean

Dim mTexto As String
Dim mValor As Long

Private Sub bExit_Click()
    If prmGrabar Then
        If MsgBox("Ud. realizó modificaciones en el conteo de billetes y no grabó." & vbCrLf & _
                        "Descarta las modificaciones ?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub bGrabar_Click()
    AccionGrabar
End Sub

Private Sub bPrint_Click()
    AccionImprimir
End Sub

Private Sub cDisponibilidad_Click()
    cUbicacion.Clear
    If vsLista.Rows > 1 Then
        vsLista.Rows = 1
        addTotal 0, True
    End If
End Sub

Private Sub cDisponibilidad_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If KeyAscii = vbKeyReturn Then
        If cDisponibilidad.ListIndex <> -1 Then
            CargoUbicaciones cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
            Foco tFecha
        End If
    End If
    
End Sub

Private Sub cUbicacion_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        If cUbicacion.ListIndex <> -1 Then Foco tQBillete: Exit Sub
        
        If Trim(cUbicacion.Text) <> "" Then
            If cDisponibilidad.ListIndex = -1 Then cDisponibilidad.SetFocus: Exit Sub
            
            If MsgBox("La ubicación ingresada no existe. " & vbCrLf & _
                           "Quiere ingresar una nueva ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Ingresar Nueva Ubicación") = vbNo Then Exit Sub
            
            AccionNuevaUbicacion
        
        Else
            bGrabar.SetFocus
        End If
    End If
    
End Sub

Private Sub AccionNuevaUbicacion()
    
    Dim mID As Long, mNombre As String
    
    With frmUbicacion
        .prmNombre = Trim(cUbicacion.Text)
        .prmIdDisponibilidad = cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
        .Show vbModal, Me
        mID = .prmAddId
        mNombre = .prmAddTexto
        Me.Refresh
    End With
    
    If mID <> 0 Then
        CargoUbicaciones cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
        cUbicacion.Text = mNombre
        Foco tQBillete
    End If
    
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.Refresh
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    
    FechaDelServidor
    Me.BackColor = RGB(255, 240, 245) 'RGB(222, 184, 135)
    
    picPie.BackColor = Me.BackColor
    picPie.BorderStyle = 0
    
    ObtengoSeteoForm Me ', WidthIni:=5040, HeightIni:=3870
'    Me.Height = 3870
'    Me.Width = 5040
    
    InicializoControles
    
    Exit Sub
errLoad:
    clsGeneral.OcurrioError "Error al iniciar el formulario.", Err.Description
End Sub

Public Function prmCargoDatos(prmDisponibilidad As Long, prmFecha As String)
    
    On Error Resume Next
    If prmDisponibilidad = 0 Then
        prmDisponibilidad = paDisponibilidad
        prmFecha = gFechaServidor
    End If
    BuscoCodigoEnCombo cDisponibilidad, prmDisponibilidad
    CargoUbicaciones prmDisponibilidad
    
    If IsDate(prmFecha) Then
        tFecha.Text = Format(prmFecha, "dd/mm/yyyy")
        tFecha.Tag = tFecha.Text

        CargoDatosDesdeBD prmDisponibilidad, tFecha.Tag
        If vsLista.Rows > 1 Then cUbicacion.SetFocus
    End If
    
End Function

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With picPie
        .Top = Me.ScaleHeight - .Height
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
    End With
    
    With vsLista
        .Left = 60
        .Height = Me.ScaleHeight - picPie.Height - .Top
        .Width = Me.ScaleWidth - (.Left * 2)
    End With
    
    bExit.Left = picPie.ScaleWidth - (bExit.Width + 60)
    bGrabar.Left = bExit.Left - (bGrabar.Width + 100)
    bPrint.Left = bGrabar.Left - (bGrabar.Width + 100)
    
    lTotal.Left = 60: lTotal.Width = bPrint.Left - lTotal.Left
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    GuardoSeteoForm Me
    EndMain
End Sub


Private Sub InicializoControles()
    
    On Error Resume Next
    Cons = "Select DisID, DisNombre from Disponibilidad " & _
                " Where DisIDSRCheque is Null " & _
                " Order by DisNombre"
    CargoCombo Cons, cDisponibilidad
    
    bExit.BackColor = Me.BackColor
    bGrabar.BackColor = Me.BackColor
    bPrint.BackColor = Me.BackColor
    cDisponibilidad.ListIndex = -1
    
    cUbicacion.Text = "": tQBillete.Clean: tIBillete.Text = ""
    
    With vsLista
        .Editable = False
        .Rows = 1: .Cols = 4
        .FormatString = ">Q Billetes|>De|>Total|<Ubicación"
        .ColWidth(0) = 900: .ColWidth(1) = 800: .ColWidth(2) = 1000: .ColWidth(3) = 1750
        
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        .AllowBigSelection = False
        .AllowSelection = False
        
        .BackColorBkg = Me.BackColor
        .BackColor = vbWhite
        .BackColorAlternate = Me.BackColor
        
        .BackColorFixed = Me.BackColor
        '.ForeColorFixed = RGB(250, 240, 230)
        
        .BorderStyle = flexBorderNone
        .HighLight = flexHighlightWithFocus
        .FocusRect = flexFocusNone
        .RowHeight(0) = 300
        .ForeColorSel = Colores.RojoClaro
        .BackColorSel = RGB(216, 191, 216)
        .ScrollBars = flexScrollBarVertical
    End With
   
    With tQBillete
        .BackColorButtonOver = Me.BackColor
        .BackColorCalculator = .BackColorButton
        '.BackColorNumberOver = .BackColorOperatorOver
    End With
    
    lTotal.Caption = ""
    lTotalU.Caption = ""
     
    With vsPrinter
        .PhysicalPage = True
        .PaperSize = vbPRPSLetter
        .Orientation = orPortrait
        .PreviewMode = pmScreen
        .PreviewPage = 1
        .Zoom = 100
        .MarginLeft = 800: .MarginRight = 350
        .MarginBottom = 750: .MarginTop = 750
        .Visible = False
    End With

End Sub


Private Sub lTotalU_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errXX
    If vsLista.Rows = 1 Then Exit Sub
    If lTotalU.Caption = "" Then Exit Sub
    
    If Button = vbRightButton Then
        vsLista.SetFocus
        vsLista.Select vsLista.Row, vsLista.Col
        ActivoMenuMover picPie.Top + lTotalU.Top + Y, X, True
    End If
errXX:
End Sub

Private Sub MnuMovTo_Click(Index As Integer)
On Error GoTo errMover

Dim bSalir As Boolean
Dim mOldUbi As Long

    'Q Billetes|<De|>Total|<Ubicación"
    Dim mNewUbi As Long, mRow As Integer
    
    mNewUbi = Val(MnuMovTo(Index).Tag)
    If mNewUbi = 0 Then Exit Sub
    
    Dim mQ As Long, mDe As String, mOldRow As Integer
    Dim mValor As Long, mTotal As Currency
    
    mOldRow = vsLista.Row
    mOldUbi = vsLista.Cell(flexcpData, mOldRow, 3)
    
    bSalir = False
    Do While Not bSalir
        mQ = vsLista.Cell(flexcpValue, mOldRow, 0)
        mDe = Trim(vsLista.Cell(flexcpText, mOldRow, 1))
        vsLista.RemoveItem mOldRow
        
        mRow = 0
    
        With vsLista
            For I = 1 To .Rows - 1          'Busco si existe billete y ubicacion nueva
                If Trim(.Cell(flexcpText, I, 1)) = mDe And _
                    .Cell(flexcpData, I, 3) = mNewUbi Then
                    mRow = I: Exit For
                End If
            Next
            If mRow = 0 Then
                .AddItem ""
                mRow = .Rows - 1
                .Cell(flexcpText, mRow, 0) = Format(mQ, "#,##0")
                .Cell(flexcpText, mRow, 1) = Format(mDe, "#,##0.00")
                
                .Cell(flexcpText, mRow, 3) = Trim(mID(MnuMovTo(Index).Caption, 2)) 'Por el & del Titulo
                mValor = mNewUbi: .Cell(flexcpData, mRow, 3) = mValor
            Else
                .Cell(flexcpText, mRow, 0) = Format(.Cell(flexcpValue, mRow, 0) + mQ, "#,##0")
            End If
                   
            mTotal = .Cell(flexcpValue, mRow, 0) * .Cell(flexcpValue, mRow, 1)
            .Cell(flexcpText, mRow, 2) = Format(mTotal, "#,##0.00")
                
            If Val(MnuTitulo.Tag) = 1 Then
                bSalir = True
            Else
                'Busco la siguinte fila con la ubicacion y la selecciono
                Dim xF As Integer, bQuedan As Boolean
                bQuedan = False
                For xF = 1 To .Rows - 1
                    If mOldUbi = .Cell(flexcpData, xF, 3) Then
                        bQuedan = True
                        mOldRow = xF 'vsLista.Select xF, 1
                        Exit For
                    End If
                Next
            End If
            
            If Not bQuedan Then bSalir = True
            
        End With
    Loop
    
    With vsLista
        '.RemoveItem mOldRow
        
        .ColSort(3) = flexSortGenericAscending
        .ColSort(1) = flexSortNumericAscending
        .Select 0, 1, 0, 1
        .Sort = flexSortUseColSort
        .Select 0, 3, 0, 3
        .Sort = flexSortUseColSort
        
        For I = 1 To .Rows - 1
            If .Cell(flexcpData, I, 3) = mNewUbi And .Cell(flexcpText, I, 1) = Format(mDe, "#,##0.00") Then
                .TopRow = I
                .Select I, 0
                Exit For
            End If
        Next
    End With
    Exit Sub

errMover:
    clsGeneral.OcurrioError "Error al mover los billetes de ubicación.", Err.Description
End Sub

Private Sub tFecha_Change()
    
    If vsLista.Rows > 1 Then
        If prmGrabar Then
            If tFecha.Tag <> "" Then
                If MsgBox("Ud. realizó modificaciones en el conteo de billetes y no grabó." & vbCrLf & _
                            "Descarta las modificaciones ?.", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    mTexto = tFecha.Tag: tFecha.Tag = ""
                    tFecha.Text = Format(mTexto, "dd/mm/yyyy")
                    tFecha.Tag = tFecha.Text
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
        End If
            
        vsLista.Rows = 1
        addTotal 0, True
        lTotalU.Caption = ""
        
    End If
    If Trim(tFecha.Tag) <> "" Then tFecha.Tag = ""
    
End Sub

Private Sub tFecha_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If IsDate(tFecha.Text) Then
            
            tFecha.Text = Format(tFecha.Text, "dd/mm/yyyy")
            tFecha.Tag = tFecha.Text
            If cDisponibilidad.ListIndex = -1 Then cDisponibilidad.SetFocus: Exit Sub
            
            CargoDatosDesdeBD cDisponibilidad.ItemData(cDisponibilidad.ListIndex), tFecha.Tag
            cUbicacion.SetFocus
        End If
        
    End If
    
End Sub

Private Sub tIBillete_GotFocus()
    tIBillete.SelStart = 0: tIBillete.SelLength = Len(tIBillete.Text)
End Sub

Private Sub tIBillete_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tIBillete.Text) Then
            tIBillete.Text = Format(tIBillete.Text, "#,##0.00")
            AgregoBillete
        End If
    End If
    
End Sub

Private Function AgregoBillete()

    If cUbicacion.ListIndex = -1 Then cUbicacion.SetFocus: Exit Function
    If Not IsNumeric(tQBillete.Text) Then Foco tQBillete: Exit Function
    If Not IsNumeric(tIBillete.Text) Then Foco tIBillete: Exit Function
    
    'Q Billetes|<De|>Total|<Ubicación"
    Dim mTotal As Currency, mRow As Integer
    mRow = 0

    With vsLista
        For I = 1 To .Rows - 1          'Busco si existe billete y ubicacion
            If .Cell(flexcpText, I, 1) = Format(tIBillete.Text, "#,##0.00") And _
                .Cell(flexcpData, I, 3) = cUbicacion.ItemData(cUbicacion.ListIndex) Then
                mRow = I: Exit For
            End If
        Next
        If mRow = 0 Then
            .AddItem ""
            mRow = .Rows - 1
        Else
            If MsgBox("Se van a sustituir " & .Cell(flexcpText, mRow, 0) & " billetes de " & .Cell(flexcpText, mRow, 1) & _
                        " por " & tQBillete.Text & " billetes." & vbCrLf & vbCrLf & _
                        "Confirma sustituir los valores ?.", vbQuestion + vbYesNo, "Sustituir Cantidades") = vbNo Then Exit Function
        End If
        
        prmGrabar = True
        
        .Cell(flexcpText, mRow, 0) = Format(tQBillete.Text, "#,##0")
        .Cell(flexcpText, mRow, 1) = Format(tIBillete.Text, "#,##0.00")
        
        addTotal .Cell(flexcpValue, mRow, 2) * -1
        mTotal = .Cell(flexcpValue, mRow, 0) * .Cell(flexcpValue, mRow, 1)
        .Cell(flexcpText, mRow, 2) = Format(mTotal, "#,##0.00")
        addTotal mTotal
        
        .Cell(flexcpText, mRow, 3) = Trim(cUbicacion.Text)
        mValor = cUbicacion.ItemData(cUbicacion.ListIndex)
        .Cell(flexcpData, mRow, 3) = mValor
                
        .ColSort(3) = flexSortGenericAscending
        .ColSort(1) = flexSortNumericAscending
        .Select 0, 1, 0, 1
        .Sort = flexSortUseColSort
        .Select 0, 3, 0, 3
        .Sort = flexSortUseColSort
        
        For I = 1 To .Rows - 1
            If .Cell(flexcpData, I, 3) = cUbicacion.ItemData(cUbicacion.ListIndex) And .Cell(flexcpText, I, 0) = Format(tQBillete.Text, "#,##0") And .Cell(flexcpText, I, 1) = Format(tIBillete.Text, "#,##0.00") Then
                .TopRow = I
                Exit For
            End If
        Next
    End With
    
    tQBillete.Clean: tIBillete.Text = ""
    Foco cUbicacion
    
End Function

Private Sub tQBillete_GotFocus()
    'tQBillete.SelStart = 0: tQBillete.SelLength = Len(tQBillete.Text)
End Sub


Private Sub tQBillete_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If IsNumeric(tQBillete.Text) Then
            If cUbicacion.ListIndex = -1 Then Foco cUbicacion: Exit Sub
            'tQBillete.Text = Format(tQBillete.Text, "#,##0")
            Foco tIBillete
        Else
            If Trim(tQBillete.Text) = "" And vsLista.Rows > 1 Then bGrabar.SetFocus
        End If
    End If
    
End Sub

Private Sub vsLista_GotFocus()
    On Error Resume Next
    If vsLista.Row = 0 Then vsLista.Row = 1
End Sub

Private Sub vsLista_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeySpace:
                If vsLista.Rows = 1 Then Exit Sub
                prmGrabar = True
                With vsLista
                    BuscoCodigoEnCombo cUbicacion, .Cell(flexcpData, .Row, 3)
                    tQBillete.Text = .Cell(flexcpText, .Row, 0)
                    tIBillete.Text = .Cell(flexcpText, .Row, 1)
                    
                    addTotal .Cell(flexcpValue, .Row, 2) * -1
                    
                    .RemoveItem .Row
                    Foco cUbicacion
                End With
        
        Case vbKeyDelete
                If vsLista.Rows = 1 Then Exit Sub
                addTotal vsLista.Cell(flexcpValue, vsLista.Row, 2) * -1
                vsLista.RemoveItem vsLista.Row
                
                prmGrabar = True
        
        Case 93
                ActivoMenuMover vsLista.Top + (vsLista.Row * vsLista.RowHeight(1)) + vsLista.RowHeight(0)
                
    End Select
    
End Sub

Private Sub CargoUbicaciones(mIDDisponibilidad As Long)

    cUbicacion.Clear: cUbicacion.Refresh
    Cons = "Select UBiCodigo, UBiNombre from UbicacionBillete Where UBiDisponibilidad = " & mIDDisponibilidad
    CargoCombo Cons, cUbicacion

    'Limpio Menu    ------------------------------
    MnuMovTo(0).Visible = True
    For I = 1 To MnuMovTo.UBound
        Unload MnuMovTo(I)
    Next
    
    Dim idxM As Integer
    'Cargo Menu con datos del combo ---------
    Cons = "Select UBiCodigo, UBiNombre from UbicacionBillete Where UBiDisponibilidad = " & mIDDisponibilidad
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        idxM = MnuMovTo.UBound + 1
        Load MnuMovTo(idxM)
        With MnuMovTo(idxM)
            .Caption = "&" & Trim(RsAux!UBiNombre)
            .Tag = RsAux!UbiCodigo
            .Visible = True
        End With
        
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If MnuMovTo.Count > 1 Then MnuMovTo(0).Visible = False
    
End Sub

Private Sub CargoDatosDesdeBD(mIDDisponibilidad As Long, mFecha As String)

On Error GoTo errCargar

    Screen.MousePointer = 11
    vsLista.Rows = 1
    prmGrabar = False
    Dim mTotal As Currency
    
    Cons = "Select * from ConteoBillete, UbicacionBillete " & _
            " Where CBiDisponibilidad = " & mIDDisponibilidad & _
            " And CBiFecha = " & Format(mFecha, "'mm/dd/yyyy'") & _
            " And CBiUbicacion = UBiCodigo"
    
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
    
        With vsLista
        
            .AddItem ""
            .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!CBiQ, "#,##0")
            .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!CBiBilleteDe, "#,##0.00")
            
            mTotal = .Cell(flexcpValue, .Rows - 1, 0) * .Cell(flexcpValue, .Rows - 1, 1)
            .Cell(flexcpText, .Rows - 1, 2) = Format(mTotal, "#,##0.00")
            addTotal mTotal
            
            .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!UBiNombre)
            mValor = RsAux!UbiCodigo: .Cell(flexcpData, .Rows - 1, 3) = mValor
        End With
    
        RsAux.MoveNext
    Loop
    RsAux.Close
    
    If vsLista.Rows = 1 Then
        Dim mFAnterior As String: mFAnterior = ""
        
        Cons = "Select Max(CBiFecha)  from ConteoBillete" & _
            " Where CBiDisponibilidad = " & mIDDisponibilidad & _
            " And CBiFecha < " & Format(mFecha, "'mm/dd/yyyy'")
        Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
        If Not RsAux.EOF Then
            If Not IsNull(RsAux(0)) Then mFAnterior = RsAux(0)
        End If
        RsAux.Close
        
        If Trim(mFAnterior) <> "" Then
            If MsgBox("Quiere cargar los billetes contados el " & Format(mFecha, "d/m/yyyy") & "," & vbCrLf & _
                        "que fueron ubicados en lugares normalmente fijos ?", vbQuestion + vbYesNo, "Cargar Conteos Fijos ?") = vbYes Then
                        
                prmGrabar = True
                
                Cons = "Select * from ConteoBillete, UbicacionBillete " & _
                            " Where CBiDisponibilidad = " & mIDDisponibilidad & _
                            " And CBiFecha = " & Format(mFAnterior, "'mm/dd/yyyy'") & _
                            " And CBiUbicacion = UBiCodigo" & _
                            " And UBiFijo = 1"
                    
                Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
                Do While Not RsAux.EOF
                
                    With vsLista
                        .AddItem ""
                        .Cell(flexcpText, .Rows - 1, 0) = Format(RsAux!CBiQ, "#,##0")
                        .Cell(flexcpText, .Rows - 1, 1) = Format(RsAux!CBiBilleteDe, "#,##0.00")
                        
                        mTotal = .Cell(flexcpValue, .Rows - 1, 0) * .Cell(flexcpValue, .Rows - 1, 1)
                        .Cell(flexcpText, .Rows - 1, 2) = Format(mTotal, "#,##0.00")
                        addTotal mTotal
                        
                        .Cell(flexcpText, .Rows - 1, 3) = Trim(RsAux!UBiNombre)
                        mValor = RsAux!UbiCodigo: .Cell(flexcpData, .Rows - 1, 3) = mValor
                    End With
                
                    RsAux.MoveNext
                Loop
                RsAux.Close
            End If
        End If
    End If
    
    If vsLista.Rows > 1 Then
        With vsLista
             .ColSort(3) = flexSortGenericAscending
            .ColSort(1) = flexSortNumericAscending
            .Select 0, 1, 0, 1
            .Sort = flexSortUseColSort
            .Select 0, 3, 0, 3
            .Sort = flexSortUseColSort
        End With
    End If
    
    Screen.MousePointer = 0
    Exit Sub
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos del conteo.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub AccionGrabar()

    If Not ValidoDatos Then Exit Sub
    
    mTexto = ""
    If vsLista.Rows = 1 Then mTexto = vbCrLf & "ATENCIÓN: Se eliminará el conteo del " & tFecha.Text
    If MsgBox("Confirma grabar el conteo de billetes." & mTexto, vbQuestion + vbYesNo, "Grabar Conteo") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    On Error GoTo errorBT
    
    Dim mDisp As Long, mFecha As String
    mDisp = cDisponibilidad.ItemData(cDisponibilidad.ListIndex)
    mFecha = tFecha.Text
    
    cBase.BeginTrans            'COMIENZO TRANSACCION------------------------------------------     !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    On Error GoTo errorET
    
    Cons = "Select * from ConteoBillete " & _
               " Where CBiDisponibilidad = " & mDisp & _
               " And CBiFecha = " & Format(mFecha, "'mm/dd/yyyy'")
                
    Set RsAux = cBase.OpenResultset(Cons, rdOpenDynamic, rdConcurValues)
    Do While Not RsAux.EOF
        RsAux.Delete
        RsAux.MoveNext
    Loop
        
    With vsLista
        For I = 1 To .Rows - 1
            RsAux.AddNew
            RsAux!CBiFecha = Format(mFecha, "mm/dd/yyyy")
            RsAux!CBiDisponibilidad = mDisp
            RsAux!CBiUbicacion = .Cell(flexcpData, I, 3)
            RsAux!CBiBilleteDe = .Cell(flexcpValue, I, 1)
            RsAux!CBiQ = .Cell(flexcpValue, I, 0)
            RsAux.Update
        Next
    End With
    
    RsAux.Close
    
    cBase.CommitTrans   'FINALIZO TRANSACCION-------------------------------------------        !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    cUbicacion.Text = "": cUbicacion.SetFocus
    Screen.MousePointer = 0
    prmGrabar = False
    Exit Sub
    
errorBT:
    clsGeneral.OcurrioError "Error al grabar los datos.", Err.Description
    Screen.MousePointer = 0: Exit Sub
errorET:
    Resume ErrorRoll
ErrorRoll:
    cBase.RollbackTrans
    clsGeneral.OcurrioError "Error al grabar los datos.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Function ValidoDatos() As Boolean
    
    ValidoDatos = False
    
    If cUbicacion.ListCount = 0 Then
        MsgBox "No hay ubicaciones ingresadas.", vbExclamation, "Posible Error "
        cDisponibilidad.SetFocus: Exit Function
    End If
    
    If cDisponibilidad.ListIndex = -1 Then cDisponibilidad.SetFocus: Exit Function
    
    If Trim(tFecha.Tag) = "" Then
        MsgBox "Falta cargar el conteo de la fecha ingresada.", vbExclamation, "Posible Error "
        Foco tFecha: Exit Function
    End If
    
    ValidoDatos = True
    
End Function

Private Function addTotal(mImporte As Currency, Optional bACero As Boolean = False)
    
    If bACero Then lTotal.Tag = 0

    If Trim(lTotal.Tag) = "" Then lTotal.Tag = 0
    lTotal.Tag = CCur(lTotal.Tag) + mImporte
    
    lTotal.Caption = "Total: " & Format(lTotal.Tag, "#,##0.00")
    
End Function

Private Sub AccionImprimir()
Dim aFormato As String, aEncabezado As String
    
    On Error GoTo errPrint
    With vsPrinter
    
    If Not .PrintDialog(pdPrinterSetup) Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    .Preview = True
    .StartDoc
            
    If .Error Then
        MsgBox "No se pudo iniciar el documento de impresión.", vbInformation, "Error"
        Screen.MousePointer = vbDefault: Exit Sub
    End If

    EncabezadoListado vsPrinter, "Conteo de Billetes", False
    
    .FileName = "Conteo de Billetes"
    .FontSize = 8: .FontBold = False
    
    .FontSize = 10: .Text = "Disponibilidad: ": vsPrinter.Text = Trim(cDisponibilidad.Text)
    vsPrinter = ""
    .Text = "Fecha: ": vsPrinter.Text = Format(tFecha.Text, "Long Date")
    
    vsPrinter = "": vsPrinter = ""
    .FontSize = 8: .FontBold = False
    
    vsLista.ExtendLastCol = False
    .RenderControl = vsLista.hwnd
    vsLista.ExtendLastCol = True
    
    For I = 1 To 7
        .Paragraph = "_____________________________________________"
    Next
    
    vsPrinter.Text = lTotal.Caption
    .EndDoc            'Cierro el Documento--------------------------------------------------------------!!!!!!!!!!!!!!
    .PrintDoc
        
    '.ZOrder 0: .Visible = True
    '.Left = 0: .Width = Me.ScaleWidth
    End With

    Screen.MousePointer = 0
    Exit Sub
    
errPrint:
    clsGeneral.OcurrioError "Error al realizar la impresión. " & Trim(Err.Description)
    Screen.MousePointer = 0
End Sub

Private Sub ActivoMenuMover(Y As Single, Optional X As Single = 0, Optional bTodo As Boolean = False)
 On Error Resume Next
    If vsLista.Rows = 1 Then Exit Sub
    If MnuMovTo.Count <= 2 Then Exit Sub
    
    Dim idUbi As Long
    idUbi = vsLista.Cell(flexcpData, vsLista.Row, 3)
    
    For I = 1 To MnuMovTo.UBound
        MnuMovTo(I).Visible = Not (Val(MnuMovTo(I).Tag) = idUbi)
    Next
    
    Dim mX As Single
    If X = 0 Then
        mX = vsLista.ColWidth(0) + vsLista.ColWidth(1) + vsLista.ColWidth(2) + 100
    Else
        mX = X
    End If
    
    If bTodo Then
        MnuTitulo.Caption = "Mover Todo a:"
        MnuTitulo.Tag = 2
    Else
        MnuTitulo.Caption = "Mover a:"
        MnuTitulo.Tag = 1
    End If
    
    PopupMenu MnuMover, , mX, Y, MnuTitulo
    
End Sub

Private Sub vsLista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errXX
    If vsLista.Rows = 1 Then Exit Sub
    
    If Button = vbRightButton Then
        vsLista.SetFocus
        vsLista.Select vsLista.MouseRow, vsLista.MouseCol
        ActivoMenuMover vsLista.Top + Y, X
    End If
errXX:
End Sub

Private Function TotalUbicacion()

    On Error GoTo errTotal
    Dim mTotal As Currency, mUbic As String
    
    mUbic = ""
    If vsLista.Rows = 1 Or vsLista.Row < 1 Then
        lTotalU.Caption = ""
        Exit Function
    End If
    
    
    With vsLista
        mUbic = .Cell(flexcpText, .Row, 3)
        For I = 1 To .Rows - 1
                       
            If Trim(mUbic) = Trim(.Cell(flexcpText, I, 3)) Then
                mTotal = mTotal + .Cell(flexcpValue, I, 2)
            End If
        Next
    End With
    
    lTotalU.Caption = Format(mTotal, "#,##0.00") & " en " & mUbic
    
errTotal:
End Function

Private Sub vsLista_SelChange()
    TotalUbicacion
End Sub
