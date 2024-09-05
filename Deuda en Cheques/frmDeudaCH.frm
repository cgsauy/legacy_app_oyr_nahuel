VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmDeudaCH 
   Caption         =   "Deuda en Cheques "
   ClientHeight    =   3885
   ClientLeft      =   2205
   ClientTop       =   3570
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDeudaCH.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   9630
   Begin VSFlex6DAOCtl.vsFlexGrid lPago 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4048
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
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
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
   Begin VSFlex6DAOCtl.vsFlexGrid lTotal 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1296
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
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   4
      GridLinesFixed  =   2
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
   Begin VB.Label lTitular 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cedula"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   45
      UseMnemonic     =   0   'False
      Width           =   8175
   End
End
Attribute VB_Name = "frmDeudaCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gCliente As Long, aTexto As String
Dim I As Integer

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    Screen.MousePointer = 11
    ObtengoSeteoForm Me, Me.Left, Me.Top, Me.Width, Me.Height
    
    'Linea de Comandos    -------------------------------------
    If Trim(Command()) <> "" Then
        aTexto = Trim(Command())
        gCliente = Val(aTexto)
    End If
    '---------------------------------------------------------------
    InicializoGrilla
       
    CargoCliente gCliente
    CargoCheques gCliente
    
End Sub

Private Sub CargoCliente(Codigo As Long)

    On Error GoTo errCargar
    lTitular.Caption = ""
    cons = "Select Nombre = (RTrim(CPeApellido1) + RTrim(' ' + CPeApellido2)+', ' + RTrim(CPeNombre1)) + RTrim(' ' + CPeNombre2) " _
           & " From CPersona " _
           & " Where CPeCliente = " & Codigo _
                                                & " UNION " _
           & " Select Nombre = (RTrim(CEmFantasia) + RTrim(' (' + CEmNombre)+ ')')  " _
           & " From CEmpresa " _
           & " Where CEmCliente = " & Codigo
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    If Not rsAux.EOF Then lTitular.Caption = Trim(rsAux!Nombre)
    rsAux.Close
    Exit Sub
    
errCargar:
    clsGeneral.OcurrioError "Error al cargar los datos del cliente.", Err.Description
    Screen.MousePointer = 0
End Sub


Private Sub CargoCheques(Cliente As Long)

Dim RsMon As rdoResultset
Dim aMonAnterior As Long, aMonNombre As String
Dim sInserto As Boolean
Dim aValor As Long

    lTotal.Rows = 0
    aMonAnterior = 0
    On Error GoTo errPago
    
    cons = "Select * from ChequeDiferido, SucursalDeBanco, BancoSSFF" _
            & " Where CDiCliente = " & Cliente _
            & " And CDiSucursal = SBaCodigo" _
            & " And CDiBanco = BanCodigo" _
            & " And CDiCobrado Is Null" _
            & " And CDiEliminado Is Null"
    Set rsAux = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
    lPago.Rows = 1
    
    Do While Not rsAux.EOF
        With lPago
            .AddItem Format(rsAux!CDiVencimiento, "dd/mm/yyyy")
            aValor = rsAux!CDiCodigo: .Cell(flexcpData, .Rows - 1, 0) = aValor
            
            
            .Cell(flexcpText, .Rows - 1, 1) = Format(rsAux!BanCodigoB, "00") & "-" & Format(rsAux!SBaCodigoS, "000") & " " & Trim(rsAux!BanNombre) & " (" & Trim(rsAux!SBaNombre) & ")"
            .Cell(flexcpText, .Rows - 1, 2) = Trim(rsAux!CDiSerie) & " " & Trim(rsAux!CDiNumero)
            
            'Cargo la Moneda--------------------------------------------------------
            If aMonAnterior <> rsAux!CDiMoneda Then
                aMonAnterior = rsAux!CDiMoneda
                
                cons = "Select * from Moneda Where MonCodigo = " & rsAux!CDiMoneda
                Set RsMon = cBase.OpenResultset(cons, rdOpenForwardOnly, rdConcurValues)
                If Not RsMon.EOF Then aMonNombre = Trim(RsMon!MonSigno)
                RsMon.Close
            End If
            .Cell(flexcpText, .Rows - 1, 3) = Trim(aMonNombre)
            '---------------------------------------------------------------------------
            .Cell(flexcpText, .Rows - 1, 4) = Format(rsAux!CDiImporte, "#,##0.00")
            .Cell(flexcpText, .Rows - 1, 5) = Format(rsAux!CDiLibrado, "dd/mm/yyyy")
            
            'Inserto en lista de Totales---------------------------------------------------------------------------------------------
            sInserto = True
            Dim mCol As Integer
            If Trim(.Cell(flexcpText, .Rows - 1, 0)) <> "" Then mCol = 2 Else mCol = 3
            With lTotal
                For I = 0 To .Rows - 1
                    If .Cell(flexcpData, I, 0) = aMonAnterior Then
                        .Cell(flexcpText, I, mCol) = Format(.Cell(flexcpValue, I, mCol) + rsAux!CDiImporte, "#,##0.00")
                        sInserto = False: Exit For
                    End If
                Next
                If sInserto Then
                    If .Rows = 0 Then .AddItem "Totales por Moneda:" Else .AddItem ""
                    .Cell(flexcpData, .Rows - 1, 0) = aMonAnterior
                    .Cell(flexcpText, .Rows - 1, 1) = Trim(aMonNombre)
                    .Cell(flexcpText, .Rows - 1, mCol) = Format(rsAux!CDiImporte, "#,##0.00")
                End If
            End With
            '---------------------------------------------------------------------------------------------------------------------------
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close
    
    On Error Resume Next
    With lTotal
    If .Rows > 0 Then
        For I = 0 To .Rows - 1
            If Trim(.Cell(flexcpText, .Rows - 1, 2)) <> "" Then .Cell(flexcpText, .Rows - 1, 2) = "Dif.    " & .Cell(flexcpText, .Rows - 1, 2)
            If Trim(.Cell(flexcpText, .Rows - 1, 3)) <> "" Then .Cell(flexcpText, .Rows - 1, 3) = "Al día   " & .Cell(flexcpText, .Rows - 1, 3)
        Next
    End If
    End With
    Exit Sub
    
errPago:
    clsGeneral.OcurrioError "Error al cargar los cheques para el documento.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    lTitular.Left = 45
    lTitular.Width = Me.ScaleWidth - (lTitular.Left * 2)
    lPago.Left = lTitular.Left
    lTotal.Left = lTitular.Left
    
    lPago.Width = lTitular.Width
    lTotal.Width = lPago.Width
    lTotal.Top = Me.Height - 1100
    lPago.Height = Me.ScaleHeight - lTotal.Height - lPago.Top

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    GuardoSeteoForm Me
    
    CierroConexion
    Set miConexion = Nothing
    Set clsGeneral = Nothing
    
End Sub

Private Sub lPago_DblClick()
    
    With lPago
        If .Rows = 0 Then Exit Sub
        EjecutarApp prmPathApp & "SeguimientoCheques.exe", .Cell(flexcpData, .Row, 0)
    End With
    
End Sub

Private Sub lPago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub InicializoGrilla()

    With lPago
        .BackColor = &HC0E0FF   'Colores.Obligatorio
        .BackColorBkg = &HC0E0FF   'Colores.Obligatorio
        .Rows = 1: .Cols = 1
        .FormatString = "<Vencimiento|<Banco|<Cheque|Mon|>Importe|Librado"
        .ColWidth(0) = 1000: .ColWidth(1) = 3500: .ColWidth(2) = 1400: .ColWidth(3) = 500: .ColWidth(4) = 1100
        .WordWrap = False: .MergeCells = flexMergeSpill: .ExtendLastCol = True
    End With

    With lTotal
        .BackColor = lTitular.BackColor ' Colores.Azul
        .BackColorBkg = .BackColor
        .FontSize = 9: .FontBold = True: .ForeColor = Colores.Blanco
        .Rows = 0: .Cols = 3
        .FormatString = "<Totales|>Moneda|<Importe Dif|>Importe Al Dia|"
        .ColWidth(0) = 4000: .ColWidth(1) = 500: .ColWidth(2) = 1800: .ColWidth(3) = 1600
        .ExtendLastCol = True
    End With
End Sub

