VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmMultiple 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Corregir Existencia"
   ClientHeight    =   7230
   ClientLeft      =   3195
   ClientTop       =   2565
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   6585
   Begin VB.CommandButton bAdd 
      Caption         =   "Agregar"
      Height          =   315
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox tEFecha 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
   Begin VSFlex6DAOCtl.vsFlexGrid vsConsulta 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   11668
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agregar artículos a la existencia con costo 0.00 y &fecha:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   4050
   End
End
Attribute VB_Name = "frmMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bAdd_Click()
On Error GoTo errAdd

    If Not IsDate(tEFecha.Text) Then tEFecha.SetFocus: Exit Sub
    
    If MsgBox("¿Confirma agregar a la existencia los registros marcados?", vbQuestion + vbYesNo + vbDefaultButton2, "Adregar") = vbNo Then Exit Sub
    
    On Error GoTo errAdd
    Screen.MousePointer = 11
    
    Dim aMin As Long, mIdX As Integer, mQAddOK As Integer
    Dim mIDArticulo As Long, mQ As Long, mFecha As String
    
    mFecha = Format(tEFecha.Text, sqlFormatoF)
    With vsConsulta
    
    For mIdX = .FixedRows To .Rows - 1
        If .Cell(flexcpChecked, mIdX, 0) = flexChecked Then
            mIDArticulo = .Cell(flexcpData, mIdX, 1)
            mQ = .Cell(flexcpValue, mIdX, 2)
    
            aMin = 0
            cons = "Select Min(ComCodigo) From CMCompra"
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If Not rsAux.EOF Then aMin = rsAux(0)
            rsAux.Close
            aMin = aMin - 1
    
            cons = "Select * from CMCompra " _
                    & " Where ComFecha = '" & mFecha & "'" _
                    & " And ComArticulo = " & mIDArticulo _
                    & " And ComCodigo = " & aMin _
                    & " And ComTipo = " & TipoCV.Comercio
            Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
            If rsAux.EOF Then
                rsAux.AddNew
                
                rsAux!ComFecha = mFecha
                rsAux!ComArticulo = mIDArticulo
                rsAux!ComCantidad = mQ
                rsAux!ComCodigo = aMin
                rsAux!ComTipo = TipoCV.Comercio
                rsAux!ComCosto = 0
                rsAux!ComQOriginal = mQ
                rsAux.Update
                
                mQAddOK = mQAddOK + 1
            Else
                MsgBox "No se pudo grabar el nuevo registro (existen compras para el mismo id).", vbCritical, "ERROR"
            End If
            rsAux.Close
        End If
    Next
    End With
    Screen.MousePointer = 0
    MsgBox mQAddOK & " Registros agregados a la existencia !!", vbInformation, "Registros agregados"
    Exit Sub

errAdd:
    clsGeneral.OcurrioError "Error al agregar los registro a la existencia.", Err.Description
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    CargoResumen
End Sub

Private Sub CargoResumen()

Dim mValor As Long
    With vsConsulta
        .Rows = .FixedRows
        .Cols = 1
        .FormatString = "|Artículos|>Q|"
        .ColWidth(0) = 400: .ColWidth(1) = 3500: .ColWidth(2) = 700
        .Editable = True
    End With
    
    
    cons = "Select ArtCodigo, ArtNombre, Q = Sum(VenCantidad), ArtID " _
           & " From CMVenta, Articulo" _
           & " Where VenArticulo = ArtID " _
           & " And VenCantidad > 0 "
    
    cons = cons & " Group by ArtCodigo, ArtNombre, ArtID" _
                       & " Union All" _
           & " Select ArtCodigo, ArtNombre, Q = Sum(VenCantidad), ArtID " _
           & " From CMVenta, Articulo" _
           & " Where VenArticulo = ArtID " _
           & " And VenCantidad < 0 "

    cons = cons & " Group by ArtCodigo, ArtNombre, ArtID" _
                       & " Order by ArtCodigo"
    Set rsAux = cBase.OpenResultset(cons, rdOpenDynamic, rdConcurValues)
    
    Dim mArticulo As String: mArticulo = ""
    
    Do While Not rsAux.EOF
        With vsConsulta
            
            If mArticulo <> Trim(rsAux!ArtNombre) Then
                mArticulo = Trim(rsAux!ArtNombre)
                .AddItem ""
                .Cell(flexcpChecked, .Rows - 1, 0) = flexUnchecked
                .Cell(flexcpText, .Rows - 1, 1) = "(" & Format(rsAux!ArtCodigo, "#,000,000") & ") " & Trim(rsAux!ArtNombre)
                mValor = rsAux!ArtID: .Cell(flexcpData, .Rows - 1, 1) = mValor
                .Cell(flexcpText, .Rows - 1, 2) = rsAux!Q
            Else
            
                If Trim(.Cell(flexcpText, .Rows - 1, 2)) = "" Then
                    .Cell(flexcpText, .Rows - 1, 2) = rsAux!Q
                Else
                    .Cell(flexcpText, .Rows - 1, 2) = .Cell(flexcpText, .Rows - 1, 2) & "/" & rsAux!Q
                End If
                
            End If
        End With
        rsAux.MoveNext
    Loop
    rsAux.Close

End Sub

Private Sub vsConsulta_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

